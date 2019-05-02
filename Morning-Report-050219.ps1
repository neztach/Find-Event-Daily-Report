#region Optional (should save me from having to manually check if there's an update
Function Invoke-ElevatedCommand {
    param (
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $Scriptblock,
        [Parameter(ValueFromPipeline = $true)]
        $InputObject,
        [switch] $EnableProfile,
        [switch] $DisplayWindow
    )
	 
    begin {
        $inputItems = New-Object -TypeName System.Collections.ArrayList
    } process {
        $null = $inputItems.Add($inputObject)
    } end {
        ## Create some temporary files for streaming input and output
        $outputFile = [IO.Path]::GetTempFileName()	
        $inputFile  = [IO.Path]::GetTempFileName()
        $errorFile  = [IO.Path]::GetTempFileName()

        ## Stream the input into the input file
        $inputItems.ToArray() | Export-CliXml -Depth 1 $inputFile

        ## Start creating the command line for the elevated PowerShell session
        $commandLine = ""
        if(-not $EnableProfile) {$commandLine += "-NoProfile "}
        if(-not $DisplayWindow) {
            $commandLine       += "-Noninteractive "
            $processWindowStyle = "Hidden"
        } else {
            $processWindowStyle = "Normal"
        }
        ## Convert the command into an encoded command for PowerShell
        $commandString  = "Set-Location '$($pwd.Path)'; " + "`$output = Import-CliXml '$inputFile' | " + "& {" + $scriptblock.ToString() + "} 2>&1 ; " + "Out-File -filepath '$errorFile' -inputobject `$error;" + "Export-CliXml -Depth 1 -In `$output '$outputFile';"
        $commandBytes   = [System.Text.Encoding]::Unicode.GetBytes($commandString)
        $encodedCommand = [Convert]::ToBase64String($commandBytes)
        $commandLine   += "-EncodedCommand $encodedCommand"

        ## Start the new PowerShell process
        $process = Start-Process -FilePath (Get-Command powershell).Definition -ArgumentList $commandLine -Passthru -Verb RunAs -WindowStyle $processWindowStyle
        $process.WaitForExit()
        $errorMessage = $(Get-Content $errorFile | Out-String)
        if($errorMessage) { Write-Error -Message $errorMessage} else { if((Get-Item $outputFile).Length -gt 0) {Import-CliXml $outputFile}}

        ## Clean up
        Remove-Item $outputFile; Remove-Item $inputFile; Remove-Item $errorFile
    }
}

Function Get-CurrentModule {
    [cmdletbinding()]
    Param([switch]$Update)

    Write-Host "Getting installed modules" -ForegroundColor Yellow
    $modules = Get-Module -ListAvailable

    ## Check if PSWinReportingV2 is installed
    $moduleinstalled = $false
    $modules | ForEach {If ($_.Name -match "PSWinReportingV2"){$moduleinstalled = $true}}
    If ($moduleinstalled -eq $false){
        ## Comment out next line if you don't want to auto install module if it isn't already.
        Invoke-ElevatedCommand {Install-Module PSWinReportingV2 -Force} -DisplayWindow
    } else {
        Write-Host "PSWinReportingV2 isn't installed"
    }

    ## Group to identify modules with multiple versions installed
    $g = $modules | Group-Object Name -NoElement | Where-Object Count -GT 1

    Write-Host "Filter to modules from the PSGallery" -ForegroundColor Yellow
    $gallery = $modules.Where({$_.RepositorySourceLocation})

    Write-Host "Comparing to online versions" -ForegroundColor Yellow
    ForEach ($module in $gallery) {
        Try {
            $online = Find-Module -Name $module.name -Repository PSGallery -ErrorAction Stop
        } Catch {
            Write-Warning "Module $($module.name) was not found in the PSGallery"
        }

        If ($online.version -gt $module.version) {
            $UpdateAvailable = $True
            If ($Update){Invoke-ElevatedCommand {Install-Module PSWinReportingV2 -Force} -DisplayWindow} 
        } Else {
            $UpdateAvailable = $False
        }
        [PSCustomObject]@{
            Name             = $module.name
            MultipleVersions = ($g.name -contains $module.name)
            InstalledVersion = $module.version
            OnlineVersion    = $online.version
            Update           = $UpdateAvailable
            Path             = $module.modulebase
        }
    }
}

Get-CurrentModule -Update
#endregion optional

Import-Module PSWinReportingV2

#region mail vars
$SMTPServer = 'mailserver.yourdomain.local'
$from       = 'ADEventReport@yourdomain.com' 
$to         = 'yourmail@yourdomain.com'
$subject    = 'Daily Report'
#endregion mail vars

#region html code and styling for mail
$html             = "<body style=`"font-family: Lucida Sans Unicode, Lucida Grande, Sans-Serif; font-size: 12pt;`">`n"
$htmlEnd          = "</body>`n"
$TableStart       = "  <table style = `"font-family: Lucida Sans Unicode, Lucida Grande, Sans-Serif; font-size: 12pt; background: #fff; margin: 10px; border-collapse: collapse; text-align: left;`">`n"
$TableEnd         = "  </table><br />`n"
$HeaderRowStart   = "    <thead>`n      <tr>`n        <th style = `"font-size: 14px; font-weight: normal; color: #039; padding: 0px 5px; border-bottom: 2px solid #6678b1;`">`n"
$HeaderRowBetween = "</th>`n        <th style = `"font-size: 14px; font-weight: normal; color: #039; padding: 0px 5px; border-bottom: 2px solid #6678b1;`">`n"
$HeaderRowEnd     = "</th>`n      </tr>`n    </thead>`n"
$TableBodyStart   = "    <tbody>`n"
$TableBodyEnd     = "    </tbody>`n"
$TableRowStart    = "      <tr>`n        <td style = `"font-size: 12px; border-bottom: 1px solid #ccc; color: #669; padding: 0px 8px;`">`n"
$CellBetween      = "</td>`n        <td style = `"font-size: 12px; border-bottom: 1px solid #ccc; color: #669; padding: 0px 8px;`">`n"
$TableRowEnd      = "</td>`n      </tr>`n"
#endregion

#region Common function to attempt to cut down on repeating code
Function Get-ThisTable($Event, [string]$Header){
    ## Look through the returned results and try to remove columns that are empty
    $tempeventheads = $Event | 
        Get-Member -MemberType 'NoteProperty' | 
        Where-Object {($_.Definition).split('=')[1] -notlike ""} | 
        Where-Object {$_.Name -notlike "*Domain Controller*"} | ## Seems irrelevant due to the "Gathered From" field generally being the same
        Where-Object {$_.Name -notlike "*Record ID*"} |         ## Not needed for a daily report for my purposes
        Select-Object -ExpandProperty 'Name'

    ## Start new section
    $htmltable  = "  <strong>$($Header)</strong>`n"
    ## Start new table + Start the table header row    ---- <table><thead><tr><th>
    $htmltable += $TableStart + $HeaderRowStart
    ## Iterate through all events for this sub-section
    For ($i=0; $i -lt $tempeventheads.count; $i++){
        ## Add header to header row
        $htmltable += $tempeventheads[$i]
        ## Check if we're on the 2nd to last event header yet
        If ($i -le ($tempeventheads.count - 2)){
            ## If not, add code between cells          ---- </th><th>
            $htmltable += $HeaderRowBetween
        } else {
            ## If we're on the last header, finish row ---- </th></tr></thead>
            $htmltable += $HeaderRowEnd
        }
    }
    ## Start the table body                            ---- <tbody>
    $htmltable += $TableBodyStart
    ## Iterate through the events
    Foreach ($ADCCC in $Event){
        ## Start the row                               ---- <tr><td>
        $htmltable += $TableRowStart
        ## Iterate through the objects for proper placement under headers
        For ($j=0; $j -lt $tempeventheads.count; $j++){
            ## Turn content into string
            $field = $($tempeventheads[$j])
            ## Add event column content to table cell
            $htmltable += $ADCCC.$field
            ## Check if we're on the 2nd to last event column yet
            If ($j -le ($tempeventheads.count - 2)){
                ## If not, add code between cells      ---- </td><td>
                $htmltable += $CellBetween
            } else {
                ## If we're on the last column, finish row ---- </td></tr>
                $htmltable += $TableRowEnd
            }
        } ## Finish this row, do the next one
    } ## finish iterating through events
    ## Finish the table ---- </tbody></table>
    $htmltable += $TableBodyEnd + $TableEnd
    ## Return the table to be added to $html
    Return $htmltable
}
#endregion Function

$Reports = @(
    'ADUserChanges'
    'ADUserChangesDetailed'
    'ADComputerChangesDetailed'
    'ADUserStatus'
    'ADUserLockouts'
    #ADUserLogon
    'ADUserUnlocked'
    'ADComputerCreatedChanged'
    'ADComputerDeleted'
    #'ADUserLogonKerberos'
    'ADGroupMembershipChanges'
    'ADGroupEnumeration'
    'ADGroupChanges'
    'ADGroupCreateDelete'
    'ADGroupChangesDetailed'
    'ADGroupPolicyChanges'
    'ADLogsClearedSecurity'
    'ADLogsClearedOther'
    #ADEventsReboots
)

## In my situation, I'm making this a sceduled task to run every morning at 7AM
$yesterday = [DateTime]::Today.AddDays(-1).AddHours(7) # Yesterday at 7AM
$today     = [DateTime]::Today.AddHours(7)             # Today at 7AM
## Query Domain for DC list, remove the domain.local to just keep the server name
$ServerList = Get-WinADForestControllers | Select-Object -Expand HostName | ForEach-Object {$_.split('.')[0]}
$Events    = Find-Events -Report $Reports -DateFrom $yesterday -DateTo $today -Servers $ServerList

### Computer Changes
## Check and make sure there are any events to even bother processing this section
If (($Events.ADComputerCreatedChanged) -OR ($Events.ADComputerDeleted) -OR ($Events.ADComputerChangesDetailed)) {
    ## If there are, add a label for this section
    $html += "  <h1>Computer Changes</h1>`n"
    ## If there are any events in "ADComputerCreatedChanged", run our function and create this table
    If ($Events.ADComputerCreatedChanged)  {$html += Get-ThisTable -Event $Events.ADComputerCreatedChanged -Header "AD Computer Created/Changed"}
    ## Same
    If ($Events.ADComputerDeleted)         {$html += Get-ThisTable -Event $Events.ADComputerDeleted -Header "AD Computer Deleted"}
    ## My results haven't yeilded any results here yet, so I'm just dumping everything til I do.
    ## Gives me the opportunity to see if this one will have to be treated differently from the others.
    If ($Events.ADComputerChangesDetailed) {
        $htmltable  = "  <strong>AD Computer Changes Detailed</strong>`n"
        $htmltable += "    <p>`n"
        foreach ($ADCCD in $Events.ADComputerCreatedChanged){$htmltable += $ADCCD}
        $htmltable += "    </p>`n"
        $htmltable += $TableBodyEnd + $TableEnd
        $html += $htmltable
    }
}
## End This Section, Repeat these steps for the rest of the way down.

### Group Changes
If (($Events.ADGroupCreateDelete) -OR ($Events.ADGroupMembershipChanges) -OR ($Events.ADGroupEnumeration) -OR ($Events.ADGroupChanges) -OR ($Events.ADGroupChangesDetailed)) {
    $html += "  <h1>Group Changes</h1>`n"
    If ($Events.ADGroupCreateDelete)      {
        $htmltable  = "  <strong>AD Group Create/Delete</strong>`n"
        $htmltable += "    <p>`n"
        foreach ($ADGCD in $Events.ADGroupCreateDelete){$htmltable += $ADGCD}
        $htmltable += "    </p>`n"
        $htmltable += $TableBodyEnd + $TableEnd
        $html += $htmltable
    }
    If ($Events.ADGroupMembershipChanges) {$html += Get-ThisTable -Event $Events.ADGroupMembershipChanges -Header "AD Group Membership Changes"}
    If ($Events.ADGroupEnumeration)       {
        $htmltable  = "  <strong>AD Group Enumeration</strong>`n"
        $htmltable += "    <p>`n"
        foreach ($ADGE in $Events.ADGroupEnumeration){$htmltable += $ADGE}
        $htmltable += "    </p>`n"
        $htmltable += $TableBodyEnd + $TableEnd
        $html += $htmltable
    }
    If ($Events.ADGroupChanges)           {$html += Get-ThisTable -Event $Events.ADGroupChanges -Header "AD Group Changes Detailed"}
    If ($Events.ADGroupChangesDetailed)   {
        $htmltable  = "  <strong>AD Group Changes Detailed</strong>`n"
        $htmltable += "    <p>`n"
        foreach ($ADGCDE in $Events.ADGroupChangesDetailed){$htmltable += $ADGCDE}
        $htmltable += "    </p>`n"
        $htmltable += $TableBodyEnd + $TableEnd
        $html += $htmltable
    }
}

### User Changes
If (($Events.ADUserChanges) -OR ($Events.ADUserChangesDetailed) -OR ($Events.ADUserLockouts) -OR ($Events.ADUserStatus) -OR ($Events.ADUserUnlocked)) {
    $html += "  <h1>User Changes</h1>`n"
    If ($Events.ADUserChanges)            {$html += Get-ThisTable -Event $Events.ADUserChanges -Header "AD User Changes"}
    If ($Events.ADUserChangesDetailed)    {
        $htmltable  = "  <strong>AD user Changes Detailed</strong>`n"
        $htmltable += "    <p>`n"
        foreach ($ADUCD in $Events.ADUserChangesDetailed){$htmltable += $ADUCD}
        $htmltable += "    </p>`n";$htmltable += $TableBodyEnd + $TableEnd
        $html += $htmltable
    }
    If ($Events.ADUserLockouts)           {$html += Get-ThisTable -Event $Events.ADUserLockouts -Header "AD User Lockouts"}
    If ($Events.ADUserStatus)             {$html += Get-ThisTable -Event $Events.ADUserStatus -Header "AD User Status Changes"}
    #            If ($ADUS."User Affected" -notlike "*Health*"){
    #                $htmltable += $tBodySt + $ADUS.Action + $CellR + $ADUS."User Affected" + $CellR + $ADUS.Who + $CellR + $ADUS.When + $CellR + $ADUS."Event ID" + $CellR + $ADUS."Gathered From" + $CellR + $ADUS."Gathered LogName" + $rEnd
    #            }
    If ($Events.ADUserUnlocked)           {$html += Get-ThisTable -Event $Events.ADUserUnlocked -Header "AD User Unlocks"}
}

#'Group Policy Changes'
If ($Events.ADGroupPolicyChanges) {
    $html += "  <h1>Group Policy Changes</h1>`n"
    If ($Events.ADGroupPolicyChanges)     {
        $htmltable  = "  <strong>AD Group Policy Changes</strong>`n"
        $htmltable += "    <p>`n"
        foreach ($ADGPC in $Events.ADGroupPolicyChanges){$htmltable += $ADGPC}
        $htmltable += "    </p>`n"
        $htmltable += $TableBodyEnd + $TableEnd
        $html += $htmltable
    }
}

#'Logs'
If (($Events.ADLogsClearedOther) -OR ($Events.ADLogsClearedSecurity)) {
    $html += "  <h2>Logs</h2>`n"
    If ($Events.ADLogsClearedOther)       {
        $htmltable  = "  <strong>AD Logs Cleared (Other)</strong>`n"
        $htmltable += "    <p>`n"
        foreach ($ADLCO in $Events.ADLogsClearedOther){$htmltable += $ADLCO}
        $htmltable += "    </p>`n"
        $htmltable += $TableBodyEnd + $TableEnd
        $html += $htmltable
    }
    If ($Events.ADLogsClearedSecurity)    {
        $htmltable  = "  <strong>AD Logs Cleared (Security)</strong>`n"
        $htmltable += "    <p>`n"
        foreach ($ADLCS in $Events.ADLogsClearedSecurity){$htmltable += $ADLCS}
        $htmltable += "    </p>`n"
        $htmltable += $TableBodyEnd + $TableEnd
        $html += $htmltable
    }
}

## Finish the html code properly ---- </body>
$html += $htmlend

## If any events at all were noted, send the html we've been creating in a mail
if($Events){Send-MailMessage -smtpserver $smtpserver -from $from -to $to -subject $subject -body $html -bodyashtml}

