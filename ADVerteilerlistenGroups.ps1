<#
.SYNOPSIS
    IAM-Skript für ABLE Management Group
.DESCRIPTION
    Dient dem Anlegen von AD-Gruppen für Verteilerlisten
.COMPONENT
    Benötigt die Module ActiveDirectory und dbatools
.PARAMETER UpdateDbatools
    Default: false
    Prüft auf Updates für das dbatools-Modul und installiert diese ggf.
.PARAMETER InstallDbatools
    Default: false
    Installiert (sofern nicht bereits vorhanden) das dbatools-Modul ohne weitere Rückfrage
.PARAMETER SQLInstance
    Default: GMZRZSQC070.ferchau.local
    Die Instanz-Notation wird unterstützt (SRV-SQL01.domain.local\INSTANCENAME)
.PARAMETER SQLDatabase
    Default: IAM
.PARAMETER SQLTable
    Default: [IAM].[dbo].[t_PowershellVerarbeitung]
.PARAMETER ExchangeFQDN
    Default: GMZRZEX001.Ferchau.local
    Wird für eine remote-Powershell zu Exchange genutzt
.PARAMETER ClearLogs
    Entfernt alle Log-Dateien älter als 30min, statt der üblichen 30 Tage
.PARAMETER WaitMultiplier
    Akzeptiert Werte zwischen 0,5 und 10
    Multiplikator der Wartezeiten, falls Überprüfungen durch langsame AD-Replikation fehlschlagen, obwohl die Ausführung erfolgreich war
.NOTES
    Author:      thmueller@cancom.de

    Version:     0.7
    Date:        March 27, 2018
    Email Handling per Auftragstabelle, Löschen von AD-Accounts implementiert
#>
#Requires -Version 4.0
##### ----- #####
# Parameter
##### ----- #####
#region param
param(
    [Parameter(Mandatory = $false)]
    [switch]$UpdateDbatools = $false,
    [Parameter(Mandatory = $false)]
    [switch]$InstallDbatools = $false,
    [Parameter(Mandatory = $false)]
    [string]$SQLInstance = "GMZRZSQC070.ferchau.local",
    [Parameter(Mandatory = $false)]
    [string]$SQLDatabase = "IAM",
    [Parameter(Mandatory = $false)]
    [string]$SQLTable = "[IAM].[dbo].[t_PowershellVerarbeitung]",
    [Parameter(Mandatory = $false)]
    [string]$SQLLogTable = "[IAM].[dbo].[t_logfile]",
    [Parameter(Mandatory = $false)]
    [switch]$RetryFailed = $false,
    [Parameter(Mandatory = $false)]
    [string]$ExchangeFQDN = "GMZRZEX001.Ferchau.local",
    [Parameter(Mandatory = $false)]
    [ValidateRange(0.5 , 10)]
    [single]$WaitMultiplier = 1,
    [Parameter(Mandatory = $false)]
    [switch]$ClearLogs = $false
)
#endregion
##### ----- #####
# Variables (Änderungen möglich)
##### ----- #####
#region var

#Wartezeiten vor der Überprüfung der ausgeführten Aktionen
$WaitSeconds = @{   'Group' = 60
                    'RemoveGroup' = 60
                    'Deaktivieren' = 90
                    'default' = 20
                }

#Spaltennamen in den SQL-Tabellen
$column = @{    'userDN' = 'Distinguishedname'  #SQLTable, varchar, DistinguishedName des zu ändernden User
                'attribute' = 'AD_Attribut'     #SQLTable, varchar, AD-Attribut oder Kommando aus (Group, RemoveGroup, Deaktivieren)
                'newValue' = 'newValue'         #SQLTable, varchar, neu zu setzender Wert, bei Gruppen DN der Gruppe, kann bei Deaktivieren leer gelassen werden
                'id' = 'ID'                     #SQLTable, bigint, IDENTITY mit Autoinkrement
                'emailtext' = 'eMailText'       #SQLTable, text, Inhalt einer zu verschickenden Mail
                'emailrecipient' = 'eMailRecipient'       #SQLTable, text, Inhalt einer zu verschickenden Mail
                'emailsubject' = 'eMailSubject'       #SQLTable, text, Inhalt einer zu verschickenden Mail
                'success' = 'success'           #SQLTable, bit, 1=erfolgreich ausgeführt, 0=nicht erfolgreich
                'completed' = 'verarbeitet'     #SQLTable, 
                'logtext' = 'Meldungstext'      #SQLLogTable, varchar, zu speichernder Log-Text
}
#Für Mail-Benachrichtigung bei Adress-Übergabe
$MailServer = "exchange2013.able-group.de" 
$MailServerport = 587
$MailFrom = "support@able-group.de"


##### ----- #####
# Ab hier sind in der Regel keine Änderungen nötig
##### ----- #####
$Scriptpath = split-path -parent $MyInvocation.MyCommand.Definition -ErrorAction SilentlyContinue
$LogPrefix = "ADUserCorrectedFromIAM_"
$LogExtension = "txt"
$LogFile = $Scriptpath + "\" + $LogPrefix + (Get-Date -Format yyyyMMdd-HH.mm).ToString() + "." + $LogExtension
$DbaToolsInstalled = Get-Module -ListAvailable -Name dbatools
$Error.Clear()

if(!($ExportPath.EndsWith("`\"))){
    $ExportPath += "`\"
}

#endregion
##### ----- #####
# Functions
##### ----- #####
#region func
function Write-Color {
    param (
        [Parameter(Mandatory = $true)]
        [String[]]$Text, 
        [Parameter(Mandatory = $false)]
        [ConsoleColor[]]$Color = "White", 
        [Parameter(Mandatory = $false)]
        [int]$StartTab = 0, 
        [Parameter(Mandatory = $false)]
        [int] $LinesBefore = 0,
        [Parameter(Mandatory = $false)]
        [int] $LinesAfter = 0, 
        [Parameter(Mandatory = $false)]
        [string] $LogFile = $LogFile, 
        [Parameter(Mandatory = $false)]
        $TimeFormat = "yyyy-MM-dd HH:mm:ss"
    ) 
    $DefaultColor = $Color[0]
    if ($LinesBefore -ne 0) {  for ($i = 0; $i -lt $LinesBefore; $i++) { Write-Host "`n" -NoNewline } } # Add empty line before
    if ($StartTab -ne 0) {  for ($i = 0; $i -lt $StartTab; $i++) { Write-Host "`t" -NoNewLine } }  # Add TABS before text
    if ($Color.Count -ge $Text.Count) {
        for ($i = 0; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -NoNewLine } 
    } else {
        for ($i = 0; $i -lt $Color.Length ; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -NoNewLine }
        for ($i = $Color.Length; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $DefaultColor -NoNewLine }
    }
    Write-Host
    if ($LinesAfter -ne 0) {  for ($i = 0; $i -lt $LinesAfter; $i++) { Write-Host } }  # Add empty line after
    if ($LogFile -ne "") {
        $TextToFile = ""
        for ($i = 0; $i -lt $Text.Length; $i++) {
            $TextToFile += $Text[$i]
        }
        Write-Output "[$([datetime]::Now.ToString($TimeFormat))] $TextToFile" | Out-File $LogFile -Encoding unicode -Append
    }
}

function Use-PSModule
{
    param (
        [parameter(Mandatory = $true)][string] $name
    )
    $retVal = $true
    if (!(Get-Module -Name $name))
    {
        $retVal = Get-Module -ListAvailable | Where-Object { $_.Name -eq $name }
        if ($retVal)
        {
            try
            {
                Import-Module $name -ErrorAction SilentlyContinue
            }
            catch
            {
                $retVal = $false
            }
        }
    }
    return $retVal
}

#endregion
##### ----- #####
# Prerequisites (Load modules etc)
##### ----- #####
#region pre
Write-Color -Text "Starte Script",", User: ",([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -Color Red,White,Cyan
Write-Color -Text "Aufruf: ",$MyInvocation.Line -Color White,Cyan
if($ClearLogs){
    Write-Color -Text "Switch ","-ClearLogs"," wurde benutzt" -Color White,Cyan,White
    $oldlogfiles = Get-ChildItem -LiteralPath $Scriptpath -Filter ($LogPrefix + "*." + $LogExtension) | Where-Object {$_.LastWriteTime -le (Get-Date).AddMinutes(-30)}
}else{
    $oldlogfiles = Get-ChildItem -LiteralPath $Scriptpath -Filter ($LogPrefix + "*." + $LogExtension) | Where-Object {$_.LastWriteTime -le (Get-Date).AddDays(-30)}
}
if ($oldlogfiles.count -gt 0){
    Write-Color -Text "Alte Logs aufräumen." -Color White
    foreach ($f in $oldlogfiles){
        Write-Color -Text "Entferne: ",$f.Name -Color White,Red
        Remove-Item $f.FullName -ErrorAction SilentlyContinue
    }
}

if (!(Get-Module -Name ActiveDirectory -ListAvailable)){
    Import-Module ServerManager -ErrorAction SilentlyContinue
    Add-WindowsFeature RSAT-AD-Powershell -ErrorAction SilentlyContinue
    
}
if(!(Use-PSModule ActiveDirectory)){
    Write-Color -Text "Script abgebrochen. ","ActiveDirectory","-Module konnte nicht geladen (oder installiert) werden" -Color White,Red,White
    return
}else{
    Write-Color -Text "ActiveDirectory","-Modul erfolgreich geladen." -Color Red,White
}

if(!$DbaToolsInstalled){
    Write-Color -Text "dbatools"," sind nicht installiert. " -Color Red,White
    if(!$InstallDbatools){
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Installs dbatools"
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Ends the script, you need to install dbatools manually"
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
        $caption = "Need to install dbatools!"
        $message = "Do you allow the script to trust PSGallery and install dtatools?"
        $result = $Host.UI.PromptForChoice($caption,$message,$choices,1)
        if($result -eq 0) {
              $InstallDbatools = $true
        }else{
              Write-Color -Text "Bitte ","dbatools"," manuell installieren und Script erneut ausführen" -Color White,Red,White
              return
        }
    }  
}else{
    if($UpdateDbatools){
        Write-Color -Text "Switch ","-UpdateDbatools"," wurde benutzt" -Color White,Cyan,White
        Write-Color -Text "Prüfe auf mögliche Updates" -Color White
        if($PSVersionTable.PSVersion.Major -ge 5){
            $localversion = (Get-Module -ListAvailable -Name dbatools).Version
            $onlineversion = (Find-Module dbatools).Version
            Write-Color -Text "Lokale dbatools-Version: ",$localversion.ToString() -Color White,Yellow
            Write-Color -Text "Aktuelle dbatools-Version: ",$onlineversion.ToString() -Color White,Green
            if($onlineversion -gt $localversion){
                Write-Color -Text "Versuche Update" -Color White
                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted 
                Update-Module dbatools 
            }
        }else{
            Write-Color -Text "Legacy Powershell Version (before 5.0): ",$PSVersionTable.PSVersion.ToString() -Color White,Yellow
            Write-Color -Text "Kein Update möglich: ","dbatools." -Color White,Red
        }
    }
}

if($InstallDbatools){
    Write-Color -Text "Switch ","-InstallDbatools"," wurde benutzt" -Color White,Cyan,White
    if(!$DbaToolsInstalled){
        Write-Color -Text "Installiere jetzt ","dbatools" -Color White,Red
        if($PSVersionTable.PSVersion.Major -ge 5){
            $pol = (Get-PSRepository -Name PSGallery).InstallationPolicy
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted 
            Install-Module dbatools -Scope CurrentUser
        }else{
            Write-Color -Text "Legacy Powershell Version (before 5.0): ",$PSVersionTable.PSVersion.ToString() -Color White,Yellow
            Write-Color -Text "Versuche via WebRequest zu installieren" -Color White
            Invoke-Expression (Invoke-WebRequest -UseBasicParsing https://dbatools.io/in)
        }
    }
}

if(!(Use-PSModule dbatools)){
    Write-Color -Text "Script abgebrochen. ","dbatools","-Module konnte nicht geladen (oder installiert) werden" -Color White,Red,White
    return
}else{
    Write-Color -Text "dbatools","-Modul erfolgreich geladen." -Color Red,White
}

#endregion
##### ----- #####
# Main Script
##### ----- #####
#region main

$query = "SELECT " + $column['id'] + "," + $column['newvalue'] + "," + $column['attribute'] + "," + $column['userDN'] + " FROM $SQLTable WHERE [AD_Attribut] in ('Group create') and verarbeitet IS NULL"

if ($RetryFailed){
    Write-Color -Text "Switch ","-Retryfailed"," wurde benutzt" -Color White,Cyan,White
    $query += " OR (" + $column['completed'] + " > DATEADD(day,-2,GETDATE()) AND success = 0)"
}
$query += " ORDER BY [" + $column['id'] + "]"

$commands = @()
$currentcommand = $null
$change_running = 0
$change_ok = 0
$change_fail = 0

try{
    Write-Color -Text "Suche nach Aufträgen in der Datenbank ",$SQLDatabase," in der SQL-Instanz ",$SQLInstance -Color White,Cyan,White,Cyan
    Write-Color -Text "Query für die Aufträge: ", $query -Color White,Cyan
    $commands = Invoke-DbaQuery -SqlInstance $SQLInstance -Query $query -Database $SQLDatabase -WarningAction Stop
    
    if($null -eq $commands.count){
        $change_total = 1
    }else{
        $change_total = $commands.Count
    }
    if($change_total -lt 1){
        Write-Color -Text "Keine Aufträge in Datenbank gefunden." -Color White
        break;
    }else{
        Write-Color -Text "Datensätze zu verarbeiten: ",$change_total -Color White,Cyan
    }

    if($commands.($column['attribute']).Contains("Group create") -or $commands.($column['attribute']).Contains("Group delete")){
        Write-Color -Text "Remote-Sessions zu Exchange etablieren" -Color Yellow
        $remotesessions = $true
        $exc_session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $ExchangeFQDN + "/PowerShell/") -Authentication Kerberos -ErrorAction Stop
        $result = Import-PSSession $exc_session -ErrorAction SilentlyContinue
    }
    
    foreach ($c in $commands){
        $change_running++
        $currentcommand = $c
        $attributeToChange = $($c.($column['attribute']))
        Write-Color -Text "Datensatz: ",$($c.($column['UserDN']))," - ",$($c.($column['attribute'])) -Color White,Cyan,White,Cyan
        try{
            switch ($attributeToChange){
                "Group create" {
                    Write-Color -Text "(",$change_running,"/",$change_total,") Auftrag: ",$ADUser.Name," (",$ADUser.Samaccountname,") ","soll in Gruppe ",$($c.($column['newValue'])) -Color White,Cyan,White,White,White,Red,White,Cyan,White,White,Cyan -StartTab 1 -LinesAfter 1
                    $OrganizationalUnit = "Ferchau.local/Administration/Gruppen/Exchange/groups/dynamic"
                    $replacingValue = $($c.($column['newValue']))
                    $executionCommand = "New-DistributionGroup -Name `"" + $replacingValue + "`" -OrganizationalUnit `"" + $OrganizationalUnit + "`" -SAMAccountName `"" + $replacingValue + "`" -Type `"Distribution`" -Alias `"" + $replacingValue + "`" -MemberDepartRestriction Closed -ErrorAction SilentlyContinue" 
                    $checkcommand = "(Get-AdGroup `"" +  $replacingValue + "`)"
                    break;
                }
                #"Group delete"{
                #    Write-Color -Text "(",$change_running,"/",$change_total,") Auftrag: ",$ADUser.Name," (",$ADUser.Samaccountname,") ","entfernen aus Gruppe ",$($c.($column['newValue'])) -Color White,Cyan,White,White,White,Red,White,Cyan,White,White,Cyan -StartTab 1 -LinesAfter 1
                #    $GroupDomain = ($($c.($column['newValue'])) -split "," | Where-Object {$_ -like "DC=*"}) -join "." -replace ("DC=","")
                #    $replacingValue = $null
                #    $executionCommand = "Get-AdGroup `"" + $($c.($column['newValue'])) + "`" -Server " + $GroupDomain + " | Remove-AdgroupMember -Members `$ADUser -Confirm:`$false -ErrorAction SilentlyContinue"
                #    $checkcommand = "(Get-ADGroup `"" + $ADUser.DistinguishedName + "`" -Properties MemberOf -Server " + $Domain + " | Select-Object -ExpandProperty Memberof | where-Object { `$_ -like `"*" + $($c.($column['newValue'])) + "*`"})"
                #    break;
                #}
            }
        }catch{
            if($currentcommand){
                $q = "UPDATE $SQLTable SET " + $column['success'] + " = 0, " + $column['completed'] + " = GETDATE() WHERE ID = '" + $($currentcommand.($column['id'])) + "'"
                $result = Invoke-DbaQuery -SqlInstance $SQLInstance -Query $q -Database $SQLDatabase
            }
            Write-Color -Text "Fehler: ",(" " + $Error[0].Exception.Message),"`r`n`tObjekt: ",(" " + $Error[0].TargetObject) -Color Red,White,Red,Cyan -LinesBefore 1 -LinesAfter 1 -StartTab 1
            Write-Color -Text "Position: ",(" " + $Error[0].InvocationInfo.PositionMessage),"`r`n`tLine: ",(" " + $Error[0].InvocationInfo.Line) -Color Red,White,Red,Cyan -LinesBefore 1 -LinesAfter 1 -StartTab 1
            Invoke-DbaQuery -SqlInstance $SQLInstance -Database $SQLDatabase -Query ("INSERT INTO " + $SQLLogTable + " (" + $column['logtext'] + ") VALUES ('Powershell-Fehler: " + $Error[0].Exception.Message + " bei Objekt: " + $Error[0].TargetObject + "')")
            Write-Color -Text "Logfile: ",$LogFile -Color Red,Yellow
            Invoke-DbaQuery -SqlInstance $SQLInstance -Database $SQLDatabase -Query ("INSERT INTO " + $SQLLogTable + " (" + $column['logtext'] + ") VALUES ('Powershell-Logfile: (" + $ENV:COMPUTERNAME + ") " + $LogFile + "')")
            continue
        }

        Write-Color -Text "(",$change_running,"/",$change_total,") Ausführung: ",$executionCommand -Color White,Cyan,White,White,White,Cyan
        try{
            $result = Invoke-Expression $executionCommand -ErrorAction SilentlyContinue
        }catch{
            Write-Color -Text "Fehler:",(" " + $Error[0].Exception.Message),"`r`n`tObjekt:",(" " + $Error[0].TargetObject) -Color Red,White,Red,Cyan -LinesBefore 1 -LinesAfter 1 -StartTab 1
            Invoke-DbaQuery -SqlInstance $SQLInstance -Database $SQLDatabase -Query ("INSERT INTO " + $SQLLogTable + " (" + $column['logtext'] + ") VALUES ('Powershell-Fehler: " + $Error[0].Exception.Message + " bei Objekt: " + $Error[0].TargetObject + "')")
        }        
        if($WaitSeconds.ContainsKey($attributeToChange)){
            [int]$totalwait = $WaitSeconds[$attributeToChange] * $WaitMultiplier
        }else{
            [int]$totalwait = $WaitSeconds['default'] * $WaitMultiplier
        }
        Write-Color -Text "(AD-Replikation abwarten) ","Wartezeit bis zur Überprüfung: ",$totalwait," Sekunden" -Color Yellow,White,Cyan,White
        Start-Sleep -Seconds $totalwait | Out-Null
        Write-Color -Text "(",$change_running,"/",$change_total,") Überprüfung: ",$checkcommand -Color White,Cyan,White,White,White,Cyan
        if ((Invoke-Expression $checkcommand -ErrorAction SilentlyContinue) -eq $replacingValue) {
            Write-Color -Text "Erfolgreich: ",$attributeToChange,":",(" " + $replacingValue) -Color Green,White,White,Green -StartTab 2 -LinesAfter 1
            $successvalue = 1
            $change_ok++
            if( [string]$($c.($column['emailsubject'])) -ne "" ){
                Write-Color -Text "Email an ",$($c.($column['emailrecipient']))," senden, Betreff: ",$($c.($column['emailsubject'])) -Color White,Cyan,White,Cyan
                Send-MailMessage -From $MailFrom -Body $($c.($column['emailtext'])) -To $($c.($column['emailrecipient'])) -Subject $($c.($column['emailsubject'])) -SmtpServer $MailServer -Port $MailServerPort | Out-Null
            }
        }else{
            Write-Color -Text "Fehlgeschlagen: ",$attributeToChange,":",(" " + $replacingValue) -Color Red,White,White,Red -StartTab 2 -LinesAfter 1
            Write-Color -Text "Letzter Fehler: ",(" " + $Error[0].Exception.Message),"`r`n`tObjekt: ",(" " + $Error[0].TargetObject) -Color Red,White,Red,Cyan -LinesBefore 1 -LinesAfter 1 -StartTab 1
            $successvalue = 0
            Invoke-DbaQuery -SqlInstance $SQLInstance -Database $SQLDatabase -Query ("INSERT INTO " + $SQLLogTable + " (" + $column['logtext'] + ") VALUES ('Powershell-Ausführung fehlgeschlagen für ID " + $($c.($column['id'])) + "')")
            $change_fail++
        }
        $q = "UPDATE $SQLTable SET " + $column['success'] + " = " + $successValue + ", " + $column['completed'] + " = GETDATE() WHERE ID = '" + $($c.($column['id'])) + "'"
        $result = Invoke-DbaQuery -SqlInstance $SQLInstance -Query $q -Database $SQLDatabase
        if ($change_running % 100 -eq 0){
            Write-Color -Text $change_running," von ",$change_total ," Änderungen ausgeführt: ",$change_ok," erfolgreich, ",$change_fail," fehlgeschlagen" -Color Cyan,White,Cyan,White,Green,White,Red,White
        }
    }
   
}
catch{
    if($currentcommand){
        $q = "UPDATE $SQLTable SET " + $column['success'] + " = 0, " + $column['completed'] + " = GETDATE() WHERE ID = '" + $($currentcommand.($column['id'])) + "'"
        $result = Invoke-DbaQuery -SqlInstance $SQLInstance -Query $q -Database $SQLDatabase
    }
    Write-Color -Text "Fehler:",(" " + $Error[0].Exception.Message),"`r`n`tObjekt:",(" " + $Error[0].CategoryInfo.Activity) -Color Red,White,Red,Cyan -LinesBefore 1 -LinesAfter 1 -StartTab 1
    Write-Color -Text "Position: ",(" " + $Error[0].InvocationInfo.PositionMessage),"`r`n`tLine: ",(" " + $Error[0].InvocationInfo.Line) -Color Red,White,Red,Cyan -LinesBefore 1 -LinesAfter 1 -StartTab 1
    Invoke-DbaQuery -SqlInstance $SQLInstance -Database $SQLDatabase -Query ("INSERT INTO " + $SQLLogTable + " (" + $column['logtext'] + ") VALUES ('Powershell-Fehler: " + $Error[0].Exception.Message + " bei Objekt: " + $Error[0].TargetObject + "')")
    Write-Color -Text "Logfile: ",$LogFile -Color Red,Yellow
    Invoke-DbaQuery -SqlInstance $SQLInstance -Database $SQLDatabase -Query ("INSERT INTO " + $SQLLogTable + " (" + $column['logtext'] + ") VALUES ('Powershell-Logfile: (" + $ENV:COMPUTERNAME + ") " + $LogFile + "')")
}
#endregion
##### ----- #####
# Cleanup
##### ----- #####
#region clean
finally{
    Invoke-DbaQuery -SqlInstance $SQLInstance -Database $SQLDatabase -Query ("INSERT INTO " + $SQLLogTable + " (" + $column['logtext'] + ") VALUES ('Powershell-Verabreitung abgeschlossen: " + $change_ok + " erfolreich, " + $change_fail + " fehlgeschlagen')")

    if($InstallDbatools){
        Set-PSRepository -Name PSGallery -InstallationPolicy $pol
    }

    if($remotesessions){
        Write-Color -Text "Remote-Sessions beenden" -Color Yellow
        $exc_session | Remove-PSSession -ErrorAction SilentlyContinue
    }
    Remove-Module ActiveDirectory -ErrorAction SilentlyContinue
    if(!(Get-Module ActiveDirectory)){
        Write-Color -Text "ActiveDirectory","-Modul erfolgreich entladen." -Color Red,White
    }
    Remove-Module dbatools -ErrorAction SilentlyContinue
    if(!(Get-Module dbatools)){
        Write-Color -Text "dbatools","-Modul erfolgreich entladen." -Color Red,White
    }
    Write-Color -Text "Beende Script" -Color Red
    exit 0
}
#endregion