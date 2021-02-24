<#
.SYNOPSIS
    IAM-Skript für ABLE Management Group
.DESCRIPTION
    Active Directory User werden durch Aufträge in einer SQL-Tabelle verändert. Darüberhinaus sind Gruppen-Management und Deaktivierung Bestandteile des Skripts
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
 .PARAMETER SQLLogTable
    Default: [IAM].[dbo].[t_logfile]  
.PARAMETER RetryFailed
    Default: false
    Aufträge der letzten 2 Tage, die nicht erfolgreich ausgeführt wurden (mit success=0 in der Datenbank markiert), werden erneut ausgeführt
.PARAMETER ExchangeFQDN
    Default: GMZRZEX001.Ferchau.local
    Wird für eine remote-Powershell zu Exchange genutzt
.PARAMETER LyncFQDN
    Default: lync2013.Ferchau.local
    Wird für eine remote-Powershell zu Lync genutzt
.PARAMETER ExportPath
    Default: \\ferchau.local\fileservice\IAM
    Pfad für Exports von Postfach und Lync-Config (möglichst UNC-Pfad verwenden)
.PARAMETER DeactivatedOU
    Default: "OU=Deaktiviert,OU=Benutzermanagement"
    Angabe der Organisationseinheit ohne den Domänenteil, da dieser individuell bestimmt wird.
.PARAMETER ClearLogs
    Entfernt alle Log-Dateien älter als 30min, statt der üblichen 30 Tage
.PARAMETER WaitMultiplier
    Akzeptiert Werte zwischen 0,5 und 10
    Multiplikator der Wartezeiten, falls Überprüfungen durch langsame AD-Replikation fehlschlagen, obwohl die Ausführung erfolgreich war
.EXAMPLE
    .\Modify-ADUserFromIAM.ps1
    Das Skript kann ohne Parameter gestartet werden
.EXAMPLE
    .\Modify-ADUserFromIAM.ps1 -UpdateDbatools
    Prüft vor dem eigentlichen Ablauf, ob Updates für das dbatools-Modul vorliegen und installiert diese ggf.
.EXAMPLE
    .\Modify-ADUserFromIAM.ps1 -RetryFailed -ClearLogs
    Aufträge der letzten 2 Tage, die nicht erfolgreich ausgeführt wurden (mit success=0 in der Datenbank markiert), werden erneut ausgeführt. Zusätzlich werden Log-Dateien älter als 30min entfernt.
    Nütlzicher Aufruf beim Testen von neuen Datenbank-Einträgen
.EXAMPLE
    .\Modify-ADUserFromIAM.ps1 -SQLInstance "MY-BRAND-NEW-SQL-MACHINE\SpecialInstance" -SQLDatabase "[newIAM_DB]" -SQLTable "[newIAM_DB].[dbo].[newTableName]" -SQLLogTable "[newIAM_DB].[dbo].[newLogTableName]"
    Sollte die dazugehörige Datenbank einmal umziehen, so können zum Testen die neuen Einstellungen übergeben werden. Nach der Migration sollten der Einfachheit und Übersichtlichkeit halber die Default-Werte im Skript angepasst werden
.OUTPUTS
    Bildschirm- und Dateiausgabe der einzelnen Arbeitsschritte. Zusammenfassung erfolgt auch als Log-Eintrag in eine SQL-Tabelle
.NOTES
    Author:      thmueller@cancom.de

    Version:     0.7
    Date:        March 27, 2018
    Email Handling per Auftragstabelle, Löschen von AD-Accounts implementiert

History:
    Version:     0.6
    Date:        March 10, 2018
    Errorhandling verbessert

    Version:     0.5
    Date:        November 8, 2018
    SQL-Spaltennamen konfigurierbar gemacht, Errorhandling und Logging verbessert, Funktionsverbesserungen
 
    Version:     0.4
    Date:        November 7, 2018
    Viel Kosmetik, Gruppen-Management für Cross-Domain optimiert, Replikations-Wartezeiten optimiert und konfigurierbar gestaltet, Mailbenachrichtigung angepasst

    Version:     0.3
    Date:        November 6, 2018
    Deaktivierung nun inkl. Lync und Exchange (ungetestet)

    Version:     0.2
    Date:        November 5, 2018
    Verbessertes Gruppen-Management, Deaktivierung optimiert

    Version:     0.1
    Date:        September 25-28, 2018
    Initial Commit

    Anpassung Lync wegen neuer Version:  sfb2015.ferchau.local 11.08.2020 GHP
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
    [string]$LyncFQDN =  "sfb2019.ferchau.local", #"gmzrzapp061.ferchau.local",
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "\\ferchau.local\fileservice\IAM",
    [Parameter(Mandatory = $false)]
    [string]$DeactivatedOU = "OU=Deaktiviert,OU=Benutzermanagement",
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
#$IsAdmin=(new-object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
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

$query = "SELECT " + $column['id'] + "," + $column['newvalue'] + "," + $column['attribute'] + "," + $column['userDN'] + "," + $column['emailtext'] + "," + $column['emailrecipient'] + "," + $column['emailsubject'] + " FROM $SQLTable WHERE [AD_Attribut] in ('Entfernen','Deaktivieren') and verarbeitet IS NULL"

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

    if($commands.($column['attribute']).Contains("Deaktivieren") -or $commands.($column['attribute']).Contains("Entfernen")){
        Write-Color -Text "Pfad für Konfigurations-Export prüfen: ",$ExportPath -Color Yellow,White
        if(!(Test-Path -Path $ExportPath)){
            New-Item $ExportPath -ItemType Directory -ErrorAction Stop | Out-Null
            Write-Color -Text $ExportPath," wurde erstellt" -Color Cyan,White
        }

        Write-Color -Text "Remote-Sessions zu Exchange und Lync etablieren" -Color Yellow
        $remotesessions = $true
        $exc_session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $ExchangeFQDN + "/PowerShell/") -Authentication Kerberos -ErrorAction Stop
        $result = Import-PSSession $exc_session -ErrorAction SilentlyContinue
        $lync_session = New-PSSession -ConnectionURI ("https://" + $LyncFQDN + "/OcsPowerShell") -Authentication negotiatewithimplicitcredential  -ErrorAction Stop 
        $result = Import-PSSession $lync_session -ErrorAction SilentlyContinue 
    }
    
    foreach ($c in $commands){
        $change_running++
        $currentcommand = $c
        $attributeToChange = $($c.($column['attribute']))
        Write-Color -Text "Datensatz: ",$($c.($column['UserDN']))," - ",$($c.($column['attribute'])) -Color White,Cyan,White,Cyan
        try{
            $Domain = ($($c.($column['UserDN'])) -split "," | Where-Object {$_ -like "DC=*"}) -join "." -replace ("DC=","")
            $DCforDomain = [string](Get-ADDomainController -DomainName $Domain -Discover).hostname
            $ADUser = $null
            #$ADUser = Get-AdUser -Filter { distinguishedname -eq $($c.($column['UserDN'])) } -Server $Domain -Properties *
            $ADuser = Get-ADUser $($c.($column['UserDN'])) -Server $Domain -Properties *
            if($ADUser -eq $null){
                Write-Color -Text $($c.($column['UserDN']))," nicht gefunden" -Color Cyan,Red
                $q = "UPDATE $SQLTable SET " + $column['success'] + " = 0, " + $column['completed'] + " = GETDATE() WHERE ID = '" + $($c.($column['id'])) + "'"
                $result = Invoke-DbaQuery -SqlInstance $SQLInstance -Query $q -Database $SQLDatabase
                $change_fail++
                continue
            }else{
                Write-Color -Text "Gefunden: ",$ADuser.SamAccountName," in Domain ",$Domain -Color Cyan,White,Cyan
            }
            switch ($attributeToChange){
                "Group" {
                    Write-Color -Text "(",$change_running,"/",$change_total,") Auftrag: ",$ADUser.Name," (",$ADUser.Samaccountname,") ","soll in Gruppe ",$($c.($column['newValue'])) -Color White,Cyan,White,White,White,Red,White,Cyan,White,White,Cyan -StartTab 1 -LinesAfter 1
                    $GroupDomain = ($($c.($column['newValue'])) -split "," | Where-Object {$_ -like "DC=*"}) -join "." -replace ("DC=","")
                    $replacingValue = $($c.($column['newValue']))
                    $executionCommand = "Get-AdGroup `"" + $replacingValue + "`" -Server " + $GroupDomain + " | Add-ADGroupMember -Members `$ADUser -ErrorAction SilentlyContinue" 
                    $checkcommand = "(Get-ADUser `"" + $ADUser.DistinguishedName + "`" -Properties MemberOf -Server " + $Domain + " | Select-Object -ExpandProperty Memberof | where-Object { `$_ -like `"*" + $replacingValue + "*`"})"
                    break;
                }
                "RemoveGroup"{
                    Write-Color -Text "(",$change_running,"/",$change_total,") Auftrag: ",$ADUser.Name," (",$ADUser.Samaccountname,") ","entfernen aus Gruppe ",$($c.($column['newValue'])) -Color White,Cyan,White,White,White,Red,White,Cyan,White,White,Cyan -StartTab 1 -LinesAfter 1
                    $GroupDomain = ($($c.($column['newValue'])) -split "," | Where-Object {$_ -like "DC=*"}) -join "." -replace ("DC=","")
                    $replacingValue = $null
                    $executionCommand = "Get-AdGroup `"" + $($c.($column['newValue'])) + "`" -Server " + $GroupDomain + " | Remove-AdgroupMember -Members `$ADUser -Confirm:`$false -ErrorAction SilentlyContinue"
                    $checkcommand = "(Get-ADUser `"" + $ADUser.DistinguishedName + "`" -Properties MemberOf -Server " + $Domain + " | Select-Object -ExpandProperty Memberof | where-Object { `$_ -like `"*" + $($c.($column['newValue'])) + "*`"})"
                    break;
                }
                {"Entfernen","Deaktivieren" -contains $_ } {
                    Write-Color -Text "(",$change_running,"/",$change_total,") Auftrag: ",$ADUser.Name," (",$ADUser.Samaccountname,") ","soll deaktiviert werden" -Color White,Cyan,White,White,White,Red,White,Cyan,White,White -StartTab 1 -LinesAfter 1
                
                    if($ADUser."msRTCSIP-PrimaryUserAddress" -or $ADuser.msExchRecipientTypeDetails){
                        $userexportpath = New-Item -Path $ExportPath -Name ($ADuser.EmployeeID + "_" + $ADuser.UserPrincipalName) -ItemType Directory -Force -ErrorAction Stop
                        Write-Color -Text $ADUser.Samaccountname,"Verzeichnis in ",$ExportPath," für Exports erstelllt" -Color Cyan,White,Cyan
                    }

                    #Lync
                    if($ADUser."msRTCSIP-PrimaryUserAddress"){
                        Write-Color -Text $ADUser.Samaccountname," ist Lync aktiviert."," Starte Export" -Color Cyan,White,Yellow -StartTab 1
                        $ADuser | Select-Object name,displayname,samaccountname,distinguishedname,msrtc* | Export-Clixml -Path ($userexportpath.FullName + "\" + $ADuser.EmployeeID + "_" + $ADuser.UserPrincipalName + "_lync_ad.xml") -ErrorAction Stop
    
                        #Export-CsUserData -FileName ($userexportpath.FullName + "\" + $ADuser.EmployeeID + "_" + $ADuser.Samaccountname + "_lync.zip") -poolFqdn $LyncFQDN -UserFilter ($ADUser."msRTCSIP-PrimaryUserAddress" -replace "sip:","") -DomainController $DCforDomain
                        Get-CsUser -Identity $ADUser.DistinguishedName -DomainController $DCforDomain | Export-Clixml -Path ($userexportpath.FullName + "\" + $ADuser.EmployeeID + "_" + $ADuser.UserPrincipalName + "_lync_csuser.xml") -ErrorAction Stop
                        Set-CsUser -Identity $ADUser.DistinguishedName -enabled $false -DomainController $DCforDomain
                        Disable-CsUser -Identity $ADUser.DistinguishedName -DomainController $DCforDomain -Confirm:$false 
                        $um = $true
                    }else{
                        $um = $false
                        Write-Color -Text $ADUser.Samaccountname," ist nicht Lync-aktiviert." -Color Cyan,White -StartTab 1
                    }

                    #Exchange
                    if($ADuser.msExchRecipientTypeDetails){
                        Write-Color -Text $ADUser.Samaccountname," hat ein Exchange-Postfach."," Starte Export, prüfe erneut in 120s" -Color Cyan,White,Yellow -StartTab 1
                    
                        $exp_req = New-MailboxExportRequest -Mailbox $ADuser.DistinguishedName -BadItemLimit 1000 -LargeItemLimit 1000 -AcceptLargeDataLoss -FilePath ($userexportpath.FullName+ "\" + $ADuser.EmployeeID + "_" + $ADuser.UserPrincipalName + ".pst") -DomainController $DCforDomain -WarningAction SilentlyContinue  -ErrorAction Stop
                        Resume-MailboxExportRequest $exp_req.Identity -DomainController $DCforDomain
                        Start-Sleep -Seconds 60 | Out-Null
                        While((Get-MailboxExportRequest $exp_req.Identity -DomainController $DCforDomain).Status -notin "CompletedWithWarning","Completed","Failed"){
                            Write-Color -Text "PST-Export noch nicht beendet:",(" " + ((Get-MailboxExportRequestStatistics $exp_req.Identity -DomainController $DCforDomain).PercentComplete).ToString()),"% fertig" -Color White,Cyan,White -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 30 | Out-Null
                        }
                        Get-MailboxExportRequestStatistics $exp_req.Identity -DomainController $DCforDomain -includeReport | Select-Object * | Out-File -FilePath ($userexportpath.FullName + "\" + $ADuser.EmployeeID + "_" + $ADuser.UserPrincipalName + "_pstexport.txt") -Encoding utf8 -ErrorAction SilentlyContinue  
                        Get-MailboxExportRequest $exp_req.Identity -DomainController $DCforDomain | Remove-MailboxExportRequest -Confirm:$false -DomainController $DCforDomain
                        if($um){
                            Disable-UMMailbox -Identity $ADuser.DistinguishedName -Confirm:$false -DomainController $DCforDomain
                        }
                        Disable-Mailbox -Identity $ADuser.DistinguishedName -Confirm:$false -DomainController $DCforDomain
                        Set-Mailbox $ADUser.Manager -EmailAddresses @{add=(($ADUser.proxyAddresses | Where-Object {$_ -match "smtp"}).ToLower() -replace "stmp:","" )} -DomainController $DCforDomain

                    }else{
                        Write-Color -Text $ADUser.Samaccountname," hat kein Exchange-Postfach." -Color Cyan,White -StartTab 1
                    }

                    $replacingValue = $true
                    $DeactivatedOU += "," + (($aduser.DistinguishedName -split "," | Where-Object {$_ -like "DC=*"}) -join ",")
                
                    #User deaktivieren
                    $executionCommand = "Get-AdUser `"" + $ADUser.DistinguishedName + "`" -Server " + $Domain + " | Set-AdUser -enabled `$false ; "
                    #User aus allen Gruppen entfernen
                    $executionCommand += "Start-Sleep -Seconds 3 ; Get-AdUser `"" + $ADUser.DistinguishedName + "`" -Server " + $Domain + " -Properties MemberOf -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Memberof | ForEach-Object { Remove-ADGroupMember -Confirm:`$false -Identity `$_ -Member `"" + $ADUser.DistinguishedName + "`" -ErrorAction SilentlyContinue} ; "
                    #User Attribute entfernen
                    $executionCommand += "Start-Sleep -Seconds 3 ; Set-AdUser `"" + $ADUser.DistinguishedName + "`" -Server " + $Domain + " -Clear Department,Description,wwwHomepage,manager,mobile,telephoneNumber,facsimileTelephoneNumber,thumbnailphoto"
                    #User in andere OU schieben
                    $executionCommand += "Start-Sleep -Seconds 3 ; Get-AdUser `"" + $ADUser.DistinguishedName + "`" -Server " + $Domain + " | Move-ADObject -TargetPath `"" + $DeactivatedOU + "`""
                    #neuen DN ermitteln
                    $newDistinguishedname = ($ADUser.DistinguishedName).Substring(0,($ADUser.DistinguishedName).IndexOf("OU=")) + $DeactivatedOU
                    $checkcommand = "[bool] ([bool]!(Get-AdUser -Filter { distinguishedname -eq `"" + $ADUser.DistinguishedName + "`"} -Server " + $Domain + ") -and !((Get-AdUser -Filter { distinguishedname -eq `"" + $newDistinguishedname + "`"} -Server " + $Domain + ").enabled) -and (((Get-AdUser -Filter { distinguishedname -eq `"" + $newDistinguishedname + "`"} -Server " + $Domain + " -Properties memberof).memberof).count -eq 0))"
                    #break;
                }
                "Entfernen" {
                    $replacingValue = $true 
                    ##Sanity Check
                    Write-Color -Text $ADUser.Samaccountname," wird entfernt." -Color Cyan,White 
                    if(!(Get-ADObject -Filter * -SearchBase $ADUser.DistinguishedName -Server $Domain).count){
                        $executionCommand = "Remove-ADUser -Identity `"" + $ADUser.DistinguishedName + "`" -Server " + $Domain + " -Confirm:`$false"
                    }else{
                        $executionCommand = "Remove-ADObject -Identity `"" + $ADUser.DistinguishedName + "`" -Server " + $Domain + " -Recursive -Confirm:`$false"
                    }
                    $checkcommand = "([bool]!(Get-AdUser -Filter { distinguishedname -eq `"" + $ADUser.DistinguishedName + "`"} -Server " + $Domain + "))"

                    break;
                }
                default {
                    $replacingValue = $($c.($column['newValue']))
                    Write-Color -Text "(",$change_running,"/",$change_total,") Auftrag: ",$ADUser.Name," (",$ADUser.Samaccountname,") ","Attribut ",$attributeToChange," ändern: ",$replacingValue -Color White,Cyan,White,White,White,Red,White,Cyan,White,White,Cyan,White,Cyan -StartTab 1 -LinesAfter 1
                    $executionCommand = "Get-AdUser `"" + $ADUser.DistinguishedName + "`" -Server " + $Domain + " | Set-AdUser -replace @{" + $attributeToChange + "=`"" + $replacingValue + "`"}" + " -Server " + $Domain + " -ErrorAction SilentlyContinue"
                    $checkcommand = "(Get-AdUser `"" + $ADUser.DistinguishedName + "`" -Server " + $Domain + " -Properties " + $attributeToChange + ")." + $attributeToChange
                    break;
                }
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
        $lync_session | Remove-PSSession -ErrorAction SilentlyContinue
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