<#
.SYNOPSIS
    Script for setting agenda rights within an O365 environment.
    Prerequisite is that there has to be made a mail enabled Security Group within Exchange Online first.
    This group has as member an (also newly created) Universal Security Group wich contains the users where the rights are assigned to.
    
    Put the newly created mail enabled Security Group within add/remove/set -mailboxpermission section of the script.         
.VERSION 
    1.0
.AUTHOR 
    Bart Tacken - Client ICT Groep
.PREREQUISITES
    PowerShell v3
    Login account for Exchange Online
    Encrypted Key and Credential file for unattended use     
.EXAMPLE
    .\Set-AgendaRightsFunction.ps1 -credpath "c:\beheer\scripts\key\saMSPstackcreds.xml" -KeyFilePath "C:\beheer\scripts\key\ClientExchangeOnline.key" -TargetUserAgenda demo03@test.nl -AccessRightsRole Reviewer -NeedAccessMB demo02@test.nl -Action Add
#>
[CmdletBinding()]
param (
        [Parameter(Mandatory=$False)] # O365 specific value, Path to XML file that includes customer Office 365 Service Account credentials
        [string]$CredPath, # "c:\beheer\scripts\key\saMSPstackcreds.xml"

        [Parameter(Mandatory=$False)] # 0365 specific value, Path to AES key file that can de-crypt the XML file containing Office 365 SA credentials
        [string]$KeyFilePath, # "C:\beheer\scripts\key\ClientExchangeOnline.key"
        
        [Parameter(Mandatory=$False)] # UPN of user whose agenda needs to be changes
        [string]$TargetUserAgenda, # "test@test.nl"

        [Parameter(Mandatory=$False)] # User name or security group name of user whose agenda needs to be changes
        [string]$TargetGroupAgenda, # "engineers-deta"
        
        [Parameter(Mandatory=$True)] # O365 Agenda access rights
        [string]$AccessRightsRole, # Reviewer (read only), Editor (read/write)
        
        [Parameter(Mandatory=$True)] # Email address of mailbox that need accessO365 Agenda access rights
        [string]$NeedAccessMB, # ,        

        [Parameter(Mandatory=$False)] # Action to be taken (Add, Remove)
        [string]$Action # 

    ) # End Param
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
[string]$DateStr = (Get-Date).ToString("s").Replace(":","-") # +"_" # Easy sortable date string    
Import-Module ActiveDirectory
#----------------------------------------------------------[Functions]----------------------------------------------------------
Function Connect-EXOnline {
    param($Credentials)
    $URL = "https://outlook.office365.com/powershell-liveid/"     
    #$URL = "https://ps.outlook.com/powershell"     
    $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $Credentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
        Import-PSSession $EXOSession
}
###################################################################################################################################
$Key = Get-Content $KeyFilePath
$credXML = Import-Clixml $CredPath #Import encrypted credential file into XML format
$secureStringPWD = ConvertTo-SecureString -String $credXML.Password -Key $key
$Credentials = New-Object System.Management.Automation.PsCredential($credXML.UserName, $secureStringPWD) # Create PScredential Object
$ErrorActionPreference = 'SilentlyContinue'
$TargetSGusers = Get-ADGroupMember $TargetGroupAgenda | select -ExpandProperty samaccountname

get-pssession | remove-pssession
Connect-EXOnline -Credentials $Credentials
Start-Transcript ('c:\windows\temp\' + "$Datestr" + '_Set_Agenda_Rights') # Start logging

# Show rights of target Agenda before taking action and check is target user is already set.
If ($TargetUserAgenda -notlike $null) {
    Write-Output "Current Agenda permissions of target agenda:"
    Get-MailboxFolderPermission -Identity ("$TargetUserAgenda" + ":\Agenda")

    If ($Action -like "Add") {
        add-MailboxFolderPermission -Identity ("$TargetUserAgenda" + ":\Agenda") -User $NeedAccessMB -AccessRights $AccessRightsRole -Confirm:$False #-whatif # werkt
    }
    If ($Action -like "Remove") {
        Remove-MailboxFolderPermission -Identity ("$TargetUserAgenda" + ":\Agenda") -User $NeedAccessMB -Confirm:$False #-whatif # werkt
    }

    Write-Output "Current Agenda permissions of target agenda after change:"
    Get-MailboxFolderPermission -Identity ("$TargetUserAgenda" + ":\Agenda")

} # End If TargetUserAgenda -notlike $Null

If ($TargetGroupAgenda -notlike $null) {
    ForEach ($MB in $TargetSGusers) { # Go through each MailBox in the target user group
        #$users = get-MailboxFolderPermission -Identity ("$MB" + ":\Agenda") | select -ExpandProperty user
      #  ForEach ($user in $users) {

            Write-Output "Current Agenda permissions of target agenda:"
            Get-MailboxFolderPermission -Identity ("$MB" + ":\Agenda")

            If ($Action -like "Add") {
                add-MailboxFolderPermission -Identity ("$MB" + ":\Agenda") -User $NeedAccessMB -AccessRights $AccessRightsRole -Confirm:$False -ea silentlycontinue #-whatif # werkt
            }
            If ($Action -like "Remove") {
                Remove-MailboxFolderPermission -Identity ("$MB" + ":\Agenda") -User $NeedAccessMB -Confirm:$False #-whatif # werkt   
            }
       # } # End ForEach
    } # End ForEach
    Write-Output "Current Agenda permissions of target agenda after change:"
    Get-MailboxFolderPermission -Identity ("$TargetUserAgenda" + ":\Agenda")

} # End If TargetGroupAgenda -notlike $Null
