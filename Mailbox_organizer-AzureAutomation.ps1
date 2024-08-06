############################################################################
##                                                                        ##
## Date: June 10. 2024                                                    ##
## Author: Peter Busted                                                   ##
##                                                                        ##
## Add all Room mailboxes to "Room mailboxes" AD Group                    ##
## Disable all Room mailboxes                                             ##
## Delete non Room mailboxes from "Room mailboxes" AD Group               ##
##                                                                        ##
## Add Shared mailboxes to "Shared mailboxes" AD Group                    ##
## Disable all Shared mailboxes                                           ##
## Delete non Shared mailboxes from "Shared mailboxes" AD Group           ##
##                                                                        ##
## Add Scheduling mailboxes to "Scheduling mailboxes" AD Group            ##
## Disable all Scheduling mailboxes                                       ##
## Delete non Scheduling mailboxes from "Scheduling mailboxes" AD Group   ##
##                                                                        ##
##                                                                        ##
## Change $GroupName and $mailboxtype if adding support for a new         ##
## mailbox type                                                           ##
##                                                                        ##
##                                                                        ##
##                                                                        ##
##                                                                        ##
##                                                                        ##
############################################################################

#Connect to ExchangeOnline and MgGraph

try
{
    "Logging in to Graph..."
    Connect-MgGraph -Identity
}
catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
}

try
{
    "Logging in to Exchange Online..."
    $organization = "xxx.onmicrosoft.com"
    Connect-ExchangeOnline -ManagedIdentity -Organization $organization

}
catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
}

#Import modules
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Users

##################
# Room mailboxes #
##################

#Mailbox type
$mailboxtype = "RoomMailbox"

#Name of group
$GroupName = "Room Mailboxes"
 
#Get Group Info 
$Group = Get-MgGroup -Search "DisplayName:$GroupName" -ConsistencyLevel:eventual
 
#Get Exisiting Members of the Group
$GroupMembers = Get-MgGroupMember -GroupId $Group.Id  | Select -ExpandProperty Id

#Get all Mailboxes
$AllMailboxes = Get-Mailbox -RecipientTypeDetails $mailboxtype -ResultSize:Unlimited | Select -ExpandProperty ExternalDirectoryObjectId

#Add all mailboxes to the group
ForEach ($Mailbox in $AllMailboxes)
{
    #Check if the mailbox is in the group
    If($GroupMembers -contains $Mailbox)
    {
        Write-Output "Mailbox $Mailbox is already a Member of $GroupName"
    }
    Else
    {
        #Add the mailbox to the group
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $Mailbox
        Write-Output "Mailbox $Mailbox is now added to $GroupName"
    }

    #Check if the mailbox account is enabled
    
    #Get Azure Account
    $MailboxAzureAccount = Get-MgUser -UserId $Mailbox -Property AccountEnabled | select AccountEnabled

    #Check if account is enabled
    If($MailboxAzureAccount.AccountEnabled -eq $True)
    {
        #Disable mailbox account
        Write-Output "$Mailbox is enabled. It will now be disabled"
        Update-Mguser -UserId $Mailbox -AccountEnabled:$false
    }
    Else
    {
        Write-Output "Mailbox $Mailbox is already disabled"
    }
}

#Make sure the group anly contains the right types of mailboxes
ForEach ($GroupMember in $GroupMembers)
{
    #Check if the member of the group is the correct mailbox type
    If($AllMailboxes -notcontains $GroupMember)
    {
        Write-Output "Mailbox $GroupMember is a member of $GroupName and will be removed"
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $GroupMember
    }
    Else
    {
        #Add the user to the group
        Write-Output "Mailbox $GroupMember is a member of $GroupName and is a $mailboxtype"
    }
}

####################
# Shared mailboxes #
####################

#Mailbox type
$mailboxtype = "SharedMailbox"

#Name of group
$GroupName = "Shared Mailboxes"
 
#Get Group Info 
$Group = Get-MgGroup -Search "DisplayName:$GroupName" -ConsistencyLevel:eventual
 
#Get Exisiting Members of the Group
$GroupMembers = Get-MgGroupMember -GroupId $Group.Id  | Select -ExpandProperty Id

#Get all Mailboxes
$AllMailboxes = Get-Mailbox -RecipientTypeDetails $mailboxtype -ResultSize:Unlimited | Select -ExpandProperty ExternalDirectoryObjectId

#Add all mailboxes to the group
ForEach ($Mailbox in $AllMailboxes)
{
    #Check if the mailbox is in the group
    If($GroupMembers -contains $Mailbox)
    {
        Write-Output "Mailbox $Mailbox is already a Member of $GroupName"
    }
    Else
    {
        #Add the mailbox to the group
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $Mailbox
        Write-Output "Mailbox $Mailbox is now added to $GroupName"
    }

    #Check if the mailbox account is enabled
    
    #Get Azure Account
    $MailboxAzureAccount = Get-MgUser -UserId $Mailbox -Property AccountEnabled | select AccountEnabled

    #Check if account is enabled
    If($MailboxAzureAccount.AccountEnabled -eq $True)
    {
        #Disable mailbox account
        Write-Output "$Mailbox is enabled. It will now be disabled"
        Update-Mguser -UserId $Mailbox -AccountEnabled:$false
    }
    Else
    {
        Write-Output "Mailbox $Mailbox is already disabled"
    }
}

#Make sure the group anly contains the right types of mailboxes
ForEach ($GroupMember in $GroupMembers)
{
    #Check if the member of the group is the correct mailbox type
    If($AllMailboxes -notcontains $GroupMember)
    {
        Write-Output "Mailbox $GroupMember is a member of $GroupName and will be removed"
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $GroupMember
    }
    Else
    {
        #Add the user to the group
        Write-Output "Mailbox $GroupMember is a member of $GroupName and is a $mailboxtype"
    }
}

########################
# Scheduling mailboxes #
########################

#Mailbox type
$mailboxtype = "SchedulingMailbox"

#Name of group
$GroupName = "Scheduling Mailboxes"
 
#Get Group Info 
$Group = Get-MgGroup -Search "DisplayName:$GroupName" -ConsistencyLevel:eventual
 
#Get Exisiting Members of the Group
$GroupMembers = Get-MgGroupMember -GroupId $Group.Id  | Select -ExpandProperty Id

#Get all Mailboxes
$AllMailboxes = Get-Mailbox -RecipientTypeDetails $mailboxtype -ResultSize:Unlimited | Select -ExpandProperty ExternalDirectoryObjectId

#Add all mailboxes to the group
ForEach ($Mailbox in $AllMailboxes)
{
    #Check if the mailbox is in the group
    If($GroupMembers -contains $Mailbox)
    {
        Write-Output "Mailbox $Mailbox is already a Member of $GroupName"
    }
    Else
    {
        #Add the mailbox to the group
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $Mailbox
        Write-Output "Mailbox $Mailbox is now added to $GroupName"
    }

    #Check if the mailbox account is enabled
    
    #Get Azure Account
    $MailboxAzureAccount = Get-MgUser -UserId $Mailbox -Property AccountEnabled | select AccountEnabled

    #Check if account is enabled
    If($MailboxAzureAccount.AccountEnabled -eq $True)
    {
        #Disable mailbox account
        Write-Output "$Mailbox is enabled. It will now be disabled"
        Update-Mguser -UserId $Mailbox -AccountEnabled:$false
    }
    Else
    {
        Write-Output "Mailbox $Mailbox is already disabled"
    }
}

#Make sure the group anly contains the right types of mailboxes
ForEach ($GroupMember in $GroupMembers)
{
    #Check if the member of the group is the correct mailbox type
    If($AllMailboxes -notcontains $GroupMember)
    {
        Write-Output "Mailbox $GroupMember is a member of $GroupName and will be removed"
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $GroupMember
    }
    Else
    {
        #Add the user to the group
        Write-Output "Mailbox $GroupMember is a member of $GroupName and is a $mailboxtype"
    }
}

