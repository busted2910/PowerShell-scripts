############################################################################
##                                                                        ##
## Date: June 10. 2024                                                    ##
## Author: PBU                                                            ##
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

#Connect-AzureAD
Connect-ExchangeOnline
Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All"

Import-Module Microsoft.Graph.Groups

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
        Write-host "Mailbox $Mailbox is already a Member of $GroupName" -f Green
    }
    Else
    {
        #Add the mailbox to the group
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $Mailbox
        Write-host -f Yellow "Mailbox $Mailbox is now added to $GroupName"
    }

    #Check if the mailbox account is enabled
    
    #Get Azure Account
    $MailboxAzureAccount = Get-MgUser -UserId $Mailbox -Property AccountEnabled | select AccountEnabled

    #Check if account is enabled
    If($MailboxAzureAccount.AccountEnabled -eq $True)
    {
        #Disable mailbox account
        Write-host "$Mailbox is enabled. It will now be disabled" -f Yellow
        Update-Mguser -UserId $Mailbox -AccountEnabled:$false
    }
    Else
    {
        Write-host -f Green "Mailbox $Mailbox is already disabled"
    }
}

#Make sure the group anly contains the right types of mailboxes
ForEach ($GroupMember in $GroupMembers)
{
    #Check if the member of the group is the correct mailbox type
    If($AllMailboxes -notcontains $GroupMember)
    {
        Write-host "Mailbox $GroupMember is a member of $GroupName and will be removed" -f Yellow
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $GroupMember
    }
    Else
    {
        #Add the user to the group
        Write-host -f Green "Mailbox $GroupMember is a member of $GroupName and is a $mailboxtype"
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
        Write-host "Mailbox $Mailbox is already a Member of $GroupName" -f Green
    }
    Else
    {
        #Add the mailbox to the group
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $Mailbox
        Write-host -f Yellow "Mailbox $Mailbox is now added to $GroupName"
    }

    #Check if the mailbox account is enabled
    
    #Get Azure Account
    $MailboxAzureAccount = Get-MgUser -UserId $Mailbox -Property AccountEnabled | select AccountEnabled

    #Check if account is enabled
    If($MailboxAzureAccount.AccountEnabled -eq $True)
    {
        #Disable mailbox account
        Write-host "$Mailbox is enabled. It will now be disabled" -f Yellow
        Update-Mguser -UserId $Mailbox -AccountEnabled:$false
    }
    Else
    {
        Write-host -f Green "Mailbox $Mailbox is already disabled"
    }
}

#Make sure the group anly contains the right types of mailboxes
ForEach ($GroupMember in $GroupMembers)
{
    #Check if the member of the group is the correct mailbox type
    If($AllMailboxes -notcontains $GroupMember)
    {
        Write-host "Mailbox $GroupMember is a member of $GroupName and will be removed" -f Yellow
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $GroupMember
    }
    Else
    {
        #Add the user to the group
        Write-host -f Green "Mailbox $GroupMember is a member of $GroupName and is a $mailboxtype"
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
        Write-host "Mailbox $Mailbox is already a Member of $GroupName" -f Green
    }
    Else
    {
        #Add the mailbox to the group
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $Mailbox
        Write-host -f Yellow "Mailbox $Mailbox is now added to $GroupName"
    }

    #Check if the mailbox account is enabled
    
    #Get Azure Account
    $MailboxAzureAccount = Get-MgUser -UserId $Mailbox -Property AccountEnabled | select AccountEnabled

    #Check if account is enabled
    If($MailboxAzureAccount.AccountEnabled -eq $True)
    {
        #Disable mailbox account
        Write-host "$Mailbox is enabled. It will now be disabled" -f Yellow
        Update-Mguser -UserId $Mailbox -AccountEnabled:$false
    }
    Else
    {
        Write-host -f Green "Mailbox $Mailbox is already disabled"
    }
}

#Make sure the group anly contains the right types of mailboxes
ForEach ($GroupMember in $GroupMembers)
{
    #Check if the member of the group is the correct mailbox type
    If($AllMailboxes -notcontains $GroupMember)
    {
        Write-host "Mailbox $GroupMember is a member of $GroupName and will be removed" -f Yellow
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $GroupMember
    }
    Else
    {
        #Add the user to the group
        Write-host -f Green "Mailbox $GroupMember is a member of $GroupName and is a $mailboxtype"
    }
}
