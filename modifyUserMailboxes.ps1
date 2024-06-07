#Purpose: This program is used to modify the members lists of a few mailboxes in AD
# !! personalized mailboxes have been removed, and replaced with a the “$targetEmail” placeholder !!

#Author: Sam Kelly
#Date: 4/29/2024

 
#FUNCTION DEFINITIONS

function AddUserToSharedMailbox
{
	#PARAMETER DEFINITIONS
	
param(
    	#this parameter accepts a user's email address
    	[Parameter(Mandatory=$true, Position=0)]
    	[string]$UserEmail
	)

	
	  Add-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess
    Add-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess
    Add-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess
    Add-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess
	  Add-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess
    Add-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess
    
}
 
function RemoveUserFromSharedMailbox
{
	#PARAMETER DEFINITIONS
	param(
    	#this parameter accepts a user's email address to remove from the six mailboxes
    	[Parameter(Mandatory=$true, Position=0)]
    	[string]$UserEmail
	)
	
    Remove-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess -Confirm:$FALSE -InheritanceType All
    Remove-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess -Confirm:$FALSE -InheritanceType All
	  Remove-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess -Confirm:$FALSE -InheritanceType All
    Remove-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess -Confirm:$FALSE -InheritanceType All
    Remove-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess -Confirm:$FALSE -InheritanceType All
    Remove-MailboxPermission -Identity $targetEmail -User $UserEmail -AccessRights FullAccess -Confirm:$FALSE -InheritanceType All
    
}
 
function VerifyMailboxPermissions
{
	param(
    	#The first parameter accepts a user's email address to add to the six mailboxes
    	[Parameter(Mandatory=$true, Position=0)]
    	[string]$UserEmail
	)
    
    Get-Mailbox $targetEmail | Get-MailboxPermission | Where-Object { $_.AccessRights -like '*FullAccess*' } | Select-Object Identity, User, AccessRights | where user -like $UserEmail
    Get-Mailbox $targetEmail | Get-MailboxPermission | Where-Object { $_.AccessRights -like '*FullAccess*' } | Select-Object Identity, User, AccessRights | where user -like $UserEmail
    
    Get-Mailbox $targetEmail | Get-MailboxPermission | Where-Object { $_.AccessRights -like '*FullAccess*' } | Select-Object Identity, User, AccessRights | where user -like $UserEmail
    Get-Mailbox $targetEmail | Get-MailboxPermission | Where-Object { $_.AccessRights -like '*FullAccess*' } | Select-Object Identity, User, AccessRights | where user -like $UserEmail
    
    Get-Mailbox $targetEmail | Get-MailboxPermission | Where-Object { $_.AccessRights -like '*FullAccess*' } | Select-Object Identity, User, AccessRights | where user -like $UserEmail
    Get-Mailbox $targetEmail | Get-MailboxPermission | Where-Object { $_.AccessRights -like '*FullAccess*' } | Select-Object Identity, User, AccessRights | where user -like $UserEmail
}
 
#display welcome banner
function WelcomeMessage
{
	Write-Host "Welcome!"
	Write-Host " "
	Write-Host "This Program is Used to Update the Mailbox Members For the Following Shared Mailboxes:"
	Write-Host " "
	Write-Host " "
	Write-Host " - $targetEmail
	Write-Host " - $targetEmail
	Write-Host " "
	Write-Host " - $targetEmail
	Write-Host " - $targetEmail
	Write-Host " "
	Write-Host " - $targetEmail
	Write-Host " - $targetEmail
	Write-Host " "
	Write-Host "_____________________________________________________________________________________ "
}
 
function ContinueOn {
	Read-Host "Press Enter to Continue..."
	Clear-Host
}
 
#MAIN
try{
	#initialize isConnected variable
	$isConnected=$TRUE
	while($isConnected){
    	try{
        	#connect to IPPSSession
            Connect-ExchangeOnline
        	$isConnected=$FALSE
    	}catch{}
    
    	do
    	{
        	WelcomeMessage
        	Write-Host " "
        	Write-Host "If you would like to update the members list, please make a selection below."
        	Write-Host " "

        	#display menu
          Write-Host "(1) Add User"
        	Write-Host "(2) View User Permissions"
          Write-Host "(3) Remove User"
          Write-Host "(4) Exit"
        	Write-Host " "
        	Write-Host " "

        	#get users input
        	$UserSelection = Read-Host "Enter Selection"

        	switch ($UserSelection){
            	"1"{
                	#ask user to enter the full user email address
                	$UserAddress = Read-Host -Prompt "Please enter your user's full email address"
                	
                  #call the add mailbox permissions function, and pass it the UserAddress value
                  AddUserToSharedMailbox -UserEmail $UserAddress
                	ContinueOn
            	}
            	"2"{
                	#ask user to enter the full user email address
                	$UserAddress = Read-Host -Prompt "Please enter your user's full email address"
                	
                  #call the verify permissions function, and pass it the UserAddress value
                  VerifyMailboxPermissions -UserEmail $UserAddress
                  ContinueOn
            	}
            	"3"{
                	#ask user to enter the full user email address
                	$UserAddress = Read-Host -Prompt "Please enter your user's full email address"
                	
                  #call the remove mailbox permissions function, and pass it the UserAddress value
                  RemoveUserFromSharedMailbox -UserEmail $UserAddress
                  ContinueOn
            	}
            	"4"{
                  Write-Host " "
                	Write-Host "Closing Connection..."
                  Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
            	}
                default {
                	Write-Host " "                  
                	Write-Host "Invalid selection, please try again."
                	Write-Host " "
                	ContinueOn
            	}
        	}
        
    	} while ($UserSelection -ne "4")
	}
}
catch{
    Write-Host "Error: Closing Connection. Please Try Again."
}
