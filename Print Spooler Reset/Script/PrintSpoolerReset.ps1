<#
---------------------------------Author--------------------------------------
                                
                               Alec McClure

-------------------------------------------------------------------------------#>

<#
.SYNOPSIS
 This script is designed to reset the print spooler on remote machines.

.DESCRIPTION
-------------------------------Description-----------------------------------
 When you run the script it will prompt you to enter a computer name that
 it will check against computer accounts in Active Directory. If it finds
 the specified computer account it will then ask you if you want to reset
 the print spooler and then reset it. Once it is reset Have the user try 
 to print again and hopefully it should work.
-------------------------------------------------------------------------------#>

<#-------------------------------Credentials-------------------------------------
This Variable gets the Admin credentials to run the script and reset the spooler
on the remote computer. If running this script to target your local machine, 
press the Esc key or exit the Credentials Box that pops up. 
#>
echo ' ','Please enter Administrator Credentials ',' ', 'Or if running on Local Computer close out of Credentials Box', ' '

$credential = Get-Credential -Message 'Enter Administrator Credentials : Domain\Username'


#-------------------------------Computer Name---------------------------------
# This Variable asks for the Computer Name that needs the Print Spooler Reset.



$user_Prompt = Read-Host "Enter Computer Name or type Show Window and press enter"

$Computer_Search = $null

if ($user_Prompt -eq 'Show Window')
    {
        $user_computerSearch_Window =  Get-ADComputer -Filter * | select -ExpandProperty Name | Out-GridView -Title 'Select a Computer' -OutputMode Single 
        $Computer_Search = $user_computerSearch_Window

    }
else 
    {
        $user_computerSearch_Inline = $user_Prompt
        $Computer_Search = $user_computerSearch_Inline
    }

#Read-Host "Enter the Computer Name"

#-----------------------------------------------------------------------------

<#--------------------------------Functions------------------------------------
 This function tests the connection to the specified computer and makes sure
 that the computer account exists in Active Directory. The output of this 
 function is a boolean value of True or False
#>
function connection_test 
    {
        Test-NetConnection -ComputerName (Get-ADComputer -Identity $Computer_Search | select -ExpandProperty Name @{n='computername';e={$_.Name}}) | select -ExpandProperty PingSucceeded
    }

<# 
 This function checks if the Print Spooler on the specified computer is either
 running or stopped. The output of this function is a String value containing
 'Running' or 'Stopped'
#>
function is_serviceRunning
    {
        gsv spooler -ComputerName $Computer_Search | select -ExpandProperty Status
    }
#-----------------------------------------------------------------------------

<#-----------------------------Test Connection---------------------------------
 This code block calls the Connection Test function and outputs whether the
 connection was successful or not.
 #>

echo ' ' , '------------------------------------Testing Connection------------------------------------------' , ' '

$connection_Status = try {connection_test} catch {echo 'Computer Not Found. Exiting...', ' ', $Error[0] ; exit}

if ($connection_Status -eq $true)
{
    echo 'Connection Successful!'
}

if ($connection_Status -eq $false)
{
    echo ' ' , 'Could not establish connection. Exiting...' , ' ' ,$Error[0]
    exit
}

echo ' ' , '------------------------------------------------------------------------------------------------' ,  ' '

#-----------------------------------------------------------------------------

<#-------------------------------Get Spooler-----------------------------------
 Here we try to get the spooler if the connection is successful and store it 
 in a variable.
 #>

try
    {
        $spooler = gwmi -ComputerName $Computer_Search -Class win32_service -Filter "Name = 'spooler'" -Credential $credential -ErrorAction Stop
    }
catch [System.Management.ManagementException]
    {
        echo 'AuthenticationError: If using this script on a local machine, close out of the Credentials Pop Up box when it appears.', ' '
        exit
    }

<#-------------------------------Reset Spooler----------------------------------
 This code block prompts if you want to reset the spooler and then stops it 
 on the remote computer and then starts it again or exits the program.
#>

$while_var = 1

while ($while_var = 1) 
{
    
    $user_response = Read-Host 'Would You Like to reset the Print Spooler? [y] or [n] press enter.'

    if ($user_response -eq 'y')
    {
    
        $spooler.stopservice() | Out-Null; echo (' ','Stopping Print Spooler', ' ')

        sleep 1
    
        $check1 = is_serviceRunning 
    
        if ($check1 -eq 'running')
        {
             echo 'Print Spooler could not be stopped'
             break
        }

        sleep 2
    
        $spooler.startservice() | Out-Null; echo (' ', 'Starting Print Spooler', ' ')

        sleep 1
    
        $check2 = is_serviceRunning
    }   

    if ($check1 -eq 'stopped')
    {
        if ($check2 -eq 'running')
        {
            echo ('Print Spooler Restarted Successfully!', ' ')
            exit
        }
    }

    if ($user_response -eq 'n')
    {
        echo ' ','Exiting...',' '
        exit
    }


    if ($user_response -ne 'y' -or 'n')
    {
        echo 'Please enter y or n.'
        continue
    }
}

#-----------------------------------------------------------------------------