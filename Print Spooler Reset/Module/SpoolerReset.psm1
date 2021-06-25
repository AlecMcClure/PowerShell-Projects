#--------------------------------Functions------------------------------------
#This is the main function to be used as the Module
function Reset-Spooler 
{
    param (
        #-------------------------------Computer Name---------------------------------
        # This Variable asks for the Computer Name that needs the Print Spooler Reset. 
    
        [string]$ComputerName
    
        #-----------------------------------------------------------------------------        
    )
    
    <#--------------------------------Functions------------------------------------
    This function tests the connection to the specified computer and makes sure
    that the computer account exists in Active Directory. The output of this 
    function is a boolean value of True or False
    #>
    
    <#-------------------------------Credentials-------------------------------------
    This Variable gets the Admin credentials to run the script and reset the spooler
    on the remote computer. If running this script to target your local machine, 
    press the Esc key or exit the Credentials Box that pops up. 
    #>
echo ' ','Please enter Administrator Credentials ',' ', 'Or if running on Local Computer close out of Credentials Box', ' '

$credential = Get-Credential -Message 'Enter Administrator Credentials : Domain\Username'

#---------------------------------------------------------------------------------

    function connection_test 
    {
        Test-NetConnection -ComputerName (Get-ADComputer -Identity $ComputerName | select -ExpandProperty Name @{n='computername';e={$_.Name}}) | select -ExpandProperty PingSucceeded
    }

    <# 
    This function checks if the Print Spooler on the specified computer is either
    running or stopped. The output of this function is a String value containing
    'Running' or 'Stopped'
    #>
    function is_serviceRunning
    {
        gsv spooler -ComputerName $ComputerName | select -ExpandProperty Status
    }
    #-----------------------------------------------------------------------------

    <#-----------------------------Test Connection---------------------------------
    This code block calls the Connection Test function and outputs whether the
    connection was successful or not.
    #>

    echo ' ' , '------------------------------------Testing Connection------------------------------------------' , ' '

    $connection_Status = try {connection_test} catch {echo 'Computer Not Found. Breaking...', ' ', $Error[0] ; break}

    if ($connection_Status -eq $true)
    {
    echo 'Connection Successful!'
    }

    if ($connection_Status -eq $false)
    {
    echo ' ' , 'Could not establish connection. Breaking...' , ' ' ,$Error[0]
    break
    }

    echo ' ' , '------------------------------------------------------------------------------------------------' ,  ' '

    #-----------------------------------------------------------------------------

    <#-------------------------------Get Spooler-----------------------------------
    Here we try to get the spooler if the connection is successful and store it 
    in a variable.
    #>

    try
    {
        $spooler = gwmi -ComputerName $ComputerName -Class win32_service -Filter "Name = 'spooler'" -Credential $credential -ErrorAction Stop
    }
    catch [System.Management.ManagementException]
    {
        echo 'AuthenticationError: If using this script on a local machine, close out of the Credentials Pop Up box when it appears.', ' '
        break
    }

    <#-------------------------------Reset Spooler----------------------------------
    This code block prompts if you want to reset the spooler and then stops it 
    on the remote computer and then starts it again or breaks the program.
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
            break
        }
    }

    if ($user_response -eq 'n')
    {
        echo ' ','Breaking...',' '
        break
    }


    if ($user_response -ne 'y' -or 'n')
    {
        echo 'Please enter y or n.'
        continue
    }
    }

    #-----------------------------------------------------------------------------
}