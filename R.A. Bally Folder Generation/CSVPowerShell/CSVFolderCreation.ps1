$CSVPath = "C:\PowerShell Testing\Database\ClientMaster.csv"

$Clients_FilePath = "C:\PowerShell Testing\Clients\"

$Clients = Import-Csv -Path $CSVPath

$script:Updated_Clients = @()

function Merge-Updates 
{
    $script:Updated_Clients | sort { [int]$_.EntryNumber } -Unique |select * | Export-Csv -Path "C:\PowerShell Testing\Database\ClientMaster.csv" -NoTypeInformation
}

[System.Collections.ArrayList]$Empty_Path = $Clients | Select * | ? ClientFolderPath -EQ ''

foreach ($client in $Empty_Path) 
{
    if ($client.Spouse -ne 'n/a')
    {       
        $FolderName_Spouse = $client.LastName + ', ' + $client.FirstName + ' & ' + $client.Spouse
        
        New-Item -ErrorAction Ignore -Path $Clients_FilePath -Name $FolderName_Spouse -ItemType Directory
        
        $Client_Folder_Path_Spouse = $Clients_FilePath + $FolderName_Spouse

        $client.ClientFolderPath = $Client_Folder_Path_Spouse

        $script:Updated_Clients += $client

        if ($client.Companies -ne 'n/a') 
        {
            $client_companies_Spouse = $client.Companies -split "\,"

            foreach ($company in $client_companies_Spouse) 
            {                
                New-Item -ErrorAction Ignore -Path $client.ClientFolderPath -Name "$company" -ItemType Directory 
            } 
        }  
    }
}

foreach ($client in $Empty_Path) 
{
    if ($client.Spouse -eq 'n/a') 
    {
        $FolderName_NoSpouse = $client.LastName + ', ' + $client.FirstName

        New-Item -ErrorAction Ignore -Path ($Clients_FilePath + $FolderName_NoSpouse) -ItemType Directory

        $Client_Folder_Path_NoSpouse = $Clients_FilePath + $FolderName_NoSpouse

        $client.ClientFolderPath = $Client_Folder_Path_NoSpouse

        $script:Updated_Clients += $client
                
        if ($client.Companies -ne 'n/a') 
        {
            $client_companies_NoSpouse = $client.Companies -split "\,"

            foreach ($company in $client_companies_NoSpouse) 
            {                
                 New-Item -ErrorAction Ignore -Path $client.ClientFolderPath -Name "$company" -ItemType Directory  
            }  
        }
    }
}

Merge-Updates