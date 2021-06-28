#<You must use install-module accesscmdlets in an administrator Powershell to run this script >#

$datasource = "C:\Users\%userprofile%\Documents\Programming\Database\ClientDatabase1.accdb"

$Clients_FilePath = "C:\Users\%userprofile%\Documents\Programming\Clients\"

$access = Connect-Access -DataSource "$datasource"

[System.Collections.ArrayList]$clients = Select-Access -Connection $access -Table "ClientMasterFile"

foreach ($client in $clients)
{
    if ($client.Spouse -ne 'n/a')
    {   
        $Client_ID = $client | select -ExpandProperty EntryNumber
        
        $FolderName_Spouse = $client.LastName + ', ' + $client.FirstName + ' & ' + $client.Spouse
        
        New-Item -ErrorAction Ignore -Path $Clients_FilePath -Name $FolderName_Spouse -ItemType Directory
        
        $Client_Folder_Path_Spouse = $Clients_FilePath + $FolderName_Spouse

        $client.ClientFolderPath = $Client_Folder_Path_Spouse

        Update-Access -Connection $access -Table "ClientMasterFile" -Columns "ClientFolderPath" -Values $Client_Folder_Path_Spouse -Where "EntryNumber = $Client_ID"
    }
}

foreach ($client in $clients) 
{
    if ($client.Spouse -eq 'n/a')
    {
        $Client_ID = $client | select -ExpandProperty EntryNumber

        $FolderName_NoSpouse = $client.LastName + ', ' + $client.FirstName

        New-Item -ErrorAction Ignore -Path ($Clients_FilePath + $FolderName_NoSpouse) -ItemType Directory

        $Client_Folder_Path_NoSpouse = $Clients_FilePath + $FolderName_NoSpouse

        $client.ClientFolderPath = $Client_Folder_Path_NoSpouse

        Update-Access -Connection $access -Table "ClientMasterFile" -Columns "ClientFolderPath" -Values $Client_Folder_Path_NoSpouse -Where "EntryNumber = $Client_ID"
    }    
}

Sleep 5

foreach ($client in $clients) 
{
    if ($client.Spouse -eq 'n/a')
    {
        if ($client.Companies -ne 'n/a')
        {
            $client_companies_NoSpouse = $client.Companies -split "\,"
            $Client_FP = $client.ClientFolderPath

            foreach ($company in $client_companies_NoSpouse) 
            {
                echo ' ',"Found $company In Client File With No Spouse, Creating Folder called $company at $Client_FP\$company",' '
                New-Item -ErrorAction Ignore -Path $client.ClientFolderPath -Name "$company" -ItemType Directory 
            }  
        }
    }
}

foreach ($client in $clients) 
{
    if ($client.Spouse -ne 'n/a')
    {
        if ($client.Companies -ne 'n/a')
        {
            $client_companies_Spouse = $client.Companies -split "\,"
            $Client_FP = $client.ClientFolderPath

            foreach ($company in $client_companies_Spouse) 
            {
                echo ' ',"Found $company In Client File With No Spouse, Creating Folder called $company at $Client_FP\$company",' '
                New-Item -ErrorAction Ignore -Path $client.ClientFolderPath -Name "$company" -ItemType Directory 
            }  
        }
    }
}

Disconnect-Access -Connection $access