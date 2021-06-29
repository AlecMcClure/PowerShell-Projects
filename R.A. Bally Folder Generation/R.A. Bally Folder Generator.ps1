#<You must use install-module accesscmdlets in an administrator Powershell to run this script >#

$datasource = "C:\Users\alecmcclure.adu\Documents\Programming\Database\ClientDatabase1.accdb"

$Clients_FilePath = "C:\Users\alecmcclure.adu\Documents\Programming\Clients\"

$access = Connect-Access -DataSource "$datasource"

[System.Collections.ArrayList]$clients = Select-Access -Connection $access -Table "ClientMasterFile"
[System.Collections.ArrayList]$Empty_Path = Invoke-Access -Connection $access -Query "SELECT * FROM ClientMasterFile WHERE (ClientFolderPath Is Null);"

function Sync-Status 
{
    $Check_Clients = Invoke-Access -Connection $access -Query "SELECT * FROM ClientMasterFile WHERE (ClientFolderPath Is Null);"

    foreach ($client in $Check_clients)
    {
        if ($client.ClientFolderPath -notlike '*')
        {
            echo "Client Path not synced. Syncing..."

            Update-Access -Connection $access -Table "ClientMasterFile" -Columns "ClientFolderPath" -Values $Client_Folder_Path_Spouse -Where "EntryNumber = $Client_ID"
        }
    }
}

function Start-Sleep($seconds) 
{
    $doneDT = (Get-Date).AddSeconds($seconds)
    while($doneDT -gt (Get-Date)) 
    {
        $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
        $percent = ($seconds - $secondsLeft) / $seconds * 100
        Write-Progress -Activity "Syncing Database" -Status "Syncing..." -PercentComplete $percent
        [System.Threading.Thread]::Sleep(500)
    }
    Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining 0 -Completed
}

echo ' ' ,"Creating Client Folders...", ' '

foreach ($client in $Empty_Path)
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

foreach ($clientp in $Empty_Path) 
{
    if ($clientp.Spouse -eq 'n/a')
    {
        $Client_IDP = $clientp | select -ExpandProperty EntryNumber

        $FolderName_NoSpouse = $clientp.LastName + ', ' + $clientp.FirstName

        New-Item -ErrorAction Ignore -Path ($Clients_FilePath + $FolderName_NoSpouse) -ItemType Directory

        $Client_Folder_Path_NoSpouse = $Clients_FilePath + $FolderName_NoSpouse

        $clientp.ClientFolderPath = $Client_Folder_Path_NoSpouse

        Update-Access -Connection $access -Table "ClientMasterFile" -Columns "ClientFolderPath" -Values $Client_Folder_Path_NoSpouse -Where "EntryNumber = $Client_IDP"
    }    
}

cls; echo "Waiting for Database to Sync Changes.";Start-Sleep -seconds 60

cls; echo ' ',"Creating Company Folders...",' '

foreach ($clientd in $clients) 
{
    if ($clientd.Spouse -eq 'n/a')
    {
        if ($clientd.Companies -ne 'n/a')
        {
            $client_companies_NoSpouse = $clientd.Companies -split "\,"
            #$Client_FP = $client.ClientFolderPath

            foreach ($company in $client_companies_NoSpouse) 
            {
                $Client_Name = $clientd.FirstName + $clientd.LastName
                #echo ' ',"Found $company In Client File With No Spouse, Creating Folder called $company at $Client_FP\$company",' '
                $client_name_NoSP = $clientd.FirstName + " " + $clientd.LastName
                try 
                {
                    New-Item -Path $clientd.ClientFolderPath -Name "$company" -ItemType Directory 
                    echo "$company Folder Created for $client_name_NoSP"  
                }
                catch 
                {
                    echo "Sync Status of $client_Name_NoSP Failed. Resyncing..."
                    Start-Sleep -seconds 5
                    New-Item -Path $clientd.ClientFolderPath -Name "$company" -ItemType Directory 
                    echo "$company Folder Created for $client_name_NoSP"  
                }
            }  
        }
    }
}

foreach ($clientf in $clients) 
{
    if ($clientf.Spouse -ne 'n/a')
    {
        if ($clientf.Companies -ne 'n/a')
        {
            $client_companies_Spouse = $clientf.Companies -split "\,"
            $Client_FP = $clientf.ClientFolderPath
            $Client_IDF = $clientf | select -ExpandProperty EntryNumber

            foreach ($company in $client_companies_Spouse) 
            {
                #echo ' ',"Found $company In Client File With No Spouse, Creating Folder called $company at $Client_FP\$company",' '
                $Client_Name = $clientf.FirstName + $clientf.LastName
                try 
                {
                    New-Item -ErrorAction Ignore -Path $clientf.ClientFolderPath -Name "$company" -ItemType Directory
                    echo "$company Folder Created for $client_name"  
                }
                catch 
                {
                    echo "Sync Status of $client_Name Failed. Resyncing..."
                    Update-Access -Connection $access -Table "ClientMasterFile" -Columns "ClientFolderPath" -Values $Client_FP -Where "EntryNumber = $Client_IDF"
                    Start-Sleep -seconds 5
                } 
            }  
        }
    }
}

echo ' ', "Database Sync Complete!", ' '

Disconnect-Access -Connection $access