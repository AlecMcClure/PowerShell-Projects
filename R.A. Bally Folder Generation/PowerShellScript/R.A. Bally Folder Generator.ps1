#<You must use install-module accesscmdlets in an administrator Powershell to run this script >#

$datasource = "C:\PowerShell Testing\Database\SampleDatabase.accdb"

$Clients_FilePath = "C:\PowerShell Testing\Clients\"

$access = Connect-Access -DataSource "$datasource"

[System.Collections.ArrayList]$Empty_Path = Invoke-Access -Connection $access -Query "SELECT * FROM ClientMasterFile WHERE (ClientFolderPath Is Null);"

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

foreach ($client in $Empty_Path) 
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

Disconnect-Access -Connection $access

echo "Waiting for Database to Sync Changes.";Start-Sleep -seconds 30

$access_Sync = Connect-Access -DataSource "$datasource"

[System.Collections.ArrayList]$clients = Select-Access -Connection $access_Sync -Table "ClientMasterFile"

echo ' ',"Creating Company Folders...",' '

foreach ($client in $clients) 
{
    if ($client.Spouse -eq 'n/a')
    {
        if ($client.Companies -ne 'n/a')
        {
            $client_companies_NoSpouse = $client.Companies -split "\,"

            foreach ($company in $client_companies_NoSpouse) 
            {
                $Client_Name = $client.FirstName + $client.LastName
                $client_name_NoSP = $client.FirstName + " " + $client.LastName
                $Client_ID = $client | select -ExpandProperty EntryNumber
                try 
                {
                    New-Item -ErrorAction Ignore -Path $client.ClientFolderPath -Name "$company" -ItemType Directory  
                }
                catch 
                {
                    echo "Sync Status of $client_Name_NoSP Failed. Resyncing..."
                    Start-Sleep -seconds 5
                    Update-Access -Connection $access -Table "ClientMasterFile" -Columns "ClientFolderPath" -Values $Client_Folder_Path_NoSpouse -Where "EntryNumber = $Client_ID" 
                }
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
            $Client_ID = $client | select -ExpandProperty EntryNumber

            foreach ($company in $client_companies_Spouse) 
            {
                $Client_Name = $client.FirstName + $client.LastName
                try 
                {
                    New-Item -ErrorAction Ignore -Path $client.ClientFolderPath -Name "$company" -ItemType Directory 
                }
                catch 
                {
                    echo "Sync Status of $client_Name Failed. Resyncing..."
                    Start-Sleep -seconds 5
                    Update-Access -Connection $access -Table "ClientMasterFile" -Columns "ClientFolderPath" -Values $Client_FP -Where "EntryNumber = $Client_ID" 
                } 
            }  
        }
    }
}

echo ' ', "Database Sync Complete!", ' '

Disconnect-Access -Connection $access_Sync