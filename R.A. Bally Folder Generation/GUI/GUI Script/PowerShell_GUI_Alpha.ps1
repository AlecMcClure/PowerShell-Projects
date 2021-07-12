[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$inputXML = @'
<Window x:Class="PowerShell_GUI.Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShell_GUI"
        mc:Ignorable="d"
        Title="R.A. Bally CMS" Height="500" Width="1000" Background="#FF252526" WindowStartupLocation="CenterScreen" WindowState="Maximized">
    <Viewbox Stretch="Uniform">
        <Grid x:Name="Grid_Grid" HorizontalAlignment="Center" VerticalAlignment="Stretch">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto" MinHeight="83"/>
                <RowDefinition Height="Auto" MinHeight="77"/>
                <RowDefinition Height="Auto" MinHeight="60"/>
                <RowDefinition Height="Auto" MinHeight="65"/>
                <RowDefinition Height="Auto" MinHeight="5"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="FirstName_Text" HorizontalAlignment="Center" Margin="52,71,480,0" TextWrapping="Wrap" VerticalAlignment="Center" Width="109" Height="18"/>
            <Label x:Name="FirstName_Label" Content="First Name" HorizontalAlignment="Center" Margin="73,40,495,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="67"/>
            <Label x:Name="Middle_Label" Content="Middle" HorizontalAlignment="Center" Margin="154,40,415,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="48"/>
            <Label x:Name="LastName_Label" Content="Last Name" HorizontalAlignment="Center" Margin="225,40,290,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="66"/>
            <Label x:Name="DateOfBirth_Label" Content="Date Of Birth" HorizontalAlignment="Center" Margin="436,40,60,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="78"/>
            <Label x:Name="Age_Label" Content="Age" HorizontalAlignment="Center" Margin="519,40,10,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="31"/>
            <Label x:Name="Spouse_Label" Content="Spouse" HorizontalAlignment="Center" Margin="393,40,220,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="50"/>
            <Label x:Name="Card_Label" Content="Card?" HorizontalAlignment="Center" Margin="627,41,25,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="24" Width="40"/>
            <TextBox x:Name="Middle_Text" HorizontalAlignment="Center" Margin="170,71,430,0" TextWrapping="Wrap" Width="16" VerticalAlignment="Center"/>
            <TextBox x:Name="LastName_Text" HorizontalAlignment="Center" Margin="195,71,270,0" TextWrapping="Wrap" VerticalAlignment="Center" Width="126" Height="18"/>
            <TextBox x:Name="Spouse_Text" HorizontalAlignment="Center" Margin="335,71,160,0" TextWrapping="Wrap" VerticalAlignment="Center" Width="91" Height="18"/>
            <TextBox x:Name="DateOfBirth_Text" HorizontalAlignment="Center" Margin="435,71,60,0" TextWrapping="Wrap" VerticalAlignment="Center" Width="81" Height="18"/>
            <TextBox x:Name="Age_Text" HorizontalAlignment="Center" Margin="525,71,10,0" TextWrapping="Wrap" VerticalAlignment="Center" Width="21" Height="18"/>
            <TextBox x:Name="Card_Text" HorizontalAlignment="Center" Margin="601,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="26" Height="18"/>
            <TextBox x:Name="StreetAddress_Text" HorizontalAlignment="Center" Margin="10,0,300,0" Grid.Row="1" TextWrapping="Wrap" Width="270" VerticalAlignment="Center" Height="18"/>
            <TextBox x:Name="City_Text" HorizontalAlignment="Center" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="120" Height="18" Margin="305,0,150,0"/>
            <TextBox x:Name="State_Text" HorizontalAlignment="Center" Margin="444,0,90,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="30" Height="18"/>
            <TextBox x:Name="ZipCode_Text" HorizontalAlignment="Center" Margin="490,0,30,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="44" Height="18"/>
            <Label x:Name="StreetAddress_Label" Content="Street Address" HorizontalAlignment="Center" Margin="102,5,400,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="86"/>
            <Label x:Name="City_Label" Content="City" HorizontalAlignment="Center" Margin="350,6,195,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="30"/>
            <Label x:Name="State_Label" Content="State" HorizontalAlignment="Center" Margin="441,5,90,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="36"/>
            <Label x:Name="ZipCode_Label" Content="Zip Code" HorizontalAlignment="Center" Margin="483,5,25,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="58"/>
            <TextBox x:Name="Companies_Text" HorizontalAlignment="Center" Grid.Row="2" TextWrapping="Wrap" VerticalAlignment="Center" Width="562" Height="18"/>
            <Label x:Name="Companies_Label" Content="Companies On File" HorizontalAlignment="Left" Grid.Row="2" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="110" Margin="296,4,0,0"/>
            <TextBox x:Name="ClientFolderPath_Text" HorizontalAlignment="Center" Grid.Row="3" Text="&#xA;" TextWrapping="Wrap" VerticalAlignment="Top" Width="505" Height="18" Margin="11,30,0,0"/>
            <Label x:Name="ClientFolderPath_Label" Content="Client File Path" HorizontalAlignment="Center" Grid.Row="3" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="90" Margin="0,4,0,0"/>
            <Button x:Name="OpenClientFolder_Button" Content="Open File" Margin="624,29,0,0" Grid.Row="3" VerticalAlignment="Top" Height="20" HorizontalAlignment="Left"/>
            <Button x:Name="SearchClient_Button" Content="Search Client" HorizontalAlignment="Center" Margin="63,16,500,0" VerticalAlignment="Top" Height="20" Width="76"/>
            <Button x:Name="NewClient_Button" Content="New Client" HorizontalAlignment="Center" VerticalAlignment="Top" Height="20" Margin="414,16,200,0" Width="76"/>
            <Button x:Name="ClearForm_Button" Content="Clear Form" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="176,16,400,0" Width="76"/>
            <Button x:Name="UpdateClient_Button" Content="Update Client" HorizontalAlignment="Center" Margin="293,16,300,0" VerticalAlignment="Top" Width="76"/>
            <TextBox x:Name="EntryNumber_Text" HorizontalAlignment="Center" Margin="10,71,620,83" TextWrapping="Wrap" Width="37" Grid.RowSpan="2" VerticalAlignment="Center"/>
            <Label x:Name="EntryNumber_Label" Content="ID" HorizontalAlignment="Center" Margin="18,42,630,0" VerticalAlignment="Top" Foreground="#FF007ACC"/>
            <DataGrid Name="Datagrid" Grid.Row="4" Background="#FF252526" Foreground="#FF007ACC" AutoGenerateColumns="True" >
                <DataGrid.Columns>
                    <DataGridCheckBoxColumn/>
                    <DataGridTextColumn Header="Entry Number" Binding="{Binding EntryNumber}"/>
                    <DataGridTextColumn Header="First Name" Binding="{Binding FirstName}"/>
                    <DataGridTextColumn Header="Middle" Binding="{Binding Middle}"/>
                    <DataGridTextColumn Header="Last Name" Binding="{Binding LastName}"/>
                    <DataGridTextColumn Header="Spouse" Binding="{Binding Spouse}"/>
                    <DataGridTextColumn Header="Date Of Birth" Binding="{Binding DateOfBirth}"/>
                    <DataGridTextColumn Header="Street Address" Binding="{Binding StreetAddress}"/>
                    <DataGridTextColumn Header="City" Binding="{Binding City}"/>
                    <DataGridTextColumn Header="State" Binding="{Binding State}"/>
                    <DataGridTextColumn Header="Zip Code" Binding="{Binding ZipCode}"/>
                    <DataGridTextColumn Header="Companies" Binding="{Binding Companies}"/>
                </DataGrid.Columns>
            </DataGrid>
            <Button x:Name="DeleteClient_Button" Content="Delete Client" HorizontalAlignment="Left" VerticalAlignment="Top" Height="20" Margin="546,16,100,0" Width="76"/>
        </Grid>
    </Viewbox>
</Window>
'@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $XAML)
    try{
        $Form=[Windows.Markup.XamlReader]::Load( $reader )
    }
    catch{
        Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
        throw
    }
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
  
$xaml.SelectNodes("//*[@Name]") | %{"trying item $($_.Name)";
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
Get-FormVariables

echo "----------------------------------------------------------------------"," "

#===========================================================================
# Use this space to add code to the various form elements in your GUI
#===========================================================================

# Database Connection

[System.Collections.ArrayList]$script:datasource = Import-Csv -Path "C:\PowerShell Testing\Database\ClientMaster.csv"

$Clients_FilePath = "C:\PowerShell Testing\Clients\"

# Functions

function Merge-Updates
{
    $script:datasource | sort { [int]$_.EntryNumber } -Unique |select * | Export-Csv -Path "C:\PowerShell Testing\Database\ClientMaster.csv" -NoTypeInformation -Force
}
function Calculate-Age($person) 
{
    foreach ($client in $person)
    {
        if ($client.DateofBirth -eq "//")
        {
            continue
        }
        else 
        {
        $currentyear = Get-Date | select -ExpandProperty Year
        $dateofbirth = get-date -Date $client.DateOfBirth
        $month = $dateofbirth.Month
        $day = $dateofbirth.Day
    
        $client.birthdaycurrent = 
        @"
        $month/$day/$currentyear
"@
    
        $dateofbirth1 = $client.DateOfBirth
        $birthdaycurrent = $client.birthdaycurrent
    
        $TimeSpan =  [DateTime]$birthdaycurrent - [DateTime]$dateofbirth1;

        $Totaldays = $TimeSpan.TotalDays

        $unformattedage = ($Totaldays/365)

        $age = "{0:n0}" -f $unformattedage
    
        $client.Age = $age
   
        }
    }
    
    Merge-Updates
}

function Search-Client
{
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$EntryNumber
    )

    try{
    [System.Collections.ArrayList]$clients = $client = $datasource | select * | ? {($_.FirstName -eq $FirstName) -or ($_.EntryNumber -eq $EntryNumber) -or ($_.LastName -eq $LastName)}
    }
    catch{
    }

        foreach ($client in $clients)
    {
        $WPFDatagrid.AddChild([pscustomobject]@{EntryNumber = $client.EntryNumber;FirstName = $client.FirstName ; Middle = $client.Middle ; LastName = $client.LastName ; Spouse = $client.Spouse ; DateOfBirth = $client.DateOfBirth; StreetAddress = $client.StreetAddress ; City = $client.City ; State = $client.State ; ZipCode = $client.ZipCode; Companies = $client.Companies})
    }

    $WPFEntryNumber_Text.Text = $client.EntryNumber
    $WPFFirstName_Text.Text = $client.FirstName
    $WPFLastName_Text.Text = $client.LastName
    $WPFMiddle_Text.Text = $client.Middle
    $WPFSpouse_Text.Text = $client.Spouse
    $WPFAge_Text.Text = $client.Age
    $WPFCard_Text.Text = $client.Card
    $WPFStreetAddress_Text.Text = $client.StreetAddress
    $WPFCity_Text.Text = $client.City
    $WPFState_Text.Text = $client.State
    $WPFZipCode_Text.Text = $client.Zipcode
    $WPFCompanies_Text.Text = $client.Companies
    $WPFClientFolderPath_Text.Text = $client.ClientFolderPath
    $WPFDateOfBirth_Text.Text = $client.DateOfBirth    
}   

    
function Open-Folder($folder)
{  
    Invoke-item $folder
}    

function Clear-Form
{
    $WPFEntryNumber_Text.Text = $null
    $WPFFirstName_Text.Text = $null
    $WPFLastName_Text.Text = $null
    $WPFMiddle_Text.Text = $null
    $WPFSpouse_Text.Text = $null
    $WPFAge_Text.Text = $null
    $WPFCard_Text.Text = $null
    $WPFStreetAddress_Text.Text = $null
    $WPFCity_Text.Text = $null
    $WPFState_Text.Text = $null
    $WPFZipCode_Text.Text = $null
    $WPFCompanies_Text.Text = $null
    $WPFClientFolderPath_Text.Text = $null
    $WPFDateOfBirth_Text.Text = $null
    $WPFDataGrid.items.Clear()
}

function Update-Client
{
    if ($WPFSpouse_Text.Text -ne "n/a")
        {
            if ($WPFClientFolderpath_Text.Text -match "&")
            {
                $WPFClientFolderpath_Text.Text = $WPFClientFolderpath_Text.Text
            }
            else
            {
                $Current_Path = $WPFClientFolderpath_Text.Text
            
                $SpouseFolder = " & " + $WPFSpouse_Text.Text
                    
                $New_Path = $Current_Path + $SpouseFolder
            
                Rename-Item $Current_Path $New_Path
            
                $WPFClientFolderpath_Text.Text = $New_Path
            }
        }

    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.FirstName = $WPFFirstName_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.LastName = $WPFLastName_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Middle = $WPFMiddle_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Spouse = $WPFSpouse_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Card = $WPFCard_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.StreetAddress = $WPFStreetAddress_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.City = $WPFCity_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.State = $WPFState_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Zipcode = $WPFZipCode_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Companies = $WPFCompanies_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.DateOfBirth = $WPFDateOfBirth_Text.Text}    
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.ClientFolderPath = $WPFClientFolderPath_Text.Text}

    if ($WPFCompanies_Text.Text -ne 'n/a') 
    {
        if ($WPFSpouse_Text.Text -ne "n/a")
        {
            $client_companies_Spouse = $WPFCompanies_Text.Text -split "\,"

            foreach ($company in $client_companies_Spouse) 
            {                
                New-Item -ErrorAction Ignore -Path $WPFClientFolderpath_Text.Text -Name "$company" -ItemType Directory 
            }
        } 
        if ($WPFSpouse_Text.Text -eq 'n/a')
        {
            $client_companies = $WPFCompanies_Text.Text -split "\,"

            foreach ($company in $client_companies) 
            {
                New-Item -ErrorAction Ignore -Path $Current_Path -Name "$company" -ItemType Directory 
            }
        }
    }

    Merge-Updates

    Clear-Form
}


function New-Client
{    
    if ($WPFSpouse_Text.Text -ne 'n/a')
    {       
        $FolderName_Spouse = $WPFLastName_Text.Text + ', ' + $WPFFirstName_Text.Text + ' & ' + $WPFSpouse_Text.Text
        
        New-Item -ErrorAction Ignore -Path $Clients_FilePath -Name $FolderName_Spouse -ItemType Directory
        
        $Client_Folder_Path_Spouse = $Clients_FilePath + $FolderName_Spouse
    }
    
    if ($WPFSpouse_Text -eq 'n/a')
    {
        $FolderName_NoSpouse = $WPFLastName_Text.Text + ', ' + $WPFFirstName_Text.Text

        New-Item -ErrorAction Ignore -Path ($Clients_FilePath + $FolderName_NoSpouse) -ItemType Directory

        $Client_Folder_Path_NoSpouse = $Clients_FilePath + $FolderName_NoSpouse
    }

    $folderpath = $null

    if ($WPFSpouse_Text.Text -ne 'n/a')
    {
        $folderpath = $Client_Folder_Path_Spouse
    }

    if ($WPFSpouse_Text.Text -eq 'n/a')
    {
        $folderpath = $Client_Folder_Path_NoSpouse
    }

    $entrynumber = $script:datasource.Count + 1

    $newRow = New-Object PsObject -Property @{ EntryNumber = $entrynumber ; FirstName = $WPFFirstName_Text.Text ; Middle = $WPFMiddle_Text.Text ; LastName = $WPFLastName_Text.Text ; Spouse = $WPFSpouse_Text.Text ; DateOfBirth = $WPFDateOfBirth_Text.Text; Age = $null ; Card = $WPFCard_Text.Text ; StreetAddress = $WPFStreetAddress_Text.Text ; City = $WPFCity_Text.Text ; State = $WPFState_Text.Text ; ZipCode = $WPFZipCode_Text.Text ; Companies = $WPFCompanies_Text.Text; ClientFolderPath = $folderpath ; birthdaycurrent = $null}
    
    $script:datasource += $newRow

    Merge-Updates
    
    Calculate-Age($script:datasource)

    Search-Client -EntryNumber $entrynumber

    Update-Client

    Clear-Form
}

function Remove-Client([string]$ClientNumber)
{
    foreach ($client in $datasource)
    { 
        if ($client.EntryNumber -ne $ClientNumber) 
        {
            Continue
        } 
        
        if ($client.EntryNumber -eq $ClientNumber)
        {
            Remove-Item -Path $client.ClientFolderPath

            $datasource.Remove($client)
        }
    }

    Merge-Updates

    Clear-Form

}


# Buttons
$WPFSearchClient_Button.Add_Click({Search-Client -FirstName $WPFFirstName_Text.Text -LastName $WPFLastName_Text.Text -EntryNumber $WPFEntryNumber_Text.Text})

$WPFOpenClientFolder_Button.Add_Click({Open-Folder($WPFClientFolderPath_Text.Text)})

$WPFClearForm_Button.Add_Click({Clear-Form})

$WPFUpdateClient_Button.Add_Click({Update-Client})

$WPFNewClient_Button.Add_Click({New-Client})

$WPFDeleteClient_Button.Add_Click({Remove-Client -ClientNumber $WPFEntryNumber_Text.Text})

$WPFFirstName_Text.Add_KeyDown({if ($_.Key -eq "Enter"){Search-Client -FirstName $WPFFirstName_Text.Text -LastName $WPFLastName_Text.Text -EntryNumber $WPFEntryNumber_Text.Text}})
$WPFLastName_Text.Add_KeyDown({if ($_.Key -eq "Enter"){Search-Client -FirstName $WPFFirstName_Text.Text -LastName $WPFLastName_Text.Text -EntryNumber $WPFEntryNumber_Text.Text}})
$WPFEntryNumber_Text.Add_KeyDown({if ($_.Key -eq "Enter"){Search-Client -EntryNumber $WPFEntryNumber_Text.Text}})

# Do Stuff

#Calculate-Age($script:datasource)

#===========================================================================
# Shows the form
#===========================================================================

#write-host "To show the form, run the following" -ForegroundColor Cyan

$WPFDatagrid.add_SelectionChanged({Search-Client -EntryNumber ($WPFDatagrid.SelectedItem | select -ExpandProperty EntryNumber)})

$WPFDatagrid.IsReadOnly = $true

$Form.ShowDialog() | Out-Null -ErrorAction Ignore
