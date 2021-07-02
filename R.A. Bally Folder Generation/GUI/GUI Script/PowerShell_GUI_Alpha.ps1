[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$inputXML = @'
<Window x:Class="PowerShell_GUI.Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShell_GUI"
        mc:Ignorable="d"
        Title="Window1" Height="450" Width="656" Background="#FF252526">
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
        <TextBox x:Name="FirstName_Text" HorizontalAlignment="Left" Margin="10,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="109" Height="18"/>
        <Label x:Name="FirstName_Label" Content="First Name" HorizontalAlignment="Left" Margin="31,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="67"/>
        <Label x:Name="Middle_Label" Content="Middle" HorizontalAlignment="Left" Margin="112,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="48"/>
        <Label x:Name="LastName_Label" Content="Last Name" HorizontalAlignment="Left" Margin="183,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="66"/>
        <Label x:Name="DateOfBirth_Label" Content="Date Of Birth" HorizontalAlignment="Left" Margin="394,44,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="79"/>
        <Label x:Name="Age_Label" Content="Age" HorizontalAlignment="Left" Margin="478,44,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="31"/>
        <Label x:Name="Spouse_Label" Content="Spouse" HorizontalAlignment="Left" Margin="312,44,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="50"/>
        <Label x:Name="Card_Label" Content="Card?" HorizontalAlignment="Left" Margin="506,44,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="40"/>
        <TextBox x:Name="Middle_Text" HorizontalAlignment="Left" Margin="128,71,0,0" TextWrapping="Wrap" Width="16"/>
        <TextBox x:Name="LastName_Text" HorizontalAlignment="Left" Margin="153,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="126" Height="18"/>
        <TextBox x:Name="Spouse_Text" HorizontalAlignment="Left" Margin="291,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="91" Height="18"/>
        <TextBox x:Name="DateOfBirth_Text" HorizontalAlignment="Left" Margin="393,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="81" Height="18"/>
        <TextBox x:Name="Age_Text" HorizontalAlignment="Left" Margin="483,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="21" Height="18"/>
        <TextBox x:Name="Card_Text" HorizontalAlignment="Left" Margin="513,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="26" Height="18"/>
        <TextBox x:Name="StreetAddress_Text" HorizontalAlignment="Left" Margin="10,0,0,0" Grid.Row="1" TextWrapping="Wrap" Width="270" VerticalAlignment="Center" Height="18"/>
        <TextBox x:Name="City_Text" HorizontalAlignment="Left" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="120" Height="18" Margin="304,0,0,0"/>
        <TextBox x:Name="State_Text" HorizontalAlignment="Left" Margin="443,0,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="29" Height="18"/>
        <TextBox x:Name="ZipCode_Text" HorizontalAlignment="Left" Margin="495,0,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="44" Height="18"/>
        <Label x:Name="StreetAddress_Label" Content="Street Address" HorizontalAlignment="Left" Margin="102,5,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="86"/>
        <Label x:Name="City_Label" Content="City" HorizontalAlignment="Left" Margin="349,5,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="30"/>
        <Label x:Name="State_Label" Content="State" HorizontalAlignment="Left" Margin="440,5,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="36"/>
        <Label x:Name="ZipCode_Label" Content="Zip Code" HorizontalAlignment="Left" Margin="488,5,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="58"/>
        <TextBox x:Name="Companies_Text" HorizontalAlignment="Center" Grid.Row="2" TextWrapping="Wrap" VerticalAlignment="Center" Width="562" Height="18"/>
        <Label x:Name="Companies_Label" Content="Companies On File" HorizontalAlignment="Center" Grid.Row="2" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="110" Margin="0,5,0,0"/>
        <TextBox x:Name="ClientFolderPath_Text" HorizontalAlignment="Left" Grid.Row="3" Text="&#xD;&#xA;" TextWrapping="Wrap" VerticalAlignment="Top" Width="512" Height="18" Margin="11,30,0,0"/>
        <Label x:Name="ClientFolderPath_Label" Content="Client File Path" HorizontalAlignment="Center" Grid.Row="3" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="90" Margin="0,4,0,0"/>
        <Button x:Name="OpenClientFolder_Button" Content="Open File" Margin="527,30,0,0" Grid.Row="3" VerticalAlignment="Top" Height="18"/>
        <Button x:Name="SearchClient_Button" Content="Search Client" HorizontalAlignment="Left" Margin="169,0,0,0" Grid.Row="4" VerticalAlignment="Center" Height="20" Width="73"/>
        <Button x:Name="NewClient_Button" Content="New Client" HorizontalAlignment="Left" Margin="339,0,0,0" Grid.Row="4" VerticalAlignment="Center" Height="20" Width="73"/>
    </Grid>
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
$datasource = "C:\Users\alecmcclure.ADU\Documents\GitHub\PowerShell-Projects\R.A. Bally Folder Generation\SampleDatabase\SampleDatabase.accdb"

$access = Connect-Access -DataSource $datasource

$clients = Select-Access -Connection $access -Table "ClientMasterFile"

function Search-Client ([string]$FirstName)
{
    [array]$client = $client = Invoke-Access -Connection $access -Query "SELECT * FROM ClientMasterFile WHERE ([FirstName] = `'$FirstName`');" | select -Property FirstName,Middle,LastName,Spouse,Age,Card,StreetAddress1,City,State,Zipcode,Companies,ClientFolderPath  
    $WPFFirstName_Text.Text = $client.FirstName
    $WPFLastName_Text.Text = $client.LastName
    $WPFMiddle_Text.Text = $client.Middle
    $WPFSpouse_Text.Text = $client.Spouse
    $WPFAge_Text.Text = $client.Age
    $WPFCard_Text.Text = $client.Card
    $WPFStreetAddress_Text.Text = $client.StreetAddress1
    $WPFCity_Text.Text = $client.City
    $WPFState_Text.Text = $client.State
    $WPFZipCode_Text.Text = $client.Zipcode
    $WPFCompanies_Text.Text = $client.Companies
    $script:WPFClientFolderPath_Text.Text = $client.ClientFolderPath
}

function Set-work($please)
{
    while ($WPFClientFolderPath_Text.Text -contains $null)
    {
        continue
    }

    Start-Process ($please)
}

function Open-Folder
{  
    Invoke-item $WPFClientFolderPath_Text.Text
}    


$WPFSearchClient_Button.Add_Click({Search-Client -FirstName $WPFFirstName_Text.Text})

$WPFOpenClientFolder_Button.Add_Click({Open-Folder})


#===========================================================================
# Shows the form
#===========================================================================

#write-host "To show the form, run the following" -ForegroundColor Cyan

$Form.ShowDialog() | Out-Null

Disconnect-Access -Connection $access
