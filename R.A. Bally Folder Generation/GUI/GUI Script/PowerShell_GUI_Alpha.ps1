[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$inputXML = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PowerShell GUI" Height="429" Width="909" Background="#FF252526">
    <Grid x:Name="Grid_Grid" HorizontalAlignment="Center" VerticalAlignment="Stretch">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto" MinHeight="80"/>
            <RowDefinition Height="Auto" MinHeight="80"/>
            <RowDefinition Height="Auto" MinHeight="60"/>
            <RowDefinition Height="Auto" MinHeight="50"/>
            <RowDefinition Height="Auto" MinHeight="20"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="FirstName_Text" HorizontalAlignment="Left" Margin="10,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="110" Height="18"/>
        <Label x:Name="FirstName_Label" Content="First Name" HorizontalAlignment="Left" Margin="31,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="67"/>
        <Label x:Name="Middle_Label" Content="Middle" HorizontalAlignment="Left" Margin="140,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="47"/>
        <Label x:Name="LastName_Label" Content="Last Name" HorizontalAlignment="Left" Margin="227,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="66"/>
        <Label x:Name="DateOfBirth_Label" Content="Date Of Birth" HorizontalAlignment="Left" Margin="525,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="79"/>
        <Label x:Name="Age_Label" Content="Age" HorizontalAlignment="Left" Margin="624,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="31"/>
        <Label x:Name="Spouse_Label" Content="Spouse" HorizontalAlignment="Left" Margin="401,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="49"/>
        <Label x:Name="Card_Label" Content="Card?" HorizontalAlignment="Left" Margin="677,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="40"/>
        <TextBox x:Name="Middle_Text" HorizontalAlignment="Left" Margin="155,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="16" Height="18"/>
        <TextBox x:Name="LastName_Text" HorizontalAlignment="Left" Margin="197,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="125" Height="18"/>
        <TextBox x:Name="Spouse_Text" HorizontalAlignment="Left" Margin="380,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="90" Height="18"/>
        <TextBox x:Name="DateOfBirth_Text" HorizontalAlignment="Left" Margin="525,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="80" Height="18"/>
        <TextBox x:Name="Age_Text" HorizontalAlignment="Left" Margin="630,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="20" Height="18"/>
        <TextBox x:Name="Card_Text" HorizontalAlignment="Left" Margin="685,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="25" Height="18"/>
        <TextBox x:Name="StreetAddress_Text" HorizontalAlignment="Left" Margin="52,0,0,0" Grid.Row="1" TextWrapping="Wrap" Width="270" VerticalAlignment="Center" Height="18"/>
        <TextBox x:Name="City_Text" HorizontalAlignment="Left" Margin="358,0,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="120" Height="18"/>
        <TextBox x:Name="State_Text" HorizontalAlignment="Left" Margin="556,0,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="18" Height="18"/>
        <TextBox x:Name="ZipCode_Text" HorizontalAlignment="Left" Margin="656,0,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="44" Height="18"/>
        <Label x:Name="StreetAddress_Label" Content="Street Address" HorizontalAlignment="Left" Margin="144,4,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="86"/>
        <Label x:Name="City_Label" Content="City" HorizontalAlignment="Left" Margin="403,4,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="30"/>
        <Label x:Name="State_Label" Content="State" HorizontalAlignment="Left" Margin="546,4,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="37"/>
        <Label x:Name="ZipCode_Label" Content="Zip Code" HorizontalAlignment="Left" Margin="649,4,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="58"/>
        <TextBox x:Name="Companies_Text" HorizontalAlignment="Center" Grid.Row="2" TextWrapping="Wrap" VerticalAlignment="Center" Width="512" Height="18"/>
        <Label x:Name="Companies_Label" Content="Companies On File" HorizontalAlignment="Center" Grid.Row="2" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="110"/>
        <TextBox x:Name="ClientFolderPath_Text" HorizontalAlignment="Center" Grid.Row="3" Text="&#xD;&#xA;" TextWrapping="Wrap" VerticalAlignment="Top" Width="445" Height="18" Margin="0,30,0,0"/>
        <Label x:Name="ClientFolderPath_Label" Content="Client File Path" HorizontalAlignment="Center" Grid.Row="3" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="89" Margin="0,4,0,0"/>
        <Button x:Name="OpenClientFolder_Button" Content="Open File" HorizontalAlignment="Left" Margin="621,28,0,0" Grid.Row="3" VerticalAlignment="Top" Height="20" Width="57"/>
        <Button x:Name="SearchClient_Button" Content="Search Client" HorizontalAlignment="Left" Margin="220,0,0,0" Grid.Row="4" VerticalAlignment="Center" Height="20" Width="73"/>
        <Button x:Name="NewClient_Button" Content="New Client" HorizontalAlignment="Left" Margin="401,0,0,0" Grid.Row="4" VerticalAlignment="Center" Height="20" Width="73"/>

    </Grid>
</Window>
'@
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read XAML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
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

#===========================================================================
# Use this space to add code to the various form elements in your GUI
#===========================================================================
$datasource = "C:\Users\alecmcclure.adu\Documents\Programming\Database\ClientDatabase1.accdb"

$access = Connect-Access -DataSource $datasource

$clients = Select-Access -Connection $access -Table "ClientMasterFile"

$FirstName_Text = $WPFFirstName_Text.Text

function Search-Client ([string]$FirstName)
{
    $WPFLastName_Text.Text = Invoke-Access -Connection $access -Query "SELECT ClientMasterFile.[LastName] FROM ClientMasterFile WHERE ([FirstName]= `'$FirstName`');" | select -ExpandProperty LastName
}

$WPFSearchClient_Button.Add_Click({Search-Client -FirstName $WPFFirstName_Text.Text})                                                                



#===========================================================================
# Shows the form
#===========================================================================

write-host "To show the form, run the following" -ForegroundColor Cyan

#echo $WPFLastName_Text.Text

$Form.ShowDialog() | out-null