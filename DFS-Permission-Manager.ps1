<#
.SYNOPSIS
    DFS und Berechtigungs-Manager
.DESCRIPTION
    Tool zur Verwaltung von Ordnern, DFS-Verknüpfungen und NTFS-Berechtigungen
.NOTES
    Erfordert: ActiveDirectory PowerShell-Modul
#>

#Requires -Version 5.1

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ============================================
# KONFIGURATION
# ============================================
$Script:Config = @{
    GroupNameSchema = "FS_{FolderName}_{Permission}"
    GroupDescriptionSchema = "Dateisystem-Berechtigung: {PermissionFull} auf {FullPath}"
    GroupOU = "OU=Dateisystem-Gruppen,OU=Gruppen,DC=domain,DC=local"
    DFSRoot = "\\domain.local\dfs"
    DefaultFileServer = "\\fileserver01"
    
    # Berechtigungstypen mit NTFS-Rechten
    PermissionTypes = @{
        "Lesen"           = @{ NTFSRights = "ReadAndExecute"; Inheritance = "ContainerInherit,ObjectInherit" }
        "Ändern"          = @{ NTFSRights = "Modify"; Inheritance = "ContainerInherit,ObjectInherit" }
        "Vollzugriff"     = @{ NTFSRights = "FullControl"; Inheritance = "ContainerInherit,ObjectInherit" }
        "Auflisten"       = @{ NTFSRights = "ReadAndExecute"; Inheritance = "None" }
        "NurDieserOrdner" = @{ NTFSRights = "Modify"; Inheritance = "None" }
    }
    
    # Kürzel für Berechtigungen (wird in {Permission} eingesetzt)
    PermissionLabels = @{
        "Lesen"           = "RO"
        "Ändern"          = "RW"
        "Vollzugriff"     = "FC"
        "Auflisten"       = "LS"
        "NurDieserOrdner" = "MO"
    }
}

# ============================================
# XAML - Vereinfacht für bessere Kompatibilität
# ============================================
$XamlString = @'
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="DFS und Berechtigungs-Manager" 
    Height="930" Width="1100"
    WindowStartupLocation="CenterScreen"
    Background="#F0F0F0">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="#0078D4" Padding="15,10">
            <TextBlock Text="📁 DFS und Berechtigungs-Manager" FontSize="20" FontWeight="Bold" Foreground="White"/>
        </Border>
        
        <TabControl Grid.Row="1" Margin="10" Name="MainTabControl">
            
            <TabItem Header="📂 1. Ordner erstellen">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <GroupBox Header="Ordner erstellen" Padding="10" Margin="0,0,0,10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="130"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="110"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Basisverzeichnis:" VerticalAlignment="Center" Margin="0,5"/>
                                <TextBox Grid.Row="0" Grid.Column="1" Name="txtBasePath" Margin="5" Padding="5,3"/>
                                <Button Grid.Row="0" Grid.Column="2" Content="📂 Durchsuchen" Name="btnBrowseBase" Margin="5" Padding="5,3"/>
                                
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Neuer Ordnername:" VerticalAlignment="Center" Margin="0,5"/>
                                <TextBox Grid.Row="1" Grid.Column="1" Name="txtNewFolderName" Margin="5" Padding="5,3"/>
                                
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Unterordner:" VerticalAlignment="Center" Margin="0,5"/>
                                <TextBox Grid.Row="2" Grid.Column="1" Name="txtSubfolders" Height="60" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Margin="5" Padding="5,3"/>
                                <TextBlock Grid.Row="2" Grid.Column="2" Text="(pro Zeile einer)" VerticalAlignment="Center" FontSize="10" Foreground="Gray"/>
                                
                                <StackPanel Grid.Row="3" Grid.Column="1" Orientation="Horizontal" Margin="5">
                                    <CheckBox Name="chkCreateDFS" Content="DFS erstellen" VerticalAlignment="Center" Margin="0,0,15,0"/>
                                    <CheckBox Name="chkCreateGroups" Content="Gruppen erstellen" VerticalAlignment="Center" IsChecked="True" Margin="0,0,15,0"/>
                                    <CheckBox Name="chkInheritanceOff" Content="Vererbung deaktivieren" VerticalAlignment="Center" IsChecked="True"/>
                                </StackPanel>
                            </Grid>
                        </GroupBox>
                        
                        <GroupBox Header="DFS-Einstellungen" Padding="10" Margin="0,0,0,10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="130"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="DFS-Namespace:" VerticalAlignment="Center" Margin="0,5"/>
                                <TextBox Grid.Row="0" Grid.Column="1" Name="txtDFSNamespace" Margin="5" Padding="5,3"/>
                                
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="DFS-Pfad:" VerticalAlignment="Center" Margin="0,5"/>
                                <TextBox Grid.Row="1" Grid.Column="1" Name="txtDFSPath" Margin="5" Padding="5,3"/>
                            </Grid>
                        </GroupBox>
                        
                        <GroupBox Header="Gruppen-Einstellungen" Padding="10" Margin="0,0,0,10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="130"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Gruppen-OU:" VerticalAlignment="Center" Margin="0,5"/>
                                <TextBox Grid.Row="0" Grid.Column="1" Name="txtGroupOU" Margin="5" Padding="5,3"/>
                                
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Namensschema:" VerticalAlignment="Center" Margin="0,5"/>
                                <TextBox Grid.Row="1" Grid.Column="1" Name="txtGroupNameSchema" Margin="5" Padding="5,3"/>
                                
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Gruppentypen:" VerticalAlignment="Center" Margin="0,5"/>
                                <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal" Margin="5">
                                    <CheckBox Name="chkGroupRead" Content="Lesen" Margin="0,0,15,0" IsChecked="True"/>
                                    <CheckBox Name="chkGroupModify" Content="Ändern" Margin="0,0,15,0" IsChecked="True"/>
                                    <CheckBox Name="chkGroupFull" Content="Vollzugriff" Margin="0,0,15,0"/>
                                    <CheckBox Name="chkGroupList" Content="Auflisten"/>
                                </StackPanel>
                            </Grid>
                        </GroupBox>
                        
                        <GroupBox Header="Vorschau" Padding="10" Margin="0,0,0,10">
                            <StackPanel>
                                <Button Content="👁️ Vorschau aktualisieren" Name="btnPreview" HorizontalAlignment="Left" Padding="10,5" Margin="0,0,0,10"/>
                                <TextBox Name="txtPreview" Height="120" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" Background="#F8F8F8"/>
                            </StackPanel>
                        </GroupBox>
                        
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button Content="✅ Alles erstellen" Name="btnCreateAll" Padding="20,8" Margin="5" Background="#107C10" Foreground="White" FontWeight="Bold"/>
                            <Button Content="📁 Nur Ordner" Name="btnCreateFolderOnly" Padding="15,8" Margin="5"/>
                            <Button Content="👥 Nur Gruppen" Name="btnCreateGroupsOnly" Padding="15,8" Margin="5"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="🔐 2. Berechtigungen">
                <Grid Margin="15">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <GroupBox Grid.Column="0" Header="Ordnerstruktur" Margin="0,0,5,0" Padding="10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                                <TextBox Name="txtPermissionPath" Width="250" Padding="5,3"/>
                                <Button Content="..." Name="btnBrowsePermPath" Width="35" Margin="5,0"/>
                                <Button Content="Laden" Name="btnLoadTree" Padding="10,3"/>
                            </StackPanel>
                            
                            <TreeView Grid.Row="1" Name="tvFolders"/>
                            
                            <StackPanel Grid.Row="2" Margin="0,10,0,0">
                                <TextBlock Text="Ausgewählt:" FontWeight="Bold"/>
                                <TextBox Name="txtSelectedFolder" IsReadOnly="True" Background="#F8F8F8" Padding="5,3"/>
                            </StackPanel>
                        </Grid>
                    </GroupBox>
                    
                    <GroupBox Grid.Column="1" Header="Berechtigungen" Margin="5,0,0,0" Padding="10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Aktuelle Berechtigungen:" FontWeight="Bold" Margin="0,0,0,5"/>
                            
                            <DataGrid Grid.Row="1" Name="dgPermissions" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="True">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Identität" Binding="{Binding Identity}" Width="*"/>
                                    <DataGridTextColumn Header="Rechte" Binding="{Binding Rights}" Width="100"/>
                                    <DataGridTextColumn Header="Vererbt" Binding="{Binding Inheritance}" Width="60"/>
                                </DataGrid.Columns>
                            </DataGrid>
                            
                            <StackPanel Grid.Row="2" Margin="0,10,0,0">
                                <TextBlock Text="Neue Berechtigung:" FontWeight="Bold" Margin="0,0,0,5"/>
                                <StackPanel Orientation="Horizontal">
                                    <ComboBox Name="cmbPermissionType" Width="150" Padding="5,3">
                                        <ComboBoxItem Content="Lesen"/>
                                        <ComboBoxItem Content="Ändern" IsSelected="True"/>
                                        <ComboBoxItem Content="Vollzugriff"/>
                                        <ComboBoxItem Content="Auflisten"/>
                                    </ComboBox>
                                    <Button Content="✨ Gruppe erstellen" Name="btnAddPermission" Padding="10,5" Margin="10,0,0,0"/>
                                </StackPanel>
                                <CheckBox Name="chkSetTraverse" Content="Traverse für Elternordner setzen" IsChecked="True" Margin="0,10,0,0"/>
                            </StackPanel>
                        </Grid>
                    </GroupBox>
                </Grid>
            </TabItem>
            
            <TabItem Header="👥 3. Benutzer zu Gruppen">
                <Grid Margin="15">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="60"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <GroupBox Grid.Column="0" Header="Benutzer suchen" Padding="10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                                <TextBox Name="txtUserSearch" Width="200" Padding="5,3"/>
                                <Button Content="🔍 Suchen" Name="btnSearchUser" Padding="10,3" Margin="5,0"/>
                            </StackPanel>
                            
                            <DataGrid Grid.Row="1" Name="dgUsers" AutoGenerateColumns="False" CanUserAddRows="False">
                                <DataGrid.Columns>
                                    <DataGridCheckBoxColumn Header="X" Binding="{Binding IsSelected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="30"/>
                                    <DataGridTextColumn Header="Name" Binding="{Binding DisplayName}" Width="*" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="Benutzer" Binding="{Binding SamAccountName}" Width="80" IsReadOnly="True"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </GroupBox>
                    
                    <StackPanel Grid.Column="1" VerticalAlignment="Center">
                        <Button Content="&gt;&gt;" Name="btnAddToGroup" Width="45" Height="30" Margin="5"/>
                        <Button Content="&lt;&lt;" Name="btnRemoveFromGroup" Width="45" Height="30" Margin="5"/>
                    </StackPanel>
                    
                    <GroupBox Grid.Column="2" Header="Gruppen" Padding="10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="120"/>
                            </Grid.RowDefinitions>
                            
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                                <TextBox Name="txtGroupSearch" Width="200" Padding="5,3"/>
                                <Button Content="🔍 Suchen" Name="btnSearchGroup" Padding="10,3" Margin="5,0"/>
                            </StackPanel>
                            
                            <ListBox Grid.Row="1" Name="lbGroups"/>
                            
                            <TextBlock Grid.Row="2" Text="Mitglieder:" FontWeight="Bold" Margin="0,10,0,5"/>
                            
                            <DataGrid Grid.Row="3" Name="dgGroupMembers" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="True">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*"/>
                                    <DataGridTextColumn Header="Typ" Binding="{Binding ObjectClass}" Width="60"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </GroupBox>
                </Grid>
            </TabItem>
            
            <TabItem Header="⚡ 4. Schnellzuweisung">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="⚡ Benutzer schnell auf Ordner berechtigen" FontSize="16" FontWeight="Bold" Margin="0,0,0,15"/>
                        
                        <GroupBox Header="1. Benutzer auswählen" Padding="10" Margin="0,0,0,10">
                            <StackPanel>
                                <StackPanel Orientation="Horizontal">
                                    <TextBox Name="txtQuickUser" Width="300" Padding="5,3"/>
                                    <Button Content="🔍 Suchen" Name="btnQuickSearchUser" Padding="10,3" Margin="10,0"/>
                                </StackPanel>
                                <ListBox Name="lbQuickUsers" Height="80" Margin="0,10,0,0"/>
                            </StackPanel>
                        </GroupBox>
                        
                        <GroupBox Header="2. Ordner auswählen" Padding="10" Margin="0,0,0,10">
                            <StackPanel Orientation="Horizontal">
                                <TextBox Name="txtQuickFolder" Width="400" Padding="5,3"/>
                                <Button Content="📂 Durchsuchen" Name="btnQuickBrowse" Padding="10,3" Margin="10,0"/>
                            </StackPanel>
                        </GroupBox>
                        
                        <GroupBox Header="3. Berechtigung" Padding="10" Margin="0,0,0,10">
                            <StackPanel>
                                <ComboBox Name="cmbQuickPermission" Width="150" HorizontalAlignment="Left" Padding="5,3">
                                    <ComboBoxItem Content="Lesen"/>
                                    <ComboBoxItem Content="Ändern" IsSelected="True"/>
                                    <ComboBoxItem Content="Vollzugriff"/>
                                </ComboBox>
                                <CheckBox Name="chkQuickTraverse" Content="Traverse für Elternordner setzen" IsChecked="True" Margin="0,10,0,0"/>
                                <CheckBox Name="chkQuickCreateGroups" Content="Fehlende Gruppen erstellen" IsChecked="True" Margin="0,5,0,0"/>
                            </StackPanel>
                        </GroupBox>
                        
                        <GroupBox Header="4. Vorschau" Padding="10" Margin="0,0,0,10">
                            <TextBox Name="txtQuickPreview" Height="100" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" Background="#F8F8F8"/>
                        </GroupBox>
                        
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button Content="👁️ Vorschau" Name="btnQuickPreview" Padding="20,8" Margin="5"/>
                            <Button Content="✅ Berechtigung setzen" Name="btnQuickApply" Padding="20,8" Margin="5" Background="#107C10" Foreground="White" FontWeight="Bold"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="⚙️ 5. Einstellungen">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <GroupBox Header="Namensschema" Padding="10" Margin="0,0,0,10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="130"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <TextBlock Grid.Row="0" Grid.ColumnSpan="2" TextWrapping="Wrap" Margin="0,0,0,10">
                                    Platzhalter: {FolderName}, {ParentFolder}, {FullPath}, {Permission} (Kürzel), {PermissionFull} (Volltext)
                                </TextBlock>
                                
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Gruppenname:" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="1" Grid.Column="1" Name="txtSettingsNameSchema" Margin="5" Padding="5,3"/>
                                
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Beschreibung:" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="2" Grid.Column="1" Name="txtSettingsDescSchema" Margin="5" Padding="5,3"/>
                            </Grid>
                        </GroupBox>
                        
                        <GroupBox Header="Berechtigungs-Kürzel" Padding="10" Margin="0,0,0,10">
                            <StackPanel>
                                <TextBlock TextWrapping="Wrap" Margin="0,0,0,10">
                                    Definiere Kürzel für {Permission} Platzhalter. Doppelklick zum Bearbeiten.
                                </TextBlock>
                                <DataGrid Name="dgPermissionLabels" Height="150" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Berechtigung" Binding="{Binding Name}" Width="120" IsReadOnly="True"/>
                                        <DataGridTextColumn Header="Kuerzel" Binding="{Binding Label, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="80"/>
                                        <DataGridTextColumn Header="NTFS-Recht" Binding="{Binding NTFSRights}" Width="*" IsReadOnly="True"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </StackPanel>
                        </GroupBox>
                        
                        <GroupBox Header="Pfade" Padding="10" Margin="0,0,0,10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="130"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Gruppen-OU:" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="0" Grid.Column="1" Name="txtSettingsOU" Margin="5" Padding="5,3"/>
                                
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="DFS-Namespace:" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="1" Grid.Column="1" Name="txtSettingsDFS" Margin="5" Padding="5,3"/>
                                
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Fileserver:" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="2" Grid.Column="1" Name="txtSettingsServer" Margin="5" Padding="5,3"/>
                            </Grid>
                        </GroupBox>
                        
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button Content="💾 Speichern" Name="btnSaveSettings" Padding="20,8" Margin="5"/>
                            <Button Content="🔄 Zurücksetzen" Name="btnResetSettings" Padding="20,8" Margin="5"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
        </TabControl>
        
        <Border Grid.Row="2" Background="#E0E0E0" Padding="10,5">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Name="lblStatus" Text="Bereit"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <TextBlock Text="AD: "/>
                    <Ellipse Name="ellADStatus" Width="12" Height="12" Fill="Gray"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

# ============================================
# GUI LADEN
# ============================================
try {
    [xml]$XAML = $XamlString
    $Reader = New-Object System.Xml.XmlNodeReader $XAML
    $Window = [Windows.Markup.XamlReader]::Load($Reader)
}
catch {
    [System.Windows.MessageBox]::Show("Fehler beim Laden der GUI: $($_.Exception.Message)", "Fehler", "OK", "Error")
    exit
}

# Controls referenzieren - sichere Methode
$Script:UI = @{}
$ControlNames = @(
    'txtBasePath', 'btnBrowseBase', 'txtNewFolderName', 'txtSubfolders',
    'chkCreateDFS', 'chkCreateGroups', 'chkInheritanceOff',
    'txtDFSNamespace', 'txtDFSPath', 'txtGroupOU', 'txtGroupNameSchema',
    'chkGroupRead', 'chkGroupModify', 'chkGroupFull', 'chkGroupList',
    'btnPreview', 'txtPreview', 'btnCreateAll', 'btnCreateFolderOnly', 'btnCreateGroupsOnly',
    'txtPermissionPath', 'btnBrowsePermPath', 'btnLoadTree', 'tvFolders',
    'txtSelectedFolder', 'dgPermissions', 'cmbPermissionType', 'btnAddPermission', 'chkSetTraverse',
    'txtUserSearch', 'btnSearchUser', 'dgUsers',
    'btnAddToGroup', 'btnRemoveFromGroup',
    'txtGroupSearch', 'btnSearchGroup', 'lbGroups', 'dgGroupMembers',
    'txtQuickUser', 'btnQuickSearchUser', 'lbQuickUsers',
    'txtQuickFolder', 'btnQuickBrowse', 'cmbQuickPermission',
    'chkQuickTraverse', 'chkQuickCreateGroups', 'txtQuickPreview',
    'btnQuickPreview', 'btnQuickApply',
    'txtSettingsNameSchema', 'txtSettingsDescSchema', 'txtSettingsOU',
    'txtSettingsDFS', 'txtSettingsServer', 'btnSaveSettings', 'btnResetSettings',
    'dgPermissionLabels',
    'lblStatus', 'ellADStatus'
)

foreach ($Name in $ControlNames) {
    $Script:UI[$Name] = $Window.FindName($Name)
}

# ============================================
# HILFSFUNKTIONEN
# ============================================

function Write-Log {
    param([string]$Message)
    $LogFile = Join-Path $env:TEMP "DFS-Permission-Manager.log"
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "[$Timestamp] $Message"
    $Script:UI.lblStatus.Text = $Message
}

function Show-Message {
    param(
        [string]$Message,
        [string]$Title = "Information",
        [string]$Icon = "Information"
    )
    [System.Windows.MessageBox]::Show($Message, $Title, "OK", $Icon)
}

function Test-ADModule {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $Script:UI.ellADStatus.Fill = [System.Windows.Media.Brushes]::Green
        return $true
    }
    catch {
        $Script:UI.ellADStatus.Fill = [System.Windows.Media.Brushes]::Red
        return $false
    }
}

function Get-CleanName {
    param([string]$Path)
    $Name = Split-Path $Path -Leaf
    $Name = $Name -replace '[\\/:*?"<>|]', '_'
    $Name = $Name -replace '\s+', '_'
    return $Name.Trim('_')
}

function Get-GroupName {
    param([string]$FolderPath, [string]$PermissionType)
    
    $FolderName = Get-CleanName -Path $FolderPath
    $ParentFolder = Get-CleanName -Path (Split-Path $FolderPath -Parent)
    $CleanPath = $FolderPath -replace '^\\\\[^\\]+\\', '' -replace '[\\/:*?"<>|]', '_'
    
    # Kuerzel aus Config holen, Fallback auf vollständigen Namen
    $PermLabel = $Script:Config.PermissionLabels[$PermissionType]
    if (-not $PermLabel) { $PermLabel = $PermissionType }
    
    $GroupName = $Script:Config.GroupNameSchema
    $GroupName = $GroupName -replace '\{FolderName\}', $FolderName
    $GroupName = $GroupName -replace '\{ParentFolder\}', $ParentFolder
    $GroupName = $GroupName -replace '\{FullPath\}', $CleanPath
    $GroupName = $GroupName -replace '\{Permission\}', $PermLabel
    $GroupName = $GroupName -replace '\{PermissionFull\}', $PermissionType
    
    if ($GroupName.Length -gt 64) { $GroupName = $GroupName.Substring(0, 64) }
    return $GroupName
}

function Get-GroupDescription {
    param([string]$FolderPath, [string]$PermissionType)
    
    # Kuerzel aus Config holen
    $PermLabel = $Script:Config.PermissionLabels[$PermissionType]
    if (-not $PermLabel) { $PermLabel = $PermissionType }
    
    $Desc = $Script:Config.GroupDescriptionSchema
    $Desc = $Desc -replace '\{FolderName\}', (Split-Path $FolderPath -Leaf)
    $Desc = $Desc -replace '\{FullPath\}', $FolderPath
    $Desc = $Desc -replace '\{Permission\}', $PermLabel
    $Desc = $Desc -replace '\{PermissionFull\}', $PermissionType
    return $Desc
}

function New-PermissionGroup {
    param([string]$GroupName, [string]$Description, [string]$OU)
    
    $Existing = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
    if ($Existing) {
        Write-Log "Gruppe '$GroupName' existiert bereits"
        return $Existing
    }
    
    $Group = New-ADGroup -Name $GroupName -GroupScope DomainLocal -GroupCategory Security -Description $Description -Path $OU -PassThru
    Write-Log "Gruppe '$GroupName' erstellt"
    return $Group
}

function Set-FolderPermission {
    param([string]$FolderPath, [string]$GroupName, [string]$PermissionType, [bool]$DisableInheritance = $false)
    
    $PermConfig = $Script:Config.PermissionTypes[$PermissionType]
    $ACL = Get-Acl -Path $FolderPath
    
    if ($DisableInheritance) {
        $ACL.SetAccessRuleProtection($true, $true)
    }
    
    $Identity = "$env:USERDOMAIN\$GroupName"
    $Rights = [System.Security.AccessControl.FileSystemRights]$PermConfig.NTFSRights
    $InheritFlags = [System.Security.AccessControl.InheritanceFlags]$PermConfig.Inheritance
    $PropFlags = [System.Security.AccessControl.PropagationFlags]::None
    $AccessType = [System.Security.AccessControl.AccessControlType]::Allow
    
    $Rule = New-Object System.Security.AccessControl.FileSystemAccessRule($Identity, $Rights, $InheritFlags, $PropFlags, $AccessType)
    $ACL.AddAccessRule($Rule)
    Set-Acl -Path $FolderPath -AclObject $ACL
    
    Write-Log "Berechtigung '$PermissionType' für '$GroupName' auf '$FolderPath' gesetzt"
}

function Set-TraversePermission {
    param([string]$FolderPath, [string]$GroupName)
    
    $ACL = Get-Acl -Path $FolderPath
    $Identity = "$env:USERDOMAIN\$GroupName"
    $Rights = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute
    $InheritFlags = [System.Security.AccessControl.InheritanceFlags]::None
    $PropFlags = [System.Security.AccessControl.PropagationFlags]::None
    $AccessType = [System.Security.AccessControl.AccessControlType]::Allow
    
    $Rule = New-Object System.Security.AccessControl.FileSystemAccessRule($Identity, $Rights, $InheritFlags, $PropFlags, $AccessType)
    $ACL.AddAccessRule($Rule)
    Set-Acl -Path $FolderPath -AclObject $ACL
    
    Write-Log "Traverse für '$GroupName' auf '$FolderPath' gesetzt"
}

function Get-ParentFolders {
    param([string]$FolderPath, [string]$RootPath)
    
    $Parents = @()
    $Current = Split-Path $FolderPath -Parent
    while ($Current -and $Current.Length -gt $RootPath.Length) {
        $Parents += $Current
        $Current = Split-Path $Current -Parent
    }
    return $Parents
}

function Search-ADUsers {
    param([string]$SearchTerm)
    
    $Filter = "(&(objectClass=user)(|(cn=*$SearchTerm*)(sAMAccountName=*$SearchTerm*)(givenName=*$SearchTerm*)(sn=*$SearchTerm*)))"
    $Users = Get-ADUser -LDAPFilter $Filter -Properties DisplayName | 
             Select-Object @{N='IsSelected';E={$false}}, 
                           @{N='DisplayName';E={$_.DisplayName}},
                           @{N='SamAccountName';E={$_.SamAccountName}},
                           @{N='DN';E={$_.DistinguishedName}}
    return $Users
}

function Get-FolderPermissions {
    param([string]$FolderPath)
    
    $ACL = Get-Acl -Path $FolderPath
    $Perms = $ACL.Access | ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.IdentityReference.Value
            Rights = $_.FileSystemRights.ToString()
            Inheritance = if ($_.IsInherited) { "Ja" } else { "Nein" }
        }
    }
    return $Perms
}

# ============================================
# EVENT HANDLER
# ============================================

# Window Loaded
$Window.Add_Loaded({
    Test-ADModule | Out-Null
    
    # Standardwerte setzen
    $Script:UI.txtBasePath.Text = $Script:Config.DefaultFileServer + "\Daten"
    $Script:UI.txtDFSNamespace.Text = $Script:Config.DFSRoot
    $Script:UI.txtGroupOU.Text = $Script:Config.GroupOU
    $Script:UI.txtGroupNameSchema.Text = $Script:Config.GroupNameSchema
    $Script:UI.txtPermissionPath.Text = $Script:Config.DefaultFileServer + "\Daten"
    
    # Settings Tab
    $Script:UI.txtSettingsNameSchema.Text = $Script:Config.GroupNameSchema
    $Script:UI.txtSettingsDescSchema.Text = $Script:Config.GroupDescriptionSchema
    $Script:UI.txtSettingsOU.Text = $Script:Config.GroupOU
    $Script:UI.txtSettingsDFS.Text = $Script:Config.DFSRoot
    $Script:UI.txtSettingsServer.Text = $Script:Config.DefaultFileServer
    
    # Permission Labels DataGrid fuellen
    $Script:PermLabelsList = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
    foreach ($Key in $Script:Config.PermissionTypes.Keys) {
        $Label = $Script:Config.PermissionLabels[$Key]
        if (-not $Label) { $Label = $Key }
        $NTFSRights = $Script:Config.PermissionTypes[$Key].NTFSRights
        $Script:PermLabelsList.Add([PSCustomObject]@{
            Name = $Key
            Label = $Label
            NTFSRights = $NTFSRights
        })
    }
    $Script:UI.dgPermissionLabels.ItemsSource = $Script:PermLabelsList
    
    Write-Log "Bereit"
})

# Ordner durchsuchen
$Script:UI.btnBrowseBase.Add_Click({
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:UI.txtBasePath.Text = $FolderBrowser.SelectedPath
    }
})

$Script:UI.btnBrowsePermPath.Add_Click({
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:UI.txtPermissionPath.Text = $FolderBrowser.SelectedPath
    }
})

$Script:UI.btnQuickBrowse.Add_Click({
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:UI.txtQuickFolder.Text = $FolderBrowser.SelectedPath
    }
})

# Vorschau
$Script:UI.btnPreview.Add_Click({
    $BasePath = $Script:UI.txtBasePath.Text
    $FolderName = $Script:UI.txtNewFolderName.Text
    $Subfolders = $Script:UI.txtSubfolders.Text -split "`r`n" | Where-Object { $_ }
    
    $Preview = "=== VORSCHAU ===`r`n`r`n"
    $MainFolder = Join-Path $BasePath $FolderName
    $Preview += "ORDNER:`r`n  $MainFolder`r`n"
    
    foreach ($Sub in $Subfolders) {
        $Preview += "  +-- $Sub`r`n"
    }
    
    if ($Script:UI.chkCreateGroups.IsChecked) {
        $Preview += "`r`nGRUPPEN:`r`n"
        $PermTypes = @()
        if ($Script:UI.chkGroupRead.IsChecked) { $PermTypes += "Lesen" }
        if ($Script:UI.chkGroupModify.IsChecked) { $PermTypes += "Ändern" }
        if ($Script:UI.chkGroupFull.IsChecked) { $PermTypes += "Vollzugriff" }
        if ($Script:UI.chkGroupList.IsChecked) { $PermTypes += "Auflisten" }
        
        foreach ($Perm in $PermTypes) {
            $GN = Get-GroupName -FolderPath $MainFolder -PermissionType $Perm
            $Preview += "  $GN`r`n"
        }
    }
    
    $Script:UI.txtPreview.Text = $Preview
})

# Alles erstellen
$Script:UI.btnCreateAll.Add_Click({
    try {
        $BasePath = $Script:UI.txtBasePath.Text
        $FolderName = $Script:UI.txtNewFolderName.Text
        $Subfolders = $Script:UI.txtSubfolders.Text -split "`r`n" | Where-Object { $_ }
        $GroupOU = $Script:UI.txtGroupOU.Text
        
        if ([string]::IsNullOrWhiteSpace($FolderName)) {
            Show-Message "Bitte Ordnernamen eingeben" "Fehler" "Warning"
            return
        }
        
        $MainFolder = Join-Path $BasePath $FolderName
        
        $Result = [System.Windows.MessageBox]::Show("Ordner '$MainFolder' erstellen?", "Bestätigung", "YesNo", "Question")
        if ($Result -ne "Yes") { return }
        
        Write-Log "Erstelle Ordner..."
        
        # Ordner erstellen
        if (-not (Test-Path $MainFolder)) {
            New-Item -Path $MainFolder -ItemType Directory -Force | Out-Null
        }
        
        foreach ($Sub in $Subfolders) {
            $SubPath = Join-Path $MainFolder $Sub
            if (-not (Test-Path $SubPath)) {
                New-Item -Path $SubPath -ItemType Directory -Force | Out-Null
            }
        }
        
        # Gruppen erstellen
        if ($Script:UI.chkCreateGroups.IsChecked) {
            Write-Log "Erstelle Gruppen..."
            
            $PermTypes = @()
            if ($Script:UI.chkGroupRead.IsChecked) { $PermTypes += "Lesen" }
            if ($Script:UI.chkGroupModify.IsChecked) { $PermTypes += "Ändern" }
            if ($Script:UI.chkGroupFull.IsChecked) { $PermTypes += "Vollzugriff" }
            if ($Script:UI.chkGroupList.IsChecked) { $PermTypes += "Auflisten" }
            
            $DisableInh = $Script:UI.chkInheritanceOff.IsChecked
            
            foreach ($Perm in $PermTypes) {
                $GN = Get-GroupName -FolderPath $MainFolder -PermissionType $Perm
                $GD = Get-GroupDescription -FolderPath $MainFolder -PermissionType $Perm
                New-PermissionGroup -GroupName $GN -Description $GD -OU $GroupOU
                Set-FolderPermission -FolderPath $MainFolder -GroupName $GN -PermissionType $Perm -DisableInheritance $DisableInh
                $DisableInh = $false
            }
            
            foreach ($Sub in $Subfolders) {
                $SubPath = Join-Path $MainFolder $Sub
                foreach ($Perm in $PermTypes) {
                    $GN = Get-GroupName -FolderPath $SubPath -PermissionType $Perm
                    $GD = Get-GroupDescription -FolderPath $SubPath -PermissionType $Perm
                    New-PermissionGroup -GroupName $GN -Description $GD -OU $GroupOU
                    Set-FolderPermission -FolderPath $SubPath -GroupName $GN -PermissionType $Perm -DisableInheritance $true
                }
            }
        }
        
        Write-Log "Fertig!"
        Show-Message "Erfolgreich erstellt!" "Erfolg" "Information"
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# Nur Ordner erstellen
$Script:UI.btnCreateFolderOnly.Add_Click({
    try {
        $BasePath = $Script:UI.txtBasePath.Text
        $FolderName = $Script:UI.txtNewFolderName.Text
        $Subfolders = $Script:UI.txtSubfolders.Text -split "`r`n" | Where-Object { $_ }
        
        if ([string]::IsNullOrWhiteSpace($FolderName)) {
            Show-Message "Bitte Ordnernamen eingeben" "Fehler" "Warning"
            return
        }
        
        $MainFolder = Join-Path $BasePath $FolderName
        
        if (-not (Test-Path $MainFolder)) {
            New-Item -Path $MainFolder -ItemType Directory -Force | Out-Null
        }
        
        foreach ($Sub in $Subfolders) {
            $SubPath = Join-Path $MainFolder $Sub
            if (-not (Test-Path $SubPath)) {
                New-Item -Path $SubPath -ItemType Directory -Force | Out-Null
            }
        }
        
        Show-Message "Ordner erstellt!" "Erfolg" "Information"
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# Ordnerstruktur laden
$Script:UI.btnLoadTree.Add_Click({
    try {
        $Path = $Script:UI.txtPermissionPath.Text
        
        if (-not (Test-Path $Path)) {
            Show-Message "Pfad nicht gefunden" "Fehler" "Error"
            return
        }
        
        $Script:UI.tvFolders.Items.Clear()
        
        $RootItem = New-Object System.Windows.Controls.TreeViewItem
        $RootItem.Header = Split-Path $Path -Leaf
        $RootItem.Tag = $Path
        $RootItem.IsExpanded = $true
        
        function Add-SubFolders {
            param($ParentItem, $FolderPath, $Depth)
            if ($Depth -gt 3) { return }
            
            try {
                $Folders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
                foreach ($Folder in $Folders) {
                    $Item = New-Object System.Windows.Controls.TreeViewItem
                    $Item.Header = $Folder.Name
                    $Item.Tag = $Folder.FullName
                    Add-SubFolders -ParentItem $Item -FolderPath $Folder.FullName -Depth ($Depth + 1)
                    $ParentItem.Items.Add($Item) | Out-Null
                }
            }
            catch {}
        }
        
        Add-SubFolders -ParentItem $RootItem -FolderPath $Path -Depth 1
        $Script:UI.tvFolders.Items.Add($RootItem) | Out-Null
        
        Write-Log "Ordnerstruktur geladen"
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# TreeView Selection
$Script:UI.tvFolders.Add_SelectedItemChanged({
    $Selected = $Script:UI.tvFolders.SelectedItem
    if ($Selected -and $Selected.Tag) {
        $Script:UI.txtSelectedFolder.Text = $Selected.Tag
        $Perms = Get-FolderPermissions -FolderPath $Selected.Tag
        $Script:UI.dgPermissions.ItemsSource = $Perms
    }
})

# Berechtigung hinzufügen
$Script:UI.btnAddPermission.Add_Click({
    try {
        $FolderPath = $Script:UI.txtSelectedFolder.Text
        $PermType = $Script:UI.cmbPermissionType.SelectedItem.Content
        $GroupOU = $Script:UI.txtGroupOU.Text
        
        if ([string]::IsNullOrWhiteSpace($FolderPath)) {
            Show-Message "Bitte Ordner auswählen" "Fehler" "Warning"
            return
        }
        
        $GN = Get-GroupName -FolderPath $FolderPath -PermissionType $PermType
        $GD = Get-GroupDescription -FolderPath $FolderPath -PermissionType $PermType
        
        New-PermissionGroup -GroupName $GN -Description $GD -OU $GroupOU
        Set-FolderPermission -FolderPath $FolderPath -GroupName $GN -PermissionType $PermType -DisableInheritance $true
        
        if ($Script:UI.chkSetTraverse.IsChecked) {
            $RootPath = $Script:UI.txtPermissionPath.Text
            $Parents = Get-ParentFolders -FolderPath $FolderPath -RootPath $RootPath
            foreach ($Parent in $Parents) {
                Set-TraversePermission -FolderPath $Parent -GroupName $GN
            }
        }
        
        $Perms = Get-FolderPermissions -FolderPath $FolderPath
        $Script:UI.dgPermissions.ItemsSource = $Perms
        
        Show-Message "Gruppe '$GN' erstellt und berechtigt!" "Erfolg" "Information"
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# Benutzersuche
$Script:UI.btnSearchUser.Add_Click({
    try {
        $SearchTerm = $Script:UI.txtUserSearch.Text
        if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
            Show-Message "Bitte Suchbegriff eingeben" "Fehler" "Warning"
            return
        }
        
        $Users = Search-ADUsers -SearchTerm $SearchTerm
        $Script:UI.dgUsers.ItemsSource = $Users
        Write-Log "$($Users.Count) Benutzer gefunden"
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# Gruppensuche
$Script:UI.btnSearchGroup.Add_Click({
    try {
        $SearchTerm = $Script:UI.txtGroupSearch.Text
        if ([string]::IsNullOrWhiteSpace($SearchTerm)) { $SearchTerm = "FS_" }
        
        $Groups = Get-ADGroup -Filter "Name -like '*$SearchTerm*'" | Select-Object Name
        $Script:UI.lbGroups.ItemsSource = $Groups.Name
        Write-Log "$($Groups.Count) Gruppen gefunden"
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# Gruppenmitglieder
$Script:UI.lbGroups.Add_SelectionChanged({
    try {
        $GroupName = $Script:UI.lbGroups.SelectedItem
        if ($GroupName) {
            $Members = Get-ADGroupMember -Identity $GroupName -ErrorAction SilentlyContinue | Select-Object Name, ObjectClass
            $Script:UI.dgGroupMembers.ItemsSource = $Members
        }
    }
    catch {}
})

# Benutzer zu Gruppe hinzufügen
$Script:UI.btnAddToGroup.Add_Click({
    try {
        $SelectedUsers = @($Script:UI.dgUsers.ItemsSource | Where-Object { $_.IsSelected })
        $GroupName = $Script:UI.lbGroups.SelectedItem
        
        if (-not $GroupName) {
            Show-Message "Bitte Gruppe auswählen" "Fehler" "Warning"
            return
        }
        
        if ($SelectedUsers.Count -eq 0) {
            Show-Message "Bitte Benutzer auswählen (Checkbox)" "Fehler" "Warning"
            return
        }
        
        foreach ($User in $SelectedUsers) {
            Add-ADGroupMember -Identity $GroupName -Members $User.SamAccountName
        }
        
        $Members = Get-ADGroupMember -Identity $GroupName | Select-Object Name, ObjectClass
        $Script:UI.dgGroupMembers.ItemsSource = $Members
        
        Show-Message "$($SelectedUsers.Count) Benutzer hinzugefügt!" "Erfolg" "Information"
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# Benutzer aus Gruppe entfernen
$Script:UI.btnRemoveFromGroup.Add_Click({
    try {
        $SelectedMember = $Script:UI.dgGroupMembers.SelectedItem
        $GroupName = $Script:UI.lbGroups.SelectedItem
        
        if (-not $GroupName -or -not $SelectedMember) {
            Show-Message "Bitte Gruppe und Mitglied auswählen" "Fehler" "Warning"
            return
        }
        
        $Result = [System.Windows.MessageBox]::Show("'$($SelectedMember.Name)' entfernen?", "Bestätigung", "YesNo", "Question")
        if ($Result -eq "Yes") {
            Remove-ADGroupMember -Identity $GroupName -Members $SelectedMember.Name -Confirm:$false
            $Members = Get-ADGroupMember -Identity $GroupName | Select-Object Name, ObjectClass
            $Script:UI.dgGroupMembers.ItemsSource = $Members
            Show-Message "Entfernt!" "Erfolg" "Information"
        }
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# Schnellzuweisung - Benutzersuche
$Script:UI.btnQuickSearchUser.Add_Click({
    try {
        $SearchTerm = $Script:UI.txtQuickUser.Text
        if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
            Show-Message "Bitte Suchbegriff eingeben" "Fehler" "Warning"
            return
        }
        
        $Users = Search-ADUsers -SearchTerm $SearchTerm
        $Script:UI.lbQuickUsers.ItemsSource = $Users | ForEach-Object { "$($_.DisplayName) ($($_.SamAccountName))" }
        $Script:UI.lbQuickUsers.Tag = $Users
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# Schnellzuweisung - Vorschau
$Script:UI.btnQuickPreview.Add_Click({
    $SelectedIndex = $Script:UI.lbQuickUsers.SelectedIndex
    $FolderPath = $Script:UI.txtQuickFolder.Text
    $PermType = $Script:UI.cmbQuickPermission.SelectedItem.Content
    
    if ($SelectedIndex -lt 0 -or [string]::IsNullOrWhiteSpace($FolderPath)) {
        Show-Message "Bitte Benutzer und Ordner auswählen" "Fehler" "Warning"
        return
    }
    
    $Users = $Script:UI.lbQuickUsers.Tag
    $User = $Users[$SelectedIndex]
    
    $GN = Get-GroupName -FolderPath $FolderPath -PermissionType $PermType
    
    $Preview = "BENUTZER: $($User.DisplayName)`r`n"
    $Preview += "ORDNER: $FolderPath`r`n"
    $Preview += "BERECHTIGUNG: $PermType`r`n"
    $Preview += "GRUPPE: $GN`r`n`r`n"
    $Preview += "AKTIONEN:`r`n"
    $Preview += "  - Gruppe erstellen/prüfen`r`n"
    $Preview += "  - Benutzer hinzufügen`r`n"
    
    if ($Script:UI.chkQuickTraverse.IsChecked) {
        $Preview += "  - Traverse-Gruppen setzen`r`n"
    }
    
    $Script:UI.txtQuickPreview.Text = $Preview
})

# Schnellzuweisung - Ausfuehren
$Script:UI.btnQuickApply.Add_Click({
    try {
        $SelectedIndex = $Script:UI.lbQuickUsers.SelectedIndex
        $FolderPath = $Script:UI.txtQuickFolder.Text
        $PermType = $Script:UI.cmbQuickPermission.SelectedItem.Content
        $GroupOU = $Script:UI.txtGroupOU.Text
        
        if ($SelectedIndex -lt 0 -or [string]::IsNullOrWhiteSpace($FolderPath)) {
            Show-Message "Bitte Benutzer und Ordner auswählen" "Fehler" "Warning"
            return
        }
        
        $Users = $Script:UI.lbQuickUsers.Tag
        $User = $Users[$SelectedIndex]
        
        $GN = Get-GroupName -FolderPath $FolderPath -PermissionType $PermType
        
        if ($Script:UI.chkQuickCreateGroups.IsChecked) {
            $GD = Get-GroupDescription -FolderPath $FolderPath -PermissionType $PermType
            New-PermissionGroup -GroupName $GN -Description $GD -OU $GroupOU
            Set-FolderPermission -FolderPath $FolderPath -GroupName $GN -PermissionType $PermType -DisableInheritance $true
        }
        
        Add-ADGroupMember -Identity $GN -Members $User.SamAccountName -ErrorAction SilentlyContinue
        Write-Log "Benutzer '$($User.SamAccountName)' zu '$GN' hinzugefügt"
        
        if ($Script:UI.chkQuickTraverse.IsChecked) {
            $RootPath = Split-Path (Split-Path $FolderPath -Parent) -Parent
            $Parents = Get-ParentFolders -FolderPath $FolderPath -RootPath $RootPath
            
            foreach ($Parent in $Parents) {
                $TGN = Get-GroupName -FolderPath $Parent -PermissionType "Auflisten"
                $TG = Get-ADGroup -Filter "Name -eq '$TGN'" -ErrorAction SilentlyContinue
                
                if (-not $TG) {
                    $TGD = Get-GroupDescription -FolderPath $Parent -PermissionType "Auflisten"
                    New-PermissionGroup -GroupName $TGN -Description $TGD -OU $GroupOU
                    Set-TraversePermission -FolderPath $Parent -GroupName $TGN
                }
                
                Add-ADGroupMember -Identity $TGN -Members $User.SamAccountName -ErrorAction SilentlyContinue
            }
        }
        
        Show-Message "Benutzer '$($User.DisplayName)' wurde berechtigt!" "Erfolg" "Information"
    }
    catch {
        Show-Message "Fehler: $($_.Exception.Message)" "Fehler" "Error"
    }
})

# Einstellungen speichern
$Script:UI.btnSaveSettings.Add_Click({
    $Script:Config.GroupNameSchema = $Script:UI.txtSettingsNameSchema.Text
    $Script:Config.GroupDescriptionSchema = $Script:UI.txtSettingsDescSchema.Text
    $Script:Config.GroupOU = $Script:UI.txtSettingsOU.Text
    $Script:Config.DFSRoot = $Script:UI.txtSettingsDFS.Text
    $Script:Config.DefaultFileServer = $Script:UI.txtSettingsServer.Text
    
    # Permission Labels aus DataGrid speichern
    foreach ($Item in $Script:PermLabelsList) {
        $Script:Config.PermissionLabels[$Item.Name] = $Item.Label
    }
    
    $Script:UI.txtGroupOU.Text = $Script:Config.GroupOU
    $Script:UI.txtGroupNameSchema.Text = $Script:Config.GroupNameSchema
    $Script:UI.txtDFSNamespace.Text = $Script:Config.DFSRoot
    $Script:UI.txtBasePath.Text = $Script:Config.DefaultFileServer + "\Daten"
    
    Show-Message "Einstellungen gespeichert!" "Erfolg" "Information"
})

# Einstellungen zurücksetzen
$Script:UI.btnResetSettings.Add_Click({
    $Script:UI.txtSettingsNameSchema.Text = "FS_{FolderName}_{Permission}"
    $Script:UI.txtSettingsDescSchema.Text = "Dateisystem-Berechtigung: {PermissionFull} auf {FullPath}"
    $Script:UI.txtSettingsOU.Text = "OU=Dateisystem-Gruppen,OU=Gruppen,DC=domain,DC=local"
    $Script:UI.txtSettingsDFS.Text = "\\domain.local\dfs"
    $Script:UI.txtSettingsServer.Text = "\\fileserver01"
    
    # Permission Labels zurücksetzen
    $DefaultLabels = @{
        "Lesen" = "RO"
        "Ändern" = "RW"
        "Vollzugriff" = "FC"
        "Auflisten" = "LS"
        "NurDieserOrdner" = "MO"
    }
    foreach ($Item in $Script:PermLabelsList) {
        if ($DefaultLabels.ContainsKey($Item.Name)) {
            $Item.Label = $DefaultLabels[$Item.Name]
        }
    }
    $Script:UI.dgPermissionLabels.Items.Refresh()
})

# ============================================
# FENSTER ANZEIGEN
# ============================================
$Window.ShowDialog() | Out-Null
