<#
    @author Camron Christ
            Camron.Christ@hcahealthcare.com

    This program allows a user to install any number of printers to any number
    of computers using the drivers installed on their local machine.
#>

# XAML outlining the GUI makeup
$xaml = @' 
<Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  xmlns:Themes="clr-namespace:Microsoft.Windows.Themes;assembly=PresentationFramework.Aero2"

  Title="WPFBatchInstallPrinter" Height="730" Width="540">
    
    <Grid Name="MainGrid" Margin="8,8">

        <Grid.RowDefinitions>
            <RowDefinition Height="24"/>
            <RowDefinition Height="24"/>
            <RowDefinition Height="24"/>
            <RowDefinition Height="14*"/>
            <RowDefinition Height="28"/>
            <RowDefinition Height="10*"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="1*"/>
            </Grid.ColumnDefinitions>
            <Label Content="Printer Name" Margin="0,0,6,0" Grid.Column="0"/>
            <Label Content="Driver" Margin="0,0,6,0" Grid.Column="1"/>
            <ComboBox Name="comboBoxPrinterConnection" Grid.Column="2" Margin="0,0,6,0" SelectedIndex="0" />
        </Grid>

        <Grid Margin="0,0,0,0" Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="1*"/>
            </Grid.ColumnDefinitions>
            <TextBox Grid.Column="0" x:Name="textBoxPrinterName" ToolTip="Printer Name" Margin="0,0,6,0"/>
            <ComboBox Grid.Column="1" x:Name="comboBoxDrivers" ToolTip="Driver" Margin="0,0,6,0"/>
            <TextBox Grid.Column="2" x:Name="textBoxPrinterConnection" ToolTip="IP Address / Server Connection" Margin="0,0,6,0"/>
            <Button Grid.Column="3" x:Name="buttonAddPrinter" Content="Add" ToolTip="Add Printer to Install List"/>
        </Grid>

        <Grid Margin="0,0,0,0" Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1*"/>
                <ColumnDefinition Width="1*"/>
            </Grid.ColumnDefinitions>
            <Label Content="Computers" Grid.Column="0" HorizontalAlignment="Center"/>
            <Label Content="Printers" Grid.Column="1" HorizontalAlignment="Center"/>

        </Grid>

        <Grid Margin="0,0,0,0" Grid.Row="3">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1*"/>
                <ColumnDefinition Width="1*"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="textBoxComputers" CharacterCasing="Upper" ToolTip="List Computers one per line:&#x0d;&#x0a;COMP001&#x0d;&#x0a;COMP002" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
            <ListBox x:Name="listBoxPrinters" Grid.Column="1" Margin="3,0,0,0"  ToolTip="Printer is Added with Add Button" Focusable="False" SelectionMode="Extended">
                <ListBox.ItemTemplate>
                    <DataTemplate DataType="tfs:WorkItem">
                        <StackPanel>
                            <TextBlock Text="{Binding Name}" />
                            <TextBlock Text="{Binding Driver}" />
                            <TextBlock Text="{Binding Port}" />
                        </StackPanel>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>
        </Grid>

        <Grid Margin="0,0,0,0" Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="6*"/>
                <ColumnDefinition Width="1*"/>
                <ColumnDefinition Width="1.5*"/>
            </Grid.ColumnDefinitions>
            <Label Name="labelConsole" Content="Console:" Grid.Column="0" Margin="0,0,6,0"/>
            <Button x:Name="buttonInstall" Content="Install" Grid.Column="1" Margin="0,3,6,3" ToolTip="Install printers to computers"/>
            <Button x:Name="buttonClearConsole" Content="Clear Console" Grid.Column="2" Margin="0,3,0,3" ToolTip="Install printers to computers"/>
        </Grid>

        <Grid Margin="0,0,0,0" Grid.Row="5">
            <TextBox x:Name="textBoxConsole" VerticalScrollBarVisibility="Auto" IsReadOnly="True" FontFamily="Lucida Console"/>
        </Grid>
    </Grid>

</Window>
'@
$Global:PrinterConnections = @(
"TCP/IP",
"Local"
)

# Converts XAML to an object with each tag becoming a field within the object
# for easy manipulation
function Convert-XAMLtoWindow{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $XAML
    )
    
    Add-Type -AssemblyName PresentationFramework
    
    $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
    $result = [Windows.Markup.XAMLReader]::Load($reader)
    $reader.Close()
    $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
    while ($reader.Read())
    {
        $name=$reader.GetAttribute('Name')
        if (!$name) { $name=$reader.GetAttribute('x:Name') }
        if($name)
        {$result | Add-Member NoteProperty -Name $name -Value $result.FindName($name) -Force}
    }
    $reader.Close()
    $result
}

function Show-WPFWindow{
    param
    (
        [Parameter(Mandatory)]
        [Windows.Window]
        $window
    )
    
    $result = $null
    $null = $syncHash.Dispatcher.InvokeAsync{
        $result = $syncHash.ShowDialog()
        Set-Variable -Name result -Value $result -Scope 1
    }.Wait()
    $result
}

# Ensures the printer being added has a valid TCP/IP address and name
# Returns Printer object
function Validate-Printer{
    
    $isValidIPAddress = ($syncHash.comboBoxPrinterConnection.Text -like $Global:PrinterConnections[0] -and $syncHash.textBoxPrinterConnection.Text.Trim() -match
        '^\b(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\b$'
    )
    $isValidPrinterName = $syncHash.textBoxPrinterName.Text -ne ""
    $isValidPrintServerConnection = ($syncHash.comboBoxPrinterConnection.Text -like $Global:PrinterConnections[1] -and
        $syncHash.textBoxPrinterConnection.Text.Trim() -match '^\\\\\w+\\\w+$'
    )
    
    if($isValidPrinterName -and ($isValidIPAddress -or $isValidPrintServerConnection))
    {
        if($syncHash.textBoxPrinterConnection.BorderBrush -like "#FFFF0000") {
            $syncHash.textBoxPrinterConnection.BorderBrush = "#FFABADB3"
        }
        # Creates Printer object
        $printer = New-Object -TypeName PSObject
        $printer | Add-Member -MemberType NoteProperty -Name Name -Value $syncHash.textBoxPrinterName.Text.Trim()
        $printer | Add-Member -MemberType NoteProperty -Name Driver -Value $syncHash.comboBoxDrivers.Text.Trim()
        $printer | Add-Member -MemberType NoteProperty -Name Port -Value $syncHash.textBoxPrinterConnection.Text.Trim()
        $printer | Add-Member -MemberType NoteProperty -Name PortType -Value $(if($isValidIPAddress){$Global:PrinterConnections[0]}else{$Global:PrinterConnections[1]})
        return $printer
    }

    if(-not $isValidPrinterName) {
        $syncHash.textBoxPrinterName.BorderBrush = "#FFFF0000"
    }

    if(-not ($isValidIPAddress -or $isValidPrintServerConnection))
    {
        $syncHash.textBoxPrinterConnection.BorderBrush = "red"
    }

    return $null
}

# Validates and adds printer to install list
function Add-PrinterToList{
    $printer = Validate-Printer
    if($null -ne $printer){
        $syncHash.listBoxPrinters.Items.Add($printer)
    }
}

# Installs printers to all computers
function Install-Printers{
    $computers = Parse-Computers
    $printers = $syncHash.listBoxPrinters.Items
    Write-Console("Initiating install...")

    #$syncHash.Host = $host
    $InstallRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash) 
    $InstallRunspace.SessionStateProxy.SetVariable("printers",$printers)
    $InstallRunspace.SessionStateProxy.SetVariable("computers",$computers)
    $InstallRunspace.SessionStateProxy.SetVariable("PrinterConnections",$global:PrinterConnections)
    $global:aSync = "" | Select-Object PowerShell,Runspace,Job
    $global:aSync.Job = $InstallRunspace.Name

    $code = {
    function Write-Console($str){
        $syncHash.Dispatcher.invoke(
        [action]{
                    $syncHash.textBoxConsole.Text += "$str`n"
                    $syncHash.textboxConsole.ScrollToEnd()
        })
    }
    [string[]]$failed = @()
    foreach($c in $computers){
        if($null -ne $c -and $c.Trim() -ne ""){
            Write-Console "`n$c`: Testing Connection..."
            if(Test-Connection -ComputerName $c -Count 2 -Quiet -ErrorAction stop){
                foreach($p in $printers){
                    
                    Write-Console "`n  $($p.Name)`: Beginning installation..."

                    # Test whether the print spooler is reachable
                    Write-Console "`tChecking Print Spooler..."
                    try
                    {
                        Get-Printer -ComputerName $c -ErrorAction Stop > $null 2>&1
                    }
                    catch
                    {
                        reg add "\\$c\HKLM\Software\Policies\Microsoft\Windows NT\Printers" /v RegisterSpoolerRemoteRpcEndPoint /t REG_DWORD /d 1 /f
                        Get-Service -ComputerName $c -Name Spooler | Restart-Service
                        Write-Console "$($c): Waiting for Spooler to restart..."
                        (Get-Service -ComputerName $c -Name Spooler).('Running', '00:00:05')
                    }

                    # If port does not exist, add port
                    Write-Console "`tChecking Port..."
                    try
                    {
                        Get-PrinterPort -ComputerName $c -Name $p.Port -ErrorAction Stop > $null 2>&1
                    }
                    catch
                    {
                        Write-Console "`tAdding Port: `'$($p.PortType) $($p.Port)`'"
                        if($p.portType -like $printerConnections[0]) {
                            Add-PrinterPort -ComputerName $c -Name $p.Port -PrinterHostAddress $p.Port
                        }
                        elseif($p.portType -like $printerConnections[1]) {
                            Add-PrinterPort -ComputerName $c -Name $p.Port
                        }
                        else {
                            Write-Console "`tPort: `'$($p.PortType) $($p.Port)`' installation failed..."
                        }
                    }

                    <# If driver does not exist, add driver
                    Write-Console "`tChecking Driver..."
                    try
                    {
                        Get-PrinterDriver -ComputerName $c -Name $p.Driver -ErrorAction Stop > $null 2>&1
                    }
                    catch
                    {
                        Write-Console "`tInstalling Driver: `'$($p.Driver)`'"
                        $sourceComputerName = hostname
                        $printerDriverINF = (Get-PrinterDriver -Name "Xerox Global Print Driver PCL6").InfPath -replace "C:\\","\\$sourceComputerName\C$"
                        Add-PrinterDriver -ComputerName $c -Name $p.Driver -InfPath $printerDriverINF
                    }
                    #>

                    # install printer
                    Write-Console "`tAdding Printer..."
                    #Add-Printer -ComputerName $c -Name $p.Name -PortName $p.Port -DriverName $p.Driver
                    rundll32 printui.dll,PrintUIEntry /if /c\\$c /b $($p.Name) /r $($p.Port) /m $($p.Driver) /q | Wait-Process
                    
                    # Verify successful installation
                    Write-Console "`tVerifying successful install..."
                    try
                    {
                        Get-Printer -ComputerName $c -Name $p.Name -ErrorAction Stop > $null 2>&1
                        Write-Console ("`t`'$($p.Name)`' Installed")
                    }
                    catch
                    {
                        Write-Console("`t`'$($p.Name)`' installation not successful")
                        $failed += "$($c): $($p.Name)"
                    }
                }
            
            }
            # PC not on network
            else
            {
                Write-Console( "`t$($c): Not found..." )
                $failed += "$($c): Not found..."
            }
        }
    }
    Write-Console("`n----------Finished------------")
    if($failed.Count -gt 0){
        Write-Console("")
        Write-Console("------Failed Device List------")
        Write-Console($failed -join "`n")
    }
    $syncHash.Dispatcher.invoke(
        [action]{
            $syncHash.buttonInstall.Content = "Install"
            $syncHash.buttonInstall.ToolTip = "Install printers to computers"
    })
    }

    $PSinstance = [powershell]::Create().AddScript($code)
    $PSinstance.Runspace = $InstallRunspace
    $global:aSync.PowerShell = $PSinstance
    $global:aSync.Runspace = $PSinstance.BeginInvoke()
    #$jobs.Add($global:aSync)
}

function Parse-Computers{
    return ($syncHash.textBoxComputers.Text -split "`n") | % {$_.ToString().Trim()}
}
function Write-Console($str){
    $syncHash.textBoxConsole.Text += "$str`n"
    $syncHash.textboxConsole.ScrollToEnd()
}

$syncHash = [hashtable]::Synchronized(@{})
$syncHash = Convert-XAMLtoWindow -XAML $xaml

[bool]$TEST = 0

$InstallRunspace = [runspacefactory]::CreateRunspace()
$InstallRunspace.ApartmentState = "STA"
$InstallRunspace.ThreadOptions = "ReuseThread"
$InstallRunspace.Name = "InstallRunspace"
$InstallRunspace.Open()

#Events
#-------------------------------------------------------------
$syncHash.buttonAddPrinter.add_click({
    Add-PrinterToList
})
$syncHash.buttonInstall.add_click({
    if($syncHash.buttonInstall.Content -eq "Install"){
        $syncHash.buttonInstall.Content = "Stop"
        $syncHash.buttonInstall.ToolTip = "Cancel install"
        Install-Printers
    } else {
        $syncHash.buttonInstall.Content = "Install"
        $syncHash.buttonInstall.ToolTip = "Install printers to computers"
        $global:aSync.PowerShell.Stop()
        Write-Console("Cancelled install")
    }
})
#====Function Contains Test Material ===================
$syncHash.buttonClearConsole.add_click({
    $syncHash.textBoxConsole.Text = ""
    
    if($TEST) {
        Write-Console ("=======TEST=======")
    }
})
$syncHash.listBoxPrinters.add_keyDown({
    if ($args[1].key -eq 'Delete'){        
        While($syncHash.listBoxPrinters.SelectedItems.Count -gt 0){
            $syncHash.listBoxPrinters.Items.RemoveAt($syncHash.listBoxPrinters.SelectedIndex)
        }
    }
})
$syncHash.listBoxPrinters.Add_PreviewDragOver({
    [System.Object]$script:sender = $args[0]
    [System.Windows.DragEventArgs]$e = $args[1]

    $e.Effects = [System.Windows.DragDropEffects]::Copy
    $e.Handled = $true
})
$syncHash.listBoxPrinters.Add_Drop({

    [System.Object]$script:sender = $args[0]
    [System.Windows.DragEventArgs]$e = $args[1]

    Write-Host $e.Data.GetData([System.Windows.DataFormats]::FileDrop)

    
})
$syncHash.textBoxPrinterName.add_gotKeyboardFocus({
    if ($syncHash.textBoxPrinterName.BorderBrush -like "#FFFF0000") {
        $syncHash.textBoxPrinterName.BorderBrush = "#FFABADB3"
    }
})
$syncHash.textBoxPrinterName.add_keyDown({
    if ($args[1].key -eq 'Enter'){
        Add-PrinterToList
    }
})
$syncHash.textBoxPrinterConnection.add_gotKeyboardFocus({
    if ($syncHash.textBoxPrinterConnection.BorderBrush -like "#FFFF0000") {
        $syncHash.textBoxPrinterConnection.BorderBrush = "#FFABADB3"
    }
})
$syncHash.textBoxPrinterConnection.add_keyDown({
    if ($args[1].key -eq 'Enter'){
        Add-PrinterToList
    }
})

$syncHash.comboBoxPrinterConnection.ItemsSource = $Global:PrinterConnections
$syncHash.comboBoxDrivers.ItemsSource = (Get-PrinterDriver).Name | Sort-Object
$syncHash.comboBoxDrivers.SelectedIndex = 0

#================= Test ===================

#=============== End Test ==================

$null = Show-WPFWindow -Window $syncHash

