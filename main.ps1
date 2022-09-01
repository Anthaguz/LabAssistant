Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

function Get-isAdmin{
    #Checking if currently run as administrator
    $elevated = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $elevated) {
        [System.Windows.MessageBox]::Show('You need to run the tool as administrator')
        Exit
    }
}
function Set-LabVMProperties {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [string]
        $PathLab,
        [Parameter(Position = 1)]  
        [string]
        $PathVM, 
        [Parameter(Position = 2)]
        [string]
        $PathVHD,
        [Parameter(Position = 3)] 
        [string]
        $VMName, 
        [Parameter(Position = 4)]
        $VMOS, 
        [Parameter(Position = 5)]
        [string]
        $VMVSwitch    
    )
    $os="asd"
    switch ($vmos) {
        "Windows Server 2008 R2" {$os="C:\Training\BaseVHDX\En_Win_Server_2008_R2.vhdx"}
        "Windows Server 2012 R2" {$os="C:\Training\BaseVHDX\En_Win_Server_2012_R2.vhdx"}
        "Windows 10" {$os="C:\Training\BaseVHDX\En_Win_Client_10.vhdx"}
        "Windows Server 2016" {$os="C:\Training\BaseVHDX\En_Win_Server_2016 .vhdx"}
        "Windows Server 2019" {$os="C:\Training\BaseVHDX\En_Win_Server_2019 .vhdx"}
        "Windows Server 2022" {$os="C:\Training\BaseVHDX\En_Win_Server_2022 .vhdx"}
        "Windows 11" {$os="C:\Training\BaseVHDX\En_Win_Client_11.vhdx"}
        Default {}
    }
    New-Item -Path $PathLab -Name $VMName -ItemType "directory"
    New-VHD -ParentPath $os -Path $PathVHD -Differencing
    if ($os -eq "C:\Training\BaseVHDX\En_Win_Server_2008_R2.vhdx")
    {New-VM -Name $VMName -MemoryStartupBytes 2GB -BootDevice VHD -VHDPath $PathVHD -Path $PathVM -Generation 1 -Switch $VMVSwitch}
    else{
    New-VM -Name $VMName -MemoryStartupBytes 2GB -BootDevice VHD -VHDPath $PathVHD -Path $PathVM -Generation 2 -Switch $VMVSwitch}
    Start-VM -name $VMName
}
Function Get-Folder($initialDirectory = "") {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory
    if ($foldername.ShowDialog() -eq "OK") { $folder += $foldername.SelectedPath }else{$folder="C:\Training\VM"}
    return $folder
}
Function Show-CVForm {
    Set-Location $global:exepath
    $xamlFile2 = "CVForm.xaml"
    $inputXML2 = Get-Content $xamlFile2 -Raw
    $inputXML2 = $inputXML2 -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [XML]$XAML2 = $inputXML2
    $reader2 = (New-Object System.Xml.XmlNodeReader $xaml2)
    #Create VM Form lifecycle
    try { $CVFormDisplay = [Windows.Markup.XamlReader]::Load($reader2) } 
    catch {
        Write-Warning $_.Exception
        throw
    }
    $xaml2.SelectNodes("//*[@Name]") | ForEach-Object {
        try { Set-Variable -Name "var_$($_.Name)" -Value $CVFormDisplay.FindName($_.Name) -ErrorAction Stop }
        catch { throw }
    }
    $Folder = "C:\training\BaseVHDX"
    if (-not (Test-Path -Path $Folder)) {
        New-Item $Folder -itemType Directory 
        [System.Windows.MessageBox]::Show(
        'The folder C:\training\BaseVHDX did not exist so it was created. Make sure you 
have the correct VHDX files in that location before running the tool. You can
use the download lab option to download the image files.')
    }else{
        Set-Location -path $Folder
        $missingvhdx=@{}
        if((Test-Path -Path ".\En_Win_Server_2008_R2.vhdx")){
            $var_CV_CBOS.items.Add("Windows Server 2008 R2")
            $missingvhdx.Add("1","vacio")
        }else{$missingvhdx.Add("1","Windows Server 2008 R2")}

        if((Test-Path -Path ".\En_Win_Server_2012_R2.vhdx")){
            $var_CV_CBOS.items.Add("Windows Server 2012 R2")
            $missingvhdx.Add("2","vacio")
        }else{$missingvhdx.Add("2","Windows Server 2012 R2")}

        if((Test-Path -Path ".\En_Win_Client_10.vhdx")){
            $var_CV_CBOS.items.Add("Windows 10")
            $missingvhdx.Add("3","vacio")
        }else{$missingvhdx.Add("3","Windows 10")}

        if((Test-Path -Path ".\En_Win_Server_2016 .vhdx")){
            $var_CV_CBOS.items.Add("Windows Server 2016")
            $missingvhdx.Add("4","vacio")
        }else{$missingvhdx.Add("4","Windows Server 2016")}

        if((Test-Path -Path ".\En_Win_Server_2019 .vhdx")){
            $var_CV_CBOS.items.Add("Windows Server 2019")
            $missingvhdx.Add("5","vacio")
        }else{$missingvhdx.Add("5","Windows Server 2019")}

        if((Test-Path -Path ".\En_Win_Server_2022 .vhdx")){
            $var_CV_CBOS.items.Add("Windows Server 2022")
            $missingvhdx.Add("6","vacio")
        }else{$missingvhdx.Add("6","Windows Server 2022")}

        if((Test-Path -Path ".\En_Win_Client_11.vhdx")){
            $var_CV_CBOS.items.Add("Windows 11")
            $missingvhdx.Add("7","vacio")
        }else{$missingvhdx.Add("7","Windows 11")}
        $message="
"
        $messagecomparison="
"
        if (-not ($missingvhdx['1'] -eq "vacio")){$message=$message+$missingvhdx['1']+"
"}
        if (-not ($missingvhdx['2'] -eq "vacio")){$message=$message+$missingvhdx['2']+"
"}
        if (-not ($missingvhdx['3'] -eq "vacio")){$message=$message+$missingvhdx['3']+"
"}
        if (-not ($missingvhdx['4'] -eq "vacio")){$message=$message+$missingvhdx['4']+"
"}
        if (-not ($missingvhdx['5'] -eq "vacio")){$message=$message+$missingvhdx['5']+"
"}
        if (-not ($missingvhdx['6'] -eq "vacio")){$message=$message+$missingvhdx['6']+"
"}
        if (-not ($missingvhdx['7'] -eq "vacio")){$message=$message+$missingvhdx['7']+"
"}

        if (-not($message -eq $messagecomparison)) {
            <# Action to perform if the condition is true #>
        [System.Windows.MessageBox]::Show("There is no valid VHDX for the following OS in the 
location C:\Training\BaseVHDX        
        "+
            $message
        )
    }
    }
    Set-Location $global:exepath
    Get-VMSwitch | ForEach-Object { $var_CV_CBVS.items.Add($_.Name) }
    $var_CV_CBOS.selectedindex = 0
    $var_CV_CBVS.selectedindex = 0
    $var_CV_tbxFolderPath.text=$global:LabLocation

    ############ VM Properties from combo box ################
    ##### To be set on an event listener #####################
    $eventVMOS = $var_CV_CBOS.selectedindex
    $eventVMVS = $var_CV_CBVS.selecteditem
    
    ############## Create VM Form ############################
    ###### Event Listeners ###################################
    $var_CV_tbxFolderPath.text="C:\Training\VM"
    $var_CV_tbxFolderPath.Add_TextChanged({
            $var_CV_lblVMPath.Content = $var_CV_tbxFolderPath.text + "\" + $var_CV_tbxCName.text
            $var_CV_lblVHDPath.Content = $var_CV_tbxFolderPath.text + "\" + $var_CV_tbxCName.text + "\" + $var_CV_tbxCName.text + ".vhdx"
        })

    $var_CV_tbxCName.Add_TextChanged({
            $var_CV_lblVMPath.Content = $var_CV_tbxFolderPath.text + "\" + $var_CV_tbxCName.text
            $var_CV_lblVHDPath.Content = $var_CV_tbxFolderPath.text + "\" + $var_CV_tbxCName.text + "\" + $var_CV_tbxCName.text + ".vhdx"
        })

    $var_CV_CBOS.Add_SelectionChanged({ $eventVMOS = $var_CV_CBOS.selecteditem })

    $var_CV_CBVS.Add_SelectionChanged({ $eventVMVS = $var_CV_CBVS.selecteditem })

    $var_CV_btnSelectPath.Add_Click({
            $var_CV_tbxFolderPath.text = Get-Folder
            $var_CV_lblVMPath.Content = $var_CV_tbxFolderPath.text + "\" + $var_CV_tbxCName.text
            $var_CV_lblVHDPath.Content = $var_CV_tbxFolderPath.text + "\" + $var_CV_tbxCName.text + "\" + $var_CV_tbxCName.text + ".vhdx"
    })

    $var_CV_btnConfirm.Add_Click({
        if ($var_CV_tbxCName.text -eq "") { [System.Windows.MessageBox]::Show("The name of the VM can't be empty") }
        else {
            Set-LabVMProperties $var_CV_tbxFolderPath.text $var_CV_lblVMPath.Content $var_CV_lblVHDPath.Content $var_CV_tbxCName.text $var_CV_CBOS.selecteditem $eventVMVS
            $CVFormDisplay.close()   
        }
        })
    $var_CV_btnCancel.Add_Click({ $CVFormDisplay.close() })

    $CVFormDisplay.showdialog()   
    $CVFormDisplay.close()
}
function Show-MainForm{
    #### XAML Reading #####
    ###### Main Form ######
    $xamlFile = "MainWindow.xaml"
    $inputXML = Get-Content $xamlFile -Raw
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [XML]$XAML = $inputXML
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)

    try { $MainFormDisplay = [Windows.Markup.XamlReader]::Load( $reader ) }
    catch {
        Write-Warning $_.Exception
        throw
    }
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        try { Set-Variable -Name "var_$($_.Name)" -Value $MainFormDisplay.FindName($_.Name) -ErrorAction Stop }
        catch { throw }
    }
    ################### Main Form ############################
    ###### Event Listeners ###################################
    $var_btnCreateVM.Add_Click({ show-CVForm })
    $var_btnDownloadLab.Add_click({Get-ADlab})
    

    $MainFormDisplay.showdialog()
    $MainFormDisplay.close()
}
function Confirm-PropertiesFolderExist{
    Set-Location -path $env:USERPROFILE"\OneDrive - Microsoft"
    $Folder = $env:USERPROFILE+"\OneDrive - Microsoft\documents\ADLab"
    "Test to see if folder [$Folder]  exists"
    if (-not (Test-Path -Path $Folder)) {
        Set-Location -path $env:USERPROFILE"\OneDrive - Microsoft"
        New-Item "documents\ADLab" -itemType Directory 
    }
    $file = $env:USERPROFILE+"\OneDrive - Microsoft\documents\ADLab\Lablocation.ini"
    if (-not(Test-Path -Path $file -PathType Leaf)) {
        try {
            $null = New-Item -ItemType File -Path $file -Force -ErrorAction Stop
            Write-Host "The file [$file] has been created."
        }
        catch {
            throw $_.Exception.Message
        }
    }
    else {
        Write-Host "Cannot create [$file] because a file with that name already exists."
        Write-Warning $file
        $global:LabLocation=Get-Content $file
    }
    Set-Location -path $PSScriptRoot
}

#Downloads the lab on the specified location. If already done once, it will save the path and not ask again.
#The path is in $env:USERPROFILE+"\OneDrive - Microsoft\documents\ADLab\Lablocation.ini"
function Get-ADlab{ 
    Write-Warning "Get-ADLAB"
    if ($null -eq $global:LabLocation){
        $LabPath=Get-Folder
        $file = $env:USERPROFILE+"\OneDrive - Microsoft\documents\ADLab\Lablocation.ini"
        $LabPath | Out-File $file -Encoding Ascii  
        $global:LabLocation = $LabPath
        New-ADMappedDrive
    }
    explorer.exe "\\adtraininglab.file.core.windows.net\adtraininglab"
}
function New-ADMappedDrive{
    If (!(Test-Path X:)){
        $connectTestResult = Test-NetConnection -ComputerName adtraininglab.file.core.windows.net -Port 445
        if ($connectTestResult.TcpTestSucceeded) {
            # Save the password so the drive will persist on reboot
            cmd.exe /C "cmdkey /add:`"adtraininglab.file.core.windows.net`" /user:`"localhost\adtraininglab`" /pass:`"f8lqYllgIS5lA+18jLwqpCVz7M546wH/MLLfUyO/DaFf/4HUPfXK/heKz9b2vPJ1K/YMznphdd4f+AStjk8zfA==`""
            # Mount the drive
            New-PSDrive -Name Z -PSProvider FileSystem -Root "\\adtraininglab.file.core.windows.net\adtraininglab" -Persist
        } else {
            Write-Error -Message "Unable to reach the Azure storage account via port 445. Check to make sure your organization or ISP is not blocking port 445, or use Azure P2S VPN, Azure S2S VPN, or Express Route to tunnel SMB traffic over a different port."
        }
    }
}
$global:exepath=Get-Location
$global:LabLocation=$null
Get-isAdmin
Confirm-PropertiesFolderExist
Show-MainForm

