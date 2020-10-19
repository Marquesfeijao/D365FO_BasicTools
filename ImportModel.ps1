#region Assemblies
Add-Type -AssemblyName System
Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
#endregion

#region Variables
$ActionFGColor = "White"
$ActionBGColor = "DarkGreen"
$FinishedFGColor = "Green"
$ExecutionStartTime = $(Get-Date)

$TaskStartTime = $(Get-Date)

$ISVModelToImport = 'C:\temp\Model'
#endregion


#region Methods
function StartImport($StartProcess) {
    $ElapsedTimeFGColor = "Cyan"
    Write-Host "****** Starting Process ******" -ForegroundColor $ElapsedTimeFGColor    
}

function ElapsedTime($TaskStartTime) {
    $ElapsedTimeFGColor = "Cyan"
    $ElapsedTime = New-TimeSpan $TaskStartTime $(Get-Date)
    Write-Host "Elapsed time:$($ElapsedTime.ToString("hh\:mm\:ss"))" -ForegroundColor $ElapsedTimeFGColor
}

function Finished($StartTime) {
    Write-Host "Finished!" -ForegroundColor $FinishedFGColor
    ElapsedTime $StartTime
}

function Pause ($message) {
    # Check if running Powershell ISE
    if ($psISE) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$message")
    }
    else {
        Write-Host "$message" -ForegroundColor Yellow
        $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

function StartServices() {
    Write-Host ""
    Write-Host "*** Starting services... ***" -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host ""

    $TaskStartTime = $(Get-Date)

    Get-Service W3SVC `
        , aspnet_state `
    | Start-Service

    ElapsedTime $TaskStartTime

    iisreset.exe
}

function StopServices() {
    Write-Host ""
    Write-Host "*** Stopping services... ***" -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host ""

    $TaskStartTime = $(Get-Date)

    Get-Service W3SVC `
        , aspnet_state `
    | Stop-Service -Force

    Finished $TaskStartTime
}

function CheckServices($ServicesStatus) {

    Get-Service W3SVC `
        , aspnet_state `
    | Where-Object {
        switch ($ServicesStatus) {
            "Running" { 
                if ($_.Status -eq "Running") {
                    StopServices
                }
            }
            "Stopped" { 
                if ($_.Status -eq "Stopped") {
                    StartServices
                }
            }
        }
    }
}

function OpenFileDialog () {

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter           = 'Models (*.axmodel)| *.axmodel'
        Multiselect      = 1
    }

    $FileBrowser.ShowDialog()
    
    if (![String]::IsNullOrEmpty($FileBrowser.FileName)) {
        ForEach ($fileName in $FileBrowser.FileNames) {
            Copy-Item $fileName -Destination $ISVModelToImport
        }
    }

    return $FileBrowser.SafeFileNames
}

function Import-DAXModel($ImportFileNames) {
    $ReturnImport = ""
    ForEach ($FileName in $ImportFileNames) {
        if ($FileName -ne "Ok") {
            $ReturnImport += Invoke-ImportModelUtil $FileName
        }
    }

    return $ReturnImport
}

function Invoke-ImportModelUtil($ModelFileName) {

    $ModelStore     = "$($env:SystemDrive)\AOSService\PackagesLocalDirectory"
    $BinPath        = Join-Path -Path $ModelStore -ChildPath "Bin"
    $ModelUtil      = Join-Path -Path $BinPath -ChildPath "ModelUtil.exe"    
    $ModelToimport  = Join-Path -Path $ISVModelToImport -ChildPath $ModelFileName

    $ReplaceArgs = @("-replace";
        "-MetadataStorePath=`"$ModelStore`"";
        "-file=`"$ModelToimport`"";
        "-force"
        ) 
    & $ModelUtil $ReplaceArgs
}
#endregion

StartImport $(Get-Date)

StopServices

$ImportFileNames = OpenFileDialog

if ($ImportFileNames[0] -eq "Ok") {
    $ReturnModels = Import-DAXModel $ImportFileNames 
    
    foreach ($modelImported in $ReturnModels) {
        Write-Host $modelImported
    }
}

StartServices


# $messageBox = [System.Windows.MessageBox]::Show('Do you want to stop the services?','Stop the services','YesNoCancel',[System.Windows.MessageBoxImage]::Exclamation)

# if ($messageBox -eq [System.Windows.MessageBoxResult]::Yes)
# {
#     [System.Windows.MessageBox]::Show('Services stopped', 'Notication', 'OK', 'Info');
# }
# else {
#     [System.Windows.MessageBox]::Show('No!!!');
# }
# StopServices 

# Pause("Press any key to start the services...")

# StartServices

Write-Host "*****************" -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
Write-Host "*** All done! ***" -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
Write-Host "*****************" -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor

ElapsedTime $ExecutionStartTime

# Pause("Press any key to continue...")