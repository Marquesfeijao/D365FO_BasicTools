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
$ElapsedTimeFGColor = "Cyan"

$ExecutionStartTime = $(Get-Date)
$TaskStartTime = $(Get-Date)

$ISVModelToImport = 'C:\temp\Model'
#endregion

#region Enum
Add-Type -TypeDefinition @"
public enum RunProcess{
    StartService        = 1,
    StopService         = 2,
    StartImportModel    = 3,
    FinishedImportModel = 4,
    AllDone             = 5
}
"@
#endregion

#region Titles Methods
function StartImport($StartProcess) {
    Write-Host "                                                                                                                        "   -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host "************************************************************************************************************************"   -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host "***************************************** Starting the process to import license ***************************************"   -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host "************************************************************************************************************************"   -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host "************************************************************************************************************************"   -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host "*********** | Start at: '$(Get-Date)' |"                                                                                    -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host "                                                                                                                        "   -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
}

function Get-Title($Title) {

    Write-Host "                                                                                                                        "   -ForegroundColor $ElapsedTimeFGColor
    Write-Host "************************************************************************************************************************"   -ForegroundColor $ElapsedTimeFGColor
    Write-Host "                                                                                                                        "   -ForegroundColor $ElapsedTimeFGColor
    Write-Host "*********** | $Title |"                                                                                                     -ForegroundColor $ElapsedTimeFGColor
    Write-Host "*********** | Start at $(Get-Date) |"                                                                                       -ForegroundColor $ElapsedTimeFGColor
    Write-Host "                                                                                                                        "   -ForegroundColor $ElapsedTimeFGColor
    Write-Host "************************************************************************************************************************"   -ForegroundColor $ElapsedTimeFGColor
    Write-Host "                                                                                                                        "   -ForegroundColor $ElapsedTimeFGColor
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
#endregion Titles Methods

#region Methods
function Pause($message) {
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

function Read-Title($runProcess) {
    switch ($runProcess) {
        [RunProcess]::StartService { "Start Service" }
        [RunProcess]::StopService { "Stop Service" }
        [RunProcess]::StartImportModel { "Start Import Service" }
        [RunProcess]::FinishedImportModel { "Finished Import Service" }
        [RunProcess]::AllDone { "All Done" }
        Default {}
    }
}

function StartServices() {
    $Title = Read-Title [RunProcess]::StartService
    Get-Title $Title

    $TaskStartTime = $(Get-Date)

    Get-Service W3SVC `
        , aspnet_state `
    | Start-Service

    ElapsedTime $TaskStartTime
}

function StopServices() {
    $Title = Read-Title [RunProcess]::StopService
    Get-Title $Title

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

function OpenFileDialog() {

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
    $Title = Read-Title [RunProcess]::StartImportModel
    Get-Title $Title

    $ReturnImport = ""
    ForEach ($FileName in $ImportFileNames) {
        if ($FileName -ne "Ok") {
            $ReturnImport += Invoke-ImportModelUtil $FileName
        }
    }

    return $ReturnImport
}

function Invoke-ImportModelUtil($ModelFileName) {

    $ModelStore = "$($env:SystemDrive)\AOSService\PackagesLocalDirectory"
    $BinPath = Join-Path -Path $ModelStore -ChildPath "Bin"
    $ModelUtil = Join-Path -Path $BinPath -ChildPath "ModelUtil.exe"    
    $ModelToimport = Join-Path -Path $ISVModelToImport -ChildPath $ModelFileName

    $ReplaceArgs = @("-delete";
        "-MetadataStorePath=`"$ModelStore`"";
        "-modelname=`"$ModelFileName`"";
        "-force"
    ) 
    & $ModelUtil $ReplaceArgs
}
#endregion

StopServices
Invoke-ImportModelUtil "FourVisionHRPlus"
StartServices
# # Show the Title 
# StartImport $(Get-Date)

# #Open dialog
# $ImportFileNames = OpenFileDialog

# #region Start import model
# if ($ImportFileNames[0] -eq "Ok") {
#     StopServices

#     $ReturnModels = Import-DAXModel $ImportFileNames 
    
#     foreach ($modelImported in $ReturnModels) {
#         Write-Host $modelImported
#     }

#     $Title = Read-Title [RunProcess]::FinishedImportModel
#     Get-Title $Title

#     StartServices

#     $Title = Read-Title [RunProcess]::AllDone
#     Get-Title $Title
# }
# #endregion Start import model

ElapsedTime $ExecutionStartTime

Pause("Press any key to continue...")

