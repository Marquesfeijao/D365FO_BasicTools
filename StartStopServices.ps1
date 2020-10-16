$ActionFGColor      = "White"
$ActionBGColor      = "DarkGreen"
$FinishedFGColor    = "Green"
$ExecutionStartTime = $(Get-Date)

$TaskStartTime = $(Get-Date)

#region Methods
function StartImport($StartProcess){
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

function StartServices(){
    Write-Host ""
    Write-Host "*** Starting services... ***" -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host ""

    $TaskStartTime = $(Get-Date)

    Get-Service DynamicsAxBatch `
        , Microsoft.Dynamics.AX.Framework.Tools.DMF.SSISHelperService.exe `
        , W3SVC `
        , MR2012ProcessService `
        , aspnet_state `
    | Start-Service

    ElapsedTime $TaskStartTime

    iisreset.exe
}

function StopServices(){
    Write-Host ""
    Write-Host "*** Stopping services... ***" -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
    Write-Host ""

    $TaskStartTime = $(Get-Date)

    Get-Service DynamicsAxBatch `
        , Microsoft.Dynamics.AX.Framework.Tools.DMF.SSISHelperService.exe `
        , W3SVC `
        , MR2012ProcessService `
        , aspnet_state `
    | Stop-Service -Force

    Finished $TaskStartTime
}
#endregion

StartImport $(Get-Date)

StopServices 

Pause("Press any key to start the services...")

StartServices

Write-Host ""
Write-Host "*** All done! ***" -ForegroundColor $ActionFGColor -BackgroundColor $ActionBGColor
Write-Host ""

ElapsedTime $ExecutionStartTime