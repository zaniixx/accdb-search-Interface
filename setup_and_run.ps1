# ============================================
# نظام البحث في قواعد بيانات الناخبين العراقيين
# Iraqi Voter Database Search System
# Setup and Run Script
# ============================================

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  نظام البحث في قواعد بيانات الناخبين" -ForegroundColor Yellow
Write-Host "  Iraqi Voter Database Search System" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Check if Python is installed
Write-Host "🔍 Checking Python installation..." -ForegroundColor Green
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✅ Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Python is not installed!" -ForegroundColor Red
    Write-Host "Please download and install Python from: https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host "Make sure to check 'Add Python to PATH' during installation" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Check if pip is available
Write-Host "🔍 Checking pip installation..." -ForegroundColor Green
try {
    $pipVersion = pip --version 2>&1
    Write-Host "✅ pip found: $pipVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ pip is not available!" -ForegroundColor Red
    Write-Host "Please reinstall Python with pip included" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Get the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Host "📁 Working directory: $scriptDir" -ForegroundColor Cyan
Write-Host ""

# Check if database files exist
Write-Host "🔍 Checking database files..." -ForegroundColor Green
$dbFiles = @(
    "البصرة- الناصرية - الموصل.accdb",
    "بابل- ديالى- الديوانية-صلاح الدين-الانبار.accdb",
    "بغداد كرخ - رصافة.accdb",
    "دهوك- اربيل-السليمانية-كركوك.accdb",
    "كربلاء-ميسان- المثنى- النجف- واسط.accdb"
)

$missingFiles = @()
foreach ($dbFile in $dbFiles) {
    if (Test-Path $dbFile) {
        Write-Host "  ✅ Found: $dbFile" -ForegroundColor Green
    } else {
        Write-Host "  ❌ Missing: $dbFile" -ForegroundColor Red
        $missingFiles += $dbFile
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "⚠️  WARNING: Some database files are missing!" -ForegroundColor Yellow
    Write-Host "The system will continue but searches in missing databases will fail." -ForegroundColor Yellow
    Write-Host ""
}

# Check if required files exist
Write-Host ""
Write-Host "🔍 Checking application files..." -ForegroundColor Green
$requiredFiles = @("app.py", "templates\index.html", "templates\results.html")
$missingAppFiles = @()

foreach ($file in $requiredFiles) {
    if (Test-Path $file) {
        Write-Host "  ✅ Found: $file" -ForegroundColor Green
    } else {
        Write-Host "  ❌ Missing: $file" -ForegroundColor Red
        $missingAppFiles += $file
    }
}

if ($missingAppFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "❌ ERROR: Critical application files are missing!" -ForegroundColor Red
    Write-Host "Please ensure all files are copied to this directory." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Install required Python packages
Write-Host "📦 Installing required Python packages..." -ForegroundColor Green
Write-Host ""

$packages = @("Flask", "pyodbc")

foreach ($package in $packages) {
    Write-Host "  Installing $package..." -ForegroundColor Cyan
    try {
        pip install $package --quiet
        Write-Host "  ✅ $package installed successfully" -ForegroundColor Green
    } catch {
        Write-Host "  ⚠️  Warning: Could not install $package" -ForegroundColor Yellow
        Write-Host "  Trying to continue anyway..." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "✅ Installation complete!" -ForegroundColor Green
Write-Host ""

# Check if Microsoft Access Database Engine is installed
Write-Host "🔍 Checking Microsoft Access Database Engine..." -ForegroundColor Green
$odbcDrivers = Get-OdbcDriver | Where-Object { $_.Name -like "*Access*" -or $_.Name -like "*Microsoft Access*" }

if ($odbcDrivers) {
    Write-Host "✅ Microsoft Access ODBC Driver found" -ForegroundColor Green
} else {
    Write-Host "⚠️  WARNING: Microsoft Access Database Engine might not be installed!" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "If you encounter database connection errors, please install:" -ForegroundColor Yellow
    Write-Host "  - Microsoft Access Database Engine 2016 Redistributable" -ForegroundColor Cyan
    Write-Host "  - Download from: https://www.microsoft.com/en-us/download/details.aspx?id=54920" -ForegroundColor Cyan
    Write-Host ""
}

Write-Host ""

# Display system information
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "           SYSTEM INFORMATION" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  📂 Application Directory: $scriptDir" -ForegroundColor White
Write-Host "  🐍 Python Version: $pythonVersion" -ForegroundColor White
Write-Host "  📦 pip Version: $pipVersion" -ForegroundColor White
Write-Host "  💾 Database Files: $($dbFiles.Count - $missingFiles.Count) / $($dbFiles.Count) found" -ForegroundColor White
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Ask user if they want to start the server
Write-Host "🚀 Ready to start the application!" -ForegroundColor Green
Write-Host ""
Write-Host "The application will start on: http://127.0.0.1:5000" -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop the server when you're done" -ForegroundColor Yellow
Write-Host ""

$response = Read-Host "Do you want to start the server now? (Y/N)"

if ($response -eq 'Y' -or $response -eq 'y' -or $response -eq 'Yes' -or $response -eq 'yes' -or $response -eq '') {
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  🚀 STARTING FLASK SERVER..." -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "📍 Access the application at: http://127.0.0.1:5000" -ForegroundColor Green
    Write-Host "📍 Or from another computer: http://YOUR_IP_ADDRESS:5000" -ForegroundColor Green
    Write-Host ""
    Write-Host "Press Ctrl+C to stop the server" -ForegroundColor Yellow
    Write-Host ""
    
    # Start the Flask application
    try {
        python app.py
    } catch {
        Write-Host ""
        Write-Host "❌ Error starting the application!" -ForegroundColor Red
        Write-Host "Error details: $_" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
} else {
    Write-Host ""
    Write-Host "👋 Setup complete! Run this script again to start the server." -ForegroundColor Cyan
    Write-Host "Or run manually: python app.py" -ForegroundColor Cyan
    Write-Host ""
    Read-Host "Press Enter to exit"
}
