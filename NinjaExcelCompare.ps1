$EnvFilePath = ".\.env.local"

If (-not (Test-Path $EnvFilePath -PathType Leaf)) {
    Write-Error "Error: .env.local file not found at $($EnvFilePath)"
    Exit 1
}

$Env = @{}
$EnvLines = Get-Content $EnvFilePath

ForEach ($Line in $EnvLines) {
    $TrimmedLine = $Line.Trim()

    If ([string]::IsNullOrEmpty($TrimmedLine) -or $TrimmedLine.StartsWith('#')) {
        Continue
    }

    $Name, $Value = $TrimmedLine -split '=', 2
    $Env.Add($Name, $Value)
}

$Body = @{
    grant_type    = "client_credentials"
    client_id     = $Env["CLIENT_ID"]
    client_secret = $Env["CLIENT_SECRET"]
    scope         = "monitoring management control"
}

$TokenResponse = Invoke-RestMethod -Method POST -Uri ("$($Env["RMM_URL"])/ws/oauth/token") -ContentType "application/x-www-form-urlencoded" -Body $Body
$Token = $TokenResponse.access_token

$Headers = @{
    Authorization = "Bearer $Token"
    Accept        = "application/json"
}

$Devices = Invoke-RestMethod -Method GET -Uri "$($Env["RMM_URL"])/v2/devices-detailed" -Headers $Headers
$DevicesJSON = $Devices | ConvertTo-Json
$DevicesArray = ConvertFrom-Json $DevicesJSON

$NinjaComputerList = @()

For ($Index = 0; $Index -lt $DevicesArray.Count; $Index++) {
    If ($null -ne $DevicesArray[$Index].displayName) {
        $DeviceName = $DevicesArray[$Index].displayName
    } Else {
        $DeviceName = $DevicesArray[$Index].systemName
    }

    # Check for no name to not add devices that were recently deleted
    If ($DeviceName -ne "") {
        $NinjaComputerList += $DeviceName
    }
}

$ReportPath = $Env["REPORT_URL"]
$ExcelPackage = Open-ExcelPackage -Path $ReportPath
$SheetNames = $ExcelPackage.Workbook.Worksheets.Name
$ExcelSheetsComputerList = @()
$TrackedSheets = @($Env["TRACKED_SHEETS"].Split(','))

ForEach ($Sheet in $SheetNames) {
    If ($TrackedSheets -contains $Sheet) {
        $Names = Import-Excel -Path $ReportPath -WorksheetName $Sheet | Select-Object -ExpandProperty Name
        $ExcelSheetsComputerList += $Names
    }   
}

ForEach ($ExcelComputer in $ExcelSheetsComputerList) {
    If (-not ($NinjaComputerList -contains $ExcelComputer)) {
        Write-Host $ExcelComputer "is present in Excel but not in Ninja!" -BackgroundColor Green
    }
}

ForEach ($NinjaComputer in $NinjaComputerList) {
    If (-not ($ExcelSheetsComputerList -contains $NinjaComputer)) {
        Write-Host $NinjaComputer "is present in Ninja but not in Excel!" -BackgroundColor Cyan
    }
}