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
    scope         = "monitoring"
}

$TokenResponse = Invoke-RestMethod -Method POST -Uri ("$($Env["RMM_URL"])/ws/oauth/token") -ContentType "application/x-www-form-urlencoded" -Body $Body

$Token = $TokenResponse.access_token

$Headers = @{
    Authorization = "Bearer $Token"
    Accept        = "application/json"
}

$Devices = Invoke-RestMethod -Method GET -Uri "$($Env["RMM_URL"])/v2/devices" -Headers $Headers

$DevicesJSON = $Devices | ConvertTo-Json

$DevicesArray = ConvertFrom-Json $DevicesJSON

$NinjaComputerList = @()

For ($Index = 0; $Index -lt $DevicesArray.Count; $Index++) {
    $DeviceName = ($DevicesArray[$Index].dnsName -replace $($Env["DNS_NAME"]), "")

    # Check for no name to not add devices that were recently deleted
    If ($DeviceName -ne "") {
        $NinjaComputerList += ($DevicesArray[$Index].dnsName -replace $($Env["DNS_NAME"]), "")
    }
}

$ReportPath = $Env["REPORT_URL"]

$ExcelObj = Import-Excel -Path $ReportPath

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
    If ($NinjaComputerList -contains $ExcelComputer) {

    } Else {
        Write-Host $ExcelComputer "is present in Excel but not in Ninja!" -BackgroundColor DarkGray
    }
}

ForEach ($NinjaComputer in $NinjaComputerList) {
    If ($ExcelSheetsComputerList -contains $NinjaComputer) {

    } Else {
        Write-Host $NinjaComputer "is present in Ninja but not in Excel!" -BackgroundColor Blue
    }
}