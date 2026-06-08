$configPath    = Join-Path $PSScriptRoot 'config.txt'
$excelPath     = [System.IO.File]::ReadAllText($configPath, [System.Text.Encoding]::UTF8).Trim()
$outPath       = Join-Path $PSScriptRoot 'data.js'
$kintoneTokenPath = Join-Path $PSScriptRoot 'kintone_token.txt'
$kintoneSubdomain = 'symgrp'
$kintoneAppId     = '269'

$targetCols = @(1,2,3,4,5,6,7,8,9,10,11,12,13,14,18,19,20,21,22,23,24,25,26)

if (-not (Test-Path $excelPath)) {
    Write-Host 'ERROR: Excel file not found' -ForegroundColor Red
    Write-Host $excelPath; exit 1
}

# ===== [1] Excel読み込み =====
Write-Host '[1/4] Reading Excel...'
$excelRows = @()
$excelHeaders = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($excelPath, 0, $true)
    $ws = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -like '*工事履歴*' -and $s.Name -notlike '*元*') { $ws = $s; break }
    }
    if ($null -eq $ws) { $ws = $wb.Worksheets.Item(1) }
    Write-Host ('  Sheet: ' + $ws.Name) -ForegroundColor Gray
    $rowCount = $ws.UsedRange.Rows.Count
    $headerRow = -1
    for ($r = 1; $r -le [Math]::Min(10, $rowCount); $r++) {
        if ($ws.Cells.Item($r, 3).Text -like '*工事名*') { $headerRow = $r; break }
    }
    if ($headerRow -eq -1) {
        for ($r = 1; $r -le [Math]::Min(10, $rowCount); $r++) {
            if ($ws.Cells.Item($r, 3).Text -ne '') { $headerRow = $r; break }
        }
    }
    if ($headerRow -eq -1) { $headerRow = 1 }
    Write-Host ('  Header row: ' + $headerRow) -ForegroundColor Gray
    foreach ($colIdx in $targetCols) {
        if ($colIdx -eq 2) { $excelHeaders += '工事番号' }
        else {
            $h = $ws.Cells.Item($headerRow, $colIdx).Text
            $excelHeaders += if ($h -ne '') { $h } else { 'Col' + $colIdx }
        }
    }
    for ($r = $headerRow + 1; $r -le $rowCount; $r++) {
        $rowData = @{}; $hasData = $false
        for ($i = 0; $i -lt $targetCols.Count; $i++) {
            $v = $ws.Cells.Item($r, $targetCols[$i]).Text
            if ($v -ne '') { $hasData = $true }
            $rowData[$excelHeaders[$i]] = $v
        }
        if ($hasData) { $rowData['_source'] = 'excel'; $rowData['_status'] = '過去'; $excelRows += $rowData }
    }
    $wb.Close($false); $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host ('[1/4] Excel: ' + $excelRows.Count + ' rows') -ForegroundColor Green
} catch {
    Write-Host ('ERROR Excel: ' + $_.Exception.Message) -ForegroundColor Red
    try { $excel.Quit() } catch {}; exit 1
}

# ===== [2] kintone全件取得（totalCountベースのページネーション）=====
Write-Host ''; Write-Host '[2/4] Reading kintone...'
$kintoneRows = @()

if (-not (Test-Path $kintoneTokenPath)) {
    Write-Host '  WARNING: kintone_token.txt not found. Skipping.' -ForegroundColor Yellow
} else {
    $token = [System.IO.File]::ReadAllText($kintoneTokenPath, [System.Text.Encoding]::UTF8).Trim()
    $baseUrl = "https://$kintoneSubdomain.cybozu.com/k/v1/records.json"
    $fieldParams = 'fields[0]=KojiNo&fields[1]=mkbKojiName&fields[2]=mkbRyakuName&fields[3]=mkbChumonNm&fields[4]=mkbSeikyuNm&fields[5]=mkbBasyo1&fields[6]=mkbBasyo2&fields[7]=mbuBumon&fields[8]=mkuJymd&fields[9]=mkuKsYmd&fields[10]=mkuKkYmd&fields[11]=amkb_kin&fields[12]=mkbSyuruiName&fields[13]=mkbDesigner'

    try {
        # まず総件数を取得
        $query = [System.Uri]::EscapeDataString('mkuKkYmd = "" and mbuBumon like "建築"')
        $countUrl = "${baseUrl}?app=${kintoneAppId}&limit=1&totalCount=true&query=${query}&${fieldParams}"
        $req = [System.Net.WebRequest]::Create($countUrl)
        $req.Method = 'GET'
        $req.Headers.Add('X-Cybozu-API-Token', $token)
        $resp = $req.GetResponse()
        $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
        $raw = $reader.ReadToEnd(); $reader.Close(); $resp.Close()
        $totalCount = ($raw | ConvertFrom-Json).totalCount
        Write-Host ('  Total records: ' + $totalCount) -ForegroundColor Cyan

        # 100件ずつ全件取得
        $limit = 100
        $offset = 0
        while ($offset -lt $totalCount) {
            $url = "${baseUrl}?app=${kintoneAppId}&limit=${limit}&offset=${offset}&query=${query}&${fieldParams}"
            $req = [System.Net.WebRequest]::Create($url)
            $req.Method = 'GET'
            $req.Headers.Add('X-Cybozu-API-Token', $token)
            $resp = $req.GetResponse()
            $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
            $raw = $reader.ReadToEnd(); $reader.Close(); $resp.Close()
            $records = ($raw | ConvertFrom-Json).records

            foreach ($rec in $records) {
                $f = @{}
                $rec.PSObject.Properties | ForEach-Object {
                    if ($_.Value -and $_.Value.PSObject.Properties['value']) {
                        $f[$_.Name] = [string]$_.Value.value
                    }
                }
                $row = @{
                    '工事番号'   = $f['KojiNo']
                    '工事名'     = $f['mkbKojiName']
                    '工事名略称' = $f['mkbRyakuName']
                    '発注者'     = $f['mkbChumonNm']
                    '請求先'     = $f['mkbSeikyuNm']
                    '施工場所'   = (($f['mkbBasyo1'] + ' ' + $f['mkbBasyo2']).Trim())
                    '担当部門'   = $f['mbuBumon']
                    '受注日'     = $f['mkuJymd']
                    '着工日'     = $f['mkuKsYmd']
                    '完成日'     = $f['mkuKkYmd']
                    '請負金額'   = $f['amkb_kin']
                    '工事種別'   = $f['mkbSyuruiName']
                    '設計者'     = $f['mkbDesigner']
                    '_source'    = 'kintone'
                    '_status'    = '稼働中'
                }
                $kintoneRows += $row
            }

            $offset += $limit
            Write-Host ('  ' + $offset + ' / ' + $totalCount + ' 件取得済み') -ForegroundColor Gray
        }
        Write-Host ('[2/4] kintone: ' + $kintoneRows.Count + ' rows') -ForegroundColor Green
    } catch {
        Write-Host ('  WARNING kintone: ' + $_.Exception.Message) -ForegroundColor Yellow
    }
}

# ===== [3] マージ（重複排除）=====
Write-Host ''; Write-Host '[3/4] Merging...'
$kintoneNoSet = @{}; $kintoneNameSet = @{}
foreach ($kr in $kintoneRows) {
    $no = $kr['工事番号']; $nm = $kr['工事名']
    if ($no -ne '') { $kintoneNoSet[$no.Trim()] = $true }
    if ($nm -ne '') { $kintoneNameSet[$nm.Substring(0, [Math]::Min(10, $nm.Length))] = $true }
}
$filteredExcel = @()
$dedupCount = 0
foreach ($er in $excelRows) {
    $no  = if ($er.ContainsKey('工事番号')) { $er['工事番号'].Trim() } else { '' }
    $nm  = if ($er.ContainsKey('工事名'))   { $er['工事名'] }           else { '' }
    $nmk = $nm.Substring(0, [Math]::Min(10, $nm.Length))
    if (($no -ne '') -and $kintoneNoSet.ContainsKey($no) -and ($nmk -ne '') -and $kintoneNameSet.ContainsKey($nmk)) {
        $dedupCount++
    } else {
        $filteredExcel += $er
    }
}
Write-Host ('  重複除外: ' + $dedupCount + '件（kintone優先）') -ForegroundColor Yellow
$merged = $kintoneRows + $filteredExcel
Write-Host ('[3/4] Total: ' + $merged.Count + ' rows') -ForegroundColor Green

# ===== [4] data.js出力 =====
Write-Host ''; Write-Host '[4/4] Generating data.js...'
$kintoneHeaderSet = @('工事番号','工事名','工事名略称','発注者','請求先','施工場所','担当部門','受注日','着工日','完成日','請負金額','工事種別','設計者')
$allHeaders = $kintoneHeaderSet + ($excelHeaders | Where-Object { $kintoneHeaderSet -notcontains $_ })
$allSheets = @{ '工事一覧' = @{ headers = $allHeaders; rows = $merged } }
$jsonObj = $allSheets | ConvertTo-Json -Depth 10 -Compress
$updateTime = (Get-Date).ToString('yyyy/MM/dd HH:mm')
$out = "// auto-generated ($updateTime)`nwindow.KOUJI_DATA = $jsonObj;`nwindow.KOUJI_UPDATE_TIME = '$updateTime';"
[System.IO.File]::WriteAllText($outPath, $out, [System.Text.Encoding]::UTF8)
Write-Host '[4/4] data.js generated' -ForegroundColor Green

Write-Host ''; Write-Host 'Pushing to GitHub...'
Set-Location $PSScriptRoot
$gitExe = 'C:\Program Files\Git\cmd\git.exe'
if (-not (Test-Path $gitExe)) { $gitExe = 'git' }
& $gitExe add data.js 2>&1 | Out-Null
& $gitExe commit -m "update: $updateTime" 2>&1 | Out-Null
$pushResult = & $gitExe push 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host ('ERROR push: ' + ($pushResult -join '')) -ForegroundColor Red; exit 1
}
Write-Host 'Pushed to GitHub' -ForegroundColor Green
Write-Host ''; Write-Host 'Done! ~1 minute to reflect.' -ForegroundColor Green
Write-Host 'URL: https://ryosuke-ema.github.io/kouji-search/'