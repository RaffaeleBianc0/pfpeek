# TODO:
# - Scrivi in un file .\rbPfPeek-log.tsv in append una riga nel formato: <timestamp> INFO <tutte le statistiche riepilogative del portafoglio, separandole con un TAB>

# --- PORTFOLIO CONFIGURATION ---
# Choose the portfolio source:
#   $true  -> rebuilds it by reading the "Movimenti" CSV export from Directa Trading
#   $false -> uses the static $manualAssets list below
$useCsvPortfolio = $true

# --- Option A: Portfolio from CSV (used if $useCsvPortfolio = $true) ---
# - If $csvPath is specified, that exact file is used.
# - If $csvPath is $null, the script automatically searches for the most recent
#   "Movimenti_*.csv" file in the script directory (or current directory).
$csvPath = $null   # e.g., "C:\Path\Movimenti_J1234_20-7-2026.csv"
$tickerSuffix = ".MI" # Market suffix to add to the tickers read from the CSV (Directa exports them without suffix)

# --- Option B: Manual portfolio (used if $useCsvPortfolio = $false) ---
# NOTE: If $useCsvPortfolio = $true, the order of this list defines the output sorting!
$manualAssets = @(
    @{ Ticker = "XD9U.MI"; AvgCost = 172.55; Qty = 49.00 }
    @{ Ticker = "EXUS.MI"; AvgCost = 33.27;  Qty = 156.00 }
    @{ Ticker = "EIMI.MI"; AvgCost = 34.64;  Qty = 43.00 }
    @{ Ticker = "WGLD.MI"; AvgCost = 261.57; Qty = 3.00 }
    @{ Ticker = "XEON.MI"; AvgCost = 144.22; Qty = 35.00 }
)

# Custom start date for the second chart (typically: first investment date)
$startDateConfig = "05/08/2024"
# ---------------------------------------------------------------------------

# --- MINIMUM CONSOLE SIZE CHECK ---
$minWidth = 114
$consoleWidth = $minWidth
if ($Host.UI.RawUI.WindowSize.Width) { $consoleWidth = $Host.UI.RawUI.WindowSize.Width }
if ($consoleWidth -lt $minWidth) {
    Write-Warning "Console window is too narrow ($consoleWidth chars)."
    Write-Warning "Minimum width required: $minWidth chars."
    Write-Warning "Please increase the console width (or decrease the font size) and try again."
    exit
}

# --- PROGRESS TRACKING ---
function Show-Progress {
    param (
        [int]$Current,
        [int]$Total,
        [string]$Activity = "Processing",
        [int]$BarWidth = 30
    )

    if ($Total -le 0) { return }

    $filled = [math]::Min($BarWidth, [math]::Max(0, [math]::Round(($Current / $Total) * $BarWidth)))
    $empty = $BarWidth - $filled

    $bar = "#" * $filled + " " * $empty
    
    Write-Host "[" -NoNewline
    Write-Host "$bar" -ForegroundColor White -NoNewline
    Write-Host "] (" -NoNewline
    Write-Host "$Current" -ForegroundColor White -NoNewline
    Write-Host "/" -NoNewline
    Write-Host "$Total" -ForegroundColor White -NoNewline
    Write-Host ") " -NoNewline
    Write-Host "$Activity" -ForegroundColor Cyan

    # If completed, print newline
    if ($Current -ge $Total) {
        Write-Host ""
    }
}
$totalSteps = 6

# --- STEP 1: INITIALIZING & PREREQUISITES CHECK ---
Show-Progress 1 $totalSteps "Initializing & checking prerequisites"

# 1. PowerShell Version Check (Requires PS 5.1+, works on both Windows PowerShell and PowerShell 7+)
$psVer = $PSVersionTable.PSVersion
if ($psVer.Major -lt 5 -or ($psVer.Major -eq 5 -and $psVer.Minor -lt 1)) {
    Write-Error "ERROR: This script requires PowerShell 5.1 or higher to run properly."
    Write-Warning "Your current version is: $psVer"
    Write-Warning "Please download and install the latest PowerShell from: https://aka.ms/powershell-release"
    exit
}

# $IsWindows exists only on PowerShell 6+ (Core). On Windows PowerShell 5.1 ("Desktop" edition)
# it is always Windows, since that edition only runs on Windows.
$IsWindowsHost = if ($PSVersionTable.PSEdition -eq 'Desktop') { $true } else { $IsWindows }

# 2. Console Color / ANSI Escape Sequences Support Check
if ($IsWindowsHost) {
    $MemberSignature = '[DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr GetStdHandle(int nStdHandle);' +
                       '[DllImport("kernel32.dll", SetLastError=true)] public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);' +
                       '[DllImport("kernel32.dll", SetLastError=true)] public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);'
    $Type = Add-Type -MemberDefinition $MemberSignature -Name "Win32Console" -Namespace "Win32" -PassThru
    $StdOutHandle = $Type::GetStdHandle(-11) # STD_OUTPUT_HANDLE
    $Mode = 0
    if ($Type::GetConsoleMode($StdOutHandle, [ref]$Mode)) {
        $Mode = $Mode -bor 4 # ENABLE_VIRTUAL_TERMINAL_PROCESSING
        if (-not $Type::SetConsoleMode($StdOutHandle, $Mode)) {
            Write-Warning "WARNING: Failed to enable Virtual Terminal Processing (ANSI colors might not render correctly)."
            Write-Warning "Please use Windows Terminal for the best visual experience: https://aka.ms/terminal"
        }
    }
}

# 3. Braille Character & UTF-8 Console Output Encoding Check
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $testBraille = [char]0x28FF
    if ([string]::IsNullOrEmpty($testBraille) -or $testBraille -eq '?') {
        throw "Failed to map Braille Unicode range."
    }
} catch {
    Write-Error "ERROR: The console does not support UTF-8 or Braille Unicode rendering."
    Write-Warning "To fix this, configure your console font to a compatible TrueType font (e.g., Cascadia Code, JetBrains Mono, or Fira Code)."
    Write-Warning "Get modern developer fonts at: https://www.nerdfonts.com"
    exit
}

# 4. Internet Connectivity Check (to Yahoo Finance)
$pingTarget = "fc.yahoo.com"
try {
    $ip = [System.Net.Dns]::GetHostAddresses($pingTarget)
    if ($null -eq $ip -or $ip.Count -eq 0) { throw "DNS resolution failed." }
} catch {
    Write-Error "ERROR: No internet connection detected or Yahoo Finance is unreachable."
    Write-Warning "Please check your network settings and proxy configuration."
    exit
}

# --------------------------------------------------

# --- STEP 2: BUILD PORTFOLIO (FROM CSV OR MANUAL) ---
if ($useCsvPortfolio) {
    Show-Progress 2 $totalSteps "Reading portfolio from Directa CSV export"

    # 1. Locate the CSV file
    if (-not $csvPath) {
        $searchDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $latestCsv = Get-ChildItem -Path $searchDir -Filter "Movimenti_*.csv" -File -ErrorAction SilentlyContinue |
                     Sort-Object LastWriteTime -Descending |
                     Select-Object -First 1
        if (-not $latestCsv) {
            Write-Error "ERROR: No 'Movimenti_*.csv' file found in '$searchDir' and no `$csvPath specified."
            Write-Warning "Set `$csvPath to the full path of your Directa export, or copy the CSV into the script folder."
            Write-Warning "Alternatively, set `$useCsvPortfolio = `$false to use the manual portfolio."
            exit
        }
        $csvPath = $latestCsv.FullName
    }
    if (-not (Test-Path $csvPath)) {
        Write-Error "ERROR: CSV file not found: $csvPath"
        exit
    }

    # 2. Helper to convert Italian-formatted numbers (e.g., "1.234,56" -> 1234.56)
    function ConvertTo-ItalianDouble ($text) {
        if ([string]::IsNullOrWhiteSpace($text)) { return 0.0 }
        $clean = $text.Trim().Replace(".", "").Replace(",", ".")
        $val = 0.0
        [void][double]::TryParse($clean, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$val)
        return $val
    }

    # 3. Read CSV and find the transaction header row
    $csvLines = Get-Content -Path $csvPath -Encoding UTF8
    $headerIdx = -1
    for ($i = 0; $i -lt $csvLines.Count; $i++) {
        if ($csvLines[$i] -match '^Data operazione;') { $headerIdx = $i; break }
    }
    if ($headerIdx -lt 0) {
        Write-Error "ERROR: Header 'Data operazione' not found in CSV: $csvPath"
        Write-Warning "Ensure the file is an original Directa Trading 'Movimenti' export."
        exit
    }

    $csvHeaderFields = @(
        "DataOperazione","DataValuta","TipoOperazione","Ticker","Isin",
        "Protocollo","Descrizione","Quantita","ImportoEuro","ImportoDivisa","Divisa","RiferimentoOrdine"
    )
    $movimenti = $csvLines[($headerIdx + 1)..($csvLines.Count - 1)] |
                 Where-Object { $_.Trim().Length -gt 0 } |
                 ConvertFrom-Csv -Delimiter ";" -Header $csvHeaderFields

    # 4. Aggregate transactions per ticker
    $positions = @{}
    foreach ($m in $movimenti) {
        $ticker = $m.Ticker
        if ([string]::IsNullOrWhiteSpace($ticker)) { continue }

        if (-not $positions.ContainsKey($ticker)) {
            $positions[$ticker] = @{ BoughtQty = 0.0; BoughtCost = 0.0; SoldQty = 0.0 }
        }

        $qty = ConvertTo-ItalianDouble $m.Quantita
        $importo = ConvertTo-ItalianDouble $m.ImportoEuro

        switch ($m.TipoOperazione) {
            "Acquisto"    { $positions[$ticker].BoughtQty  += $qty; $positions[$ticker].BoughtCost += [math]::Abs($importo) }
            "Commissioni" { $positions[$ticker].BoughtCost += [math]::Abs($importo) }
            "Vendita"     { $positions[$ticker].SoldQty    += $qty }
            default { }
        }
    }

    # 5. Build $assets sorted according to $manualAssets sequence
    # Map custom index for each full ticker (e.g., "XD9U.MI" -> 0)
    $customOrderMap = @{}
    for ($i = 0; $i -lt $manualAssets.Count; $i++) {
        $customOrderMap[$manualAssets[$i].Ticker] = $i
    }

    # Map and sort tickers read from CSV
    $sortedCsvTickers = $positions.Keys | Sort-Object {
        $fullTicker = "$_$tickerSuffix"
        if ($customOrderMap.ContainsKey($fullTicker)) {
            return $customOrderMap[$fullTicker]
        } else {
            return 99999 # Place at bottom if ticker from CSV is not in $manualAssets
        }
    }

    $assets = @()
    foreach ($ticker in $sortedCsvTickers) {
        $p = $positions[$ticker]
        $netQty = $p.BoughtQty - $p.SoldQty
        if ($p.BoughtQty -le 0 -or $netQty -le 0.0001) { continue }

        $avgCost = $p.BoughtCost / $p.BoughtQty
        $assets += @{
            Ticker  = "$ticker$tickerSuffix"
            AvgCost = [math]::Round($avgCost, 4)
            Qty     = [math]::Round($netQty, 4)
        }
    }

    if ($assets.Count -eq 0) {
        Write-Error "ERROR: No open positions found in CSV: $csvPath"
        exit
    }
} else {
    Show-Progress 2 $totalSteps "Loading manual portfolio configuration"
    $assets = $manualAssets
}

# --- PORTFOLIO METRICS (available only when using Directa CSV export) ---
$csvMetricsAvailable = $useCsvPortfolio
if ($csvMetricsAvailable) {

    $numeroAcquisti = @($movimenti | Where-Object { $_.TipoOperazione -eq 'Acquisto' }).Count
    $numeroVendite  = @($movimenti | Where-Object { $_.TipoOperazione -eq 'Vendita' }).Count

    function Get-SommaImporti ($tipoOperazione) {
        $tot = ($movimenti | Where-Object { $_.TipoOperazione -eq $tipoOperazione } |
                ForEach-Object { ConvertTo-ItalianDouble $_.ImportoEuro } |
                Measure-Object -Sum).Sum
        if ($null -eq $tot) { return 0.0 }
        return $tot
    }

    $totaleCommissioni    = [math]::Abs((Get-SommaImporti "Commissioni"))
    $totaleBollo          = [math]::Abs((Get-SommaImporti "Bollo portafoglio titoli*"))
    $totaleRitenute       = [math]::Abs((Get-SommaImporti "Rit. etf"))
    $totaleCostiSostenuti = $totaleCommissioni + $totaleBollo + $totaleRitenute

    $totaleVersato   = Get-SommaImporti "Conferimento con bonifico"
    $totalePrelevato = [math]::Abs((Get-SommaImporti "Prelievo bonifico"))

    $numeroTitoliMovimentati = $positions.Keys.Count
    $numeroPosizioniAperte   = $assets.Count
    $numeroPosizioniChiuse   = $numeroTitoliMovimentati - $numeroPosizioniAperte

    $dataAcquisti = @()
    foreach ($m in $movimenti) {
        if ($m.TipoOperazione -eq 'Acquisto') {
            $d = [DateTime]::MinValue
            if ([DateTime]::TryParseExact($m.DataOperazione, "dd-MM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
                $dataAcquisti += $d
            }
        }
    }
    $dataPrimoInvestimento = if ($dataAcquisti.Count -gt 0) { ($dataAcquisti | Sort-Object)[0] } else { $null }
    $giorniInvestimento = if ($dataPrimoInvestimento) { [int]([DateTime]::Today - $dataPrimoInvestimento).TotalDays } else { 0 }

    # ============================================================
    # CORRECTED CASHFLOW EXTRACTION FOR MWRR
    # ============================================================
    # IMPORTANT: Cash flows are from the INVESTOR's perspective:
    #   - Deposits (Conferimenti): NEGATIVE (money leaves investor)
    #   - Withdrawals (Prelievi): POSITIVE (money returns to investor)
    #   - Costs (Commissioni, Bollo, Rit.etf): NEGATIVE
    #   - Final value: POSITIVE (liquidation proceeds)
    # ============================================================
    $cashflows = New-Object System.Collections.Generic.List[PSCustomObject]

    foreach ($m in $movimenti) {
        $d = [DateTime]::MinValue
        if (-not [DateTime]::TryParseExact($m.DataOperazione, "dd-MM-yyyy", 
            [System.Globalization.CultureInfo]::InvariantCulture, 
            [System.Globalization.DateTimeStyles]::None, [ref]$d)) { 
            continue 
        }
        
        $amount = ConvertTo-ItalianDouble $m.ImportoEuro
        
        switch ($m.TipoOperazione) {
            # NEGATIVE: money OUT from investor (deposits into portfolio)
            "Conferimento con bonifico" {
                $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = -[math]::Abs($amount) })
            }
            # POSITIVE: money IN to investor (withdrawals from portfolio)
            "Prelievo bonifico" {
                $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = [math]::Abs($amount) })
            }
            # NEGATIVE: costs (fees, stamp duty, taxes)
            "Commissioni" {
                $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = -[math]::Abs($amount) })
            }
            "Bollo portafoglio titoli*" {
                $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = -[math]::Abs($amount) })
            }
            "Rit. etf" {
                $cashflows.Add([PSCustomObject]@{ Date = $d; Amount = -[math]::Abs($amount) })
            }
            # INTERNAL OPERATIONS (no net cash impact - excluded from XIRR)
            "Acquisto" { continue }
            "Vendita" { continue }
            default { continue }
        }
    }

    # Sort cashflows by date
    $cashflows = [System.Collections.Generic.List[PSCustomObject]]($cashflows | Sort-Object Date)

    # ============================================================
    # ROBUST XIRR CALCULATION FUNCTION
    # ============================================================
    function Get-XirrRobust ($cashflows) {
        if ($cashflows.Count -lt 2) { return $null }
        
        $sorted = @($cashflows | Sort-Object Date)
        $baseDate = $sorted[0].Date

        function Get-Npv ($rate) {
            $sum = 0.0
            foreach ($cf in $sorted) {
                $days = ($cf.Date - $baseDate).TotalDays
                if ($days -eq 0) {
                    $sum += $cf.Amount
                } else {
                    $sum += $cf.Amount / [math]::Pow((1 + $rate), ($days / 365.0))
                }
            }
            return $sum
        }

        # Try different initial guesses
        $guesses = @(-0.99, -0.5, -0.1, 0.0, 0.01, 0.05, 0.1, 0.15, 0.2, 0.3, 0.5, 1.0, 2.0, 5.0)
        
        foreach ($guess in $guesses) {
            $rate = $guess
            for ($iter = 0; $iter -lt 200; $iter++) {
                $npv = Get-Npv $rate
                
                # Derivative approximation
                $epsilon = 0.00001
                $npvEps = Get-Npv ($rate + $epsilon)
                $derivative = ($npvEps - $npv) / $epsilon
                
                if ([math]::Abs($derivative) -lt 1e-12) { break }
                
                $newRate = $rate - $npv / $derivative
                
                # Check for convergence
                if ([math]::Abs($newRate - $rate) -lt 1e-10) {
                    # Validate result
                    if ($newRate -gt -0.999 -and $newRate -lt 100) {
                        return $newRate
                    }
                    break
                }
                
                # Prevent overflow
                if ($newRate -lt -0.999 -or $newRate -gt 100) { break }
                
                $rate = $newRate
            }
        }
        
        # Fallback: try binary search if Newton failed
        $low = -0.99
        $high = 10.0
        $npvLow = Get-Npv $low
        $npvHigh = Get-Npv $high
        
        if ($npvLow * $npvHigh -ge 0) { return $null }
        
        for ($iter = 0; $iter -lt 100; $iter++) {
            $mid = ($low + $high) / 2
            $npvMid = Get-Npv $mid
            
            if ([math]::Abs($npvMid) -lt 1e-8) { return $mid }
            
            if ($npvLow * $npvMid -lt 0) {
                $high = $mid
                $npvHigh = $npvMid
            } else {
                $low = $mid
                $npvLow = $npvMid
            }
        }
        
        return $null
    }
    # ============================================================
}

# 1. ROBUST PARSING OF CONFIG DATE
$dateParsed = [DateTime]::MinValue
if (-not [DateTime]::TryParseExact($startDateConfig, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dateParsed)) {
    Write-Error "ERROR: Configured start date ($startDateConfig) is invalid. Use dd/MM/yyyy format."
    exit
}
$customStartUnix = (New-Object DateTimeOffset ($dateParsed)).ToUnixTimeSeconds()
$nowUnix = (New-Object DateTimeOffset ([DateTime]::UtcNow)).ToUnixTimeSeconds()

$tickers = ($assets.Ticker) -join ","

# Initialize web session
$userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
$webSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$webSession.Headers.Add("User-Agent", $userAgent)

# --- DATA INTERPOLATION FUNCTION (Simplified Linear Interpolation) ---
function Invoke-InterpolateData ($originalValues, $targetSize) {
    if ($originalValues.Count -lt 2) {
        return ,([double]$originalValues[0]) * $targetSize
    }
    
    $interpolated = [double[]]::new($targetSize)
    $origCount = $originalValues.Count

    for ($i = 0; $i -lt $targetSize; $i++) {
        $percent = $i / ($targetSize - 1)
        $exactIdx = $percent * ($origCount - 1)
        
        $lowIdx = [math]::Floor($exactIdx)
        $highIdx = [math]::Ceiling($exactIdx)
        $weight = $exactIdx - $lowIdx

        $lowVal = [double]$originalValues[$lowIdx]
        $highVal = [double]$originalValues[$highIdx]

        $interpolated[$i] = $lowVal + ($highVal - $lowVal) * $weight
    }
    return $interpolated
}

# --- BRAILLE SPARKLINE CHART GENERATOR ---
function Get-BrailleSparkline ($values, $height = 10, $targetWidth, $visibleRatio = $null, [switch]$Filled) {
    $validValues = @($values | Where-Object { $_ -is [valueType] })
    if ($validValues.Count -eq 0) { return "  [No data available to render chart]" }

    $logicalWidth = $targetWidth * 2
    
    $colsToPlot = $logicalWidth
    if ($null -ne $visibleRatio) {
        $colsToPlot = [int][math]::Round($visibleRatio * $logicalWidth)
        if ($colsToPlot -lt 2) { $colsToPlot = 2 }
        if ($colsToPlot -gt $logicalWidth) { $colsToPlot = $logicalWidth }
    }

    $resizedValues = Invoke-InterpolateData -originalValues $validValues -targetSize $colsToPlot

    $min = ($resizedValues | Measure-Object -Minimum).Minimum
    $max = ($resizedValues | Measure-Object -Maximum).Maximum
    $range = $max - $min
    if ($range -eq 0) { $range = 1 }

    $gridHeight = [int]($height * 4)
    $grid = [byte[]]::new($gridHeight * $logicalWidth)

    for ($col = 0; $col -lt $colsToPlot; $col++) {
        $norm = ([double]$resizedValues[$col] - [double]$min) / [double]$range
        $targetY = [int][math]::Round($norm * ($gridHeight - 1))
        
        if ($Filled) {
            for ($row = 0; $row -le $targetY; $row++) {
                $grid[$row * $logicalWidth + $col] = 1
            }
        } else {
            $grid[$targetY * $logicalWidth + $col] = 1
        }
    }

    $brailleRows = [string[]]::new($height)
    for ($r = ($height - 1); $r -ge 0; $r--) {
        $sb = [System.Text.StringBuilder]::new($targetWidth)
        $idx0 = $r * 4
        $idx1 = $r * 4 + 1
        $idx2 = $r * 4 + 2
        $idx3 = $r * 4 + 3

        for ($c = 0; $c -lt $logicalWidth; $c += 2) {
            $c1 = $c
            $c2 = $c + 1

            $p1 = if ($grid[$idx3 * $logicalWidth + $c1] -eq 1) { 1 } else { 0 }
            $p2 = if ($grid[$idx2 * $logicalWidth + $c1] -eq 1) { 2 } else { 0 }
            $p3 = if ($grid[$idx1 * $logicalWidth + $c1] -eq 1) { 4 } else { 0 }
            $p4 = if ($grid[$idx3 * $logicalWidth + $c2] -eq 1) { 8 } else { 0 }
            $p5 = if ($grid[$idx2 * $logicalWidth + $c2] -eq 1) { 16 } else { 0 }
            $p6 = if ($grid[$idx1 * $logicalWidth + $c2] -eq 1) { 32 } else { 0 }
            $p7 = if ($grid[$idx0 * $logicalWidth + $c1] -eq 1) { 64 } else { 0 }
            $p8 = if ($grid[$idx0 * $logicalWidth + $c2] -eq 1) { 128 } else { 0 }

            $codeOffset = $p1 + $p2 + $p3 + $p4 + $p5 + $p6 + $p7 + $p8
            $null = $sb.Append([char](0x2800 + $codeOffset))
        }
        $brailleRows[($height - 1) - $r] = $sb.ToString()
    }
    return $brailleRows
}

# --- TIMELINE GENERATORS ---
function Get-TimelineRow ($timestamps, $targetWidth) {
    if ($null -eq $timestamps -or $timestamps.Count -lt 2) { return " " * $targetWidth }
    $origCount = $timestamps.Count
    $timelineChars = [char[]](" " * $targetWidth)
    $useMonth = $targetWidth -ge 60

    for ($origIdx = 0; $origIdx -lt $origCount; $origIdx++) {
        $epoch = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
        $date = $epoch.AddSeconds($timestamps[$origIdx]).ToLocalTime()

        if ($date.DayOfWeek -eq [DayOfWeek]::Monday) {
            $percent = $origIdx / ($origCount - 1)
            $consoleX = [int][math]::Round($percent * ($targetWidth - 1))

            $dayStr = if ($useMonth) { "$($date.Day)/$($date.Month)" } else { $date.Day.ToString() }
            $len = $dayStr.Length
            $startPos = [math]::Max(0, $consoleX - [int]($len / 2))
            
            if ($startPos + $len -le $targetWidth) {
                $canWrite = $true
                for ($k = 0; $k -lt $len; $k++) {
                    if ($timelineChars[$startPos + $k] -ne ' ') { $canWrite = $false; break }
                }
                if ($canWrite) {
                    for ($k = 0; $k -lt $len; $k++) { $timelineChars[$startPos + $k] = $dayStr[$k] }
                }
            }
        }
    }
    return -join $timelineChars
}

function Get-IntradayTimelineRow ($targetWidth) {
    $timelineChars = [char[]](" " * $targetWidth)
    $points = @(
        @{ Label = "Prev"; Hour = 9; Min = 0; IsPrev = $true }
        @{ Label = "09";   Hour = 9; Min = 0 }
        @{ Label = "10";   Hour = 10; Min = 0 }
        @{ Label = "11";   Hour = 11; Min = 0 }
        @{ Label = "12";   Hour = 12; Min = 0 }
        @{ Label = "13";   Hour = 13; Min = 0 }
        @{ Label = "14";   Hour = 14; Min = 0 }
        @{ Label = "15";   Hour = 15; Min = 0 }
        @{ Label = "16";   Hour = 16; Min = 0 }
        @{ Label = "17";   Hour = 17; Min = 0 }
        @{ Label = "17:30"; Hour = 17; Min = 30 }
    )
    
    foreach ($p in $points) {
        if ($p.IsPrev) {
            $percent = 0.0
        } else {
            $totalMinutesStart = 9 * 60
            $totalMinutesEnd = (17 * 60) + 30
            $currentMinutes = $p.Hour * 60 + $p.Min
            $percent = 0.12 + (0.88 * (($currentMinutes - $totalMinutesStart) / ($totalMinutesEnd - $totalMinutesStart)))
        }
        
        $lblText = $p.Label
        $consoleX = [int][math]::Round($percent * ($targetWidth - 1))
        
        $startPos = [math]::Max(0, $consoleX - [int]($lblText.Length / 2))
        if ($startPos + $lblText.Length -gt $targetWidth) { $startPos = $targetWidth - $lblText.Length }
        
        $canWrite = $true
        for ($k = 0; $k -lt $lblText.Length; $k++) {
            if ($timelineChars[$startPos + $k] -ne ' ') { $canWrite = $false; break }
        }
        if ($canWrite) {
            for ($k = 0; $k -lt $lblText.Length; $k++) { $timelineChars[$startPos + $k] = $lblText[$k] }
        }
    }
    return -join $timelineChars
}

function Get-YearlyTimelineRow ($timestamps, $targetWidth) {
    if ($null -eq $timestamps -or $timestamps.Count -lt 2) { return " " * $targetWidth }

    $timelineChars = [char[]](" " * $targetWidth)
    $epoch = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
    
    $monthlyPoints = New-Object System.Collections.Generic.List[PSCustomObject]
    $lastMonthKey = ""

    for ($i = 0; $i -lt $timestamps.Count; $i++) {
        $date = $epoch.AddSeconds($timestamps[$i]).ToLocalTime()
        $monthKey = $date.ToString("MM/yy")
        if ($monthKey -ne $lastMonthKey) {
            $monthlyPoints.Add([PSCustomObject]@{ Index = $i; Label = $monthKey })
            $lastMonthKey = $monthKey
        }
    }

    if ($monthlyPoints.Count -eq 0) { return " " * $targetWidth }

    $requiredSpace = 8
    $maxLabelsAllowed = [math]::Max(1, [int][math]::Floor($targetWidth / $requiredSpace))
    $step = if ($monthlyPoints.Count -gt $maxLabelsAllowed) { [int][math]::Ceiling($monthlyPoints.Count / $maxLabelsAllowed) } else { 1 }

    for ($idx = 0; $idx -lt $monthlyPoints.Count; $idx += $step) {
        $point = $monthlyPoints[$idx]
        $percent = $point.Index / ($timestamps.Count - 1)
        $consoleX = [int][math]::Round($percent * ($targetWidth - 1))
        
        $lbl = $point.Label
        $startPos = [math]::Max(0, $consoleX - [int]($lbl.Length / 2))
        if ($startPos + $lbl.Length -gt $targetWidth) { $startPos = $targetWidth - $lbl.Length }

        $canWrite = $true
        for ($k = 0; $k -lt $lbl.Length; $k++) {
            if ($timelineChars[$startPos + $k] -ne ' ') { $canWrite = $false; break }
        }
        if ($canWrite) {
            for ($k = 0; $k -lt $lbl.Length; $k++) { $timelineChars[$startPos + $k] = $lbl[$k] }
        }
    }
    return -join $timelineChars
}

# --- YAHOO CRUMB RETRIEVAL (with GDPR consent fallback, no browser login required) ---
# Helper: extract the value="" of an <input> tag identified by its name="" attribute,
# regardless of the order in which the HTML attributes appear.
function Get-HtmlInputValue {
    param([string]$Html, [string]$Name)
    $tagPattern = "<input\b[^>]*name\s*=\s*[""']$Name[""'][^>]*>"
    $tagMatch = [regex]::Match($Html, $tagPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $tagMatch.Success) { return $null }
    $valMatch = [regex]::Match($tagMatch.Value, "value\s*=\s*[""']([^""']*)[""']", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($valMatch.Success) { return $valMatch.Groups[1].Value }
    return $null
}

# Strategy 1 ("basic"): works out-of-the-box for most US-based IPs.
function Get-YahooCrumbBasic {
    param($WebSession)
    try {
        $null = Invoke-WebRequest -Uri "https://fc.yahoo.com" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10 -ErrorAction SilentlyContinue
    } catch {}
    try {
        $resp = Invoke-WebRequest -Uri "https://query2.finance.yahoo.com/v1/test/getcrumb" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10
        $c = $resp.Content.Trim()
        if ($c -and $c -notmatch "Too Many Requests") { return $c }
    } catch {}
    return $null
}

# Strategy 2 ("consent/GDPR"): required for most EU-based IPs (including Italy) the first
# time a fresh session hits Yahoo, since the "basic" flow gets silently redirected to the
# cookie-consent page instead of receiving usable cookies. This replicates - purely over
# HTTP, with no browser and no Yahoo account needed - what happens when a person manually
# accepts Yahoo's cookie banner.
function Get-YahooCrumbConsent {
    param($WebSession)
    try {
        $consentResp = Invoke-WebRequest -Uri "https://guce.yahoo.com/consent" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10
        $html = $consentResp.Content

        $csrfToken = Get-HtmlInputValue -Html $html -Name "csrfToken"
        $sessionId = Get-HtmlInputValue -Html $html -Name "sessionId"
        if (-not $csrfToken -or -not $sessionId) { return $null }

        $bodyString = "agree=agree&agree=agree" +
            "&consentUUID=default" +
            "&sessionId=$([uri]::EscapeDataString($sessionId))" +
            "&csrfToken=$([uri]::EscapeDataString($csrfToken))" +
            "&originalDoneUrl=$([uri]::EscapeDataString('https://finance.yahoo.com/'))" +
            "&namespace=yahoo"

        $null = Invoke-WebRequest -Uri "https://consent.yahoo.com/v2/collectConsent?sessionId=$sessionId" `
            -WebSession $WebSession -Method Post -Body $bodyString -ContentType "application/x-www-form-urlencoded" `
            -UseBasicParsing -TimeoutSec 10
        $null = Invoke-WebRequest -Uri "https://guce.yahoo.com/copyConsent?sessionId=$sessionId" `
            -WebSession $WebSession -UseBasicParsing -TimeoutSec 10

        $resp = Invoke-WebRequest -Uri "https://query2.finance.yahoo.com/v1/test/getcrumb" -WebSession $WebSession -UseBasicParsing -TimeoutSec 10
        $c = $resp.Content.Trim()
        if ($c -and $c -notmatch "Too Many Requests") { return $c }
    } catch {}
    return $null
}

function Get-YahooCrumb {
    param($WebSession)
    $c = Get-YahooCrumbBasic -WebSession $WebSession
    if (-not $c) { $c = Get-YahooCrumbConsent -WebSession $WebSession }
    return $c
}

# --- FETCH DATA WITH TIMEOUT AND TRY/CATCH ---
try {
    Show-Progress 3 $totalSteps "Connecting to Yahoo Finance and getting session token (crumb)"

    $crumb = Get-YahooCrumb -WebSession $webSession
    if (-not $crumb) {
        Write-Error "ERROR: Unable to obtain a Yahoo Finance session token (crumb), even after attempting the GDPR consent fallback."
        Write-Warning "Yahoo Finance may be temporarily blocking automated requests from this IP. Try again in a few minutes."
        exit
    }

    Show-Progress 4 $totalSteps "Downloading real-time data"

    $apiUrl = "https://query2.finance.yahoo.com/v7/finance/quote?symbols=$tickers&crumb=$crumb"
    $response = Invoke-RestMethod -Uri $apiUrl -WebSession $webSession -Method Get -TimeoutSec 10
    $quotes = $response.quoteResponse.result

    Show-Progress 5 $totalSteps "Downloading historical data"

    # 3. PARALLEL DOWNLOAD OF HISTORICAL DATA VIA RUNSPACE POOL
    # (ForEach-Object -Parallel is a PowerShell 7+ only feature; a runspace pool
    #  works identically on Windows PowerShell 5.1 and PowerShell 7+, out of the box.)
    $today = [DateTime]::Today
    $startTime = (New-Object DateTimeOffset ($today.AddHours(9))).ToUnixTimeSeconds()
    $endTime = (New-Object DateTimeOffset ($today.AddHours(17).AddMinutes(30))).ToUnixTimeSeconds()
    $intradaySteps = 271
    $stepInterval = ($endTime - $startTime) / ($intradaySteps - 1)

    $assetWorkerScript = {
        param($asset, $webSession, $customStartUnix, $nowUnix, $today, $dateParsed, $startTime, $stepInterval, $intradaySteps)

        $ticker = $asset.Ticker
        $avgCost = $asset.AvgCost

        $res = @{
            Ticker = $ticker
            HistoricalCloses = $null
            GlobalTimestamps = $null
            CustomCloses = $null
            CustomTimestamps = $null
            AlignedIntraday = $null
        }

        function Invoke-SafeRest ($uri) {
            try { return Invoke-RestMethod -Uri $uri -WebSession $webSession -Method Get -TimeoutSec 10 } catch { return $null }
        }

        # --- CUSTOM HISTORY ---
        $diffDays = ($today - $dateParsed).TotalDays
        $interval = if ($diffDays -gt 180) { "1wk" } else { "1d" }
        $customChartUrl = "https://query2.finance.yahoo.com/v8/finance/chart/${ticker}?period1=$customStartUnix&period2=$nowUnix&interval=$interval"
        $customChartRes = Invoke-SafeRest $customChartUrl
        if ($customChartRes -and $customChartRes.chart.result) {
            $res.CustomTimestamps = $customChartRes.chart.result[0].timestamp
            $cCloses = $customChartRes.chart.result[0].indicators.quote[0].close
            $cleanedC = @()
            $lastValidC = $avgCost
            foreach ($p in $cCloses) {
                if ($null -ne $p -and $p -is [valueType]) { $lastValidC = [double]$p }
                $cleanedC += $lastValidC
            }
            $res.CustomCloses = $cleanedC
        }

        # --- INTRADAY ---
        $intraUrl = "https://query2.finance.yahoo.com/v8/finance/chart/${ticker}?range=1d&interval=2m"
        $intraRes = Invoke-SafeRest $intraUrl
        
        $intraPoints = @{}
        if ($intraRes -and $intraRes.chart.result) {
            $intraTimes = $intraRes.chart.result[0].timestamp
            $intraCloses = $intraRes.chart.result[0].indicators.quote[0].close
            if ($intraTimes) {
                for ($k = 0; $k -lt $intraTimes.Count; $k++) {
                    $t = $intraTimes[$k]
                    $p = $intraCloses[$k]
                    if ($null -ne $p -and $p -is [valueType]) { $intraPoints[$t] = [double]$p }
                }
            }
        }

        $alignedIntraday = [double[]]::new($intradaySteps)
        $prevClosePrice = $avgCost
        if ($intraRes -and $intraRes.chart.result) {
            $prevClosePrice = [double]$intraRes.chart.result[0].meta.chartPreviousClose
        }
        $alignedIntraday[0] = $prevClosePrice
        $lastKnownIntraPrice = $prevClosePrice

        for ($step = 1; $step -lt $intradaySteps; $step++) {
            $targetTimestamp = $startTime + (($step - 1) * $stepInterval)
            $bestKey = $null
            $minDiff = 600
            foreach ($key in $intraPoints.Keys) {
                $diff = [math]::Abs($key - $targetTimestamp)
                if ($diff -lt $minDiff) {
                    $minDiff = $diff
                    $bestKey = $key
                }
            }
            if ($null -ne $bestKey) { $lastKnownIntraPrice = $intraPoints[$bestKey] }
            $alignedIntraday[$step] = $lastKnownIntraPrice
        }
        $res.AlignedIntraday = $alignedIntraday

        return $res
    }

    # Run the worker script block once per asset, in parallel, via a runspace pool.
    $throttleLimit = [math]::Max(1, [math]::Min(8, $assets.Count))
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $throttleLimit)
    $runspacePool.Open()

    $pendingJobs = @()
    foreach ($asset in $assets) {
        $psInstance = [powershell]::Create()
        $psInstance.RunspacePool = $runspacePool
        [void]$psInstance.AddScript($assetWorkerScript).AddParameters(@{
            asset            = $asset
            webSession       = $webSession
            customStartUnix  = $customStartUnix
            nowUnix          = $nowUnix
            today            = $today
            dateParsed       = $dateParsed
            startTime        = $startTime
            stepInterval     = $stepInterval
            intradaySteps    = $intradaySteps
        })
        $pendingJobs += [PSCustomObject]@{
            Pipe   = $psInstance
            Handle = $psInstance.BeginInvoke()
        }
    }

    $parallelResults = foreach ($job in $pendingJobs) {
        try {
            $job.Pipe.EndInvoke($job.Handle)
        } catch {
            Write-Warning "WARNING: A background download task failed: $_"
        } finally {
            $job.Pipe.Dispose()
        }
    }

    $runspacePool.Close()
    $runspacePool.Dispose()

    $customData     = @{}
    $intradayData   = @{}
    $globalTimestamps = $null
    $yearlyTimestamps = $null
    $customTimestamps = $null

    foreach ($r in $parallelResults) {
        $t = $r.Ticker
        if ($r.CustomCloses)     { $customData[$t]     = $r.CustomCloses }
        if ($r.AlignedIntraday)  { $intradayData[$t]   = $r.AlignedIntraday }

        if (-not $globalTimestamps -and $r.GlobalTimestamps) { $globalTimestamps = $r.GlobalTimestamps }
        if (-not $yearlyTimestamps -and $r.YearlyTimestamps) { $yearlyTimestamps = $r.YearlyTimestamps }
        if (-not $customTimestamps -and $r.CustomTimestamps) { $customTimestamps = $r.CustomTimestamps }
    }

    $portfolioPrevCloseValue = 0
    foreach ($asset in $assets) {
        $quote = $quotes | Where-Object { $_.symbol -eq $asset.Ticker }
        if ($quote) {
            $prevPrice = [double]$quote.regularMarketPreviousClose
            $portfolioPrevCloseValue += ($prevPrice * [double]$asset.Qty)
        } else {
            $portfolioPrevCloseValue += ([double]$asset.AvgCost * [double]$asset.Qty)
        }
    }

} catch {
    Write-Error "Error retrieving data from Yahoo Finance: $_"
    exit
}

Show-Progress 6 $totalSteps "Rendering"

# --- DYNAMIC CONSOLE DIMENSIONS CALCULATION ---
$consoleHeight = 40
if ($Host.UI.RawUI.WindowSize.Height) { $consoleHeight = $Host.UI.RawUI.WindowSize.Height }
$targetChartWidth = $consoleWidth - 2
$widthThreshold2 = 140

$assetRowMultiplier = if ($consoleWidth -lt 130) { 4 } else { 3 } 
$metricsRows = if ($csvMetricsAvailable) { if ($consoleWidth -ge $widthThreshold2) { 3 } else { 5 } } else { 0 }
$fixedRows = 10 + $metricsRows + ($assets.Count * $assetRowMultiplier)
$availableHeightForCharts = $consoleHeight - $fixedRows
$calculatedChartHeight = [math]::Max(8, [int][math]::Floor($availableHeightForCharts / 2))

# Custom Portfolio Calculations
$portfolioCustom = @()
$customCount = $customData.Values | ForEach-Object { $_.Count } | Measure-Object -Min
$minCustom = $customCount.Minimum

if ($minCustom -gt 0) {
    for ($idx = 0; $idx -lt $minCustom; $idx++) {
        $idxValue = 0
        foreach ($asset in $assets) {
            $priceOnIdx = if ($customData.ContainsKey($asset.Ticker)) { $customData[$asset.Ticker][$idx] } else { $asset.AvgCost }
            $idxValue += ($priceOnIdx * [double]$asset.Qty)
        }
        $portfolioCustom += $idxValue
    }
} else {
    $portfolioCustom = @(0)
}

# Intraday Portfolio Calculations
$portfolioIntraday = @()
for ($step = 0; $step -lt $intradaySteps; $step++) {
    $stepValue = 0
    foreach ($asset in $assets) {
        $priceAtStep = if ($intradayData.ContainsKey($asset.Ticker)) { $intradayData[$asset.Ticker][$step] } else { $asset.AvgCost }
        $stepValue += ($priceAtStep * [double]$asset.Qty)
    }
    $portfolioIntraday += $stepValue
}

$localNowUnix = (New-Object DateTimeOffset ([DateTime]::Now)).ToUnixTimeSeconds()
$currentStepLimit = [int][math]::Floor(($localNowUnix - $startTime) / $stepInterval) + 1
$currentStepLimit = [math]::Max(1, [math]::Min($currentStepLimit, $intradaySteps - 1))

$portfolioIntradayFiltered = $portfolioIntraday[0..$currentStepLimit]
$intradayRatio = $currentStepLimit / ($intradaySteps - 1)

# Chart Generation
$intradayChartRows = Get-BrailleSparkline -values $portfolioIntradayFiltered -height $calculatedChartHeight -targetWidth $targetChartWidth -visibleRatio $intradayRatio
$intradayTimelineRow = Get-IntradayTimelineRow -targetWidth $targetChartWidth

$customChartRows = Get-BrailleSparkline -values $portfolioCustom -height $calculatedChartHeight -targetWidth $targetChartWidth -Filled
$customTimelineRow = Get-YearlyTimelineRow -timestamps $customTimestamps -targetWidth $targetChartWidth

# Real-time Metrics
$portfolio = @()
$totalCost = 0
$totalValue = 0
$totalDayChange = 0

foreach ($asset in $assets) {
    $quote = $quotes | Where-Object { $_.symbol -eq $asset.Ticker }
    if (-not $quote) { continue }

    $currentPrice = [double]$quote.regularMarketPrice
    $prevClose = [double]$quote.regularMarketPreviousClose
    $displayName = $quote.longName

    $cost = [double]$asset.AvgCost * [double]$asset.Qty
    $value = $currentPrice * [double]$asset.Qty
    $gain = $value - $cost
    $gainPct = if ($cost -ne 0) { ($gain / $cost) * 100 } else { 0 }

    $dayChangePerShare = $currentPrice - $prevClose
    $dayChangeTotal = $dayChangePerShare * [double]$asset.Qty
    $dayChangePct = if ($prevClose -ne 0) { ($dayChangePerShare / $prevClose) * 100 } else { 0 }

    $totalCost += $cost
    $totalValue += $value
    $totalDayChange += $dayChangeTotal

    $portfolio += [PSCustomObject]@{
        Ticker         = $asset.Ticker
        Name           = $displayName
        CurrentPrice   = $currentPrice
        AvgCost        = $asset.AvgCost
        Qty            = $asset.Qty
        Value          = $value
        Gain           = $gain
        GainPct        = $gainPct
        DayChangeTotal = $dayChangeTotal
        DayChangePct   = $dayChangePct
    }
}

$totalChange = $totalValue - $totalCost
$totalChangePct = if ($totalCost -ne 0) { ($totalChange / $totalCost) * 100 } else { 0 }
$totalDayChangePct = if (($totalValue - $totalDayChange) -ne 0) { ($totalDayChange / ($totalValue - $totalDayChange)) * 100 } else { 0 }

# ============================================================
# CORRECTED MWRR CALCULATION
# Add the FINAL portfolio value as a positive cash flow
# (as if liquidating the entire portfolio at today's value)
# ============================================================
$mwrr = $null
if ($csvMetricsAvailable) {
    $cashflows.Add([PSCustomObject]@{ 
        Date = [DateTime]::Today
        Amount = $totalValue   # POSITIVE: money coming IN to investor
    })
    $mwrr = Get-XirrRobust $cashflows
}
# ============================================================

# --- CONSOLE PRINTING (BUFFERED OUTPUT TO PREVENT FLICKER) ---
$ESC = [char]27
$Green = "$ESC[92m"
$Red = "$ESC[91m"
$Gray = "$ESC[90m"
$White = "$ESC[97m"
$Orange = "$ESC[33m"
$Reset = "$ESC[0m"

function Get-TrendColor ($val) { 
	if ($val -gt 0) { return $Green } elseif ($val -lt 0) { return $Red } else { return $Gray } 
}

function Get-TrendArrow ($val) { 
	if ($val -gt 0) { return [char]0x2191 } elseif ($val -lt 0) { return [char]0x2193 } else { return [char]0x2192 } 
}

$outputBuffer = [System.Text.StringBuilder]::new()

$null = $outputBuffer.Append("$ESC[2J$ESC[H")

# A. INTRADAY Chart Block
$null = $outputBuffer.AppendLine("${White}Portfolio performance - Today${Reset}")

$intraFirstVal = $portfolioPrevCloseValue
$intraLastVal = if ($portfolioIntradayFiltered.Count -gt 0) { $portfolioIntradayFiltered[-1] } else { 0 }
$intraChartColor = Get-TrendColor ($intraLastVal - $intraFirstVal)

# Percentuale e delta: allineati esattamente al calcolo di "Day change" più sotto,
# basato sulle quotazioni realtime (non sulla serie storica intraday del grafico,
# che ha una granularità di 2 minuti e può risultare leggermente disallineata).
$intraChangePct = $totalDayChangePct

# Formatta la percentuale con segno e colore
$pctSign = if ($intraChangePct -gt 0) { "+" } elseif ($intraChangePct -lt 0) { "" } else { "" }
$pctColor = Get-TrendColor $intraChangePct
$pctFormatted = "$pctColor$pctSign$($intraChangePct.ToString('F2').Replace('.', ','))%$Reset"
$cleanPctStr = "$pctSign$($intraChangePct.ToString('F2').Replace('.', ','))%"

# Delta assoluto = totalDayChange (stesso valore di "Day change"), con segno esplicito e colore trend
$intraDelta = $totalDayChange
$deltaColor = Get-TrendColor $intraDelta
$deltaSign = if ($intraDelta -gt 0) { "+" } else { "" }   # il "-" per i negativi arriva già da ToString
$deltaValueStr = $deltaSign + $intraDelta.ToString('F2').Replace('.', ',')

# Testo pulito (senza codici di colore) per il calcolo preciso della lunghezza visibile
$cleanNowStr = "Now: " + $deltaValueStr + " (" + $cleanPctStr + ")"

# Testo effettivo con i colori ANSI da stampare
$labelPrev = "Prev: " + $intraFirstVal.ToString('F2')
$labelNow = "Now: " + "${deltaColor}${deltaValueStr}${Reset}" + " (" + $pctFormatted + ")"

# Allineamento orizzontale corrispondente all'ora corrente sul grafico (stessa proporzione
# usata per disegnare lo sparkline, $intradayRatio), centrando l'etichetta su quel punto.
$consoleXNow = [int][math]::Round($intradayRatio * ($targetChartWidth - 1))
$startNow = $consoleXNow - [int]($cleanNowStr.Length / 2)

# Non deve mai superare il margine destro della console
$maxStartNow = $targetChartWidth - $cleanNowStr.Length
if ($startNow -gt $maxStartNow) { $startNow = $maxStartNow }

# Non deve mai sovrapporsi all'etichetta "Prev:"
$minStartNow = $labelPrev.Length + 2
if ($startNow -lt $minStartNow) { $startNow = $minStartNow }

$spacesCount = $startNow - $labelPrev.Length
if ($spacesCount -lt 1) { $spacesCount = 1 }

# Costruzione della riga senza troncamenti o stringhe fisse sproporzionate
$headerValuesLine = $labelPrev + (" " * $spacesCount) + $labelNow
$null = $outputBuffer.AppendLine(" ${Reset}${headerValuesLine}${Reset}")

foreach ($row in $intradayChartRows) {
    $null = $outputBuffer.AppendLine(" ${intraChartColor}${row}${Reset}")
}
$null = $outputBuffer.AppendLine(" ${Gray}${intradayTimelineRow}${Reset}")
$null = $outputBuffer.AppendLine("-" * $consoleWidth)

# D. Custom Historical Chart Block
$null = $outputBuffer.AppendLine("${White}Portfolio performance - Since ${startDateConfig}${Reset}")
$cFirstVal = if ($portfolioCustom.Count -gt 0) { $portfolioCustom[0] } else { 0 }
$cLastVal = if ($portfolioCustom.Count -gt 0) { $portfolioCustom[-1] } else { 0 }
$cChartColor = Get-TrendColor ($cLastVal - $cFirstVal)

foreach ($row in $customChartRows) {
    $null = $outputBuffer.AppendLine(" ${cChartColor}${row}${Reset}")
}
$null = $outputBuffer.AppendLine(" ${Gray}${customTimelineRow}${Reset}")
$null = $outputBuffer.AppendLine("=" * $consoleWidth)

# E. Summary Row
$dcColor = Get-TrendColor $totalDayChange
$tcColor = Get-TrendColor $totalChange

if ($csvMetricsAvailable -and $null -ne $mwrr) {
    $mwrrColor = Get-TrendColor $mwrr
    $mwrrStr = "MWRR: ${mwrrColor}$(Get-TrendArrow $mwrr) $(($mwrr * 100).ToString('F2'))% annual${Reset}"
} else {
    $mwrrStr = ""
}

$null = $outputBuffer.AppendLine(" P/L: ${tcColor}$(Get-TrendArrow $totalChange) $($totalChange.ToString('F2')) ($($totalChangePct.ToString('F2'))%)$Reset   $mwrrStr   Day change: ${dcColor}$(Get-TrendArrow $totalDayChange) $($totalDayChange.ToString('F2')) ($($totalDayChangePct.ToString('F2'))%)$Reset   Cost: ${White}$($totalCost.ToString('F2'))$Reset   Value: ${White}$($totalValue.ToString('F2'))$Reset")

# E-bis. Portfolio Metrics (Impaginazione reattiva alla larghezza console)
if ($csvMetricsAvailable) {

    # Cost ratio = impatto dei costi sul guadagno totale (realizzato + non realizzato).
    # Guadagno totale = valore attuale + prelevato - versato.
    # Se il guadagno non è positivo, il rapporto non ha un senso economico pulito -> N/A.
    $totalGainAllTime = $totalValue + $totalePrelevato - $totaleVersato
    if ($totaleVersato -ne 0) {
        $costRatio = if ($totalGainAllTime -gt 0) { ($totaleCostiSostenuti / $totalGainAllTime) * 100 } else { $null }
    } elseif ($totalCost -ne 0) {
        $costRatio = ($totaleCostiSostenuti / $totalCost) * 100
    } else {
        $costRatio = $null
    }
    $costRatioStr = if ($null -ne $costRatio) { "$($costRatio.ToString('F2'))%" } else { "N/A" }
    $dataPrimoStr = if ($dataPrimoInvestimento) { $dataPrimoInvestimento.ToString('dd/MM/yyyy') } else { "N/A" }

    $m2 = "${Gray}Expenses:${Reset} ${Orange}$($totaleCostiSostenuti.ToString('F2'))${Reset} (${Orange}${costRatioStr} ${Gray}of gain${Reset})"
    $m3 = "${Gray}Commissions${Reset} ${Orange}$($totaleCommissioni.ToString('F2'))${Reset} + ${Gray}Stamp duty${Reset} ${Orange}$($totaleBollo.ToString('F2'))${Reset} + ${Gray}Capital gain taxes${Reset} ${Orange}$($totaleRitenute.ToString('F2'))${Reset}"
    $m4 = "${Gray}Trades:${Reset} ${White}${numeroAcquisti} buy & ${numeroVendite} sell${Reset}"
    $m5 = "${Gray}Positions:${Reset} ${White}${numeroPosizioniAperte} open & ${numeroPosizioniChiuse} closed${Reset}"
    $m6 = "${Gray}First invest:${Reset} ${White}${dataPrimoStr}${Reset} (${White}${giorniInvestimento}d ago${Reset})"
    $m7 = "${Gray}Deposited:${Reset} ${White}$($totaleVersato.ToString('F2'))${Reset}  ${Gray}Withdrawn:${Reset} ${White}$($totalePrelevato.ToString('F2'))${Reset}"

    if ($consoleWidth -ge $minWidth) {
        $null = $outputBuffer.AppendLine(" $m2  =  $m3")
        $null = $outputBuffer.AppendLine(" $m4  $m5  $m6  $m7")
    } else {
        $null = $outputBuffer.AppendLine(" $m2 =")
        $null = $outputBuffer.AppendLine("  = $m3")
        $null = $outputBuffer.AppendLine(" $m4  $m5")
        $null = $outputBuffer.AppendLine(" $m6  $m7")
    }
    $null = $outputBuffer.AppendLine("-" * $consoleWidth)
}

# F. Individual Assets

function Get-AssetCategory ($ticker, $name) {
    $t = $ticker.ToUpper()
    $n = $name.ToUpper()

    # Commodities / Gold: Yellow Background (43), Black Text (30)
    if ($t -match "GLD" -or $n -match "GOLD" -or $n -match "COMMODITY" -or $n -match "PHYSICAL") {
        return @{ Label = " COMMODITY "; Color = "$ESC[43;30m" }
    }
    # Money Market / Cash: Dark Gray Background (100), Bright White Text (97)
    if ($t -match "XEON" -or $t -match "LEON" -or $n -match "OVERNIGHT" -or $n -match "LIQUIDITY" -or $n -match "MONEY MARKET") {
        return @{ Label = " MONEY MKT "; Color = "$ESC[100;97m" }
    }
    # Fixed Income / Bonds: Blue Background (44), Bright White Text (97)
    if ($t -match "EM35" -or $n -match "BOND" -or $n -match "TREASURY" -or $n -match "GOVT" -or $n -match "CORP" -or $n -match "OBBL") {
        return @{ Label = "   BOND    "; Color = "$ESC[44;97m" }
    }
    
    # Equities (Default): Red Background (41), Bright White Text (97)
    return @{ Label = "  EQUITY   "; Color = "$ESC[41;97m" }
}

foreach ($item in $portfolio) {
    $weight = ($item.Value / $totalValue) * 100
    $gainColor = Get-TrendColor $item.Gain
    $dcColor = Get-TrendColor $item.DayChangeTotal

    # Label generation and formatting
    $cat = Get-AssetCategory $item.Ticker $item.Name
    $labelFormatted = "$($cat.Color)$($cat.Label)$Reset"

    # Print asset header with prepended label
    $null = $outputBuffer.AppendLine("$labelFormatted ${White}$($item.Ticker) ${Gray}$($item.Name)${Reset}")

    $posRow = " ${Gray}Qty:${Reset} {0,4:F0}   ${Gray}Avg cost:${Reset} {1,6:F2}   ${Gray}Price:${Reset} {2,6:F2}   ${Gray}Value: ${White}{3,8:F2} ({4,5:F2}%)${Reset}" -f `
        $item.Qty, $item.AvgCost, $item.CurrentPrice, $item.Value, $weight

    $rawGainText = "{0} {1,8:F2} ({2,5:F2}%)" -f (Get-TrendArrow $item.Gain), $item.Gain, $item.GainPct
    $rawDayText  = "{0} {1,7:F2} ({2,5:F2}%)" -f (Get-TrendArrow $item.DayChangeTotal), $item.DayChangeTotal, $item.DayChangePct

    $gainFormatted = "{0,-18}" -f $rawGainText
    $dayFormatted  = "{0,-18}" -f $rawDayText

    $coloredGain = "${gainColor}${gainFormatted}${Reset}"
    $coloredDay  = "${dcColor}${dayFormatted}${Reset}"

    $plString = "${Gray}P/L:${Reset} ${coloredGain}   ${Gray}Day:${Reset} $coloredDay"

    if ($consoleWidth -ge $widthThreshold2) {
        $null = $outputBuffer.AppendLine("$posRow   $plString")
    } else {
        $null = $outputBuffer.AppendLine("$posRow")
        $null = $outputBuffer.AppendLine("  $plString")
    }
    $null = $outputBuffer.AppendLine("")
}

Write-Host $outputBuffer.ToString() -NoNewline