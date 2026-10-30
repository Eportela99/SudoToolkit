<#
    TechRoute - Tecnica Systems
    Reads the Helpdesk "Tickets Report" export (.xlsx) and builds a Google Maps route
    starting at the office. Pure PowerShell 5.1 + WinForms, no Excel needed.

    Route ordering: stops always go from the farthest (from the office) to the closest.
      By Open Date : takes the N oldest open tickets, then orders them by distance.
      By Distance  : takes the N farthest open tickets, ordered by distance.
      Custom       : the tech picks the tickets, ordered by distance.
    Addresses are geocoded with the ArcGIS World geocoder (free, no key, handles typos).
#>
param([switch]$SelfTest)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#region Config
$script:App = @{
    Name    = 'TechRoute'
    Version = '1.2.0'
    Root    = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent ([Environment]::GetCommandLineArgs()[0]) }
}
$script:App.Assets = Join-Path $script:App.Root 'Assets'
$script:Office = [pscustomobject]@{
    Address = '7959 NW 21st St, Doral, FL 33122'
    Lat     = 25.7942316   # fallback if the geocoder is unreachable
    Lon     = -80.3258367
}
$script:MaxStopsPerLink = 10   # Google Maps: destination + 9 waypoints per link
$script:Tickets  = @()
$script:TechName = ''
$script:FilePath = ''
$script:GeoCache = @{}         # session only (address -> result)

$script:Theme = @{
    Bg      = [Drawing.Color]::FromArgb(24, 26, 31)
    Panel   = [Drawing.Color]::FromArgb(34, 37, 44)
    Button  = [Drawing.Color]::FromArgb(48, 52, 62)
    Hover   = [Drawing.Color]::FromArgb(62, 67, 80)
    Text    = [Drawing.Color]::FromArgb(232, 234, 238)
    Muted   = [Drawing.Color]::FromArgb(150, 156, 168)
    Brand   = [Drawing.Color]::FromArgb(132, 204, 22)
    Warn    = [Drawing.Color]::FromArgb(250, 190, 60)
    Font    = New-Object Drawing.Font('Segoe UI', 10)
    Bold    = New-Object Drawing.Font('Segoe UI', 10, [Drawing.FontStyle]::Bold)
    Big     = New-Object Drawing.Font('Segoe UI Semibold', 13)
    Title   = New-Object Drawing.Font('Segoe UI Black', 20, [Drawing.FontStyle]::Italic)
}
#endregion

#region Excel reader (xlsx = zip of XML)
function Read_ZipText($zip, [string]$name) {
    $entry = $zip.GetEntry($name)
    if (-not $entry) { return $null }
    $reader = New-Object IO.StreamReader($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Close() }
}

function Get_ColumnIndex([string]$cellRef) {
    $letters = ($cellRef -replace '\d', '').ToUpper()
    $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n - 1
}

function Convert_ToDate($value) {
    if ($value -is [double]) { return [DateTime]::FromOADate($value) }
    $s = "$value".Trim()
    if ($s -match '^0*(\d{1,2})/0*(\d{1,2})/(\d{4})') {
        try { return New-Object DateTime ([int]$Matches[3]), ([int]$Matches[1]), ([int]$Matches[2]) } catch { }
    }
    $d = [DateTime]::MinValue
    if ([DateTime]::TryParse($s, [ref]$d)) { return $d }
    return [DateTime]::MaxValue
}

function Read_TicketsXlsx([string]$path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $shared = @()
        $ssText = Read_ZipText $zip 'xl/sharedStrings.xml'
        if ($ssText) { $shared = @(([xml]$ssText).sst.si | ForEach-Object { $_.InnerText }) }
        $sheetText = Read_ZipText $zip 'xl/worksheets/sheet1.xml'
        if (-not $sheetText) { throw 'The file has no worksheet (xl/worksheets/sheet1.xml).' }
        [xml]$sheet = $sheetText
    } finally { $zip.Dispose() }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $sheet.worksheet.sheetData.row) {
        $cells = @{}
        $i = 0
        foreach ($c in @($row.c)) {
            if ($c.r) { $i = Get_ColumnIndex $c.r }
            $val = switch ($c.t) {
                'inlineStr' { $c.is.InnerText }
                's'         { $shared[[int]$c.v] }
                'str'       { $c.v }
                default     { if ($null -ne $c.v -and "$($c.v)" -ne '') { [double]$c.v } else { '' } }
            }
            $cells[$i] = $val
            $i++
        }
        $rows.Add($cells)
    }
    if ($rows.Count -lt 1) { throw 'The file is empty.' }

    # map headers by name so column order does not matter
    $head = @{}
    foreach ($k in $rows[0].Keys) { $head[("$($rows[0][$k])").Trim().ToLower()] = $k }
    $need = 'ticket no.', 'open date', 'ticket subject', 'description', 'customer name', 'customer address'
    $missing = @($need | Where-Object { -not $head.ContainsKey($_) })
    if ($missing) { throw "This does not look like a Helpdesk Tickets Report. Missing columns: $($missing -join ', ')" }
    $get = { param($cells, $name) if ($head.ContainsKey($name)) { "$($cells[$head[$name]])".Trim() } else { '' } }

    $list = New-Object System.Collections.Generic.List[object]
    for ($r = 1; $r -lt $rows.Count; $r++) {
        $cells = $rows[$r]
        $num = & $get $cells 'ticket no.'
        if (-not $num) { continue }
        $openRaw = if ($head.ContainsKey('open date')) { $cells[$head['open date']] } else { '' }
        $item = [pscustomobject]@{
            Number      = $num -replace '\.0+$', ''
            Status      = & $get $cells 'ticket status'
            OpenDate    = Convert_ToDate $openRaw
            Subject     = & $get $cells 'ticket subject'
            Description = (& $get $cells 'description') -replace '\s*\r?\n\s*', ' '
            Customer    = & $get $cells 'customer name'
            Phone       = & $get $cells 'customer phone'
            Address     = & $get $cells 'customer address'
            City        = & $get $cells 'city'
            State       = & $get $cells 'state'
            Tech        = & $get $cells 'technician'
            Lat = $null; Lon = $null; Miles = $null; GeoAddress = ''; GeoNote = ''
        }
        $list.Add($item)
    }
    return , $list.ToArray()
}

function Test_IsOpen($ticket) { return $ticket.Status -notmatch 'closed|cancel|complete|resolved' }
function Get_OpenTickets { return @($script:Tickets | Where-Object { Test_IsOpen $_ }) }

function Find_TicketsFile {
    $candidates = @()
    foreach ($dir in $script:App.Assets, $script:App.Root, (Join-Path $env:USERPROFILE 'Downloads')) {
        if (Test-Path $dir) { $candidates += @(Get-ChildItem $dir -Filter 'Tickets Report*.xlsx' -File -ErrorAction SilentlyContinue) }
    }
    $newest = $candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newest) { return $newest.FullName }
    return $null
}
#endregion

#region Geocoding + distance
function Get_FullAddress($t) {
    $a = $t.Address -replace '\bN\.W\b', 'NW' -replace '\bS\.W\b', 'SW' -replace '\bN\.E\b', 'NE' -replace '\bS\.E\b', 'SE'
    if ($t.City -and $a -notmatch [regex]::Escape($t.City)) { $a += ", $($t.City)" }
    $state = if ($t.State -match '^(FL|FLORIDA)$' -or -not $t.State) { 'FL' } else { $t.State }
    if ($a -notmatch "\b($state|FLORIDA)\b") { $a += ", $state" }
    return $a
}

function Invoke_Geocode([string]$address) {
    $key = $address.ToUpper()
    if ($script:GeoCache.ContainsKey($key)) { return $script:GeoCache[$key] }
    $url = 'https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates' +
           '?f=json&maxLocations=1&countryCode=USA&SingleLine=' + [uri]::EscapeDataString($address)
    $result = $null
    try {
        $c = (Invoke-RestMethod -Uri $url -TimeoutSec 20).candidates | Select-Object -First 1
        if ($c -and $c.score -ge 70) {
            $result = [pscustomobject]@{ Lat = [double]$c.location.y; Lon = [double]$c.location.x; Address = $c.address; Score = [double]$c.score }
        }
    } catch { throw "Could not reach the map service. Check the internet connection.`n$($_.Exception.Message)" }
    $script:GeoCache[$key] = $result
    return $result
}

function Get_Miles([double]$lat1, [double]$lon1, [double]$lat2, [double]$lon2) {
    $rad = [Math]::PI / 180
    $dLat = ($lat2 - $lat1) * $rad; $dLon = ($lon2 - $lon1) * $rad
    $a = [Math]::Sin($dLat / 2) * [Math]::Sin($dLat / 2) +
         [Math]::Cos($lat1 * $rad) * [Math]::Cos($lat2 * $rad) * [Math]::Sin($dLon / 2) * [Math]::Sin($dLon / 2)
    return 3958.8 * 2 * [Math]::Atan2([Math]::Sqrt($a), [Math]::Sqrt(1 - $a))
}

# Geocodes the given tickets (and the office once). Asks the tech to fix addresses that are not found.
function Resolve_Locations($tickets, $statusLabel) {
    if (-not $script:Office.Geocoded) {
        $o = Invoke_Geocode $script:Office.Address
        if ($o) { $script:Office.Lat = $o.Lat; $script:Office.Lon = $o.Lon }
        $script:Office | Add-Member -NotePropertyName Geocoded -NotePropertyValue $true -Force
    }
    $i = 0
    foreach ($t in $tickets) {
        $i++
        if ($null -ne $t.Lat) { continue }
        if ($statusLabel) { $statusLabel.Text = "Locating $i of $($tickets.Count): $($t.Customer)..."; [Windows.Forms.Application]::DoEvents() }
        $query = Get_FullAddress $t
        $geo = Invoke_Geocode $query
        while (-not $geo) {
            $fixed = Show_Input 'Address not found' ("Ticket #$($t.Number) - $($t.Customer)`nThe map could not find this address. Fix it, or Cancel to leave this ticket out.") $query
            if (-not $fixed) { break }
            $query = $fixed
            $geo = Invoke_Geocode $query
        }
        if ($geo) {
            $t.Lat = $geo.Lat; $t.Lon = $geo.Lon; $t.GeoAddress = $geo.Address
            $t.Miles = [Math]::Round((Get_Miles $script:Office.Lat $script:Office.Lon $geo.Lat $geo.Lon), 1)
            $t.GeoNote = if ($geo.Score -lt 90) { 'check address' } else { '' }
        }
    }
    if ($statusLabel) { $statusLabel.Text = '' }
}
#endregion

#region Route building
# Orders by distance (farthest first) and merges tickets at the same location into one stop.
function New_Stops($tickets) {
    $stops = New-Object System.Collections.Generic.List[object]
    foreach ($t in ($tickets | Where-Object { $null -ne $_.Lat } | Sort-Object @{ e = { $_.Miles }; Descending = $true }, Number)) {
        $same = $stops | Where-Object { (Get_Miles $_.Lat $_.Lon $t.Lat $t.Lon) -lt 0.05 } | Select-Object -First 1
        if ($same) { $same.Tickets += $t; continue }
        $stops.Add([pscustomobject]@{ Lat = $t.Lat; Lon = $t.Lon; Miles = $t.Miles; Address = $t.GeoAddress; Customer = $t.Customer; Tickets = @($t) })
    }
    return , $stops.ToArray()
}

function Get_MapsLinks($stops) {
    $links = @()
    $origin = $script:Office.Address
    for ($i = 0; $i -lt $stops.Count; $i += $script:MaxStopsPerLink) {
        $chunk = @($stops[$i..([Math]::Min($i + $script:MaxStopsPerLink, $stops.Count) - 1)])
        $dest = $chunk[-1].Address
        $url = 'https://www.google.com/maps/dir/?api=1&travelmode=driving' +
               '&origin=' + [uri]::EscapeDataString($origin) +
               '&destination=' + [uri]::EscapeDataString($dest)
        if ($chunk.Count -gt 1) {
            $url += '&waypoints=' + (($chunk[0..($chunk.Count - 2)] | ForEach-Object { [uri]::EscapeDataString($_.Address) }) -join '%7C')
        }
        $links += $url
        $origin = $dest
    }
    return , $links
}
#endregion

#region QR code (offline encoder, byte mode, ECC L..H auto)
# Built in so the route link (customer addresses) never goes to an online QR service.
# Compiled on first use; follows the QR Code Model 2 spec (ISO/IEC 18004).
$script:QrSource = @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

public static class TechRouteQr {
    // index = ECC level (L, M, Q, H), then version
    static readonly int[,] EccPerBlock = {
        {-1, 7,10,15,20,26,18,20,24,30,18,20,24,26,30,22,24,28,30,28,28,28,28,30,30,26,28,30,30,30,30,30,30,30,30,30,30,30,30,30,30},
        {-1,10,16,26,18,24,16,18,22,22,26,30,22,22,24,24,28,28,26,26,26,26,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28},
        {-1,13,22,18,26,18,24,18,22,20,24,28,26,24,20,30,24,28,28,26,30,28,30,30,30,30,28,30,30,30,30,30,30,30,30,30,30,30,30,30,30},
        {-1,17,28,22,16,22,28,26,26,24,28,24,28,22,24,24,30,28,28,26,28,30,24,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30}};
    static readonly int[,] NumBlocks = {
        {-1,1,1,1,1,1,2,2,2,2,4,4,4,4,4,6,6,6,6,7,8,8,9,9,10,12,12,12,13,14,15,16,17,18,19,19,20,21,22,24,25},
        {-1,1,1,1,2,2,4,4,4,5,5,5,8,9,9,10,10,11,13,14,16,17,17,18,20,21,23,25,26,28,29,31,33,35,37,38,40,43,45,47,49},
        {-1,1,1,2,2,4,4,6,6,8,8,8,10,12,16,12,17,16,18,21,20,23,23,25,27,29,34,34,35,38,40,43,45,48,51,53,56,59,62,65,68},
        {-1,1,1,2,4,4,4,5,6,8,8,11,11,16,16,18,16,19,21,25,25,25,34,30,32,35,37,40,42,45,48,51,54,57,60,63,66,70,74,77,81}};
    static readonly int[] EclFormat = { 1, 0, 3, 2 };

    static int size, version, ecl;
    static bool[,] dark, isFn;

    public static int Version { get { return version; } }

    // true = dark module, indexed [y, x]
    public static bool[,] Encode(string text) {
        byte[] data = Encoding.UTF8.GetBytes(text);
        version = 0;
        for (int v = 1; v <= 40; v++) {
            if (DataBits(data.Length, v) <= DataCodewords(v, 0) * 8) { version = v; break; }
        }
        if (version == 0) throw new ArgumentException("Link too long for a QR code (" + data.Length + " bytes).");
        ecl = 0;
        for (int e = 1; e <= 3; e++) { if (DataBits(data.Length, version) <= DataCodewords(version, e) * 8) ecl = e; }

        // data bits: byte mode, length, payload, terminator, padding
        var bits = new List<bool>();
        AppendBits(bits, 4, 4);
        AppendBits(bits, data.Length, version <= 9 ? 8 : 16);
        foreach (byte b in data) AppendBits(bits, b, 8);
        int capacity = DataCodewords(version, ecl) * 8;
        AppendBits(bits, 0, Math.Min(4, capacity - bits.Count));
        AppendBits(bits, 0, (8 - bits.Count % 8) % 8);
        for (int pad = 0xEC; bits.Count < capacity; pad ^= 0xEC ^ 0x11) AppendBits(bits, pad, 8);
        byte[] codewords = new byte[bits.Count / 8];
        for (int i = 0; i < bits.Count; i++) if (bits[i]) codewords[i >> 3] |= (byte)(0x80 >> (i & 7));

        size = version * 4 + 17;
        dark = new bool[size, size]; isFn = new bool[size, size];
        DrawFunctionPatterns();
        DrawCodewords(AddEccAndInterleave(codewords));

        int best = 0, bestPenalty = int.MaxValue;
        for (int m = 0; m < 8; m++) {
            ApplyMask(m); DrawFormatBits(m);
            int p = Penalty();
            if (p < bestPenalty) { best = m; bestPenalty = p; }
            ApplyMask(m);   // xor again = undo
        }
        ApplyMask(best); DrawFormatBits(best);
        return dark;
    }

    // black on white with a 4-module quiet zone; scale = pixels per module
    public static Bitmap Render(string text, int scale) {
        bool[,] m = Encode(text);
        int n = m.GetLength(0), px = (n + 8) * scale;
        var bmp = new Bitmap(px, px);
        using (var g = Graphics.FromImage(bmp)) {
            g.Clear(Color.White);
            for (int y = 0; y < n; y++)
                for (int x = 0; x < n; x++)
                    if (m[y, x]) g.FillRectangle(Brushes.Black, (x + 4) * scale, (y + 4) * scale, scale, scale);
        }
        return bmp;
    }

    public static int ModuleCount(string text) {
        byte[] data = Encoding.UTF8.GetBytes(text);
        for (int v = 1; v <= 40; v++) if (DataBits(data.Length, v) <= DataCodewords(v, 0) * 8) return v * 4 + 17;
        throw new ArgumentException("Link too long for a QR code (" + data.Length + " bytes).");
    }

    static int DataBits(int len, int v) { return 4 + (v <= 9 ? 8 : 16) + len * 8; }

    static void AppendBits(List<bool> bits, int val, int len) {
        for (int i = len - 1; i >= 0; i--) bits.Add(((val >> i) & 1) != 0);
    }

    static int RawDataModules(int v) {
        int r = (16 * v + 128) * v + 64;
        if (v >= 2) {
            int na = v / 7 + 2;
            r -= (25 * na - 10) * na - 55;
            if (v >= 7) r -= 36;
        }
        return r;
    }

    static int DataCodewords(int v, int e) { return RawDataModules(v) / 8 - EccPerBlock[e, v] * NumBlocks[e, v]; }

    static void SetFn(int x, int y, bool d) { dark[y, x] = d; isFn[y, x] = true; }

    static void DrawFunctionPatterns() {
        for (int i = 0; i < size; i++) { SetFn(6, i, i % 2 == 0); SetFn(i, 6, i % 2 == 0); }
        DrawFinder(3, 3); DrawFinder(size - 4, 3); DrawFinder(3, size - 4);
        int[] al = AlignmentPositions();
        int na = al.Length;
        for (int i = 0; i < na; i++)
            for (int j = 0; j < na; j++)
                if (!(i == 0 && j == 0 || i == 0 && j == na - 1 || i == na - 1 && j == 0)) DrawAlignment(al[i], al[j]);
        DrawFormatBits(0);   // reserve; real bits drawn after masking
        DrawVersion();
    }

    static void DrawFinder(int x, int y) {
        for (int dy = -4; dy <= 4; dy++)
            for (int dx = -4; dx <= 4; dx++) {
                int dist = Math.Max(Math.Abs(dx), Math.Abs(dy)), xx = x + dx, yy = y + dy;
                if (xx >= 0 && xx < size && yy >= 0 && yy < size) SetFn(xx, yy, dist != 2 && dist != 4);
            }
    }

    static void DrawAlignment(int x, int y) {
        for (int dy = -2; dy <= 2; dy++)
            for (int dx = -2; dx <= 2; dx++) SetFn(x + dx, y + dy, Math.Max(Math.Abs(dx), Math.Abs(dy)) != 1);
    }

    static int[] AlignmentPositions() {
        if (version == 1) return new int[0];
        int na = version / 7 + 2;
        int step = (version * 8 + na * 3 + 5) / (na * 4 - 4) * 2;
        int[] r = new int[na];
        r[0] = 6;
        for (int i = na - 1, pos = size - 7; i >= 1; i--, pos -= step) r[i] = pos;
        return r;
    }

    static bool Bit(int x, int i) { return ((x >> i) & 1) != 0; }

    static void DrawFormatBits(int mask) {
        int data = EclFormat[ecl] << 3 | mask, rem = data;
        for (int i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >> 9) * 0x537);
        int bits = (data << 10 | rem) ^ 0x5412;
        for (int i = 0; i <= 5; i++) SetFn(8, i, Bit(bits, i));
        SetFn(8, 7, Bit(bits, 6)); SetFn(8, 8, Bit(bits, 7)); SetFn(7, 8, Bit(bits, 8));
        for (int i = 9; i < 15; i++) SetFn(14 - i, 8, Bit(bits, i));
        for (int i = 0; i < 8; i++) SetFn(size - 1 - i, 8, Bit(bits, i));
        for (int i = 8; i < 15; i++) SetFn(8, size - 15 + i, Bit(bits, i));
        SetFn(8, size - 8, true);
    }

    static void DrawVersion() {
        if (version < 7) return;
        int rem = version;
        for (int i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >> 11) * 0x1F25);
        int bits = version << 12 | rem;
        for (int i = 0; i < 18; i++) {
            bool b = Bit(bits, i);
            int a = size - 11 + i % 3, c = i / 3;
            SetFn(a, c, b); SetFn(c, a, b);
        }
    }

    static byte[] AddEccAndInterleave(byte[] data) {
        int nb = NumBlocks[ecl, version], eccLen = EccPerBlock[ecl, version];
        int raw = RawDataModules(version) / 8, numShort = nb - raw % nb, shortLen = raw / nb;
        byte[] div = RsDivisor(eccLen);
        var blocks = new List<byte[]>();
        for (int i = 0, k = 0; i < nb; i++) {
            int datLen = shortLen - eccLen + (i < numShort ? 0 : 1);
            byte[] dat = new byte[datLen];
            Array.Copy(data, k, dat, 0, datLen);
            k += datLen;
            byte[] ecc = RsRemainder(dat, div);
            byte[] block = new byte[shortLen + 1];
            Array.Copy(dat, 0, block, 0, datLen);   // short blocks keep a gap at index datLen
            Array.Copy(ecc, 0, block, shortLen + 1 - eccLen, eccLen);
            blocks.Add(block);
        }
        var result = new List<byte>();
        for (int i = 0; i < shortLen + 1; i++)
            for (int j = 0; j < nb; j++)
                if (i != shortLen - eccLen || j >= numShort) result.Add(blocks[j][i]);
        return result.ToArray();
    }

    static byte[] RsDivisor(int degree) {
        byte[] r = new byte[degree];
        r[degree - 1] = 1;
        int root = 1;
        for (int i = 0; i < degree; i++) {
            for (int j = 0; j < degree; j++) {
                r[j] = (byte)RsMul(r[j], root);
                if (j + 1 < degree) r[j] ^= r[j + 1];
            }
            root = RsMul(root, 0x02);
        }
        return r;
    }

    static byte[] RsRemainder(byte[] data, byte[] div) {
        byte[] r = new byte[div.Length];
        foreach (byte b in data) {
            int factor = b ^ r[0];
            Array.Copy(r, 1, r, 0, r.Length - 1);
            r[r.Length - 1] = 0;
            for (int i = 0; i < r.Length; i++) r[i] ^= (byte)RsMul(div[i], factor);
        }
        return r;
    }

    static int RsMul(int x, int y) {
        int z = 0;
        for (int i = 7; i >= 0; i--) {
            z = (z << 1) ^ ((z >> 7) * 0x11D);
            z ^= ((y >> i) & 1) * x;
        }
        return z;
    }

    static void DrawCodewords(byte[] data) {
        int i = 0;
        for (int right = size - 1; right >= 1; right -= 2) {
            if (right == 6) right = 5;
            for (int vert = 0; vert < size; vert++)
                for (int j = 0; j < 2; j++) {
                    int x = right - j;
                    bool upward = ((right + 1) & 2) == 0;
                    int y = upward ? size - 1 - vert : vert;
                    if (!isFn[y, x] && i < data.Length * 8) {
                        dark[y, x] = Bit(data[i >> 3], 7 - (i & 7));
                        i++;
                    }
                }
        }
    }

    static void ApplyMask(int mask) {
        for (int y = 0; y < size; y++)
            for (int x = 0; x < size; x++) {
                bool inv;
                switch (mask) {
                    case 0: inv = (x + y) % 2 == 0; break;
                    case 1: inv = y % 2 == 0; break;
                    case 2: inv = x % 3 == 0; break;
                    case 3: inv = (x + y) % 3 == 0; break;
                    case 4: inv = (x / 3 + y / 2) % 2 == 0; break;
                    case 5: inv = x * y % 2 + x * y % 3 == 0; break;
                    case 6: inv = (x * y % 2 + x * y % 3) % 2 == 0; break;
                    default: inv = ((x + y) % 2 + x * y % 3) % 2 == 0; break;
                }
                if (inv && !isFn[y, x]) dark[y, x] = !dark[y, x];
            }
    }

    static bool At(int x, int y, bool rows) { return rows ? dark[y, x] : dark[x, y]; }

    // spec penalty rules N1-N4 (lower = easier to scan)
    static int Penalty() {
        int p = 0, darkCount = 0;
        for (int pass = 0; pass < 2; pass++) {
            bool rows = pass == 0;
            for (int a = 0; a < size; a++) {
                int run = 1;
                for (int b = 1; b < size; b++) {
                    if (At(b, a, rows) == At(b - 1, a, rows)) {
                        run++;
                        if (run == 5) p += 3; else if (run > 5) p++;
                    } else run = 1;
                }
                // 1:1:3:1:1 finder look-alike with 4 light modules on one side
                for (int b = 0; b + 10 < size; b++) {
                    bool core = At(b + 4, a, rows) && !At(b + 5, a, rows) && At(b + 6, a, rows) && At(b + 7, a, rows) && At(b + 8, a, rows) && !At(b + 9, a, rows) && At(b + 10, a, rows);
                    bool core2 = At(b, a, rows) && !At(b + 1, a, rows) && At(b + 2, a, rows) && At(b + 3, a, rows) && At(b + 4, a, rows) && !At(b + 5, a, rows) && At(b + 6, a, rows);
                    if (core && !At(b, a, rows) && !At(b + 1, a, rows) && !At(b + 2, a, rows) && !At(b + 3, a, rows)) p += 40;
                    if (core2 && !At(b + 7, a, rows) && !At(b + 8, a, rows) && !At(b + 9, a, rows) && !At(b + 10, a, rows)) p += 40;
                }
            }
        }
        for (int y = 0; y < size; y++)
            for (int x = 0; x < size; x++) {
                if (dark[y, x]) darkCount++;
                if (x + 1 < size && y + 1 < size) {
                    bool c = dark[y, x];
                    if (c == dark[y, x + 1] && c == dark[y + 1, x] && c == dark[y + 1, x + 1]) p += 3;
                }
            }
        int total = size * size;
        int k = (Math.Abs(darkCount * 20 - total * 10) + total - 1) / total - 1;
        p += Math.Max(0, k) * 10;
        return p;
    }
}
'@

# Returns a QR code bitmap of $text, about $targetPx wide (whole pixels per module).
function New_QrBitmap([string]$text, [int]$targetPx = 480) {
    if (-not ('TechRouteQr' -as [type])) { Add-Type -TypeDefinition $script:QrSource -ReferencedAssemblies System.Drawing }
    $scale = [Math]::Max(2, [Math]::Floor($targetPx / ([TechRouteQr]::ModuleCount($text) + 8)))
    return [TechRouteQr]::Render($text, $scale)
}
#endregion

#region Ticket forms (printing)
# Redraws "Ticket Detail - Helpdesk - Tecnica Systems.pdf" with GDI+, one page per stop.
# Coordinates are in the PDF's screen pixels (950 px = 8.5 in); Draw_TicketForm scales them to the page.
$script:FormScale = 850 / 950
$script:PrintState = @{ Index = 0; Stops = @() }

function Get_FormFields($stop) {
    $tk = @($stop.Tickets)
    $multi = $tk.Count -gt 1
    return [pscustomobject]@{
        Number      = ($tk | ForEach-Object { $_.Number }) -join ' / '
        Date        = (Get-Date).ToString('MM/dd/yyyy')
        Company     = $tk[0].Customer
        Subject     = ($tk | ForEach-Object { if ($multi) { "#$($_.Number): $($_.Subject)" } else { $_.Subject } }) -join "`n"
        Description = ($tk | ForEach-Object { if ($multi) { "#$($_.Number): $($_.Description)" } else { $_.Description } }) -join "`n"
        Address     = if ($stop.Address) { $stop.Address } else { Get_FullAddress $tk[0] }
    }
}

# $g must already map 1 unit = 1/100 inch on a Letter page (850 x 1100).
function Draw_TicketForm([Drawing.Graphics]$g, $stop) {
    $f = Get_FormFields $stop
    $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
    $g.ScaleTransform($script:FormScale, $script:FormScale)
    $g.TranslateTransform(0, -18)   # the browser header line of the PDF is not printed
    $font  = New-Object Drawing.Font('Arial', 14, [Drawing.FontStyle]::Regular, [Drawing.GraphicsUnit]::World)
    $bold  = New-Object Drawing.Font('Arial', 14, [Drawing.FontStyle]::Bold, [Drawing.GraphicsUnit]::World)
    $big   = New-Object Drawing.Font('Arial', 24, [Drawing.FontStyle]::Regular, [Drawing.GraphicsUnit]::World)
    $small = New-Object Drawing.Font('Arial', 11, [Drawing.FontStyle]::Regular, [Drawing.GraphicsUnit]::World)
    $pen   = New-Object Drawing.Pen([Drawing.Color]::Black, 1.6)
    $thin  = New-Object Drawing.Pen([Drawing.Color]::FromArgb(170, 170, 170), 2)
    $ink   = [Drawing.Brushes]::Black
    $mid   = New-Object Drawing.StringFormat; $mid.LineAlignment = 'Center'; $mid.Trimming = 'EllipsisWord'
    $L = 45; $R = 908
    $cell = { param($text, $x1, $y1, $x2, $y2, $fnt = $font)
        $g.DrawString($text, $fnt, $ink, (New-Object Drawing.RectangleF(($x1 + 6), $y1, ($x2 - $x1 - 10), ($y2 - $y1))), $mid) }

    # header
    if ($script:FormLogo) { $g.DrawImage($script:FormLogo, 50, 40, 260, 83) }
    $g.DrawString("TICKET NUMBER : $($f.Number)", $big, $ink, 461, 47)
    $g.DrawString("Date: $($f.Date)", $big, $ink, 461, 83)

    # company / subject / description / address (rows grow when the text is long)
    $y = 133
    foreach ($row in @(@('Company', $f.Company), @('Subject', $f.Subject), @('Description', $f.Description), @('Address', $f.Address))) {
        $need = $g.MeasureString($row[1], $font, ($R - 200 - 10)).Height + 14
        $h = [Math]::Min(150, [Math]::Max(34, [Math]::Ceiling($need)))
        $g.DrawRectangle($pen, $L, $y, ($R - $L), $h); $g.DrawLine($pen, 200, $y, 200, ($y + $h))
        & $cell $row[0] $L $y 200 ($y + $h)
        & $cell $row[1] 200 $y $R ($y + $h)
        $y += $h
    }
    $shift = $y - 269   # extra height used by long subjects/descriptions

    # service description
    $g.DrawString('SERVICE DESCRIPTION', $font, $ink, 44, (279 + $shift))
    $y = 307 + $shift
    $g.DrawRectangle($pen, $L, $y, ($R - $L), 166)
    $g.DrawLine($pen, $L, ($y + 55), $R, ($y + 55)); $g.DrawLine($pen, $L, ($y + 110), $R, ($y + 110))
    $g.DrawLine($pen, 293, $y, 293, ($y + 166))
    $g.DrawLine($pen, 458, $y, 458, ($y + 55)); $g.DrawLine($pen, 740, $y, 740, ($y + 55))
    $g.DrawString("Arrival Date:`n______/______/______", $font, $ink, 53, ($y + 8))
    & $cell 'Arrival Time: _______' 297 $y 458 ($y + 55)
    & $cell 'Depart Date: ______/______/______' 462 $y 740 ($y + 55)
    & $cell 'Depart Time: _______' 744 $y $R ($y + 55)
    & $cell "Reference Workorder/Service`nOrder #" $L ($y + 55) 293 ($y + 110)
    & $cell 'Type of Service' $L ($y + 110) 293 ($y + 166)
    & $cell "Contract Maintenance: ___ COD: ___ Billable: ___ Warranty: ___ Depot: ___ Install: ___`nOther: ___" 293 ($y + 110) $R ($y + 166)

    # items table: header + 3 rows
    $y = 473 + $shift
    $cols = 45, 113, 240, 407, 470, 631, 772, 908
    $heads = 'ITEM', 'PRODUCT', 'DESCRIPTION', 'QTY', 'RESOLUTION', 'UNIT PRICE', 'EXT. PRICE'
    $g.DrawRectangle($pen, $L, $y, ($R - $L), 135)
    foreach ($x in $cols[1..6]) { $g.DrawLine($pen, $x, $y, $x, ($y + 135)) }
    foreach ($dy in 36, 69, 102) { $g.DrawLine($pen, $L, ($y + $dy), $R, ($y + $dy)) }
    for ($i = 0; $i -lt 7; $i++) { & $cell $heads[$i] $cols[$i] $y $cols[$i + 1] ($y + 36) $bold }

    # detail description: header + rows (fewer rows if the top grew, so it stays one page)
    $y = 630 + $shift
    $rows = [Math]::Max(2, 6 - [Math]::Ceiling($shift / 33.5))
    $h = 35 + $rows * 33.5
    $g.DrawRectangle($pen, $L, $y, ($R - $L), $h); $g.DrawLine($pen, 810, $y, 810, ($y + $h))
    for ($i = 0; $i -lt $rows; $i++) { $g.DrawLine($pen, $L, ($y + 35 + $i * 33.5), $R, ($y + 35 + $i * 33.5)) }
    & $cell 'Detail Description of Products & Services Provided' $L $y 810 ($y + 35) $bold
    & $cell 'Price' 810 $y $R ($y + 35) $bold

    # disclaimer + signatures
    $legal = 'Customer acknowledges satisfactory completion of the above described work and understands and agrees that Tecnica Business Systems, LLC is not responsible for any loss of business or any loss of data whether occurring as a result of any imperfection due to normal wear and tear, negligence, equipment failure, power disturbances occurring from natural disasters or utility company service interruptions and/or during the process of providing the services. Data recovery and system restoration due to negligence or natural disasters including but not limited to utility company service interruption is billable at a service rate using the normal time and material billing schedule at the time when services are provided. Tecnica Business Systems, LLC service rates are available and provided at your request by calling (800) 423-4046 and speaking to our Support Services Administrator.'
    $g.DrawString($legal, $small, $ink, (New-Object Drawing.RectangleF($L, 888, ($R - $L), 100)))
    $g.DrawString('Customer Name: _____________________________________', $font, $ink, 44, 1034)
    $g.DrawString('Tecnica Representative: _______________________________', $font, $ink, 478, 1034)
    $g.DrawRectangle($thin, 46, 1056, 190, 26)
    $g.DrawString('Customer Signature: __________________________________', $font, $ink, 44, 1127)
    $g.DrawString('Tecnica Rep. Signature: ______________________________', $font, $ink, 478, 1127)
    $g.DrawString('Date: __________________________', $font, $ink, 44, 1180)

    foreach ($o in $font, $bold, $big, $small, $pen, $thin, $mid) { $o.Dispose() }
}

function New_FormsDocument($stops) {
    $doc = New-Object Drawing.Printing.PrintDocument
    $doc.DocumentName = "TechRoute ticket forms $((Get-Date).ToString('MM-dd-yyyy'))"
    $doc.DefaultPageSettings.Margins = New-Object Drawing.Printing.Margins(0, 0, 0, 0)
    $script:PrintState.Stops = @($stops)
    $doc.Add_BeginPrint({ $script:PrintState.Index = 0 })
    $doc.Add_PrintPage({
        param($s, $e)
        # draw from the physical page corner, not the printer's printable-area corner
        $e.Graphics.TranslateTransform(-$e.PageSettings.HardMarginX, -$e.PageSettings.HardMarginY)
        $e.Graphics.ScaleTransform(($e.PageBounds.Width / 850), ($e.PageBounds.Height / 1100))
        Draw_TicketForm $e.Graphics $script:PrintState.Stops[$script:PrintState.Index]
        $script:PrintState.Index++
        $e.HasMorePages = $script:PrintState.Index -lt $script:PrintState.Stops.Count
    })
    return $doc
}

function Invoke_PrintForms($stops, [switch]$Preview) {
    $doc = New_FormsDocument $stops
    try {
        if ($Preview) {
            $dlg = New-Object Windows.Forms.PrintPreviewDialog
            $dlg.Document = $doc; $dlg.Width = 900; $dlg.Height = 1000; $dlg.StartPosition = 'CenterScreen'
            if ($script:Icon) { $dlg.Icon = $script:Icon }
            $null = $dlg.ShowDialog(); $dlg.Dispose()
        } else {
            $pd = New-Object Windows.Forms.PrintDialog
            $pd.Document = $doc; $pd.UseEXDialog = $true
            if ($pd.ShowDialog() -eq 'OK') {
                # always one-sided (set after the dialog, which can bring back the printer's default)
                $doc.PrinterSettings.Duplex = [Drawing.Printing.Duplex]::Simplex
                $doc.DefaultPageSettings.PrinterSettings.Duplex = [Drawing.Printing.Duplex]::Simplex
                $doc.Print()
            }
            $pd.Dispose()
        }
    } catch { Show_Message "Printing failed:`n$($_.Exception.Message)" 'Print' 'Error' }
    finally { $doc.Dispose() }
}
#endregion

#region UI helpers
function New_Button([string]$text, [int]$w = 260, [int]$h = 52, $font = $script:Theme.Big) {
    $b = New-Object Windows.Forms.Button
    $b.Text = $text; $b.Width = $w; $b.Height = $h; $b.Font = $font
    $b.FlatStyle = 'Flat'; $b.FlatAppearance.BorderSize = 0
    $b.FlatAppearance.MouseOverBackColor = $script:Theme.Hover
    $b.BackColor = $script:Theme.Button; $b.ForeColor = $script:Theme.Text
    $b.Cursor = [Windows.Forms.Cursors]::Hand
    $b.Margin = New-Object Windows.Forms.Padding(6)
    return $b
}

function New_Label([string]$text, $font = $script:Theme.Font, $color = $script:Theme.Text) {
    $l = New-Object Windows.Forms.Label
    $l.Text = $text; $l.Font = $font; $l.ForeColor = $color; $l.AutoSize = $true; $l.UseMnemonic = $false
    return $l
}

function New_Form([string]$title, [int]$w, [int]$h) {
    $f = New-Object Windows.Forms.Form
    $f.Text = $title; $f.ClientSize = New-Object Drawing.Size($w, $h)
    $f.StartPosition = 'CenterScreen'; $f.BackColor = $script:Theme.Bg; $f.ForeColor = $script:Theme.Text; $f.Font = $script:Theme.Font
    if ($script:Icon) { $f.Icon = $script:Icon }
    return $f
}

function New_Grid {
    $g = New-Object Windows.Forms.DataGridView
    $g.Dock = 'Fill'; $g.BackgroundColor = $script:Theme.Bg; $g.BorderStyle = 'None'; $g.GridColor = $script:Theme.Button
    $g.RowHeadersVisible = $false; $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false; $g.SelectionMode = 'FullRowSelect'; $g.MultiSelect = $false
    $g.AutoSizeRowsMode = 'AllCells'; $g.EnableHeadersVisualStyles = $false
    $g.ColumnHeadersDefaultCellStyle.BackColor = $script:Theme.Panel
    $g.ColumnHeadersDefaultCellStyle.ForeColor = $script:Theme.Brand
    $g.ColumnHeadersDefaultCellStyle.Font = $script:Theme.Bold
    $g.ColumnHeadersHeightSizeMode = 'AutoSize'
    $g.DefaultCellStyle.BackColor = $script:Theme.Bg; $g.DefaultCellStyle.ForeColor = $script:Theme.Text
    $g.DefaultCellStyle.SelectionBackColor = $script:Theme.Hover; $g.DefaultCellStyle.SelectionForeColor = $script:Theme.Text
    $g.DefaultCellStyle.WrapMode = 'True'; $g.DefaultCellStyle.Padding = New-Object Windows.Forms.Padding(4)
    $g.AlternatingRowsDefaultCellStyle.BackColor = $script:Theme.Panel
    return $g
}

function Add_GridColumn($grid, [string]$header, [int]$fill, [string]$type = 'Text') {
    $col = if ($type -eq 'Check') { New-Object Windows.Forms.DataGridViewCheckBoxColumn } else { New-Object Windows.Forms.DataGridViewTextBoxColumn }
    $col.HeaderText = $header; $col.SortMode = 'NotSortable'
    if ($fill -gt 0) { $col.AutoSizeMode = 'Fill'; $col.FillWeight = $fill } else { $col.AutoSizeMode = 'AllCells' }
    if ($type -ne 'Check') { $col.ReadOnly = $true }
    $null = $grid.Columns.Add($col)
}

function New_BottomBar {
    $bar = New-Object Windows.Forms.FlowLayoutPanel
    $bar.Dock = 'Bottom'; $bar.Height = 64; $bar.BackColor = $script:Theme.Panel
    $bar.FlowDirection = 'RightToLeft'; $bar.Padding = New-Object Windows.Forms.Padding(8, 4, 8, 4)
    return $bar
}

function Show_Message([string]$text, [string]$title = $script:App.Name, [string]$icon = 'Information') {
    $null = [Windows.Forms.MessageBox]::Show($text, $title, 'OK', $icon)
}

# Simple text prompt. Returns the text, or $null on Cancel.
function Show_Input([string]$title, [string]$prompt, [string]$default = '') {
    $f = New_Form $title 520 190
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $l = New_Label $prompt; $l.Location = New-Object Drawing.Point(16, 14); $l.MaximumSize = New-Object Drawing.Size(490, 0)
    $tb = New-Object Windows.Forms.TextBox
    $tb.Text = $default; $tb.Location = New-Object Drawing.Point(16, 80); $tb.Width = 488
    $tb.BackColor = $script:Theme.Panel; $tb.ForeColor = $script:Theme.Text; $tb.BorderStyle = 'FixedSingle'
    $ok = New_Button 'OK' 110 38 $script:Theme.Bold; $ok.Location = New-Object Drawing.Point(278, 130); $ok.DialogResult = 'OK'
    $cancel = New_Button 'Cancel' 110 38 $script:Theme.Bold; $cancel.Location = New-Object Drawing.Point(394, 130); $cancel.DialogResult = 'Cancel'
    $f.Controls.AddRange(@($l, $tb, $ok, $cancel)); $f.AcceptButton = $ok; $f.CancelButton = $cancel
    $res = $f.ShowDialog()
    $text = $tb.Text.Trim(); $f.Dispose()
    if ($res -eq 'OK' -and $text) { return $text }
    return $null
}

# Asks how many tickets the tech will do. Returns a number, or $null on Cancel.
function Show_CountPrompt([int]$max) {
    $f = New_Form 'How many tickets?' 420 200
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $l = New_Label "How many tickets will you do today?`nYou have $max open ticket(s)."; $l.Location = New-Object Drawing.Point(16, 14)
    $num = New-Object Windows.Forms.NumericUpDown
    $num.Minimum = 1; $num.Maximum = [Math]::Max(1, $max); $num.Value = [Math]::Min(5, [Math]::Max(1, $max))
    $num.Font = $script:Theme.Big; $num.Width = 100; $num.Location = New-Object Drawing.Point(16, 70)
    $num.BackColor = $script:Theme.Panel; $num.ForeColor = $script:Theme.Text
    $all = New_Button 'All' 90 38 $script:Theme.Bold; $all.Location = New-Object Drawing.Point(130, 68)
    $all.Add_Click({ $num.Value = $num.Maximum; $f.DialogResult = 'OK' })
    $ok = New_Button 'Make route' 130 38 $script:Theme.Bold; $ok.Location = New-Object Drawing.Point(150, 140); $ok.DialogResult = 'OK'
    $cancel = New_Button 'Cancel' 110 38 $script:Theme.Bold; $cancel.Location = New-Object Drawing.Point(290, 140); $cancel.DialogResult = 'Cancel'
    $f.Controls.AddRange(@($l, $num, $all, $ok, $cancel)); $f.AcceptButton = $ok; $f.CancelButton = $cancel
    $res = $f.ShowDialog()
    $value = [int]$num.Value; $f.Dispose()
    if ($res -eq 'OK') { return $value }
    return $null
}
#endregion

#region Screens
function Show_TicketList {
    $open = Get_OpenTickets
    $f = New_Form "My Tickets - $($script:TechName)" 1000 600
    $grid = New_Grid
    Add_GridColumn $grid 'Ticket #' 0
    Add_GridColumn $grid 'Customer Name' 22
    Add_GridColumn $grid 'Subject' 30
    Add_GridColumn $grid 'Description' 48
    foreach ($t in ($open | Sort-Object OpenDate, Number)) { $null = $grid.Rows.Add($t.Number, $t.Customer, $t.Subject, $t.Description) }
    $bar = New_BottomBar
    $close = New_Button 'Close' 140 44 $script:Theme.Bold; $close.Add_Click({ $f.Close() })
    $hidden = $script:Tickets.Count - $open.Count
    $info = New_Label ("$($open.Count) open ticket(s)" + $(if ($hidden) { "   ($hidden closed hidden)" } else { '' })) $script:Theme.Font $script:Theme.Muted
    $info.Margin = New-Object Windows.Forms.Padding(12, 18, 12, 0)
    $bar.Controls.AddRange(@($close, $info))
    $f.Controls.AddRange(@($grid, $bar))
    $null = $f.ShowDialog(); $f.Dispose()
}

function Show_CustomPicker {
    $open = @(Get_OpenTickets | Sort-Object OpenDate, Number)
    $f = New_Form 'Custom route - pick today''s tickets' 1000 600
    $grid = New_Grid
    $grid.SelectionMode = 'CellSelect'
    Add_GridColumn $grid 'Do' 0 'Check'
    Add_GridColumn $grid 'Ticket #' 0
    Add_GridColumn $grid 'Customer Name' 22
    Add_GridColumn $grid 'Subject' 30
    Add_GridColumn $grid 'Description' 48
    foreach ($t in $open) { $null = $grid.Rows.Add($false, $t.Number, $t.Customer, $t.Subject, $t.Description) }
    # toggle the check when clicking anywhere on the row
    $grid.Add_CellClick({ param($s, $e) if ($e.RowIndex -ge 0) { $c = $s.Rows[$e.RowIndex].Cells[0]; $c.Value = -not [bool]$c.Value; $s.EndEdit() } })
    $bar = New_BottomBar
    $go = New_Button 'Make route' 160 44 $script:Theme.Bold; $go.ForeColor = $script:Theme.Brand; $go.DialogResult = 'OK'
    $cancel = New_Button 'Cancel' 120 44 $script:Theme.Bold; $cancel.DialogResult = 'Cancel'
    $bar.Controls.AddRange(@($go, $cancel))
    $f.Controls.AddRange(@($grid, $bar)); $f.CancelButton = $cancel
    $picked = @()
    if ($f.ShowDialog() -eq 'OK') {
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) { if ([bool]$grid.Rows[$i].Cells[0].Value) { $picked += $open[$i] } }
    }
    $f.Dispose()
    return , $picked
}

# QR code of the Google Maps link(s), to open the route on the phone. One code per link part.
function Show_RouteQr($links) {
    try { $images = @($links | ForEach-Object { New_QrBitmap $_ 480 }) }
    catch { Show_Message "Could not make the QR code:`n$($_.Exception.Message)" 'QR code' 'Error'; return }
    $f = New_Form 'Route QR code' ([Math]::Max(560, 300 + 102 * $images.Count)) 640
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $pic = New-Object Windows.Forms.PictureBox
    $pic.Dock = 'Fill'; $pic.SizeMode = 'CenterImage'; $pic.BackColor = $script:Theme.Bg
    $state = @{ Index = 0 }
    $msg = 'Scan with the phone camera to open the route in Google Maps.'
    $info = New_Label $msg $script:Theme.Font $script:Theme.Muted
    $info.Dock = 'Top'; $info.AutoSize = $false; $info.Height = 40; $info.TextAlign = 'MiddleCenter'
    $bar = New_BottomBar
    $close = New_Button 'Close' 110 44 $script:Theme.Bold; $close.DialogResult = 'Cancel'
    $save = New_Button 'Save image' 140 44 $script:Theme.Bold
    $save.Add_Click({
        $dlg = New-Object Windows.Forms.SaveFileDialog
        $dlg.Filter = 'PNG image (*.png)|*.png'
        $part = if ($images.Count -gt 1) { " part $($state.Index + 1)" } else { '' }
        $dlg.FileName = "TechRoute $(Get-Date -Format 'yyyy-MM-dd')$part.png"
        if ($dlg.ShowDialog($f) -eq 'OK') { $images[$state.Index].Save($dlg.FileName, [Drawing.Imaging.ImageFormat]::Png) }
        $dlg.Dispose()
    }.GetNewClosure())
    $bar.Controls.AddRange(@($close, $save))
    $show = {
        param([int]$i)
        $state.Index = $i; $pic.Image = $images[$i]
        if ($images.Count -gt 1) { $info.Text = "Part $($i + 1) of $($images.Count) - $msg" }
    }.GetNewClosure()
    if ($images.Count -gt 1) {
        for ($p = $images.Count - 1; $p -ge 0; $p--) {
            $b = New_Button "Part $($p + 1)" 90 44 $script:Theme.Bold; $b.ForeColor = $script:Theme.Brand; $b.Tag = $p
            $b.Add_Click({ param($s) & $show $s.Tag }.GetNewClosure())
            $bar.Controls.Add($b)
        }
    }
    & $show 0
    $f.Controls.AddRange(@($pic, $info, $bar)); $f.CancelButton = $close
    $null = $f.ShowDialog(); $f.Dispose()
    $images | ForEach-Object { $_.Dispose() }
}

function Show_Route($stops, [string]$mode, $skipped) {
    $links = Get_MapsLinks $stops
    $f = New_Form "Route - $mode" 1180 600
    $grid = New_Grid
    Add_GridColumn $grid 'Stop' 0
    Add_GridColumn $grid 'Miles' 0
    Add_GridColumn $grid 'Customer' 22
    Add_GridColumn $grid 'Address' 30
    Add_GridColumn $grid 'Ticket(s)' 40
    $null = $grid.Rows.Add('Start', '0', 'OFFICE', $script:Office.Address, '')
    $n = 0
    $allOpen = Get_OpenTickets
    foreach ($s in $stops) {
        $n++
        $tk = ($s.Tickets | ForEach-Object { "#$($_.Number)  $($_.Subject)" + $(if ($_.GeoNote) { "  [$($_.GeoNote)]" } else { '' }) }) -join "`n"
        # other open tickets at the same place that were not selected
        $ids = $s.Tickets | ForEach-Object { $_.Number }
        $extra = $allOpen | Where-Object { $ids -notcontains $_.Number -and $null -ne $_.Lat -and (Get_Miles $s.Lat $s.Lon $_.Lat $_.Lon) -lt 0.05 }
        if (-not $extra) {
            $key = ($s.Tickets[0].Address -replace '[^A-Za-z0-9]', '').ToUpper()
            $extra = $allOpen | Where-Object { $ids -notcontains $_.Number -and ($_.Address -replace '[^A-Za-z0-9]', '').ToUpper() -eq $key }
        }
        if ($extra) { $tk += "`nAlso open here: " + (($extra | ForEach-Object { "#$($_.Number)" }) -join ', ') }
        $part = if ($links.Count -gt 1) { "  (part $([Math]::Floor(($n - 1) / $script:MaxStopsPerLink) + 1))" } else { '' }
        $null = $grid.Rows.Add("$n$part", "$($s.Miles)", $s.Customer, $s.Address, $tk)
    }

    $bar = New_BottomBar
    $close = New_Button 'Close' 120 44 $script:Theme.Bold; $close.Add_Click({ $f.Close() })
    $copy = New_Button 'Copy link' 140 44 $script:Theme.Bold
    $copy.Add_Click({ [Windows.Forms.Clipboard]::SetText(($links -join "`r`n")); $copy.Text = 'Copied!' }.GetNewClosure())
    $qr = New_Button 'QR code' 120 44 $script:Theme.Bold; $qr.Tag = $links; $qr.Add_Click({ param($s) Show_RouteQr @($s.Tag) })
    $print = New_Button 'Print forms' 140 44 $script:Theme.Bold; $print.Add_Click({ Invoke_PrintForms $stops })
    $preview = New_Button 'Preview' 110 44 $script:Theme.Bold; $preview.Add_Click({ Invoke_PrintForms $stops -Preview })
    $bar.Controls.AddRange(@($close, $copy, $qr, $print, $preview))
    for ($p = $links.Count - 1; $p -ge 0; $p--) {
        $label = if ($links.Count -gt 1) { "Google Maps part $($p + 1)" } else { 'Open in Google Maps' }
        $b = New_Button $label 210 44 $script:Theme.Bold; $b.ForeColor = $script:Theme.Brand
        $b.Tag = $links[$p]
        $b.Add_Click({ param($s) Start-Process $s.Tag })
        $bar.Controls.Add($b)
    }
    $note = "$($stops.Count) stop(s), farthest to closest."
    if ($links.Count -gt 1) { $note += " Google Maps allows $script:MaxStopsPerLink stops per link, so the route is split." }
    if ($skipped) { $note += "`nLeft out (address not found): " + (($skipped | ForEach-Object { "#$($_.Number)" }) -join ', ') }
    $info = New_Label $note $script:Theme.Font $(if ($skipped) { $script:Theme.Warn } else { $script:Theme.Muted })
    $info.Margin = New-Object Windows.Forms.Padding(12, 12, 12, 0)
    $bar.Controls.Add($info)
    $f.Controls.AddRange(@($grid, $bar))
    $f.Add_Shown({
        $q = [Windows.Forms.MessageBox]::Show($f, "Print the ticket forms for the $($stops.Count) store(s) on this route?`n(one page per store)", 'Print ticket forms', 'YesNo', 'Question')
        if ($q -eq 'Yes') { Invoke_PrintForms $stops }
    })
    $null = $f.ShowDialog(); $f.Dispose()
}

function Start_Route([string]$mode, $statusLabel) {
    $open = Get_OpenTickets
    if (-not $open) { Show_Message 'You have no open tickets.'; return }
    $chosen = @()
    switch ($mode) {
        'By Open Date' {
            $count = Show_CountPrompt $open.Count
            if (-not $count) { return }
            $chosen = @($open | Sort-Object OpenDate, { [long]($_.Number -replace '\D', '0') } | Select-Object -First $count)
        }
        'By Distance' {
            $count = Show_CountPrompt $open.Count
            if (-not $count) { return }
            Resolve_Locations $open $statusLabel
            $chosen = @($open | Where-Object { $null -ne $_.Lat } | Sort-Object Miles -Descending | Select-Object -First $count)
        }
        'Custom' {
            $chosen = Show_CustomPicker
            if (-not $chosen -or $chosen.Count -eq 0) { return }
        }
    }
    Resolve_Locations $chosen $statusLabel
    $skipped = @($chosen | Where-Object { $null -eq $_.Lat })
    $stops = New_Stops $chosen
    if ($stops.Count -eq 0) { Show_Message 'None of the selected addresses could be found on the map.' $script:App.Name 'Warning'; return }
    Show_Route $stops $mode $skipped
}

function Show_RouteMenu($statusLabel) {
    $f = New_Form 'Plan a route' 380 330
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $panel = New-Object Windows.Forms.FlowLayoutPanel
    $panel.Dock = 'Fill'; $panel.FlowDirection = 'TopDown'; $panel.Padding = New-Object Windows.Forms.Padding(54, 16, 0, 0)
    $panel.Controls.Add((New_Label 'How do you want to order the route?' $script:Theme.Bold $script:Theme.Muted))
    $script:RouteChoice = $null
    foreach ($m in 'By Open Date', 'By Distance', 'Custom') {
        $b = New_Button $m
        $b.Add_Click({ param($s) $script:RouteChoice = $s.Text; $f.Close() })
        $panel.Controls.Add($b)
    }
    $f.Controls.Add($panel)
    $null = $f.ShowDialog(); $f.Dispose()
    if ($script:RouteChoice) {
        try { Start_Route $script:RouteChoice $statusLabel }
        catch { Show_Message $_.Exception.Message 'Route error' 'Error' }
    }
}

function Import_File([string]$path) {
    $script:Tickets = Read_TicketsXlsx $path
    $script:FilePath = $path
    $techs = @($script:Tickets | ForEach-Object { $_.Tech } | Where-Object { $_ } | Select-Object -Unique)
    $script:TechName = if ($techs.Count -eq 1) { $techs[0] } elseif ($techs.Count -gt 1) { "$($techs.Count) technicians" } else { 'Technician' }
}

function Show_Home {
    $f = New_Form "$($script:App.Name) - Tecnica Systems" 460 470
    $f.FormBorderStyle = 'FixedSingle'; $f.MaximizeBox = $false

    $title = New_Label 'TechRoute' $script:Theme.Title $script:Theme.Brand
    $title.Location = New-Object Drawing.Point(24, 18)
    $tech = New_Label '' $script:Theme.Big; $tech.Location = New-Object Drawing.Point(28, 72)
    $file = New_Label '' $script:Theme.Font $script:Theme.Muted; $file.Location = New-Object Drawing.Point(28, 100)
    $file.MaximumSize = New-Object Drawing.Size(410, 0)

    $refresh = {
        $open = @(Get_OpenTickets).Count
        $tech.Text = "$($script:TechName)  -  $open open ticket(s)"
        $file.Text = if ($script:FilePath) { [IO.Path]::GetFileName($script:FilePath) + '  (' + (Get-Item $script:FilePath).LastWriteTime.ToString('MM/dd/yyyy h:mm tt') + ')' } else { 'No tickets file loaded' }
    }

    $status = New_Label '' $script:Theme.Font $script:Theme.Warn; $status.Location = New-Object Drawing.Point(28, 410)

    $bList = New_Button 'List My Tickets' 400 56; $bList.Location = New-Object Drawing.Point(30, 150)
    $bList.Add_Click({ if ($script:FilePath) { Show_TicketList } })
    $bRoute = New_Button 'Plan a Route' 400 56; $bRoute.Location = New-Object Drawing.Point(30, 216); $bRoute.ForeColor = $script:Theme.Brand
    $bRoute.Add_Click({ if ($script:FilePath) { Show_RouteMenu $status } })
    $bLoad = New_Button 'Load another tickets file...' 400 44 $script:Theme.Bold; $bLoad.Location = New-Object Drawing.Point(30, 292)
    $bLoad.Add_Click({
        $dlg = New-Object Windows.Forms.OpenFileDialog
        $dlg.Filter = 'Excel tickets report (*.xlsx)|*.xlsx'
        $dlg.InitialDirectory = Join-Path $env:USERPROFILE 'Downloads'
        if ($dlg.ShowDialog() -eq 'OK') {
            try { Import_File $dlg.FileName; & $refresh } catch { Show_Message $_.Exception.Message 'Could not read file' 'Error' }
        }
    })
    $bExit = New_Button 'Exit' 400 44 $script:Theme.Bold; $bExit.Location = New-Object Drawing.Point(30, 346)
    $bExit.Add_Click({ $f.Close() })

    $ver = New_Label "v$($script:App.Version)  -  Tecnica Systems" (New-Object Drawing.Font('Segoe UI', 8)) $script:Theme.Muted
    $ver.Location = New-Object Drawing.Point(28, 440)

    $f.Controls.AddRange(@($title, $tech, $file, $bList, $bRoute, $bLoad, $bExit, $status, $ver))
    & $refresh
    [Windows.Forms.Application]::Run($f)
}
#endregion

#region Embedded images (SMS-Toolkit.ico 16-64 px, Tecnica_Form_Logo.png)
$script:IconB64 = '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'
$script:LogoPngB64 = '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'
#endregion

#region Main
$ico = Join-Path $script:App.Assets 'SMS-Toolkit.ico'
# files in Assets win; the embedded copies keep the exe working on its own
$script:Icon = if (Test-Path $ico) { New-Object Drawing.Icon($ico) } else { New-Object Drawing.Icon((New-Object IO.MemoryStream(, [Convert]::FromBase64String($script:IconB64)))) }
$logo = Join-Path $script:App.Assets 'Tecnica_Form_Logo.png'
$script:FormLogo = if (Test-Path $logo) { [Drawing.Image]::FromFile($logo) } else { [Drawing.Image]::FromStream((New-Object IO.MemoryStream(, [Convert]::FromBase64String($script:LogoPngB64)))) }

if ($SelfTest) {
    $path = Find_TicketsFile
    Import_File $path
    "File: $path"; "Tech: $($script:TechName)  tickets: $($script:Tickets.Count)  open: $(@(Get_OpenTickets).Count)"
    $open = Get_OpenTickets
    Resolve_Locations $open $null
    "Office: $($script:Office.Lat), $($script:Office.Lon)"
    $open | Sort-Object OpenDate, Number | ForEach-Object { '{0} {1:MM/dd} {2,5} mi  {3,-22} {4}' -f $_.Number, $_.OpenDate, $_.Miles, $_.Customer, $_.GeoAddress }
    $chosen = @($open | Sort-Object OpenDate, Number | Select-Object -First 4)
    "--- By Open Date, 4 tickets:"
    $stops = New_Stops $chosen
    $stops | ForEach-Object { '{0,5} mi  {1}  [{2}]' -f $_.Miles, $_.Customer, (($_.Tickets | ForEach-Object Number) -join ',') }
    "--- By Distance, All:"
    $stops = New_Stops $open
    $stops | ForEach-Object { '{0,5} mi  {1}  [{2}]' -f $_.Miles, $_.Customer, (($_.Tickets | ForEach-Object Number) -join ',') }
    $links = Get_MapsLinks $stops
    $links
    # QR codes of the links (saved as qr_N.png when TECHROUTE_SNAPSHOT is set)
    $i = 0
    foreach ($l in $links) {
        $i++
        $bmp = New_QrBitmap $l 480
        "QR part ${i}: version $([TechRouteQr]::Version), $($bmp.Width) px, $($l.Length) chars"
        if ($env:TECHROUTE_SNAPSHOT) { $bmp.Save((Join-Path $env:TECHROUTE_SNAPSHOT "qr_$i.png")) }
        $bmp.Dispose()
    }
    # render the ticket forms to PNG (TECHROUTE_SNAPSHOT=<dir>) to check the layout without printing
    if ($env:TECHROUTE_SNAPSHOT) {
        $i = 0
        foreach ($s in $stops) {
            $i++
            $bmp = New-Object Drawing.Bitmap 1275, 1650
            $g = [Drawing.Graphics]::FromImage($bmp); $g.Clear([Drawing.Color]::White); $g.ScaleTransform(1.5, 1.5)
            Draw_TicketForm $g $s
            $bmp.Save((Join-Path $env:TECHROUTE_SNAPSHOT "form_$i.png")); $g.Dispose(); $bmp.Dispose()
        }
        "Forms rendered: $i"
    }
    return
}

try {
    $path = Find_TicketsFile
    if ($path) { Import_File $path }
} catch { Show_Message $_.Exception.Message 'Could not read tickets file' 'Error' }
Show_Home
#endregion
