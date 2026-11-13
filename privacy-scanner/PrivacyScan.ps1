# 매장 PC 개인정보 파일 점검 (PrivacyScan.ps1)
# - 인터넷 통신 기능 없음. 검사 결과는 화면에만 표시됩니다.
# - 검사 대상: txt/csv, 엑셀(xlsx/xlsm/xls), 워드(docx/doc), 한글(hwpx/hwp), PDF, 사진(jpg/png 등, Windows OCR)
# - 체크한 파일을 영구 삭제(선택 시 0으로 덮어쓴 뒤 삭제)할 수 있습니다.
# 실행: 같은 폴더의 Run-PrivacyScan.bat 더블클릭

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-WindowStyle', 'Hidden', '-File', "`"$PSCommandPath`"")
    return
}

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, Microsoft.VisualBasic

trap {
    [Windows.Forms.MessageBox]::Show("프로그램 오류가 발생했습니다.`r`n`r`n$($_.Exception.Message)", '개인정보 파일 점검', 'OK', 'Error') | Out-Null
    break
}

# ---------------------------------------------------------------------------
# 파일 내용 읽기 도우미 (C#) - PDF, Office XML, IFilter, 바이너리 문자열
# ---------------------------------------------------------------------------
$helperSource = @'
using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Xml;

public static class PsXml {
    static readonly HashSet<string> Breaks = new HashSet<string>(new string[] { "p", "tc", "c", "si", "row", "tr", "td", "tab", "br", "cell", "lineBreak" });
    public static string Text(Stream s, int maxChars) {
        XmlReaderSettings set = new XmlReaderSettings();
        set.DtdProcessing = DtdProcessing.Prohibit;
        set.XmlResolver = null;
        set.CloseInput = true;
        StringBuilder sb = new StringBuilder();
        using (XmlReader xr = XmlReader.Create(s, set)) {
            while (xr.Read()) {
                switch (xr.NodeType) {
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.SignificantWhitespace:
                        sb.Append(xr.Value); break;
                    case XmlNodeType.Element:
                        if (xr.IsEmptyElement && Breaks.Contains(xr.LocalName)) sb.Append(' ');
                        break;
                    case XmlNodeType.EndElement:
                        if (Breaks.Contains(xr.LocalName)) sb.Append(' ');
                        break;
                }
                if (sb.Length > maxChars) break;
            }
        }
        return sb.ToString();
    }
}

public static class PsRaw {
    static string Clean(string s) {
        return Regex.Replace(s, "[^ -~가-힣ㄱ-ㆎ]+", " ");
    }
    public static string Text(byte[] d) {
        StringBuilder sb = new StringBuilder();
        sb.Append(Clean(Encoding.Unicode.GetString(d))).Append('\n');
        if (d.Length > 1) sb.Append(Clean(Encoding.Unicode.GetString(d, 1, d.Length - 1))).Append('\n');
        sb.Append(Clean(Encoding.GetEncoding(28591).GetString(d)));
        return sb.ToString();
    }
}

public static class PsIFilter {
    [StructLayout(LayoutKind.Sequential)]
    public struct FULLPROPSPEC { public Guid guidPropSet; public uint ulKind; public IntPtr data; }
    [StructLayout(LayoutKind.Sequential)]
    public struct STAT_CHUNK {
        public uint idChunk; public int breakType; public int flags; public uint locale;
        public FULLPROPSPEC attribute; public uint idChunkSource; public uint cwcStartSource; public uint cwcLenSource;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct FILTERREGION { public uint idChunk; public uint cwcStart; public uint cwcExtent; }

    [ComImport, Guid("89BCB740-6119-101A-BCB7-00DD010655AF"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFilter {
        [PreserveSig] int Init(uint grfFlags, uint cAttributes, IntPtr aAttributes, out uint pdwFlags);
        [PreserveSig] int GetChunk(out STAT_CHUNK pStat);
        [PreserveSig] int GetText(ref uint pcwcBuffer, IntPtr awcBuffer);
        [PreserveSig] int GetValue(out IntPtr ppPropValue);
        [PreserveSig] int BindRegion(FILTERREGION origPos, ref Guid riid, out IntPtr ppunk);
    }

    [DllImport("query.dll", CharSet = CharSet.Unicode)]
    static extern int LoadIFilter(string pwcsPath, [MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter, [MarshalAs(UnmanagedType.IUnknown)] out object ppIUnk);

    public static string Read(string path, int maxChars) {
        object unk;
        int hr = LoadIFilter(path, null, out unk);
        if (hr != 0 || unk == null) throw new Exception("IFilter 없음 0x" + hr.ToString("X8"));
        IFilter f = (IFilter)unk;
        StringBuilder sb = new StringBuilder();
        IntPtr buf = Marshal.AllocHGlobal(8192 * 2);
        try {
            uint fl;
            hr = f.Init(1 | 2 | 16, 0, IntPtr.Zero, out fl);
            if (hr < 0) throw new Exception("IFilter 초기화 실패 0x" + hr.ToString("X8"));
            int guard = 0;
            while (sb.Length < maxChars && guard++ < 200000) {
                STAT_CHUNK st;
                hr = f.GetChunk(out st);
                if (hr == unchecked((int)0x80041707) || hr == unchecked((int)0x80041708)) continue;
                if (hr != 0) break;
                if ((st.flags & 1) == 0) continue;
                while (sb.Length < maxChars) {
                    uint n = 8192;
                    hr = f.GetText(ref n, buf);
                    if (hr < 0) break;
                    if (n > 0) sb.Append(Marshal.PtrToStringUni(buf, (int)Math.Min(n, 8192u)));
                    if (hr == 0x00041709) break;
                }
                sb.Append('\n');
            }
        } finally {
            Marshal.FreeHGlobal(buf);
            Marshal.ReleaseComObject(f);
        }
        return sb.ToString();
    }
}

public static class PsPdf {
    static readonly Encoding L1 = Encoding.GetEncoding(28591);

    public static string Extract(byte[] data, int maxImages, int maxChars, List<byte[]> images) {
        string s = L1.GetString(data);
        Dictionary<int, string> m1 = new Dictionary<int, string>();
        Dictionary<int, string> m2 = new Dictionary<int, string>();
        List<byte[]> contents = new List<byte[]>();
        int pos = 0;
        while (pos < s.Length) {
            int k = s.IndexOf("stream", pos, StringComparison.Ordinal);
            if (k < 0) break;
            if (k >= 3 && string.CompareOrdinal(s, k - 3, "end", 0, 3) == 0) { pos = k + 6; continue; }
            int start = k + 6;
            if (start < s.Length && s[start] == '\r') start++;
            if (start < s.Length && s[start] == '\n') start++;
            int end = s.IndexOf("endstream", start, StringComparison.Ordinal);
            if (end < 0) break;
            pos = end + 9;
            int dictStart = s.LastIndexOf("obj", k, StringComparison.Ordinal);
            if (dictStart < 0 || k - dictStart > 4096) dictStart = Math.Max(0, k - 1024);
            string dict = s.Substring(dictStart, k - dictStart);
            int len = end - start;
            while (len > 0 && (data[start + len - 1] == '\n' || data[start + len - 1] == '\r')) len--;
            if (len <= 0) continue;
            bool isImage = dict.Contains("/Image");
            if (dict.Contains("/DCTDecode")) {
                if (isImage && images != null && images.Count < maxImages && len > 10000 && !dict.Contains("/FlateDecode")) {
                    byte[] b = new byte[len];
                    Buffer.BlockCopy(data, start, b, 0, len);
                    images.Add(b);
                }
                continue;
            }
            if (isImage) continue;
            if (dict.Contains("/Length1") || dict.Contains("/Length2") || dict.Contains("Type1C") || dict.Contains("/OpenType") || dict.Contains("/XRef") || dict.Contains("/ObjStm") || dict.Contains("/Metadata")) continue;
            if (dict.Contains("/Filter") && !dict.Contains("/FlateDecode")) continue;
            byte[] raw = new byte[len];
            Buffer.BlockCopy(data, start, raw, 0, len);
            byte[] dec = dict.Contains("/FlateDecode") ? Inflate(raw) : raw;
            if (dec == null || dec.Length == 0) continue;
            string ds = L1.GetString(dec);
            if (ds.Contains("begincmap")) { ParseCMap(ds, m1, m2); continue; }
            contents.Add(dec);
        }
        StringBuilder sb = new StringBuilder();
        foreach (byte[] c in contents) {
            if (sb.Length > maxChars) break;
            ExtractText(L1.GetString(c), m1, m2, sb);
            sb.Append('\n');
        }
        return sb.ToString();
    }

    static byte[] Inflate(byte[] raw) {
        MemoryStream outMs = new MemoryStream();
        try {
            int off = (raw.Length > 2 && (raw[0] & 0x0F) == 8) ? 2 : 0;
            using (MemoryStream ms = new MemoryStream(raw, off, raw.Length - off))
            using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress)) {
                byte[] buf = new byte[65536];
                int n;
                while ((n = ds.Read(buf, 0, buf.Length)) > 0) {
                    outMs.Write(buf, 0, n);
                    if (outMs.Length > 50000000) break;
                }
            }
        } catch { }
        return outMs.Length > 0 ? outMs.ToArray() : null;
    }

    static string Hex16(string h) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i + 4 <= h.Length; i += 4) sb.Append((char)Convert.ToInt32(h.Substring(i, 4), 16));
        return sb.ToString();
    }

    static void Put(int code, int hexWidth, string val, Dictionary<int, string> m1, Dictionary<int, string> m2) {
        if (hexWidth <= 2) m1[code] = val; else m2[code] = val;
    }

    static void ParseCMap(string s, Dictionary<int, string> m1, Dictionary<int, string> m2) {
        try {
            foreach (Match blk in Regex.Matches(s, @"beginbfchar(.*?)endbfchar", RegexOptions.Singleline)) {
                foreach (Match m in Regex.Matches(blk.Groups[1].Value, @"<([0-9A-Fa-f]{1,8})>\s*<([0-9A-Fa-f]*)>")) {
                    Put(Convert.ToInt32(m.Groups[1].Value, 16), m.Groups[1].Value.Length, Hex16(m.Groups[2].Value), m1, m2);
                }
            }
            foreach (Match blk in Regex.Matches(s, @"beginbfrange(.*?)endbfrange", RegexOptions.Singleline)) {
                foreach (Match m in Regex.Matches(blk.Groups[1].Value, @"<([0-9A-Fa-f]{1,8})>\s*<([0-9A-Fa-f]{1,8})>\s*(<[0-9A-Fa-f]*>|\[[^\]]*\])")) {
                    int lo = Convert.ToInt32(m.Groups[1].Value, 16);
                    int hi = Convert.ToInt32(m.Groups[2].Value, 16);
                    int width = m.Groups[1].Value.Length;
                    if (hi < lo || hi - lo > 65535) continue;
                    string d = m.Groups[3].Value;
                    if (d.StartsWith("[")) {
                        MatchCollection parts = Regex.Matches(d, @"<([0-9A-Fa-f]*)>");
                        for (int i = 0; i < parts.Count && lo + i <= hi; i++) Put(lo + i, width, Hex16(parts[i].Groups[1].Value), m1, m2);
                    } else {
                        string b = Hex16(d.Substring(1, d.Length - 2));
                        if (b.Length == 0) continue;
                        string head = b.Substring(0, b.Length - 1);
                        char last = b[b.Length - 1];
                        for (int c = lo; c <= hi; c++) Put(c, width, head + (char)(last + (c - lo)), m1, m2);
                    }
                }
            }
        } catch { }
    }

    static string Decode(byte[] b, Dictionary<int, string> m1, Dictionary<int, string> m2) {
        if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(b, 2, b.Length - 2);
        if (m2.Count > 0 && b.Length >= 2 && b.Length % 2 == 0) {
            StringBuilder sb = new StringBuilder();
            int hit = 0;
            for (int k = 0; k + 1 < b.Length; k += 2) {
                string v;
                if (m2.TryGetValue((b[k] << 8) | b[k + 1], out v)) { sb.Append(v); hit++; }
            }
            if (hit > 0) return sb.ToString();
        }
        if (m1.Count > 0) {
            StringBuilder sb = new StringBuilder();
            foreach (byte x in b) {
                string v;
                if (m1.TryGetValue(x, out v)) sb.Append(v); else sb.Append((char)x);
            }
            return sb.ToString();
        }
        return L1.GetString(b);
    }

    static void ExtractText(string c, Dictionary<int, string> m1, Dictionary<int, string> m2, StringBuilder sb) {
        int i = 0, n = c.Length;
        while (i < n) {
            char ch = c[i];
            if (ch == '(') {
                List<byte> bytes = new List<byte>();
                int depth = 1; i++;
                while (i < n && depth > 0) {
                    char x = c[i];
                    if (x == '\\' && i + 1 < n) {
                        char e = c[i + 1]; i += 2;
                        switch (e) {
                            case 'n': bytes.Add(10); break;
                            case 'r': bytes.Add(13); break;
                            case 't': bytes.Add(9); break;
                            case 'b': bytes.Add(8); break;
                            case 'f': bytes.Add(12); break;
                            case '\r': if (i < n && c[i] == '\n') i++; break;
                            case '\n': break;
                            default:
                                if (e >= '0' && e <= '7') {
                                    int v = e - '0'; int cnt = 1;
                                    while (cnt < 3 && i < n && c[i] >= '0' && c[i] <= '7') { v = v * 8 + (c[i] - '0'); i++; cnt++; }
                                    bytes.Add((byte)(v & 0xFF));
                                } else bytes.Add((byte)e);
                                break;
                        }
                        continue;
                    }
                    if (x == '(') depth++;
                    else if (x == ')') { depth--; if (depth == 0) { i++; break; } }
                    bytes.Add((byte)x); i++;
                }
                sb.Append(Decode(bytes.ToArray(), m1, m2));
            } else if (ch == '<') {
                if (i + 1 < n && c[i + 1] == '<') { i += 2; continue; }
                int j = c.IndexOf('>', i + 1);
                if (j < 0) break;
                string h = Regex.Replace(c.Substring(i + 1, j - i - 1), "[^0-9A-Fa-f]", "");
                if (h.Length % 2 == 1) h += "0";
                byte[] b = new byte[h.Length / 2];
                for (int k = 0; k < b.Length; k++) b[k] = Convert.ToByte(h.Substring(k * 2, 2), 16);
                sb.Append(Decode(b, m1, m2));
                i = j + 1;
            } else if (ch == '%') {
                while (i < n && c[i] != '\n' && c[i] != '\r') i++;
            } else if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch == '\'' || ch == '"') {
                int j = i;
                while (j < n && ((c[j] >= 'A' && c[j] <= 'Z') || (c[j] >= 'a' && c[j] <= 'z') || c[j] == '*' || c[j] == '\'' || c[j] == '"')) j++;
                string op = c.Substring(i, j - i);
                if (op == "BI") { int e = c.IndexOf("EI", j, StringComparison.Ordinal); i = e < 0 ? n : e + 2; continue; }
                if (op == "Td" || op == "TD" || op == "T*" || op == "Tm" || op == "BT" || op == "ET" || op == "'" || op == "\"") sb.Append(' ');
                i = j;
            } else i++;
        }
    }
}
'@
Add-Type -TypeDefinition $helperSource -ReferencedAssemblies 'System.Xml' -Language CSharp

# ---------------------------------------------------------------------------
# 검사 작업 (별도 스레드에서 실행)
# ---------------------------------------------------------------------------
$scanScript = {
    param($roots, $sync, $q, [bool]$full)
    $ErrorActionPreference = 'Stop'
    function Post($o) { $q.Enqueue($o) }

    try {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.Runtime.WindowsRuntime

        # --- Windows OCR 준비 ---
        $ocr = $null; $ocrMax = 2600
        try {
            $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
            $null = [Windows.Graphics.Imaging.BitmapTransform, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
            $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
            $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
            $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
            $asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
                $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetGenericArguments().Count -eq 1 -and
                $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
            } | Select-Object -First 1
            $ocr = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
            $ocrMax = [int][Windows.Media.Ocr.OcrEngine]::MaxImageDimension
        } catch { $ocr = $null }

        function Await($op, [type]$t) {
            $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op))
            $task.GetAwaiter().GetResult()
        }

        function Read-Ocr([byte[]]$bytes) {
            $ms = New-Object IO.MemoryStream(, $bytes)
            $ras = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($ms)
            try {
                $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($ras)) ([Windows.Graphics.Imaging.BitmapDecoder])
                $w = [double]$dec.PixelWidth; $h = [double]$dec.PixelHeight
                if ($w -lt 200 -or $h -lt 120) { return '' }
                $scale = [Math]::Min(1.0, ($ocrMax - 10) / [Math]::Max($w, $h))
                $tf = New-Object Windows.Graphics.Imaging.BitmapTransform
                $tf.ScaledWidth = [uint32][Math]::Max(1, [Math]::Floor($w * $scale))
                $tf.ScaledHeight = [uint32][Math]::Max(1, [Math]::Floor($h * $scale))
                $bmp = Await ($dec.GetSoftwareBitmapAsync(
                        [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
                        [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied,
                        $tf,
                        [Windows.Graphics.Imaging.ExifOrientationMode]::RespectExifOrientation,
                        [Windows.Graphics.Imaging.ColorManagementMode]::DoNotColorManage)) ([Windows.Graphics.Imaging.SoftwareBitmap])
                try {
                    $res = Await ($ocr.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType = WindowsRuntime])
                    return (@($res.Lines | ForEach-Object { $_.Text }) -join "`n")
                } finally { $bmp.Dispose() }
            } finally { $ras.Dispose(); $ms.Dispose() }
        }

        function Read-Bytes([string]$p) {
            $fs = New-Object IO.FileStream($p, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
            try {
                $b = New-Object byte[] $fs.Length
                $off = 0
                while ($off -lt $b.Length) { $n = $fs.Read($b, $off, $b.Length - $off); if ($n -le 0) { break }; $off += $n }
                return , $b
            } finally { $fs.Dispose() }
        }

        function Read-TextFile([string]$p) {
            $b = Read-Bytes $p
            if ($b.Length -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($b) }
            if ($b.Length -ge 2 -and $b[0] -eq 0xFE -and $b[1] -eq 0xFF) { return [Text.Encoding]::BigEndianUnicode.GetString($b) }
            try { return (New-Object Text.UTF8Encoding($false, $true)).GetString($b) }
            catch { return [Text.Encoding]::GetEncoding(949).GetString($b) }
        }

        # --- 찾을 개인정보 형태 ---
        $patterns = [ordered]@{
            '주민·외국인번호' = [regex]'(?<!\d)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\s?-?\s?[1-8]\d{6}(?!\d)'
            '휴대폰번호'     = [regex]'(?<!\d)01[016789][\s.\-]?\d{3,4}[\s.\-]?\d{4}(?!\d)'
            '운전면허번호'   = [regex]'(?<!\d)\d{2}\s?-\s?\d{2}\s?-\s?\d{6}\s?-\s?\d{2}(?!\d)'
            '카드번호'       = [regex]'(?<!\d)\d{4}[\s\-]?\d{4}[\s\-]?\d{4}[\s\-]?\d{4}(?!\d)'
            '신분증 문구'    = [regex]'주\s*민\s*등\s*록\s*증|운\s*전\s*면\s*허\s*증|외\s*국\s*인\s*등\s*록\s*증|대\s*한\s*민\s*국\s*여\s*권|(?i:passport)'
        }
        $nameRe = '신분증|주민등록|주민번호|운전면허|외국인등록|여권|개인정보|가입신청서'
        $maskRe = [regex]'(?<=\d[\s.\-]?\d[\s.\-]?\d[\s.\-]?\d[\s.\-]?\d[\s.\-]?\d[\s.\-]?)\d'

        function Test-Rrn([string]$v) {
            if ($v -match '-') { return $true }
            $d = $v -replace '\D', ''
            if ($d.Length -ne 13) { return $false }
            $w = 2, 3, 4, 5, 6, 7, 8, 9, 2, 3, 4, 5; $sum = 0
            for ($i = 0; $i -lt 12; $i++) { $sum += ([int]$d[$i] - 48) * $w[$i] }
            return (((11 - ($sum % 11)) % 10) -eq ([int]$d[12] - 48))
        }
        function Test-Card([string]$v) {
            $d = $v -replace '\D', ''
            if ($d -match '^(\d)\1+$') { return $false }
            $sum = 0; $alt = $false
            for ($i = $d.Length - 1; $i -ge 0; $i--) {
                $n = [int]$d[$i] - 48
                if ($alt) { $n *= 2; if ($n -gt 9) { $n -= 9 } }
                $sum += $n; $alt = -not $alt
            }
            return (($sum % 10) -eq 0)
        }
        function Add-Hits([string]$text, $found) {
            if ([string]::IsNullOrEmpty($text)) { return }
            foreach ($kind in $patterns.Keys) {
                foreach ($m in $patterns[$kind].Matches($text)) {
                    if ($kind -eq '주민·외국인번호' -and -not (Test-Rrn $m.Value)) { continue }
                    if ($kind -eq '카드번호' -and -not (Test-Card $m.Value)) { continue }
                    if (-not $found.Contains($kind)) {
                        $s = [Math]::Max(0, $m.Index - 20)
                        $e = [Math]::Min($text.Length, $m.Index + $m.Length + 20)
                        $ex = ($text.Substring($s, $e - $s) -replace '\s+', ' ').Trim()
                        $found[$kind] = @{ N = 0; Sample = $maskRe.Replace($ex, '*') }
                    }
                    $found[$kind].N++
                    if ($found[$kind].N -ge 9999) { break }
                }
            }
        }

        # --- 검사 대상 확장자 ---
        $textExt = @('.txt', '.csv', '.log')
        $zipExt = @('.xlsx', '.xlsm', '.docx', '.pptx', '.hwpx')
        $pdfExt = @('.pdf')
        $binExt = @('.xls', '.doc', '.hwp', '.ppt')
        $imgExt = @('.jpg', '.jpeg', '.jfif', '.png', '.bmp', '.gif', '.tif', '.tiff', '.webp', '.heic', '.heif')
        $allExt = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($e in ($textExt + $zipExt + $pdfExt + $binExt + $imgExt)) { [void]$allExt.Add($e) }

        $skipAny = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($n in @('$Recycle.Bin', 'System Volume Information', 'AppData', 'node_modules', '.git', '$WinREAgent', 'Config.Msi')) { [void]$skipAny.Add($n) }
        $skipRoot = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($n in @('Windows', 'Windows.old', 'Program Files', 'Program Files (x86)', 'ProgramData', 'Recovery', 'PerfLogs', 'MSOCache', 'Intel', 'AMD', 'NVIDIA')) { [void]$skipRoot.Add($n) }

        # --- 1단계: 파일 목록 만들기 ---
        $stack = New-Object 'Collections.Generic.Stack[string]'
        foreach ($r in $roots) { if ([IO.Directory]::Exists($r)) { $stack.Push($r) } }
        $visited = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $targets = New-Object 'Collections.Generic.List[string]'
        $seen = 0; $errors = 0; $skipped = 0; $inspected = 0
        $sw = [Diagnostics.Stopwatch]::StartNew()

        while ($stack.Count -gt 0 -and -not $sync.Stop) {
            $dir = $stack.Pop()
            if ($dir.Length -gt 400) { $skipped++; continue }
            if (-not $visited.Add($dir.TrimEnd('\'))) { continue }
            try { $entries = [IO.Directory]::GetFileSystemEntries($dir) } catch { $errors++; continue }
            $atRoot = ([IO.Path]::GetPathRoot($dir).TrimEnd('\') -eq $dir.TrimEnd('\'))
            foreach ($p in $entries) {
                try {
                    $a = [int][IO.File]::GetAttributes($p)
                    if ($a -band 0x10) {
                        $leaf = [IO.Path]::GetFileName($p)
                        if ($skipAny.Contains($leaf)) { continue }
                        if ($full -and $atRoot -and $skipRoot.Contains($leaf)) { continue }
                        # 예전 방식 연결 폴더(숨김·시스템 바로가기 폴더)는 건너뜀
                        if (($a -band 0x400) -and ($a -band 0x6)) { continue }
                        $stack.Push($p); continue
                    }
                    $seen++
                    if ($a -band 0x441000) { $skipped++; continue }   # 클라우드 전용(내려받지 않은) 파일
                    if ($allExt.Contains([IO.Path]::GetExtension($p))) { $targets.Add($p) }
                } catch { $errors++ }
            }
            if ($sw.ElapsedMilliseconds -ge 300) { Post @{ T = 'count'; Seen = $seen; Targets = $targets.Count }; $sw.Restart() }
        }

        # --- 2단계: 내용 검사 ---
        $done = 0; $total = $targets.Count
        Post @{ T = 'pct'; Done = 0; Total = $total; Pct = 0 }
        if (-not $ocr) { Post @{ T = 'msg'; Text = 'Windows 글자 인식(OCR)을 사용할 수 없어 사진은 파일명만 검사합니다.' } }

        foreach ($path in $targets) {
            if ($sync.Stop) { break }
            try {
                $fi = New-Object IO.FileInfo($path)
                $ext = $fi.Extension.ToLowerInvariant()
                $found = [ordered]@{}
                $note = ''
                $read = $true
                Add-Hits $fi.Name $found
                if ($fi.Name -match $nameRe) { $found['파일명'] = @{ N = 1; Sample = $fi.Name } }
                try {
                    if ($textExt -contains $ext) {
                        if ($fi.Length -le 32MB) { Add-Hits (Read-TextFile $path) $found } else { $read = $false; $note = '용량이 커서 내용 미검사' }
                    }
                    elseif ($zipExt -contains $ext) {
                        if ($fi.Length -le 100MB) {
                            $fs = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
                            try {
                                $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Read)
                                try {
                                    $imgCount = 0
                                    foreach ($en in $zip.Entries) {
                                        $nm = $en.FullName
                                        if ($nm -match '^(xl/sharedStrings\.xml|xl/worksheets/sheet\d+\.xml|word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml|ppt/slides/slide\d+\.xml|Contents/section\d+\.xml)$') {
                                            Add-Hits ([PsXml]::Text($en.Open(), 30000000)) $found
                                        }
                                        elseif ($ocr -and $imgCount -lt 15 -and $en.Length -ge 15KB -and $en.Length -le 30MB -and $nm -match '^((xl|word|ppt)/media|BinData)/[^/]+\.(png|jpe?g|bmp|gif|tiff?)$') {
                                            $imgCount++
                                            $es = $en.Open(); $mem = New-Object IO.MemoryStream
                                            try { $es.CopyTo($mem) } finally { $es.Dispose() }
                                            try { Add-Hits (Read-Ocr $mem.ToArray()) $found } catch { }
                                        }
                                    }
                                } finally { $zip.Dispose() }
                            } finally { $fs.Dispose() }
                        } else { $read = $false; $note = '용량이 커서 내용 미검사' }
                    }
                    elseif ($pdfExt -contains $ext) {
                        if ($fi.Length -le 100MB) {
                            $imgs = New-Object 'Collections.Generic.List[byte[]]'
                            $t = [PsPdf]::Extract((Read-Bytes $path), 10, 30000000, $imgs)
                            Add-Hits $t $found
                            if ($ocr) { foreach ($im in $imgs) { try { Add-Hits (Read-Ocr $im) $found } catch { } } }
                            if ($found.Count -eq 0) { try { Add-Hits ([PsIFilter]::Read($path, 30000000)) $found } catch { } }
                            if ($t.Trim().Length -eq 0 -and $imgs.Count -eq 0) { $note = 'PDF 글자를 읽지 못함(암호화·특수 형식)' }
                            elseif ($imgs.Count -gt 0 -and -not $ocr) { $note = 'PDF 속 사진은 OCR 없어 미검사' }
                        } else { $read = $false; $note = '용량이 커서 내용 미검사' }
                    }
                    elseif ($binExt -contains $ext) {
                        $ok = $false
                        try { $t = [PsIFilter]::Read($path, 30000000); if ($t.Trim().Length -gt 0) { Add-Hits $t $found; $ok = $true } } catch { }
                        if (-not $ok) {
                            if ($fi.Length -le 40MB) { Add-Hits ([PsRaw]::Text((Read-Bytes $path))) $found; $note = '부분 검사(전용 읽기 기능 없음)' }
                            else { $read = $false; $note = '용량이 커서 내용 미검사' }
                        }
                    }
                    elseif ($imgExt -contains $ext) {
                        if (-not $ocr) { $read = $false; $note = 'OCR 없음 - 파일명만 검사' }
                        elseif ($fi.Length -lt 15KB) { $read = $false }
                        elseif ($fi.Length -le 30MB) { Add-Hits (Read-Ocr (Read-Bytes $path)) $found }
                        else { $read = $false; $note = '용량이 커서 내용 미검사' }
                    }
                } catch {
                    $errors++; $read = $false; $note = '내용을 읽지 못함'
                }
                if ($read) { $inspected++ } else { $skipped++ }

                if ($found.Count -gt 0) {
                    $kinds = @($found.Keys)
                    $lines = @(foreach ($k in $kinds) { "$k ($($found[$k].N)건): $($found[$k].Sample)" })
                    if ($note) { $lines += "※ $note" }
                    $first = $found[$kinds[0]]
                    Post @{
                        T       = 'hit'
                        Path    = $path
                        Kinds   = $kinds
                        KindText = (($kinds | ForEach-Object { if ($found[$_].N -gt 1) { "$_ $($found[$_].N)" } else { $_ } }) -join ', ')
                        Summary = $first.Sample
                        Detail  = ($lines -join "`r`n")
                        Ticks   = $fi.LastWriteTimeUtc.Ticks
                        Length  = $fi.Length
                    }
                }
            } catch { $errors++ }
            $done++
            if ($sw.ElapsedMilliseconds -ge 250 -or $done -eq $total) {
                Post @{ T = 'pct'; Done = $done; Total = $total; Pct = $(if ($total) { [int][Math]::Floor(100.0 * $done / $total) } else { 100 }) }
                $sw.Restart()
            }
        }
        Post @{ T = 'done'; Seen = $seen; Total = $total; Done = $done; Inspected = $inspected; Skipped = $skipped; Errors = $errors; Stopped = [bool]$sync.Stop }
    } catch {
        Post @{ T = 'fatal'; Text = $_.Exception.Message }
    }
}

# ---------------------------------------------------------------------------
# 화면
# ---------------------------------------------------------------------------
[Windows.Forms.Application]::EnableVisualStyles()
$font = New-Object Drawing.Font('Malgun Gothic', 9.5)
$kindsAll = @('전체', '주민·외국인번호', '휴대폰번호', '운전면허번호', '카드번호', '신분증 문구', '파일명')

$script:hits = New-Object Collections.ArrayList
$script:ps = $null; $script:handle = $null; $script:q = $null; $script:sync = $null
$script:rebuilding = $false; $script:doneInfo = $null; $script:scanLabel = ''

function New-Btn([string]$text, [int]$w) {
    $b = New-Object Windows.Forms.Button
    $b.Text = $text; $b.Width = $w; $b.Height = 32
    $b.Margin = New-Object Windows.Forms.Padding(3)
    return $b
}
function Format-Size([long]$n) {
    if ($n -ge 1GB) { return ('{0:N1} GB' -f ($n / 1GB)) }
    if ($n -ge 1MB) { return ('{0:N1} MB' -f ($n / 1MB)) }
    if ($n -ge 1KB) { return ('{0:N0} KB' -f ($n / 1KB)) }
    return "$n B"
}

$form = New-Object Windows.Forms.Form
$form.Text = '개인정보 파일 점검 (인터넷 전송 없음)'
$form.Size = New-Object Drawing.Size(1200, 740)
$form.MinimumSize = New-Object Drawing.Size(900, 520)
$form.StartPosition = 'CenterScreen'
$form.Font = $font

$layout = New-Object Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'; $layout.ColumnCount = 1; $layout.RowCount = 5
[void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
[void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 100)))
[void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 96)))
[void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
[void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
$form.Controls.Add($layout)

# 위쪽: 검사 버튼
$top = New-Object Windows.Forms.FlowLayoutPanel
$top.AutoSize = $true; $top.Dock = 'Fill'; $top.WrapContents = $true
$btnQuick = New-Btn '빠른 검사' 100
$btnFull = New-Btn '전체 드라이브 검사' 150
$btnFolder = New-Btn '폴더 골라 검사' 130
$btnStop = New-Btn '중지' 70; $btnStop.Enabled = $false
$lblFilter = New-Object Windows.Forms.Label
$lblFilter.Text = '보기:'; $lblFilter.AutoSize = $true; $lblFilter.Margin = New-Object Windows.Forms.Padding(24, 11, 3, 0)
$cmb = New-Object Windows.Forms.ComboBox
$cmb.DropDownStyle = 'DropDownList'; $cmb.Width = 170; $cmb.Margin = New-Object Windows.Forms.Padding(3, 7, 3, 3)
[void]$cmb.Items.AddRange([object[]]$kindsAll); $cmb.SelectedIndex = 0
$top.Controls.AddRange(@($btnQuick, $btnFull, $btnFolder, $btnStop, $lblFilter, $cmb))
$layout.Controls.Add($top, 0, 0)

# 가운데: 결과 목록
$lv = New-Object Windows.Forms.ListView
$lv.View = 'Details'; $lv.CheckBoxes = $true; $lv.FullRowSelect = $true; $lv.GridLines = $true
$lv.HideSelection = $false; $lv.Dock = 'Fill'
foreach ($c in @(@('종류', 190), @('파일명', 230), @('위치', 320), @('수정일', 125), @('크기', 75), @('발견 내용 (숫자 일부 가림)', 400))) {
    [void]$lv.Columns.Add($c[0], $c[1])
}
$layout.Controls.Add($lv, 0, 1)

# 선택 항목 상세
$txtDetail = New-Object Windows.Forms.TextBox
$txtDetail.Multiline = $true; $txtDetail.ReadOnly = $true; $txtDetail.ScrollBars = 'Vertical'; $txtDetail.Dock = 'Fill'
$txtDetail.Text = '검사 버튼을 누르면 시작합니다. 빠른 검사는 바탕화면·문서·다운로드·사진 폴더만 검사합니다.'
$layout.Controls.Add($txtDetail, 0, 2)

# 진행 상황
$statusRow = New-Object Windows.Forms.TableLayoutPanel
$statusRow.ColumnCount = 2; $statusRow.RowCount = 1; $statusRow.Dock = 'Fill'; $statusRow.AutoSize = $true
[void]$statusRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 260)))
[void]$statusRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
$pb = New-Object Windows.Forms.ProgressBar
$pb.Dock = 'Fill'; $pb.Height = 20; $pb.Margin = New-Object Windows.Forms.Padding(3, 6, 3, 3)
$lblStatus = New-Object Windows.Forms.Label
$lblStatus.AutoSize = $true; $lblStatus.Margin = New-Object Windows.Forms.Padding(8, 8, 3, 3); $lblStatus.Text = '대기 중'
$statusRow.Controls.Add($pb, 0, 0); $statusRow.Controls.Add($lblStatus, 1, 0)
$layout.Controls.Add($statusRow, 0, 3)

# 아래쪽: 파일 작업
$bottom = New-Object Windows.Forms.FlowLayoutPanel
$bottom.AutoSize = $true; $bottom.Dock = 'Fill'; $bottom.WrapContents = $true
$btnOpen = New-Btn '파일 열기' 90
$btnShow = New-Btn '폴더에서 보기' 120
$btnCheckAll = New-Btn '보이는 항목 모두 체크' 170
$btnUncheck = New-Btn '체크 모두 해제' 120
$chkWipe = New-Object Windows.Forms.CheckBox
$chkWipe.Text = '0으로 덮어쓴 뒤 삭제 (복구 어렵게)'; $chkWipe.Checked = $true; $chkWipe.AutoSize = $true
$chkWipe.Margin = New-Object Windows.Forms.Padding(24, 10, 3, 3)
$btnDelete = New-Btn '체크한 파일 영구 삭제' 190
$btnDelete.BackColor = [Drawing.Color]::FromArgb(200, 40, 40); $btnDelete.ForeColor = [Drawing.Color]::White
$btnDelete.FlatStyle = 'Flat'; $btnDelete.Font = New-Object Drawing.Font('Malgun Gothic', 9.5, [Drawing.FontStyle]::Bold)
$btnDelete.Margin = New-Object Windows.Forms.Padding(16, 3, 3, 3)
$bottom.Controls.AddRange(@($btnOpen, $btnShow, $btnCheckAll, $btnUncheck, $chkWipe, $btnDelete))
$layout.Controls.Add($bottom, 0, 4)

# --- 목록 관리 ---
function Test-Visible($h) {
    $k = [string]$cmb.SelectedItem
    return ($k -eq '전체' -or ($h.Kinds -contains $k))
}
function New-Row($h) {
    $it = New-Object Windows.Forms.ListViewItem([string]$h.KindText)
    [void]$it.SubItems.Add([IO.Path]::GetFileName($h.Path))
    [void]$it.SubItems.Add([IO.Path]::GetDirectoryName($h.Path))
    [void]$it.SubItems.Add((New-Object DateTime([long]$h.Ticks, [DateTimeKind]::Utc)).ToLocalTime().ToString('yyyy-MM-dd HH:mm'))
    [void]$it.SubItems.Add((Format-Size ([long]$h.Length)))
    [void]$it.SubItems.Add([string]$h.Summary)
    $it.Checked = [bool]$h.Checked
    $it.Tag = $h
    return $it
}
function Update-Counts {
    $checked = @($script:hits | Where-Object { $_.Checked }).Count
    $btnDelete.Text = "체크한 파일 영구 삭제 ($checked)"
}
function Rebuild-List {
    $script:rebuilding = $true
    $lv.BeginUpdate()
    $lv.Items.Clear()
    $rows = New-Object Collections.ArrayList
    foreach ($h in $script:hits) { if (Test-Visible $h) { [void]$rows.Add((New-Row $h)) } }
    if ($rows.Count) { $lv.Items.AddRange([Windows.Forms.ListViewItem[]]$rows.ToArray()) }
    $lv.EndUpdate()
    $script:rebuilding = $false
    Update-Counts
}

function Set-Scanning([bool]$on) {
    $btnQuick.Enabled = -not $on; $btnFull.Enabled = -not $on; $btnFolder.Enabled = -not $on
    $btnDelete.Enabled = -not $on
    $btnStop.Enabled = $on
}

function Get-QuickRoots {
    $list = New-Object Collections.ArrayList
    foreach ($sf in @('Desktop', 'MyDocuments', 'MyPictures', 'CommonDesktopDirectory', 'CommonDocuments')) {
        try { [void]$list.Add([Environment]::GetFolderPath($sf)) } catch { }
    }
    $dl = $null
    try { $dl = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -ErrorAction Stop).'{374DE290-123F-4565-9164-39C4925E467B}' } catch { }
    if (-not $dl) { $dl = Join-Path $env:USERPROFILE 'Downloads' }
    [void]$list.Add($dl)
    return @($list | Where-Object { $_ -and [IO.Directory]::Exists($_) } | Select-Object -Unique)
}

function Start-Scan([string[]]$roots, [bool]$full, [string]$label) {
    if ($script:ps) { return }
    if (-not $roots -or $roots.Count -eq 0) {
        [Windows.Forms.MessageBox]::Show('검사할 폴더를 찾지 못했습니다.', '개인정보 파일 점검') | Out-Null
        return
    }
    $script:hits.Clear(); Rebuild-List
    $script:doneInfo = $null; $script:scanLabel = $label
    $txtDetail.Text = "검사 범위:`r`n" + ($roots -join "`r`n")
    $script:sync = [hashtable]::Synchronized(@{ Stop = $false })
    $script:q = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'; $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()
    $script:ps = [powershell]::Create()
    $script:ps.Runspace = $rs
    [void]$script:ps.AddScript($scanScript.ToString()).AddArgument($roots).AddArgument($script:sync).AddArgument($script:q).AddArgument($full)
    $script:handle = $script:ps.BeginInvoke()
    $pb.Style = 'Marquee'; $pb.Value = 0
    $lblStatus.Text = "$label 시작 - 파일 목록을 만드는 중..."
    Set-Scanning $true
    $timer.Start()
}

function Finish-Scan {
    $timer.Stop()
    $err = $null
    try { [void]$script:ps.EndInvoke($script:handle) } catch { $err = $_.Exception.Message }
    try { $script:ps.Runspace.Dispose(); $script:ps.Dispose() } catch { }
    $script:ps = $null; $script:handle = $null
    $pb.Style = 'Continuous'
    Set-Scanning $false
    $d = $script:doneInfo
    if ($d) {
        $pb.Value = 100
        $head = $(if ($d.Stopped) { "$($script:scanLabel) 중지됨" } else { "$($script:scanLabel) 완료" })
        $lblStatus.Text = "$head · 후보 $($script:hits.Count)개 · 전체 파일 $($d.Seen)개 중 검사 대상 $($d.Total)개 (내용 검사 $($d.Inspected) / 건너뜀 $($d.Skipped) / 읽기 오류 $($d.Errors))"
    } elseif ($err) {
        $lblStatus.Text = "검사 중 오류: $err"
    }
    Update-Counts
}

$timer = New-Object Windows.Forms.Timer
$timer.Interval = 200
$timer.Add_Tick({
    if (-not $script:q) { return }
    $rows = New-Object Collections.ArrayList
    $o = $null; $n = 0
    while ($n -lt 500 -and $script:q.TryDequeue([ref]$o)) {
        $n++
        switch ([string]$o.T) {
            'count' { $lblStatus.Text = "파일 목록을 만드는 중... 파일 $($o.Seen)개 확인 / 검사 대상 $($o.Targets)개" }
            'pct' {
                $pb.Style = 'Continuous'; $pb.Value = [Math]::Max(0, [Math]::Min(100, [int]$o.Pct))
                $lblStatus.Text = "내용 검사 중 $($o.Done) / $($o.Total) ($($o.Pct)%) · 후보 $($script:hits.Count)개"
            }
            'hit' {
                $o.Checked = $false
                [void]$script:hits.Add($o)
                if (Test-Visible $o) { [void]$rows.Add((New-Row $o)) }
            }
            'msg' { $txtDetail.AppendText("`r`n$($o.Text)") }
            'fatal' { $txtDetail.AppendText("`r`n검사 오류: $($o.Text)") }
            'done' { $script:doneInfo = $o }
        }
    }
    if ($rows.Count) {
        $script:rebuilding = $true
        $lv.BeginUpdate(); $lv.Items.AddRange([Windows.Forms.ListViewItem[]]$rows.ToArray()); $lv.EndUpdate()
        $script:rebuilding = $false
    }
    if ($script:handle -and $script:handle.IsCompleted -and $script:q.IsEmpty) { Finish-Scan }
})

# --- 영구 삭제 ---
function Get-ErrText($e) {
    $x = $e.Exception
    while ($x.InnerException) { $x = $x.InnerException }
    return $x.Message
}

function Remove-FileSecure([string]$path, [bool]$wipe) {
    $fi = New-Object IO.FileInfo($path)
    if ($fi.Attributes -band [IO.FileAttributes]::ReadOnly) {
        $fi.Attributes = [IO.FileAttributes]([int]$fi.Attributes -band (-bnot [int][IO.FileAttributes]::ReadOnly))
    }
    if ($wipe) {
        $fs = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $buf = New-Object byte[] 1048576
            $left = $fs.Length
            while ($left -gt 0) {
                $n = [int][Math]::Min([long]$buf.Length, $left)
                $fs.Write($buf, 0, $n); $left -= $n
            }
            $fs.Flush($true)
            $fs.SetLength(0)
        } finally { $fs.Dispose() }
        # 파일 이름도 남지 않도록 무작위 이름으로 바꾼 뒤 삭제
        $tmp = Join-Path $fi.DirectoryName ([Guid]::NewGuid().ToString('N') + '.tmp')
        [IO.File]::Move($path, $tmp)
        $path = $tmp
    }
    [IO.File]::Delete($path)
}

function Invoke-Delete {
    $targets = @($script:hits | Where-Object { $_.Checked })
    if ($targets.Count -eq 0) { $targets = @($lv.SelectedItems | ForEach-Object { $_.Tag }) }
    if ($targets.Count -eq 0) {
        [Windows.Forms.MessageBox]::Show('삭제할 파일 왼쪽 네모칸에 체크하세요.', '영구 삭제') | Out-Null
        return
    }
    $list = (@($targets | Select-Object -First 15 | ForEach-Object { $_.Path })) -join "`r`n"
    if ($targets.Count -gt 15) { $list += "`r`n... 외 $($targets.Count - 15)개" }
    $how = $(if ($chkWipe.Checked) { '내용을 0으로 덮어쓴 뒤 삭제합니다.' } else { '휴지통을 거치지 않고 바로 삭제합니다.' })
    $r = [Windows.Forms.MessageBox]::Show("$($targets.Count)개 파일을 영구 삭제합니다.`r`n$how 삭제한 파일은 되돌릴 수 없습니다.`r`n`r`n$list", '영구 삭제 확인', 'YesNo', 'Warning', 'Button2')
    if ($r -ne 'Yes') { return }
    if ($targets.Count -ge 10) {
        $t = [Microsoft.VisualBasic.Interaction]::InputBox("$($targets.Count)개 파일을 영구 삭제하려면 아래 칸에  삭제  라고 입력하세요.", '최종 확인', '')
        if ($t.Trim() -ne '삭제') { return }
    }

    $ok = 0; $fail = New-Object Collections.ArrayList; $i = 0
    $form.Cursor = 'WaitCursor'; Set-Scanning $true; $btnStop.Enabled = $false
    try {
        foreach ($h in $targets) {
            $i++
            $lblStatus.Text = "삭제 중 $i / $($targets.Count)"
            [Windows.Forms.Application]::DoEvents()
            try {
                $fi = New-Object IO.FileInfo($h.Path)
                if (-not $fi.Exists) { $h.Gone = $true; [void]$fail.Add("$($h.Path)`r`n   → 이미 없는 파일"); continue }
                if ($fi.LastWriteTimeUtc.Ticks -ne [long]$h.Ticks -or $fi.Length -ne [long]$h.Length) {
                    [void]$fail.Add("$($h.Path)`r`n   → 검사 후 내용이 바뀜 (다시 검사 후 확인하세요)"); continue
                }
                Remove-FileSecure $h.Path $chkWipe.Checked
                $h.Gone = $true; $ok++
            } catch {
                [void]$fail.Add("$($h.Path)`r`n   → $(Get-ErrText $_)")
            }
        }
    } finally {
        $form.Cursor = 'Default'; Set-Scanning $false
    }
    foreach ($h in @($script:hits | Where-Object { $_.Gone })) { $script:hits.Remove($h) }
    Rebuild-List
    $lblStatus.Text = "영구 삭제 완료 $ok개 / 실패 $($fail.Count)개"
    $msg = "영구 삭제 완료: $ok개"
    if ($fail.Count) {
        $msg += "`r`n삭제하지 못함: $($fail.Count)개 (파일이 열려 있으면 프로그램을 닫고 다시 시도하세요)`r`n`r`n" + ((@($fail) | Select-Object -First 20) -join "`r`n")
    }
    [Windows.Forms.MessageBox]::Show($msg, '영구 삭제 결과', 'OK', $(if ($fail.Count) { 'Warning' } else { 'Information' })) | Out-Null
}

# --- 버튼 연결 ---
$btnQuick.Add_Click({ Start-Scan (Get-QuickRoots) $false '빠른 검사' })
$btnFull.Add_Click({
    $drives = @([IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady } | ForEach-Object { $_.RootDirectory.FullName })
    $r = [Windows.Forms.MessageBox]::Show("다음 드라이브 전체를 검사합니다: $($drives -join ', ')`r`n(Windows·Program Files 등 시스템 폴더는 제외)`r`n`r`n사진이 많으면 오래 걸릴 수 있습니다. 시작할까요?", '전체 드라이브 검사', 'YesNo', 'Question')
    if ($r -eq 'Yes') { Start-Scan $drives $true '전체 드라이브 검사' }
})
$btnFolder.Add_Click({
    $dlg = New-Object Windows.Forms.FolderBrowserDialog
    $dlg.Description = '검사할 폴더를 고르세요 (하위 폴더 포함)'
    if ($dlg.ShowDialog($form) -eq 'OK') { Start-Scan @($dlg.SelectedPath) $false '폴더 검사' }
})
$btnStop.Add_Click({ if ($script:sync) { $script:sync.Stop = $true; $lblStatus.Text = '중지하는 중...' } })
$cmb.Add_SelectedIndexChanged({ Rebuild-List })
$lv.Add_ItemChecked({
    param($s, $e)
    if ($script:rebuilding) { return }
    $e.Item.Tag.Checked = $e.Item.Checked
    Update-Counts
})
$lv.Add_SelectedIndexChanged({
    if ($lv.SelectedItems.Count -eq 1) {
        $h = $lv.SelectedItems[0].Tag
        $txtDetail.Text = "$($h.Path)`r`n$($h.Detail)"
    }
})
$lv.Add_DoubleClick({ if ($lv.SelectedItems.Count -eq 1) { try { Start-Process -FilePath $lv.SelectedItems[0].Tag.Path } catch { } } })
$btnOpen.Add_Click({
    if ($lv.SelectedItems.Count -ge 1) {
        try { Start-Process -FilePath $lv.SelectedItems[0].Tag.Path } catch { [Windows.Forms.MessageBox]::Show((Get-ErrText $_), '파일 열기') | Out-Null }
    }
})
$btnShow.Add_Click({
    if ($lv.SelectedItems.Count -ge 1) { Start-Process explorer.exe -ArgumentList "/select,`"$($lv.SelectedItems[0].Tag.Path)`"" }
})
$btnCheckAll.Add_Click({
    foreach ($it in $lv.Items) { $it.Tag.Checked = $true }
    Rebuild-List
})
$btnUncheck.Add_Click({
    foreach ($h in $script:hits) { $h.Checked = $false }
    Rebuild-List
})
$btnDelete.Add_Click({ Invoke-Delete })
$form.Add_FormClosing({
    if ($script:sync) { $script:sync.Stop = $true }
    $timer.Stop()
    if ($script:ps) { try { [void]$script:ps.BeginStop($null, $null) } catch { } }
})

Update-Counts
[void]$form.ShowDialog()
[Environment]::Exit(0)
