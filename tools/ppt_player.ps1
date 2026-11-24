# 會議室主電腦：PowerPoint 放映器
# - 用 PowerPoint 放映原始簡報：動畫特效、音效、影片都會正常播放
# - 接上「會場連線」：主控台可換公所、上一步／下一步（下一步會照順序播動畫）、回選單
# - 簡報筆、鍵盤照常可用
param([int]$SelfTest = 0, [string]$Send = '')
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# 讓放映視窗一定跳到最前面：Windows 會擋背景程式搶前景，所以用 AttachThreadInput 借前景權限
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class NtFg {
  [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr p);
  [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
  [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool f);
  [DllImport("user32.dll")] static extern bool AllowSetForegroundWindow(int pid);
  [DllImport("user32.dll")] static extern void keybd_event(byte v, byte sc, uint f, UIntPtr e);
  [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint f);
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f, IntPtr l);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, System.Text.StringBuilder s, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, System.Text.StringBuilder s, int n);
  // 找 PowerPoint 全螢幕放映視窗（類別 screenClass）；標題含簡報檔名的優先
  public static IntPtr FindShow(string part) {
    IntPtr best = IntPtr.Zero, any = IntPtr.Zero;
    EnumWindows(delegate (IntPtr h, IntPtr l) {
      if (!IsWindowVisible(h)) return true;
      var c = new System.Text.StringBuilder(64); GetClassName(h, c, 64);
      if (c.ToString() != "screenClass") return true;
      if (any == IntPtr.Zero) any = h;
      var t = new System.Text.StringBuilder(512); GetWindowText(h, t, 512);
      if (!string.IsNullOrEmpty(part) && t.ToString().Contains(part)) { best = h; return false; }
      return true;
    }, IntPtr.Zero);
    return best != IntPtr.Zero ? best : any;
  }
  public static void Allow() { try { AllowSetForegroundWindow(-1); } catch {} }
  public static bool IsFront(IntPtr h) { return GetForegroundWindow() == h; }
  public static void Front(IntPtr h) {
    if (h == IntPtr.Zero) return;
    uint me = GetCurrentThreadId();
    uint fg = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero);
    uint tg = GetWindowThreadProcessId(h, IntPtr.Zero);
    if (fg != me) AttachThreadInput(me, fg, true);
    if (tg != me) AttachThreadInput(me, tg, true);
    // 放映視窗設為「最上層」，就算前景被別的程式佔住，畫面也一定蓋在最上面
    SetWindowPos(h, new IntPtr(-1), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040);
    // 按一下 Alt 解除 Windows 的前景鎖定，讓簡報筆、鍵盤的翻頁鍵送到放映視窗
    if (GetForegroundWindow() != h) { keybd_event(0x12, 0, 0, UIntPtr.Zero); keybd_event(0x12, 0, 2, UIntPtr.Zero); }
    ShowWindow(h, 5); BringWindowToTop(h); SetForegroundWindow(h);
    if (tg != me) AttachThreadInput(me, tg, false);
    if (fg != me) AttachThreadInput(me, fg, false);
  }
}
"@

$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
# 從別台電腦／隨身碟複製來的簡報會被標成「來自網際網路」，PowerPoint 會用受保護的檢視開啟；這裡先解除封鎖（不需系統管理員權限）
try { Get-ChildItem -LiteralPath $dir -File | Unblock-File -ErrorAction SilentlyContinue } catch {}
# 只留一個放映器：再開一次 .bat 時，先把舊的放映器關掉，避免兩個同時收主控台指令（翻頁重複、簡報開兩次）
if ($SelfTest -eq 0 -and -not $Send) {
  try {
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe' or Name='pwsh.exe'" |
      Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -like '*ppt_player.ps1*' -and $_.CommandLine -notlike '*-Send*' -and $_.CommandLine -notlike '*-SelfTest*' } |
      ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
  } catch {}
}
$TOWNS = @('鹿谷鄉','集集鎮','名間鄉','南投市','仁愛鄉','水里鄉','埔里鎮','竹山鎮','中寮鄉','魚池鄉','信義鄉','草屯鎮','國姓鄉')
$PREFIX = 'nt115v/ndnb7hg6ad1l/'
$BROKERS = @('wss://broker.emqx.io:8084/mqtt', 'wss://broker.hivemq.com:8884/mqtt', 'wss://test.mosquitto.org:8081/mqtt')
$utf8 = New-Object System.Text.UTF8Encoding($false)

# ---------- 各公所檔案（檔名開頭 01_～13_，依計時器順序） ----------
function Find-File([int]$i) {
  $pre = '{0:D2}_' -f ($i + 1)
  $all = Get-ChildItem -LiteralPath $dir -File | Where-Object { $_.Name.StartsWith($pre) -and $_.Extension -match '^\.(pptx|ppt|ppsx|pps|pptm|pdf)$' } | Sort-Object Name
  $ppt = $all | Where-Object { $_.Extension -ne '.pdf' } | Select-Object -First 1
  if ($ppt) { return $ppt }
  return ($all | Select-Object -First 1)
}

# ---------- PowerPoint ----------
$script:pp = $null; $script:pres = $null; $script:town = -1; $script:kind = ''
$script:view = 'menu'; $script:page = 0; $script:pages = 0; $script:lastSent = ''
function Close-Current {
  if ($script:pres) { try { $script:pres.Close() } catch {} ; $script:pres = $null }
  $script:kind = ''; $script:view = 'menu'; $script:page = 0; $script:pages = 0
}
function Get-View {
  if (-not $script:pres) { return $null }
  try { return $script:pres.SlideShowWindow.View } catch { return $null }
}
function Show-Hwnd {
  # PowerShell 讀不到 SlideShowWindow.HWND，改用視窗類別與標題找放映視窗
  $part = ''
  try { $part = [IO.Path]::GetFileNameWithoutExtension([string]$script:pres.Name) } catch {}
  try { return [NtFg]::FindShow($part) } catch { return [IntPtr]::Zero }
}
# 把全螢幕放映拉到最前面；剛開始放映的 3 秒內，Tick 會再檢查，沒在最前面就再拉一次
$script:focusUntil = $null
function Focus-Show {
  if (-not $script:pres) { return }
  try { $script:pres.SlideShowWindow.Activate() } catch {}
  try { [NtFg]::Front((Show-Hwnd)) } catch {}
}
function Begin-Show { if ($script:form) { try { $script:form.WindowState = 'Minimized' } catch {} }; [NtFg]::Allow() }
function Poll-Focus {
  if (-not $script:focusUntil) { return }
  if ((Get-Date) -gt $script:focusUntil) { $script:focusUntil = $null; return }
  if ((Show-Count) -gt 0) { $h = Show-Hwnd; if ($h -ne [IntPtr]::Zero -and -not [NtFg]::IsFront($h)) { Focus-Show } }
}
function Open-Town([int]$i) {
  if ($i -lt 0 -or $i -ge $TOWNS.Count) { return }
  $f = Find-File $i
  $script:town = $i; $script:resume = 1
  if (-not $f) { Close-Current; Set-Status ($TOWNS[$i] + '：資料夾裡沒有這家的簡報檔'); Push-State; return }
  Set-Status ('開啟中：' + $f.Name)
  try {
    if ($f.Extension -eq '.pdf') {
      Close-Current; Start-Process -FilePath $f.FullName
      $script:kind = 'pdf'; $script:view = 'play'
      Set-Status ($TOWNS[$i] + '：PDF 已開啟（主控台上／下頁會送翻頁鍵給 PDF 視窗）')
      Push-State; return
    }
    if (-not $script:pp) { $script:pp = New-Object -ComObject PowerPoint.Application }
    Close-Current
    # Open(檔名, 唯讀, 無標題, 有視窗)：-1 = 是，0 = 否
    $script:pres = $script:pp.Presentations.Open($f.FullName, -1, 0, -1)
    $ss = $script:pres.SlideShowSettings
    try { $ss.ShowType = 1; $ss.LoopUntilStopped = 0; $ss.AdvanceMode = 2 } catch {}
    Begin-Show
    [void]$ss.Run()
    $script:kind = 'ppt'; $script:view = 'play'; $script:pages = $script:pres.Slides.Count; $script:page = 1
    Focus-Show; $script:focusUntil = (Get-Date).AddSeconds(3)
    Set-Status ('放映中：' + $TOWNS[$i] + '（簡報筆、鍵盤、主控台都可翻頁）')
  } catch {
    Set-Status ('PowerPoint 開啟失敗，改直接開檔（請按 F5 放映）：' + $f.Name)
    try { Start-Process -FilePath $f.FullName } catch {}
    $script:kind = 'file'; $script:view = 'play'
  }
  Push-State
}
# 主控台「全螢幕放映」：放映中就把放映視窗拉到最前面；不小心按 Esc 退出時，從剛才那頁重新全螢幕放映
$script:resume = 1
function Start-Full {
  if ($script:kind -eq 'ppt' -and $script:pres) {
    $n = Show-Count
    if ($n -gt 0) { Begin-Show; Focus-Show; $script:focusUntil = (Get-Date).AddSeconds(3); Set-Status '已切回全螢幕放映'; Push-State; return }
    try {
      $total = [int]$script:pres.Slides.Count
      $start = [Math]::Max(1, [Math]::Min($total, [int]$script:resume))
      $ss = $script:pres.SlideShowSettings
      try { $ss.ShowType = 1; $ss.LoopUntilStopped = 0; $ss.AdvanceMode = 2 } catch {}
      $ss.RangeType = 2; $ss.StartingSlide = $start; $ss.EndingSlide = $total
      Begin-Show
      [void]$ss.Run()
      Focus-Show; $script:focusUntil = (Get-Date).AddSeconds(3)
      $script:view = 'play'; $script:page = $start; $script:pages = $total
      Set-Status ('全螢幕放映：' + $TOWNS[$script:town] + '（從第 ' + $start + ' 頁）')
      Push-State; return
    } catch {}
  }
  if ($script:town -ge 0) { Open-Town $script:town }
}
# 備案影片：PowerPoint 舊版（如 2007）播不出嵌入影片時，主控台按「播放本頁影片」
# 會用 Windows Media Player 全螢幕蓋在放映畫面上播放，播完自動關掉、回到原本那頁
# 檔案放在 影片備案\，檔名「NN_鄉鎮名_第PP頁.mp4」（NN = 公所序號、PP = 頁碼）
$script:vStart = $null; $script:vEnd = $null; $script:vProc = $null
function Find-Video([int]$town, [int]$page) {
  $vd = Join-Path $dir '影片備案'
  if (-not (Test-Path -LiteralPath $vd)) { return $null }
  $pre = '{0:D2}_' -f ($town + 1)
  return (Get-ChildItem -LiteralPath $vd -File | Where-Object { $_.Name.StartsWith($pre) -and $_.Name -match ('第0*' + $page + '頁') } | Select-Object -First 1)
}
function Video-Seconds($f) {
  try {
    $sh = New-Object -ComObject Shell.Application
    $it = $sh.Namespace($f.DirectoryName).ParseName($f.Name)
    $t = [string]$sh.Namespace($f.DirectoryName).GetDetailsOf($it, 27)
    $t = ($t -replace '[^0-9:]', '')
    if ($t) { return [TimeSpan]::Parse($t).TotalSeconds }
  } catch {}
  return 0
}
function Start-Video {
  if ($script:town -lt 0) { Set-Status '請先開啟公所簡報'; return }
  $pg = $(if ($script:page -ge 1) { $script:page } else { $script:resume })
  $f = Find-Video $script:town $pg
  if (-not $f) { Set-Status ($TOWNS[$script:town] + ' 第 ' + $pg + ' 頁沒有備案影片'); return }
  Stop-Video $false
  $wmp = @("$env:ProgramFiles\Windows Media Player\wmplayer.exe", "${env:ProgramFiles(x86)}\Windows Media Player\wmplayer.exe") | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
  $sec = Video-Seconds $f
  $script:vStart = Get-Date
  try {
    if ($wmp) { $script:vProc = Start-Process -FilePath $wmp -ArgumentList ('"' + $f.FullName + '" /fullscreen') -PassThru }
    else { $script:vProc = Start-Process -FilePath $f.FullName -PassThru }
  } catch { Set-Status ('影片開不起來：' + $f.Name); $script:vStart = $null; return }
  $script:vEnd = $(if ($sec -gt 0) { $script:vStart.AddSeconds($sec + 2) } else { $null })
  Set-Status ('播放影片：' + $f.Name + $(if ($sec -gt 0) { '（約 ' + [int]$sec + ' 秒，播完自動回簡報）' } else { '（播完請按主控台「停止影片」回簡報）' }))
}
function Stop-Video([bool]$manual) {
  if (-not $script:vStart) { return }
  $t0 = $script:vStart.AddSeconds(-2)
  try { if ($script:vProc -and -not $script:vProc.HasExited) { $script:vProc.Kill() } } catch {}
  try { Get-Process -Name wmplayer -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -ge $t0 } | ForEach-Object { try { $_.Kill() } catch {} } } catch {}
  $script:vStart = $null; $script:vEnd = $null; $script:vProc = $null
  if ($script:kind -eq 'ppt' -and $script:pres) { Focus-Show }
  Set-Status '影片結束，回到簡報'
}
function Poll-Video {
  if (-not $script:vStart) { return }
  if ($script:vEnd -and (Get-Date) -ge $script:vEnd) { Stop-Video $false; return }
  # 有人自己把播放器關掉：開始 3 秒後找不到播放器，就回到簡報
  if (((Get-Date) - $script:vStart).TotalSeconds -gt 3) {
    $alive = $false
    try { if ($script:vProc -and -not $script:vProc.HasExited) { $alive = $true } } catch {}
    if (-not $alive) { try { $t0 = $script:vStart.AddSeconds(-2); if (Get-Process -Name wmplayer -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -ge $t0 }) { $alive = $true } } catch {} }
    if (-not $alive -and $script:vEnd) { Stop-Video $false }
  }
}
# 只播聲音：主控台按某個聲音檔，就在背景播它的聲音（不開視窗、不蓋住簡報、不影響翻頁）
# 檔案放在 影片備案\（mp4、mp3、wav、m4a、wma 都可以），檔名開頭 NN_ 是公所序號
$script:wmp = $null; $script:aud = ''; $script:audT = $null
function Media-List {
  $vd = Join-Path $dir '影片備案'
  if (-not (Test-Path -LiteralPath $vd)) { return @() }
  return @(Get-ChildItem -LiteralPath $vd -File | Where-Object { $_.Extension -match '^\.(mp4|mp3|wav|m4a|wma|wmv)$' } | Sort-Object Name | ForEach-Object { $_.Name })
}
function Play-Audio([string]$name) {
  $vd = Join-Path $dir '影片備案'
  $f = Get-ChildItem -LiteralPath $vd -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $name } | Select-Object -First 1
  if (-not $f) { Set-Status ('找不到聲音檔：' + $name); return }
  try {
    if (-not $script:wmp) { $script:wmp = New-Object -ComObject WMPlayer.OCX }
    $script:wmp.settings.volume = 100
    $script:wmp.URL = $f.FullName
    $script:wmp.controls.play()
    $script:aud = $f.Name; $script:audT = Get-Date
    Set-Status ('播放聲音：' + $f.Name + '（簡報照常進行）')
  } catch { $script:aud = ''; Set-Status ('聲音播不出來：' + $f.Name + '，可直接雙擊這個檔案播放') }
  Push-State
  if ($script:kind -eq 'ppt' -and $script:pres) { if ((Show-Count) -gt 0) { Focus-Show } }
}
function Stop-Audio {
  try { if ($script:wmp) { $script:wmp.controls.stop() } } catch {}
  $script:aud = ''; $script:audT = $null; Set-Status '聲音已停止'; Push-State
}
function Poll-Audio {
  if (-not $script:aud -or -not $script:wmp) { return }
  if (((Get-Date) - $script:audT).TotalSeconds -lt 2) { return }
  try { $st = [int]$script:wmp.playState } catch { return }
  if ($st -eq 1 -or $st -eq 8 -or $st -eq 0) { $script:aud = ''; $script:audT = $null; Set-Status '聲音播完了'; Push-State }
}
function Do-Cmd($c) {
  if (-not $c) { return }
  $a = [string]$c.a
  switch ($a) {
    'town' { Open-Town ([int]$c.v) }
    'menu' { Close-Current; Set-Status '已回到選單，點公所即可放映'; Show-Form; Push-State }
    'show' { Start-Full }
    'video' { Start-Video }
    'videostop' { Stop-Video $true }
    'audio' { Play-Audio ([string]$c.v) }
    'audiostop' { Stop-Audio }
    default {
      if ($script:kind -eq 'pdf') {
        if ($a -eq 'next') { [System.Windows.Forms.SendKeys]::SendWait('{PGDN}') }
        elseif ($a -eq 'prev') { [System.Windows.Forms.SendKeys]::SendWait('{PGUP}') }
        return
      }
      $v = Get-View; if (-not $v) { return }
      try {
        if ($a -eq 'next') { $v.Next() }
        elseif ($a -eq 'prev') { $v.Previous() }
        elseif ($a -eq 'page') { $n = [int]$c.v; if ($n -ge 1 -and $n -le $script:pages) { $v.GotoSlide($n) } }
      } catch {}
      Poll-Show
    }
  }
}
# 檢查放映狀態：PowerPoint 播影片或忙碌時 COM 會暫時沒回應，不可以因此判定放映結束（舊版會因此跳回清單）
$script:lastPoll = [datetime]::MinValue
function Show-Count {
  try { return [int]$script:pp.SlideShowWindows.Count } catch { return -1 }  # -1 = 暫時問不到
}
function Poll-Show {
  if ($script:kind -ne 'ppt') { return }
  if (((Get-Date) - $script:lastPoll).TotalMilliseconds -lt 500) { return }
  $script:lastPoll = Get-Date
  $n = Show-Count
  if ($n -eq 0) {
    if ($script:view -eq 'play') { if ($script:page -ge 1) { $script:resume = $script:page }; $script:view = 'menu'; $script:page = 0; Set-Status '放映已結束（按 Esc 或播完），點公所可再播放'; Push-State; Show-Form }
    return
  }
  if ($n -lt 0) { return }
  $v = Get-View; if (-not $v) { return }
  try {
    $p = [int]$v.CurrentShowPosition
    $changed = $false
    if ($script:view -ne 'play') { $script:view = 'play'; $changed = $true }
    if ($p -ge 1 -and $p -ne $script:page) { $script:page = $p; $changed = $true }
    if ($changed) { Push-State }
  } catch {}
}

# ---------- 會場連線（MQTT over WebSocket） ----------
function B-Str([string]$s) { $b = $utf8.GetBytes($s); $o = New-Object byte[] ($b.Length + 2); $o[0] = [byte]($b.Length -shr 8); $o[1] = [byte]($b.Length -band 255); [Array]::Copy($b, 0, $o, 2, $b.Length); return ,$o }
function B-Cat { $l = New-Object System.Collections.Generic.List[byte]; foreach ($x in $args) { if ($x -is [byte[]]) { $l.AddRange($x) } else { foreach ($y in $x) { $l.Add([byte]$y) } } }; return ,$l.ToArray() }
function B-Packet([int]$type, [byte[]]$body) {
  $n = $body.Length; $rl = New-Object System.Collections.Generic.List[byte]
  do { $d = $n % 128; $n = [math]::Floor($n / 128); if ($n -gt 0) { $d = $d -bor 128 }; $rl.Add([byte]$d) } while ($n -gt 0)
  return (B-Cat ([byte[]]@([byte]$type)) $rl.ToArray() $body)
}
function Rid { return ([guid]::NewGuid().ToString('N').Substring(0, 12)) }
function Now-Ms { return [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() }
function State-Json([bool]$online) {
  $o = [ordered]@{ online = $online }
  if ($online) {
    $o.view = $script:view; $o.town = $script:town; $o.name = $(if ($script:town -ge 0) { $TOWNS[$script:town] } else { '' })
    $o.page = $script:page; $o.pages = $script:pages; $o.engine = 'ppt'; $o.pid = $script:PLAYERID
    $o.media = [string[]](Media-List); $o.aud = $script:aud
  }
  $o._id = Rid; $o._ts = Now-Ms
  return ($o | ConvertTo-Json -Compress)
}
$script:PLAYERID = 'ppt-' + ([guid]::NewGuid().ToString('N').Substring(0, 6))
$conns = @()
foreach ($u in $BROKERS) { $conns += @{ url = $u; ws = $null; ok = $false; ct = $null; rt = $null; raw = (New-Object byte[] 65536); buf = (New-Object System.Collections.Generic.List[byte]); next = [datetime]::MinValue; retry = 0; started = [datetime]::MinValue; ping = [datetime]::MinValue; pid = 1 } }
$seen = @{}
function C-Send($c, [byte[]]$data) {
  try { $t = $c.ws.SendAsync((New-Object System.ArraySegment[byte] -ArgumentList (, $data)), [System.Net.WebSockets.WebSocketMessageType]::Binary, $true, [Threading.CancellationToken]::None); [void]$t.Wait(3000); return $true } catch { return $false }
}
function C-Drop($c) {
  try { if ($c.ws) { $c.ws.Abort(); $c.ws.Dispose() } } catch {}
  $c.ws = $null; $c.ct = $null; $c.rt = $null; $c.buf.Clear()
  if ($c.ok) { $c.ok = $false; Update-Net }
  $c.retry = [math]::Min($c.retry + 1, 5); $c.next = (Get-Date).AddSeconds([math]::Min(30, [math]::Pow(2, $c.retry)))
}
function C-Open($c) {
  try {
    $c.ws = New-Object System.Net.WebSockets.ClientWebSocket
    $c.ws.Options.AddSubProtocol('mqtt')
    try { $c.ws.Options.Proxy = [System.Net.WebRequest]::DefaultWebProxy } catch {}
    $c.ct = $c.ws.ConnectAsync([Uri]$c.url, [Threading.CancellationToken]::None); $c.started = Get-Date
  } catch { C-Drop $c }
}
function C-Publish($c, [string]$topic, [string]$json, [bool]$retain) {
  if (-not $c.ok) { return }
  $h = 0x30; if ($retain) { $h = 0x31 }
  [void](C-Send $c (B-Packet $h (B-Cat (B-Str ($PREFIX + $topic)) $utf8.GetBytes($json))))
}
function Push-State {
  if ($Send) { return }
  $j = State-Json $true
  foreach ($c in $conns) { C-Publish $c 'player/state' $j $true }
  if ($SelfTest) { Write-Host ('STATE ' + $j) }
}
function C-Parse($c) {
  while ($c.buf.Count -ge 2) {
    $mul = 1; $len = 0; $i = 1
    do { if ($i -ge $c.buf.Count) { return }; $byte = $c.buf[$i]; $i++; $len += ($byte -band 127) * $mul; $mul *= 128 } while ($byte -band 128)
    if ($c.buf.Count -lt $i + $len) { return }
    $hdr = $c.buf[0]; $type = $hdr -shr 4; $body = $c.buf.GetRange($i, $len).ToArray(); $c.buf.RemoveRange(0, $i + $len)
    if ($type -eq 2) {
      if ($body.Length -ge 2 -and $body[1] -eq 0) {
        $c.ok = $true; $c.retry = 0; $c.ping = Get-Date; Update-Net
        $id = $c.pid; $c.pid++
        [void](C-Send $c (B-Packet 0x82 (B-Cat ([byte[]]@([byte]($id -shr 8), [byte]($id -band 255))) (B-Str ($PREFIX + 'player/cmd')) ([byte[]]@(0)))))
        Push-State
      } else { C-Drop $c; return }
    } elseif ($type -eq 3) {
      $tl = ($body[0] -shl 8) -bor $body[1]; $topic = $utf8.GetString($body, 2, $tl); $p = 2 + $tl; if ($hdr -band 6) { $p += 2 }
      $txt = $utf8.GetString($body, $p, $body.Length - $p)
      if ($topic -eq ($PREFIX + 'player/cmd')) {
        try { $o = $txt | ConvertFrom-Json } catch { $o = $null }
        if ($o) {
          $k = [string]$o._id
          if ($k -and $seen.ContainsKey($k)) { continue }
          if ($k) { $seen[$k] = Get-Date; if ($seen.Count -gt 300) { $seen.Clear(); $seen[$k] = Get-Date } }
          # 剛連上時收到的舊指令（保留訊息）不執行
          if ($o._ts -and ((Now-Ms) - [int64]$o._ts) -gt 15000) { continue }
          if ($SelfTest) { Write-Host ('CMD ' + $txt) }
          if (-not $Send) { Do-Cmd $o }
        }
      }
    }
  }
}
function Pump {
  foreach ($c in $conns) {
    try {
      if (-not $c.ws) { if ((Get-Date) -ge $c.next) { C-Open $c }; continue }
      if ($c.ct) {
        if (-not $c.ct.IsCompleted) { if (((Get-Date) - $c.started).TotalSeconds -gt 10) { C-Drop $c }; continue }
        if ($c.ct.IsFaulted -or $c.ws.State -ne 'Open') { C-Drop $c; continue }
        $c.ct = $null
        $will = State-Json $false
        if ($Send) {
          $vh = B-Cat (B-Str 'MQTT') ([byte[]]@(4, 0x02, 0, 30)); $pay = B-Str ('nt115-send-' + (Rid))
        } else {
          $vh = B-Cat (B-Str 'MQTT') ([byte[]]@(4, (0x02 -bor 0x04 -bor 0x20), 0, 30))
          $pay = B-Cat (B-Str ('nt115-ppt-' + (Rid))) (B-Str ($PREFIX + 'player/state')) (B-Str $will)
        }
        if (-not (C-Send $c (B-Packet 0x10 (B-Cat $vh $pay)))) { C-Drop $c; continue }
      }
      if ($c.ws.State -ne 'Open') { C-Drop $c; continue }
      for ($k = 0; $k -lt 20; $k++) {
        if (-not $c.rt) { $c.rt = $c.ws.ReceiveAsync((New-Object System.ArraySegment[byte] -ArgumentList (, $c.raw)), [Threading.CancellationToken]::None) }
        if (-not $c.rt.IsCompleted) { break }
        if ($c.rt.IsFaulted -or $c.rt.IsCanceled) { C-Drop $c; break }
        $r = $c.rt.Result; $c.rt = $null
        if ($r.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) { C-Drop $c; break }
        for ($j = 0; $j -lt $r.Count; $j++) { $c.buf.Add($c.raw[$j]) }
        C-Parse $c
        if (-not $c.ws) { break }
      }
      if ($c.ok -and ((Get-Date) - $c.ping).TotalSeconds -ge 20) { [void](C-Send $c ([byte[]]@(0xC0, 0))); $c.ping = Get-Date }
    } catch { C-Drop $c }
  }
}
$script:beat = Get-Date
function Tick {
  Pump; Poll-Show; Poll-Focus; Poll-Video; Poll-Audio
  if (((Get-Date) - $script:beat).TotalSeconds -ge 20) { $script:beat = Get-Date; Push-State }
}

# ---------- 視窗 ----------
$script:form = $null; $script:status = $null; $script:net = $null
function Set-Status([string]$t) { if ($script:status) { $script:status.Text = $t }; if ($SelfTest) { Write-Host ('STATUS ' + $t) } }
function Update-Net {
  $n = @($conns | Where-Object { $_.ok }).Count
  $t = $(if ($n -gt 0) { "● 已連上會場連線（$n 條），主控台可遙控" } else { '○ 未連上會場連線：本機照常播放（點公所、簡報筆翻頁）' })
  if ($script:net) { $script:net.Text = $t; $script:net.ForeColor = $(if ($n -gt 0) { [System.Drawing.Color]::FromArgb(30, 140, 90) } else { [System.Drawing.Color]::FromArgb(190, 60, 50) }) }
  if ($SelfTest) { Write-Host ('NET ' + $n) }
}
function Show-Form { if ($script:form) { try { $script:form.WindowState = 'Normal'; $script:form.Activate() } catch {} } }

if ($Send) {
  # 測試用：送一個指令給放映器（例如 -Send '{"a":"next"}'），送出即結束
  $SelfTest = 1; $end = (Get-Date).AddSeconds(15)
  while ((Get-Date) -lt $end -and -not @($conns | Where-Object { $_.ok }).Count) { Pump; Start-Sleep -Milliseconds 100 }
  $end = (Get-Date).AddSeconds(2); while ((Get-Date) -lt $end) { Pump; Start-Sleep -Milliseconds 100 }
  $o = $Send | ConvertFrom-Json; $h = [ordered]@{}; foreach ($pr in $o.PSObject.Properties) { $h[$pr.Name] = $pr.Value }; $h._id = Rid; $h._ts = Now-Ms
  $j = $h | ConvertTo-Json -Compress; $n = 0
  foreach ($c in $conns) { if ($c.ok) { $hb = 0x30; [void](C-Send $c (B-Packet $hb (B-Cat (B-Str ($PREFIX + 'player/cmd')) $utf8.GetBytes($j)))); $n++ } }
  Write-Host ("SENT $n $j"); exit 0
}
if ($SelfTest -gt 0) {
  # 測試模式：不開視窗，只跑連線 N 秒，印出收到的指令
  $end = (Get-Date).AddSeconds($SelfTest)
  while ((Get-Date) -lt $end) { Tick; Start-Sleep -Milliseconds 100 }
  foreach ($c in $conns) { if ($c.ok) { C-Publish $c 'player/state' (State-Json $false) $true } }
  Close-Current; try { if ($script:pp) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:pp) } } catch {}
  exit 0
}

$form = New-Object System.Windows.Forms.Form
$script:form = $form
$form.Text = 'PowerPoint 放映器（115年鄉鎮市考核）'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(600, 820)
$form.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 12)

$net = New-Object System.Windows.Forms.Label
$net.Dock = 'Top'; $net.Height = 40; $net.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 0)
$script:net = $net

$panel = New-Object System.Windows.Forms.FlowLayoutPanel
$panel.Dock = 'Fill'; $panel.FlowDirection = 'TopDown'; $panel.WrapContents = $false; $panel.AutoScroll = $true
$panel.Padding = New-Object System.Windows.Forms.Padding(12)

$status = New-Object System.Windows.Forms.Label
$status.Dock = 'Bottom'; $status.Height = 56; $status.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$status.Text = '點公所即用 PowerPoint 放映（動畫、音效、影片都會播放）。放映中按 Esc 回到這裡。'
$script:status = $status

$form.Controls.Add($panel); $form.Controls.Add($status); $form.Controls.Add($net)

for ($i = 0; $i -lt $TOWNS.Count; $i++) {
  $f = Find-File $i
  $b = New-Object System.Windows.Forms.Button
  $b.Width = 540; $b.Height = 46; $b.TextAlign = 'MiddleLeft'; $b.Tag = $i
  $b.Text = ('{0:D2}  {1}' -f ($i + 1), $TOWNS[$i]) + $(if ($f) { $(if ($f.Extension -eq '.pdf') { '　（PDF）' } else { '' }) } else { '　（尚未提供）' })
  if (-not $f) { $b.ForeColor = [System.Drawing.Color]::Gray }
  $b.Add_Click({ param($s, $e) Open-Town ([int]$s.Tag) })
  $panel.Controls.Add($b)
}
$endBtn = New-Object System.Windows.Forms.Button
$endBtn.Text = '結束目前放映並關閉簡報'; $endBtn.Width = 540; $endBtn.Height = 46
$endBtn.Add_Click({ Close-Current; Set-Status '已關閉目前簡報'; Push-State })
$panel.Controls.Add($endBtn)
$quitBtn = New-Object System.Windows.Forms.Button
$quitBtn.Text = '完全關閉放映器（會議結束時按）'; $quitBtn.Width = 540; $quitBtn.Height = 46
$quitBtn.ForeColor = [System.Drawing.Color]::FromArgb(190, 60, 50)
$quitBtn.Add_Click({ $form.Close() })
$panel.Controls.Add($quitBtn)

Update-Net
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 150
$timer.Add_Tick({ Tick })
$timer.Start()
$form.Add_FormClosing({
  $timer.Stop()
  foreach ($c in $conns) { if ($c.ok) { C-Publish $c 'player/state' (State-Json $false) $true } }
  Close-Current
  # PowerPoint 沒有其他開著的檔案就一起關掉（有別的檔案開著就不動它）
  try { if ($script:pp -and [int]$script:pp.Presentations.Count -eq 0) { $script:pp.Quit() } } catch {}
  try { if ($script:pp) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:pp) } } catch {}
  $script:pp = $null
})
[void]$form.ShowDialog()
# 視窗關掉後確實結束整個程式，不留在背景
foreach ($c in $conns) { try { $c.ws.Dispose() } catch {} }
[Environment]::Exit(0)
