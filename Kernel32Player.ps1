# 需要在 STA 模式下运行（Windows PowerShell 默认是 STA）
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Host "请在 STA 模式下运行此脚本，例如：powershell.exe -STA -File .\Kernel32Player.ps1" -ForegroundColor Yellow
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ============================
# Kernel32 Beep
# ============================
$beepCs = @'
using System;
using System.Runtime.InteropServices;
using System.Threading;

public static class Kernel32PlayerBeep
{
    [DllImport("kernel32.dll")]
    public static extern bool Beep(uint dwFreq, uint dwDuration);

    /// <summary>
    /// 按下播放一次超长Beep（在后台线程中，避免卡 UI）
    /// </summary>
    public static void StartNote(string id, uint freq)
    {
        Thread t = new Thread(() =>
        {
            Beep(freq, 4294967294);
        });
        t.IsBackground = true;
        t.Start();
    }

    /// <summary>
    /// 松开播放一次1Hz，0ms的Beep
    /// </summary>
    public static void StopNote(string id)
    {
        Thread t = new Thread(() =>
        {
            Beep(1, 0);
        });
        t.IsBackground = true;
        t.Start();
    }

    /// <summary>
    /// 测试用
    /// </summary>
    public static void TestBeep(uint freq, uint duration)
    {
        Beep(freq, duration);
    }

    /// <summary>
    /// MIDI 播放用：在后台线程播放指定时长 Beep
    /// </summary>
    public static void PlayNote(uint freq, uint durationMs)
    {
        Thread t = new Thread(() => { Beep(freq, durationMs); });
        t.IsBackground = true;
        t.Start();
    }
}
'@

try {
    Add-Type -TypeDefinition $beepCs -Language CSharp -ErrorAction Stop | Out-Null
} catch {
    Write-Host "C# Beep 类型编译失败：" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host $_.Exception.InnerException.Message -ForegroundColor Red
    }
    return
}

# ============================
# MIDI Parser (C# Add-Type)
# ============================
$midiParserCs = @'
using System;
using System.Collections.Generic;
using System.IO;

public class MidiNoteEvent
{
    public long StartMs { get; set; }
    public long DurationMs { get; set; }
    public int MidiNote { get; set; }
}

public static class MidiParser
{
    public static List<MidiNoteEvent> Parse(string filePath)
    {
        var notes = new List<MidiNoteEvent>();
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var br = new BinaryReader(fs))
        {
            var header = new string(br.ReadChars(4));
            if (header != "MThd")
                throw new Exception("Invalid MIDI header");
            br.ReadBytes(4);
            int format = ReadBigEndianShort(br);
            int numTracks = ReadBigEndianShort(br);
            int division = ReadBigEndianShort(br);
            int ticksPerQuarter = (division & 0x7FFF);
            double currentTempo = 500000.0;

            var pendingNotes = new Dictionary<int, long>();

            for (int trackIdx = 0; trackIdx < numTracks; trackIdx++)
            {
                var chunkId = new string(br.ReadChars(4));
                if (chunkId != "MTrk")
                    throw new Exception("Expected MTrk chunk");
                int chunkLen = ReadBigEndianInt(br);
                long chunkEnd = fs.Position + chunkLen;
                byte runningStatus = 0;
                long tick = 0;
                double micros = 0;

                while (fs.Position < chunkEnd)
                {
                    int delta = ReadVariableLength(br);
                    tick += delta;
                    micros += delta * (currentTempo / ticksPerQuarter);

                    byte b = br.ReadByte();
                    byte status;
                    if ((b & 0x80) != 0)
                    {
                        status = b;
                        runningStatus = b;
                    }
                    else
                    {
                        status = runningStatus;
                        fs.Seek(-1, SeekOrigin.Current);
                    }

                    if (status >= 0xF0)
                    {
                        if (status == 0xFF)
                        {
                            byte metaType = br.ReadByte();
                            int metaLen = ReadVariableLength(br);
                            byte[] metaData = br.ReadBytes(metaLen);
                            if (metaType == 0x51 && metaLen == 3)
                            {
                                int usPerQuarter = (metaData[0] << 16) | (metaData[1] << 8) | metaData[2];
                                currentTempo = usPerQuarter;
                            }
                        }
                        else if (status == 0xF0 || status == 0xF7)
                        {
                            int len = ReadVariableLength(br);
                            br.ReadBytes(len);
                        }
                        continue;
                    }

                    int cmd = status & 0xF0;

                    if (cmd == 0x80 || cmd == 0x90)
                    {
                        int note = br.ReadByte();
                        int vel = br.ReadByte();
                        bool isOff = (cmd == 0x80 || vel == 0);
                        long startMicros = (long)micros;
                        if (isOff && pendingNotes.ContainsKey(note))
                        {
                            long start = pendingNotes[note];
                            pendingNotes.Remove(note);
                            notes.Add(new MidiNoteEvent
                            {
                                StartMs = start / 1000,
                                DurationMs = (startMicros - start) / 1000,
                                MidiNote = note
                            });
                        }
                        else if (!isOff)
                        {
                            pendingNotes[note] = startMicros;
                        }
                    }
                    else if (cmd == 0xA0 || cmd == 0xB0 || cmd == 0xE0)
                    {
                        br.ReadByte();
                        br.ReadByte();
                    }
                    else if (cmd == 0xC0 || cmd == 0xD0)
                    {
                        br.ReadByte();
                    }
                }
            }
        }
        notes.Sort((a, b) => a.StartMs.CompareTo(b.StartMs));
        return notes;
    }

    static int ReadBigEndianShort(BinaryReader br)
    {
        byte[] b = br.ReadBytes(2);
        return (b[0] << 8) | b[1];
    }

    static int ReadBigEndianInt(BinaryReader br)
    {
        byte[] b = br.ReadBytes(4);
        return (b[0] << 24) | (b[1] << 16) | (b[2] << 8) | b[3];
    }

    static int ReadVariableLength(BinaryReader br)
    {
        int v = 0;
        for (int i = 0; i < 4; i++)
        {
            byte b = br.ReadByte();
            v = (v << 7) | (b & 0x7F);
            if ((b & 0x80) == 0) break;
        }
        return v;
    }
}
'@

try {
    Add-Type -TypeDefinition $midiParserCs -Language CSharp -ErrorAction Stop | Out-Null
} catch {
    Write-Host "C# MIDI Parser 编译失败：" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

# ============================
# XAML UI 定义
# ============================
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Kernel32Player"
        Width="1100"
        Height="360"
        Background="Black"
        ResizeMode="CanMinimize"
        WindowStartupLocation="CenterScreen"
        AllowsTransparency="False"
        FontFamily="Consolas">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Margin="0,0,0,10" Padding="10"
                Background="#111111"
                CornerRadius="4">
            <TextBlock Text="Kernel32Player"
                       HorizontalAlignment="Center"
                       VerticalAlignment="Center"
                       Foreground="Gray"
                       FontSize="32"
                       FontWeight="Bold"
                       FontStyle="Italic"
                       TextOptions.TextFormattingMode="Display"
                       TextOptions.TextRenderingMode="Aliased">
                <TextBlock.Effect>
                    <DropShadowEffect Color="DarkGray" BlurRadius="8" ShadowDepth="2" Opacity="0.8"/>
                </TextBlock.Effect>
            </TextBlock>
        </Border>

        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
            <Button x:Name="TestButton"
                    Content="测试"
                    Width="80"
                    Height="30"
                    Margin="0,0,10,0"
                    Background="#222222"
                    Foreground="Lime"
                    BorderBrush="Gray"/>
            <Button x:Name="MidiOpenButton"
                    Content="打开 MIDI"
                    Width="100"
                    Height="30"
                    Margin="0,0,10,0"
                    Background="#222222"
                    Foreground="Cyan"
                    BorderBrush="Gray"/>
            <Button x:Name="MidiStopButton"
                    Content="停止"
                    Width="80"
                    Height="30"
                    Margin="0,0,10,0"
                    Background="#222222"
                    Foreground="Orange"
                    BorderBrush="Gray"
                    IsEnabled="False"/>
            <TextBlock Text="使用鼠标点击琴键或使用键盘zsxdcvgbhnjm,l.;/q2w3e4rt6y7ui9o0p-[]演奏，c3-c6。"
                       VerticalAlignment="Center"
                       Foreground="LightGray"/>
        </StackPanel>

        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="300"/>
            </Grid.ColumnDefinitions>

            <ScrollViewer x:Name="PianoScroll"
                          Grid.Column="0"
                          HorizontalScrollBarVisibility="Visible"
                          VerticalScrollBarVisibility="Disabled"
                          Background="Black">
                <Canvas x:Name="PianoCanvas"
                        Height="180"
                        Background="Black"/>
            </ScrollViewer>

            <Border Grid.Column="1"
                    Margin="10,0,0,0"
                    Background="#111111"
                    BorderBrush="#333333"
                    BorderThickness="1"
                    CornerRadius="4">
                <TextBox x:Name="LogBox"
                         Background="#111111"
                         Foreground="Lime"
                         FontSize="12"
                         IsReadOnly="True"
                         TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Disabled"
                         BorderThickness="0"
                         Padding="5"/>
            </Border>
        </Grid>
    </Grid>
</Window>
"@

$xmlReader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window    = [Windows.Markup.XamlReader]::Load($xmlReader)

$pianoCanvas    = $window.FindName('PianoCanvas')
$logBox         = $window.FindName('LogBox')
$testButton     = $window.FindName('TestButton')
$pianoScroll    = $window.FindName('PianoScroll')
$midiOpenButton = $window.FindName('MidiOpenButton')
$midiStopButton = $window.FindName('MidiStopButton')

$script:midiStopRequested = $false
$script:midiPlaybackJob   = $null
$script:mainWindow       = $window
$script:mainMidiStopBtn  = $midiStopButton
$script:mainMidiOpenBtn  = $midiOpenButton
$script:midiStopRequested = $false
$script:midiPlaybackJob   = $null
$script:mainWindow       = $window
$script:mainMidiStopBtn  = $midiStopButton
$script:mainMidiOpenBtn  = $midiOpenButton
$script:midiNotes        = @()
$script:midiPlayIndex    = 0
$script:midiStartTick    = 0
$script:midiPlayTimer    = $null

# ============================
# 日志函数
# ============================
function Write-Log {
    param(
        [string]$Message
    )
    if (-not $logBox) { return }
    $ts = (Get-Date).ToString("HH:mm:ss")
    $logBox.AppendText("[$ts] $Message`r`n")
    $logBox.ScrollToEnd()
}

# ============================
# 键盘映射与音高计算
# ============================
$keyCharsString = "zsxdcvgbhnjm,l.;/q2w3e4rt6y7ui9o0p-[]"
$keyChars = $keyCharsString.ToCharArray() | Where-Object { $_ -ne [char]0 }

$startMidi = 48
$noteNames = "C","C#","D","D#","E","F","F#","G","G#","A","A#","B"

$script:keyMap      = @{}
$script:keyButtons  = @{}
$script:pressedKeys = @{}

$whiteKeyWidth  = 40
$whiteKeyHeight = 180
$blackKeyWidth  = 26
$blackKeyHeight = 110

$whiteIndex = 0

for ($i = 0; $i -lt $keyChars.Count; $i++) {
    $ch   = [string]$keyChars[$i]
    $midi = $startMidi + $i
    $index = $midi % 12
    $noteName = $noteNames[$index]

    $octave = [int][math]::Floor($midi / 12) - 1

    $n    = $midi - 69
    $freq = 440.0 * [math]::Pow(2.0, $n / 12.0)

    if ($freq -lt 37)    { $freq = 37 }
    if ($freq -gt 32767) { $freq = 32767 }

    $isBlack = $noteName.Contains("#")

    if (-not $isBlack) {
        $x = $whiteIndex * $whiteKeyWidth
        $whiteIndex++
    } else {
        $x = $whiteIndex * $whiteKeyWidth - ($blackKeyWidth / 2.0)
    }

    if (-not $isBlack) {
        $btn = New-Object System.Windows.Controls.Button
        $btn.Width       = $whiteKeyWidth
        $btn.Height      = $whiteKeyHeight
        $btn.Background  = [System.Windows.Media.Brushes]::White
        $btn.BorderBrush = [System.Windows.Media.Brushes]::Gray
        $btn.Foreground  = [System.Windows.Media.Brushes]::Black
        $btn.Padding     = New-Object System.Windows.Thickness(0)
        $btn.FontSize    = 16
        $btn.FontWeight  = "Bold"
        $btn.VerticalContentAlignment   = 'Bottom'
        $btn.HorizontalContentAlignment = 'Center'

        if ($noteName -eq "C" -and $octave -ge 3 -and $octave -le 6) {
            $btn.Content = "C$octave"
        } else {
            $btn.Content = ""
        }

        [System.Windows.Controls.Canvas]::SetLeft($btn, $x)
        [System.Windows.Controls.Canvas]::SetTop($btn, 0)
        [System.Windows.Controls.Panel]::SetZIndex($btn, 0)
    }
    else {
        $btn = New-Object System.Windows.Controls.Button
        $btn.Width       = $blackKeyWidth
        $btn.Height      = $blackKeyHeight
        $btn.Background  = [System.Windows.Media.Brushes]::Black
        $btn.BorderBrush = [System.Windows.Media.Brushes]::DarkGray
        $btn.Foreground  = [System.Windows.Media.Brushes]::White
        $btn.Padding     = New-Object System.Windows.Thickness(0)
        $btn.FontSize    = 8

        [System.Windows.Controls.Canvas]::SetLeft($btn, $x)
        [System.Windows.Controls.Canvas]::SetTop($btn, 0)
        [System.Windows.Controls.Panel]::SetZIndex($btn, 1)
    }

    $btn.Tag = $ch

    [void]$pianoCanvas.Children.Add($btn)

    $info = [pscustomobject]@{
        Char      = $ch
        Midi      = $midi
        NoteName  = $noteName
        Octave    = $octave
        Frequency = $freq
        IsBlack   = $isBlack
    }
    $script:keyMap[$ch]     = $info
    $script:keyButtons[$ch] = $btn
}

$pianoCanvas.Width = $whiteIndex * $whiteKeyWidth

# ============================
# 键盘视觉动画
# ============================
function Get-KeyBaseBrush {
    param($info)
    if ($info.IsBlack) {
        return [System.Windows.Media.Brushes]::Black
    } else {
        return [System.Windows.Media.Brushes]::White
    }
}

function Invoke-KeyVisualDown {
    param(
        [System.Windows.Controls.Button]$Button,
        $Info
    )
    $Button.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::Gold)
    $Button.BorderBrush = [System.Windows.Media.Brushes]::Orange

    $anim = New-Object System.Windows.Media.Animation.DoubleAnimation
    $anim.From = 1.0
    $anim.To   = 0.7
    $anim.Duration = [System.Windows.Duration]::op_Implicit([System.TimeSpan]::FromMilliseconds(80))
    $anim.AutoReverse = $true

    $Button.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $anim)
}

function Invoke-KeyVisualUp {
    param(
        [System.Windows.Controls.Button]$Button,
        $Info
    )
    $Button.Background  = Get-KeyBaseBrush $Info
    $Button.BorderBrush = [System.Windows.Media.Brushes]::Gray
    $Button.Opacity     = 1.0
}

# ============================
# 音符触发逻辑
# ============================
function Invoke-NoteDown {
    param(
        [string]$Char
    )
    if (-not $script:keyMap.ContainsKey($Char)) { return }

    $toRelease = @($script:pressedKeys.Keys | Where-Object { $_ -ne $Char })
    foreach ($c in $toRelease) {
        Invoke-NoteUp -Char $c
    }

    if ($script:pressedKeys.ContainsKey($Char)) {
        return
    }

    $info = $script:keyMap[$Char]
    $script:pressedKeys[$Char] = $true

    if ($script:keyButtons.ContainsKey($Char)) {
        $btn = $script:keyButtons[$Char]
        Invoke-KeyVisualDown -Button $btn -Info $info
    }

    $freqUint = [uint32][math]::Round($info.Frequency)
    [Kernel32PlayerBeep]::StartNote($Char, $freqUint)

    Write-Log "按下: '$Char' -> $($info.NoteName)$($info.Octave) ($([math]::Round($info.Frequency)) Hz)"
}

function Invoke-NoteUp {
    param(
        [string]$Char
    )
    if (-not $script:pressedKeys.ContainsKey($Char)) {
        return
    }

    $script:pressedKeys.Remove($Char)

    if ($script:keyMap.ContainsKey($Char)) {
        $info = $script:keyMap[$Char]
    } else {
        $info = $null
    }

    if ($script:keyButtons.ContainsKey($Char)) {
        $btn = $script:keyButtons[$Char]
        if ($info) {
            Invoke-KeyVisualUp -Button $btn -Info $info
        }
    }

    [Kernel32PlayerBeep]::StopNote($Char)

    if ($info) {
        Write-Log "松开: '$Char' -> $($info.NoteName)$($info.Octave)"
    } else {
        Write-Log "松开: '$Char'"
    }
}

# ============================
# 鼠标事件绑定
# ============================
foreach ($entry in $script:keyButtons.GetEnumerator()) {
    $btn  = $entry.Value
    $char = [string]$entry.Key

    $btn.Add_PreviewMouseLeftButtonDown({
        param($sender, $e)
        $c = [string]$sender.Tag
        Invoke-NoteDown -Char $c
        $e.Handled = $true
    })

    $btn.Add_PreviewMouseLeftButtonUp({
        param($sender, $e)
        $c = [string]$sender.Tag
        Invoke-NoteUp -Char $c
        $e.Handled = $true
    })

    $btn.Add_MouseLeave({
        param($sender, $e)
        $c = [string]$sender.Tag
        if ($script:pressedKeys.ContainsKey($c)) {
            Invoke-NoteUp -Char $c
        }
    })
}

# ============================
# 键盘事件绑定
# ============================
function Convert-KeyToChar {
    param(
        [System.Windows.Input.Key]$Key
    )
    if ($Key -ge [System.Windows.Input.Key]::A -and $Key -le [System.Windows.Input.Key]::Z) {
        $ch = $Key.ToString().ToLower()
        return $ch
    }

    if ($Key -ge [System.Windows.Input.Key]::D0 -and $Key -le [System.Windows.Input.Key]::D9) {
        $num = [int]$Key - [int][System.Windows.Input.Key]::D0
        return "$num"
    }

    switch ($Key) {
        ([System.Windows.Input.Key]::OemMinus)        { return "-" }
        ([System.Windows.Input.Key]::OemPlus)         { return "=" }
        ([System.Windows.Input.Key]::OemOpenBrackets) { return "[" }
        ([System.Windows.Input.Key]::OemCloseBrackets) { return "]" }
        ([System.Windows.Input.Key]::OemComma)        { return "," }
        ([System.Windows.Input.Key]::OemPeriod)       { return "." }
        ([System.Windows.Input.Key]::OemSemicolon)    { return ";" }
        ([System.Windows.Input.Key]::OemQuestion)     { return "/" }
        ([System.Windows.Input.Key]::OemQuotes)       { return "'" }
    }

    return $null
}

$window.Add_PreviewKeyDown({
    param($sender, $e)
    $char = Convert-KeyToChar $e.Key
    if ($null -ne $char -and $script:keyMap.ContainsKey($char)) {
        Invoke-NoteDown -Char $char
        $e.Handled = $true
    }
})

$window.Add_PreviewKeyUp({
    param($sender, $e)
    $char = Convert-KeyToChar $e.Key
    if ($null -ne $char -and $script:keyMap.ContainsKey($char)) {
        Invoke-NoteUp -Char $char
        $e.Handled = $true
    }
})

# ============================
# MIDI 播放
# ============================
function Start-MidiPlayback {
    param([string]$FilePath)

    $script:midiStopRequested = $false
    try {
        $script:midiNotes = [MidiParser]::Parse($FilePath)
    } catch {
        Write-Log "MIDI 解析失败: $($_.Exception.Message)"
        return
    }
    if ($script:midiNotes.Count -eq 0) {
        Write-Log "MIDI 文件中没有音符。"
        return
    }

    $script:midiPlayIndex = 0
    $script:midiStartTick = [Environment]::TickCount
    $script:midiPendingStops = @{}

    if ($script:midiPlayTimer) {
        $script:midiPlayTimer.Stop()
        $script:midiPlayTimer = $null
    }

    $script:midiPlayTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:midiPlayTimer.Interval = [TimeSpan]::FromMilliseconds(20)
    $script:midiPlayTimer.Add_Tick({
        $elapsed = [Environment]::TickCount - $script:midiStartTick

        foreach ($k in @($script:midiPendingStops.Keys)) {
            if ($script:midiPendingStops[$k] -le $elapsed) {
                [Kernel32PlayerBeep]::StopNote($k)
                $script:midiPendingStops.Remove($k)
            }
        }

        while ($script:midiPlayIndex -lt $script:midiNotes.Count -and -not $script:midiStopRequested) {
            $ev = $script:midiNotes[$script:midiPlayIndex]
            if ($ev.StartMs -gt $elapsed) { break }
            $freq = 440.0 * [math]::Pow(2.0, ($ev.MidiNote - 69) / 12.0)
            if ($freq -lt 37)    { $freq = 37 }
            if ($freq -gt 32767) { $freq = 32767 }
            $durMs = [int][math]::Max(1, [math]::Min($ev.DurationMs, 60000))
            $noteId = "midi_$($ev.MidiNote)_$($ev.StartMs)"
            [Kernel32PlayerBeep]::StartNote($noteId, [uint32][math]::Round($freq))
            $script:midiPendingStops[$noteId] = $ev.StartMs + $durMs
            $script:midiPlayIndex++
        }

        if ($script:midiPlayIndex -ge $script:midiNotes.Count) {
            $maxEnd = 0
            foreach ($v in $script:midiPendingStops.Values) { if ($v -gt $maxEnd) { $maxEnd = $v } }
            if ($elapsed -ge $maxEnd -or $script:midiPendingStops.Count -eq 0) {
                $script:midiPlayTimer.Stop()
                $script:midiPlayTimer = $null
                $script:mainMidiStopBtn.IsEnabled = $false
                $script:mainMidiOpenBtn.IsEnabled = $true
                Write-Log "MIDI 播放结束。"
            }
        }
        if ($script:midiStopRequested) {
            foreach ($k in @($script:midiPendingStops.Keys)) {
                [Kernel32PlayerBeep]::StopNote($k)
            }
            $script:midiPendingStops.Clear()
            $script:midiPlayTimer.Stop()
            $script:midiPlayTimer = $null
            $script:mainMidiStopBtn.IsEnabled = $false
            $script:mainMidiOpenBtn.IsEnabled = $true
            Write-Log "MIDI 播放已停止。"
        }
    })
    $script:midiPlayTimer.Start()
}

$midiOpenButton.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = "MIDI 文件 (*.mid;*.midi)|*.mid;*.midi|所有文件 (*.*)|*.*"
    $dialog.Title = "选择 MIDI 文件"
    $dialog.Multiselect = $false
    $dialog.CheckFileExists = $true
    $result = $dialog.ShowDialog($window)
    if ($result -eq $true -and $dialog.FileName) {
        $script:midiStopRequested = $false
        $midiOpenButton.IsEnabled = $false
        $midiStopButton.IsEnabled = $true
        Write-Log "正在解析 MIDI: $($dialog.FileName)"
        try {
            $notes = [MidiParser]::Parse($dialog.FileName)
            Write-Log "解析完成，共 $($notes.Count) 个音符，开始播放 (Kernel32 Beep)..."
            Start-MidiPlayback -FilePath $dialog.FileName
        } catch {
            Write-Log "MIDI 解析失败: $($_.Exception.Message)"
            $midiOpenButton.IsEnabled = $true
            $midiStopButton.IsEnabled = $false
        }
    } else {
        Write-Log "未选择文件。"
    }
})

$midiStopButton.Add_Click({
    $script:midiStopRequested = $true
    $midiStopButton.IsEnabled = $false
    $midiOpenButton.IsEnabled = $true
    Write-Log "已请求停止 MIDI 播放。"
})

# ============================
# 测试按钮
# ============================
$testButton.Add_Click({
    Write-Log "测试按钮：播放测试 Beep（750Hz，500ms）..."
    [Kernel32PlayerBeep]::TestBeep(750, 500)
    Write-Log "测试完成。"
})

Write-Log "Kernel32Player 已启动。"
Write-Log "使用鼠标点击琴键或使用键盘zsxdcvgbhnjm,l.;/q2w3e4rt6y7ui9o0p-[]演奏，c3-c6。"

$null = $window.ShowDialog()