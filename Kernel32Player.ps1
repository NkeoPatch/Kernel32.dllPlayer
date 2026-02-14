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
            Beep(freq, 4294967294);  //
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
            Beep(1, 0);   // 37Hz, 0ms
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
            <RowDefinition Height="Auto"/>   <!-- 标题 -->
            <RowDefinition Height="Auto"/>   <!-- 测试按钮 -->
            <RowDefinition Height="*"/>      <!-- 键盘 + 日志 -->
        </Grid.RowDefinitions>

        <!-- 顶部标题 -->
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

        <!-- 测试按钮 -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
            <Button x:Name="TestButton"
                    Content="测试"
                    Width="80"
                    Height="30"
                    Margin="0,0,10,0"
                    Background="#222222"
                    Foreground="Lime"
                    BorderBrush="Gray"/>
            <TextBlock Text="使用鼠标点击琴键或使用键盘zsxdcvgbhnjm,l.;/q2w3e4rt6y7ui9o0p-[]演奏，c3-c6。"
                       VerticalAlignment="Center"
                       Foreground="LightGray"/>
        </StackPanel>

        <!-- 主体 -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="300"/>
            </Grid.ColumnDefinitions>

            <!-- 钢琴键盘（横向滚动） -->
            <ScrollViewer x:Name="PianoScroll"
                          Grid.Column="0"
                          HorizontalScrollBarVisibility="Visible"
                          VerticalScrollBarVisibility="Disabled"
                          Background="Black">
                <Canvas x:Name="PianoCanvas"
                        Height="180"
                        Background="Black"/>
            </ScrollViewer>

            <!-- 日志窗口 -->
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

# 解析 XAML 并获取控件
$xmlReader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window    = [Windows.Markup.XamlReader]::Load($xmlReader)

$pianoCanvas = $window.FindName('PianoCanvas')
$logBox      = $window.FindName('LogBox')
$testButton  = $window.FindName('TestButton')
$pianoScroll = $window.FindName('PianoScroll')

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

# 你指定的键盘序列
$keyCharsString = "zsxdcvgbhnjm,l.;/q2w3e4rt6y7ui9o0p-[]"
$keyChars = $keyCharsString.ToCharArray() | Where-Object { $_ -ne [char]0 }

# C3 的 MIDI 号：48；从 C3 开始连续映射
$startMidi = 48  # C3
$noteNames = "C","C#","D","D#","E","F","F#","G","G#","A","A#","B"

# 存储：字符 -> 音符信息
$script:keyMap      = @{} # char -> [pscustomobject]
$script:keyButtons  = @{} # char -> Button
$script:pressedKeys = @{} # char -> $true / $false

# 键盘视觉参数
$whiteKeyWidth  = 40
$whiteKeyHeight = 180
$blackKeyWidth  = 26
$blackKeyHeight = 110

$whiteIndex = 0

for ($i = 0; $i -lt $keyChars.Count; $i++) {
    $ch   = [string]$keyChars[$i]
    $midi = $startMidi + $i      # 从 C3 连续上行
    $index = $midi % 12
    $noteName = $noteNames[$index]

    # MIDI 到八度：C4(60) -> 4
    $octave = [int][math]::Floor($midi / 12) - 1

    # 计算频率（等音分，A4=440Hz）
    $n    = $midi - 69
    $freq = 440.0 * [math]::Pow(2.0, $n / 12.0)

    # Beep 限制范围（37 - 32767）
    if ($freq -lt 37)    { $freq = 37 }
    if ($freq -gt 32767) { $freq = 32767 }

    $isBlack = $noteName.Contains("#")

    # 确定在 Canvas 中的位置
    if (-not $isBlack) {
        $x = $whiteIndex * $whiteKeyWidth
        $whiteIndex++
    } else {
        # 黑键位于当前白键与前一白键之间
        $x = $whiteIndex * $whiteKeyWidth - ($blackKeyWidth / 2.0)
    }

    # 创建按键按钮：白键 / 黑键
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

        # 只在 C3~C6 上写字
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

    # Tag 存储键字符
    $btn.Tag = $ch

    # 加入 Canvas
    [void]$pianoCanvas.Children.Add($btn)

    # 保存映射
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

# 根据白键数量设置 Canvas 宽度，以便滚动
$pianoCanvas.Width = $whiteIndex * $whiteKeyWidth

# ============================
# 键盘视觉动画函数
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

    # 按下动画
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
        # 已经在按下状态，忽略重复
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

    # 停止持续 Beep，发送37hz 50ms Beep
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

    # 鼠标拖出键面时，若仍按下则视为松开
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

    # 处理字母 A-Z
    if ($Key -ge [System.Windows.Input.Key]::A -and $Key -le [System.Windows.Input.Key]::Z) {
        $ch = $Key.ToString().ToLower()
        return $ch
    }

    # 数字键 0-9（主键盘行）
    if ($Key -ge [System.Windows.Input.Key]::D0 -and $Key -le [System.Windows.Input.Key]::D9) {
        $num = [int]$Key - [int][System.Windows.Input.Key]::D0
        return "$num"
    }

    switch ($Key) {
        # 顶部符号行
        ([System.Windows.Input.Key]::OemMinus)        { return "-" }
        ([System.Windows.Input.Key]::OemPlus)         { return "=" }
        ([System.Windows.Input.Key]::OemOpenBrackets) { return "[" }
        ([System.Windows.Input.Key]::OemCloseBrackets) { return "]" }

        # 主键盘中下排
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
# 测试按钮
# ============================
$testButton.Add_Click({
    Write-Log "测试按钮：播放测试 Beep（750Hz，500ms）..."
    [Kernel32PlayerBeep]::TestBeep(750, 500)
    Write-Log "测试完成。"
})

# 初始日志
Write-Log "Kernel32Player 已启动。"
Write-Log "使用鼠标点击琴键或使用键盘zsxdcvgbhnjm,l.;/q2w3e4rt6y7ui9o0p-[]演奏，c3-c6。"

# 显示窗口（禁止最大化：XAML 中已设置 ResizeMode="CanMinimize"）
$null = $window.ShowDialog()