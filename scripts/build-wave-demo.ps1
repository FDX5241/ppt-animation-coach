$ErrorActionPreference = 'Stop'

$outputDir = 'C:/Users/35168/Documents/Codex/2026-09-13/skill-skill-github-github-skill-ppt/outputs'
$repoExampleDir = 'C:/Users/35168/Documents/Codex/2026-09-13/skill-skill-github-github-skill-ppt/work/ppt-animation-coach/examples'
New-Item -ItemType Directory -Force $outputDir | Out-Null
New-Item -ItemType Directory -Force $repoExampleDir | Out-Null

$finalPath = Join-Path $outputDir 'wave-animation-demo.pptx'
$examplePath = Join-Path $repoExampleDir 'wave-animation-demo.pptx'
Remove-Item -LiteralPath $finalPath, $examplePath -Force -ErrorAction SilentlyContinue

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1
$presentation = $pp.Presentations.Add()
$presentation.PageSetup.SlideWidth = 960
$presentation.PageSetup.SlideHeight = 540

$blankLayout = 12
$slideCount = 6
$bg = 0x271609
$ink = 0xF9F2E1
$muted = 0xC2B494
$cyan = 0xDCCD26
$blue = 0xB27D19

function Add-Text($slide, $text, $left, $top, $width, $height, $size, $color, $bold = $false) {
  $box = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
  $box.TextFrame.TextRange.Text = $text
  $box.TextFrame.TextRange.Font.Name = 'Aptos Display'
  $box.TextFrame.TextRange.Font.Size = $size
  $box.TextFrame.TextRange.Font.Bold = $bold
  $box.TextFrame.TextRange.Font.Color.RGB = $color
  $box.TextFrame.MarginLeft = 0
  $box.TextFrame.MarginRight = 0
  $box.TextFrame.MarginTop = 0
  $box.TextFrame.MarginBottom = 0
  return $box
}

function Add-WaveBand($slide, $offset, $top, $fillColor, $scale = 1.0) {
  $shapes = New-Object System.Collections.ArrayList
  $x = -240 + $offset
  $baseY = $top
  $waveWidth = 190
  $waveHeight = 74 * $scale
  $rect = $slide.Shapes.AddShape(1, $x, $baseY + 28 * $scale, 1440, 180 * $scale)
  $rect.Name = ('waveRect_' + $top + '_' + $offset)
  $rect.Fill.ForeColor.RGB = $fillColor
  $rect.Line.Visible = 0
  [void]$shapes.Add($rect)
  for ($i = 0; $i -lt 9; $i++) {
    $oval = $slide.Shapes.AddShape(9, $x + ($i * $waveWidth), $baseY, $waveWidth + 12, $waveHeight)
    $oval.Name = ('waveOval_' + $top + '_' + $offset + '_' + $i)
    $oval.Fill.ForeColor.RGB = $fillColor
    $oval.Line.Visible = 0
    [void]$shapes.Add($oval)
  }
  $shapeNames = @($shapes | ForEach-Object { $_.Name })
  $group = $slide.Shapes.Range($shapeNames).Group()
  return $group
}

for ($frame = 0; $frame -lt $slideCount; $frame++) {
  $slide = $presentation.Slides.Add($frame + 1, $blankLayout)
  $slide.FollowMasterBackground = 0
  $slide.Background.Fill.ForeColor.RGB = $bg

  Add-Text $slide 'Flowing wave' 72 56 500 48 30 $ink $true | Out-Null
  Add-Text $slide 'Editable PowerPoint frame animation' 74 108 500 28 14 $muted | Out-Null
  Add-Text $slide ('Frame ' + ($frame + 1) + ' / ' + $slideCount) 790 64 100 22 12 $muted | Out-Null

  $offset = ($frame * 115) % 190
  $back = Add-WaveBand $slide $offset 316 $blue 1.0
  $front = Add-WaveBand $slide ($offset + 70) 352 $cyan 0.78
  $front.Fill.Transparency = 0.08

  Add-Text $slide 'Duplicate the wave, shift it, and advance automatically' 72 458 700 24 14 $ink | Out-Null
  $notes = @"
海浪流动动画：制作说明

视觉原理
用一组重复的圆形波峰和底部矩形组成可编辑的波浪带。每一帧只改变波浪带的水平位置，再让页面按时间自动推进，就能产生连续流动的感觉。

复现步骤
1. 插入一个圆形，复制多个并横向排列，形成波峰。
2. 在波峰下方加一个矩形，统一填充颜色并组合。
3. 复制波浪组，每一页向左或向右移动固定距离。
4. 在“切换”中关闭单击换页，设置自动换页时间为 0.35 秒。
5. 把最后一页设置为循环播放，预览连续运动效果。

可调参数
- 移动距离：决定流动速度和连续性。
- 自动换页时间：越短越快，建议 0.25 到 0.60 秒之间测试。
- 前后波浪的颜色、透明度和垂直位置：决定层次感。

兼容性
本示例使用可编辑形状和自动换页，优先考虑 PowerPoint 与 WPS 播放兼容性。请在目标版本中实际预览，WPS 对自动换页和循环播放的细节可能存在差异。

来源
参考教程：PPT制作 | 海浪动画
作者：星瀚文化课件定制
链接：https://www.bilibili.com/video/BV1kX4y1f7EX/
参考日期：2026-09-13
"@
  $notesShape = $slide.NotesPage.Shapes.Placeholders.Item(2)
  $notesShape.TextFrame.TextRange.Text = $notes
  $slide.SlideShowTransition.AdvanceOnClick = 0
  $slide.SlideShowTransition.AdvanceOnTime = -1
  $slide.SlideShowTransition.AdvanceTime = 0.35
  $slide.SlideShowTransition.EntryEffect = 0
}

$presentation.SlideShowSettings.LoopUntilStopped = -1
$presentation.SaveAs($finalPath, 24)
$presentation.SaveCopyAs($examplePath, 24)
$presentation.Close()
$pp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
Write-Output $finalPath
