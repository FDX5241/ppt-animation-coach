param(
  [Parameter(Mandatory=$true)][string]$CharacterAPath,
  [Parameter(Mandatory=$true)][string]$CharacterBPath,
  [Parameter(Mandatory=$true)][string]$ObjectSharpPath,
  [Parameter(Mandatory=$true)][string]$ObjectMidPath,
  [Parameter(Mandatory=$true)][string]$ObjectFrontPath,
  [string]$BackgroundPath,
  [string]$Title = 'PARALLAX',
  [string]$Subtitle = 'MORPH',
  [Parameter(Mandatory=$true)][string]$OutputPath,
  [switch]$AutoAdvance
)

$ErrorActionPreference = 'Stop'
$required = @($CharacterAPath,$CharacterBPath,$ObjectSharpPath,$ObjectMidPath,$ObjectFrontPath)
if ($BackgroundPath) { $required += $BackgroundPath }
foreach($path in $required) {
  if(-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "File not found: $path" }
}
$CharacterAPath = (Resolve-Path -LiteralPath $CharacterAPath).Path
$CharacterBPath = (Resolve-Path -LiteralPath $CharacterBPath).Path
$ObjectSharpPath = (Resolve-Path -LiteralPath $ObjectSharpPath).Path
$ObjectMidPath = (Resolve-Path -LiteralPath $ObjectMidPath).Path
$ObjectFrontPath = (Resolve-Path -LiteralPath $ObjectFrontPath).Path
if ($BackgroundPath) { $BackgroundPath = (Resolve-Path -LiteralPath $BackgroundPath).Path }
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputPath) | Out-Null

function Add-PictureNamed($slide,$path,$name,$box) {
  $shape = $slide.Shapes.AddPicture($path,0,-1,$box[0],$box[1],$box[2],$box[3])
  $shape.Name = $name
  return $shape
}

function Set-Box($shape,$box) {
  $shape.Left=$box[0]; $shape.Top=$box[1]; $shape.Width=$box[2]; $shape.Height=$box[3]
}

$ppt = New-Object -ComObject PowerPoint.Application
$presentation = $ppt.Presentations.Add()
$presentation.PageSetup.SlideWidth = 960
$presentation.PageSetup.SlideHeight = 540
$slide = $presentation.Slides.Add(1,12)

if ($BackgroundPath) {
  [void](Add-PictureNamed $slide $BackgroundPath 'background' @(0,0,960,540))
} else {
  $slide.FollowMasterBackground = 0
  $slide.Background.Fill.ForeColor.RGB = 0x241810
}

$backA = @(
  @(36,56,30,25),@(148,88,37,31),@(282,36,29,25),
  @(440,74,34,29),@(610,42,30,25),@(805,68,38,32),
  @(72,322,27,23),@(566,294,34,29),@(878,300,28,24)
)
$backB = @(
  @(18,61,30,25),@(125,93,37,31),@(255,42,29,25),
  @(409,80,34,29),@(574,48,30,25),@(764,74,38,32),
  @(25,328,27,23),@(516,300,34,29),@(820,306,28,24)
)
for($i=0;$i -lt 9;$i++) {
  $item = Add-PictureNamed $slide $ObjectSharpPath ("object-back-{0:D2}" -f ($i+1)) $backA[$i]
  $item.Rotation = (-24 + $i*19) % 360
}

[void](Add-PictureNamed $slide $CharacterAPath 'character-a' @(118,25,297,490))
[void](Add-PictureNamed $slide $CharacterBPath 'character-b' @(1040,25,300,490))

$titleBox = $slide.Shapes.AddTextbox(1,478,164,300,100)
$titleBox.Name = 'title-main'
$titleBox.TextFrame.TextRange.Text = $Title
$titleBox.TextFrame.TextRange.Font.Name = 'Microsoft YaHei'
$titleBox.TextFrame.TextRange.Font.Size = 62
$titleBox.TextFrame.TextRange.Font.Bold = -1
$titleBox.TextFrame.TextRange.Font.Color.RGB = 0x00F0E4D2
$titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$titleBox.TextFrame.VerticalAnchor = 3
$titleBox.TextFrame.MarginLeft = 0
$titleBox.TextFrame.MarginRight = 0

$subtitleBox = $slide.Shapes.AddTextbox(1,468,262,320,54)
$subtitleBox.Name = 'title-sub'
$subtitleBox.TextFrame.TextRange.Text = $Subtitle
$subtitleBox.TextFrame.TextRange.Font.Name = 'Microsoft YaHei'
$subtitleBox.TextFrame.TextRange.Font.Size = 28
$subtitleBox.TextFrame.TextRange.Font.Bold = -1
$subtitleBox.TextFrame.TextRange.Font.Color.RGB = 0x00DDD1BF
$subtitleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$subtitleBox.TextFrame.MarginLeft = 0
$subtitleBox.TextFrame.MarginRight = 0

$midA = @(
  @(-24,142,76,65),@(174,214,63,54),@(354,128,72,61),
  @(565,116,84,71),@(720,236,70,59),@(864,174,80,68)
)
$midB = @(
  @(-115,150,76,65),@(74,222,63,54),@(238,136,72,61),
  @(431,124,84,71),@(568,244,70,59),@(694,182,80,68)
)
for($i=0;$i -lt 6;$i++) {
  $item = Add-PictureNamed $slide $ObjectMidPath ("object-mid-{0:D2}" -f ($i+1)) $midA[$i]
  $item.Rotation = (12 + $i*33) % 360
}

$frontA = @(@(-55,360,230,195),@(760,-75,255,216),@(830,405,245,208))
$frontB = @(@(-330,370,245,208),@(330,-60,270,229),@(490,410,260,221))
for($i=0;$i -lt 3;$i++) {
  $item = Add-PictureNamed $slide $ObjectFrontPath ("object-front-{0:D2}" -f ($i+1)) $frontA[$i]
  $item.Rotation = (-18 + $i*42) % 360
}

$slide.SlideShowTransition.AdvanceOnClick = -1
if($AutoAdvance) {
  $slide.SlideShowTransition.AdvanceOnTime = -1
  $slide.SlideShowTransition.AdvanceTime = 2.2
}
$slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Parallax state A. All matched objects already exist. Character B starts outside the right edge. Back objects are sharp and small, middle objects are softly focused, and foreground objects are strongly defocused.'

$duplicate = $slide.Duplicate()
$slide2 = $duplicate.Item(1)
$slide2.MoveTo(2)
Set-Box $slide2.Shapes.Item('character-a') @(-355,30,297,490)
Set-Box $slide2.Shapes.Item('character-b') @(592,25,300,490)
Set-Box $slide2.Shapes.Item('title-main') @(174,160,300,100)
Set-Box $slide2.Shapes.Item('title-sub') @(164,258,320,54)
for($i=0;$i -lt 9;$i++) { Set-Box $slide2.Shapes.Item(("object-back-{0:D2}" -f ($i+1))) $backB[$i] }
for($i=0;$i -lt 6;$i++) { Set-Box $slide2.Shapes.Item(("object-mid-{0:D2}" -f ($i+1))) $midB[$i] }
for($i=0;$i -lt 3;$i++) { Set-Box $slide2.Shapes.Item(("object-front-{0:D2}" -f ($i+1))) $frontB[$i] }

$slide2.SlideShowTransition.EntryEffect = 3954
$slide2.SlideShowTransition.Duration = 2.0
$slide2.SlideShowTransition.AdvanceOnClick = -1
if($AutoAdvance) {
  $slide2.SlideShowTransition.AdvanceOnTime = -1
  $slide2.SlideShowTransition.AdvanceTime = 2.2
}
$slide2.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Parallax state B. This slide is a full duplicate of state A. Character A exits left while character B enters from the right. Back objects move least, middle objects move farther, and foreground objects move farthest. Morph is applied to this slide.'

$presentation.SaveAs($OutputPath,24)
$presentation.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output $OutputPath
