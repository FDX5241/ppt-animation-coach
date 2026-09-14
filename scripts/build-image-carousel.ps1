param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string[]]$ImagePaths,
  [Parameter(Mandatory = $true)]
  [string]$OutputPath,
  [switch]$AutoAdvance
)

$ErrorActionPreference = 'Stop'
if ($ImagePaths.Count -ne 5) { throw 'ImagePaths must contain exactly five image files.' }
foreach ($imagePath in $ImagePaths) {
  if (-not (Test-Path -LiteralPath $imagePath -PathType Leaf)) { throw "Image not found: $imagePath" }
}

$outputFolder = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Force $outputFolder | Out-Null
$ppt = New-Object -ComObject PowerPoint.Application
$presentation = $ppt.Presentations.Add()
$presentation.PageSetup.SlideWidth = 960
$presentation.PageSetup.SlideHeight = 540
$blankLayout = 12

function Add-Text($slide, $text, $left, $top, $width, $height, $size, $color, $bold = $false) {
  $box = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
  $box.TextFrame.TextRange.Text = $text
  $box.TextFrame.TextRange.Font.Name = 'Arial'
  $box.TextFrame.TextRange.Font.Size = $size
  $box.TextFrame.TextRange.Font.Bold = $bold
  $box.TextFrame.TextRange.Font.Color.RGB = $color
  $box.TextFrame.MarginLeft = 0; $box.TextFrame.MarginRight = 0
  $box.TextFrame.MarginTop = 0; $box.TextFrame.MarginBottom = 0
}

function Add-Card($slide, $name, $left, $top, $width, $height, $shadow) {
  $shadowShape = $slide.Shapes.AddShape(5, $left + 8, $top + 10, $width, $height)
  $shadowShape.Name = $shadow
  $shadowShape.Fill.ForeColor.RGB = 0xD8D8D8
  $shadowShape.Fill.Transparency = 0.25
  $shadowShape.Line.Visible = 0
  $card = $slide.Shapes.AddShape(5, $left, $top, $width, $height)
  $card.Name = $name
  $card.Fill.Solid(); $card.Fill.ForeColor.RGB = 0xFFFFFF
  $card.Line.Visible = 0
  return $card
}

$slides = @()
$slide = $presentation.Slides.Add(1, $blankLayout)
$slides += $slide
$slide.FollowMasterBackground = 0
$slide.Background.Fill.ForeColor.RGB = 0xFBFAF8
Add-Text $slide 'Image Carousel' 62 42 500 42 28 0x1E1E1E $true
Add-Text $slide 'Keep matching picture frames for a smooth Morph transition' 64 85 650 24 14 0x59636E
Add-Text $slide 'Each image moves clockwise into the next position' 64 474 700 22 14 0x59636E

[void](Add-Card $slide 'photo-center' 300 200 360 220 'shadow-center')
[void](Add-Card $slide 'photo-rightMid' 585 225 255 185 'shadow-rightMid')
[void](Add-Card $slide 'photo-rightSmall' 760 250 180 145 'shadow-rightSmall')
[void](Add-Card $slide 'photo-leftSmall' 20 250 180 145 'shadow-leftSmall')
[void](Add-Card $slide 'photo-leftMid' 120 225 255 185 'shadow-leftMid')
$slide.Shapes.Item('photo-center').ZOrder(0)
$slide.Shapes.Item('photo-rightSmall').ZOrder(1)

for ($i = 2; $i -le 5; $i++) {
  $duplicate = $slide.Duplicate()
  $slide = $duplicate.Item(1)
  $slide.MoveTo($i)
  $slides += $slide
}

$slotNames = @('photo-center', 'photo-rightMid', 'photo-rightSmall', 'photo-leftSmall', 'photo-leftMid')
$shadowNames = @('shadow-center', 'shadow-rightMid', 'shadow-rightSmall', 'shadow-leftSmall', 'shadow-leftMid')
$firstMapping = @(0, 4, 3, 2, 1)
for ($page = 1; $page -le 5; $page++) {
  $current = $slides[$page - 1]
  for ($slot = 0; $slot -lt 5; $slot++) {
    $sourceIndex = $firstMapping[($slot - ($page - 1) + 5) % 5]
    $current.Shapes.Item($slotNames[$slot]).Fill.UserPicture($ImagePaths[$sourceIndex])
  }
  $current.Shapes.Item('photo-center').ZOrder(0)
  $current.Shapes.Item('photo-rightSmall').ZOrder(1)
  foreach ($shadowName in $shadowNames) { $current.Shapes.Item($shadowName).ZOrder(1) }
  $current.SlideShowTransition.AdvanceOnClick = -1
  if ($AutoAdvance) { $current.SlideShowTransition.AdvanceOnTime = -1; $current.SlideShowTransition.AdvanceTime = 2.0 }
  if ($page -gt 1) { $current.SlideShowTransition.EntryEffect = 3954; $current.SlideShowTransition.Duration = 2.0 }
  $notes = "Image carousel construction`r`n`r`nDuplicate the entire slide so the five picture frames keep their correspondence. Do not paste a canvas selection into an existing slide: object IDs can change and Morph may fail.`r`n`r`nWithin a slide, the image positions are filled counterclockwise. Across slides, each image moves clockwise to the next frame.`r`r`n`r`nDefault: five slides, two-second Morph. Test WPS separately if needed."
  $current.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $notes
}

$presentation.SlideShowSettings.LoopUntilStopped = -1
$presentation.SaveAs($OutputPath, 24)
$presentation.Close(); $ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output $OutputPath
