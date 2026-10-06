param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Algorithm_Framework_English.pptx'),
    [string]$PreviewPath = (Join-Path $PSScriptRoot '..\out\Algorithm_Framework_English_preview.png')
)

$ErrorActionPreference = 'Stop'
$scale = 0.6
$ink = 0x2B2117
$line = 0x5D4833
$muted = 0xC2B6AA
$blue = 0xA87339
$blueFill = 0xFAF2EA
$green = 0x66804B
$greenFill = 0xEFF4EA
$gold = 0x3480AD
$goldFill = 0xE5F3FB
$panelFill = 0xFFFFFF
$softFill = 0xFBF9F7

function P([double]$v) { return [single]($v * $script:scale) }
function Rgb([int]$r,[int]$g,[int]$b) { return $r -bor ($g -shl 8) -bor ($b -shl 16) }

function Add-Rect {
    param($Slide,[string]$Name,[double]$X,[double]$Y,[double]$W,[double]$H,[int]$Fill,[int]$Stroke,[double]$Weight=1.2,[bool]$Rounded=$true)
    $type = if ($Rounded) { 5 } else { 1 }
    $s = $Slide.Shapes.AddShape($type,(P $X),(P $Y),(P $W),(P $H))
    $s.Name = $Name
    $s.Fill.Visible = -1
    $s.Fill.ForeColor.RGB = $Fill
    $s.Line.Visible = -1
    $s.Line.ForeColor.RGB = $Stroke
    $s.Line.Weight = $Weight
    if ($Rounded) { try { $s.Adjustments.Item(1) = 0.06 } catch {} }
    return $s
}

function Add-Text {
    param($Slide,[string]$Name,[string]$Text,[double]$X,[double]$Y,[double]$W,[double]$H,[double]$Size=14,[bool]$Bold=$false,[int]$Align=2,[int]$Color=0x17212B,[bool]$Italic=$false)
    $s = $Slide.Shapes.AddTextbox(1,(P $X),(P $Y),(P $W),(P $H))
    $s.Name = $Name
    $s.Line.Visible = 0
    $s.Fill.Visible = 0
    $tf = $s.TextFrame2
    $tf.MarginLeft = 1.5; $tf.MarginRight = 1.5; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tf.WordWrap = -1; $tf.AutoSize = 0; $tf.VerticalAnchor = 3
    $tr = $tf.TextRange
    $tr.Text = $Text
    $tr.ParagraphFormat.Alignment = $Align
    $tr.Font.Name = 'Arial'
    $tr.Font.Size = [single]($Size * $script:scale)
    $tr.Font.Bold = if ($Bold) { -1 } else { 0 }
    $tr.Font.Italic = if ($Italic) { -1 } else { 0 }
    $tr.Font.Fill.ForeColor.RGB = $Color
    return $s
}

function Add-Line {
    param($Slide,[string]$Name,[double]$X1,[double]$Y1,[double]$X2,[double]$Y2,[int]$Color=0x334155,[bool]$Dashed=$false,[bool]$Arrow=$false,[double]$Weight=1.8)
    $s = $Slide.Shapes.AddLine((P $X1),(P $Y1),(P $X2),(P $Y2))
    $s.Name = $Name
    $s.Line.Visible = -1
    $s.Line.ForeColor.RGB = $Color
    $s.Line.Weight = $Weight
    if ($Dashed) { $s.Line.DashStyle = 4 }
    if ($Arrow) { $s.Line.EndArrowheadStyle = 2; $s.Line.EndArrowheadLength = 2; $s.Line.EndArrowheadWidth = 2 }
    return $s
}

function Add-ArrowPath {
    param($Slide,[string]$Name,[double[][]]$Points,[int]$Color=0x334155,[bool]$Dashed=$false,[double]$Weight=1.8)
    for ($i=0; $i -lt $Points.Count - 1; $i++) {
        Add-Line -Slide $Slide -Name ("$Name`_$i") -X1 $Points[$i][0] -Y1 $Points[$i][1] -X2 $Points[$i+1][0] -Y2 $Points[$i+1][1] -Color $Color -Dashed $Dashed -Arrow ($i -eq $Points.Count - 2) -Weight $Weight | Out-Null
    }
}

$outputFullPath = [IO.Path]::GetFullPath($OutputPath)
$previewFullPath = [IO.Path]::GetFullPath($PreviewPath)
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($outputFullPath))
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($previewFullPath))
$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Add(0)
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540
    $slide = $pres.Slides.Add(1,12)
    $slide.Name = 'English Algorithm Framework'
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

    # Background panels
    Add-Rect $slide 'Online interaction panel' 40 88 300 555 (Rgb 255 255 255) (Rgb 170 182 194) 1.1 | Out-Null
    Add-Rect $slide 'Branch rollout panel' 375 88 650 555 (Rgb 255 255 255) (Rgb 170 182 194) 1.1 | Out-Null
    Add-Rect $slide 'Payment panel' 1060 88 500 290 (Rgb 255 255 255) (Rgb 170 182 194) 1.1 | Out-Null
    Add-Rect $slide 'Learning panel' 1060 410 500 233 (Rgb 255 255 255) (Rgb 170 182 194) 1.1 | Out-Null
    Add-Rect $slide 'Full branch band' 405 158 590 200 (Rgb 246 250 253) (Rgb 183 201 217) 1.0 | Out-Null
    Add-Rect $slide 'Counterfactual branch band' 405 386 590 220 (Rgb 252 250 246) (Rgb 221 207 181) 1.0 | Out-Null

    # Flow arrows, placed behind nodes
    Add-ArrowPath $slide 'State to report' @(@(190,199),@(190,211)) | Out-Null
    Add-ArrowPath $slide 'Report to branch sample' @(@(190,260),@(190,272)) | Out-Null
    Add-ArrowPath $slide 'Branch sample to UAV decision' @(@(190,321),@(190,333)) | Out-Null
    Add-ArrowPath $slide 'UAV decision to execution' @(@(190,382),@(190,394)) | Out-Null
    Add-ArrowPath $slide 'Execution to settlement' @(@(190,443),@(190,455)) | Out-Null
    Add-ArrowPath $slide 'Settlement to next state' @(@(190,504),@(190,516)) | Out-Null
    Add-ArrowPath $slide 'State transition feedback' @(@(80,540),@(55,540),@(55,174),@(80,174)) | Out-Null
    Add-ArrowPath $slide 'Branch to full rollout' @(@(300,284),@(350,284),@(350,261),@(420,261)) | Out-Null
    Add-ArrowPath $slide 'Branch to counterfactual rollout' @(@(300,308),@(366,308),@(366,490),@(420,490)) | Out-Null
    Add-ArrowPath $slide 'Full policy to MEC' @(@(580,261),@(615,261)) | Out-Null
    Add-ArrowPath $slide 'MEC to welfare' @(@(782,261),@(817,261)) | Out-Null
    Add-ArrowPath $slide 'Mask to counterfactual policy' @(@(580,490),@(615,490)) | Out-Null
    Add-ArrowPath $slide 'Counterfactual policy to environment' @(@(782,490),@(817,490)) | Out-Null
    Add-ArrowPath $slide 'Counterfactual data to critic' @(@(972,490),@(1040,490),@(1040,197),@(1084,197)) -Color (Rgb 37 99 166) -Dashed $true | Out-Null
    Add-ArrowPath $slide 'Full action to payment comparison' @(@(502,218),@(502,145),@(1040,145),@(1040,306),@(1084,306)) -Color (Rgb 37 99 166) -Dashed $true | Out-Null
    Add-ArrowPath $slide 'Critic to value comparison' @(@(1310,236),@(1310,258)) | Out-Null
    Add-ArrowPath $slide 'Payment to online settlement' @(@(1310,347),@(1310,370),@(356,370),@(356,478),@(300,478)) -Color (Rgb 37 99 166) -Dashed $true | Out-Null
    Add-ArrowPath $slide 'UAV replay to update' @(@(1215,505),@(1240,505)) | Out-Null
    Add-ArrowPath $slide 'User replay to update' @(@(1215,581),@(1240,581)) | Out-Null
    Add-Line $slide 'Legend divider' 45 690 1555 690 (Rgb 215 222 229) $false $false 0.9 | Out-Null
    Add-Line $slide 'Legend solid arrow' 70 730 130 730 (Rgb 51 65 85) $false $true 1.8 | Out-Null
    Add-Line $slide 'Legend dashed arrow' 430 730 490 730 (Rgb 37 99 166) $true $true 1.8 | Out-Null

    # Main nodes
    Add-Rect $slide 'System state' 80 150 220 49 (Rgb 247 249 251) (Rgb 100 116 139) 1.2 | Out-Null
    Add-Rect $slide 'User reports' 80 211 220 49 (Rgb 234 242 250) (Rgb 57 115 168) 1.2 | Out-Null
    Add-Rect $slide 'Branch sampling' 80 272 220 49 (Rgb 251 243 229) (Rgb 173 128 52) 1.2 | Out-Null
    Add-Rect $slide 'Online UAV decision' 80 333 220 49 (Rgb 234 242 250) (Rgb 57 115 168) 1.2 | Out-Null
    Add-Rect $slide 'Online MEC execution' 80 394 220 49 (Rgb 247 249 251) (Rgb 100 116 139) 1.2 | Out-Null
    Add-Rect $slide 'Online price settlement' 80 455 220 49 (Rgb 251 243 229) (Rgb 173 128 52) 1.2 | Out-Null
    Add-Rect $slide 'Next state' 80 516 220 49 (Rgb 247 249 251) (Rgb 100 116 139) 1.2 | Out-Null
    Add-Rect $slide 'UAV policy' 425 218 155 86 (Rgb 234 242 250) (Rgb 57 115 168) 1.2 | Out-Null
    Add-Rect $slide 'MEC environment' 622 218 160 86 (Rgb 247 249 251) (Rgb 100 116 139) 1.2 | Out-Null
    Add-Rect $slide 'Full welfare' 824 218 148 86 (Rgb 234 244 239) (Rgb 75 128 102) 1.2 | Out-Null
    Add-Rect $slide 'Mask user' 425 447 155 86 (Rgb 251 243 229) (Rgb 173 128 52) 1.2 | Out-Null
    Add-Rect $slide 'Counterfactual policy' 622 447 160 86 (Rgb 234 242 250) (Rgb 57 115 168) 1.2 | Out-Null
    Add-Rect $slide 'Reduced system' 824 447 148 86 (Rgb 234 244 239) (Rgb 75 128 102) 1.2 | Out-Null
    Add-Rect $slide 'Counterfactual critic' 1090 158 440 78 (Rgb 251 243 229) (Rgb 173 128 52) 1.2 | Out-Null
    Add-Rect $slide 'Value comparison' 1090 265 440 82 (Rgb 234 244 239) (Rgb 75 128 102) 1.2 | Out-Null
    Add-Rect $slide 'UAV replay buffer' 1090 474 125 62 (Rgb 247 249 251) (Rgb 100 116 139) 1.2 | Out-Null
    Add-Rect $slide 'UAV SAC update' 1247 474 283 62 (Rgb 234 242 250) (Rgb 57 115 168) 1.2 | Out-Null
    Add-Rect $slide 'User replay buffer' 1090 550 125 62 (Rgb 247 249 251) (Rgb 100 116 139) 1.2 | Out-Null
    Add-Rect $slide 'User SAC update' 1247 550 283 62 (Rgb 234 242 250) (Rgb 57 115 168) 1.2 | Out-Null

    # All figure text is kept as independent, editable text objects.
    Add-Text $slide 'Figure title' 'Long-Term Externality-Aware Multi-Agent SAC' 340 14 920 48 28 $true | Out-Null
    Add-Text $slide 'Online heading' '1. Online service and settlement' 64 102 285 38 18 $true 1 | Out-Null
    Add-Text $slide 'Rollout heading' '2. Branch-specific training rollouts' 399 102 610 38 18 $true 1 | Out-Null
    Add-Text $slide 'Payment heading' '3. Long-term externality payment' 1084 102 450 38 19 $true 1 | Out-Null
    Add-Text $slide 'Learning heading' '4. Per-agent SAC learning' 1084 424 450 38 19 $true 1 | Out-Null
    Add-Text $slide 'Full branch heading' 'Full-system branch (κ = 0)' 429 169 500 30 16 $true 1 | Out-Null
    Add-Text $slide 'Counterfactual heading' 'User-removal counterfactual (κ = n)' 429 397 520 30 16 $true 1 | Out-Null

    Add-Text $slide 'State label' "State and local observations`ns(t), {oᵢ(t)}" 85 151 210 47 12.5 | Out-Null
    Add-Text $slide 'Report label' "User agents report types`nθ̂ₙ(t) ~ πₙ(· | oₙ(t))" 85 212 210 47 12.5 | Out-Null
    Add-Text $slide 'Branch label' 'Select training branch κ' 85 273 210 47 12.5 | Out-Null
    Add-Text $slide 'Online action label' "UAV selects joint action`nTask admission and location" 85 334 210 47 12.5 | Out-Null
    Add-Text $slide 'Online execution label' "Execute selected branch`nMEC service and UAV movement" 85 395 210 47 12.5 | Out-Null
    Add-Text $slide 'Online settlement label' "Charge externality price pₙ`nFull-system branch only" 85 456 210 47 12.5 | Out-Null
    Add-Text $slide 'Payment transfer label' 'pₙ' 315 458 26 22 12 $false 2 (Rgb 37 99 166) | Out-Null
    Add-Text $slide 'Next state label' "Observe next state`ns(t+1), continue to next slot" 85 517 210 47 12.5 | Out-Null
    Add-Text $slide 'Full policy label' "UAV policy`na₀ ~ π₀(· | o₀)" 430 233 145 58 15 | Out-Null
    Add-Text $slide 'MEC label' "MEC environment`nExecute joint action" 627 233 150 58 15 | Out-Null
    Add-Text $slide 'Welfare label' "Welfare`nand transitions" 829 233 138 58 15 | Out-Null
    Add-Text $slide 'Full buffer note' 'Store transitions in per-agent replay buffers' 500 314 400 32 14 | Out-Null
    Add-Text $slide 'Mask label' "Mask user n`no₀⁻ⁿ" 430 462 145 58 15 | Out-Null
    Add-Text $slide 'Counterfactual action label' "Counterfactual`na₀⁻ⁿ ~ π₀" 627 462 150 58 15 | Out-Null
    Add-Text $slide 'Reduced welfare label' "Reduced system`nWelfare transition" 829 462 138 58 15 | Out-Null
    Add-Text $slide 'Counterfactual buffer note' 'Store counterfactual transition in the UAV replay buffer' 455 555 500 32 14 | Out-Null
    Add-Text $slide 'Critic label' "UAV SAC critic Q₀⁻ⁿ`nEvaluate the removal environment" 1100 169 420 56 16 | Out-Null
    Add-Text $slide 'Comparison label' "Compare long-term values`npₙ = Q₀⁻ⁿ,* − Q₀⁻ⁿ(Π₋ₙ(a₀*))" 1100 276 420 60 16 | Out-Null
    Add-Text $slide 'UAV buffer label' "UAV replay`nbuffer" 1095 480 115 50 14 | Out-Null
    Add-Text $slide 'UAV update label' "Update UAV actor`nand critics" 1252 480 273 50 14 | Out-Null
    Add-Text $slide 'User buffer label' "User n replay`nbuffer" 1095 556 115 50 14 | Out-Null
    Add-Text $slide 'User update label' "Update user n actor and critics`nfor each n ∈ N" 1252 556 273 50 13 | Out-Null
    Add-Text $slide 'Legend environment' 'Environment interaction / data flow' 144 718 250 28 14 | Out-Null
    Add-Text $slide 'Legend learning' 'Learning signal / value transfer' 504 718 260 28 14 | Out-Null

    try { $pres.BuiltInDocumentProperties.Item('Title').Value = 'Long-Term Externality-Aware Multi-Agent SAC Framework' } catch {}
    $pres.SaveAs($outputFullPath,24)
    $slide.Export($previewFullPath,'PNG',1920,1080)
    Write-Output "PPTX=$outputFullPath"
    Write-Output "PREVIEW=$previewFullPath"
    Write-Output "SHAPES=$($slide.Shapes.Count)"
}
finally {
    if ($null -ne $pres) { try { $pres.Close() } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($pres) } catch {} }
    if ($null -ne $ppt) { try { $ppt.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
