param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\out\长期外部性SAC训练框架_可编辑.pptx'),
    [string]$PreviewPath = (Join-Path $PSScriptRoot '..\out\长期外部性SAC训练框架_可编辑_预览.png')
)

$ErrorActionPreference = 'Stop'

function Get-Rgb([int]$Red, [int]$Green, [int]$Blue) {
    return $Red -bor ($Green -shl 8) -bor ($Blue -shl 16)
}

function Set-LineStyle {
    param(
        [object]$Shape,
        [int]$Color,
        [double]$Weight = 1.0,
        [bool]$Dashed = $false,
        [bool]$Arrow = $false
    )

    $Shape.Line.Visible = -1
    $Shape.Line.ForeColor.RGB = $Color
    $Shape.Line.Weight = $Weight
    if ($Dashed) {
        $Shape.Line.DashStyle = 4
    }
    if ($Arrow) {
        $Shape.Line.EndArrowheadStyle = 2
        $Shape.Line.EndArrowheadLength = 2
        $Shape.Line.EndArrowheadWidth = 2
    }
}

function Add-Frame {
    param(
        [object]$Slide,
        [string]$Name,
        [double]$X,
        [double]$Y,
        [double]$Width,
        [double]$Height,
        [int]$Color
    )

    $shape = $Slide.Shapes.AddShape(5, $X, $Y, $Width, $Height)
    $shape.Name = $Name
    $shape.Fill.Visible = 0
    Set-LineStyle -Shape $shape -Color $Color -Weight 0.9 -Dashed $true
    try { $shape.Adjustments.Item(1) = 0.06 } catch {}
    return $shape
}

function Add-Label {
    param(
        [object]$Slide,
        [string]$Name,
        [string]$Text,
        [double]$X,
        [double]$Y,
        [double]$Width,
        [double]$Height,
        [double]$FontSize,
        [int]$Color,
        [bool]$Bold = $false,
        [bool]$WhiteBackground = $false,
        [int]$Alignment = 2,
        [string]$FontName = 'SimSun'
    )

    $shape = $Slide.Shapes.AddTextbox(1, $X, $Y, $Width, $Height)
    $shape.Name = $Name
    $shape.Line.Visible = 0
    if ($WhiteBackground) {
        $shape.Fill.Visible = -1
        $shape.Fill.ForeColor.RGB = Get-Rgb 255 255 255
        $shape.Fill.Transparency = 0
    } else {
        $shape.Fill.Visible = 0
    }

    $frame = $shape.TextFrame2
    $frame.MarginLeft = 1
    $frame.MarginRight = 1
    $frame.MarginTop = 0
    $frame.MarginBottom = 0
    $frame.WordWrap = -1
    $frame.VerticalAnchor = 3
    $range = $frame.TextRange
    $range.Text = $Text
    $range.ParagraphFormat.Alignment = $Alignment
    $range.Font.Name = $FontName
    try { $range.Font.NameFarEast = $FontName } catch {}
    $range.Font.Size = $FontSize
    $range.Font.Fill.ForeColor.RGB = $Color
    if ($Bold) {
        $range.Font.Bold = -1
    }
    return $shape
}

function Add-Node {
    param(
        [object]$Slide,
        [string]$Name,
        [string]$Text,
        [double]$X,
        [double]$Y,
        [double]$Width,
        [double]$Height,
        [int]$FillColor,
        [int]$LineColor,
        [double]$FontSize = 12.0,
        [int]$MathFromParagraph = 0,
        [int]$ShapeType = 5
    )

    $shape = $Slide.Shapes.AddShape($ShapeType, $X, $Y, $Width, $Height)
    $shape.Name = $Name
    $shape.Fill.Visible = -1
    $shape.Fill.ForeColor.RGB = $FillColor
    $shape.Fill.Transparency = 0
    Set-LineStyle -Shape $shape -Color $LineColor -Weight 0.85
    if ($ShapeType -eq 5) {
        try { $shape.Adjustments.Item(1) = 0.05 } catch {}
    }

    $frame = $shape.TextFrame2
    $frame.MarginLeft = 5
    $frame.MarginRight = 5
    $frame.MarginTop = 3
    $frame.MarginBottom = 3
    $frame.WordWrap = -1
    $frame.AutoSize = 0
    $frame.VerticalAnchor = 3

    $range = $frame.TextRange
    $range.Text = $Text
    $range.ParagraphFormat.Alignment = 2
    $range.Font.Name = 'SimSun'
    try { $range.Font.NameFarEast = 'SimSun' } catch {}
    $range.Font.Size = $FontSize
    $range.Font.Fill.ForeColor.RGB = $LineColor

    if ($MathFromParagraph -gt 0) {
        try {
            $paragraphCount = $range.Paragraphs().Count
            for ($index = $MathFromParagraph; $index -le $paragraphCount; $index++) {
                $paragraph = $range.Paragraphs($index, 1)
                $paragraph.Font.Name = 'Cambria Math'
                $paragraph.Font.Size = [Math]::Max(9.5, $FontSize - 0.5)
                $paragraph.Font.Italic = -1
            }
        } catch {}
    }
    return $shape
}

function Add-Arrow {
    param(
        [object]$Slide,
        [string]$Name,
        [double]$X1,
        [double]$Y1,
        [double]$X2,
        [double]$Y2,
        [int]$Color,
        [bool]$Dashed = $false,
        [int]$ConnectorType = 1,
        [double]$Weight = 1.25
    )

    if ($ConnectorType -eq 1) {
        $shape = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
    } else {
        $shape = $Slide.Shapes.AddConnector($ConnectorType, $X1, $Y1, $X2, $Y2)
    }
    $shape.Name = $Name
    Set-LineStyle -Shape $shape -Color $Color -Weight $Weight -Dashed $Dashed -Arrow $true
    return $shape
}

function Add-LineSegment {
    param(
        [object]$Slide,
        [string]$Name,
        [double]$X1,
        [double]$Y1,
        [double]$X2,
        [double]$Y2,
        [int]$Color,
        [bool]$Dashed = $false,
        [double]$Weight = 1.25
    )

    $shape = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
    $shape.Name = $Name
    Set-LineStyle -Shape $shape -Color $Color -Weight $Weight -Dashed $Dashed
    return $shape
}

$black = Get-Rgb 25 25 25
$flowFill = Get-Rgb 244 244 255
$actorFill = Get-Rgb 235 248 252
$criticFill = Get-Rgb 237 252 237
$paymentFill = Get-Rgb 255 244 233
$memoryFill = Get-Rgb 247 247 247

$outputFullPath = [IO.Path]::GetFullPath($OutputPath)
$previewFullPath = [IO.Path]::GetFullPath($PreviewPath)
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($outputFullPath))
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($previewFullPath))

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Add(0)
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540
    $slide = $presentation.Slides.Add(1, 12)
    $slide.Name = '可编辑训练框架'
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.ForeColor.RGB = Get-Rgb 255 255 255

    # Frames are created first so all arrows and nodes remain selectable above them.
    $outerFrame = Add-Frame -Slide $slide -Name '00_整体框架_虚线框' -X 55 -Y 27 -Width 870 -Height 493 -Color $black
    $learningFrame = Add-Frame -Slide $slide -Name '10_SAC学习模块_虚线框' -X 410 -Y 67 -Width 195 -Height 445 -Color $black
    $paymentFrame = Add-Frame -Slide $slide -Name '20_长期外部性支付模块_虚线框' -X 650 -Y 75 -Width 260 -Height 332 -Color $black

    # Main online-decision flow.
    $arrowStateReport = Add-Arrow -Slide $slide -Name '30_箭头_状态到上报' -X1 212.5 -Y1 112 -X2 212.5 -Y2 133 -Color $black
    $arrowReportDecision = Add-Arrow -Slide $slide -Name '31_箭头_上报到决策' -X1 212.5 -Y1 195 -X2 212.5 -Y2 216 -Color $black
    $arrowDecisionEnv = Add-Arrow -Slide $slide -Name '32_箭头_决策到环境' -X1 212.5 -Y1 278 -X2 212.5 -Y2 300 -Color $black

    # State feedback loop, assembled from editable line segments.
    $feedbackBottom = Add-LineSegment -Slide $slide -Name '33_状态反馈_横线下' -X1 135 -Y1 343 -X2 82 -Y2 343 -Color $black
    $feedbackVertical = Add-LineSegment -Slide $slide -Name '34_状态反馈_竖线' -X1 82 -Y1 343 -X2 82 -Y2 81 -Color $black
    $feedbackTop = Add-Arrow -Slide $slide -Name '35_状态反馈_横线与箭头' -X1 82 -Y1 81 -X2 135 -Y2 81 -Color $black

    # Experience replay links.
    $arrowUserMemory = Add-Arrow -Slide $slide -Name '36_箭头_环境到用户经验池' -X1 172 -Y1 387 -X2 143 -Y2 408 -Color $black -Dashed $true -ConnectorType 3 -Weight 1.15
    $arrowUavMemory = Add-Arrow -Slide $slide -Name '37_箭头_环境到UAV经验池' -X1 255 -Y1 387 -X2 306 -Y2 408 -Color $black -Dashed $true -ConnectorType 3 -Weight 1.15

    # Policy-gradient and policy-output links.
    $arrowUserGradient = Add-Arrow -Slide $slide -Name '38_箭头_用户策略梯度' -X1 507 -Y1 186 -X2 507 -Y2 153 -Color $black -Dashed $true -Weight 1.15
    $arrowUavGradient = Add-Arrow -Slide $slide -Name '39_箭头_UAV策略梯度' -X1 507 -Y1 394 -X2 507 -Y2 362 -Color $black -Dashed $true -Weight 1.15
    $arrowUserAction = Add-Arrow -Slide $slide -Name '40_箭头_用户Actor到上报' -X1 427 -Y1 123 -X2 290 -Y2 164 -Color $black -ConnectorType 3
    $arrowUavAction = Add-Arrow -Slide $slide -Name '41_箭头_UAVActor到决策' -X1 427 -Y1 332 -X2 290 -Y2 247 -Color $black -ConnectorType 3

    # Long-term externality links and internal flow.
    $arrowActorCounterfactual = Add-Arrow -Slide $slide -Name '42_箭头_UAVActor到反事实样本' -X1 587 -Y1 332 -X2 666 -Y2 124 -Color $black -Dashed $true -ConnectorType 3
    $arrowCriticValue = Add-Arrow -Slide $slide -Name '43_箭头_UAVCritic到福利对照' -X1 587 -Y1 441 -X2 666 -Y2 224 -Color $black -Dashed $true -ConnectorType 3
    $arrowSampleValue = Add-Arrow -Slide $slide -Name '44_箭头_样本到福利对照' -X1 780 -Y1 155 -X2 780 -Y2 191 -Color $black
    $arrowValuePayment = Add-Arrow -Slide $slide -Name '45_箭头_福利对照到支付' -X1 780 -Y1 256 -X2 780 -Y2 289 -Color $black

    # Left online-decision nodes.
    $state = Add-Node -Slide $slide -Name '50_系统状态' -Text "系统状态`r`s(t)" -X 135 -Y 50 -Width 155 -Height 62 -FillColor $flowFill -LineColor $black -FontSize 13.5 -MathFromParagraph 2
    $report = Add-Node -Slide $slide -Name '51_移动用户上报' -Text "移动用户上报`r`θ̂ₙ(t) = πₙ(oₙ(t))" -X 135 -Y 133 -Width 155 -Height 62 -FillColor $flowFill -LineColor $black -FontSize 12.5 -MathFromParagraph 2
    $decision = Add-Node -Slide $slide -Name '52_UAV联合决策' -Text "UAV 联合决策`r`a₀(t) = π₀(o₀(t))" -X 135 -Y 216 -Width 155 -Height 62 -FillColor $flowFill -LineColor $black -FontSize 12.5 -MathFromParagraph 2
    $environment = Add-Node -Slide $slide -Name '53_UAV辅助MEC环境' -Text "UAV 辅助`r`MEC 环境`r`卸载、传输、计算与飞行" -X 135 -Y 300 -Width 155 -Height 87 -FillColor $flowFill -LineColor $black -FontSize 12.5

    # Editable cylinder shapes for replay buffers.
    $userMemory = Add-Node -Slide $slide -Name '54_用户经验池' -Text "用户经验池`r`{ℬₙ}ₙ∈N" -X 75 -Y 408 -Width 135 -Height 76 -FillColor $memoryFill -LineColor $black -FontSize 11.5 -MathFromParagraph 2 -ShapeType 13
    $uavMemory = Add-Node -Slide $slide -Name '55_UAV经验池' -Text "UAV 经验池 ℬ₀`r`完整与反事实转移" -X 224 -Y 408 -Width 168 -Height 76 -FillColor $memoryFill -LineColor $black -FontSize 11.5 -ShapeType 13

    # Multi-agent SAC nodes.
    $userActor = Add-Node -Slide $slide -Name '60_用户Actor' -Text "用户 Actor πₙ`r`生成类型上报" -X 427 -Y 93 -Width 160 -Height 60 -FillColor $actorFill -LineColor $black -FontSize 11.8
    $userCritic = Add-Node -Slide $slide -Name '61_用户Critic' -Text "用户 Critic Qₙ`r`评估上报策略" -X 427 -Y 186 -Width 160 -Height 60 -FillColor $criticFill -LineColor $black -FontSize 11.8
    $uavActor = Add-Node -Slide $slide -Name '62_UAV_Actor' -Text "UAV Actor π₀`r`生成联合决策" -X 427 -Y 302 -Width 160 -Height 60 -FillColor $actorFill -LineColor $black -FontSize 11.8
    $uavCritic = Add-Node -Slide $slide -Name '63_UAV_Critic' -Text "UAV Critic`r`Q₀, Q̂₀`r`评估完整与反事实价值" -X 427 -Y 394 -Width 160 -Height 94 -FillColor $criticFill -LineColor $black -FontSize 11.5 -MathFromParagraph 2

    # Long-term externality payment nodes.
    $counterfactual = Add-Node -Slide $slide -Name '70_反事实样本构造' -Text "反事实样本构造`r`κ = n,   o₀⁻ⁿ(t)" -X 666 -Y 93 -Width 228 -Height 62 -FillColor $paymentFill -LineColor $black -FontSize 12.3 -MathFromParagraph 2
    $valueCompare = Add-Node -Slide $slide -Name '71_反事实长期福利对照' -Text "反事实长期福利对照`r`Q₀⁻ⁿ,* 与 Q₀⁻ⁿ(Π₋ₙ(a₀*))" -X 666 -Y 191 -Width 228 -Height 65 -FillColor $paymentFill -LineColor $black -FontSize 12.0 -MathFromParagraph 2
    $payment = Add-Node -Slide $slide -Name '72_外部性支付' -Text "外部性支付`r`pₙ = Q₀⁻ⁿ,* −`r`Q₀⁻ⁿ(Π₋ₙ(a₀*))" -X 666 -Y 289 -Width 228 -Height 92 -FillColor $paymentFill -LineColor $black -FontSize 12.0 -MathFromParagraph 2

    # Titles and line annotations are separate editable text objects.
    $title = Add-Label -Slide $slide -Name '80_总标题' -Text '所提长期外部性感知多智能体 SAC 训练框架' -X 300 -Y 12 -Width 380 -Height 30 -FontSize 17 -Color $black -Bold $true -WhiteBackground $true
    $learningTitle = Add-Label -Slide $slide -Name '81_SAC模块标题' -Text '多智能体 SAC 学习模块' -X 402 -Y 49 -Width 212 -Height 27 -FontSize 15.5 -Color $black -Bold $true -WhiteBackground $true
    $paymentTitle = Add-Label -Slide $slide -Name '82_支付模块标题' -Text '长期外部性支付模块' -X 681 -Y 57 -Width 198 -Height 27 -FontSize 15.5 -Color $black -Bold $true -WhiteBackground $true
    $feedbackLabel = Add-Label -Slide $slide -Name '83_状态反馈标签' -Text 's(t+1)' -X 24 -Y 205 -Width 58 -Height 25 -FontSize 10.5 -Color $black -FontName 'Cambria Math'
    $userGradientLabel = Add-Label -Slide $slide -Name '84_用户策略梯度标签' -Text '策略梯度' -X 514 -Y 159 -Width 60 -Height 22 -FontSize 9.5 -Color $black -Alignment 1
    $uavGradientLabel = Add-Label -Slide $slide -Name '85_UAV策略梯度标签' -Text '策略梯度' -X 514 -Y 367 -Width 60 -Height 22 -FontSize 9.5 -Color $black -Alignment 1

    # Useful metadata for collaborators opening the file later.
    try {
        $presentation.BuiltInDocumentProperties.Item('Title').Value = '长期外部性感知多智能体 SAC 训练框架（可编辑）'
        $presentation.BuiltInDocumentProperties.Item('Comments').Value = '所有框、文字、经验池与连线均为 PowerPoint 原生可编辑对象。'
    } catch {}

    $presentation.SaveAs($outputFullPath, 24)
    $slide.Export($previewFullPath, 'PNG', 1920, 1080)

    Write-Output "PPTX=$outputFullPath"
    Write-Output "PREVIEW=$previewFullPath"
    Write-Output "SHAPES=$($slide.Shapes.Count)"
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) } catch {}
    }
    if ($null -ne $powerPoint) {
        try { $powerPoint.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
