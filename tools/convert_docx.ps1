param(
    [string]$InputDocx = (Join-Path $PSScriptRoot '..\李锡泽_7.12.docx'),
    [string]$OutputTex = (Join-Path $PSScriptRoot '..\src\main.tex')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$mathNs = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
$wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function Get-MathChild([System.Xml.XmlNode]$Node, [string]$Name) {
    foreach ($child in $Node.ChildNodes) {
        if ($child.NamespaceURI -eq $mathNs -and $child.LocalName -eq $Name) { return $child }
    }
    return $null
}

function Get-MathAttr([System.Xml.XmlNode]$Node, [string]$Name) {
    if ($null -eq $Node) { return $null }
    $attr = $Node.Attributes.GetNamedItem($Name)
    if ($null -eq $attr) { $attr = $Node.Attributes.GetNamedItem("w:$Name") }
    if ($null -eq $attr) { return $null }
    return $attr.Value
}

function Convert-MathText([string]$Text) {
    $map = @{
        'α'='\alpha '; 'β'='\beta '; 'γ'='\gamma '; 'δ'='\delta '; 'ε'='\epsilon '; 'ζ'='\zeta '; 'η'='\eta '; 'θ'='\theta '; 'ϑ'='\vartheta '; 'ι'='\iota '; 'κ'='\kappa '; 'λ'='\lambda '; 'μ'='\mu '; 'ν'='\nu '; 'ξ'='\xi '; 'ο'='o'; 'π'='\pi '; 'ρ'='\rho '; 'σ'='\sigma '; 'τ'='\tau '; 'υ'='\upsilon '; 'φ'='\phi '; 'χ'='\chi '; 'ψ'='\psi '; 'ω'='\omega '
        '∞'='\infty '; '∈'='\in '; '∉'='\notin '; '∅'='\emptyset '; '∑'='\sum '; '∏'='\prod '; '∫'='\int '; '∂'='\partial '; '∇'='\nabla '; '√'='\sqrt '; '≤'='\leq '; '≥'='\geq '; '≠'='\neq '; '≈'='\approx '; '≜'='\triangleq '; '∼'='\sim '; '±'='\pm '; '×'='\times '; '⋅'='\cdot '; '∙'='\cdot '; '⋯'='\cdots '; '…'='\ldots '; '→'='\rightarrow '; '←'='\leftarrow '; '↦'='\mapsto '; '∀'='\forall '; '∃'='\exists '; '∧'='\land '; '∨'='\lor '; '¬'='\neg '; '⊂'='\subset '; '⊆'='\subseteq '; '∪'='\cup '; '∩'='\cap '; '∥'='\parallel '; '‖'='\Vert '; '−'='-'; '–'='-'; '—'='-'; '⁡'=''; '′'="'"
    }
    $result = [System.Text.StringBuilder]::new()
    foreach ($char in $Text.ToCharArray()) {
        $key = [string]$char
        $handled = $false
        switch ([int][char]$char) {
            0x200A { $handled = $true }
            0x200B { $handled = $true }
            0x2004 { [void]$result.Append(' '); $handled = $true }
            0x27E8 { [void]$result.Append('\langle '); $handled = $true }
            0x27E9 { [void]$result.Append('\rangle '); $handled = $true }
            0x2223 { [void]$result.Append('\mid '); $handled = $true }
            0x2216 { [void]$result.Append('\setminus '); $handled = $true }
        }
        if ($handled) { continue }
        switch -CaseSensitive ($key) {
            'Γ' { [void]$result.Append('\Gamma '); continue }
            'Δ' { [void]$result.Append('\Delta '); continue }
            'Θ' { [void]$result.Append('\Theta '); continue }
            'Λ' { [void]$result.Append('\Lambda '); continue }
            'Ξ' { [void]$result.Append('\Xi '); continue }
            'Π' { [void]$result.Append('\Pi '); continue }
            'Σ' { [void]$result.Append('\Sigma '); continue }
            'Φ' { [void]$result.Append('\Phi '); continue }
            'Ψ' { [void]$result.Append('\Psi '); continue }
            'Ω' { [void]$result.Append('\Omega '); continue }
        }
        if ($map.ContainsKey($key)) { [void]$result.Append($map[$key]); continue }
        switch ($key) {
            '\' { [void]$result.Append('\backslash ') }
            '#' { [void]$result.Append('\#') }
            '%' { [void]$result.Append('\%') }
            '&' { [void]$result.Append('\mathbin{\&}') }
            '$' { [void]$result.Append('\$') }
            '~' { [void]$result.Append('\sim ') }
            '^' { [void]$result.Append('\mathbin{\hat{}}') }
            default { [void]$result.Append($char) }
        }
    }
    return $result.ToString()
}

function Convert-MathNode([System.Xml.XmlNode]$Node) {
    if ($Node.NamespaceURI -ne $mathNs) { return '' }
    switch ($Node.LocalName) {
        'oMath' { return (($Node.ChildNodes | ForEach-Object { Convert-MathNode $_ }) -join '') }
        'oMathPara' { return (($Node.ChildNodes | ForEach-Object { Convert-MathNode $_ }) -join '') }
        'r' {
            $value = (($Node.ChildNodes | Where-Object { $_.NamespaceURI -eq $mathNs -and $_.LocalName -eq 't' } | ForEach-Object { Convert-MathText $_.InnerText }) -join '')
            $style = Get-MathAttr (Get-MathChild (Get-MathChild $Node 'rPr') 'sty') 'val'
            if ($style -eq 'p' -and $value -match '[A-Za-z]') { return "\mathrm{$value}" }
            return $value
        }
        'sSub' { return "{0}_{{{1}}}" -f (Convert-MathNode (Get-MathChild $Node 'e')),(Convert-MathNode (Get-MathChild $Node 'sub')) }
        'sSup' { return "{0}^{{{1}}}" -f (Convert-MathNode (Get-MathChild $Node 'e')),(Convert-MathNode (Get-MathChild $Node 'sup')) }
        'sSubSup' { return "{0}_{{{1}}}^{{{2}}}" -f (Convert-MathNode (Get-MathChild $Node 'e')),(Convert-MathNode (Get-MathChild $Node 'sub')),(Convert-MathNode (Get-MathChild $Node 'sup')) }
        'f' { return "\frac{{{0}}}{{{1}}}" -f (Convert-MathNode (Get-MathChild $Node 'num')),(Convert-MathNode (Get-MathChild $Node 'den')) }
        'rad' {
            $degree = Convert-MathNode (Get-MathChild $Node 'deg')
            $body = Convert-MathNode (Get-MathChild $Node 'e')
            if ([string]::IsNullOrWhiteSpace($degree)) { return "\sqrt{$body}" }
            return "\sqrt[$degree]{$body}"
        }
        'd' {
            $properties = Get-MathChild $Node 'dPr'
            $beg = Get-MathAttr (Get-MathChild $properties 'begChr') 'val'
            $end = Get-MathAttr (Get-MathChild $properties 'endChr') 'val'
            if ($null -eq $beg) { $beg = '(' }; if ($null -eq $end) { $end = ')' }
            $pairs = @{ '('='\left('; ')'='\right)'; '['='\left['; ']'='\right]'; '{'='\left\lbrace '; '}'='\right\rbrace '; '|'='\left|'; '‖'='\left\lVert '; '⌈'='\left\lceil '; '⌉'='\right\rceil '; '⌊'='\left\lfloor '; '⌋'='\right\rfloor ' }
            $left = if ($pairs.ContainsKey($beg)) { $pairs[$beg] } else { "\left$beg" }
            $right = if ($pairs.ContainsKey($end)) { $pairs[$end] } else { "\right$end" }
            return "$left$(Convert-MathNode (Get-MathChild $Node 'e'))$right"
        }
        'nary' {
            $properties = Get-MathChild $Node 'naryPr'
            $char = Get-MathAttr (Get-MathChild $properties 'chr') 'val'
            $ops = @{ '∑'='\sum'; '∏'='\prod'; '∫'='\int'; '∬'='\iint'; '∭'='\iiint'; '⋂'='\bigcap'; '⋃'='\bigcup' }
            $op = if ($char -and $ops.ContainsKey($char)) { $ops[$char] } elseif ($char) { Convert-MathText $char } else { '\sum' }
            $sub = Convert-MathNode (Get-MathChild $Node 'sub'); $sup = Convert-MathNode (Get-MathChild $Node 'sup'); $body = Convert-MathNode (Get-MathChild $Node 'e')
            if ($sub) { $op += "_{$sub}" }; if ($sup) { $op += "^{$sup}" }
            return "$op $body"
        }
        'acc' {
            $char = Get-MathAttr (Get-MathChild (Get-MathChild $Node 'accPr') 'chr') 'val'
            $accents = @{ 'ˆ'='hat'; 'ˇ'='check'; 'ˉ'='bar'; '˙'='dot'; '¨'='ddot'; '⃗'='vec'; '˜'='tilde'; '¯'='overline' }
            $command = if ($char -and $accents.ContainsKey($char)) { $accents[$char] } else { 'hat' }
            return "\$command{$(Convert-MathNode (Get-MathChild $Node 'e'))}"
        }
        'func' { return "$(Convert-MathNode (Get-MathChild $Node 'fName')) $(Convert-MathNode (Get-MathChild $Node 'e'))" }
        'fName' { return (Convert-MathNode (Get-MathChild $Node 'e')) }
        'limLow' { return "$(Convert-MathNode (Get-MathChild $Node 'e'))_{$(Convert-MathNode (Get-MathChild $Node 'lim'))}" }
        'limUpp' { return "$(Convert-MathNode (Get-MathChild $Node 'e'))^{$(Convert-MathNode (Get-MathChild $Node 'lim'))}" }
        'eqArr' {
            $rows = @($Node.ChildNodes | Where-Object { $_.NamespaceURI -eq $mathNs -and $_.LocalName -eq 'e' } | ForEach-Object { Convert-MathNode $_ })
            return "\begin{aligned}`n$($rows -join ' \\ ' )`n\end{aligned}"
        }
        'm' {
            $rows = @($Node.ChildNodes | Where-Object { $_.NamespaceURI -eq $mathNs -and $_.LocalName -eq 'mr' } | ForEach-Object { Convert-MathNode $_ })
            return "\begin{bmatrix}$($rows -join ' \\ ')\end{bmatrix}"
        }
        'mr' {
            return (@($Node.ChildNodes | Where-Object { $_.NamespaceURI -eq $mathNs -and $_.LocalName -eq 'e' } | ForEach-Object { Convert-MathNode $_ }) -join ' & ')
        }
        'e' { return (($Node.ChildNodes | ForEach-Object { Convert-MathNode $_ }) -join '') }
        default { return (($Node.ChildNodes | ForEach-Object { Convert-MathNode $_ }) -join '') }
    }
}

function Escape-Text([string]$Text) {
    $Text = $Text.Replace([string][char]0x00A0, ' ').Replace([string][char]0x200A, '').Replace([string][char]0x200B, '').Replace([string][char]0x2004, ' ')
    $Text = $Text.Replace([string][char]0x03BA, '@@KAPPA@@').Replace([string][char]0x03B5, '@@EPSILON@@').Replace([string][char]0x2208, '@@IN@@')
    $escaped = $Text.Replace('\','\textbackslash{}').Replace('&','\&').Replace('%','\%').Replace('$','\$').Replace('#','\#').Replace('_','\_').Replace('{','\{').Replace('}','\}').Replace('~','\textasciitilde{}').Replace('^','\textasciicircum{}')
    return $escaped.Replace('@@KAPPA@@','\ensuremath{\kappa}').Replace('@@EPSILON@@','\ensuremath{\epsilon}').Replace('@@IN@@','\ensuremath{\in}')
}

$zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $InputDocx))
try {
    $documentEntry = $zip.GetEntry('word/document.xml')
    $reader = [IO.StreamReader]::new($documentEntry.Open())
    [xml]$document = $reader.ReadToEnd(); $reader.Close()

    $imageDir = Join-Path (Split-Path $OutputTex) 'images'
    New-Item -ItemType Directory -Force -Path $imageDir | Out-Null
    foreach ($entry in $zip.Entries | Where-Object { $_.FullName -like 'word/media/*' }) {
        $target = Join-Path $imageDir ([IO.Path]::GetFileName($entry.FullName))
        $input = $entry.Open(); $output = [IO.File]::Create($target); $input.CopyTo($output); $output.Close(); $input.Close()
    }

    $ns = [Xml.XmlNamespaceManager]::new($document.NameTable)
    $ns.AddNamespace('w',$wordNs); $ns.AddNamespace('m',$mathNs)
    $lines = [System.Collections.Generic.List[string]]::new()
    @(
        '% Auto-generated from 李锡泽_7.12.docx by tools/convert_docx.ps1.',
        '\documentclass[UTF8,11pt,a4paper]{ctexart}',
        '\usepackage[margin=2.54cm]{geometry}',
        '\usepackage{amsmath,amssymb,mathtools}',
        '\usepackage{graphicx}',
        '\usepackage{booktabs}',
        '\usepackage{float}',
        '\setlength{\parindent}{2em}',
        '\setlength{\parskip}{0.35em}',
        '\begin{document}',
        '\title{UAV辅助移动边缘计算中的长期动态激励与卸载优化}',
        '\author{李锡泽}',
        '\date{}',
        '\maketitle'
    ) | ForEach-Object { $lines.Add($_) }

    $sectionHeadings = @('Introduction','系统模型','仿真')
    $subsectionHeadings = @('通信模型','计算卸载模型','UAV能耗模型','动态拍卖','问题建模','RL设置','基于VCG的定价规则','反事实子拍卖的采样与 Q 函数更新')
    $subsubsectionHeadings = @('UAV设置','MU设置','UAV 观测','UAV动作','UAV奖励')
    foreach ($p in $document.SelectNodes('//w:body/w:p',$ns)) {
        $plain = (($p.SelectNodes('./w:r/w:t',$ns) | ForEach-Object { $_.InnerText }) -join '').Trim()
        $mathNodes = @($p.SelectNodes('./m:oMath | ./m:oMathPara',$ns))
        $hasDrawing = $null -ne $p.SelectSingleNode('.//w:drawing',$ns)
        if (-not $plain -and $mathNodes.Count -eq 0 -and -not $hasDrawing) { continue }
        if ($sectionHeadings -contains $plain) { $lines.Add("\section{$(Escape-Text $plain)}"); continue }
        if ($subsectionHeadings -contains $plain) { $lines.Add("\subsection{$(Escape-Text $plain)}"); continue }
        if ($subsubsectionHeadings -contains $plain) { $lines.Add("\subsubsection{$(Escape-Text $plain)}"); continue }
        if ($hasDrawing) {
            $lines.Add('\begin{figure}[H]'); $lines.Add('\centering'); $lines.Add('\includegraphics[width=0.7\linewidth]{images/image1.png}'); $lines.Add('\end{figure}'); continue
        }
        if ($mathNodes.Count -gt 0 -and -not $plain) {
            $formula = ($mathNodes | ForEach-Object { Convert-MathNode $_ }) -join '\quad '
            $lines.Add('\begin{equation*}'); $lines.Add($formula); $lines.Add('\end{equation*}'); continue
        }
        if ($mathNodes.Count -eq 0 -and $plain -match '\\' -and $plain -notmatch '[\u4E00-\u9FFF]') {
            $lines.Add('\begin{equation*}'); $lines.Add($plain); $lines.Add('\end{equation*}'); continue
        }
        $pieces = [System.Collections.Generic.List[string]]::new()
        foreach ($child in $p.ChildNodes) {
            if ($child.NamespaceURI -eq $wordNs -and $child.LocalName -eq 'r') {
                $text = (($child.SelectNodes('./w:t',$ns) | ForEach-Object { $_.InnerText }) -join '')
                if ($text) { $pieces.Add((Escape-Text $text)) }
            } elseif ($child.NamespaceURI -eq $mathNs -and ($child.LocalName -eq 'oMath' -or $child.LocalName -eq 'oMathPara')) {
                $pieces.Add(('$' + (Convert-MathNode $child) + '$'))
            }
        }
        if ($pieces.Count -gt 0) { $lines.Add(($pieces -join '')) }
    }
    if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -match '^\\section\{.*\}$') {
        $lines.RemoveAt($lines.Count - 1)
    }
    $lines.Add('\end{document}')
    [IO.File]::WriteAllLines((Resolve-Path (Split-Path $OutputTex)).Path + '\' + (Split-Path $OutputTex -Leaf), $lines, [Text.UTF8Encoding]::new($false))
} finally {
    $zip.Dispose()
}
