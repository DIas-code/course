$ErrorActionPreference = 'Stop'
$path = 'C:\Users\Dias\Desktop\course\УМКД_ПМ01_Готовый.docx'
$out  = 'C:\Users\Dias\AppData\Local\Temp\claude\C--Users-Dias-Desktop-course\767528e1-1ebd-40dc-8f49-105500f0c51f\scratchpad\headings.txt'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$lines = New-Object System.Collections.Generic.List[string]
try {
    $doc = $word.Documents.Open($path, $false, $true)   # ConfirmConversions=false, ReadOnly=true
    $doc.Repaginate()
    foreach ($p in $doc.Paragraphs) {
        $t = $p.Range.Text
        if ($null -eq $t) { continue }
        $t = $t.Trim()
        if ($t.Length -eq 0 -or $t.Length -gt 200) { continue }
        $hit = $false
        if ($t -ceq 'CONTENT') { $hit = $true }
        elseif ($t -ceq 'INTRODUCTION') { $hit = $true }
        elseif ($t -ceq 'I. ACADEMIC CALENDAR OF MODULE') { $hit = $true }
        elseif ($t -cmatch '^(II|III|IV|V|VI|VII|VIII)\. LO ') { $hit = $true }
        elseif ($t -ceq '1. ACADEMIC CALENDAR') { $hit = $true }
        elseif ($t -ceq '2. GLOSSARY') { $hit = $true }
        elseif ($t -ceq '3. LECTURE MATERIALS') { $hit = $true }
        elseif ($t -cmatch '^4\. (PRACTICAL CLASS|LABORATORY WORK|INDUSTRIAL TRAINING) MATERIALS$') { $hit = $true }
        elseif ($t -ceq '5. LIST OF SOFTWARE AND MULTIMEDIA SUPPORT FOR TRAINING SESSIONS') { $hit = $true }
        elseif ($t -ceq '6. LIST OF LITERATURE') { $hit = $true }
        elseif ($t -ceq 'IX. SELF-STUDY TOPICS FOR THE MODULE') { $hit = $true }
        elseif ($t -ceq 'X. EXAM QUESTIONS') { $hit = $true }
        if ($hit) {
            $pg = $p.Range.Information(3)   # wdActiveEndPageNumber
            $lines.Add(("{0}`t{1}" -f $pg, $t))
        }
    }
    $doc.Close($false)   # do not save
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
[System.IO.File]::WriteAllLines($out, $lines, [System.Text.UTF8Encoding]::new($false))
Write-Output ("Collected {0} headings -> {1}" -f $lines.Count, $out)
