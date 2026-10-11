$ErrorActionPreference = 'Stop'
$files = @(
    @{ Input = 'C:\Users\LHW66\Documents\Codex\2026-09-16\bang\outputs\第一题.docx'; Pdf = 'C:\Users\LHW66\Documents\Codex\2026-09-16\bang\work\qa\第一题\第一题.pdf' },
    @{ Input = 'C:\Users\LHW66\Documents\Codex\2026-09-16\bang\outputs\第二题.docx'; Pdf = 'C:\Users\LHW66\Documents\Codex\2026-09-16\bang\work\qa\第二题\第二题.pdf' }
)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    foreach ($file in $files) {
        $dir = Split-Path -Parent $file.Pdf
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $doc = $word.Documents.Open($file.Input, $false, $true)
        try {
            $doc.ExportAsFixedFormat($file.Pdf, 17, $false)
        }
        finally {
            $doc.Close($false)
        }
        Write-Output $file.Pdf
    }
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
