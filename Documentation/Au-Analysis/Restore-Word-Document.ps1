$ErrorActionPreference = "Stop"

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$base64File = Join-Path $scriptDirectory "Au_Analysis_Micro-to-SEM_Workflow_Guide_DRAFT.docx.b64"
$outputFile = Join-Path $scriptDirectory "Au_Analysis_Micro-to-SEM_Workflow_Guide_DRAFT.docx"

if (-not (Test-Path -LiteralPath $base64File)) {
    throw "The Base64 source file was not found: $base64File"
}

$base64Text = Get-Content -LiteralPath $base64File -Raw
$documentBytes = [Convert]::FromBase64String($base64Text)
[IO.File]::WriteAllBytes($outputFile, $documentBytes)

Write-Host "Word document restored successfully:"
Write-Host $outputFile
Write-Host "Size: $($documentBytes.Length) bytes"
