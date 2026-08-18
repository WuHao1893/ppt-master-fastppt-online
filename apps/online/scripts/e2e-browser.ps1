$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$webUrl = if ($env:E2E_WEB_URL) { $env:E2E_WEB_URL } else { 'http://localhost:5173' }
$apiUrl = if ($env:E2E_API_URL) { $env:E2E_API_URL } else { 'http://127.0.0.1:8787' }
$session = "fastppt-e2e-$PID"
$flow = (Resolve-Path (Join-Path $PSScriptRoot 'e2e-flow.js')).Path
$outputDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'output\playwright'
$screenshot = Join-Path $outputDir 'e2e-latest.png'

try {
  $health = Invoke-RestMethod "$apiUrl/api/v1/health"
  if (-not $health.ok) { throw 'FastPPT API health check failed.' }
  Invoke-WebRequest $webUrl -UseBasicParsing | Out-Null
  New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

  & npx --yes --package '@playwright/cli' playwright-cli "-s=$session" open $webUrl
  if ($LASTEXITCODE -ne 0) { throw 'Unable to open the browser session.' }
  & npx --yes --package '@playwright/cli' playwright-cli "-s=$session" run-code --filename $flow
  if ($LASTEXITCODE -ne 0) { throw 'Browser E2E flow failed.' }
  & npx --yes --package '@playwright/cli' playwright-cli "-s=$session" resize 1440 900
  & npx --yes --package '@playwright/cli' playwright-cli "-s=$session" screenshot --filename $screenshot --full-page
  if ($LASTEXITCODE -ne 0) { throw 'Browser E2E screenshot failed.' }
} finally {
  & npx --yes --package '@playwright/cli' playwright-cli "-s=$session" close 2>$null | Out-Null
}

Write-Output "Browser E2E passed. Screenshot: $screenshot"
