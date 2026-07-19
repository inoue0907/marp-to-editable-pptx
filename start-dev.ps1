$ErrorActionPreference = "Stop"

$vite = Join-Path $PSScriptRoot "node_modules\.bin\vite.cmd"
$node = Get-Command node -ErrorAction SilentlyContinue

if (-not $node) {
  $bundledNodeDir = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin"
  if (Test-Path (Join-Path $bundledNodeDir "node.exe")) {
    $env:Path = "$bundledNodeDir;$env:Path"
    $node = Get-Command node -ErrorAction SilentlyContinue
  }
}

if (-not $node) {
  Write-Error "Node.jsが見つかりません。https://nodejs.org/ からNode.js LTSをインストールしてください。"
}

if (-not (Test-Path $vite)) {
  Write-Host "依存関係をインストールしています..." -ForegroundColor Cyan
  $pnpm = Get-Command pnpm -ErrorAction SilentlyContinue
  $npm = Get-Command npm -ErrorAction SilentlyContinue
  if ($pnpm) {
    & $pnpm.Source install
  } elseif ($npm) {
    & $npm.Source install
  } else {
    Write-Error "npmまたはpnpmが見つかりません。Node.js LTSを再インストールしてください。"
  }
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

Write-Host "Marp to PowerPointを起動します: http://localhost:4173" -ForegroundColor Green
& $vite --host 127.0.0.1 --port 4173
