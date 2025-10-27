# ===============================
# 🧠 VBA_PARSER Portable Builder
# Autor: Adelina Trandafir
# Descriere: creează o versiune portabilă (folder cu .exe + DLL)
# ===============================

# Configurare căi
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Dist = Join-Path $Root "dist"

$Projects = @(
    "VBA_VIEWER",
    "VBA_TOKENIZER",
    "VBA_CORE"
)

# 1️⃣ Curăță vechiul folder dist
if (Test-Path $Dist) {
    Remove-Item -Recurse -Force $Dist
}
New-Item -ItemType Directory -Path $Dist | Out-Null

Write-Host "🧩 Building portable version from compiled projects..." -ForegroundColor Cyan

# 2️⃣ Verifică existența executabilelor
foreach ($proj in $Projects) {
    $binPath = Join-Path $Root "$proj\bin\Release"
    if (!(Test-Path $binPath)) {
        Write-Host "⚠️  Folderul $binPath nu există. Rulează mai întâi Build (Release)." -ForegroundColor Yellow
        continue
    }

    # 3️⃣ Copiază fișierele din bin\Release
    Write-Host "📦 Copiez din $proj..." -ForegroundColor Green
    Copy-Item "$binPath\*" $Dist -Recurse -Force
}

# 4️⃣ Elimină fișiere inutile (PDB, XML, manifest)
Get-ChildItem $Dist -Recurse -Include *.pdb, *.xml, *.manifest | Remove-Item -Force -ErrorAction SilentlyContinue

# 5️⃣ Verifică executabilele
$exes = Get-ChildItem $Dist -Filter "*.exe"
if ($exes.Count -gt 0) {
    Write-Host "✅ Build complet! Executabile disponibile:" -ForegroundColor Cyan
    $exes | ForEach-Object { Write-Host "   → " $_.Name -ForegroundColor White }
    Write-Host "`n📁 Folder final: $Dist" -ForegroundColor Cyan
} else {
    Write-Host "❌ Niciun executabil găsit. Verifică build-ul în Release mode." -ForegroundColor Red
}

# 6️⃣ Opțional: deschide folderul
Start-Process $Dist
