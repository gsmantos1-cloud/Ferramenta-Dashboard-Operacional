# ══════════════════════════════════════════════════════════════
#  GS Mantos — Backup Automático do Banco de Dados
#  Roda a cada hora via Agendador de Tarefas do Windows
#  Mantém: 48 backups horários + 30 backups diários
# ══════════════════════════════════════════════════════════════

$BaseDir   = "C:\Users\l3ti\Ferramenta de Atualizacoes de Pedidos"
$DbPath    = Join-Path $BaseDir "pedidos.db"
$BackupDir = Join-Path $BaseDir "backups"
$LogFile   = Join-Path $BackupDir "backup.log"

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts  $msg" | Out-File -Append -Encoding utf8 -FilePath $LogFile
}

# ── Verifica se o banco existe ──────────────────────────────────
if (-not (Test-Path $DbPath)) {
    Log "ERRO: banco não encontrado em $DbPath"
    exit 1
}

# ── Cria subpastas ──────────────────────────────────────────────
$HourlyDir = Join-Path $BackupDir "horario"
$DailyDir  = Join-Path $BackupDir "diario"
New-Item -ItemType Directory -Force -Path $HourlyDir | Out-Null
New-Item -ItemType Directory -Force -Path $DailyDir  | Out-Null

# ── Nome com timestamp ──────────────────────────────────────────
$Stamp   = Get-Date -Format "yyyy-MM-dd_HH-mm"
$DayStr  = Get-Date -Format "yyyy-MM-dd"

$HourlyDest = Join-Path $HourlyDir "pedidos_$Stamp.db"
$DailyDest  = Join-Path $DailyDir  "pedidos_$DayStr.db"

# ── Copia o banco usando SQLite online backup (via Python) ──────
#    Isso é seguro mesmo com o Flask rodando (WAL mode safe copy)
$PythonScript = @"
import sqlite3, shutil, sys
src = r'$DbPath'
dst = sys.argv[1]
try:
    src_con = sqlite3.connect(src)
    dst_con = sqlite3.connect(dst)
    src_con.backup(dst_con)
    dst_con.close()
    src_con.close()
    print('OK')
except Exception as e:
    print('ERRO:', e)
    sys.exit(1)
"@

$TmpScript = Join-Path $env:TEMP "gs_backup_tmp.py"
$PythonScript | Out-File -Encoding utf8 -FilePath $TmpScript

# Backup horário
$result = & python $TmpScript $HourlyDest 2>&1
if ($LASTEXITCODE -eq 0) {
    $size = [math]::Round((Get-Item $HourlyDest).Length / 1KB, 1)
    Log "Backup horário OK → $HourlyDest ($size KB)"
} else {
    Log "ERRO no backup horário: $result"
    exit 1
}

# Backup diário (sobrescreve o do dia, mantém um por dia)
& python $TmpScript $DailyDest 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    Log "Backup diário OK → $DailyDest"
}

Remove-Item $TmpScript -ErrorAction SilentlyContinue

# ── Limpeza: mantém apenas últimos 48 backups horários ─────────
$HourlyFiles = Get-ChildItem $HourlyDir -Filter "*.db" | Sort-Object LastWriteTime -Descending
if ($HourlyFiles.Count -gt 48) {
    $HourlyFiles | Select-Object -Skip 48 | ForEach-Object {
        Remove-Item $_.FullName -Force
        Log "Removido backup horário antigo: $($_.Name)"
    }
}

# ── Limpeza: mantém apenas últimos 30 backups diários ──────────
$DailyFiles = Get-ChildItem $DailyDir -Filter "*.db" | Sort-Object LastWriteTime -Descending
if ($DailyFiles.Count -gt 30) {
    $DailyFiles | Select-Object -Skip 30 | ForEach-Object {
        Remove-Item $_.FullName -Force
        Log "Removido backup diário antigo: $($_.Name)"
    }
}

# ── Mantém log com no máximo 500 linhas ────────────────────────
if (Test-Path $LogFile) {
    $lines = Get-Content $LogFile
    if ($lines.Count -gt 500) {
        $lines | Select-Object -Last 400 | Out-File -Encoding utf8 -FilePath $LogFile
    }
}

Log "─── Fim do backup ───────────────────────────────────────────"
