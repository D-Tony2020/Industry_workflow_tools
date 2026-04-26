# =============================================================
#  工厂订单转换工具 - 自动升级脚本
#  作用: 自动迁移用户数据 + 备份旧版本 + 部署新版本
#  调起方式: 双击同目录下的「升级.bat」
#
#  开发者可选参数（普通用户无需关心）:
#    -OldPath  <旧部署目录路径>   跳过文件夹选择对话框（用于自动化测试）
#    -NoLaunch                   完成后不弹出"是否启动"对话框（测试用）
# =============================================================

param(
    [string]$OldPath = '',
    [switch]$NoLaunch
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- 升级包自身路径（即本脚本所在目录）----------
$pkg = $PSScriptRoot
if (-not $pkg) { $pkg = Split-Path -Parent $MyInvocation.MyCommand.Path }
$pkg = $pkg.TrimEnd('\')

Write-Host '============================================================'
Write-Host '  工厂订单转换工具 - 升级助手'
Write-Host '============================================================'
Write-Host ('升级包目录: {0}' -f $pkg)
Write-Host ''

# ---------- Step 1: 选择旧部署目录（或使用传入参数）----------
if ($OldPath) {
    $old = $OldPath.TrimEnd('\')
    Write-Host ('[非交互模式] 使用传入旧部署目录: {0}' -f $old)
} else {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = '请选择您当前的「订单转换工具」部署目录（即里面有 mapping_table.xlsx 的那个文件夹）'
    $dialog.ShowNewFolderButton = $false

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host '用户已取消，未做任何修改。'
        exit 0
    }
    $old = $dialog.SelectedPath.TrimEnd('\')
}

# ---------- Step 2: 校验旧目录 ----------
if (-not (Test-Path (Join-Path $old '订单转换工具.exe'))) {
    [System.Windows.Forms.MessageBox]::Show(
        ('选定目录不是有效的部署目录(找不到 订单转换工具.exe):' + "`r`n`r`n" + $old + "`r`n`r`n" + '请重新双击「升级.bat」并选择正确的目录。'),
        '目录无效', 'OK', 'Error'
    ) | Out-Null
    exit 1
}

# 防止"旧目录 = 升级包目录"自我吞噬
if ((Resolve-Path $old).Path -eq (Resolve-Path $pkg).Path) {
    [System.Windows.Forms.MessageBox]::Show(
        '您选择的目录就是升级包本身。请把升级包先解压到一个临时位置(比如桌面)，然后双击其中的「升级.bat」。',
        '目录冲突', 'OK', 'Error'
    ) | Out-Null
    exit 1
}

Write-Host ('旧部署目录: {0}' -f $old)
Write-Host ''

# ---------- Step 3: 读取用户数据白名单 ----------
$whitelist_file = Join-Path $pkg 'user_data_files.txt'
$whitelist = @('mapping_table.xlsx', 'settings.json')   # 默认 fallback
if (Test-Path $whitelist_file) {
    $whitelist = Get-Content $whitelist_file -Encoding UTF8 |
        Where-Object { $_ -and -not $_.TrimStart().StartsWith('#') } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ }
}
Write-Host ('用户数据白名单: {0}' -f ($whitelist -join ', '))
Write-Host ''

# ---------- Step 4: 备份旧目录 ----------
$ts = Get-Date -Format 'yyyyMMdd-HHmmss'
$old_parent = Split-Path $old -Parent
$old_leaf   = Split-Path $old -Leaf
$backup     = Join-Path $old_parent ('{0}-backup-{1}' -f $old_leaf, $ts)

Write-Host ('[1/4] 备份旧目录到: {0}' -f $backup)
Move-Item -LiteralPath $old -Destination $backup -Force
Write-Host '      [OK] 已备份'
Write-Host ''

# ---------- Step 5: 把用户数据从备份目录迁移到升级包 ----------
Write-Host '[2/4] 迁移用户数据到新版本...'
$migrated = @()
foreach ($name in $whitelist) {
    $src = Join-Path $backup $name
    if (Test-Path -LiteralPath $src) {
        $dst = Join-Path $pkg $name
        if ((Get-Item -LiteralPath $src).PSIsContainer) {
            if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Recurse -Force }
            Copy-Item -LiteralPath $src -Destination $dst -Recurse -Force
        } else {
            Copy-Item -LiteralPath $src -Destination $dst -Force
        }
        $migrated += $name
        Write-Host ('      [OK] {0}' -f $name)
    }
}
if ($migrated.Count -eq 0) {
    Write-Host '      (备份目录中未发现白名单文件，可能是首次部署或客户尚未编辑映射表)'
}
Write-Host ''

# ---------- Step 6: 把升级包内容(去掉升级辅助文件)复制到旧目录原位置 ----------
Write-Host ('[3/4] 部署新版本到原位置: {0}' -f $old)
New-Item -ItemType Directory -Path $old -Force | Out-Null

# 这些文件不复制到客户部署目录（它们只属于升级包本身）
$exclude = @('升级.bat', '升级.ps1', 'user_data_files.txt', '部署说明.txt')

foreach ($item in Get-ChildItem -LiteralPath $pkg) {
    if ($exclude -contains $item.Name) { continue }
    if ($item.PSIsContainer) {
        Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $old $item.Name) -Recurse -Force
    } else {
        Copy-Item -LiteralPath $item.FullName -Destination $old -Force
    }
}
Write-Host '      [OK] 已部署'
Write-Host ''

# ---------- Step 7: 完成提示 ----------
Write-Host '[4/4] 升级完成!'
Write-Host ''

$summary = @"
升级完成 [OK]

  - 旧版本备份: $backup
  - 新版本部署: $old
  - 已迁移数据: $($migrated -join '、')

如发现问题需回滚，删除新版目录后把上面的备份目录改回原名即可。
是否立即启动新版本?
"@

if ($NoLaunch) {
    Write-Host $summary
    Write-Host ''
    Write-Host '[非交互模式] 跳过启动确认'
} else {
    $result = [System.Windows.Forms.MessageBox]::Show(
        $summary, '升级完成', 'YesNo', 'Information'
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $exe = Join-Path $old '订单转换工具.exe'
        if (Test-Path -LiteralPath $exe) {
            Start-Process -FilePath $exe
        }
    }
}
