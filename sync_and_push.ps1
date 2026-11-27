# GitHub 자동 동기화 스크립트
# 소스 파일을 C:\kr-tools 로 복사하고 변경분이 있으면 commit + push

$REPO = "C:\kr-tools"
$EXCLUDE = @("*.log","done_*","Thumbs.db","desktop.ini")

function Sync-Dir($src, $dst, $excludeFiles = @(), $excludeDirs = @("__pycache__","chrome-profile","masked_output","screenshots",".git",".claude")) {
    if (-not (Test-Path $dst)) { New-Item -ItemType Directory -Force -Path $dst | Out-Null }
    Get-ChildItem $src -File | Where-Object {
        $name = $_.Name
        -not ($excludeFiles | Where-Object { $name -like $_ })
    } | ForEach-Object {
        $dest = Join-Path $dst $_.Name
        if (-not (Test-Path $dest) -or $_.LastWriteTime -gt (Get-Item $dest).LastWriteTime) {
            Copy-Item $_.FullName $dest -Force
        }
    }
    Get-ChildItem $src -Directory | Where-Object { $_.Name -notin $excludeDirs } | ForEach-Object {
        Sync-Dir $_.FullName (Join-Path $dst $_.Name) $excludeFiles $excludeDirs
    }
}

# 소스 → 레포 동기화
Sync-Dir "C:\krlaw"                              "$REPO\law-mcp-server"   $EXCLUDE @("__pycache__",".claude",".git")
Sync-Dir "C:\hometax-axtax"                      "$REPO\hometax-axtax"    $EXCLUDE @("__pycache__","chrome-profile","masked_output",".git")
Sync-Dir "C:\Users\admin\.claude\skills"         "$REPO\skills"           $EXCLUDE @("__pycache__")
Sync-Dir "C:\Users\admin\.claude\scheduled-tasks" "$REPO\scheduled-tasks" $EXCLUDE @("__pycache__","screenshots")

# 변경 감지 후 commit + push
Set-Location $REPO
$status = git status --porcelain
if ($status) {
    $date = Get-Date -Format "yyyy-MM-dd HH:mm"
    git add .
    git commit -m "자동 백업 $date"
    git push origin master
    Write-Host "[$date] push 완료"
} else {
    Write-Host "[$(Get-Date -Format 'HH:mm')] 변경 없음"
}
