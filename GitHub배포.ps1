# GitHub 배포 스크립트 — 더블클릭으로 실행하세요
$env:PATH = "C:\Program Files\GitHub CLI;$env:PATH"
$ErrorActionPreference = "Continue"
$REPO_DIR = "c:\Users\2359_오희수\Desktop\ax-유통기한임박재고관리\expiry-stock-manager"
$REPO_NAME = "expiry-stock-manager"

Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║   유통기한 임박재고 관리 — 웹 배포 마법사  ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

Set-Location $REPO_DIR

# ── STEP 1: GitHub 로그인 ─────────────────────────────
$loginStatus = & gh auth status 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [1/3] GitHub 로그인" -ForegroundColor Yellow
    Write-Host "        브라우저가 열립니다. GitHub에 로그인하고" -ForegroundColor White
    Write-Host "        화면의 코드(숫자 8자리)를 붙여넣으세요." -ForegroundColor White
    Write-Host ""
    gh auth login --web --git-protocol https
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  로그인에 실패했습니다. 다시 시도해주세요." -ForegroundColor Red
        Read-Host "  [Enter]를 눌러 종료"
        exit
    }
    Write-Host "  ✅ 로그인 완료!" -ForegroundColor Green
} else {
    $user = & gh api user --jq ".login" 2>&1
    Write-Host "  [1/3] GitHub 로그인  ✅ 이미 로그인됨 ($user)" -ForegroundColor Green
}
Write-Host ""

# ── STEP 2: 저장소 생성 및 코드 업로드 ───────────────
Write-Host "  [2/3] GitHub 저장소 생성 및 코드 업로드 중..." -ForegroundColor Yellow

# 이미 있는지 확인
$existingRepo = & gh repo view $REPO_NAME --json url --jq ".url" 2>&1
if ($LASTEXITCODE -eq 0 -and $existingRepo -match "github.com") {
    Write-Host "        저장소가 이미 존재합니다. 코드를 최신 버전으로 업데이트합니다..." -ForegroundColor White
    $user = & gh api user --jq ".login" 2>&1
    git remote remove origin 2>$null
    git remote add origin "https://github.com/$user/$REPO_NAME.git"
    git push -u origin master --force
} else {
    & gh repo create $REPO_NAME --private --source . --push --remote origin
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "  업로드 중 오류가 발생했습니다." -ForegroundColor Red
    Read-Host "  [Enter]를 눌러 종료"
    exit
}

$repoUrl = & gh repo view --json url --jq ".url" 2>&1
Write-Host "  ✅ 업로드 완료!  저장소: $repoUrl" -ForegroundColor Green
Write-Host ""

# ── STEP 3: Streamlit Cloud 배포 안내 ────────────────
Write-Host "  [3/3] 웹사이트 배포 — 마지막 단계" -ForegroundColor Yellow
Write-Host ""
Write-Host "  ┌─────────────────────────────────────────┐"
Write-Host "  │  아래 순서로 진행하세요 (2~3분 소요)    │"
Write-Host "  ├─────────────────────────────────────────┤"
Write-Host "  │  ① share.streamlit.io 브라우저로 열기  │"
Write-Host "  │  ② GitHub 계정으로 로그인              │"
Write-Host "  │  ③ 'Create app' 클릭                   │"
Write-Host "  │  ④ Repository: $REPO_NAME 선택         │"
Write-Host "  │  ⑤ Main file path: app.py              │"
Write-Host "  │  ⑥ Deploy! 클릭 → URL 발급             │"
Write-Host "  └─────────────────────────────────────────┘"
Write-Host ""
Write-Host "  지금 share.streamlit.io 를 열까요? [Y/N] " -ForegroundColor Cyan -NoNewline
$answer = Read-Host
if ($answer -match "[Yy]") {
    Start-Process "https://share.streamlit.io"
    Write-Host ""
    Write-Host "  브라우저가 열렸습니다." -ForegroundColor Green
}
Write-Host ""
Write-Host "  배포 완료 후 URL(예: https://xxx.streamlit.app)을" -ForegroundColor Cyan
Write-Host "  북마크에 저장하면 언제 어디서나 접속 가능합니다." -ForegroundColor Cyan
Write-Host ""
Read-Host "  [Enter]를 눌러 종료"
