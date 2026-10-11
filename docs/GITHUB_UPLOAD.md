# GitHub에 처음 올리기

정리된 작업 폴더에는 Git 저장소가 초기화되어 있고 공개할 파일이 스테이징되어 있습니다. ZIP에는 `.git`과 실제 운영 설정을 넣지 않았으므로 ZIP을 풀었다면 아래 초기화 명령부터 실행하세요.

## 1. 빈 원격 저장소 생성

GitHub에서 `commerce-automation-tools` 저장소를 생성합니다. 로컬에 README와 .gitignore가 있으므로 GitHub 생성 화면에서는 README, .gitignore, LICENSE를 추가하지 않습니다.

## 2. 로컬 변경 확인 및 커밋

정리된 폴더에서 터미널을 열고 실행합니다. 아래 명령에는 실제 운영 설정과 엑셀 결과가 포함되지 않도록 `.gitignore`가 적용됩니다.

```powershell
git status
git diff --cached --stat
git commit -m "Organize commerce automation tools and documentation"
```

Git 작성자 이름과 이메일이 아직 설정되지 않았다면 먼저 본인 정보를 설정합니다. 이메일은 GitHub에서 제공하는 비공개 noreply 주소를 사용할 수 있습니다.

```powershell
git config user.name "본인 이름"
git config user.email "본인 Git 작성자 이메일"
```

## 3. 원격 연결 및 푸시

`YOUR_USERNAME`은 실제 GitHub 계정명으로 바꿉니다. GitHub 로그인 인증이 필요할 수 있습니다.

```powershell
git remote add origin https://github.com/YOUR_USERNAME/commerce-automation-tools.git
git push -u origin main
```

## ZIP을 풀어 사용하는 경우

커밋 전에 아래 초기화 명령을 실행합니다.

```powershell
git init -b main
git add .
```

## 이후 변경

```powershell
git add .
git diff --cached --stat
git commit -m "Update automation tools"
git push
```

`config.local.json`은 실행에만 사용하는 개인 설정입니다. `git add -f`로 강제 추가하지 않습니다. 공개 ZIP은 운영 설정을 제외한 배포본입니다.
