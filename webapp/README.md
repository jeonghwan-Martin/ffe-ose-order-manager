# FF&E / OS&E 발주관리 웹앱

호텔/모텔 리모델링 프로젝트의 룸타입·배치·FF&E·OS&E 발주 관리를 위한 웹앱입니다.
데이터는 구글 Apps Script를 통해 저장되어, 팀원들이 로그인 없이 같은 링크로 접속해 함께 씁니다.

## 처음 설정하는 순서

### 1) 데이터 저장소(Apps Script) 먼저 배포하기

`apps-script/README.md`를 따라 진행하세요. 끝나면 웹 앱 URL이 하나 나옵니다.

### 2) 이 프로젝트를 깃허브에 올리기

1. https://github.com 에서 계정이 없다면 새로 만드세요 (무료)
2. 새 저장소(Repository)를 만듭니다 — 이름은 예: `ffe-ose-order-manager` (Public으로 설정)
3. 이 폴더 전체를 그 저장소에 업로드합니다. 터미널을 쓸 줄 아시면:
   ```
   git init
   git add .
   git commit -m "init"
   git branch -M main
   git remote add origin https://github.com/<본인계정>/ffe-ose-order-manager.git
   git push -u origin main
   ```
   터미널이 낯설다면, 깃허브 웹사이트에서 "Add file → Upload files"로 폴더를 통째로 드래그해서 올려도 됩니다.

   > 이 단계는 Claude Code를 쓰시면 훨씬 수월해요 — "이 프로젝트를 깃허브 저장소에 올려줘"라고 요청하시면 됩니다.

### 3) 저장소 이름에 맞게 `vite.config.js` 수정

저장소 이름이 `ffe-ose-order-manager`가 아니라면, `vite.config.js`의 `base` 값을 저장소 이름에 맞게 고쳐주세요.

### 4) `src/api.js`에 Apps Script URL 붙여넣기

1단계에서 받은 웹 앱 URL을 `APPS_SCRIPT_URL`에 넣으세요.

### 5) 깃허브 페이지(Pages) 켜기

1. 깃허브 저장소 페이지에서 **Settings → Pages**
2. **Source**를 "GitHub Actions"로 선택
3. `main` 브랜치에 push하면 `.github/workflows/deploy.yml`이 자동으로 빌드·배포합니다 (몇 분 소요)
4. 배포가 끝나면 `https://<본인계정>.github.io/ffe-ose-order-manager/` 주소로 접속됩니다 — 이 링크를 팀원들과 공유하시면 돼요.

## 로컬에서 미리 보기 (선택)

```
npm install
npm run dev
```

## 폴더 구조

- `src/App.jsx` — 화면 전체 로직 (룸타입/배치/FF&E/OS&E/카탈로그/저장)
- `src/api.js` — Apps Script 백엔드와 통신하는 부분 (URL 설정 위치)
- `apps-script/Code.gs` — 데이터를 저장하는 백엔드 코드
- `.github/workflows/deploy.yml` — 깃허브에 push하면 자동으로 웹사이트를 배포하는 설정
