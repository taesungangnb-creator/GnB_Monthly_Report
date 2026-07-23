# GnB 성적표 입력 시스템

기본정보 입력 → 학생 성적 입력표 → 최종 성적표(GnB Monthly Report) 3단계로 구성된 웹앱입니다.
React + Vite + Recharts로 만들어졌습니다.

## GitHub에 올리기

1. GitHub에서 새 저장소(Repository)를 만듭니다. (예: `gnb-report-app`, Public/Private 무관)
2. 이 폴더 전체를 그대로 로컬 컴퓨터에 저장한 뒤, 터미널에서 아래 명령을 순서대로 실행합니다.

```bash
cd gnb-report-app
git init
git add .
git commit -m "GnB 성적표 입력 시스템 초기 버전"
git branch -M main
git remote add origin https://github.com/사용자명/저장소명.git
git push -u origin main
```

`node_modules`, `dist` 폴더는 `.gitignore`에 이미 등록되어 있어 자동으로 제외됩니다.

## Vercel 배포하기

1. https://vercel.com 접속 → GitHub 계정으로 로그인
2. "Add New... → Project" 클릭 → 방금 올린 저장소 선택 → Import
3. Framework Preset이 자동으로 **Vite**로 인식됩니다. (Build Command: `npm run build`, Output Directory: `dist` — 자동 설정되므로 그대로 두면 됩니다)
4. "Deploy" 클릭 → 잠시 후 배포 완료, `https://프로젝트명.vercel.app` 주소 생성

이후 GitHub에 코드를 push할 때마다 Vercel이 자동으로 재배포합니다(기존 G-Plum Hub와 동일한 방식).

## 로컬에서 미리보기

```bash
npm install
npm run dev
```

## 폴더 구조

```
gnb-report-app/
├── index.html
├── package.json
├── vite.config.js
├── src/
│   ├── main.jsx
│   └── App.jsx      ← 실제 앱 로직 전체 (3단계 화면)
└── README.md
```

## 주요 기능

- 3단계 입력 흐름 (기본정보 → 성적 입력표 → 최종 성적표)
- 성적 입력표: 영역별 만점 고정, 초과 입력 방지, 엑셀 템플릿 다운로드/일괄 업로드
- 최종 성적표: 학생별 인쇄, 전체 인쇄(학생 수만큼 A4 페이지 자동 생성), 엑셀로 결과 다운로드
- 인쇄는 A4 한 장 기준으로 스타일이 맞춰져 있습니다 (Chrome 브라우저에서 가장 안정적으로 동작합니다)

## 다음 확장 예정

- 교재별 엑셀 데이터 업로드 → 문항수/만점 자동 세팅, 교재 선택지 확장
- 입력 데이터 저장/불러오기 (현재는 새로고침 시 초기화됨)
