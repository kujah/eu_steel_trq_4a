# EU STEEL TRQ 판재류 소진 현황

EU 2026/1457 규정(판재류 카테고리 1.A~10)에 속한 `09xxxx` order number 데이터를 수집해 `allocation`, `used`, `balance`, `utilization`을 카테고리별 탭으로 보여주는 정적 대시보드입니다.

## Files

- `fetch_data.py`: EU TARIC 데이터 수집 후 `public/data/orders.json`과 카테고리별 시트를 가진 Excel 파일 생성
- `site/*`: 정적 대시보드 화면 (상단 탭으로 13개 카테고리 전환)
- `scripts/build.mjs`: 배포용 `dist/` 생성
- `.github/workflows/update-data.yml`: 매일 데이터 갱신
- `.github/workflows/deploy-pages.yml`: GitHub Pages 배포

## Workflow Note

- 데이터 갱신과 배포는 `.github/workflows/update-data.yml`과 `.github/workflows/deploy-pages.yml` 두 워크플로우로 나뉘어 있으며 `main`에 변경이 반영되면 실행됩니다.
- `update-data.yml`은 데이터 갱신 완료 후 GitHub Pages 배포까지 함께 수행합니다.
- `deploy-pages.yml`은 일반 코드 변경에 대한 `push` 배포 담당입니다.

## Local Run

```powershell
cd C:\Users\kujah\Desktop\coding\eu-steel-trq-4a-dashboard
py fetch_data.py
npm run dev
```

브라우저에서 `http://localhost:3000` 접속.
