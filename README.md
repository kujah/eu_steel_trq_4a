# EU STEEL TRQ 4A 소진 현황

EU TARIC quota 페이지에서 `09xxxx` order number 데이터를 수집해 `allocation`, `used`, `balance`, `utilization`을 보여주는 정적 대시보드입니다.

## Files

- `fetch_data.py`: EU TARIC 데이터 수집 후 `public/data/orders.json`과 Excel 파일 생성
- `site/*`: 정적 대시보드 화면
- `scripts/build.mjs`: 배포용 `dist/` 생성
- `.github/workflows/update-data.yml`: 매일 19:00 Europe/Brussels 기준 데이터 갱신
- `.github/workflows/deploy-pages.yml`: GitHub Pages 배포

## Local Run

```powershell
cd C:\Users\kujah\.codex\eu-steel-trq-4a-dashboard
py fetch_data.py
npm run dev
```

브라우저에서 `http://localhost:3000` 접속.
