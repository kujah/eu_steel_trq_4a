# EU STEEL TRQ 판재류 소진 현황

EU 2026/1457 규정(판재류 카테고리 1.A~10)에 속한 `09xxxx` order number 데이터를 수집해 `allocation`, `used`, `balance`, `utilization`을 카테고리별 탭으로 보여주는 정적 대시보드입니다.

## Files

- `fetch_data.py`: EU TARIC 데이터 수집 후 `public/data/orders.json`과 카테고리별 시트를 가진 Excel 파일 생성
- `public/data/history/`: 지나간 기준기간의 현황 JSON과 화면용 목록. 대시보드 하단의 **과거 기준기간**에서 조회
- `site/*`: 정적 대시보드 화면 (상단 탭으로 13개 카테고리 전환)
- `scripts/build.mjs`: 배포용 `dist/` 생성
- `.github/workflows/update-data.yml`: 매일 데이터 갱신
- `.github/workflows/deploy-pages.yml`: GitHub Pages 배포

## Workflow Note

- 데이터 갱신과 배포는 `.github/workflows/update-data.yml`과 `.github/workflows/deploy-pages.yml` 두 워크플로우로 나뉘어 있으며 `main`에 변경이 반영되면 실행됩니다.
- `update-data.yml`은 데이터 갱신 완료 후 GitHub Pages 배포까지 함께 수행합니다.
- 자동 갱신은 매일 01:00 UTC(한국시간 10:00)에 시작합니다. UTC 날짜에 유효한 쿼터 기간을 선택합니다.
- `deploy-pages.yml`은 일반 코드 변경에 대한 `push` 배포 담당입니다.
- 수집기는 TARIC 목록의 각 기간을 확인해 수집일(UTC)에 유효한 기준기간을 현재 화면에 표시합니다. 지난 기간은 종료 후 TARIC에서 다시 조회해 최초 1회 보관하며, 과거 화면의 `Generated`는 그 자료를 수집한 시점입니다. 일별 과거 스냅샷은 아닙니다.

## Local Run

```powershell
cd C:\Users\kujah\Desktop\coding\eu-steel-trq-4a-dashboard
py fetch_data.py
npm run dev
```

브라우저에서 `http://localhost:3000` 접속.
