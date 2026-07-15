# Clove Container Tracker

Excel 파일 업데이트 → GitHub Actions 자동 실행 → GCS 배포

## 🔗 링크
https://storage.googleapis.com/wholesale-us-map/restock_tracker.html

## 📋 사용 방법
1. `All_Orders_Rolling_Pre-Orders.xlsx` 파일 업데이트
2. GitHub에 파일 업로드 (push)
3. GitHub Actions가 자동으로 HTML 재생성 + GCS 업로드
4. 링크 공유

## ⚙️ 설정 (최초 1회)
GitHub repo Settings → Secrets → `GCS_SA_KEY` 추가
(Google Cloud Service Account JSON 키)
