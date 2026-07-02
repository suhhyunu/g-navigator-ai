# G-Navigator AI (ATLAS AI)

생성형 AI 기반 국내·글로벌 물류 의사결정 내비게이터

> Navigate Every Vehicle, Anywhere.

## 소개

G-Navigator AI는 국내·국제 자동차 부품 물류의 경로, 비용, 리스크 판단을 한 화면에서 지원하는 AI 기반 물류 인텔리전스 플랫폼입니다. 기존 TMS/WMS가 데이터 조회에 머무는 것과 달리, 몬테카를로 시뮬레이션과 생성형 AI(Gemini)를 결합해 실행 가능한 의사결정까지 제시합니다.

## 핵심 기능

### 01. Global Route Command Center
- 다중 경유지 경로 설계 및 AI 최적 항로 생성
- 실시간 AIS 선박 추적(국내 60척·국제 150척 규모) 및 항만 혼잡도 확인
- 지정학적 리스크·태풍 경로 지도 시각화

### 02. AI Strategy & Decision Hub
- 전략 A/B/C(신속·균형·경제) 자동 생성 및 동시 비교
- 운송 수락·조건부·비추천 판단 AI
- 자연어 질의응답 및 AI 리스크 브리핑 (Gemini)

### 03. Risk & Cost Intelligence
- 몬테카를로 시뮬레이션(1,000회) 기반 비용·납기 예측
- 통관일수·HS 관세·CBAM·보험료 자동 계산
- 리스크 임계치 초과 시 대안 경로 자동 제안

### 04. Insights & Collaboration
- 진행중 배송 현황판 및 운행 패턴 분석
- PDF·Excel 리포트 생성, 이메일·Slack 공유
- 즐겨찾기 경로 저장 및 경로 비교

## AI 활용 방식

| 생성형 AI (Gemini) | 규칙·통계 기반 |
|---|---|
| 자연어 질의응답 | 전략 A/B/C 추천 (몬테카를로 시뮬레이션) |
| AI 리스크 브리핑 (뉴스 요약) | 운송 수락 판단 (확률 계산) |
| 운행 패턴 인사이트 | 리스크 자동 경고 (임계값 조건문) |

숫자로 정확히 계산 가능한 문제는 시뮬레이션·통계로, 언어 이해·요약·자유 질의가 필요한 문제만 생성형 AI로 처리합니다.

## Tech Stack

Python · Streamlit · Folium · Altair · Google Gemini API · SQLite

**외부 연동 (API 키 등록 시 실데이터로 자동 전환, 없으면 시뮬레이션 데이터 사용)**
- AISHub — 실시간 선박 AIS 데이터
- 기상청(공공데이터포털) — 태풍 경로
- Frankfurter — 실시간 환율
- 관세청 unipass — HS코드 관세율
- Google Gemini — AI 어시스턴트

## 설치 및 실행

```bash
# 1. 저장소 클론
git clone <이 저장소 URL>
cd <저장소 폴더>

# 2. 패키지 설치
pip install -r requirements.txt

# 3. 실행
streamlit run atlas_ai_merged.py
```

## 사용법

1. 앱 실행 후 좌측 사이드바에서 언어(한/영)와 API 키(선택)를 설정합니다.
2. `국내운송` / `국제운송` 탭에서 출발지·도착지·화물 정보를 입력합니다.
3. AI 추천 전략(A/B/C) 중 하나를 선택하고 시뮬레이션을 실행합니다.
4. 결과를 PDF/Excel로 내보내거나 이메일/Slack으로 공유할 수 있습니다.
5. `홈 대시보드` 탭에서 진행중인 배송 현황과 운행 패턴을 확인합니다.

## 참고 사항

- API 키가 없어도 모든 기능이 시뮬레이션 데이터로 동작합니다.
- SQLite DB(`atlas_ai_data.db`)는 로컬에 자동 생성되며, 즐겨찾기·시뮬레이션 이력·배송 현황이 저장됩니다.
- 항만 혼잡도·빈 화차 정보 등 일부 데이터는 실시간 공개 API가 없어 추정치로 대체되어 있습니다 (코드 주석 참고).

## Team ATLAS

문의: hyeonuseo199@gmail.com
