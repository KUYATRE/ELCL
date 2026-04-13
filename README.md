# Electrical Capacity & Breaker Helper v2

## 추가 반영 내용
- 드래그 오류 수정 (`QMimeData()` 직접 생성)
- 외부 QSS 스타일 적용
- 크롬 느낌의 버튼/입력창/스크롤바 스타일
- 카드/버튼 그림자 효과 추가
- 캔버스에서 선택 항목 삭제 버튼 추가
- 차단기 아래에 하위 차단기 추가 가능
- 하위 차단기까지 포함한 재귀 합산 계산

## 실행
```bash
pip install -r requirements.txt
python main.py
```

## 사용 방법
- 파트 관리 탭에서 부품 등록
- 차단기 캔버스 탭 좌측 라이브러리에서 부품 드래그
- 부품을 차단기 위에 놓으면 하위 부하 추가
- 하위 차단기 템플릿을 차단기 위에 놓으면 하위 차단기 추가
- 차단기 클릭 시 안전율 수정
- 부하 더블클릭 시 수량 수정
- 항목 선택 후 `선택 항목 삭제` 버튼으로 삭제
- 최상위 차단기는 삭제 제한

## 저장 파일
- DB: `data/parts.db`
- 캔버스: `data/breaker_canvas_layout.json`
- 로그: `logs/app.log`
