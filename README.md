# PDF Parser PyQt5 GUI - Invoice & Packing List

PDF 파일에서 인보이스와 패킹리스트 데이터를 추출하여 Excel 파일로 변환하는 현대적인 GUI 도구입니다.

## 기능

- **🎨 현대적인 GUI**: PyQt5 기반의 아름답고 직관적인 인터페이스
- **📁 드래그 앤 드롭**: 파일을 쉽게 선택할 수 있는 파일 브라우저
- **📊 실시간 진행상황**: 진행률 바와 상세한 로그로 변환 과정 확인
- **🔄 멀티스레딩**: UI 블로킹 없이 백그라운드에서 안전한 변환 작업
- **💾 스마트 파일명**: 선택한 파일을 기반으로 출력 파일명 자동 생성
- **🎯 안정성**: PyQt5의 강력한 이벤트 시스템으로 안정적인 동작

## 사용법

### PyQt5 GUI 실행파일 사용 (Windows)

1. `PDF-Parser-PyQt5.exe` 파일을 다운로드
2. 실행파일을 더블클릭하여 실행
3. 현대적인 GUI에서 파일 선택:
   - **인보이스 파일**: `*CI.pdf` 형식의 파일 선택
   - **패킹리스트 파일**: `*PL.pdf` 형식의 파일 선택
4. 출력 Excel 파일명 확인/수정
5. **"📄 Excel로 변환"** 버튼 클릭
6. 실시간 진행상황 확인
7. 변환 완료 후 파일 자동 열기 옵션

### 주요 특징

- **크로스 플랫폼**: Windows, macOS, Linux 지원
- **반응형 UI**: 창 크기 조절 가능한 유연한 레이아웃
- **상세한 피드백**: 변환 과정의 모든 단계를 실시간으로 표시
- **오류 처리**: 친절한 오류 메시지와 복구 가이드

### 출력 파일

- `{파일명}.xlsx`: 파싱된 데이터가 포함된 Excel 파일
  - **Invoice 시트**: 인보이스 데이터 + 집계 테이블
  - **Packing_List 시트**: 패킹리스트 데이터 + 집계 테이블

## 개발자용

### 요구사항

```bash
pip install -r requirements.txt
```

### 로컬 실행

```bash
python main.py
```

### 빌드

```bash
python build.py
```

## 지원 형식

### 인보이스 필드
- EDI, DeliveryNo, InvoiceNo, InvoiceDate
- ShipmentNo, TotalQuantity
- EanCode, Ref, Ref00, Description
- Quantity, UnitPrice, TotalPriceUsd, Country

### 패킹리스트 필드
- EDI, DeliveryNo, ShipmentNo, Brand
- EAN, REF, REF_00, Description
- Qty, Batch, MfgDate, ExpDate, Dg

## 라이선스

MIT License