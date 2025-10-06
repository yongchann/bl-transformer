# PDF Parser GUI - Invoice & Packing List

PDF 파일에서 인보이스와 패킹리스트 데이터를 추출하여 Excel 파일로 변환하는 GUI 도구입니다.

## 기능

- **🖥️ 직관적인 GUI**: 사용하기 쉬운 그래픽 인터페이스
- **📁 파일 선택**: 인보이스와 패킹리스트 파일을 개별적으로 선택
- **📊 실시간 진행상황**: 변환 과정을 실시간으로 확인
- **📋 상세 결과**: 파싱된 데이터의 상세 정보 표시
- **💾 자동 파일명**: 선택한 파일을 기반으로 출력 파일명 자동 생성

## 사용법

### GUI 실행파일 사용 (Windows)

1. `PDF-Parser-GUI.exe` 파일을 다운로드
2. 실행파일을 더블클릭하여 실행
3. GUI에서 파일 선택:
   - **인보이스 파일**: `*CI.pdf` 형식의 파일 선택
   - **패킹리스트 파일**: `*PL.pdf` 형식의 파일 선택
4. 출력 Excel 파일명 확인/수정
5. **"📄 Excel로 변환"** 버튼 클릭
6. 변환 완료 후 파일 열기 선택

### 주요 특징

- **유연한 파일 선택**: 어떤 폴더의 파일이든 선택 가능
- **선택적 처리**: 인보이스만, 패킹리스트만, 또는 둘 다 처리 가능
- **안전한 처리**: 별도 스레드에서 변환 작업 수행으로 UI 응답성 유지

### 출력 파일

- `{파일명}_parsed_data.xlsx`: 파싱된 데이터가 포함된 Excel 파일
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