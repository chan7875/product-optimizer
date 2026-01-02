# SMD 데이터 검증 탭 구현 계획

## 목표
사용자가 지정한 폴더 내의 하위 폴더들을 검색하여 `jsonInfo.txt` 파일을 찾고, 해당 파일의 내용을 파싱하여 GUI에 표시합니다. UI는 제공된 이미지와 유사하게 좌측의 필터/트리 영역과 우측의 상세 정보 테이블로 구성합니다.

## 사용자 요구사항 분석
1. **폴더 검색**: 하위 폴더 재귀 검색 -> `jsonInfo.txt` 찾기.
2. **데이터 파싱**:
   - `basicInfo`: `pcbCode`, `smdCode` (품목코드), `smdNm` (품목명), `pcbSize`.
   - `cadFileInfo`: `neutralFileNm`, `gerberFileNm`, `matrList` (BOM 파일명 포함).
3. **UI 구성**:
   - **탭 추가**: 메인 윈도우에 새로운 탭 추가.
   - **레이아웃**: Splitter 사용 (좌: 선택/검색, 우: 상세 그리드).
   - **스타일**: 제공된 이미지와 유사한 룩앤필 (아이콘 툴바 등).

## 변경 제안

### 1. Backend Logic
- **Module**: `data_loader.py` (신규 생성 또는 `main_ui.py` 내 클래스)
- **Function**: `scan_directory(path)` -> `List[Dict]`
  - `os.walk`를 사용하여 `jsonInfo.txt` 탐색.
  - `json.load`로 파일 읽기.
  - 필요한 필드 추출 및 구조화.

### 2. GUI Logic (`main_ui.py`)
#### [NEW] Class `SMDVerificationTab(QWidget)`
- **Layout**: `QHBoxLayout` with `QSplitter`.
- **Widgets**:
  - **Left Panel**: `QTreeWidget` or `QListWidget`. (SMD Code 목록 표시)
  - **Right Panel**: `QTableWidget`. (상세 정보 표시)
    - Columns: `PCB Code`, `SMD Code`, `SMD Name`, `PCB Size`, `Neutral File`, `Gerber File`, `BOM File` 등.
  - **Top Toolbar**: "폴더 열기(Load Folder)" 버튼 필요.

#### [MODIFY] Class `MainWindow`
- `SMDVerificationTab`을 탭 위젯의 **첫 번째** 탭으로 추가 (User Request: "Main tab" 추가).

## 검증 계획
1. 더미 `jsonInfo.txt` 파일을 포함한 테스트 폴더 구조 생성.
2. 앱 실행 후 "SMD Data Verification" 탭 선택.
3. "Load Folder" 버튼 클릭 -> 테스트 폴더 선택.
4. 좌측 리스트와 우측 테이블에 데이터가 정상적으로 파싱되어 표시되는지 확인.
