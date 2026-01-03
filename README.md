# 생산 계획 최적화 프로그램 매뉴얼

이 문서는 개발된 생산 최적화 도구에 대한 포괄적인 가이드를 제공합니다.

## 1. 개요
이 프로그램은 자재 교체를 최소화하여 SMT 생산 순서를 최적화하기 위해 고안된 두 가지 주요 Python 스크립트로 구성됩니다.

-   **`optimize_plan.py`**: BOM 데이터를 분석하여 자재를 분류(공통 vs 개별)하고, 선택적으로 전체 공통 부품 목록을 추출합니다.
-   **`optimize_sequence.py`**: TSP(Traveling Salesperson Problem) 알고리즘을 사용하여 우선순위 항목, 레이어 순서, 자재 변경 등을 고려한 최적의 생산 순서를 계산합니다.

## 2. 필수 조건
-   Python 설치 필요.
-   필수 라이브러리: `ortools`
    ```powershell
    pip install ortools
    ```
-   **입력 파일** (`D:\Develoment\ProductOptimize\Input` 폴더 내):
    -   `BOM.txt`: 자재 명세서(Bill of Materials).
    -   `item_list.txt`: 생산 계획(수량, 시간).
    -   `common_material_list.csv`: 고정 공통 부품 목록 (최적화 계산 시 제외됨).

## 3. 사용법 및 옵션

### A. 분석 및 공통 부품 추출 (`optimize_plan.py`)
BOM 구조를 분석합니다.

| 옵션 | 설명 | 예시 |
| :--- | :--- | :--- |
| **(없음)** | 기존 `common_material_list.csv`를 사용하여 BOM을 분석합니다. | `python optimize_plan.py` |
| `--extract-common` | 모든 항목에 교차하여 사용되는 공통 부품을 추출합니다. | `python optimize_plan.py --extract-common` |
| `--items` | 위 옵션과 함께 사용하여 추출 범위를 제한합니다. | `python optimize_plan.py --extract-common --items "A,B"` |

### B. 순서 최적화 (`optimize_sequence.py`)
최적화된 생산 일정을 생성합니다.

| 옵션 | 설명 | 예시 |
| :--- | :--- | :--- |
| **(없음)** | 모든 작업을 TSP를 사용하여 최적화합니다. | `python optimize_sequence.py` |
| `--priority` | 특정 항목을 우선 생산하도록 지정합니다. | `python optimize_sequence.py --priority "ItemA,ItemB"` |
| `--layer` | 레이어 순서를 강제합니다 (TB=Top→Bottom, BT=Bottom→Top). | `python optimize_sequence.py --layer TB` |
| `--manual` | 사용자가 지정한 수동 순서를 평가합니다. | `python optimize_sequence.py --manual "(ItemA,Top),(ItemB,Bot)"` |

**로직 참고:**
-   **공통 자재 제외**: 프로그램은 `common_material_list.csv`를 로드하여 정확성을 위해 "개별 자재" 수 및 교체 비용 계산에서 이 부품들을 **엄격히 제외**합니다.

### C. 생산 시간 산출 (`calculate_schedule.py`)
엑셀 형태의 생산 계획을 읽어 생산 시간을 계산합니다.

| 옵션 | 설명 | 예시 |
| :--- | :--- | :--- |
| `--file` | 엑셀 파일 경로. (필수) | `python calculate_schedule.py --file "Schedule.xlsx"` |
| `--date` | 생산 시간을 산출할 기준 날짜 (엑셀 헤더와 일치해야 함). | `python calculate_schedule.py ... --date "2024-01-01"` |

**주요 기능:**
-   **시간 계산**: `(CycleTime * Array * 수량 / 60) + 13분(준비시간)` 공식을 각 작업면(Top/Bottom)별로 적용.
-   **결과 출력**: 화면에 총 생산 시간(분, 소수점 첫째자리 반올림)과 가동률 표시.
-   **파일 생성**: `Input/item_list_from_excel.txt` 파일을 생성하여 최적화 프로그램 입력으로 활용 가능.

## 4. 출력 설명 (`optimization_sequence.csv`)

결과 파일은 `D:\Develoment\ProductOptimize\Output\optimization_sequence.csv`에 저장됩니다.

| 열(Column) | 설명 |
| :--- | :--- |
| `Index` | 생산 순서 번호. |
| `Item_Code`, `Layer` | 제품 식별자 및 레이어(면). |
| `Qty` | 생산 수량. |
| `Prod_Time` | 예상 생산 시간. |
| `Total_Count` | 이 레이어의 전체 부품 수. |
| `Common_Count` | 공통 자재 목록에서 발견된 부품 수. |
| `Individual_Count` | 이 모델 고유의 부품 수 (전체 - 공통). |
| `Transition_Shared_Count` | 직전 작업과 공유되는 **개별 부품** 수. (높을수록 효율적). |
| `Selection_Reason` | 이 작업이 이곳에 배치된 이유 설명. |
| `Individual_Materials` | 특정 개별 자재 코드 목록. |

---
**실행 예시 (레이어 우선순위 포함 전체 최적화):**
```powershell
python D:\Develoment\ProductOptimize\optimize_sequence.py --layer TB
```

## 5. 버전 관리 (Git)

이 프로젝트는 Github를 통해 버전 관리가 되고 있습니다.

-   **저장소 주소**: `https://github.com/chan7875/product-optimizer.git`

### 기본 명령어 가이드

1.  **현재 상태 확인**
    ```powershell
    git status
    ```
2.  **변경된 파일 스테이징 (준비)**
    ```powershell
    git add .
    ```
3.  **변경 사항 커밋 (저장)**
    ```powershell
    git commit -m "작업 내용에 대한 메시지"
    ```
4.  **Github에 업로드 (Push)**
    ```powershell
    git push
    ```
5.  **Github에서 다운로드 (Pull)**
    ```powershell
## 6. GUI 프로그램 사용법

콘솔 명령어가 아닌 그래픽 인터페이스를 통해 편리하게 작업을 수행할 수 있습니다.

**실행 방법:**
```powershell
python D:\Develoment\ProductOptimize\main_ui.py
```

### A. 스케줄 관리 탭 (Schedule Management)
생산 계획 엑셀 파일을 로드하여 라인별 생산 시간과 가동율을 확인하고 관리합니다.

1.  **엑셀 로드 (`Load Excel Schedule`)**:
    -   생산 계획이 담긴 엑셀 파일을 불러옵니다.
    -   날짜 형식은 `YYYY-MM-DD` 형태로 자동 변환되어 상단 헤더에 표시됩니다.
    
2.  **화면 구성 (Quad-View)**:
    -   **좌측 상단** (고정): 라인별 요약 라벨 (S01~S04, Util%)이 "연배열" 열에 표시됩니다.
    -   **우측 상단** (고정/요약): 라인별/일자별 총 생산 시간(분)과 가동율(%)을 보여줍니다.
        -   **상세 팝업**: 우측 상단 테이블의 셀을 **더블 클릭**하면, 해당 일자/라인의 상세 생산 내역(아이템, 레이어, 수량, T/T)을 팝업창으로 확인할 수 있습니다.
    -   **좌측 하단** (고정): 품목 코드, 레이어 등 제품 기본 정보가 표시됩니다 (스크롤 고정).
    -   **우측 하단** (메인): 일자별 생산 수량 데이터가 표시됩니다 (자유 스크롤).

3.  **데이터 수정 및 재산출**:
    -   테이블의 수량 데이터를 직접 수정할 수 있습니다.
    -   `Calculate Production Time` 버튼을 누르면 **현재 화면에 입력된 값**을 기준으로 즉시 생산 시간과 요약 정보를 재계산합니다.
    -   이때 팝업되는 결과창은 선택한 날짜의 상세 리포트를 보여줍니다.

4.  **설정 시간 (Setup Time)**:
    -   화면 상단의 입력 필드(S01 ~ S04)에서 라인별 모델 교체 시간(분)을 설정할 수 있습니다. (기본값: 0)
    -   계산 시 이 값이 반영됩니다.

### B. 최적화 수행 탭 (Optimization)
기존의 `optimize_sequence.py` 기능을 GUI에서 수행합니다.

1.  **파일 선택**: BOM, 생산 계획, 공통 부품 리스트 파일을 선택합니다.
2.  **옵션 설정**: 우선순위(Priority), 레이어(Layer) 옵션 등을 입력합니다.
3.  **실행 (`Run Optimization`)**: 최적화 스크립트를 실행하고 결과를 화면에 표시합니다.

### C. 날짜별 최적화 (Date Range Optimization)
**Optimization 탭**에서 특정 날짜 범위에 대해 일괄 최적화를 수행합니다.

1.  **날짜 범위 선택**:
    -   `Run Optimization` 버튼 클릭 시 날짜 범위(Start Date ~ End Date)를 선택하는 팝업이 표시됩니다.
2.  **일괄 실행**:
    -   선택한 기간 내의 모든 날짜에 대해 자동으로 생산 계획을 생성하고 최적화를 수행합니다.
    -   해당 날짜에 데이터가 없는 경우 자동으로 건너뛰거나 빈 결과를 표시합니다.
3.  **결과 확인 (탭 뷰)**:
    -   결과 화면이 탭(Start Date, Start Date+1, ...) 형식으로 표시되어 날짜별 최적화 결과를 쉽게 전환하며 확인할 수 있습니다.
4.  **엑셀 내보내기**:
    -   `Export Result to Excel` 버튼을 클릭하면 **현재 보고 있는 탭**의 데이터를 엑셀로 저장합니다.

### D. SMD 데이터 검증 (SMD Data Verification)
JSON 기반의 SMD/PCB 데이터를 검증하고 CubicSMT 실행을 돕는 탭입니다.

1.  **폴더 로드 (`Load Folder`)**:
    -   최상위 폴더를 선택하면 하위 디렉토리를 재귀적으로 탐색하여 `jsonInfo.txt` 파일을 찾습니다.
    -   기본 경로: `Y:\CadDesign\Manufacture\NW\Design_25`
2.  **데이터 보기 (Split View)**:
    -   **좌측 (트리)**: SMD Code 및 PCB Code 구조를 트리 형태로 표시합니다. 항목 클릭 시 우측 테이블이 필터링됩니다.
        -   **필터 초기화**: 트리 상단 헤더(`SMD Code / PCB Code`)를 클릭하면 전체 목록을 다시 볼 수 있습니다.
    -   **우측 (테이블)**: PCB Code, Rev, SMD Code, BOM/Gerber 파일명 등 상세 정보를 표시합니다.
        -   **BOM File**: 여러 개의 파일이 있는 경우 줄바꿈으로 구분되어 표시됩니다.
3.  **SMD Pro 실행**:
    -   테이블의 `Check` 박스를 선택한 후 상단의 `SMD Pro 실행` 버튼을 클릭합니다.
    -   선택된 항목의 경로에 있는 `jsonInfo.txt`를 인자로 하여 `CubicSMT.exe`를 실행합니다.
4.  **CAD-BOM 체크 리포트 (CAD-BOM Report Viewer)**:
    -   우측 상단 테이블의 행을 **클릭**하면 해당 모델에 대한 검증 리포트를 조회합니다.
    -   **검색 경로**: `L:\CADBomReport`
    -   **파일명 규칙**: `{SMD Code}_{PCB Code}`로 시작하는 엑셀 파일.
    -   **화면 표시**: 우측 하단 탭 영역에 엑셀 파일의 모든 시트 내용을 표시합니다. (1/2 화면 분할)
