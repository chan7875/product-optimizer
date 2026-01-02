# 날짜 범위 최적화 구현 계획

## 목표
사용자가 단일 날짜 대신 **시작 날짜**와 **종료 날짜** 범위를 입력하여, 해당 기간 내의 모든 생산 계획에 대해 일괄적으로 최적화를 수행할 수 있도록 합니다. 결과는 날짜별로 구분하여 출력합니다.

## 사용자 검토 필요 사항
> [!NOTE]
> 결과 화면은 기존의 단일 데이터 테이블 대신 **탭(Tab) 형태**로 변경하여, 각 날짜별 결과 테이블을 별도의 탭에서 확인할 수 있도록 구현할 예정입니다. (예: `2025-12-15` 탭, `2025-12-16` 탭)

## 변경 제안

### GUI 로직 (`main_ui.py`)
#### [MODIFY] [main_ui.py](file:///c:/Develoment/product-optimizer/main_ui.py)
- **Class `OptimizationTab`**:
    - **UI 변경**:
        - 결과 표시 영역(`result_table`)을 `QTabWidget`(`result_tabs`)으로 교체합니다.
    - **Method `run_optimization`**:
        - `QInputDialog` 대신 커스텀 다이얼로그(`DateRangeDialog`)를 띄워 **Start Date**와 **End Date**를 입력받습니다.
        - 시작 날짜부터 종료 날짜까지 반복문을 수행합니다.
        - 각 날짜에 대해:
            1. `ScheduleTab`에서 해당 날짜의 생산 데이터를 가져옵니다.
            2. `Input/item_list.txt`를 생성합니다.
            3. BOM을 병합하여 `Input/BOM.txt`를 생성합니다.
            4. `optimize_plan.py`와 `optimize_sequence.py`를 실행합니다.
            5. `Output/optimization_sequence.csv` 결과를 읽어옵니다.
            6. 결과가 있다면, `result_tabs`에 새로운 탭을 추가하고 테이블을 생성하여 데이터를 표시합니다.
    - **Method `export_result`**:
        - 현재 선택된 탭의 결과만 내보내거나, 모든 탭의 결과를 엑셀 시트별로 내보내도록 수정합니다 (우선 현재 선택된 탭 기준).

## 검증 계획

### 수동 검증
1. 앱 실행 후 `Schedule Management` 탭에서 스케줄 로드.
2. `Production Optimization` 탭 이동.
3. `Run Optimization` 클릭.
4. 날짜 범위 입력 (예: `2025-12-14` ~ `2025-12-15`).
5. 결과 영역에 탭이 2개 생성되는지 확인.
6. 각 탭의 내용이 해당 날짜의 최적화 결과인지 확인.
