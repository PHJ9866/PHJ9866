# Validation Check Excel Mapping Tool

Line List(Master)와 여러 개의 Instrument Datasheet(다중 시트)를 비교해서,
Line No 기준으로 Process Data(운전/설계 압력·온도)가 일치하는지 검증하고
불일치 항목을 하이라이트한 Validation Report Excel을 만들어주는 GUI 도구입니다.

기존 콘솔(`input()`) 기반 매핑 방식 대신, 모든 시트의 컬럼을 자동으로 인식하고
표 형태로 보여준 뒤 필요한 부분만 고쳐서 확정하는 방식으로 동작합니다.

## 설치

```bash
pip install -r requirements.txt
```

Python 표준 라이브러리인 tkinter가 필요합니다. Windows/macOS 공식 설치본에는
기본 포함되어 있고, Linux는 `sudo apt install python3-tk` 등으로 설치할 수 있습니다.

## 실행

```bash
python app.py
```

## 사용 방법

1. **① Line List (Master) 파일**: Line No / Process Data가 있는 Excel 파일을 선택합니다.
2. **② Instrument Datasheet 파일**: 여러 개 선택할 수 있고, 파일 안의 모든 시트를 자동으로 참조합니다.
3. **③ 스캔 & 매핑 확인**: Tag No, Line No와 Process Data 6종 컬럼을 시트별로 자동 인식합니다.
   Line List(Master) 행이 맨 위에, 그 아래로 한 칸 띄운 뒤 Instrument 시트들이 나열됩니다.
   - 초록색 행: 자동/저장된 매핑으로 정상 인식됨
   - 노란색 행: 시트 이름으로 유추한 기본값을 사용 중 (확인 필요)
   - 빨간색 행: 인식 실패 (직접 입력 필요)
   표에서 컬럼 문자(B, AC ...) 위에 마우스를 올리면 실제 헤더/샘플 값이 툴팁으로 바로 보이고,
   행을 더블클릭하면 그 자리에서 직접 컬럼을 입력해 고칠 수 있습니다. 수동으로 고친 컬럼은
   `*` 표시로 구분되고, 인식하지 못한 컬럼은 "매핑 필요"로 표시됩니다(해당 항목이 없는 시트는
   비워두면 됩니다). 상단의 **시트 이름 필터**에 텍스트(예: PE)를 입력하면 엑셀 필터처럼
   일치하는 시트만 화면에 보여줘서 (실제 포함 여부는 바꾸지 않고) 마우스 오버로 미리 확인해볼
   수 있고, 필요하면 "필터에 일치하는 시트만 포함" 버튼으로 그대로 포함 여부에 반영할 수
   있습니다. 필요 없는 시트는 "포함" 칸을 클릭해서 개별적으로 제외할 수도 있습니다.
   확인한 매핑은 `config.json`에 저장되어 다음 실행부터는 같은 시트를 다시 스캔할 필요가 없습니다.
4. **④ Report 생성**: 저장 경로를 지정하면 Validation Report(.xlsx)가 생성됩니다.
   - Line별로 Master 행(회색) + 매칭되는 Instrument 행이 나열됩니다.
   - Process Data가 다르면 빨간색, 같으면 초록색으로 표시됩니다.
   - 매칭되는 Instrument가 없는 Line은 노란색 "No Related Item"으로 표시됩니다.
   - Line(Master) 행의 `Result`는 그 Line에 속한 Instrument 중 하나라도 FAIL이면 FAIL로
     표시됩니다 - 엑셀에서 Result 열을 FAIL로 필터링해도 Line 행이 같이 남아있도록 하기 위함입니다.
   - 맨 끝 `Remark` 열은 색칠 없이 흰 배경으로 생성되어, 검토하면서 메모를 남길 수 있습니다.
   - `Summary` 시트에 전체 통계가 함께 저장됩니다.

## Process Data (6종)

Operating Pressure, Max Design Pressure, Operating Temperature,
Max Operating Temperature, Min Design Temperature, Max Design Temperature

Validation 시트에서는 Line No / Tag No / Type 다음, `P_Oper` → `P_Design_Max` →
`T_Oper` → `T_Oper_Max` → `T_Min_Design` → `T_Max_Design` → Source Sheet → Result → Remark 순으로 배치됩니다.

## 비교 규칙

- Operating Pressure: Master 값이 `10~15` 같은 범위면 Instrument 값이 그 범위 안에 있으면 PASS
- Max Design Pressure / Max Operating Temperature / Min·Max Design Temperature: 허용 오차(0.05) 이내면 PASS
- Operating Temperature: Master 값이 `AMB`면 N/A 처리, 그 외에는 문자열 완전 일치
- 최종 PASS/FAIL은 Operating Pressure, Max Design Pressure, Max Operating Temperature,
  Min/Max Design Temperature 기준으로 판정됩니다 (Operating Temperature는 AMB 표기가 흔해 판정에서 제외)

## 컬럼 자동 인식이 실패하기 쉬운 경우와 대응

- **병합된 그룹 헤더** (예: "Pressure"라는 제목 아래 Operating/Design 컬럼이 나뉘는 경우): 그룹
  라벨이 맨 왼쪽 컬럼에만 저장되는 병합 셀의 특성을 고려해, 헤더 인식 시 그룹 라벨을 오른쪽으로
  전파시켜 모든 하위 컬럼이 인식되도록 처리합니다.
- **헤더 위의 번호 매기기 행** (예: "1 2 3 ..."): 데이터 시작 위치를 숫자처럼 보이는 행 하나가
  아니라 연속된 여러 행으로 판단해서, 번호 행을 실제 데이터로 착각하지 않도록 합니다.
- **설명/타입 컬럼이 우연히 키워드를 포함하는 경우** (예: "Pressure Transmitter"라는 텍스트):
  헤더만으로 판단하지 않고, 실제 데이터가 공정값(숫자 또는 ATM/AMB/F.V 같은 표기)인지도 함께
  확인해서 텍스트 컬럼이나 빈 컬럼을 걸러냅니다. Test/Hydrotest Pressure 컬럼도 제외됩니다.
- **Line No 헤더가 "No"처럼 무의미한 경우**: 키워드로 못 찾으면 데이터 모양(301-ATM-0007처럼
  대시로 구분된 형태)으로 Line No 컬럼을 직접 찾아냅니다.
- **Differential Pressure처럼 별도의 Min/Nor/Max 그룹이 있는 경우**: DP 트랜스미터 자체의
  차압 범위(Differential Pressure)는 제외하고, 실제 공정 압력을 나타내는 "Pressure" 그룹의
  Nor(Normal, 운전값) 컬럼을 Operating Pressure로 사용합니다.
- 그래도 틀리게 잡히면, ③ 매핑 화면에서 해당 행을 더블클릭해 직접 컬럼을 지정하면 됩니다.

## 파일 구성

- `app.py` : GUI (tkinter)
- `engine.py` : 파일 파싱, 컬럼 자동 매핑, 비교 로직, 리포트 생성
- `config.json` : 시트별로 확인된 컬럼 매핑 저장 (자동 생성, 최초 실행 시 없음)
