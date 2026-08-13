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
   "파일 선택" 버튼 대신 탐색기에서 파일을 끌어다 놓아도 됩니다.
2. **② Instrument Datasheet 파일**: 여러 개 선택할 수 있고, 파일 안의 모든 시트를 자동으로 참조합니다.
   여러 개를 한 번에 끌어다 놓아도 됩니다.
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
   있습니다. 필요 없는 시트는 "포함" 칸을 클릭해서 개별적으로 제외할 수도 있습니다. 매핑 수정은
   시트 하나하나 독립적으로 적용됩니다 - " - PE"/" - HE"/" - BU"처럼 종류가 같은 시트라도 컬럼
   레이아웃이 다를 수 있으니, 한 시트를 고쳐도 다른 시트는 바뀌지 않습니다.
   확인한 매핑과 포함/제외 여부는 시트 이름별로 `validation_mapping_config.json`에 저장되어,
   다음 실행 때 Instrument 파일 이름이 바뀌어도(예: 파일명 끝 날짜가 매일 달라지는 경우) 시트
   이름이 같으면 그대로 적용되고 다시 물어보지 않습니다.
   스캔을 시작하면 이전에 저장된 매핑 기록이 있는지, 있다면 몇 개 시트가 저장되어 있고
   마지막으로 언제 저장됐는지 진행 로그에 표시됩니다. 매핑을 확인/저장할 때는 이번에
   적용된 매핑이 이전 기록과 비교해서 달라진 점이 있는지(새로 추가된 시트, 컬럼이 바뀐 항목,
   포함/제외 여부 변경) 시트별로 진행 로그에 나열되고, 달라진 게 없으면 "이전 매핑과 동일함"으로
   표시됩니다.
4. **④ Report 생성**: 저장 경로를 지정하면 Validation Report(.xlsx)가 생성됩니다.
   - Line별로 Master 행(회색) + 매칭되는 Instrument 행이 나열됩니다.
   - Process Data가 다르면 빨간색, 같으면 초록색으로 표시됩니다.
   - 매칭되는 Instrument가 없는 Line은 색칠 없이 "No Related Item"으로 표시됩니다.
   - Line(Master) 행의 `Result`는 그 Line에 속한 Instrument 중 하나라도 FAIL이면 FAIL로
     표시됩니다 - 엑셀에서 Result 열을 FAIL로 필터링해도 Line 행이 같이 남아있도록 하기 위함입니다.
   - 맨 끝 `Remark` 열은 색칠 없이 흰 배경으로 생성되어, 검토하면서 메모를 남길 수 있습니다.
   - 모든 셀에 옅은 테두리가 적용되어 표 형태로 바로 보기 편합니다.
   - `Unmatched Instruments` 시트에 어떤 Line No와도 매칭되지 않은 Instrument Tag들이
     (Tag No / Line No / Source Sheet / Source File과 함께) 나열됩니다.
   - `Result` 다음 `Checked` 열은 검토 완료 여부를 표시하는 드롭다운 칸입니다(칸을 클릭하면
     "V" 선택 가능). Report를 생성할 때마다 바로 이전에 만든 Report 파일 경로가
     `validation_mapping_config.json`에 자동으로 기억되어, 다음번에 새로 Report를 만들면
     같은 Line No/Tag No 조합에 대해 이전 Report에서 체크했던 내용과 `Remark`에 적어둔
     메모가 자동으로 이어받아집니다(진행 로그에 몇 건을 이어받았는지 표시됨). 매칭되는
     Instrument가 없어 "No Related Item"으로 표시된 행도 동일하게 이어받아집니다. 이미
     검토한 항목을 또 검토하는 중복 작업을 막기 위한 기능입니다.

## Process Data (6종)

Operating Pressure, Max Design Pressure, Operating Temperature,
Max Operating Temperature, Min Design Temperature, Max Design Temperature

Validation 시트에서는 Line No / Tag No 다음, `P_Oper` → `P_Design_Max` →
`T_Oper` → `T_Oper_Max` → `T_Min_Design` → `T_Max_Design` → Source Sheet → Result → Remark 순으로 배치됩니다.

## 비교 규칙

- 모든 숫자 비교는 **사실상 완전 일치**를 요구합니다 (부동소수점 오차 정도만 허용, 예:
  4.899999999999999 vs 4.9는 PASS). 3.975 vs 4처럼 실제로 다른 값은 조금만 달라도 FAIL로
  표시됩니다.
- Operating Pressure: Master 값이 `10~15` 같은 범위면 Instrument 값이 그 범위 안에 있으면 PASS
- Operating Temperature: Master 값이 `AMB`면 Instrument 값도 `AMB`일 때만 PASS - 구체적인
  숫자가 들어있으면 FAIL로 표시됩니다 (컬럼을 잘못 매핑했을 가능성이 높다는 신호이기도 합니다)
- 최종 PASS/FAIL은 Process Data 6개 항목 전부를 기준으로 판정됩니다
- Line(Master) 행의 Result는 그 Line에 속한 Instrument 중 하나라도 FAIL이면 FAIL로 표시됩니다

## 컬럼 자동 인식이 실패하기 쉬운 경우와 대응

- **그룹 헤더** (예: "Pressure"/"Design Pressure"라는 제목 아래 Min/Nor/Max 컬럼이 나뉘는 경우):
  실제 병합 셀이면 병합 정보를 그대로 이용하고, 병합 없이 그냥 옆 칸을 비워둔 경우에도 그
  칸이 자신만의 하위 항목(Min/Nor/Max 등)을 갖고 있으면 그룹 제목을 전파시켜 인식합니다
  (완전히 비어있는 진짜 스페이서 칸에는 전파하지 않습니다).
- **헤더 위의 번호 매기기 행** (예: "1 2 3 ..."): 데이터 시작 위치를 숫자처럼 보이는 행 하나가
  아니라 연속된 여러 행으로 판단해서, 번호 행을 실제 데이터로 착각하지 않도록 합니다.
- **설명/타입 컬럼이 우연히 키워드를 포함하는 경우** (예: "Pressure Transmitter"라는 텍스트):
  헤더만으로 판단하지 않고, 실제 데이터가 공정값(숫자 또는 ATM/AMB/F.V 같은 표기)인지도 함께
  확인해서 텍스트 컬럼이나 빈 컬럼을 걸러냅니다. Test/Hydrotest Pressure 컬럼도 제외됩니다.
- **Line No 헤더가 "No"처럼 무의미한 경우**: 키워드로 못 찾으면 데이터 모양(301-ATM-0007처럼
  대시로 구분된 형태)으로 Line No 컬럼을 직접 찾아냅니다.
- **Operating/Design Pressure·Temperature와 이름이 비슷한 다른 항목이 있는 경우**: DP
  트랜스미터 자체의 차압(Differential Pressure), 인라인 기기의 압력 손실(Pressure Drop),
  유체의 임계압력(Critical Pressure)·증기압(Vapor Pressure), Flow/Viscosity/Density 같은
  유체 물성치, PSV·Rupture Disc 전용 값(Set/Back/Built-up Pressure, Overpressure,
  Blowdown 등), Service/Fluid/Description/UOM/Size/Class 같은 설명·단위 컬럼은 전부
  제외하고, 실제 공정 압력·온도를 나타내는 컬럼만 사용합니다.
- **Line No 헤더가 다른 의미의 "Line"을 포함하는 경우**: "Line"과 "No"/"Number"가 함께 있는
  헤더만 인정하고, Service/Size/Class/Fluid/Diameter/UOM/Unit/Description이 포함된 컬럼은
  제외합니다.
- 그래도 틀리게 잡히면, ③ 매핑 화면에서 해당 행을 더블클릭해 직접 컬럼을 지정하면 됩니다.

## 매핑 정보는 어디에 저장되나

`validation_mapping_config.json` 파일이 **Line List(Master) 파일이 있는 폴더**에 자동으로
생성됩니다 (도구 코드가 있는 폴더가 아닙니다). 그래서:

- 이 도구 자체를 업데이트해서 새 폴더에 다시 받아도, Line List가 있는 원래 작업 폴더는
  그대로이므로 이전에 확인한 매핑이 계속 적용됩니다.
- Instrument 파일 이름이 매일 바뀌어도(파일명 끝 날짜 등) 시트 이름만 같으면 다시 물어보지
  않습니다.
- 컬럼 매핑은 " - PE"/" - HE"/" - BU" 같은 접미사를 뗀 시트 종류 단위로 공유되고, 포함/제외
  여부는 시트 이름 전체 단위로 따로 저장됩니다.

## 파일 구성

- `app.py` : GUI (tkinter)
- `engine.py` : 파일 파싱, 컬럼 자동 매핑, 비교 로직, 리포트 생성
- `validation_mapping_config.json` : 시트별 컬럼 매핑 + 포함/제외 여부 저장
  (Line List 파일이 있는 폴더에 자동 생성)
