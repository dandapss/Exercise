import os
import fitz  # PyMuPDF
import re
import openpyxl
from datetime import datetime
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import column_index_from_string

The_King = "LJH_SexyGuy"


## 월을 숫자 → 영문으로 변환
def mon(text):
    """월(MM)을 영문 월(JAN, FEB 등)로 변환"""
    months = {
        "01": "JAN", "02": "FEB", "03": "MAR", "04": "APR", "05": "MAY", "06": "JUN",
        "07": "JUL", "08": "AUG", "09": "SEP", "10": "OCT", "11": "NOV", "12": "DEC"
    }
    return months.get(text, "")

## PO Number page가 달라 못받을 경우 H6에 있는 값 받아오기
def fill_missing_pno(ws):
    """H열(pno)에 값이 없는 경우 기본값 채워 넣기"""
    column_index = column_index_from_string("H")  # H = 8
    first_row = 7  # 데이터는 H6부터 시작

    for row in range(first_row, ws.max_row + 1):
        cell = ws.cell(row=row, column=column_index)
        if cell.value is None or str(cell.value).strip() == "":
            cell.value = "=H6"

## (H열) 마지막 줄 찾기
def get_last_filled_row(ws, column_letter="H"):
    column_index = column_index_from_string(column_letter)  # 'H' => 8
    last_row = 6  # 최소 4행 보장 (H열 4행까지는 무조건 값이 있음)
    
    for row in range(6, ws.max_row + 1):  # 5행부터 검색
        if ws.cell(row=row, column=column_index).value:  # 값이 있는 마지막 행 찾기
            last_row = row
    
    return last_row  # 최소 4행 보장

## 시트별 디자인 구성성
def get_or_create_sheet(wb, sheet_name):
    """Excel 시트를 가져오거나 없으면 새로 생성"""

    sheet_name = sheet_name[:31]
    
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    
    ws = wb.create_sheet(sheet_name)

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
        ws.column_dimensions[col].width = 15
    ws.column_dimensions['H'].width = 35

    # 첫 번째 행: 파이썬 파일 돌린 시간
    ws.merge_cells('B1:H1')
    ws['B1'] = datetime.now()
    ws['B1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['B1'].font = Font(bold=True)

    # 두 번째 행: 뭐 적어야 되는지 모름
    ws.merge_cells('B2:H2')
    ws.row_dimensions[2].height = 30
    ws['B2'].alignment = Alignment(horizontal='center', vertical='center')
    ws['B2'].font = Font(bold=True)
    ws['B2'] = "JUN HYEOK LEE!!!!!! The Master of Logistics"

    # 세 번째 행: 인바운드 & 아웃 바운드
    ws.merge_cells('B3:D3')
    Green_Fill = PatternFill(start_color="00FF00", end_color="FF9999", fill_type="mediumGray")
    ws['B3'].fill = Green_Fill
    ws['B3'].alignment = Alignment(horizontal='center', vertical='center')
    ws['B3'].font = Font(bold=True)
    ws['B3'] = "Inbound"

    ws.merge_cells('F3:H3')
    Red_Fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="mediumGray")
    ws['F3'].fill = Red_Fill
    ws['F3'].alignment = Alignment(horizontal='center', vertical='center')
    ws['F3'].font = Font(bold=True)
    ws['F3'] = "Outbound"
    
    # 네 번째 행: 열 제목(데이터 종류)
    ws.append(["Month", "PO No", "Date", "QTY (MT)", "On Stock", "QTY (MT)", "Date", "PO No"])
    for col in range(1,9):
        cell = ws.cell(row=4, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')        
    
    # 다섯 번째 행: 빈 값
    ws.append([""] * 8)

    return ws

### 색상 추가 부분!! 필요 없을 경우 삭제
def apply_conditional_formatting(ws, max_row):
    """D3 값을 기준으로 D4:D1000 범위에 조건부 서식 적용"""

    # if max_row is None or max_row < 7:
    #     max_row = 100
    
    # 색상 정의
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # 초록색
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 노란색
    orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # 주황색
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # 빨간색

    # 범위 정의
    data_range = f"E6:E{max_row}"

    # thick_border = Border(
    #         left = Side(style="thick"),
    #         right = Side(style="thick"),
    #         top = Side(style="thick"),
    #         bottom = Side(style="thick")
    #     )
    
    # for row in range(6, max_row+1):
    #     cell = ws.cell(row=row, column=5)
    #     cell.border = thick_border

    # 조건부 서식 추가 (D3을 기준으로 계산)
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["E6>=E$5*0.6"], stopIfTrue=True, fill=green_fill)  # 60% 이상 초록색
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["AND(E6>=E$5*0.4, E6<E$5*0.6)"], stopIfTrue=True, fill=yellow_fill)  # 40% 이상 노란색
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["AND(E6>0, E6<E$5*0.4)"], stopIfTrue=True, fill=orange_fill)  # 40% 미만 주황색
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["E6<0"], stopIfTrue=True, fill=red_fill)  # 음수(마이너스 값) 빨간색
    )
    # 0인 경우 색상 없음 (기본값 유지)

    # 셀 테두리 변경
    thick = Side(style="thick")
    thin = Side(style="thin")
    double = Side(style="double")
    medium = Side(style="medium")

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            cell.border = Border(right=double)

    for cell in ws[5]:
        cell.border = Border(top=double)

    ws['A5'].border = Border(right=double, top=double)

    first_row = 4
    last_row = max_row
    column = 5
    for row in range(first_row, last_row+1):
        cell = ws.cell(row=row, column=column)
        cell2 = ws.cell(row=row, column=1)

        if row == first_row:
            cell.border = Border(top=thick, left=thick, right=thick, bottom=None)
            # cell2.border = Border(top=medium, left=medium, right=medium, bottom=None)
        elif row == last_row:
            cell.border = Border(top=None, left=thick, right=thick, bottom=thick)
            # cell2.border = Border(top=None, left=medium, right=medium, bottom=medium)
        else:
            cell.border = Border(top=None, bottom=None, left=thick, right=thick)
            # cell2.border = Border(top=None, bottom=None, left=medium, right=medium)

######################################################
################ 회사별 함수 분류 #####################
######################################################
def process_smp_iberica(text, filename, ws):
    """SMP Ibérica 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []

    while i < len(cleaned_list):
        if cleaned_list[i].strip().startswith("W ") and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]}")
            i+=2
        elif cleaned_list[i].strip().endswith("number/date") and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]}")
            i+=2
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1
    
    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # material = "9120491 ASA LI941 F94484 (LG)"

        if line.startswith("W ") or line.startswith("D "):
            if len(extracted_text) >= 3:
                quantity = extracted_text[2]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = int(quantity)/100000
                    written_date = f"{extracted_text[1][:2]}-{extracted_text[1][2:4]}-{extracted_text[1][4:]}"
                    written_month = f"{mon(extracted_text[1][2:4])}-{extracted_text[1][6:]}"
                    # print(f"[SMP Ibérica] 데이터 추가: {quantity}")
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, "pno"])

        if "number/date" in line:
            if len(extracted_text) >= 4:
                pno = extracted_text[3]
                first_row = 6
                last_row = ws.max_row
                column = 8
                for row in range(first_row, last_row+1):
                    cell = ws.cell(row=row, column=column)
                    cell.value = pno             
       
        
def process_samvardhana(text, filename, ws):
    """Samvardhana Motherson 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []

    while i < len(cleaned_list):
        if cleaned_list[i].strip().endswith("number/date") and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]}")
            i+=2
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1
    
    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # material = "9122188 ASA LI 941V NEGRO 9B9 (LG)"           
        
        if line.startswith("W ") or line.startswith("D "):
            if len(extracted_text) >= 4:
                quantity = extracted_text[3]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = float(quantity)/100000
                    written_date = f"{extracted_text[2][:2]}-{extracted_text[2][2:4]}-{extracted_text[2][4:]}"
                    written_month = f"{mon(extracted_text[2][2:4])}-{extracted_text[2][6:]}"
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, "pno"])
                    # print(f"[Samvardhana Motherson] 데이터 추가: {quantity}")

        if "number/date" in line:
            if len(extracted_text) >= 4:
                pno = extracted_text[3]
                first_row = 6
                last_row = ws.max_row
                column = 8
                for row in range(first_row, last_row+1):
                    cell = ws.cell(row=row, column=column)
                    cell.value = pno  


def process_samvardhana2(text, filename, ws):
    """Samvardhana Motherson 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []
    while i < len(cleaned_list):
        if cleaned_list[i].strip().startswith("D") and i + 1 < len(cleaned_list):           
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]} {cleaned_list[i+2]} {cleaned_list[i+3]} {cleaned_list[i+4]}")
            i+=5
        elif cleaned_list[i].strip().startswith("Sch Agr No") and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i-1]} {cleaned_list[i]}")
            i+=1
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1
    
    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # material = "LG LI941V_V94841_ASA"       
        
        if line.startswith("W ") or line.startswith("D"):
            if "Date" not in line:
                if len(extracted_text) >= 5:
                    quantity = extracted_text[3]
                    if quantity.isdigit() and int(quantity) > 0:
                        qty = float(quantity)/100000
                        written_date = f"{extracted_text[1][:2]}-{extracted_text[1][2:4]}-{extracted_text[1][4:]}"
                        written_month = f"{mon(extracted_text[1][2:4])}-{extracted_text[1][6:]}"
                        ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date])
                        # print(f"[Samvardhana Motherson] 데이터 추가: {quantity}")
        
        if "Sch Agr No" in line:
            if len(extracted_text) >= 4:
                pno = extracted_text[0]
                first_row = 6
                last_row = ws.max_row
                column = 8
                for row in range(first_row, last_row+1):
                    cell = ws.cell(row=row, column=column)
                    cell.value = pno  

# "Blanket P/O #" 기준으로 PO No. 추출하였으나 동일 회사 다른 파일에 해당 부분이 같음.. 재 확인 필요.
# "9) 11.02.2025 Forecast"  &  "homecmsdat96LG0045"
def process_PLASTICOS(text, filename, ws):
    """Samvardhana Motherson 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]

    for line in cleaned_list:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # print(f"@@@@@@@@@@@@@@@line {line}")
        # print(f"@@@@@@@@@@@@@@@extracted_text {extracted_text}")

        if "P/O" in line and "#:" in line:
            if int(extracted_text[7]) > 0:
                pno = extracted_text[7]
                first_row = 6
                last_row = ws.max_row
                column = 8                
                for row in range(first_row, last_row+1):
                    cell = ws.cell(row=row, column=column)
                    cell.value = pno
        
        if "Week" in line and len(extracted_text) > 5:
            quantity = extracted_text[3]
            if quantity.isdigit() and int(quantity) > 0:
                qty = float(quantity)/100000
                for_date = extracted_text[0].split("/")
                if for_date[0] != "RELEASES":
                    if len((for_date)[0]) < 2:
                        for_date[0] = f"0{for_date[0]}"
                    written_date = f"{for_date[0]}-{for_date[1]}-{for_date[2]}"
                    written_month = f"{mon(for_date[1])}-{for_date[2]}"
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, pno])
                    # print(f"[Samvardhana Motherson] 데이터 추가: {quantity}")
            else:
                print("숫자로 변환 불가")        
            
        if "Month" in line and len(extracted_text) > 5:
            quantity = extracted_text[3]
            if quantity.isdigit() and int(quantity) > 0:
                qty = float(quantity)/100000
                for_date = extracted_text[0].split("/")
                if len((for_date)[0]) < 2:
                    for_date[0] = f"0{for_date[0]}"
                written_date = f"{for_date[0]}-{for_date[1]}-{for_date[2]}"
                written_month = f"{mon(for_date[1])}-{for_date[2]}"
                ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, pno])
                # print(f"[Samvardhana Motherson] 데이터 추가: {quantity}")
            else:
                print("숫자로 변환 불가")

        # date1 to date2 Floating Forecast 나왔을 경우 date2를 표시. >> for_date = extracted_text[2].split("/")
        if "Floating" in line and len(extracted_text) > 8:
            if "Forecast" in line:
                quantity = extracted_text[5]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = float(quantity)/100000
                    for_date = extracted_text[2].split("/")
                    if len((for_date)[0]) < 2:
                        for_date[0] = f"0{for_date[0]}"
                    written_date = f"{for_date[0]}-{for_date[1]}-{for_date[2]}"
                    written_month = f"{mon(for_date[1])}-{for_date[2]}"
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, pno])
                    # print(f"[Samvardhana Motherson] 데이터 추가: {quantity}")
            else:
                print("숫자로 변환 불가")                        


def process_OGGIONNI(text, filename, ws):
    """Samvardhana Motherson 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []
    while i < len(cleaned_list):
        if cleaned_list[i].strip().startswith("Date") and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i-1]} {cleaned_list[i]}")
            i+=1
        elif cleaned_list[i].strip().startswith("kg") and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i-4]} {cleaned_list[i]}")
            i+=1
        elif cleaned_list[i].strip().startswith("PO Number") and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i-1]} {cleaned_list[i]}")
            i+=1
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1

    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]

        if line.endswith("Date") and len(extracted_text) >= 2 and "Delivery" not in line:
            written_date = f"{extracted_text[0][:2]}-{extracted_text[0][2:4]}-{extracted_text[0][4:]}"
            written_month = f"{mon(extracted_text[0][2:4])}-{extracted_text[0][6:]}"
        
        if line.endswith("Number") and "PO" in line:
            if int(extracted_text[0]) > 0:
                pno = extracted_text[0]
        
        if line.endswith("kg") and len(extracted_text) >= 2:
            quantity = extracted_text[0]
            if quantity.isdigit() and int(quantity) > 0:
                qty = float(quantity)/1000
            else:
                print("숫자로 변환 불가")
        
            ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, pno])    


def process_NIRAKSV(text, filename, ws):
    """N.I.R.A. KSV 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []
    while i < len(cleaned_list):
        if "ABS LG ER400 M95007" in cleaned_list[i].strip():
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+2]} {cleaned_list[i+5]} {cleaned_list[i+6]}")
            i+=6
        elif "Bestelldatum" in cleaned_list[i].strip():
            merge_lines.append(f"{cleaned_list[i-1]} {cleaned_list[i]} {cleaned_list[i+1]}")
            i+=2
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1
    
    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # material = "ABS LG ER400 M95007 schwarz"
        
        if line.startswith("ABS LG ER400"):
            if len(extracted_text) >= 5:
                quantity = extracted_text[5]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = float(quantity)/100000
                    written_date = f"{extracted_text[6][:2]}-{extracted_text[6][2:4]}-{extracted_text[6][4:]}"
                    written_month = f"{mon(extracted_text[6][2:4])}-{extracted_text[6][4:]}"
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date])
                    # print(f"[Samvardhana Motherson] 데이터 추가: {quantity}")
        
        if "Bestelldatum" in line:
            if len(extracted_text) >= 3:
                pno = extracted_text[0]

                first_row = 6
                last_row = ws.max_row
                column = 8
                for row in range(first_row, last_row+1):
                    cell = ws.cell(row=row, column=column)
                    cell.value = pno  

###### 조건을 이해 1도 못함.. 준혁이형한테 물어보고 작업 진행 ###### [   1) 18.02.2025 104542 WK08   ]
# 날짜 안보임
# PO No. 모르겠음
# Quantity 가로 정렬된 숫자 뽑아내야 하는지?
def process_MaflowPlastic(text, filename, ws):
    """Maflow Plastic Poland 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []
    while i < len(cleaned_list):
        if "0075A00054100GR" in cleaned_list[i].strip():
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+2]} {cleaned_list[i+5]} {cleaned_list[i+6]}")
            i+=6
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1
    
    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # material = "ABS LG ER400 M95007 schwarz"
        print(f"@@@@@@@@@@@@@@@line {line}")
        print(f"@@@@@@@@@@@@@@@extracted_text {extracted_text}")
        
        if line.startswith("ABS LG ER400"):
            if len(extracted_text) >= 5:
                quantity = extracted_text[5]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = float(quantity)/100000
                    written_date = f"{extracted_text[6][:2]}-{extracted_text[6][2:4]}-{extracted_text[6][4:]}"
                    written_month = f"{mon(extracted_text[6][2:4])}-{extracted_text[6][4:]}"
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date])
                    # print(f"[Samvardhana Motherson] 데이터 추가: {quantity}")
        
        if "Bestelldatum" in line:
            if len(extracted_text) >= 3:
                pno = extracted_text[0]

                first_row = 6
                last_row = ws.max_row
                column = 8
                for row in range(first_row, last_row+1):
                    cell = ws.cell(row=row, column=column)
                    cell.value = pno  


def process_ITWSlovakia(text, filename, ws):
    """Maflow Plastic Poland 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []
    while i < len(cleaned_list):
        if cleaned_list[i].strip().endswith(" P") and i + 1 < len(cleaned_list):
            if cleaned_list[i+4].strip().endswith(" P") and i + 1 < len(cleaned_list):
                merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]} {cleaned_list[i+2]} {cleaned_list[i+3]} Seob4")
                i+=4
            elif cleaned_list[i+4].strip().endswith("Sincerly"):
                merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]} {cleaned_list[i+2]} {cleaned_list[i+3]} Seob4")
                i+=4
            else:
                merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]} {cleaned_list[i+2]} {cleaned_list[i+3]} {cleaned_list[i+4]} Seob5")
                i+=5
        elif cleaned_list[i].strip().startswith("Number:") and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]}")
            i+=2
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1
    
    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # material = "ABS LG ER400 M95007 schwarz"

        if "Number:" in line:
            if len(extracted_text) >= 2:
                pno = (f"{extracted_text[1]} {extracted_text[2]}")
                first_row = 6
                last_row = ws.max_row
                column = 8
                for row in range(first_row, last_row+1):
                    cell = ws.cell(row=row, column=column)
                    cell.value = pno 
        
        if line.endswith("Seob4"):
            if len(extracted_text) >= 5:
                quantity = extracted_text[3]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = float(quantity)/100000
                    if len(extracted_text[0]) == 5:
                        for_date = f"0{extracted_text[0]}"
                    else:
                        for_date = extracted_text[0]
                    print(f"@@@ {for_date}")
                    written_date = f"{for_date[:2]}-{for_date[2:4]}-{for_date[4:]}"
                    written_month = f"{mon(for_date[2:4])}-{for_date[4:]}"
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, pno])
                    # print(f"[ITWSlovakia] 데이터 추가: {quantity}")

        if line.endswith("Seob5"):
            if len(extracted_text) >= 6:
                quantity = extracted_text[4]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = float(quantity)/100000
                    if len(extracted_text[0]) == 5:
                        for_date = f"0{extracted_text[0]}"
                    else:
                        for_date = extracted_text[0]
                    written_date = f"{for_date[:2]}-{for_date[2:4]}-{for_date[4:]}"
                    written_month = f"{mon(for_date[2:4])}-{for_date[4:]}"
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, pno])
                    # print(f"[ITWSlovakia] 데이터 추가: {quantity}")


###### PO No. 부분 어디인지 물어보기@@@@@@@@ [    4) 27.03.2025 ZA00820_1906033_LG EU    ]
def process_BoryszewKunststofftechnik(text, filename, ws):
    """Maflow Plastic Poland 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []
    while i < len(cleaned_list):
        if cleaned_list[i].strip().endswith("Gesamtpreis") and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+3]} {cleaned_list[i+6]} Seob_Gesamtpreis")
            i+=6
        elif "Liefertermin:" in cleaned_list[i] and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]}")
            i+=2
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1
    
    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # material = "ABS LG ER400 M95007 schwarz"
        print(f"@@@@@@@@@@@@@@@line {line}")
        print(f"@@@@@@@@@@@@@@@extracted_text {extracted_text}")

        if "Number:" in line:
            if len(extracted_text) >= 2:
                pno = (f"{extracted_text[1]} {extracted_text[2]}")
                first_row = 6
                last_row = ws.max_row
                column = 8
                for row in range(first_row, last_row+1):
                    cell = ws.cell(row=row, column=column)
                    cell.value = pno 

        if "Liefertermin:" in line:
            for_date = extracted_text[1]
            written_date = f"{for_date[:2]}-{for_date[2:4]}-{for_date[4:]}"
            written_month = f"{mon(for_date[2:4])}-{for_date[4:]}"
            for row in range(6, ws.max_row + 1):
                    cell = ws.cell(row=row, column=7)
                    cell2 = ws.cell(row=row, column=1)
                    cell.value = written_date 
                    cell2.value = written_month

       
        if "Seob_Gesamtpreis" in line:
            if len(extracted_text) >= 4:
                quantity = extracted_text[2]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = float(quantity)/1000
                    ws.append(["written_month", The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, "written_date"])
                    # print(f"[ITWSlovakia] 데이터 추가: {quantity}")         


# 동일 회사 다른 PO No.를 가진 파일을 동시에 돌려야 할때가 있는지. 
# "15) 15.01.2025 Bestellung Nr. 7000533558 (003)"   &   "8) 05.03.2025 Bestellung Nr. 7000540404"
# Quantity 빼고는 처음 받은 변수를 저장하게 되어 있음.
# 해당 조건 추가 필요시 모든 코드 수정 필요.
# 04.08.25 추가 임시 저장용 변수 추가 및 하기에서 내용 확인 검증 진행.
def process_ProXAutomotiveAG(text, filename, ws):
    """ProX Automotive AG 문서를 처리하는 함수"""
    print(f"📄 Processing file: {filename}")

    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    merge_lines = []

    # 임시 저장용 변수들
    current_date = ""
    current_month = ""
    current_pno = ""

    i = 0
    while i < len(cleaned_list):
        if "Bestellnummer / Datum" in cleaned_list[i] and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]} Seob_Bestellnummer")
            i += 2
        elif "LG ASA LI941" in cleaned_list[i] and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]} Seob_qty")
            i += 2
        else:
            merge_lines.append(cleaned_list[i])
            i += 1

    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]

        if "Seob_Bestellnummer" in line and len(extracted_text) >= 8:
            for_date = extracted_text[7]
            current_pno = extracted_text[5]
            current_date = f"{for_date[:2]}-{for_date[2:4]}-{for_date[4:]}"
            current_month = f"{mon(for_date[2:4])}-{for_date[4:]}"
            print(f"📌 날짜/PNO 저장: {current_date}, {current_pno}")

        if "Seob_qty" in line and len(extracted_text) >= 9:
            quantity = extracted_text[7]
            if quantity.isdigit() and int(quantity) > 0:
                qty = float(quantity) / 1000000
                if current_date and current_pno:
                    print(f"✅ 저장된 정보 기반 데이터 추가: {qty}, {current_date}, {current_pno}")
                    ws.append([current_month, filename, datetime.now().strftime("%Y-%m-%d"), filename, "On Stock", qty, current_date, current_pno])
                else:
                    print("❌ 날짜나 PNO 정보가 누락됨 → 건너뜀")


def process_FinkeAnwendungstechnik(text, filename, ws):
    """Maflow Plastic Poland 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []
    while i < len(cleaned_list):
        if "Bestellung Nr:" in cleaned_list[i].strip() and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]} Seob_BestellungNr")
            i+=2
        elif "ABS XR 401 BK 9001" in cleaned_list[i].strip() and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]} Seob_qty")
            i+=2
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1
    
    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # material = "ABS LG ER400 M95007 schwarz"
        print(f"@@@@@@@@@@@@@@@line {line}")
        print(f"@@@@@@@@@@@@@@@extracted_text {extracted_text}")

        if "Seob_BestellungNr" in line:
            for_date = extracted_text[4]
            pno = extracted_text[2]
            written_date = f"{for_date[:2]}-{for_date[2:4]}-{for_date[4:]}"
            written_month = f"{mon(for_date[2:4])}-{for_date[4:]}"
            for row in range(6, ws.max_row + 1):
                    cell = ws.cell(row=row, column=7)
                    cell2 = ws.cell(row=row, column=1)
                    cell_pno = ws.cell(row=row, column=8)
                    cell.value = written_date 
                    cell2.value = written_month
                    cell_pno.value = pno
       
        if "Seob_qty" in line:
            if len(extracted_text) >= 7:
                quantity = extracted_text[6]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = float(quantity)/100000
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, pno])
                    # print(f"[ITWSlovakia] 데이터 추가: {quantity}")


def process_FormzeugGmbH(text, filename, ws):
    """Formzeug GmbH 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    i = 0
    merge_lines = []
    while i < len(cleaned_list):
        if "Bestellung " in cleaned_list[i].strip() and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]} Seob_BestellungNr")
            i+=2
        elif "ASA LI941-F94484" in cleaned_list[i].strip() and i + 1 < len(cleaned_list):
            merge_lines.append(f"{cleaned_list[i-2]} {cleaned_list[i]} Seob_qty")
            i+=1
        else:
            merge_lines.append(str(cleaned_list[i]))
            i += 1
    
    for line in merge_lines:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        # material = "ABS LG ER400 M95007 schwarz"
        # print(f"@@@@@@@@@@@@@@@line {line}")
        # print(f"@@@@@@@@@@@@@@@extracted_text {extracted_text}")

        if "Seob_BestellungNr" in line:
            for_date = extracted_text[2]
            pno = extracted_text[1]
            written_date = f"{for_date[:2]}-{for_date[2:4]}-{for_date[4:]}"
            written_month = f"{mon(for_date[2:4])}-{for_date[4:]}"
            for row in range(6, ws.max_row + 1):
                    cell = ws.cell(row=row, column=7)
                    cell2 = ws.cell(row=row, column=1)
                    cell_pno = ws.cell(row=row, column=8)
                    cell.value = written_date 
                    cell2.value = written_month
                    cell_pno.value = pno
       
        if "Seob_qty" in line:
            if len(extracted_text) >= 5:
                quantity = extracted_text[0]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = float(quantity)/1000
                    ws.append([written_month, The_King, datetime.now().strftime("%Y-%m-%d"), The_King, "On Stock", qty, written_date, pno])
                    # print(f"[ITWSlovakia] 데이터 추가: {quantity}")



#########################################################
################ 회사별 함수 분류 끝 #####################
#########################################################


def extract_info(folder_path, output_excel):
    """폴더 내 모든 PDF를 읽고 키워드별로 처리"""
    extracted_data = []

    # 기존 Excel 파일이 있으면 로드, 없으면 새 파일 생성
    if os.path.exists(output_excel):
        wb = openpyxl.load_workbook(output_excel)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # 기본 생성되는 'Sheet' 삭제

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)

            try:
                doc = fitz.open(file_path)

                for page in doc:
                    text = page.get_text("text")
                    print(f"📄 {filename} - 페이지 텍스트 읽음")
                    print(f"@@@@ Text 임: {text}")

                    if "SMP Ibérica" in text:
                        ws = get_or_create_sheet(wb, "SMP Ibérica")
                        process_smp_iberica(text, filename, ws)

                    elif "Samvardhana Motherson Peguform" in text and "9122188 ASA LI 941V NEGRO 9B9 (LG)" in text:
                        ws = get_or_create_sheet(wb, "Samvardhana Motherson Peguform")
                        process_samvardhana(text, filename, ws)

                    elif "Samvardhana Motherson Innovative" in text:
                        ws = get_or_create_sheet(wb, "Samvardhana Motherson Innovative")
                        process_samvardhana2(text, filename, ws)

                    elif "OGGIONNI" in text and "5335630000" in text:
                        ws = get_or_create_sheet(wb, "OGGIONNI")
                        process_OGGIONNI(text, filename, ws)

                    elif "PLASTICOS ABC SPAIN" in text:
                        ws = get_or_create_sheet(wb, "PLASTICOS")
                        process_PLASTICOS(text, filename, ws)

                    elif "N.I.R.A. KSV" in text:
                        ws = get_or_create_sheet(wb, "N.I.R.A. KSV")
                        process_NIRAKSV(text, filename, ws)
                        
                    elif "Maflow Plastics Poland" in text:
                        ws = get_or_create_sheet(wb, "Maflow Plastics Poland")
                        process_NIRAKSV(text, filename, ws)

                    elif "ITW Slovakia" in text:
                        ws = get_or_create_sheet(wb, "ITW Slovakia")
                        process_ITWSlovakia(text, filename, ws)

                    elif "Boryszew Kunststofftechnik" in text:
                        ws = get_or_create_sheet(wb, "Boryszew")
                        process_BoryszewKunststofftechnik(text, filename, ws)

                    elif "Pro-X Automotive AG" in text:
                        ws = get_or_create_sheet(wb, "Pro-X Automotive AG")
                        process_ProXAutomotiveAG(text, filename, ws)        

                    elif "Finke Anwendungstechnik" in text:
                        ws = get_or_create_sheet(wb, "Finke Anwendungstechnik")
                        process_FinkeAnwendungstechnik(text, filename, ws)

                    elif "Formzeug GmbH" in text:
                        ws = get_or_create_sheet(wb, "Formzeug GmbH")
                        process_FormzeugGmbH(text, filename, ws)
                        
                                    
                            
                    else:
                        print(f"⚠️ {filename}: 지정된 키워드 없음. 스킵.")

            except Exception as e:
                print(f"❌ {filename} 처리 중 오류 발생: {e}")


    for sheet in wb.sheetnames:
        ws = wb[sheet]

        max_row_f = ws.max_row
        for row in range(max_row_f, 0, -1):
            if ws[f'F{row}'].value is not None:
                last_row_f = row

                ############################
                ## 값 못 받아오는 부분 때문에 임시 코드
                ############################
                if last_row_f < 6:
                    last_row_f = 6
                break

        for row in range(6, last_row_f + 1):
            F_value = f"F{row}"
            if F_value:
                ws[f"E{row}"] = f"=E{row-1}-F{row}"
        
        ## PO Number page가 달라 못받을 경우 H6에 있는 값 받아오기
        fill_missing_pno(ws)

        ## 색상 추가!! 필요 없을 경우 아래 한줄만 삭제
        apply_conditional_formatting(ws, last_row_f)  # 각 시트에 조건부 서식 적용

        ws.freeze_panes = 'B5'

    if os.path.exists(output_excel):
        os.remove(output_excel)
    wb.save(output_excel)
    print(f"✅ 함수 추가 및 엑셀 파일 저장 완료: {output_excel}")
       

# 실행
folder_path = r"C:\Users\82109\Desktop\개인\Python Test"
output_excel = os.path.join(folder_path, f'{datetime.now().strftime("%Y-%m-%d")}.xlsx')
datetime.now().strftime("%Y-%m-%d")

extract_info(folder_path, output_excel)
