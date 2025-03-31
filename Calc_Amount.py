import os
import fitz  # PyMuPDF
import re
import openpyxl
from datetime import datetime
from openpyxl.styles import Font, Alignment
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule


def mon(text):
    """월(MM)을 영문 월(JAN, FEB 등)로 변환"""
    months = {
        "01": "JAN", "02": "FEB", "03": "MAR", "04": "APR", "05": "MAY", "06": "JUN",
        "07": "JUL", "08": "AUG", "09": "SEP", "10": "OCT", "11": "NOV", "12": "DEC"
    }
    return months.get(text, "")


def get_or_create_sheet(wb, sheet_name):
    """Excel 시트를 가져오거나 없으면 새로 생성"""
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    
    ws = wb.create_sheet(sheet_name)

    # 첫 번째 행: 파이썬 파일 돌린 시간
    ws.merge_cells('B1:H1')
    ws['B1'] = datetime.now()
    ws['B1'].alignment = Alignment(horizontal='center', vertical='bottom')
    ws['B1'].font = Font(bold=True)

    # 두 번째 행: 뭐 적어야 되는지 모름
    ws.append("Lee Jun Hyeok the Master of Logistics")

    # 세 번째 행: 인바운드 & 아웃 바운드
    ws.merge_cells('B3:D3')
    ws.merge_cells('F3:H3')
    
    # 네 번째 행: 열 제목(데이터 종류)
    ws.append(["Month", "PO No", "Date", "QTY (MT)", "On Stock", "QTY (MT)", "Date", "PO No"])
    ws['B4'].font = Font(bold=True)
    
    # 다섯 번째 행: 빈 값
    ws.append([""] * 8)

    return ws


def process_smp_iberica(text, filename, ws):
    """SMP Ibérica 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    
    for line in cleaned_list:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        
        if line.startswith("W ") or line.startswith("D "):
            if len(extracted_text) >= 3:
                quantity = extracted_text[2] if len(extracted_text) == 3 else extracted_text[3]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = int(quantity)/100
                    written_date = f"{extracted_text[1][:2]}-{extracted_text[1][2:4]}-{extracted_text[1][4:]}"
                    written_month = f"{mon(extracted_text[1][2:4])}-{extracted_text[1][6:]}"
                    ws.append([written_month, filename, datetime.now().strftime("%Y-%m-%d"), "", "On Stock", qty, written_date, "PO No"])
                    print(f"[SMP Ibérica] 데이터 추가: {quantity}")


def process_samvardhana(text, filename, ws):
    """Samvardhana Motherson 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    
    for line in cleaned_list:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        
        if line.startswith("W ") or line.startswith("D "):
            if len(extracted_text) >= 4:
                quantity = extracted_text[3]
                if quantity.isdigit() and int(quantity) > 0:
                    qty = int(quantity)/100
                    written_date = f"{extracted_text[2][:2]}-{extracted_text[2][2:4]}-{extracted_text[2][4:]}"
                    written_month = f"{mon(extracted_text[2][2:4])}-{extracted_text[2][6:]}"
                    ws.append([written_month, filename, datetime.now().strftime("%Y-%m-%d"), "", "On Stock", qty, written_date, "PO Num"])
                    print(f"[Samvardhana Motherson] 데이터 추가: {quantity}")


### 색상 추가 부분!! 필요 없을 경우 삭제
def apply_conditional_formatting(ws, max_row):
    """D3 값을 기준으로 D4:D1000 범위에 조건부 서식 적용"""
    
    # 색상 정의
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # 초록색
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 노란색
    orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # 주황색
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # 빨간색

    # 범위 정의
    data_range = f"E7:E{max_row}"

    # 조건부 서식 추가 (D3을 기준으로 계산)
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["E4>=E$3*0.6"], stopIfTrue=True, fill=green_fill)  # 60% 이상 초록색
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["AND(E4>=E$3*0.4, E4<E$3*0.6)"], stopIfTrue=True, fill=yellow_fill)  # 40% 이상 노란색
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["AND(E4>0, E4<E$3*0.4)"], stopIfTrue=True, fill=orange_fill)  # 40% 미만 주황색
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["E4<0"], stopIfTrue=True, fill=red_fill)  # 음수(마이너스 값) 빨간색
    )
    # 0인 경우 색상 없음 (기본값 유지)


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

                    if "SMP Ibérica" in text:
                        ws = get_or_create_sheet(wb, "SMP Ibérica")
                        process_smp_iberica(text, filename, ws)

                    elif "Samvardhana Motherson" in text:
                        ws = get_or_create_sheet(wb, "Samvardhana Motherson")
                        process_samvardhana(text, filename, ws)

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
                break

        for row in range(7, last_row_f + 1):
            F_value = f"F{row}"
            if F_value:
                ws[f"E{row}"] = f"=E{row-1}-F{row}"

        ## 색상 추가!! 필요 없을 경우 아래 한줄만 삭제
        apply_conditional_formatting(ws, last_row_f)  # 각 시트에 조건부 서식 적용

    wb.save(output_excel)
    print(f"✅ 함수 추가 완료: {output_excel}")

    print(f"✅ 엑셀 파일 저장 완료: {output_excel}")



# 실행
folder_path = r"C:\Users\82109\Desktop\개인\Python Test"
output_excel = os.path.join(folder_path, f'{datetime.now().strftime("%Y-%m-%d")}.xlsx')

extract_info(folder_path, output_excel)
