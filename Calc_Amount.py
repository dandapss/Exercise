import os
import fitz
import pandas as pd
import re
import pdfplumber
from datetime import datetime
from openpyxl import Workbook, load_workbook


def mon(text):
    months = {"01": "JAN", "02": "FEB", "03": "MAR", "04": "APR", "05": "MAY", "06": "JUN", 
              "07": "JUL", "08": "AUG", "09": "SEP", "10": "OCT", "11": "NOV", "12": "DEC"}
    return months.get(text, "")

def sheet_name(wb, text):
    if text in wb.sheetnames:
        return wb[text]
    else:
        ws = wb.create_sheet(text)
        return ws

def extract_info(folder_path,output_excel):
    extracted_data = []
    wb = load_workbook(output_excel)
    sheet_title = "Default"
    #################################################
    # wb 사용할수 있도록 variable 이든 변수든 뭐든 설정 해야함. 그러면 sheet 가능할듯듯

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)

            try:
                doc = fitz.open(file_path)
                for page in doc:
                    text = page.get_text("text")
                    print(f"This is extracted text: {text}")

                    if "SMP Ibérica" in text:
                        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                        if sheet_title == "Default" or sheet_title.startswith("Sheet"):
                            print("phase #1")
                            sheet_title = "SMP Ibérica"
                            
                        elif sheet_title == "SMP Ibérica":
                            print("phase #2")
                        else:
                            print("phase #3")
                            sheet_name(wb, "SMP Ibérica")

                        lines = text.replace(",", "").replace(".","").split("\n")
                        print("lineslineslineslineslineslineslineslineslineslineslineslineslineslineslines")
                        # 아래에서 clean이 안된상태로 나와 다시 strip 해줘야함.. 왜 필요?
                        cleaned_list = [item for item in lines if item and str(item).strip()]
                        i = 0
                        merge_lines = []
                        while i < len(cleaned_list):
                            if cleaned_list[i].strip().startswith("W ") and len(cleaned_list[i]) == 2:
                                merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]}")
                                i+=2
                            else:
                                merge_lines.append(str(cleaned_list[i]))
                                i += 1
                        print("cleaned_listcleaned_listcleaned_listcleaned_listcleaned_listcleaned_listcleaned_listcleaned_listcleaned_list")
                        for line in merge_lines:
                            extracted_texts = re.split(r'[ ]', line)
                            extracted_text = [item for item in extracted_texts if item and str(item).strip()]
                            print(f"@@@@@@@@@@@@@@@@@@@@@@ {extracted_text}")

                            if line.startswith("W ") or line.startswith("D "):
                                if len(extracted_text) == 3:
                                    if int(extracted_text[2]) > 0:
                                        written_date = (f"{extracted_text[1][:2]}-{extracted_text[1][2:4]}-{extracted_text[1][4:]}")
                                        written_month = (f"{mon(extracted_text[1][2:4])}-{extracted_text[1][6:]}")
                                        extracted_data.append([filename, written_month, Datetime, "Whole Number", "On Stock", extracted_text[2], written_date, "PO No"])
                                        print(f"%%%%%%%%%%%%%% {extracted_text[2]}")
                                elif len(extracted_text) > 3:
                                    if int(extracted_text[3]) > 0:
                                        written_date = (f"{extracted_text[2][:2]}-{extracted_text[2][2:4]}-{extracted_text[2][4:]}")
                                        written_month = (f"{mon(extracted_text[2][2:4])}-{extracted_text[2][6:]}")
                                        extracted_data.append([filename,written_month, Datetime, "Whole Number", "On Stock", extracted_text[3], written_date, "PO No"])
                                        print(f"%%%%%%%%%%%%%% {extracted_text[3]}")




                    elif "Samvardhana Motherson Peguform" in text:
                        print("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&")
                        if sheet_title == "Default" or sheet_title.startswith("Sheet"):
                            sheet_title = "Samvardhana Motherson Peguform"
                        elif sheet_title.startswith("Samvardhana"):
                            print("Stay")
                        else:
                            print("phase #3")
                            ############# 03.31.2025 요 아래가 안돈다!!!!
                            print(wb.sheetnames)
                            sheet_name(wb, "Samvardhana")
                        
                        lines = text.replace(",", "").replace(".","").split("\n")
                        # 아래에서 clean이 안된상태로 나와 다시 strip 해줘야함.. 왜 필요?
                        cleaned_list = [item for item in lines if item and str(item).strip()]
                        i = 0
                        merge_lines = []
                        while i < len(cleaned_list):
                            if cleaned_list[i].strip().startswith("W ") and len(cleaned_list[i]) == 2:
                                merge_lines.append(f"{cleaned_list[i]}{cleaned_list[i+1]}")
                                i+=2
                            else:
                                merge_lines.append(str(cleaned_list[i]))
                                i += 1

                        for line in merge_lines:
                            extracted_texts = re.split(r'[ ]', line)
                            extracted_text = [item for item in extracted_texts if item and str(item).strip()]
                            print(f"@@@@@@@@@@@@@@@@@@@@@@ {extracted_text}")

                            if line.startswith("W ") or line.startswith("D "):
                                if len(extracted_text) == 3:
                                    if int(extracted_text[2]) > 0:
                                        written_date = (f"{extracted_text[1][:2]}-{extracted_text[1][2:4]}-{extracted_text[1][4:]}")
                                        written_month = (f"{mon(extracted_text[1][2:4])}-{extracted_text[1][6:]}")
                                        extracted_data.append([filename, written_month, Datetime, "Whole Number", "On Stock", extracted_text[2], written_date, "PO No"])
                                        print(f"%%%%%%%%%%%%%% {extracted_text[2]}")
                                elif len(extracted_text) > 3:
                                    if int(extracted_text[3]) > 0:
                                        written_date = (f"{extracted_text[2][:2]}-{extracted_text[2][2:4]}-{extracted_text[2][4:]}")
                                        written_month = (f"{mon(extracted_text[2][2:4])}-{extracted_text[2][6:]}")
                                        extracted_data.append([filename,written_month, Datetime, "Whole Number", "On Stock", extracted_text[3], written_date, "PO No"])
                                        print(f"%%%%%%%%%%%%%% {extracted_text[3]}")



            except Exception as e:
                print(f"Error reading {filename}: {e}")

    if extracted_data:
        df = pd.DataFrame(extracted_data, columns=["Month", "PO No", "Date", "QTY (MT)", "On Stock", "QTY (MT)", "Date", "PO No"])
        empty_row = pd.DataFrame([[""] * len(df.columns)], columns=df.columns)
        df = pd.concat([df.iloc[:0], empty_row, df.iloc[0:]], ignore_index=True)


        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_title if sheet_title else "Default", index=False)
        print(f"✅ 엑셀 파일 저장 완료: {output_excel}")
    else:
        print("❌ 추출된 데이터가 없습니다.")

    wb = load_workbook(output_excel)

    for sheet in wb.sheetnames:
        ws = wb[sheet]

        max_row_f = ws.max_row
        for row in range(max_row_f, 0, -1):
            if ws[f'F{row}'].value is not None:
                last_row_f = row
                break

        for row in range(3, last_row_f + 1):
            # D_value = ws[f"D{row-1}"]
            # F_value = ws[f"F{row}"]
            # if F_value:
            #     ws[f"D{row}"] = f"={D_value}-{F_value}"
            F_value = f"F{row}"
            if F_value:
                ws[f"D{row}"] = f"=D{row-1}-F{row}"

    wb.save(output_excel)
    print(f"✅ 함수 추가 완료: {output_excel}")

Datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
folder_path = r"C:\Users\82109\Desktop\개인\Python Test"
filename = f"Result_{Datetime}.xlsx"
output_excel = os.path.join(folder_path, "date.xlsx")

extract_info(folder_path,output_excel)





#############################################################################################################################

import os
import fitz  # PyMuPDF
import re
import openpyxl
from datetime import datetime


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

    # 첫 번째 행: 빈 값
    ws.append([""] * 8)
    
    # 두 번째 행: 열 제목(데이터 종류)
    ws.append(["파일명", "월", "날짜", "단위", "재고 상태", "수량", "작성일", "PO 번호"])
    
    # 세 번째 행: 빈 값
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
                    written_date = f"{extracted_text[1][:2]}-{extracted_text[1][2:4]}-{extracted_text[1][4:]}"
                    written_month = f"{mon(extracted_text[1][2:4])}-{extracted_text[1][6:]}"
                    ws.append([filename, written_month, datetime.now().strftime("%Y-%m-%d"), "Whole Number", "On Stock", quantity, written_date, "PO No"])
                    print(f"[SMP Ibérica] 데이터 추가: {quantity}")


def process_samvardhana(text, filename, ws):
    """Samvardhana Motherson 문서를 처리하는 함수"""
    lines = text.replace(",", "").replace(".", "").split("\n")
    cleaned_list = [item.strip() for item in lines if item.strip()]
    
    for line in cleaned_list:
        extracted_texts = re.split(r'\s+', line)
        extracted_text = [item for item in extracted_texts if item.strip()]
        
        if line.startswith("M ") or line.startswith("N "):
            if len(extracted_text) >= 4:
                quantity = extracted_text[3]
                if quantity.isdigit() and int(quantity) > 0:
                    written_date = f"{extracted_text[2][:2]}-{extracted_text[2][2:4]}-{extracted_text[2][4:]}"
                    written_month = f"{mon(extracted_text[2][2:4])}-{extracted_text[2][6:]}"
                    ws.append([filename, written_month, datetime.now().strftime("%Y-%m-%d"), "Whole Number", "Stock", quantity, written_date, "PO Num"])
                    print(f"[Samvardhana Motherson] 데이터 추가: {quantity}")


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

    # Excel 파일 저장
    wb.save(output_excel)
    print(f"✅ 엑셀 파일 저장 완료: {output_excel}")


# 실행
folder_path = r"C:\Users\82109\Desktop\개인\Python Test"
output_excel = os.path.join(folder_path, "date.xlsx")

extract_info(folder_path, output_excel)


################################################################################################################
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
import openpyxl

def apply_conditional_formatting(ws):
    """D3 값을 기준으로 D4:D1000 범위에 조건부 서식 적용"""
    
    # 색상 정의
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # 초록색
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 노란색
    orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # 주황색
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # 빨간색

    # 범위 정의
    data_range = "D4:D1000"

    # 조건부 서식 추가 (D3을 기준으로 계산)
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["D4>=D$3*0.6"], stopIfTrue=True, fill=green_fill)  # 60% 이상 초록색
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["AND(D4>=D$3*0.4, D4<D$3*0.6)"], stopIfTrue=True, fill=yellow_fill)  # 40% 이상 노란색
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["AND(D4>0, D4<D$3*0.4)"], stopIfTrue=True, fill=orange_fill)  # 40% 미만 주황색
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["D4<0"], stopIfTrue=True, fill=red_fill)  # 음수(마이너스 값) 빨간색
    )
    # 0인 경우 색상 없음 (기본값 유지)

# 📝 엑셀 파일 로드 & 적용
output_excel = "C:/Users/82109/Desktop/개인/Python Test/date.xlsx"
wb = openpyxl.load_workbook(output_excel)

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    apply_conditional_formatting(ws)  # 각 시트에 조건부 서식 적용

wb.save(output_excel)
print(f"✅ 조건부 서식 적용 완료: {output_excel}")


