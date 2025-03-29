import os
import fitz
import pandas as pd
import re
import pdfplumber
from datetime import datetime
from openpyxl import load_workbook


def mon(text):
    months = {"01": "JAN", "02": "FEB", "03": "MAR", "04": "APR", "05": "MAY", "06": "JUN", 
              "07": "JUL", "08": "AUG", "09": "SEP", "10": "OCT", "11": "NOV", "12": "DEC"}
    return months.get(text, "")

def sheet_name(wb, text):
    if text in wb.sheetnames:
        return wb[sheet_name]
    else:
        return wb.create_sheet(title=text)



def extract_info(folder_path,output_excel):
    extracted_data = []

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)

            try:
                doc = fitz.open(file_path)

                for page in doc:
                    text = page.get_text("text")
                    print(f"This is extracted text: {text}")


                    if "SMP Ibérica" in text:
                        if sheet_title == "Default" or sheet_title.startswith("Sheet"):
                            sheet_title = "SMP Ibérica"
                        else:
                            return wb.create_sheet(title="SMP Ibérica")

                        
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




                    elif "Samvardhana Motherson Peguform" in text:
                        if sheet_title == "Default" or sheet_title.startswith("Sheet"):
                            sheet_title = "Samvardhana Motherson Peguform"
                        else:
                            wb.create_sheet(title="Samvardhana Motherson Peguform")

                        
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
