import os
import fitz
import pandas as pd
import re
import pdfplumber
from datetime import datetime

def make_chart(cleaned_list, start_first_word, start_last_word, end_first_word, end_lastword):
    counter = 1
    i = 0
    capture = False
    chart = []
    for line in cleaned_list:
        if cleaned_list[i].startswith(start_first_word) and cleaned_list[i].endswith(start_last_word):
            capture = True
            continue
        if cleaned_list[i].startswith(end_first_word) and cleaned_list[i].endswith(end_lastword):
            capture = False
            break
        if capture:
            chart.append(f"{counter}. {line.strip()}")
            counter+=1
        
        return chart

def extract_info(folder_path, key_dict, output_excel):
    extracted_data = []

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)

            try:
                doc = fitz.open(file_path)

                for page in doc:
                    text = page.get_text("text")
                    print(f"This is extracted text: {text}")
                    lines = text.replace(",", "").replace(".","").split("\n")
                    # 데이터 조정.....
                    # W (스페이스 포함) 로 시작하는 줄은 다음 줄과 병합 되도록 설정.
                    # if line for line in lines startswith("W "):
                    # or
                    # for line in lines:
                    #     if line.startswith("W "):
                    #         다음 줄과 병합.
                    cleaned_list = [item for item in lines if item and str(item).strip()]
                    i = 0
                    merge_lines = []
                    while i < len(cleaned_list):
                        if cleaned_list[i].strip().startswith("W "):
                            merge_lines.append(f"{cleaned_list[i]}{cleaned_list[i+1]}")
                            i+=2
                        elif cleaned_list[i].strip().endswith("number/date"):
                            merge_lines.append(f"{cleaned_list[i]} schedule!!/ {cleaned_list[i+1]}")
                            i+=2
                        elif cleaned_list[i].strip().startswith("PO Number"):
                            merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i-1]}")
                            i+=1
                        elif cleaned_list[i].strip() == ("Material") and cleaned_list[i-2].strip() == ("Item"):
                            merge_lines.append(f"{cleaned_list[i-1]} {cleaned_list[i]}")
                            i+=1
                        elif "Part #" in cleaned_list[i].strip():
                            group = {"Part #", "Vendor Part #"}
                            if all(word in cleaned_list[i] for word in group):
                                merge_lines.append(f"{cleaned_list[i]} {cleaned_list[i+1]}")
                                i+=1
                        else:
                            merge_lines.append(str(cleaned_list[i]))
                            i += 1
                    print(f"This is splitted lines: {merge_lines}")
                    for line in merge_lines:
                        ############################################################### 03/19/25
                        # What if instead of make cleaned_list, try <if line is not "">
                        # + it works but could only resolve either " " or ""      ##### 03/20/25
                        ###############################################################
                        if line:
                            print(f"This is line: {line}")
                            lower_line = line.lower()
                            print(f"This is lower_line: {lower_line}")
                            for column, keyword in key_dict.items():
                                print(f"This is column: {column}")
                                print(f"This is keyword: {keyword}")
                                ############################################################################ 03/19/25
                                # Need to understand this part <k in lower_line for k in keyword>
                                ############################################################################
                                if any(k in lower_line for k in keyword):
                                    ############################################################################################################## 03/19/25
                                    # What if I also want to split empty space? Simply just add one more split? or change something inside of [ ]?
                                    ##############################################################################################################
                                    extracted_text = re.split(r'[:= ]', line)
                                    print(f"######################This is extracted_text: {extracted_text}")
                                    if column == "Material":
                                        nott = {"Raw", "Thank"}
                                        if not any(word in extracted_text for word in nott):
                                            sentence = "".join(extracted_text)
                                            if sentence.startswith("Item") and sentence.endswith("Description"): 
                                                continue
                                            elif "FabricationCum" in sentence:
                                                material_code = sentence.split("Cum")[2].strip()
                                                extracted_data.append([filename,column,keyword,line.strip(),Datetime,material_code])
                                            elif "Part#" in sentence:
                                                material_code = sentence.split("ModelYear")[1].split("StandardPackage")[0].strip()
                                                print(f"!@#####@#@!#!@#!@#@#@#@#@#@#@#@#@#@## {material_code}")
                                                extracted_data.append([filename,column,keyword,line.strip(),Datetime,material_code])
                                            else:
                                                material_code = " ".join([word for word in extracted_text if word != "Material"])
                                                extracted_data.append([filename,column,keyword,line.strip(),Datetime,material_code])
                                    elif column == "Price":
                                        cleaned_price_list = [item for item in extracted_text if item and str(item).strip()]
                                        print(f"############################ cleeaned_price_list: {cleaned_price_list}")
                                        if cleaned_price_list[0] == "W" or cleaned_price_list[0] == "D":
                                            if len(cleaned_price_list) > 3:
                                                if cleaned_price_list[3].endswith("W") or cleaned_price_list[3].endswith("D"):
                                                    divided_first = "".join(cleaned_price_list[3][:3])
                                                    print(f"############################ divided_first: {divided_first}")
                                                    if int(divided_first) > 0:
                                                        this_date = (f"{cleaned_price_list[2][:2]}-{cleaned_price_list[2][2:4]}-{cleaned_price_list[2][4:]}")
                                                        extracted_data.append([filename,column,keyword,line.strip(),this_date,divided_first.strip()])
                                                    if int(cleaned_price_list[6]) > 0:
                                                        this_date = (f"{cleaned_price_list[5][:2]}-{cleaned_price_list[5][2:4]}-{cleaned_price_list[5][4:]}")
                                                        extracted_data.append([filename,column,keyword,line.strip(),this_date,cleaned_price_list[6]])
                                                else:
                                                    extracted_data.append([filename,column,keyword,line.strip(),Datetime,cleaned_price_list[3].strip()])
                                                    print(f"############################ cleaned_price_list fianllllll: {cleaned_price_list}")
                                            elif int(cleaned_price_list[2]) > 0:
                                                print(f"This is price_code: {cleaned_price_list}")
                                                this_date = (f"{cleaned_price_list[1][:2]}-{cleaned_price_list[1][2:4]}-{cleaned_price_list[1][4:]}")
                                                extracted_data.append([filename,column,keyword,line.strip(),this_date,cleaned_price_list[2]])
                                            else:
                                                print("Too Short")
                                        elif cleaned_price_list[1] == "Week" or cleaned_price_list[1] == "Month":
                                            if cleaned_price_list[2] == "Raw":
                                                if int(cleaned_price_list[4]) > 0:
                                                    this_date = cleaned_price_list[0]
                                                    new_price = int(cleaned_price_list[4])/100
                                                    extracted_data.append([filename,column,keyword,line.strip(),this_date,new_price])
                                            elif int(cleaned_price_list[3]) > 0:
                                                this_date = cleaned_price_list[0]
                                                new_price = int(cleaned_price_list[3])/100
                                                extracted_data.append([filename,column,keyword,line.strip(),this_date,new_price])
                                        elif cleaned_price_list[3] == "Floating":
                                            if cleaned_price_list[1] == "to" and int(cleaned_price_list[5]) > 0:
                                                new_price = int(cleaned_price_list[5])/100
                                                extracted_data.append([filename,column,keyword,line.strip(),this_date,new_price])
                                        else:
                                            print("Too Short")

                                    elif column == "Order Number":

                                        ######################################################### "/" 이나 추가 문자 없는 상태로 스플릿 해야하는데 조건이 안걸림. 
                                        # if "schedule!!" in lower_line:
                                        #     Scheduling_list = "".join(lower_line.split("schedule!!")).split("/")
                                        # else:
                                        #     Scheduling_list.split()
                                        #########################################################
                                        if "schedule!!" in lower_line:
                                            Scheduling_list = "".join(lower_line.split("schedule!!")).split("/")
                                            print(f"This is scheduling list: {Scheduling_list}")
                                        else:
                                            Scheduling_list = (lower_line.split("#:"))[1].split()
                                            print(f"This is scheduling list: {Scheduling_list}")

                                        if len(Scheduling_list) > 2:
                                            if int(Scheduling_list[2]) > 0:
                                                extracted_data.append([filename,column,keyword,line.strip(),Datetime,str(Scheduling_list[2].strip())])
                                            else:
                                                extracted_data.append([filename,column,keyword,line.strip(),Datetime,str(Scheduling_list[1].strip())])
                                        elif Scheduling_list[0].startswith("po"):
                                            extracted_data.append([filename,column,keyword,line.strip(),Datetime,str(extracted_text[2].strip())])
                                        else:
                                            extracted_data.append([filename,column,keyword,line.strip(),Datetime,str(Scheduling_list[0].strip())])
                                    elif column == "Company Name":
                                        C_Name = keyword
                                        extracted_data.append([filename,column,C_Name,line.strip(),Datetime,str(extracted_text[1].strip())])
                                    else:
                                        extracted_data.append([filename,column,keyword,line.strip(),Datetime,extracted_text])

            except Exception as e:
                print(f"Error reading {filename}: {e}")

    if extracted_data:
        df = pd.DataFrame(extracted_data, columns=["파일명", "컬럼", "키워드", "전체 문장", "날짜", "추출 값"])
        df.to_excel(output_excel, index=False, engine="openpyxl")
        print(f"✅ 엑셀 파일 저장 완료: {output_excel}")
    else:
        print("❌ 추출된 데이터가 없습니다.")

############################################################################################################## 03/19/25
# To add both .pdf and .xlsx 
# Final_Data = []
# Final_Data.append(extracted_data_for_pdf)
# Final_Data.append(extracted_data_for_xlsx)
# if Final_Data: bla bla bla
##############################################################################################################

# Data needed to extract
key_dict = {
    "Material" : ["material", "part #"],
    "Price" : ["w ", "d ", "week", "month", "floating"],
    "Order Number" : ["number/date", "po number", "PO", "p/o #"],
    "Company Name" : ["pl0770", "smpibéricas.l.u.", "samvardhana motherson peguform", "biesterfeld polybass s.p.a"]
}

# Current Date
Datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

# Path
folder_path = r"C:\Users\82109\Desktop\개인\Python Test"
filename = f"Result_{Datetime}.xlsx"
output_excel = os.path.join(folder_path, "date.xlsx")

extract_info(folder_path, key_dict, output_excel)
