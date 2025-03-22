import os
import fitz
import pandas as pd
import re
import pdfplumber

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
                            merge_lines.append(f"{cleaned_list[i]} schedule!! {cleaned_list[i+1]}")
                            i+=2
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
                                        material_code = " ".join([word for word in extracted_text if word != "Material"])
                                        extracted_data.append([filename,column,keyword,line.strip(),material_code])
                                    elif column == "Price":
                                        cleaned_price_list = [item for item in extracted_text if item and str(item).strip()]
                                        if cleaned_price_list[0] == "W" or cleaned_price_list[0] == "D":
                                            if int(cleaned_price_list[2]) > 0:
                                                print(f"This is price_code: {cleaned_price_list}")
                                                extracted_data.append([filename,column,keyword,line.strip(),cleaned_price_list])
                                    elif column == "Scheduling":
                                        Scheduling_list = lower_line.split("schedule!!")
                                        test = "".join(Scheduling_list)
                                        tt = test.split("/")
                                        print(f"This is scheduling list: {tt}")
                                        if int(tt[1]) > 0:
                                            extracted_data.append([filename,column,keyword,line.strip(),tt[1]])
                                        else:
                                            extracted_data.append([filename,column,keyword,line.strip(),tt[2]])
                                    else:
                                        extracted_data.append([filename,column,keyword,line.strip(),extracted_text])

            except Exception as e:
                print(f"Error reading {filename}: {e}")

    if extracted_data:
        df = pd.DataFrame(extracted_data, columns=["파일명", "컬럼", "키워드", "전체 문장", "추출 값"])
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

key_dict = {
    "Material" : ["material", "LG Chem"],
    "Price" : ["w ", "d "],
    "Scheduling" : ["number/date"]
}

folder_path = r"C:\Users\82109\Desktop\개인\Python Test"
output_excel = r"C:\Users\82109\Desktop\개인\Python Test\date.xlsx"

extract_info(folder_path, key_dict, output_excel)
