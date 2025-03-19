import os
import fitz
import pandas as pd
import re
import pdfplumber

def extract_info(pdf_folder, key_dict, output_excel):
    extracted_data = []

    for filename in os.listdir(pdf_folder):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(pdf_folder, filename)

            try:
                doc = fitz.open(file_path)

                for page in doc:
                    text = page.get_text("text")
                    print(f"This is extracted text: {text}")
                    lines = text.replace(",", "").replace(".","").split("\n")
                    cleaned_list = [item for item in lines if item and str(item).strip()]
                    print(f"This is splitted lines: {cleaned_list}")
                    for line in cleaned_list:
                        ############################################################### 03/19/25
                        # What if instead of make cleaned_list, try <if line is not "">
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
                                    extracted_text = re.split(r'[:=]', line)
                                    print(f"This is extracted_text: {extracted_text}")
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
    "name" : ["solomon", "LG Chem", "solomon"],
    "car" : ["benz", "audi"]
}

pdf_folder = r"C:\Users\82109\Desktop\개인\Python Test"
output_excel = r"C:\Users\82109\Desktop\개인\Python Test\date.xlsx"

extract_info(pdf_folder, key_dict, output_excel)
