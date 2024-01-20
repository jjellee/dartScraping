
import pandas as pd
import xlwings as xw
import os

# 파일명이 숫자인지 확인하는 함수
def is_number(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def flatten_multiindex(index):
    """Flatten a MultiIndex to a single-level Index by concatenating level values."""
    return ['_'.join(map(str, entry)) if isinstance(entry, tuple) else entry for entry in index]

def extract_strings_from_file(file_path):
    results = []
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            # Splitting the line at the first occurrence of ':'
            parts = line.split(':', 1)
            if len(parts) == 2:
                left_string, right_string = parts
                results.append((left_string.strip(), right_string.strip()))
    return results

def convert_html_table_to_excel(company_submitter, html_file_path, excel_writer, sheet_name, start_row):
    # Read the HTML table
    tables = pd.read_html(html_file_path, encoding='utf-8')
    #print(tables[0].columns)
     # 첫 번째 행을 무시하고 두 번째 행을 인덱스로 설정
    #tables[0].columns = tables[0].iloc[1]
    #table_to_write = tables[0].iloc[2:]

    # Write the modified HTML table DataFrame to the Excel file
    #table_to_write.to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, index=True)
    
    # Assuming tables[0] is the DataFrame you want to write
    # Write the HTML table DataFrame to the Excel file
    tables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, index=True)
    
    # Get the workbook and the sheet
    workbook  = excel_writer.book
    worksheet = excel_writer.sheets[sheet_name]

     # Write company_submitter[0] and company_submitter[1] in specified cells using openpyxl method
    worksheet.cell(row=start_row + 1, column=2, value='회사명')  # Writing in the first cell of start_row
    worksheet.cell(row=start_row + 1, column=3, value=company_submitter[0])  # Writing in the first cell of start_row

    worksheet.cell(row=start_row + 2, column=2, value='제출인')  # Writing in the first cell of the next row
    worksheet.cell(row=start_row + 2, column=3, value=company_submitter[1])  # Writing in the first cell of the next row

    #worksheet.delete_rows(start_row + 3)
    worksheet.delete_rows(start_row + 5)

    return start_row + len(tables[0]) + 7  # Return the new start row for the next table

def createFirstVersionExcel(equityFolder) :
    files = os.listdir(equityFolder)
    filtered_files = [f for f in files if (f.endswith('.txt') or f.endswith('.html')) and is_number(os.path.splitext(f)[0])]
    sorted_files = sorted(filtered_files, key=lambda x: int(os.path.splitext(x)[0]))

    company_submitter_list = []
    # 파일 순회
    for file in sorted_files:
        if file.endswith('.txt') : 
            stringPairs = extract_strings_from_file(os.path.join(equityFolder,file))
            company = stringPairs[0][1]
            submitter = stringPairs[1][1]
            company_submitter_list.append((company, submitter))
    #print(company_submitter_list[0][0])

    xlsxFile = equityFolder + '_v1' + '.xlsx'
    xlsxFilePath = os.path.join(equityFolder, xlsxFile)
    with pd.ExcelWriter(xlsxFilePath, engine='openpyxl') as writer:
        idx = 0
        start_row = 1  # Initialize the starting row
        for file in sorted_files:
            if file.endswith('.html'):
                print(f"Processing file: {file}")
                htmlFile = os.path.join(equityFolder, file)
                sheet_name = equityFolder  # Name of the consolidated sheet
                # Assuming txtContent is the text content related to the html file
                txtContent = "Your text content here"  # Replace with actual content
                #print('start_row:' + str(start_row))
                start_row = convert_html_table_to_excel(company_submitter_list[idx], htmlFile, writer, sheet_name, start_row)
                idx = idx + 1

    # 엑셀 파일을 다시 열고 첫 번째 열 삭제
    app = xw.App(visible=False)  # 엑셀 애플리케이션을 보이지 않게 설정
    book = app.books.open(xlsxFilePath)  # 엑셀 파일 열기

    try:
        sheet = book.sheets[sheet_name]  # 워크시트 선택
        sheet.range('A:A').delete()  # 첫 번째 열 삭제
        book.save()  # 변경 사항 저장
    finally:
        book.close()  # 파일 닫기
        app.quit()  # 엑셀 애플리케이션 종료
    return xlsxFilePath

def calculateFirstVersionExcel(xlsxFilePath) :
    # 엑셀 파일을 불러옵니다.
    df = pd.read_excel(xlsxFilePath)

    # 첫 번째 열을 순회하며 '회사명'이라는 단어를 찾습니다.
    company_name_row = None
    for index, value in enumerate(df.iloc[:, 0]):
        if '회사명' in str(value):
            company_name_row = index
            #print(f"'회사명'은 {index + 1}번째 행에 있습니다.")

            # '회사명'이 있는 행에서 3행 뒤부터 '증감'이라는 단어를 찾습니다.
            row_index = company_name_row + 3
            if '보고사유' in df.iloc[row_index, 0] :
                # Case 1
                # 단순히 증감 * 처분 단가
            elif '성명' in df.iloc[row_index, 0] :
                #Case 2
            else :

            '''
            for col_index in range(len(df.columns)):
                if '증감' in str(df.iloc[row_index, col_index]):
                    print(f"'증감'은 {row_index + 1}행 {col_index + 1}열에 있습니다.")
            '''
    # Case 1 : 보고사유, 처분단가열 1개
    
    
    # Case 2 : 성명 (명칭), 처분단가열 2개


def main () :
    equityFolder = '2024.01.18_지분공시'  # Update the folder path
    xlsxFilePath = createFirstVersionExcel(equityFolder)
    #xlsxFilePath = 'E:/bbAutomation/dartScraping/2024.01.18_지분공시/2024.01.18_지분공시_v1.xlsx'
    calculateFirstVersionExcel(xlsxFilePath)


main()

'''
file_path=os.path.join(os.getcwd(), 'example.xlsx')
# 파일이 존재하지 않으면 새 파일 생성
if not os.path.exists(file_path):
    wb = xw.Book()  # 새 워크북 생성
    wb.save(file_path)  # 파일로 저장
else:
    wb = xw.Book(file_path)  # 기존 파일 열기

sheet = wb.sheets['Sheet1']

# 엑셀에 숫자를 세로로 쓰기
numbers = [1, 2, 3, 4, 5]  # 예시 데이터
sheet.range('A1').options(transpose=True).value = numbers

# VBA의 SUM 함수를 사용하여 합 구하기
sum_formula = f'=SUM(A1:A{len(numbers)})'
sheet.range('B1').value = sum_formula

# 계산된 합 가져오기
total_sum = sheet.range('B1').value
print("계산된 합:", total_sum)

# 엑셀 파일 저장 및 닫기
wb.save(file_path)
wb.close()
'''
