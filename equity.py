
import pandas as pd
import xlwings as xw
import os, sys
import numpy as np

secondTableColumn = 14

def number_to_alphabet(number):
    # 숫자를 알파벳으로 변환 (1 -> 'A', 2 -> 'B', ...)
    return chr(64 + number)

# 파일명이 숫자인지 확인하는 함수
def is_number(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

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

    # 작업 테이블
    tables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, index=True)
    
    # 원본 테이블 
    tables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, startcol=secondTableColumn, index=True)

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

def HTMLtoExcel(equityFolder) :
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

    #end_row=0
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
        #end_row = start_row - 4
    column_letter = number_to_alphabet(secondTableColumn)

    #print(end_row)

    # 엑셀 파일을 다시 열고 첫 번째 열 삭제
    app = xw.App(visible=False)  # 엑셀 애플리케이션을 보이지 않게 설정
    book = app.books.open(xlsxFilePath)  # 엑셀 파일 열기
    
    try:
        sheet = book.sheets[sheet_name]  # 워크시트 선택
        sheet.range('A:A').delete()  # 첫 번째 열 삭제
        sheet.range(f'{column_letter}:{column_letter}').delete()  # 지정된 열 삭제
        book.save()  # 변경 사항 저장
    finally:
        book.close()  # 파일 닫기
        app.quit()  # 엑셀 애플리케이션 종료

    
    return xlsxFilePath

def adjustnExcel(xlsxFilePath) : 
    app = xw.App(visible=False)  # Excel 애플리케이션을 보이지 않게 설정
    book = app.books.open(xlsxFilePath)  # Excel 파일 열기

    try:
        sheet = book.sheets[0]  # 첫 번째 시트 선택
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row  # 첫 번째 열의 마지막 행 찾기
        row_index = 1
        print('last_row : ' + str(last_row))
        while row_index <= last_row:
            value = sheet.range(f'A{row_index}').value  # 각 행의 A열 값 읽기
            if value == '회사명':
                row_index += 1  # '회사명' 행 다음부터 검사 시작
                while row_index <= last_row + 1:
                    # 현재 행 전체가 비어있는지 확인
                    row_values = sheet.range(f'{row_index}:{row_index}').value
                    if all(cell is None for cell in row_values):
                        #print(f'Row {row_index}')  # 비어 있는 행의 번호를 출력
                        break
                    row_index += 1
                tableEndValue = sheet.range(f'A{row_index-1}').value  # 테이블 끝행의 A열 값 읽기
                print('Row' + str(row_index-1) + ':' + tableEndValue)
                if tableEndValue != '합 계':
                    sheet.api.Rows(row_index).Insert()  # 새 행 삽입
                    sheet.range(f'A{row_index}').value = '합 계'  # 새 행의 A열에 '합 계' 입력
                    last_row += 1  # 행이 추가되었으므로 마지막 행 번호 업데이트
            else:
                row_index += 1

    finally:
        book.close()  # 파일 닫기
        app.quit()  # Excel 애플리케이션 종료


def main () :
    equityFolder = '2024.01.18_지분공시'  # Update the folder path
    xlsxFilePath = HTMLtoExcel(equityFolder)
    #xlsxFilePath = 'E:/bbAutomation/dartScraping/2024.01.18_지분공시/2024.01.18_지분공시_v1.xlsx'
    #xlsxFilePath = '/Users/yee/Documents/dartScraping/2024.01.18_지분공시/2024.01.18_지분공시_v1.xlsx'
    adjustnExcel(xlsxFilePath)
    #calculateFirstVersionExcel(xlsxFilePath)


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

def calculateFirstVersionExcel(xlsxFilePath) :
    # 엑셀 파일을 불러옵니다.
    df = pd.read_excel(xlsxFilePath)

    tables_info = []
        tables_info.append({
        'top_left_cell': (left_table_start_row, left_table_start_col_label),
        'bottom_right_cell': (table_data_end_row - 1, df.columns[left_table_end_col]),
    })

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
                # 단순히 증감 * 처분 단가로 평균단가 구하기
                for col_index in range(len(df.columns)):
                    if '증감' in str(df.iloc[row_index, col_index]):
                        deltaRow = row_index
                        deltaCol = col_index
                        print(f"'증감'은 {row_index + 2}행 " + number_to_alphabet(col_index + 1) + "열에 있습니다.")
                        break
                # '증감'의 '합계'가 존재하지 않을 경우
                

            elif '성명' in df.iloc[row_index, 0] :
                # Case 2
                # 처분 단가 2열 중 숫자인 것만, 0은 숫자X, 혼잡(숫자+문자)인 경우 숫자만 가져오기
                for col_index in range(len(df.columns)):
                    if '증감' in str(df.iloc[row_index, col_index]):
                        print(f"'증감'은 {row_index + 2}행 " + number_to_alphabet(col_index + 1) + "열에 있습니다.")
                        break
            else :
                print('알지 못하는 형식입니다')
                sys.exit(0)
'''










'''
def adjustnExcel(xlsxFilePath) : 
    # 엑셀 파일을 불러옵니다.
    df = pd.read_excel(xlsxFilePath)

    # Get the shape of the DataFrame
    num_rows, num_cols = df.shape
    print (num_rows, num_cols)

    # Create a list to store new rows
    row_index = 0
    while row_index < num_rows:
        if df.iloc[row_index, 0] == '회사명':
            #print('회사명 : ' + df.iloc[row_index, 1] + ' row_index : ' + str(row_index + 2))
            while row_index < num_rows and not df.iloc[row_index].isnull().all():
                row_index += 1
            #print('isnuall all : ' + str(row_index + 2))
            if row_index < num_rows and df.iloc[row_index - 1, 0] != '합 계':
                # 새로운 행을 생성하고 첫 번째 열에 '합 계'를 추가합니다.
                new_row = ['합 계'] + [np.nan] * (num_cols - 1)
                # DataFrame에 새로운 행을 삽입합니다.
                #df = pd.concat([df.iloc[:row_index], pd.DataFrame([new_row], columns=df.columns), df.iloc[row_index:]]).reset_index(drop=True)
                df = pd.concat([df.iloc[:row_index], pd.DataFrame([new_row], columns=df.columns), df.iloc[row_index:]], ignore_index=True)
                print('new_row inserted at index: ' + str(row_index + 2))
                # 행을 추가했으므로 num_rows 업데이트
                num_rows += 1
                #print (num_rows, num_cols)
        row_index += 1

    # Save the modified DataFrame to a new Excel file
    secondxlsxFilePath = os.path.join(equityFolder, secondxlsxFile)
    df.to_excel(secondxlsxFilePath, index=False)

    if is_row_empty == True :
        print(f"Row {row_index+2} is empty: {is_row_empty}")

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
                # 단순히 증감 * 처분 단가로 평균단가 구하기
                for col_index in range(len(df.columns)):
                    if '증감' in str(df.iloc[row_index, col_index]):
                        deltaRow = row_index
                        deltaCol = col_index
                        print(f"'증감'은 {row_index + 2}행 " + number_to_alphabet(col_index + 1) + "열에 있습니다.")
                        break
                # '증감'의 '합계'가 존재하지 않을 경우
                

            elif '성명' in df.iloc[row_index, 0] :
                # Case 2
                # 처분 단가 2열 중 숫자인 것만, 0은 숫자X, 혼잡(숫자+문자)인 경우 숫자만 가져오기
                for col_index in range(len(df.columns)):
                    if '증감' in str(df.iloc[row_index, col_index]):
                        print(f"'증감'은 {row_index + 2}행 " + number_to_alphabet(col_index + 1) + "열에 있습니다.")
                        break
            else :
                print('알지 못하는 형식입니다')
                sys.exit(0)
'''