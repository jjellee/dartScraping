
import pandas as pd
import xlwings as xw
import os, sys
import numpy as np
import re
from itertools import groupby

secondTableColumn = 14
thirdTableColumn = 22

def extract_number_from_filename(filename):
    # 정규 표현식을 사용하여 파일명에서 숫자 부분을 찾습니다.
    match = re.search(r'\d{1,3}', filename)
    if match:
        # 숫자 부분을 찾았다면, 이를 반환합니다.
        return int(match.group())
    else:
        # 파일명에 숫자가 없다면 None을 반환합니다.
        return None
    
def number_to_alphabet(number):
    # 숫자를 알파벳으로 변환 (1 -> 'A', 2 -> 'B', ...)
    return chr(64 + number)

# 파일명이 숫자인지 확인하는 함수
def is_number(s):
    if s is None:
        return False
    try:
        int(s)
        return True
    except ValueError:
        return False
    
def is_number_in_string(s):
    if s is None:
        return False

    if isinstance(s, (int, float)):
        return True  # s가 숫자 타입인 경우 True 반환

    if isinstance(s, str):
        if re.search(r'\d', s):  # 문자열에 숫자가 있는지 확인
            return True
        try:
            int(s)  # 숫자로 변환 가능한지 시도
            return True
        except ValueError:
            return False

    return False  # s가 문자열이 아닌 다른 타입인 경우 False 반환
    
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

#def convert_html_table_to_excel(company_submitter, html_file_path, excel_writer, sheet_name, start_row):

def convert_html_table_to_excel(company_submitter, tradeHTMLlfilePath, reporterHTMLfilePath, shareRatioHTMLfilePath, excel_writer, sheet_name, start_row) : 
    # Read the HTML table
    tradeTables = pd.read_html(tradeHTMLlfilePath, encoding='utf-8')
    reporterTables = pd.read_html(reporterHTMLfilePath, encoding='utf-8')
    shareRatioTables = pd.read_html(shareRatioHTMLfilePath, encoding='utf-8')

    #print(tradeTables[0])
    #print(tables[0].columns)
     # 첫 번째 행을 무시하고 두 번째 행을 인덱스로 설정
    #tables[0].columns = tables[0].iloc[1]
    #table_to_write = tables[0].iloc[2:]

    # Write the modified HTML table DataFrame to the Excel file
    #table_to_write.to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, index=True)
    
    # Assuming tables[0] is the DataFrame you want to write
    # Write the HTML table DataFrame to the Excel file

    # 작업 테이블
    tradeTables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, index=True)
    #reporterTables[0].iloc[1:].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, startcol=secondTableColumn, index=True)
    reporterTables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, startcol=secondTableColumn, index=True)
    shareRatioTables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, startcol=thirdTableColumn, index=True)

    # 원본 테이블 
    #tables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, startcol=secondTableColumn, index=True)

    # Get the workbook and the sheet
    #workbook  = excel_writer.book
    worksheet = excel_writer.sheets[sheet_name]

     # Write company_submitter[0] and company_submitter[1] in specified cells using openpyxl method
    worksheet.cell(row=start_row + 1, column=2, value='회사명')  # Writing in the first cell of start_row
    worksheet.cell(row=start_row + 1, column=3, value=company_submitter[0])  # Writing in the first cell of start_row

    worksheet.cell(row=start_row + 1, column=5, value='보고서명')  # Writing in the first cell of start_row
    worksheet.cell(row=start_row + 1, column=6, value=company_submitter[2])  # Writing in the first cell of start_row

    worksheet.cell(row=start_row + 2, column=2, value='제출인')  # Writing in the first cell of the next row
    worksheet.cell(row=start_row + 2, column=3, value=company_submitter[1])  # Writing in the first cell of the next row

    #worksheet.delete_rows(start_row + 3)
    worksheet.delete_rows(start_row + 5)

    return start_row + max(len(tradeTables[0]), len(reporterTables[0]), len(shareRatioTables[0])) + 7  # Return the new start row for the next table

def extract_number_from_filename(filename):
    match = re.search(r'\d{1,3}', filename)
    if match:
        return int(match.group())
    else:
        return 9999  # 파일명에 숫자가 없는 경우 큰 수를 반환

def sortedTextFiles(folderPath) :
    # 폴더 내의 모든 파일을 리스트로 가져옵니다.
    files = [f for f in os.listdir(folderPath) if os.path.isfile(os.path.join(folderPath, f)) and f.endswith('.txt')]
    sorted_files = sorted(files, key=lambda x: extract_number_from_filename(x))
    return sorted_files

def extract_details_from_filename(filename):
    # 파일명에서 숫자 추출
    number_match = re.search(r'\d{1,3}', filename)
    number = int(number_match.group()) if number_match else None

    # 문자열 순서 정의
    order_dict = {'세부변동내역': 1, '보고자에관한상황': 2, '소유특정증권등의수및소유비율': 3}
    for key in order_dict:
        if key in filename:
            return (number, order_dict[key])
    return (number, None)

def sortedHTMLFiles(folderPath) :
    files = [f for f in os.listdir(folderPath) if os.path.isfile(os.path.join(folderPath, f)) and f.endswith('.html')]
    # 파일을 정렬 기준에 따라 정렬합니다.
    sorted_files = sorted(files, key=lambda x: extract_details_from_filename(x))
    return sorted_files

def HTMLtoExcel(equityFolder) :
    STFs = sortedTextFiles(equityFolder)
    
    company_submitter_list = []
    for file in STFs:
        stringPairs = extract_strings_from_file(os.path.join(equityFolder,file))
        company = stringPairs[0][1]
        submitter = stringPairs[1][1]
        reportName = stringPairs[2][1]
        company_submitter_list.append((company, submitter, reportName))

    SHFs = sortedHTMLFiles(equityFolder)
    #end_row=0
    xlsxFile = equityFolder + '_v1' + '.xlsx'
    xlsxFilePath = os.path.join(equityFolder, xlsxFile)
    sheet_name = equityFolder  # Name of the consolidated sheet


    with pd.ExcelWriter(xlsxFilePath, engine='openpyxl') as writer:
        idx = 0
        start_row = 1
        for key, group in groupby(SHFs, key=lambda x: extract_details_from_filename(x)[0]):
            grouped_files = list(group)
            if len(grouped_files) != 3 :
                print(f'html파일이 3개가 이닙니다 : {grouped_files}')
                sys.exit(-1)
            print(f"Processing files: {grouped_files}")
            #order_dict = {'세부변동내역': 1, '보고자에관한상황': 2, '소유특정증권등의수및소유비율': 3}
            tradeHTMLlfilePath = os.path.join(equityFolder, grouped_files[0]) # 세부변동내역
            reporterHTMLfilePath = os.path.join(equityFolder, grouped_files[1]) # 보고자에관한상황
            shareRatioHTMLfilePath = os.path.join(equityFolder, grouped_files[2]) # 소유특정증권등의수및소유비율
            # Assuming txtContent is the text content related to the html file
            #print('start_row:' + str(start_row))
            #print(reporterHTMLfilePath)
            start_row = convert_html_table_to_excel(company_submitter_list[idx], tradeHTMLlfilePath, reporterHTMLfilePath, shareRatioHTMLfilePath, writer, sheet_name, start_row)
            idx = idx + 1

    '''
    with pd.ExcelWriter(xlsxFilePath, engine='openpyxl') as writer:
        idx = 0
        start_row = 1  # Initialize the starting row
        for file in SHFs:
            print(f"Processing file: {file}")
            htmlFile = os.path.join(equityFolder, file)
            sheet_name = equityFolder  # Name of the consolidated sheet
            # Assuming txtContent is the text content related to the html file
            txtContent = "Your text content here"  # Replace with actual content
            #print('start_row:' + str(start_row))
            start_row = convert_html_table_to_excel(company_submitter_list[idx], htmlFile, writer, sheet_name, start_row)
            idx = idx + 1
    '''
        #end_row = start_row - 4
    
    
    
    #print(end_row)

    # 엑셀 파일을 다시 열고 첫 번째 열 삭제
    app = xw.App(visible=False)  # 엑셀 애플리케이션을 보이지 않게 설정
    book = app.books .open(xlsxFilePath)  # 엑셀 파일 열기
    
    try:
        sheet = book.sheets[sheet_name]  # 워크시트 선택
        sheet.range('A:A').delete()  # 첫 번째 열 삭제
        column_letter = number_to_alphabet(secondTableColumn)
        sheet.range(f'{column_letter}:{column_letter}').delete()  # 지정된 열 삭제
        column_letter = number_to_alphabet(thirdTableColumn-1) # 'secondTableColumn'에 해당하는 열 하나를 삭제했기 때문에
        sheet.range(f'{column_letter}:{column_letter}').delete()  # 지정된 열 삭제
        book.save()  # 변경 사항 저장
    finally:
        book.close()  # 파일 닫기
        app.quit()  # 엑셀 애플리케이션 종료
   
    return xlsxFilePath


def count_numeric_rows(sheet, start_row, end_row, col):
    count = 0
    for row in range(start_row, end_row):
        cell_value = sheet.range((row, col)).value
        if is_number_in_string(cell_value):
            count += 1
    return count

def tableForm(sheet, indexRow) :
    case = None
    deltaCol = None
    priceCol = None
    remarksCol = None
    sumRow = None

    #'합 계' row 구하기  : sumRow  
    row = indexRow
    while True :
        value = sheet.range(f'A{row}').value
        if value == '합 계' :
            sumRow = row
            break
        row += 1

    # case, deltaCol, priceCol, remarksCol 구하기
    col = 1  # Start from the first colum
    while True :
        cell_value = sheet.range((indexRow, col)).value
        if col == 1:
            if '성명' in cell_value : #취득/처분 단가 2개 열 : 첫 열은 '성명 (명칭)'
                case = 2
            elif cell_value == '보고사유' : #취득/처분 단가 1개 열 : 첫 열은 '보고사유'
                case = 1
            else :
                print('새로운 폼! 처리 필요')
        elif cell_value == '증감' :
            deltaCol = col
        elif '취득/처분 단가' in cell_value :
            if case == 1 :
                priceCol = col
            elif case == 2 :
                # 숫자가 써져있는 row 개수가 더 많은 컬럼을 선택
                count_col = count_numeric_rows(sheet, indexRow + 1, sumRow, col) #숫자와 문자가 혼합되어 있는 경우에도 숫자로 인식. 보통 '(원)'도 들어감.
                count_next_col = count_numeric_rows(sheet, indexRow + 1, sumRow, col + 1)
                if count_next_col > count_col:
                    priceCol = col + 1
                else:
                    priceCol = col
                col += 1 #'취득/처분 단가' 열이 하나 더 있음
        elif cell_value == '비 고' : # 마지막 열
            remarksCol = col
            break
        col += 1

    return case, deltaCol, priceCol, remarksCol, sumRow

def addAveragePriceColumn(sheet, row) : 
    col = 1  # Start from the first colum
    while True :
        cell_value = sheet.range((row, col)).value
        if cell_value == '비 고' :
            target_cell = sheet.range(row, col + 1)
            target_cell.value = '증감X취득/처분 단가'

            # Apply bold formatting
            target_cell.api.Font.Bold = True

            # Add borders to the cell
            for border_id in range(7, 13):  # These are the border index values for Excel
                target_cell.api.Borders(border_id).LineStyle = 1  # Solid line
            break
        col += 1

def sort_and_write_back(sheet, start_row, end_row):
    data = []  # 데이터를 저장할 리스트

    # 데이터 추출
    for row in range(start_row, end_row):
        row_data = []  # 현재 행의 데이터를 저장할 리스트
        for col in range(1, secondTableColumn):  # sheet.ncols: 시트의 열 개수
            row_data.append(sheet.range((row, col)).value)
        data.append(row_data)

    # 데이터 정렬 (A열 값으로 정렬)
    data.sort(key=lambda x: x[0])

    # 정렬된 데이터를 시트에 다시 작성
    current_row = start_row
    for row_data in data:
        for col, value in enumerate(row_data, start=1):
            sheet.range((current_row, col)).value = value
        current_row += 1

def calculateAveragePrice(xlsxFilePath) : 
    app = xw.App(visible=False)  # Excel 애플리케이션을 보이지 않게 설정
    book = app.books.open(xlsxFilePath)  # Excel 파일 열기

    try:
        # '합 계' 행 추가 + '증감X취득/처분 단가' 열 추가
        sheet = book.sheets[0]  # 첫 번째 시트 선택
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row  # 첫 번째 열의 마지막 행 찾기
        row = 1
        print('last_row : ' + str(last_row))
        while row <= last_row :
            value = sheet.range(f'A{row}').value  # 각 행의 A열 값 읽기
            if value == '회사명':
                addAveragePriceColumn(sheet, row + 3)   
                row += 1  # '회사명' 행 다음부터 검사 시작
                while row <= last_row + 1:
                    # 현재 행 전체가 비어있는지 확인
                    row_values = sheet.range(f'{row}:{row}').value
                    #if all(cell is None for cell in row_values[:secondTableColumn - 1]):
                    if row_values[0] is None :
                        #print(f'Row {row}')  
                        break
                    row += 1
                tableEndValue = sheet.range(f'A{row-1}').value  # 테이블 끝행의 A열 값 읽기
                #print('Row' + str(row-1) + ':' + tableEndValue)
                if tableEndValue != '합 계':
                    sheet.api.Rows(row).Insert()  # 새 행 삽입
                    sheet.range(f'A{row}').value = '합 계'  # 새 행의 A열에 '합 계' 입력
                    last_row += 1  # 행이 추가되었으므로 마지막 행 번호 업데이트
            else:
                row += 1

        
        # 모든 테이블에 대해서 Case 2인 경우 (첫 번째 열이 '성명 (명칭)'), 동일인의 거래 내역을 붙여서 정렬
        row = 1
        while row <= last_row :
            value = sheet.range(f'A{row}').value  # 각 행의 A열 값 읽기
            if value == '회사명' :
                tableIndexRow = row + 3     
                case, deltaCol, priceCol, remarksCol, sumRow = tableForm(sheet, tableIndexRow) #테이블 형식, 증감열, 비고열
                if case == 2 : 
                    sort_and_write_back(sheet, tableIndexRow + 1, sumRow)
                    row = sumRow
            row += 1
        
        # '평균 취득/처분 단가' 구하기
        row = 1
        while row <= last_row :
            value = sheet.range(f'A{row}').value  # 각 행의 A열 값 읽기
            if value == '회사명' :
                tableIndexRow = row + 3     
                case, deltaCol, priceCol, remarksCol, sumRow = tableForm(sheet, tableIndexRow) #테이블 형식, 증감열, 비고열

                # 1. '증감X취득/처분 단가' 열 값 채워넣고 합계 구하기
                # Case 1, 2 모두 동일
                for row_index in range(tableIndexRow + 1, sumRow) :
                    delta_cell_value = sheet.range(row_index, deltaCol).value
                    price_cell_value = sheet.range(row_index, priceCol).value

                    #증감이 숫자가 아니면 스킾
                    if not is_number(delta_cell_value) :
                        continue

                    # Extract numeric part from the price cell value
                    if isinstance(price_cell_value, str):
                        # Find all numeric parts and concatenate them
                        numbers = re.findall(r'[0-9]+', price_cell_value)
                        concatenated_number = ''.join(numbers)
                        if concatenated_number:
                            # Convert the concatenated number to a float
                            numeric_price = float(concatenated_number)
                        else:
                            numeric_price = 0  # Default to 0 if no numbers found

                            # Calculate the product and write it back to Excel
                        product = delta_cell_value * numeric_price
                        sheet.range(row_index, remarksCol + 1).value = product
                    else:
                        numeric_price = price_cell_value  # Use the value directly if it's already a number
                        # Calculate the product and write it back to Excel
                        #print('delta_cell_value:' + str(delta_cell_value) + ', numeric_price:' + str(numeric_price))
                        product = delta_cell_value * numeric_price
                        sheet.range(row_index, remarksCol + 1).value = product
                        delta_cell = sheet.range(row_index, deltaCol).get_address(0, 0)  # Address of the deltaCol cell
                        price_cell = sheet.range(row_index, priceCol).get_address(0, 0)  # Address of the priceCol cell
                        target_cell = (row_index, remarksCol + 1)  # Target cell for the result
                        sheet.range(target_cell).formula = f'=PRODUCT({delta_cell}, {price_cell})'  # Set the formula for multiplication
                
                # 2. 증감 합계, 취득/처분 단가 합계, 취득/처분 단가 평균 구하기
                # 증감 합계
                cell_value = sheet.range((sumRow, deltaCol)).value
                if not is_number(cell_value) : #증감 합계가 구해져 있지 않은 경우
                    start_cell = sheet.range((tableIndexRow + 1, deltaCol)).get_address(0, 0)  # Get the address of the start cell
                    end_cell = sheet.range((sumRow - 1, deltaCol)).get_address(0, 0)  # Get the address of the end cell
                    sheet.range((sumRow, deltaCol)).formula = f'=SUM({start_cell}:{end_cell})'  # Set the SUM formula
                    #print('Case 1 합계 구함')
                
                # 취득/처분 단가 합계
                start_cell = sheet.range((tableIndexRow + 1, remarksCol + 1)).get_address(0, 0)  # Get the address of the start cell
                end_cell = sheet.range((sumRow - 1, remarksCol + 1)).get_address(0, 0)  # Get the address of the end cell
                sheet.range((sumRow, remarksCol + 1)).formula = f'=SUM({start_cell}:{end_cell})'  # Set the SUM formula
                
                # 취득/처분 단가 평균
                # Get the addresses of the cells in Excel's A1 notation
                dividend_cell = sheet.range((sumRow, remarksCol + 1)).get_address(0, 0)
                divisor_cell = sheet.range((sumRow, deltaCol)).get_address(0, 0)
                
                # Set the formula for division
                sheet.range((sumRow, remarksCol)).formula = f'={dividend_cell}/{divisor_cell}'
                row = sumRow + 1
                '''
                elif case == 2 :             
                    print('Case 2 unsupported')
                    row += 1
                '''
            else :
                row += 1
        
        

        book.save()  # 변경 사항 저장
    finally:
        book.close()  # 파일 닫기
        app.quit()  # Excel 애플리케이션 종료


def main () :
    equityFolder = '2024.01.26_지분공시'  # Update the folder path
    xlsxFilePath = HTMLtoExcel(equityFolder)
    #xlsxFilePath = 'E:/bbAutomation/dartScraping/2024.01.18_지분공시/2024.01.18_지분공시_v1.xlsx'
    #xlsxFilePath = '/Users/yee/Documents/dartScraping/2024.01.18_지분공시/2024.01.18_지분공시_v1.xlsx'
    calculateAveragePrice(xlsxFilePath)
    #calculateFirstVersionExcel(xlsxFilePath)


main()