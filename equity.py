
import pandas as pd
import xlwings as xw
import os, sys
import numpy as np
import re

secondTableColumn = 15

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

def tableForm(sheet, indexRow) :
    case = None
    deltaCol = None
    priceCol = None
    remarksCol = None
    sumRow = None
    col = 1  # Start from the first colum
    while True :
        cell_value = sheet.range((indexRow, col)).value
        if col == 1:
            if '성명' in cell_value : #취득/처분 단가 2개 열
                case = 2
            elif cell_value == '보고사유' : #취득/처분 단가 1개 열
                case = 1
            else :
                print('새로운 폼! 처리 필요')
        if cell_value == '비 고' :
            remarksCol = col
            break
        elif cell_value == '증감' :
            deltaCol = col
        elif '취득/처분 단가' in cell_value :
            priceCol = col
        col += 1

    #'합 계' row 구하기    
    row = indexRow
    while True :
        value = sheet.range(f'A{row}').value
        if value == '합 계' :
            sumRow = row
            break
        row += 1
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

def adjustExcel(xlsxFilePath) : 
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
                addAveragePriceColumn(sheet, row+3)   
                row += 1  # '회사명' 행 다음부터 검사 시작
                while row <= last_row + 1:
                    # 현재 행 전체가 비어있는지 확인
                    row_values = sheet.range(f'{row}:{row}').value
                    if all(cell is None for cell in row_values) : # 비어 있는 행인지 확인
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
        
        # '평균 취득/처분 단가' 구하기
        row = 1
        while row <= last_row :
            value = sheet.range(f'A{row}').value  # 각 행의 A열 값 읽기
            if value == '회사명' :
                tableIndexRow = row + 3     
                case, deltaCol, priceCol, remarksCol, sumRow = tableForm(sheet, tableIndexRow) #테이블 형식, 증감열, 비고열
                if case == 1 :
                    # 1. '증감X취득/처분 단가' 열 값 채워넣고 합계 구하기
                    for row_index in range(tableIndexRow + 1, sumRow) :
                        delta_cell_value = sheet.range(row_index, deltaCol).value
                        price_cell_value = sheet.range(row_index, priceCol).value

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
                            #numeric_price = price_cell_value  # Use the value directly if it's already a number
                            # Calculate the product and write it back to Excel
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
                        print('Case 1 합계 구함')
                    
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
                elif case == 2 :             
                    print('Case 2 unsupported')
                    row += 1
            else :
                row += 1
        
        

        book.save()  # 변경 사항 저장
    finally:
        book.close()  # 파일 닫기
        app.quit()  # Excel 애플리케이션 종료


def main () :
    equityFolder = '2024.01.22_지분공시'  # Update the folder path
    xlsxFilePath = HTMLtoExcel(equityFolder)
    #xlsxFilePath = 'E:/bbAutomation/dartScraping/2024.01.18_지분공시/2024.01.18_지분공시_v1.xlsx'
    #xlsxFilePath = '/Users/yee/Documents/dartScraping/2024.01.18_지분공시/2024.01.18_지분공시_v1.xlsx'
    adjustExcel(xlsxFilePath)
    #calculateFirstVersionExcel(xlsxFilePath)


main()