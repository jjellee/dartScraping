import pandas as pd
import xlwings as xw
import os, sys
import numpy as np
import re
from itertools import groupby
from dateutil import parser
import datetime
import time

secondTableColumn = 14
thirdTableColumn = 22
fourthTableColumn = 42
num_push_row_down = 0

# Excel 열들에 대한 포맷을 설정하는 함수
def set_number_format_with_comma(sheet, columns, last_row):
    for column in columns:
        for row in range(1, last_row + 1):
            sheet.range(f'{column}{row}').number_format = '#,##0'

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

'''
def alphabet_to_number(alphabet):
    # 알파벳을 숫자로 변환 ('A' -> 1, 'B' -> 2, ...)
    return ord(alphabet) - 64
'''

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

def deleteRow_specificRange(sheet, start_row, end_row, start_col, end_col) :
    # start_row부터 end_row까지의 행들을 삭제하고, 아래 행들을 위로 이동
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            sheet.cell(row=row, column=col).value = sheet.cell(row=row + 1, column=col).value

    #print('start_row:'+str(start_row))
    #print('end_row:'+str(end_row))
    # end_row의 start_col부터 end_col까지의 셀 값 삭제
    for col in range(start_col, end_col + 1):
        sheet.cell(row=end_row, column=col).value = None


def push_row_down(sheet, row_number, start_col, end_col):
    #print(sheet.range((row_number, start_col), (row_number, end_col)).value)
    # 다음 행에 데이터가 있는지 확인
    next_row_data = sheet.range((row_number + 1, start_col), (row_number + 1, end_col)).value
    if any(cell is not None for cell in next_row_data):
        # 다음 행에 데이터가 있다면, 그 행도 아래로 밀기
        push_row_down(sheet, row_number + 1, start_col, end_col)

    # 현재 행의 데이터를 아래 행으로 복사
    for col in range(start_col, end_col + 1):
        cell_value = sheet.range((row_number, col)).value
        sheet.range((row_number + 1, col)).value = cell_value

    # 원래 행의 첫 번째 셀에 '합 계' 작성, 나머지 셀은 비워냄
    sheet.range((row_number, start_col)).value = '합 계'
    global num_push_row_down
    num_push_row_down += 1
    #print('num_push_row_down:' + str(num_push_row_down))
    for col in range(start_col + 1, end_col + 1):
        sheet.range((row_number, col)).value = None


'''
def push_row_down(sheet, row_number, start_col, end_col):
    # 마지막 행 번호를 찾기
    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end('up').row
    
    # 마지막 행부터 시작 행까지 역순으로 데이터를 한 행씩 아래로 이동
    for row in range(last_row, row_number, -1):
        for col in range(start_col, end_col + 1):
            cell_value_above = sheet.range((row - 1, col)).value
            sheet.range((row, col)).value = cell_value_above

    # 원래 행의 첫 번째 셀에 '합 계' 작성, 나머지 셀은 비워냄
    sheet.range((row_number, start_col)).value = '합 계'
    for col in range(start_col + 1, end_col + 1):
        sheet.range((row_number, col)).value = None
'''
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
    reporterTables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, startcol=secondTableColumn, index=True)
    shareRatioTables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, startcol=thirdTableColumn, index=True)

    # 원본 테이블 
    #tables[0].to_excel(excel_writer, sheet_name=sheet_name, startrow=start_row + 2, startcol=secondTableColumn, index=True)

    # Get the workbook and the sheet
    #workbook  = excel_writer.book
    worksheet = excel_writer.sheets[sheet_name]

    #tradeTables
    startRow = start_row + 5
    endRow = startRow + len(tradeTables[0])
    startCol = 1
    endCol = secondTableColumn - 1
    #print('회사명 row : ' + str(start_row + 1))
    deleteRow_specificRange(worksheet, startRow, endRow, startCol, endCol)
    
    #reporterTables
    startRow = start_row + 3
    endRow = startRow + len(reporterTables[0])
    startCol = secondTableColumn
    endCol = thirdTableColumn - 1
    deleteRow_specificRange(worksheet, startRow, endRow, startCol, endCol)
    
    #shareRatioTables
    if '임원' in company_submitter[2] : 
        startRow = start_row + 5
    else :
        startRow = start_row + 6
    endRow = startRow + len(shareRatioTables[0])
    startCol = thirdTableColumn
    endCol = fourthTableColumn - 1
    deleteRow_specificRange(worksheet, startRow, endRow, startCol, endCol)

    
    # 'Unnamed' 제거
    if '임원' in company_submitter[2] : 
        #column_letter = number_to_alphabet(thirdTableColumn+2)
        #print('cell : ' + str(start_row+3) + f':{column_letter}')
        #print('cell : ' + str(start_row+3) + ':' + str(thirdTableColumn+2))
        #rint(worksheet.cell(row=start_row + 3, column=thirdTableColumn+2).value)
        if 'Unnamed' in worksheet.cell(row=start_row + 3, column=thirdTableColumn+2).value :
            worksheet.cell(row=start_row + 3, column=thirdTableColumn+2).value = None
        if 'Unnamed' in worksheet.cell(row=start_row + 4, column=thirdTableColumn+2).value :
            worksheet.cell(row=start_row + 4, column=thirdTableColumn+2).value = None

    # Write company_submitter[0] and company_submitter[1] in specified cells using openpyxl method
    worksheet.cell(row=start_row + 1, column=2, value='회사명')  # Writing in the first cell of start_row
    worksheet.cell(row=start_row + 1, column=3, value=company_submitter[0])  # Writing in the first cell of start_row

    worksheet.cell(row=start_row + 1, column=5, value='보고서명')  # Writing in the first cell of start_row
    worksheet.cell(row=start_row + 1, column=6, value=company_submitter[2])  # Writing in the first cell of start_row

    worksheet.cell(row=start_row + 2, column=2, value='제출인')  # Writing in the first cell of the next row
    worksheet.cell(row=start_row + 2, column=3, value=company_submitter[1])  # Writing in the first cell of the next row

    #worksheet.delete_rows(start_row + 3)
    #worksheet.delete_rows(start_row + 5)

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
    xlsxFile = equityFolder + '_detail' + '.xlsx'
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

def addDeltaMultiplyPricetColumn(sheet, row) : 
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

def parse_custom_date_string(date_str):
    try :
        # '년', '월', '일'을 제거하고 '-'로 대체
        cleaned_date_str = re.sub(r'년|월', '-', date_str)
        cleaned_date_str = re.sub(r'일', '', cleaned_date_str)
        # 연속된 '-' 제거
        cleaned_date_str = re.sub(r'--', '-', cleaned_date_str)
        # 빈칸 제거
        cleaned_date_str = cleaned_date_str.strip()
        # 날짜 파싱
        return parser.parse(cleaned_date_str)
    except parser.ParserError:
        # 파싱 에러 발생 시 None 반환
        return None
    
def sort_and_write_back1(sheet, start_row, end_row, end_col):
    data = []  # 데이터를 저장할 리스트

    # 데이터 추출
    for row in range(start_row, end_row + 1):  # +1을 추가하여 end_row를 포함하도록 함
        row_data = []  # 현재 행의 데이터를 저장할 리스트
        for col in range(1, end_col):  # +1을 추가하여 secondTableColumn을 포함하도록 함
            row_data.append(sheet.range((row, col)).value)
        data.append(row_data)

    # 데이터 정렬 (A열 값으로 정렬)
    try :
        data.sort(key=lambda x: (x[0], parse_custom_date_string(x[1])))
    except Exception as e:
        print(f"Sorting error: {e}")
        print('sort only by name')
        data.sort(key=lambda x: x[0])
    # 정렬된 데이터를 시트에 다시 작성
    current_row = start_row
    for row_data in data:
        for col, value in enumerate(row_data, start=1):
            sheet.range((current_row, col)).value = value
        current_row += 1

def sort_and_write_back2(sheet, start_row, end_row, end_col):
    data = []  # 데이터를 저장할 리스트

    # 데이터 추출
    for row in range(start_row, end_row + 1):  # +1을 추가하여 end_row를 포함하도록 함
        row_data = []  # 현재 행의 데이터를 저장할 리스트
        for col in range(1, end_col):  # +1을 추가하여 secondTableColumn을 포함하도록 함
            row_data.append(sheet.range((row, col)).value)
        data.append(row_data)

    # 데이터 정렬 (A열 값으로 정렬)
    try :
        data.sort(key=lambda x: (x[0], parse_custom_date_string(x[2])))
    except Exception as e:
        print(f"Sorting error: {e}")
        print('sort only by name')
        data.sort(key=lambda x: x[0])

    # 정렬된 데이터를 시트에 다시 작성
    newTransactionEndRow = end_row
    current_row = start_row
    previous_value = data[0][0]  # 이전 행의 첫 번째 열의 값을 초기화
    #print(data)
    for row_data in data:
        value = row_data[0]  # 현재 행의 첫 번째 열의 값을 가져옴
        # 값이 이전 값과 다르면 행 이동 및 '합 계' 작성
        if value != previous_value:
            #print('current_row:' + str(current_row) + ' previous_value:' + previous_value + ' value:' + value )
            # current_row부터 end_row까지 한 행씩 아래로 이동
            push_row_down(sheet, current_row, 1, end_col)
            current_row += 1
            newTransactionEndRow += 1  # 행을 이동시켰기 때문에 end_row도 조정
        for col, cell_value in enumerate(row_data, start=1):
            sheet.range((current_row, col)).value = cell_value

        current_row += 1
        previous_value = value  # 이전 값 업데이트
    return newTransactionEndRow
    

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
        #print('last_row : ' + str(last_row))
        while row <= last_row :
            value = sheet.range(f'A{row}').value  # 각 행의 A열 값 읽기
            if value == '회사명':
                addDeltaMultiplyPricetColumn(sheet, row + 3)   
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
                    #sheet.api.Rows(row).Insert()  # 새 행 삽입
                    sheet.range(f'A{row}').value = '합 계'  # 새 행의 A열에 '합 계' 입력
                    #last_row += 1  # 행이 추가되었으므로 마지막 행 번호 업데이트
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
                            numeric_price = int(concatenated_number)
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

            else :
                row += 1

        book.save()  # 변경 사항 저장
    finally:
        book.close()  # 파일 닫기
        app.quit()  # Excel 애플리케이션 종료

def getReportType(sheet, row, col) :
    # 주어진 열 문자를 열 번호로 변환 (예: 'A' -> 1, 'B' -> 2, ..., 'E' -> 5)
    col_number = ord(col.upper()) - 64
    
    # 해당 셀의 값을 반환
    value = sheet.range(row, col_number).value
    if '대량' in value : #주식등의대량보유상황보고서
        return 2
    return 1 #임원주요주주특정증권등소유상황보고서


def getNumberOfBuyers(sheet, startRow, endRow) :
    unique_first_column_values = set()  # 첫 번째 열의 유니크한 값들을 저장할 집합

    # 데이터 추출 및 유니크한 값 계산
    for row in range(startRow, endRow):
        value = sheet.range((row, 1)).value  # 첫 번째 열의 값 추출
        unique_first_column_values.add(value)  # 집합에 추가

    return len(unique_first_column_values)  # 유니크한 첫 번째 열 값들의 개수 반환

def makeForm1(sheet, transactionStartRow) : 
    row = transactionStartRow
    while True :
        # 현재 행 A열의 값이 비어있는지 확인
        row_values = sheet.range(f'{row}:{row}').value        
        if row_values[0] is None :
            break
        row += 1
    transactionEndRow = row - 1
    tableEndValue = sheet.range(f'A{transactionEndRow}').value  # 테이블 끝행의 A열 값 읽기
    #print('Row' + str(row-1) + ':' + tableEndValue)
    
    # '합 계'행 존재 유무 확인
    if tableEndValue != '합 계':
        sheet.range(f'A{transactionEndRow + 1}').value = '합 계'  # 새 행의 A열에 '합 계' 입력
    
    row = transactionStartRow
    while True :
        # 현재 행 A열의 값이 '합 계'인지 확인
        row_values = sheet.range(f'{row}:{row}').value        
        if row_values[0] == '합 계' :
            break
        row += 1
    transactionEndRow = row - 1
    
    sort_and_write_back1(sheet, transactionStartRow, transactionEndRow, secondTableColumn)

    return transactionEndRow

def form2priceColOneCol(sheet, tableIndexRow) :
    col = 1
    leftPriceCol = None
    RightPriceCol = None
    while True :
        if '단가' in sheet.range((tableIndexRow, col)).value :
            leftPriceCol = col # 왼쪽 '취득/처분 단가'
            RightPriceCol = leftPriceCol + 1
            break
        col += 1
    
    transactionStartRow = tableIndexRow + 1
    row = transactionStartRow
    while True :
        #첫 번쨰열 값이 없으면 break
        fisrtColValue = sheet.range((row, 1)).value
        if fisrtColValue is None : 
            break
        leftCellValue = sheet.range((row, leftPriceCol)).value
        rightCellValue = sheet.range((row, RightPriceCol)).value

        #왼쪽 처분단가 열에 숫자가 없고, 오른쪽 처분단가 열에 숫자가 있다면
        #오른쪽 -> 왼쪽 복사
        convertedLeftCellValue = convertStringToNumber(leftCellValue)
        convertedRightCellValue = convertStringToNumber(rightCellValue)
        if (not is_number(convertedLeftCellValue) or convertedLeftCellValue == 0) and is_number(convertedRightCellValue) :
            if convertedRightCellValue != 0 :
                sheet.range((row, leftPriceCol)).value = sheet.range((row, RightPriceCol)).value
                print('row '+ str(row) + ' copy ' + str(rightCellValue) + ' to ' + str(leftCellValue))
        row += 1
        

def makeForm2(sheet, transactionStartRow) : 
    row = transactionStartRow
    #transactionEndRow 구하기
    while True :
        # 현재 행 A열의 값만 확인
        first_cell_value = sheet.range((row, 1)).value  # 1은 A열을 의미
        if first_cell_value is None:
            break
        row += 1
    transactionEndRow = row - 1

    # '처분 단가' (왼쪽)열 하나로
    tableIndexRow = transactionStartRow - 1
    form2priceColOneCol(sheet, tableIndexRow)

    RowsToAdd = getNumberOfBuyers(sheet, transactionStartRow, transactionEndRow)
    row = transactionStartRow
    newTransactionEndRow = transactionEndRow
    if RowsToAdd > 0 and RowsToAdd < 50:
        print('RowsToAdd:'+str(RowsToAdd))
        # row 행 전체가 빌 때까지 row를 1씩 증가
        while True:
            # 현재 행 전체를 가져옴
            row_values = sheet.range(f"{row}:{row}").value
            # 현재 행이 비어있는지 확인 (모든 셀이 None 또는 빈 문자열인지)
            if all(cell is None or cell == '' for cell in row_values):
                break
            else:
                row += 1
        sheet.range(f"{row}:{row + RowsToAdd}").api.Insert(Shift=1)
        newTransactionEndRow = sort_and_write_back2(sheet, transactionStartRow, transactionEndRow, secondTableColumn)
    # 테이블 마지막 행에 '합 계'
    sheet.range((newTransactionEndRow + 1, 1)).value = '합 계'
    #print('endRow:' + str(endRow))
    # 반환 row : '합 계' 행
    return newTransactionEndRow, RowsToAdd


def getForm1Detail(sheet, tableIndexRow) :
    deltaCol = None
    priceCol = None
    remarksCol = None
    productCol = None
    sumRows = []
    endRow = None

    #'합 계' row 구하기  : sumRow  
    row = tableIndexRow
    while True :
        value = sheet.range(f'A{row}').value
        if value == '합 계' :
            sumRows.append(row)
        elif value is None :
            break
        row += 1
    endRow = sumRows[-1]
    # deltaCol, priceCol, productCol, remarksCol 구하기
    col = 1  # Start from the first colum
    while True :
        cell_value = sheet.range((tableIndexRow, col)).value
        if cell_value == '증감' :
            deltaCol = col
        elif '취득/처분 단가' in cell_value :
            priceCol = col
        elif cell_value == '비 고' : # 마지막 열
            remarksCol = col
            break
        col += 1
    productCol = remarksCol + 1

    return deltaCol, priceCol, productCol, remarksCol, endRow, sumRows

def getForm2Detail(sheet, tableIndexRow) :
    deltaCol = None
    priceCol = None
    remarksCol = None
    productCol = None
    sumRows = []
    endRow = None

    #'합 계' row 구하기  : sumRow  
    row = tableIndexRow
    while True :
        value = sheet.range(f'A{row}').value
        if value == '합 계' :
            sumRows.append(row)
        elif value is None :
            break
        row += 1
    endRow = sumRows[-1]
    # deltaCol, priceCol, productCol, remarksCol 구하기
    col = 1  # Start from the first colum
    while True :
        cell_value = sheet.range((tableIndexRow, col)).value
        if cell_value == '증감' :
            deltaCol = col
        elif '취득/처분 단가' in cell_value :
            priceCol = col
        elif cell_value == '비 고' : # 마지막 열
            remarksCol = col
            break
        col += 1
    priceCol = priceCol - 1 #왼쪽 '취득/처분 단가'
    productCol = remarksCol + 1

    return deltaCol, priceCol, productCol, remarksCol, endRow, sumRows

def convertStringToNumber(cell_value) :
    numeric = None
    if isinstance(cell_value, str):
        # Find all numeric parts and concatenate them
        numbers = re.findall(r'[0-9]+', cell_value)
        concatenated_number = ''.join(numbers)
        if concatenated_number:
            # Convert the concatenated number to a float
            numeric = int(concatenated_number)
    elif isinstance(cell_value, float) :
        numeric = cell_value
    return numeric

def update_sums_in_table(sheet, sumRows, deltaCol, transactionStartRow):
    for i in range(len(sumRows)):
        start_row = sumRows[i - 1] + 1 if i > 0 else transactionStartRow
        end_row = sumRows[i] - 1
        start_cell = sheet.range((start_row, deltaCol)).get_address(0, 0)  # Get the address of the start cell
        end_cell = sheet.range((end_row, deltaCol)).get_address(0, 0)  # Get the address of the end cell
        sheet.range((sumRows[i], deltaCol)).formula = f'=SUM({start_cell}:{end_cell})'  # Set the SUM formula
        '''
        sum_value = 0
        for row in range(start_row, end_row + 1):
            cell_value = sheet.range((row, deltaCol)).value
            if is_number(cell_value):
                sum_value += cell_value
        if sheet.range((sumRows[i], deltaCol)).value is None :
            sheet.range((sumRows[i], deltaCol)).value = sum_value
        '''


def update_delta_product_price_col_in_table(sheet, sumRows, deltaCol, priceCol, productCol, transactionStartRow) : 
    for i in range(len(sumRows)):
        start_row = sumRows[i - 1] + 1 if i > 0 else transactionStartRow
        end_row = sumRows[i] - 1
        sum_value = 0
        for row in range(start_row, end_row + 1):
            '''
            delta_cell_value = sheet.range(row, deltaCol).value
            price_cell_value = sheet.range(row, priceCol).value
            
            if is_number(delta_cell_value) and is_number(price_cell_value) :
                product = delta_cell_value * price_cell_value
                sum_value += cell_value
        if sheet.range((sumRows[i], deltaCol)).value is None :
            sheet.range((sumRows[i], deltaCol)).value = sum_value
            '''
            delta_cell = sheet.range(row, deltaCol).get_address(0, 0)  # Address of the deltaCol cell
            price_cell = sheet.range(row, priceCol).get_address(0, 0)  # Address of the priceCol cell
            target_cell = (row, productCol)  # Target cell for the result
            sheet.range(target_cell).formula = f'=PRODUCT({delta_cell}, {price_cell})' 
        
        start_cell = sheet.range((start_row, productCol)).get_address(0, 0)  # Get the address of the start cell
        end_cell = sheet.range((end_row, productCol)).get_address(0, 0)  # Get the address of the end cell
        sheet.range((sumRows[i], productCol)).formula = f'=SUM({start_cell}:{end_cell})'  # Set the SUM formula

        # 취득/처분 단가 평균
        # Get the addresses of the cells in Excel's A1 notation
        dividend_cell = sheet.range((sumRows[i], productCol)).get_address(0, 0)
        divisor_cell = sheet.range((sumRows[i], deltaCol)).get_address(0, 0)
        
        # Set the formula for division
        sheet.range((sumRows[i], productCol-1)).formula = f'={dividend_cell}/{divisor_cell}'

def calculateForm1(sheet, tableIndexRow) :   
    deltaCol, priceCol, productCol, remarksCol, endRow, sumRows = getForm1Detail(sheet, tableIndexRow) #테이블 형식, 증감열, 비고열

    # 1.'처분 단가'열에서 숫자가 아닌 부분 삭제
    for row in range(tableIndexRow + 1, endRow) :
        cell_value = sheet.range(row, priceCol).value
        sheet.range(row, priceCol).value = convertStringToNumber(cell_value)

    # 2.sumRows 리스트의 각 멤버들 사이의 '증감' 열 값들의 합을 계산하고, 해당 값을 각 sumRows 멤버의 셀에 저장합니다.
    update_sums_in_table(sheet, sumRows, deltaCol, tableIndexRow + 1)

    # 3.'증감X취득/처분 단가' 열 값 채워넣고 합계 구하기
    update_delta_product_price_col_in_table(sheet, sumRows, deltaCol, priceCol, productCol, tableIndexRow + 1)

    return sumRows[-1]

def calculateForm2(sheet, tableIndexRow) : 
    deltaCol, priceCol, productCol, remarksCol, endRow, sumRows = getForm2Detail(sheet, tableIndexRow) #테이블 형식, 증감열, 비고열

    # 1.'처분 단가'열에서 숫자가 아닌 부분 삭제
    for row in range(tableIndexRow + 1, endRow) :
        cell_value = sheet.range(row, priceCol).value
        sheet.range(row, priceCol).value = convertStringToNumber(cell_value)

    # 2.sumRows 리스트의 각 멤버들 사이의 '증감' 열 값들의 합을 계산하고, 해당 값을 각 sumRows 멤버의 셀에 저장합니다.
    update_sums_in_table(sheet, sumRows, deltaCol, tableIndexRow + 1)

    # 3.'증감X취득/처분 단가' 열 값 채워넣고 합계 구하기
    update_delta_product_price_col_in_table(sheet, sumRows, deltaCol, priceCol, productCol, tableIndexRow + 1)
    print(sheet.range(tableIndexRow-3, 2).value)
    return sumRows[-1]



def improvement_calculateAveragePrice(xlsxFilePath) :
    app = xw.App(visible=False)  # Excel 애플리케이션을 보이지 않게 설정
    book = app.books.open(xlsxFilePath)  # Excel 파일 열기
    try :
        sheet = book.sheets[0]  # 첫 번째 시트 선택
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row  # 첫 번째 열의 마지막 행 찾기
        #print('last_row : ' + str(last_row))
        row = 1
        while row <= last_row :
            value = sheet.range(f'A{row}').value  # 각 행의 A열 값 읽기
            if value == '회사명':
                print(sheet.range(row, 2).value)
                print(sheet.range(row+1, 2).value)
                print(sheet.range(row, 5).value)
                form = getReportType(sheet, row, 'E')
                if form == 1 :
                    row = makeForm1(sheet, row+4)
                elif form == 2 :
                    global num_push_row_down
                    num_push_row_down = 0
                    row, RowsToAdd = makeForm2(sheet, row+4)
                    last_row += RowsToAdd
            row += 1

        # '증감' 합, '증감X취득/처분 단가'합, 평균 단가 구하기
        row = 1
        while row <= last_row :
            value = sheet.range(f'A{row}').value  # 각 행의 A열 값 읽기
            if value == '회사명':
                addDeltaMultiplyPricetColumn(sheet, row + 3)  # '증감X취득/처분 단가' 열 추가 (Form1&2 공통)
                form = getReportType(sheet, row, 'E')
                # 임원주요주주특정증권등소유상황보고서 Form1 : 매매주체 1인
                if form == 1 :
                    row = calculateForm1(sheet, row + 3) # '합 계'행 존재 유무 확인
                # 주식등의대량보유상황보고서 Form2 : 매매주체 다수
                elif form == 2:
                    row = calculateForm2(sheet, row + 3) # 마지막 '합 계'가 포함된 행
            row += 1
        
        # D, E, F, G, H, I, K, L 열의 포맷을 '숫자'로 설정하고, '1000단위 구분기호' 추가
        set_number_format_with_comma(sheet, ['D', 'E', 'F', 'G', 'H', 'I', 'K', 'L'], last_row)

        book.save()  # 변경 사항 저장
    finally:
        book.close()  # 파일 닫기
        app.quit()  # Excel 애플리케이션 종료
'''
def writeSummaryForm1(sheet_d, sheet_s, row) :
    companyName = sheet_d.range(row, 1).value
    submitter = sheet_d.range(row+1, 1).value
    tableIndexRow = row + 3
    
    todayDate = datetime.now().strftime("%Y-%m-%d")
    startDate, endDate, transactionCount = getForm1StartEndDate() #매수주체는 1명

    [companyName, todayDate, submitter]
'''
def writeSummaryFile(equityFolder, detailFilePath) :
    # 공시 회사, 공시일, 변동일(S), 변동일(E), S~E (수), 매매, 공시주체, 이름, 출생년도, 수량(증감), 변동후, 지분율, 단가, 총액, 비고
    # 임원주요주주특정증권등소유상황보고서
    # 주식등의대량보유상황보고서
    summaryFileName = equityFolder + '_summary' + '.xlsx'
    summaryFilePath = os.path.join(equityFolder, summaryFileName)

    app = xw.App(visible=False)  # Excel 애플리케이션을 보이지 않게 설정
    book_d = None
    book_s = None

    try:
        book_d = app.books.open(detailFilePath)
        sheet_d = book_d.sheets[0]
        # 파일이 존재하는지 확인하고, 없으면 새로 생성합니다.
        if not os.path.exists(summaryFilePath):
            book_s = app.books.add()  # 새로운 책을 추가합니다.
            book_s.save(summaryFilePath)  # 파일을 저장합니다.
        else:
            book_s = app.books.open(summaryFilePath)


        # 요약 파일 첫 행 입력
        sheet_s = book_s.sheets[0]
        # 첫 번째 행에 헤더들을 입력합니다.
        headers = ['공시 회사', '공시일', '변동일(S)', '변동일(E)', 'S~E (수)', '매매', '공시주체', '이름', '출생년도', '수량(증감)', '변동후', '지분율', '단가', '총액', '비고']
        sheet_s.range('A1').value = headers

        row_d = 1
        row_s = 2
        '''
        while row <= last_row :
            value = sheet_d.range(f'A{row}').value  # 각 행의 A열 값 읽기
            if value == '회사명':
                addDeltaMultiplyPricetColumn(sheet_d, row + 3)  # '증감X취득/처분 단가' 열 추가 (Form1&2 공통)
                form = getReportType(sheet_d, row, 'E')
                # 임원주요주주특정증권등소유상황보고서 Form1 : 매매주체 1인
                if form == 1 :
                    row_s = writeSummaryForm1(sheet_d, sheet_s, row_d, row_s) 
                # 주식등의대량보유상황보고서 Form2 : 매매주체 다수
                elif form == 2:
                    row_s = writeSummaryForm2(sheet_d, sheet_s, row_d, row_s) 
                    last_row += num_push_row_down
                    #print('last_row : ' + str(last_row))
            row += 1
        '''

        book_d.save()  # 변경 사항 저장
        book_s.save()  # 변경 사항 저장
    except Exception as e:
        print("An error occurred:", e)
    finally:
        # 파일이 열려 있으면 닫습니다.
        if book_d:
            book_d.close()
        if book_s:
            book_s.close()
        app.quit()  # Excel 애플리케이션 종료

def main () :
    equityFolder = '2024.02.02_지분공시'  # Update the folder path
    xlsxFilePath = HTMLtoExcel(equityFolder)
    
    #calculateAveragePrice(xlsxFilePath)
    
    improvement_calculateAveragePrice(xlsxFilePath)
    #xlsxFilePath = '/Users/yee/Documents/dartScraping/2024.01.31_지분공시/2024.01.31_지분공시_detail.xlsx'
    #xlsxFilePath = 'E:/bbAutomation/dartScraping/2024.01.31_지분공시/2024.01.31_지분공시_detail.xlsx'
    writeSummaryFile(equityFolder, xlsxFilePath)

# 시작 시간 기록
start_time = time.time()

main()

# 종료 시간 기록
end_time = time.time()

# 실행 시간 계산
execution_time = end_time - start_time
print(f"Execution time: {execution_time} seconds")