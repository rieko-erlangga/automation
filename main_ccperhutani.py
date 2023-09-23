print('\nPART 1: AUTO UPLOAD EXCEL')

from undetected_chromedriver import ChromeOptions
from undetected_chromedriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta
from time import sleep

#------------------------------FUNCTION LISTS-------------------------------

def located_element(selector:str, element:str):
    sleep(1) # Add delay so element uploaded successfully
    return driver.find_element(selector, element)

def autofill_textbox(element:str, value:str):
    located_element('name', element).send_keys(value + Keys.ENTER)

def auto_click_recaptcha():
    try:
        recaptcha_frame = located_element('tag name', 'iframe')
        driver.switch_to.frame(recaptcha_frame)
        located_element('class name', 'rc-anchor').click()
        driver.switch_to.default_content()
    except:
        auto_click_recaptcha()

def auto_click_button(element:str):
    located_element('class name', element).click()

def auto_select_menu(attribut:str, element:str, value:str):
    located_element(attribut, element).send_keys(value + Keys.ARROW_DOWN + Keys.ENTER)
    
def auto_hover_menu(element:str):
    menu = located_element('id', element)
    ActionChains(driver).move_to_element(menu).perform()

def auto_click_menu(element:str):
    menu = located_element('id', element)
    ActionChains(driver).move_to_element(menu).click().perform()

def auto_filter_date_time():
    current_time = datetime.now().time()
    
    if current_time < datetime.strptime('22:50', '%H:%M').time() \
    or current_time > datetime.strptime('23:59', '%H:%M').time():
        print('Time will be set for yesterday!')
        previous_date = (datetime.now() - timedelta(days=1)) #

        located_element('name', 'start_date').send_keys(
            previous_date.strftime('%d%m')
            
            # Use one of them if the date selection is jumbled
            #previous_date.strftime('%d')
            #+ Keys.RIGHT + previous_date.strftime('%m')
            + Keys.RIGHT + previous_date.strftime('%Y')
            + Keys.RIGHT + '000001'
            )

        located_element('name', 'end_date').send_keys(
            previous_date.strftime('%d%m')
            
            # Use one of them if the date selection is jumbled
            #previous_date.strftime('%d')
            #+ Keys.RIGHT + previous_date.strftime('%m')
            + Keys.RIGHT + previous_date.strftime('%Y')
            + Keys.RIGHT + '235959'
            )

def auto_download_file(url:str, attribut:str, element:str, value:str):
    driver.get(url)
    auto_filter_date_time()

    if url != 'https://SECRET':
        auto_select_menu(attribut, element, value)

    element = 'css-sg6n02'
    auto_click_button(element)
    try:
        element = 'fm-tab-export'
        auto_hover_menu(element)
        element = 'fm-tab-export-excel'
        auto_click_menu(element)
    except:
        print('CWC is still empty!')

    sleep(4)

#------------------------------IMPLEMENTATION-------------------------------

options = ChromeOptions()

# To avoid having to log in again, use Google Chrome Portable
chrome_path = \
'C:/SECRET/GoogleChromePortable64/GoogleChromePortable.exe'
options.binary_location = chrome_path

argument_1 = '--no-sandbox'
options.add_argument(argument_1)

argument_2 = '--app=https://accounts.google.com/'
options.add_argument(argument_2)

argument_3 = '--headless'
#options.add_argument(argument_3)

argument_4 = '--start-maximized'
options.add_argument(argument_4)

prefs = {
    'credentials_enable_service': False,
    'profile.password_manager_enable': False
}
options.add_experimental_option('prefs', prefs)

print('Create undetected Chrome instance')
driver = Chrome(options)

try:

    print('Add delay so not detected as bot')
    timeout = 60
    driver.implicitly_wait(timeout)
    '''
    print('Sign in to Google account that already passed recaptcha')
    value = 'https://accounts.google.com/'
    driver.get(value)

    print('> Autofill username')
    element = 'identifier'
    value = 'SECRET@gmail.com'
    autofill_textbox(element, value)

    print('> Autofill password')
    element = 'Passwd'
    value = 'SECRET'
    autofill_textbox(element, value)

    sleep(1) # Add time and manual confirmation first when Google add extra security
    '''
    print("Navigate to to the company's omnichannel")
    value = 'https://SECRET'
    driver.get(value)
    '''
    print('> Autofill tenant')
    element = 'tenant'
    value = 'SECRET'
    autofill_textbox(element, value)

    print('> Autofill username')
    element = 'username'
    value = 'SECRET'
    autofill_textbox(element, value)

    print('> Autofill password')
    element = 'password'
    value = 'SECRET'
    autofill_textbox(element, value)

    print('Try to bypass recaptcha')
    auto_click_recaptcha()

    sleep(1)

    element = 'MuiButton-root'
    auto_click_button(element)
    print('Successfully login')
    '''
    print('Download daily reports - if any')
    url = 'https://SECRET'
    attribut = 'name'
    element = 'type_date'
    value = 'date start interaction'
    auto_download_file(url, attribut, element, value)

    print('Download answered calls - if any')
    url = 'https://SECRET'
    attribut = 'id'
    element = 'combo-box-demo'
    value = 'Answered'
    auto_download_file(url, attribut, element, value)

    print('Download abandoned calls - if any')    
    url = 'https://SECRET'
    attribut = 'id'
    element = 'combo-box-demo'
    value = 'Abandoned'
    auto_download_file(url, attribut, element, value)

    print('Download manual interaction - if any')
    url = 'https://SECRET'
    attribut = 'name'
    element = 'type_date'
    value = 'date start'
    auto_download_file(url, attribut, element, value)

    print('Download outbound calls - if any')
    url = 'https://SECRET'
    auto_download_file(url, attribut=None, element=None, value=None)

    print("Add time to ensure all file's uploaded successfully")
    sleep(3)
    
finally:

    driver.close()
    driver.quit()


print('\nPART 2: AUTO EDIT EXCEL')

import pandas
import xlwings
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from collections import Counter
from datetime import timedelta
from datetime import datetime
from datetime import date
import os

#------------------------------FUNCTION LISTS-------------------------------

Visible = False

file_name = 'report_service_ticket.xlsx'
file_name_2 = 'report_service_voice.xlsx'
file_name_3 = 'report_service_voice (1).xlsx'
file_name_4 = 'report_service_interaction.xlsx'
file_name_5 = 'report_sales_cal_log.xlsx'

sheet_name = 'Flexmonster Pivot Table'
sheet_name_2 = 'Data Tabulation'
sheet_name_3 = 'Pivot Table'
sheet_name_4 = 'Answered Calls'
sheet_name_5 = 'Abandoned Calls'
sheet_name_6 = 'Manual'
sheet_name_7 = 'Outbound Calls'

destination_column = 'A:A'
fill_color = (0, 0, 0)
fill_color_2 = (233, 233, 233)

def sheet(sheet_name:str):
    return workbook.sheets[sheet_name]

def sheet_range(sheet_name:str, set_ranges:str):
    return sheet(sheet_name).range(set_ranges)

def delete_columns(sheet_name:str, set_ranges:str):
    sheet_range(sheet_name, set_ranges).api.Delete()

def insert_column(sheet_name:str, destination_column:str, direction:str):
    sheet_range(sheet_name, destination_column).insert(direction)

def cell_formula(sheet_name:str, target_cell:str, fill_formula:str):
    sheet_range(sheet_name, target_cell).formula = fill_formula

def last_row(sheet_name:str):
    return sheet_range(sheet_name, 'A1').end('down').row

def move_column(sheet_name:str, set_ranges:str, destination_column:str, direction:str):
    sheet_range(sheet_name, set_ranges).api.Cut()
    insert_column(sheet_name, destination_column, direction)

def align_center(sheet_name:str, set_ranges:str):
    sheet_range(sheet_name, set_ranges).api.HorizontalAlignment = -4108

def count_rows(sheet_name:str, target_cell:str):
    cell_formula(sheet_name, target_cell, f'=SUBTOTAL(3,A2:A{last_row(sheet_name)})')
    align_center(sheet_name, target_cell)

def give_cell_values(sheet_name:str, set_cells:str, value:str):
    sheet_range(sheet_name, set_cells).value = value

def create_rt_aht(sheet_name:str, set_ranges:str, destination_column:str, direction:str,
                  set_cells:str, set_title:str, column_1:str, column_2:str):

    move_column(sheet_name, set_ranges, destination_column, direction)
    give_cell_values(sheet_name, set_cells, set_title)

    for row in range(2, last_row(sheet_name)+1):
        cell_formula(sheet_name, f'A{row}', f'={column_1}{row}-{column_2}{row}')

def create_average(sheet_name:str, target_cell:str, get_cell:str):
    result = f'=TEXT(SUBTOTAL(1,{get_cell}2:{get_cell}{last_row(sheet_name)}), "h:mm:ss") & " - " \
             & TEXT(SUBTOTAL(1,{get_cell}2:{get_cell}{last_row(sheet_name)})*86400, "#0") & " seconds"'
        
    cell_formula(sheet_name, target_cell, result)

def merge_cell(sheet_name:str, set_ranges:str):
    sheet_range(sheet_name, set_ranges).merge()

def range_color(sheet_name:str, set_ranges:str, fill_color:tuple):
    sheet_range(sheet_name, set_ranges).color = fill_color

def text_color(sheet_name:str, set_ranges:str):
    sheet_range(sheet_name, set_ranges).api.Font.Color = 16777215   

def format_column(sheet_name:str, set_ranges:str, set_formats:str):
    sheet_range(sheet_name, set_ranges).number_format = set_formats

def wrap_text_sheet(sheet_name:str, set_ranges:str):
    sheet(sheet_name).api.UsedRange.RowHeight = 15
    sheet_range(sheet_name, set_ranges).api.WrapText = False

def hide_gridlines(sheet_name:str):
    sheet(sheet_name).api.Application.ActiveWindow.DisplayGridlines = False

def auto_filter(sheet_name:str, set_ranges:str):
    sheet_range(sheet_name, set_ranges).api.AutoFilter(1)

def rename_sheet(sheet_name:str, sheet_name_2:str):
    sheet(sheet_name).name = sheet_name_2

def new_sheet(set_value:str):
    workbook.sheets.add().name = set_value

def locate_row(sheet_name_3:str, set_ranges:str, value:str):
    return [row.row for row in sheet_range(
        sheet_name_3, set_ranges) if row.value and value in row.value]

def sorted_row(sheet_name:str, set_range:str, value, reverse:bool):
    return sorted(locate_row(sheet_name, set_range, value), reverse = reverse)

def delete_rows(sheet_name_3:str, value:str):
    for row in sorted_row(sheet_name_3, f'A2:A{last_row(sheet_name_3)}', value, True):
        sheet(sheet_name_3).api.Rows(row).Delete()

def auto_pivot_table(file_name:str, sheet_name_2:str, value:str, sheet_name_3:str):
    data_frame = pandas.read_excel(file_name, sheet_name_2, engine='openpyxl')
    pivot_table = pandas.pivot_table(data_frame, value, aggfunc='count', index='category_name')
    book = load_workbook(file_name)
    sheet = book[sheet_name_3]

    for row in dataframe_to_rows(pivot_table, index=True, header=True):
        sheet.append(row)

    sheet.delete_rows(2)
    book.save(file_name)
    book.close()

def delete_cell_value(sheet_name:str):
    for row in sorted_row(sheet_name, f'B2:B{last_row(sheet_name)}', 'Manual', True):
        sheet(sheet_name).range(f'C{row}').value = 0

def get_cell_values(sheet_name_3:str, set_cell:str):
    return sheet_range(sheet_name_3, set_cell).value

def add_tooltip(sheet_name:str, tooltip_text_1:str, tooltip_text_2:str):
    sheet(sheet_name)[f'B{last_row(sheet_name)+1}'].api.AddComment(tooltip_text_1)
    sheet(sheet_name)[f'C{last_row(sheet_name)+1}'].api.AddComment(tooltip_text_2)
    sheet(sheet_name)[f'D{last_row(sheet_name)+1}'].api.AddComment(tooltip_text_2)

def title_table(sheet_name_3:str, set_ranges:str, set_title:str, fill_color:str):
    give_cell_values(sheet_name_3, set_ranges, set_title)
    align_center(sheet_name_3, set_ranges)
    range_color(sheet_name_3, set_ranges, fill_color)
    text_color(sheet_name_3, set_ranges)

def delete_redundant_category(sheet_name_3:str, values:list):
    for value in values:
        for row in locate_row(sheet_name_3, f'A2:A{last_row(sheet_name_3)}', value):
            cell_value = get_cell_values(sheet_name_3, f'A{row}').replace('-', '>')
            give_cell_values(sheet_name_3, f'A{row}', cell_value[12:].strip())

def partial_delete(sheet_name_3:str):
    sheet_range(sheet_name_3, f'A7:B{last_row(sheet_name_3)}').api.Delete()

def get_cell_length(sheet_name_3:str, set_cell:str):
    cell_value = get_cell_values(sheet_name_3, set_cell)
    return len(cell_value) if cell_value is not None else 0

def max_cell_length(sheet_name_3:str):
    cell_values = [get_cell_length(sheet_name_3, f'A{row}') for row in range(2, 7)]
    return max(cell_values) - 5

def set_row_width(sheet_name_3:str, set_cell:str):
    sheet_range(sheet_name_3, set_cell).column_width = max_cell_length(sheet_name_3)

def select_row(sheet_name:str, set_ranges:str):
    sheet(sheet_name).api.Range(set_ranges).Select()    

def split_row(sheet_name_2:str, set_ranges:str):
    select_row(sheet_name_2, set_ranges)
    sheet(sheet_name_2).api.Application.ActiveWindow.SplitRow = 1

def freeze_top_row(sheet_name_2:str, set_ranges:str):
    split_row(sheet_name_2, set_ranges)
    sheet(sheet_name_2).api.Application.ActiveWindow.FreezePanes = True

def sort_column(sheet_name_3:str, key:str, order:int, orientation:int):
    sheet_range(sheet_name_3, f'A2:I{last_row(sheet_name_3)}').api.Sort(
        Key1 = sheet_range(sheet_name_3, f'{key}2').api, Order1=order, Orientation=orientation)

def disable_headings(sheet_name:str):
    sheet(sheet_name).book.app.api.ActiveWindow.DisplayHeadings = False

def reverse_sheet(sheet_names:list):
    for sheet_name in sheet_names:
        sheet(sheet_name).api.Move(Before = workbook.sheets[0].api)

def font_bold(sheet_name_4:str, set_range:str):
    sheet_range(sheet_name_4, set_range).font.bold = True

def set_column_width(sheet_name_4:str, set_cell:str, value:str):
    sheet_range(sheet_name_4, set_cell).column_width = len(value)

def cell_values():
    return [('A1', '    date_time    '),
            ('B1', 'date_time_connect'),
            ('C1', '  date_time_end  ')]

def format_table(sheet_name:str, cell_values:list):
    for cell, value in cell_values:
        give_cell_values(sheet_name, cell, value)
        set_column_width(sheet_name, cell, value)
        range_color(sheet_name, cell, fill_color_2)
        font_bold(sheet_name, cell)

def title_table_2(sheet_name_4:str):
    cell_value_2 = [
        *cell_values(),
        ('D1', '  unique_id  '),
        ('E1', '    cl_id    '),
        ('F1', 'wait_time'),
        ('G1', '      talk_time     '),
        ('H1', 'ring_time'),
        ('I1', '  dst  ')]

    format_table(sheet_name_4, cell_value_2)

def title_table_3(sheet_name_5:str):
    cell_value_3 = [
        *cell_values(),
        ('D1', '   event   '),
        ('E1', ' unique_id '),
        ('F1', '   cl_id   '),
        ('G1', 'wait_time')]

    format_table(sheet_name_5, cell_value_3)

def title_table_5(sheet_name_7:str):
    cell_value_4 = [
        ('A1', 'status_call     '),
        ('B1', 'reason_call    '),
        ('C1', '    date_start   '),
        ('D1', '   date_answer   '),
        ('E1', '    date_end     '),
        ('F1', '  agent  '),
        ('G1', '      to_id     '),
        ('H1', 'duration')]

    format_table(sheet_name_7, cell_value_4)

def add_tooltip_2(sheet_name_4:str, target_cell:str, tooltip_text:str):
    sheet(sheet_name_4)[target_cell].api.AddComment(tooltip_text)

def delete_redundant_session_id(sheet_name:str):
    for row in range(2, last_row(sheet_name)+1):
        row_target = f'AB{row}:AC{row}'
        row_target_2 = f'AB{row-1}:AC{row-1}'
        if get_cell_values(sheet_name, row_target) == get_cell_values(sheet_name, row_target_2):
              sheet(sheet_name).api.Rows(row).Delete()

def remove_seconds(datetime_list):
    return [datetime.replace(second=0, microsecond=0) for datetime in datetime_list]

def delete_rows_2(sheet_name:str, rows_to_delete:list):
    rows_to_delete.reverse()
    for row in rows_to_delete:
        sheet(sheet_name).api.Rows(row).Delete()

def delete_inconsistence_voice(sheet_name_2:str, sheet_name_4:str):
    channel_name_value = sheet_range(sheet_name_2, f'B2:B{last_row(sheet_name_2)}').value
    voice_amount = channel_name_value.count('Voice')

    date_start_interaction_value = sheet_range(sheet_name_2, f'M2:M{last_row(sheet_name_2)}').value
    date_start_interaction_no_seconds = remove_seconds(date_start_interaction_value)

    date_time_connect = sheet_range(sheet_name_4, f'B2:B{last_row(sheet_name_4)}').value
    date_time_connect_no_seconds = \
    remove_seconds(date_time_connect) if isinstance(date_time_connect, list) else [remove_seconds([date_time_connect])[0]]

    deleted_rows = set()
    rows_to_delete_sheet_2 = []
    for row, (channel_name, date_start) in enumerate(zip(channel_name_value, date_start_interaction_no_seconds)):
        if channel_name == 'Voice' and date_start != date_time_connect and date_start_interaction_no_seconds.count(date_start) > 1:
            if date_start not in deleted_rows:
                rows_to_delete_sheet_2.append(row+2)
                deleted_rows.add(date_start)
        
    rows_to_delete_sheet_4 = [row+2 for row, date_time in enumerate(date_time_connect_no_seconds)
                              if date_time not in date_start_interaction_no_seconds]

    if voice_amount > last_row(sheet_name_4)-1:
        delete_rows_2(sheet_name_2, rows_to_delete_sheet_2)
    elif voice_amount < last_row(sheet_name_4)-1:
        delete_rows_2(sheet_name_4, rows_to_delete_sheet_4)

def voice_data(sheet_name_2:str, sheet_name_4:str):
    rows_with_voice = sorted_row(sheet_name_2, f'B2:B{last_row(sheet_name_2)}', 'Voice', True)

    for get_row, set_row in enumerate(rows_with_voice):
        # date_open = date_time
        sheet_range(sheet_name_2, f'N{set_row}').value = sheet_range(sheet_name_4, f'A{get_row + 2}').value
        # date_start_interaction = date_pickup_interaction = date_time_connect
        sheet_range(sheet_name_2, f'M{set_row}').value = sheet_range(sheet_name_2, f'W{set_row}').value \
        = sheet_range(sheet_name_4, f'B{get_row + 2}').value
        # date_close = date_time_end
        sheet_range(sheet_name_2, f'O{set_row}').value = sheet_range(sheet_name_4, f'C{get_row + 2}').value

def convert_to_time(value:int):
    hours, remainder = divmod(int(value), 3600)
    minutes, seconds = divmod(remainder, 60)
    return f'{hours}:{minutes:02d}:{seconds:02d}'

def adjust_time(sheet_name_4: str):
    for row in range(2, last_row(sheet_name_4)+1):
        give_cell_values(sheet_name_4, f'F{row}', convert_to_time(get_cell_values(sheet_name_4, f'F{row}')))
        give_cell_values(sheet_name_4, f'G{row}', convert_to_time(get_cell_values(sheet_name_4, f'G{row}')))
        give_cell_values(sheet_name_4, f'H{row}', convert_to_time(get_cell_values(sheet_name_4, f'H{row}')))

def create_average_2(sheet_name_4:str, target_cell:str, get_cell:str, fill_color:tuple):
    cell_formula(sheet_name_4, target_cell,
             f'=IF(MOD(AVERAGE({get_cell}2:{get_cell}{last_row(sheet_name_4)}), 1) < 0.6, \
             TEXT(FLOOR.MATH(AVERAGE({get_cell}2:{get_cell}{last_row(sheet_name_4)}))/86400, "h:mm:ss") & " - " & \
             FLOOR.MATH(AVERAGE({get_cell}2:{get_cell}{last_row(sheet_name_4)})) & " seconds", \
             \
             TEXT(AVERAGE({get_cell}2:{get_cell}{last_row(sheet_name_4)})/86400, "h:mm:ss") & " - " & \
             TEXT(AVERAGE({get_cell}2:{get_cell}{last_row(sheet_name_4)}), "#0") & " seconds")'
             )
    range_color(sheet_name_4, target_cell, fill_color)
    text_color(sheet_name_4, target_cell)

def create_average_3(sheet_name_2:str, sheet_name_4:str):
    sheet_range(sheet_name_2, 'B:B').api.AutoFilter(2, 'Voice')
    give_cell_values(sheet_name_4, f'G{last_row(sheet_name_4)+1}', last_row_value(sheet_name_2, 'D'))
    sheet_range(sheet_name_2, 'B:B').api.AutoFilter(2, '<>')

    range_color(sheet_name_4, target_cell, fill_color)
    text_color(sheet_name_4, target_cell)

def create_count_2(sheet_name_5:str, get_cell:str):
    formula = (f'="Count: " & COUNTA({get_cell}2:{get_cell}{last_row(sheet_name_5)})')
    set_cell = f'{get_cell}{last_row(sheet_name_5)+1}'
    give_cell_values(sheet_name_5, set_cell, formula)
    range_color(sheet_name_5, set_cell, fill_color)
    text_color(sheet_name_5, set_cell)

def hide_formula_bar():
    xlwings.App(visible = Visible).api.DisplayFormulaBar = Visible

def bold_character(sheet_name_3:str, cell:str, value:str):
    sheet_range(sheet_name_3, cell).api.GetCharacters(Start=len(value)+2).Font.Bold = True

def acd(sheet_name_2:str, channel_name:str):
    return sheet_range(sheet_name_2, f'B2:B{last_row(sheet_name_2)}').value.count(channel_name)

def abd(sheet_name_5:str):
    return last_row(sheet_name_5) - 1

def last_row_value(sheet_name_2:str, cell:str):
    return sheet_range(sheet_name_2, f'{cell}2').expand().end("down").value

def partial_bold(sheet_name_3:str, cell:str, cell_value:str, value:str):
    give_cell_values(sheet_name_3, cell, f'{cell_value}: {value}')
    bold_character(sheet_name_3, cell, cell_value)

def result(sheet_name_3:str, target_cell:str):
    range_color(sheet_name_3, target_cell, fill_color_2)
    return ''

def result_2(sheet_name_3:str, target_cell:str, value):
    range_color(sheet_name_3, target_cell, (255, 255, 255))
    return value

def data_template(sheet_name_2:str, sheet_name_3:str, channel_name:str, start_cell:int):
    sheet_range(sheet_name_2, 'B:B').api.AutoFilter(2, channel_name, 7)
    last_row_value_B = last_row_value(sheet_name_2, 'B')
    rt_cell = f'A{start_cell}'
    acd_cell = f'A{start_cell+1}'
    abd_cell = f'A{start_cell+2}'
    scr_cell = f'A{start_cell+3}'
    aht_cell = f'A{start_cell+4}'
    data = [
        (rt_cell, 'COF', value:= \
         result(sheet_name_3, rt_cell) if last_row_value_B == 0 #
         else result_2(sheet_name_3, 'A11', int(last_row_value_B))
         ),
        (acd_cell, 'ACD',
         result(sheet_name_3, acd_cell) if last_row_value_B == 0 #
         else result_2(sheet_name_3, 'A11', int(last_row_value_B))
         ),
        (abd_cell, 'ABD', result(sheet_name_3, abd_cell)),
        (scr_cell, 'SCR', result(sheet_name_3, scr_cell) if value == '' else '100%'),
        (aht_cell, 'AHT',
         result(sheet_name_3, aht_cell) if last_row_value(sheet_name_2, 'D') == '' #
         else result_2(sheet_name_3, 'A11', str(last_row_value(sheet_name_2, 'D')).split('-')[0]))
        ]
    if 'Voice' not in channel_name:
        rt_cell = f'A{start_cell+5}'
        data.append((rt_cell, 'RT',
                     result(sheet_name_3, rt_cell) if last_row_value(sheet_name_2, 'C') == '' #
                     else result_2(sheet_name_3, 'A11', str(last_row_value(sheet_name_2, 'C')).split('-')[0]))
                     )
    for cell, cell_value, value in data:
        partial_bold(sheet_name_3, cell, cell_value, value)

    sheet_range(sheet_name_2, 'B:B').api.AutoFilter(2, '<>')

def set_cof_outbound_calls(sheet_name_3:str, sheet_name_7:str):
    last_row_value_B = last_row_value(sheet_name_7, 'B')
    partial_bold(sheet_name_3, 'A11', 'COF',
                 result(sheet_name_3, 'A11') if last_row_value_B == '' #
                 else result_2(sheet_name_3, 'A11', int(last_row_value_B)))

def other_social_media_template(sheet_name_2:str, sheet_name_3:str):
    filters = (
        'TW Comment', 'TW Message', 'IG Comment', 'IG Message', 'FB Comment',
        'FB Message', 'Manual Twitter', 'Manual instagram', 'Manual Facebook'
        )
    data_template(sheet_name_2, sheet_name_3, filters, 30)

    data = [
        ('A37', 'TW Comment', acd(sheet_name_2, filters[0])),
        ('A38', 'TW Message', acd(sheet_name_2, filters[1]) + acd(sheet_name_2, filters[6])),
        ('A39', 'IG Comment', acd(sheet_name_2, filters[2])),
        ('A40', 'IG Message', acd(sheet_name_2, filters[3]) + acd(sheet_name_2, filters[7])),
        ('A41', 'FB Comment', acd(sheet_name_2, filters[4])),
        ('A42', 'FB Message', acd(sheet_name_2, filters[5]) + acd(sheet_name_2, filters[8]))
        ] 
    for cell, cell_value, value in data:
        partial_bold(sheet_name_3, cell, cell_value, value:= \
                     result(sheet_name_3, cell) if value == 0
                     else result_2(sheet_name_3, cell, value))

def attendance_adherence(sheet_name_3:str):
    global wfo
    wfo = tuple(set(get_cell_values(sheet_name_2, f'A2:A{last_row(sheet_name_2)}')))
    data = [ 
        ('A45', 'Plan', ''),
        ('A46', 'WFH', '0'),
        ('A47', 'WFO', value:= len(wfo)), 
        ('A49', 'Realization', ''),
        ('A50', 'WFH', '0'),
        ('A51', 'WFO', value), 
        ('A53', 'Adherence', '100%')
        ]
    for cell, cell_value, value in data:
        partial_bold(sheet_name_3, cell, cell_value, value)

def date_object(sheet_name_2:str):
    return datetime.strptime(str(get_cell_values(sheet_name_2, 'L2')), '%Y-%m-%d %H:%M:%S')

def template():
    global date
    date = date_object(sheet_name_2).strftime('%B %d,  %Y')

    global input_date
    input_date =  date_object(sheet_name_2)

    global new_file_name
    new_file_name = f'Automated PHT Daily Reports - {date}.xlsx'
    
    return [
        ('A1', f'DAILY REPORT: CC SECRET - {date}'),
        ('A3', 'INBOUND'),
        ('A10', 'OUTBOUND'),
        ('A13', 'WHATSAPP'),
        ('A21', 'EMAIL'),
        ('A29', 'OTHER SOCIAL MEDIA'),
        ('A44', 'ATTENDANCE & ADHERENCE')
        ]
def bottom_border(sheet_name_3:str, target_cell:str):
    sheet_range(sheet_name_3, target_cell).api.Borders(9).LineStyle = 1

def report_template(sheet_name_3:str, template:list):
    set_range = 'A1;A3;A10;A13;A21;A29;A44'
    for cell, cell_value in template:
        if cell in set_range:
            give_cell_values(sheet_name_3, cell, f'{cell_value}')

    data_template(sheet_name_2, sheet_name_3, ('Voice', 'Manual Voice'), 4)
    partial_bold(sheet_name_3, 'A11', 'COF', result(sheet_name_3, 'A11'))
    data_template(sheet_name_2, sheet_name_3, ('Whatsapp', 'Manual WA'), 14)
    data_template(sheet_name_2, sheet_name_3, ('Email', 'Manual Email'), 22)
    other_social_media_template(sheet_name_2, sheet_name_3)
    attendance_adherence(sheet_name_3)

    range_color(sheet_name_3, set_range, fill_color)
    text_color(sheet_name_3, set_range)
    align_center(sheet_name_3, set_range)
    
    set_range_2 = 'A4;A5;A6;A7;A11;A14;A15;A16;A17;A18;A22;A23;A24;A25;A26;A30;A31;A32;' \
    'A33;A34;A35;A36;A37;A38;A39;A40;A41;A45;A46;A47;A48;A49;A50;A51;A52;A56;A57;A58;A59'
    bottom_border(sheet_name_3, set_range_2)

def top_5_interaction_1_column(sheet_name_3:str):
    for row in range(2, 7):
        cell_value = get_cell_values(sheet_name_3, f'A{row}')
        value = get_cell_values(sheet_name_3, f'B{row}')
        
        if value is not None:
            value = int(value)
            partial_bold(sheet_name_3, f'A{row}', cell_value, value)

def adjust_voice_amount(sheet_name_3:str, sheet_name_2:str):
    data_template(sheet_name_2, sheet_name_3, 'Voice', 4)

def adjust_abandoned_call(sheet_name_2:str, sheet_name_3:str, sheet_name_5:str):
    acd_value = acd(sheet_name_2, 'Voice')
    abd_value = abd(sheet_name_5)
    total_value = acd_value + abd_value

    data = [
        ('A4', 'COF', total_value),
        ('A6', 'ABD', abd_value),
        ('A7', 'SCR', f'{acd_value / total_value:.0%}')
    ]
    for cell, cell_value, value in data:
        partial_bold(sheet_name_3, cell, cell_value, value)

def delete_other_than_manual(sheet_name_6:str):
    values = 'TW', 'IG', 'FB', 'Voice', 'Whatsapp', 'Email'

    for value in values:
        delete_rows(sheet_name_6, value)

def title_table_4(sheet_name_6:str):
    cell_values = [
        ('A1', '  channel  '),
        ('B1', '  agent  '),
        ('C1', '   date_origin  '),
        ('D1', '   date_start   '),
        ('E1', '    date_end     '),
        ('F1', ' from_id '),
        ('G1', '   source   ')
        ]
    format_table(sheet_name_6, cell_values)

def adjust_manual(sheet_name_2:str, sheet_name_6:str):
    rows_with_manual = sorted_row(sheet_name_2, f'B2:B{last_row(sheet_name_2)}', 'Manual', False)

    for get_row, set_row in enumerate(rows_with_manual):
        cell_value = get_cell_values(sheet_name_6, f'B{get_row + 2}')
        value = 'Manual ' + cell_value if cell_value is not None else ''
        give_cell_values(sheet_name_2, f'B{set_row}', value)

    other_social_media_template(sheet_name_2, sheet_name_3)

def check_manual_exists(sheet_name_6:str, sheet_name_2:str):
    return any('Manual' in row.value for row in sheet_range(sheet_name_6, f'C2:C{last_row(sheet_name_6)}'))

def save_data_for_google_form(sheet_name_2:str):
    non_voice = (
        'Whatsapp', 'Email', 'TW Comment', 'TW Message', 'IG Comment', 'IG Message', 'FB Comment', 'FB Message',
        'Manual WA', 'Manual Email', 'Manual Twitter', 'Manual instagram', 'Manual Facebook'
        )
    global data_to_fill
    data_to_fill = []
    
    for name in wfo:
        sheet_range(sheet_name_2, 'A:A').api.AutoFilter(1, name)

        sheet_range(sheet_name_2, 'B:B').api.AutoFilter(2, 'Voice')
        productivity_voice = int(last_row_value(sheet_name_2, 'B'))

        sheet_range(sheet_name_2, 'B:B').api.AutoFilter(2, non_voice, 7)
        productivity_non_voice = int(last_row_value(sheet_name_2, 'B'))
        average_response_time_non_voice = str(last_row_value(sheet_name_2, 'C')).split('-')[0]        
        average_handling_time_non_voice = str(last_row_value(sheet_name_2, 'D')).split('-')[0]
        remarks = ' '
        
        data = {
            'name': name,
            'productivity_voice': productivity_voice,
            'productivity_non_voice': productivity_non_voice,
            'average_response_time_non_voice': average_response_time_non_voice,
            'average_handling_time_non_voice': average_handling_time_non_voice,
            'remarks': remarks
            }
        data_to_fill.append(data)
        
    sheet_range(sheet_name_2, 'B:B').api.AutoFilter(2, non_voice, 7)
    sheet_range(sheet_name_2, 'A:A').api.AutoFilter(1, wfo, 7)

    return data_to_fill

def save_data_for_whatsapp(sheet_name_3:str):
    cell_values = get_cell_values(sheet_name_3, 'A1:A60')

    del cell_values[34]
    del cell_values[26]
    del cell_values[18]

    whatsapp_data = [cell_value if cell_value is not None else '' for cell_value in cell_values]

    global whatsapp_message
    whatsapp_message = '\n'.join(whatsapp_data)

def save_data_for_google_form_2(sheet_name_3:str):
    global voice_cof
    voice_cof = get_cell_values(sheet_name_3, 'A4').split(':')[1]

    global voice_acd
    voice_acd = get_cell_values(sheet_name_3, 'A5').split(':')[1]

    global voice_abd
    voice_abd = get_cell_values(sheet_name_3, 'A6').split(':')[1]

    global voice_scr
    voice_scr = get_cell_values(sheet_name_3, 'A7').split(':')[1]

    global voice_aht
    voice_aht = get_cell_values(sheet_name_3, 'A8').split(': ')[-1]

    global outbound_cof
    outbound_cof = get_cell_values(sheet_name_3, 'A11').split(':')[1]

    global wa_cof
    wa_cof = get_cell_values(sheet_name_3, 'A14').split(':')[1]

    global wa_acd
    wa_acd = get_cell_values(sheet_name_3, 'A15').split(':')[1]

    global wa_abd
    wa_abd = get_cell_values(sheet_name_3, 'A16').split(':')[1]

    global wa_scr
    wa_scr = get_cell_values(sheet_name_3, 'A17').split(':')[1]

    global wa_aht
    wa_aht = get_cell_values(sheet_name_3, 'A18').split(': ')[-1]

    global wa_rt
    wa_rt = get_cell_values(sheet_name_3, 'A19').split(': ')[-1]

    global email_cof
    email_cof = get_cell_values(sheet_name_3, 'A22').split(':')[1]

    global email_acd
    email_acd = get_cell_values(sheet_name_3, 'A23').split(':')[1]

    global email_abd
    email_abd = get_cell_values(sheet_name_3, 'A24').split(':')[1]

    global email_scr
    email_scr = get_cell_values(sheet_name_3, 'A25').split(':')[1]

    global email_aht
    email_aht = get_cell_values(sheet_name_3, 'A26').split(': ')[-1]

    global email_rt
    email_rt = get_cell_values(sheet_name_3, 'A27').split(': ')[-1]

    global other_cof
    other_cof = get_cell_values(sheet_name_3, 'A30').split(':')[1]

    global other_acd
    other_acd = get_cell_values(sheet_name_3, 'A31').split(':')[1]

    global other_abd
    other_abd = get_cell_values(sheet_name_3, 'A32').split(':')[1]

    global other_scr
    other_scr = get_cell_values(sheet_name_3, 'A33').split(':')[1]

    global other_aht
    other_aht = get_cell_values(sheet_name_3, 'A34').split(': ')[-1]

    global other_rt
    other_rt = get_cell_values(sheet_name_3, 'A35').split(': ')[-1]

    global top_1
    top_1 = get_cell_values(sheet_name_3, 'A56')

    global top_2
    top_2 = get_cell_values(sheet_name_3, 'A57')

    global top_3
    top_3 = get_cell_values(sheet_name_3, 'A58')

    global top_4
    top_4 = get_cell_values(sheet_name_3, 'A59')

    global top_5
    top_5 = get_cell_values(sheet_name_3, 'A60')

#------------------------------IMPLEMENTATION-------------------------------

try:
    
    print('\nDATA TABULATION TAB:')

    print(f'Open {file_name}')
    with xlwings.App(visible = Visible).books.open(file_name) as workbook:

        print('Hide formula bar')
        hide_formula_bar()

        print('Delete redundant session_id - if any')
        delete_redundant_session_id(sheet_name)

        get_cell = 'A'
        target_cell = f'A{last_row(sheet_name)+1}'
        direction = 'right'

        print('Delete unecessary columns')
        set_ranges = 'D:M;V:X;AC:AD;AG:AG;AM:AM;AR:AY;BA:BF;BI:BJ'
        delete_columns(sheet_name, set_ranges)

        set_ranges = 'N:N'
        set_cells = 'A1'
        set_ranges_2 = 'A:A'

        set_title = 'average_handling_time'
        print(f'Create {set_title} column')
          
        column_1 = 'L'; column_2 = 'J'

        print(f'Set {set_title} = date_close - date_start_interaction')
        create_rt_aht(sheet_name, set_ranges, destination_column, direction,
                          set_cells, set_title, column_1, column_2)

        print('Formulate average of AHT in last cell')
        create_average(sheet_name, target_cell, get_cell)
    
        set_title = 'response_time'
        print(f'Create {set_title} column')
          
        column_1 = 'W'; column_2 = 'K'

        print(f'Set {set_title} = date_pickup_interaction - date_start_interaction')
        create_rt_aht(sheet_name, set_ranges, destination_column, direction,
                          set_cells, set_title, column_1, column_2)

        print('Formulate average of response time in last cell')
        create_average(sheet_name, target_cell, get_cell)

        print('Color cell of AHT and response time')
        target_range = f'A2:B{last_row(sheet_name)}'
        range_color(sheet_name, target_range, fill_color_2)

        set_ranges = 'R:R'

        print('Move channel_name')
        move_column(sheet_name, set_ranges, destination_column, direction)

        print('Formulate count of channel name')
        count_rows(sheet_name, target_cell)

        print('Move created_by_name')
        set_ranges = 'O:O'
        move_column(sheet_name, set_ranges, destination_column, direction)  

        target_range = f'B{last_row(sheet_name)+1}:D{last_row(sheet_name)+1}'

        print('Color title')
        range_color(sheet_name, target_range, fill_color)

        print('Color formulate cell')
        text_color(sheet_name, target_range)

        print('Change Manual response_time to 0')
        delete_cell_value(sheet_name)

        print('Align some columns paragraph to center')
        set_ranges = 'A:A;C:E;H:H;J:J;L:P;R:R;W:Z'
        align_center(sheet_name, set_ranges)

        print('Format response_time and average_handling_time')
        set_ranges = 'C:D'
        set_formats = 'h:mm:ss'
        format_column(sheet_name, set_ranges, set_formats)

        print('Format customer_hp')
        set_ranges = 'J:J'
        set_formats = '0'
        format_column(sheet_name, set_ranges, set_formats)

        print('Add tooltips')
        tooltip_text_1 = '\n        Dynamic\n        Counts'
        tooltip_text_2 = '\n        Dynamic\n        Average'
        add_tooltip(sheet_name, tooltip_text_1, tooltip_text_2)

        print('Change AB1 column name to ticket_Invoice/email')
        set_cells = 'AB1'
        value = 'ticket_Invoice/email'
        give_cell_values(sheet_name, set_cells, value)

        set_ranges = 'A:AC'

        print('Wrap text')
        wrap_text_sheet(sheet_name, set_ranges)

        print('Add auto filter')
        auto_filter(sheet_name, set_ranges)

        print('Hide gridlines')
        hide_gridlines(sheet_name)

        print(f'Rename {sheet_name} sheet to {sheet_name_2}')
        rename_sheet(sheet_name, sheet_name_2)

        print('Freeze top row')
        set_ranges = 'A1:AC1'
        freeze_top_row(sheet_name_2, set_ranges)

        print('Disable headings')
        disable_headings(sheet_name_2)

        print(f'Make {sheet_name_3} sheet')
        new_sheet(sheet_name_3)

        print('Save workbook')
        workbook.save()

    print('\nPIVOT TABLE TAB:')

    print('Make calculations of Top 5 Interaction')
    value = 'ticket_id'
    auto_pivot_table(file_name, sheet_name_2, value, sheet_name_3)

    with xlwings.App(visible = Visible).books.open(file_name) as workbook:

        direction = 'down'

        set_ranges = 'A1:B1'
        set_title = 'TOP 5 INTERACTIONS'

        print(f'Set title of {set_title}')
        title_table(sheet_name_3, set_ranges, set_title, fill_color)

        print('Delete lines 5 down')
        value = 'ain'
        delete_rows(sheet_name_3, value)

        print('Delete redundant: category > sub_category')
        values = ['Informasi - Informasi', 'Permintaan - Permintaan']
        delete_redundant_category(sheet_name_3, values)

        print('Sort descending')
        key='B'; order=2; orientation=1
        sort_column(sheet_name_3, key, order, orientation)

        print('Make it one column')
        
        partial_delete(sheet_name_3)
        
        top_5_interaction_1_column(sheet_name_3)
        
        set_ranges = 'B:B'
        delete_columns(sheet_name_3, set_ranges)

        print('Set row width')
        set_row_width(sheet_name_3, 'A1')
        
        print('Set view')
        set_ranges = 'A1:B6'; destination_column = 'A61'
        move_column(sheet_name_3, set_ranges, destination_column, direction)
        
        print('Calculate daily report')
        report_template(sheet_name_3, template())

        print('Disable headings')
        disable_headings(sheet_name_3)

        print('Hide gridlines')
        hide_gridlines(sheet_name_3)

        print('Save workbook')
        workbook.save()

    print('\nMANUAL TAB:')

    print('Check Manual input')
    with xlwings.App(visible = Visible).books.open(file_name_4) as workbook:

        data_voice = sheet(sheet_name).used_range.value

    with xlwings.App(visible = Visible).books.open(file_name) as workbook:

        new_sheet(sheet_name_6)

        sheet_range(sheet_name_6, 'A1').value = data_voice

        if check_manual_exists(sheet_name_6, sheet_name_2):

            print('Create Manual sheet')
            print(f'Export {file_name_4} to {sheet_name_6} sheet')

            print('Delete unecessary columns')
            set_ranges = 'A:B;D:I;K:L;O:Q;S:W;Y:Y;AA:AG'
            delete_columns(sheet_name_6, set_ranges)

            print('Delete non Manual rows')
            delete_other_than_manual(sheet_name_6)

            print('Adjust title table')
            title_table_4(sheet_name_6)

            print('Move source to column B')
            set_ranges = 'G:G'
            destination_column = 'B:B'
            direction = 'right'
            move_column(sheet_name_6, set_ranges, destination_column, direction)

            print('Adjust Manual calculations with pivot data')
            adjust_manual(sheet_name_2, sheet_name_6)

            print('Calculate manual counts')
            get_cell = 'B'
            create_count_2(sheet_name_6, get_cell)

            print('Set align center')
            set_ranges = 'A:F'
            align_center(sheet_name_6, set_ranges)

            print('Freeze top row')

            set_ranges = 'G1:H1'
            merge_cell(sheet_name_6, set_ranges)

            set_ranges = 'A1:H1'
            select_row(sheet_name_6, set_ranges)

            set_ranges = 'A1:G1'
            freeze_top_row(sheet_name_6, set_ranges)

            print('Set auto filter')
            set_ranges = 'A:H'
            auto_filter(sheet_name_6, set_ranges)

            print('Disable headings')
            disable_headings(sheet_name_6)

            print('Hide gridlines')
            hide_gridlines(sheet_name_6)
            
        else:

            print('No Manual input')
            workbook.sheets[sheet_name_6].delete()

        print('Save workbook')
        workbook.save()

        print('Save data for Whatsapp')
        save_data_for_whatsapp(sheet_name_3)

        print('Save data for Google Form: Agent Performance')
        save_data_for_google_form(sheet_name_2)

        print('Save data for Google Form: Daily Reports')
        save_data_for_google_form_2(sheet_name_3)

        os.remove(file_name_4)

except FileNotFoundError:
    
    print('CWC is still empty! Unable to tabulate data')

try:
    
    print('\nANSWERED CALLS TAB:')

    print('Check answered calls')
    with xlwings.App(visible = Visible).books.open(file_name_2) as workbook:

        data_voice = sheet(sheet_name).used_range.value
    
    with xlwings.App(visible = Visible).books.open(file_name) as workbook:

        print(f'Create {sheet_name_4} sheet')
        new_sheet(sheet_name_4)

        print(f'Retrieve the data from {file_name_2}')
        sheet_range(sheet_name_4, 'A1').value = data_voice

        print('Delete unecessary columns')
        set_ranges = 'D:F;M:M'
        delete_columns(sheet_name_4, set_ranges)

        print('Adjust each title')
        title_table_2(sheet_name_4)

        print('Align center')
        set_ranges = 'A:D;F:I'
        align_center(sheet_name_4, set_ranges)

        print('Format cl_id column')
        set_ranges = 'E:E'
        set_formats = '0'
        format_column(sheet_name_4, set_ranges, set_formats)

        #adjust_time(sheet_name_4)

        print('Add tooltip to average cell')
        target_cell = f'G{last_row(sheet_name_4)+1}'
        tooltip_text = '\n\n        Average'
        add_tooltip_2(sheet_name_4, target_cell, tooltip_text)

        print('Delete inconsistence voice')
        delete_inconsistence_voice(sheet_name_2, sheet_name_4)
        voice_data(sheet_name_2, sheet_name_4)

        print('Create average of talk_time')
        get_cell = 'G'
        target_cell = f'G{last_row(sheet_name_4)+1}'
        #create_average_2(sheet_name_4, target_cell, get_cell, fill_color)
        create_average_3(sheet_name_2, sheet_name_4)

        print('Adjust voice amount to pivot data') 
        adjust_voice_amount(sheet_name_3, sheet_name_2)

        print('Sort ascending date_time')
        key='A'; order=1; orientation=1
        sort_column(sheet_name_4, key, order, orientation)

        print('Focus title')
        set_ranges = 'A1:I1'
        select_row(sheet_name_4, set_ranges)

        print('Disable headings')
        disable_headings(sheet_name_4)

        print('Hide gridlines')
        hide_gridlines(sheet_name_4)

        print('Adjust tab order')
        sheet_names = ['Answered Calls', 'Data Tabulation', 'Pivot Table']
        reverse_sheet(sheet_names)

        print('Save workbook')
        workbook.save()

        print('Save data for Whatsapp')
        save_data_for_whatsapp(sheet_name_3)

        print('Save data for Google Form: Agent Performance')
        save_data_for_google_form(sheet_name_2)

        print('Save data for Google Form: Daily Reports')
        save_data_for_google_form_2(sheet_name_3)

        os.remove(file_name_2)

except FileNotFoundError:

    if os.path.exists(file_name):
        print('No answered calls!')
    else:
        print(f'Manual input please - if any')

try:
    
    print('\nABANDONED CALLS TAB:')

    print('Check abandoned calls')
    with xlwings.App(visible = Visible).books.open(file_name_3) as workbook:

        data_voice = sheet(sheet_name).used_range.value

    with xlwings.App(visible = Visible).books.open(file_name) as workbook:

        print(f'Create {sheet_name_5} sheet')
        new_sheet(sheet_name_5)

        print(f'Retrieve the data from {file_name_3}')
        sheet_range(sheet_name_5, 'A1').value = data_voice

        print('Delete unecessary columns')
        set_ranges = 'D:E'
        delete_columns(sheet_name_5, set_ranges)

        print('Create title table')
        title_table_3(sheet_name_5)

        print('Align center')
        set_ranges = 'A:G'
        align_center(sheet_name_5, set_ranges)

        print('Sort ascending date_time')
        key='A'; order=1; orientation=1
        sort_column(sheet_name_5, key, order, orientation)

        print('Calculate abandoned counts')
        get_cell = 'D'
        create_count_2(sheet_name_5, get_cell)

        print('Adjust abandoned to pivot data')
        adjust_abandoned_call(sheet_name_2, sheet_name_3, sheet_name_5)

        print('Select title')
        set_ranges = 'A1:G1'
        select_row(sheet_name_5, set_ranges)

        print('Disable headings')
        disable_headings(sheet_name_5)

        print('Hide gridlines')
        hide_gridlines(sheet_name_5)

        print('Adjust tab order')
        if os.path.exists(file_name_2):
            sheet_names = ['Abandoned Calls', 'Answered Calls', 'Data Tabulation', 'Pivot Table']
        else:
            sheet_names = ['Abandoned Calls', 'Data Tabulation', 'Pivot Table']
        reverse_sheet(sheet_names)
        
        print('Save workbook')
        workbook.save()

        print('Adjust data for Whatsapp')
        save_data_for_whatsapp(sheet_name_3)

        print('Adjust data for Google Form: Daily Reports')
        save_data_for_google_form_2(sheet_name_3)

        os.remove(file_name_3)

except FileNotFoundError:

    if os.path.exists(file_name):
        print('No abandoned calls!')
    else:
        print(f'Manual input please - if any')

try:
    
    print('\nOUTBOUND CALLS TAB:')

    print('Check outbound calls')
    with xlwings.App(visible = Visible).books.open(file_name_5) as workbook:

        data_voice = sheet(sheet_name).used_range.value

    with xlwings.App(visible = Visible).books.open(file_name) as workbook:

        print(f'Create {sheet_name_7} sheet')
        new_sheet(sheet_name_7)

        print(f'Retrieve the data from {file_name_5}')
        sheet_range(sheet_name_7, 'A1').value = data_voice

        print('Delete unecessary columns')
        set_ranges = 'A:E;G:G;I:K;O:O;Q:R;T:U;W:AA'
        delete_columns(sheet_name_7, set_ranges)

        print('Create title table')
        title_table_5(sheet_name_7)

        print('Format to_id')
        set_ranges = 'G:G'
        set_formats = '0'
        format_column(sheet_name_7, set_ranges, set_formats)

        print('Align center')
        set_ranges = 'C:H'
        align_center(sheet_name_7, set_ranges)

        print('Select title')
        set_ranges = 'A1:H1'
        select_row(sheet_name_7, set_ranges)

        print('Disable headings')
        disable_headings(sheet_name_7)

        print('Add auto filter')
        auto_filter(sheet_name_7, set_ranges)

        print('Formulate outbound counts')
        target_cell = f'B{last_row(sheet_name_7)+1}'
        count_rows(sheet_name_7, target_cell)

        print('Color formulate cell')
        text_color(sheet_name_7, target_cell)
        range_color(sheet_name_7, target_cell, fill_color)

        print('Add tooltips')
        tooltip_text = '\n        Dynamic\n        Counts'
        add_tooltip_2(sheet_name_7, target_cell, tooltip_text)

        print('Freeze top row')
        freeze_top_row(sheet_name_7, set_ranges)

        print('Hide gridlines')
        hide_gridlines(sheet_name_7)

        print('Adjust tab order')
        if os.path.exists(file_name_3):
            sheet_names = ['Outbound Calls', 'Abandoned Calls', 'Answered Calls', 'Data Tabulation', 'Pivot Table']
        else:
            sheet_names = ['Outbound Calls', 'Answered Calls', 'Data Tabulation', 'Pivot Table']
        reverse_sheet(sheet_names)

        print('Adjust outbound to pivot data')
        set_cof_outbound_calls(sheet_name_3, sheet_name_7)

        print('Save workbook')
        workbook.save()

        print('Adjust data for Whatsapp')
        save_data_for_whatsapp(sheet_name_3)

        print('Adjust data for Google Form: Daily Reports')
        save_data_for_google_form_2(sheet_name_3)

        os.remove(file_name_5)

except FileNotFoundError:
    print('No outbound calls!')

if os.path.exists(file_name):
    print(f'\nRename file to: {new_file_name}')
    os.rename(file_name, new_file_name)
else:
    print(f'\nFile {file_name} not found!')


print('\n\nPART 3: AUTO INPUT GOOGLE FORM')
print(input_date)

from undetected_chromedriver import ChromeOptions
from undetected_chromedriver import Chrome
from selenium.webdriver.common.keys import Keys
from time import sleep
from random import randint

#------------------------------FUNCTION LISTS-------------------------------

def navigate_to_google_form(url:str):
    return driver.get(url)

def located_element(selector:str, element:str):
    sleep(1) # Add delay so element uploaded successfully
    return driver.find_element(selector, element)

def auto_click_button():
    located_element('class name', 'NPEfkd').click()

def seconds_to_time_str(seconds:int):
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f'{hours}:{minutes:02d}:{seconds:02d}'

def time_str_to_seconds(time_str:str):
    hours, minutes, seconds = map(int, time_str.split(':'))
    total_seconds = hours * 3600 + minutes * 60 + seconds
    return total_seconds
  
def autofill_performance_agent_SECRET():
    print('Autofill Google form: Performance Agent SECRET')
    
    names = {
        'SECRET1': 'SECRET', 'SECRET2': 'SECRET', 'SECRET3': 'SECRET',
        'SECRET4': 'SECRET', 'SECRET5': 'SECRET'
        }

    for data in data_to_fill:
        print("Navigate to Google Form: Performance Agent SECRET")
        
        url = 'https://SECRET'
        navigate_to_google_form(url)
        #driver.get(value)

        name = names[data['name']]
        print(f'Name: {name}')

        productivity_voice = data['productivity_voice']
        print(f'Productivity voice: {productivity_voice}')

        productivity_non_voice = data['productivity_non_voice']
        print(f'Productivity non voice: {productivity_non_voice}')

        average_response_time_non_voice = data['average_response_time_non_voice']
        print(f'Average response time non voice: {average_response_time_non_voice}')

        average_handling_time_non_voice = data['average_handling_time_non_voice']
        print(f'Average handling time non voice: {average_handling_time_non_voice}')

        remarks = data['remarks']

        located_element('class name', 'whsOnd').send_keys(
            input_date.strftime('%d%m') + Keys.RIGHT + input_date.strftime('%Y')
            + Keys.TAB + name
            + Keys.TAB + str(productivity_voice)
            + Keys.TAB + str(productivity_non_voice)
            + Keys.TAB + average_response_time_non_voice
            + Keys.TAB + str(average_handling_time_non_voice)
            + Keys.TAB + remarks
            )
            
        auto_click_button()

        print(f"Add time to ensure {data['name']}'s data uploaded successfully")
        sleep(1)
  
def autofill_daily_report_SECRET():
    print('Autofill Google form: Daily Report SECRET')
    url = 'https://SECRET'

    print("Navigate to Google Form: Daily Report SECRET")
    navigate_to_google_form(url)

    located_element('class name', 'whsOnd').send_keys(
        input_date.strftime('%d%m') + Keys.RIGHT + input_date.strftime('%Y')
        + Keys.TAB + (get_voice_cof := voice_cof)
        + Keys.TAB + (get_voice_acd := voice_acd)
        + Keys.TAB + (get_voice_abd := voice_abd)
        + Keys.TAB + (get_voice_scr := voice_scr)
        + Keys.TAB + (get_voice_aht := voice_aht)
        + Keys.TAB + (get_wa_cof := wa_cof)
        + Keys.TAB + (get_wa_acd := wa_acd)
        + Keys.TAB + (get_wa_abd := wa_abd)
        + Keys.TAB + (get_wa_scr := wa_scr)
        + Keys.TAB + (get_wa_aht := wa_aht)
        + Keys.TAB + (get_wa_rt := wa_rt)
        + Keys.TAB + (get_email_cof:= email_cof)
        + Keys.TAB + (get_email_acd := email_acd)
        + Keys.TAB + (get_email_abd := email_abd)
        + Keys.TAB + (get_email_scr := email_scr)
        + Keys.TAB + (get_email_aht := email_aht)
        + Keys.TAB + (get_email_rt := email_rt)
        + Keys.TAB + (get_other_cof := other_cof)
        + Keys.TAB + (get_other_acd := other_acd)
        + Keys.TAB + (get_other_abd := other_abd)
        + Keys.TAB + (get_other_scr := other_scr)
        + Keys.TAB + (get_other_aht := other_aht)
        + Keys.TAB + (get_other_rt := other_rt)
        + Keys.TAB + (get_outbound_cof := outbound_cof)
        + Keys.TAB + (get_top_1 := '' if top_1 is None else top_1)
        + Keys.TAB + (get_top_2 := '' if top_2 is None else top_2)
        + Keys.TAB + (get_top_3 := '' if top_3 is None else top_3)
        + Keys.TAB + (get_top_4 := '' if top_4 is None else top_4)
        + Keys.TAB + (get_top_5 := '' if top_5 is None else top_5)
        )
        
    auto_click_button()

    print('Add time to ensure data uploaded successfully')
    sleep(1)

#------------------------------IMPLEMENTATION-------------------------------
        
options = ChromeOptions()

chrome_path = \
'D:/SECRET/GoogleChromePortable64/GoogleChromePortable.exe'
options.binary_location = chrome_path

argument_1 = '--no-sandbox'
options.add_argument(argument_1)

argument_2 = '--app=https://accounts.google.com/'
#options.add_argument(argument_2)

argument_3 = '--headless'
#options.add_argument(argument_3)

argument_4 = '--start-maximized'
options.add_argument(argument_4)

prefs = {
    'credentials_enable_service': False,
    'profile.password_manager_enable': False
}
options.add_experimental_option('prefs', prefs)

print('Create undetected Chrome instance')
driver = Chrome(options)

try:

    print('Add delay so not detected as bot')
    timeout = 60
    driver.implicitly_wait(timeout)

    autofill_performance_agent_SECRET()
    autofill_daily_report_SECRET()

finally:

    driver.close()
    driver.quit()


print('\n\nPART 4: AUTO SEND WHATSAPP DAILY REPORTS')

import pywhatkit
from time import sleep

def hour():
    return int(datetime.now().strftime('%H'))

def minute():
    return int((datetime.now() + timedelta(minutes=1.5)).strftime('%M'))

pywhatkit.sendwhatmsg_to_group('BjX6w61sUFiL5kHgjV1sjp', whatsapp_message, hour(), minute(), tab_close=True)
sleep(3)
pywhatkit.sendwhatmsg('+62SECRET', whatsapp_message, hour(), minute(), tab_close=True)

#pywhatkit.sendwhatmsg_to_group('EAl65pvSECRET', whatsapp_message, hour(), minute(), tab_close=True)
#pywhatkit.sendwhatmsg('+62SECRET', whatsapp_message, hour(), minute(), tab_close=True)



