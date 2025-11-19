import sys
from datetime import datetime
import json
import os
import os.path
from pathlib import Path

import keyboard
import xlsxwriter
from striprtf.striprtf import rtf_to_text
from colorama import Fore


# Name in dic property
class CC:
    CATALOG = '#CATALOG'
    DEBUG_NAME = '#DEBUGNAME'
    FORM_NAME = '#NAMEFORM'
    DEVELOPER = '#DEVELOPER'
    CLIENT = '#CLIENT'
    START_PATH = '#STARTPATH'
    SYSTEM_SOURCE_TYPE = '#SYSTEMTYPEISH'
    PROTOTYPE = '#PROTOTYPE'
    DO = '#DO'
    CODE = '#CODE'
    HOUR = '#HOUR'
    VIPFCOM_BASENAME = '#VIPFCOM_BASENAME'
    VIPFCOM_FILE_PATH = '#FULLPATHVIPFCOM'
    DEFRES_BASENAME = '#DEFRESBASENAME'
    DEFRES_FILE_PATH = '#FULLPATHDEFRES'
    NOTE = '#PRIM'
    RELEASE_DATE = '#DATE'
    INTERFACE_NAME = 'INTERFACENAME'
    INTERFACE = 'INTERFACE'
    NAME_IN_LIST = 'NAMEINLIST'

class AppSettings:
    folder_process:str
    folder_result:str
    show_skip:bool
    use_diff :bool
    ext_vip :bool
    ext_slk :bool
    ext_rtf :bool
    ext_gcd :bool
    ext_fr3 :bool

    def __init__(self):
        self.folder_process = ''
        self.folder_result = ''
        self.show_skip = False
        self.use_diff = True
        self.ext_vip = True
        self.ext_slk = True
        self.ext_rtf = True
        self.ext_gcd = True
        self.ext_fr3 = False


# Печать в Excel
# https://xlsxwriter.readthedocs.io/worksheet.html#write
# Пакеты:
# XlsxWriter	    1.2.8	1.2.8
# setuptools	    41.2.0	46.1.3


#  ------------------------------------------------------------------------
# SLK Cyrillic dict
a_dict = {"\\x1bNAA": "А", "\\x1bNBA": "Б", "\\x1bNCA": "В", "\\x1bNDA": "Г", "\\x1bNHA": "Д", "\\x1bNJA": "Е",
          "\\x1bNH ": "Ё", "\\x1bNa": "Ж", "\\x1bNKC": "З", "\\x1bNAE": "И", "\\x1bNBE": "Й", "\\x1bNCE": "К",
          "\\x1bNHE": "Л", "\\x1bNAI": "М", "\\x1bNBI": "Н", "\\x1bNCI": "О", "\\x1bNHI": "П", "\\x1bNb": "Р",
          "\\x1bNDN": "С", "\\x1bNAO": "Т", "\\x1bNBO": "У", "\\x1bNCO": "Ф", "\\x1bNDO": "Х", "\\x1bNHO": "Ц",
          "\\x1b-7": "Ч", "\\x1bNi": "Ш", "\\x1bNAU": "Щ", "\\x1bNBU": "Ъ", "\\x1bNCU": "Ы", "\\x1bNHU": "Ь",
          "\\x1b-=": "Э", "\\x1bNl": "Ю", "\\x1bN{": "Я", "\\x1bNAa": "а", "\\x1bNBa": "б", "\\x1bNCa": "в",
          "\\x1bNDa": "г", "\\x1bNHa": "д", "\\x1bNJa": "е", "\\x1b+8": "ё", "\\x1bNq": "ж", "\\x1bNKc": "з",
          "\\x1bNAe": "и", "\\x1bNBe": "й", "\\x1bNCe": "к", "\\x1bNHe": "л", "\\x1bNAi": "м", "\\x1bNBi": "н",
          "\\x1bNCi": "о", "\\x1bNHi": "п", "\\x1bNs": "р", "\\x1bNDn": "с", "\\x1bNAo": "т", "\\x1bNBo": "у",
          "\\x1bNCo": "ф", "\\x1bNDo": "х", "\\x1bNHo": "ц", "\\x1b/7": "ч", "\\x1bNy": "ш", "\\x1bNAu": "щ",
          "\\x1bNBu": "ъ", "\\x1bNCu": "ы", "\\x1bNHu": "ь", "\\x1b/=": "э", "\\x1bN|": "ю", "\\x1bNHy": "я"
          }

# Fixed property DEFRES
property_list = [CC.FORM_NAME, CC.DEVELOPER, CC.CLIENT, CC.START_PATH
    , CC.DO, CC.CODE, CC.HOUR, CC.NOTE, CC.RELEASE_DATE]


def print_error(*values):
    for value in values:
        print(f'{Fore.RED}{value}')


def print_warning(*values):
    for value in values:
        print(f'{Fore.YELLOW}{value}')


def print_header(*values):
    for value in values:
        print(f'{Fore.BLUE}{value}')


def print_step1(*values):
    for value in values:
        print(f'{Fore.BLUE}{value}')


def print_step2(*values):
    for value in values:
        print(f'{Fore.LIGHTBLUE_EX}{value}')


def print_result(*values):
    for value in values:
        print(f'{Fore.GREEN}{value}')


def print_value(*values):
    for value in values:
        print(f'{Fore.WHITE}{value}')

# Clear name
def clear_form_name(value):
    return (value.replace('\'', '').replace('\"', '').replace('"', '').replace('`', '').
            replace('’', '').replace(';','').replace('‘', '').replace('\\', '').replace('\'','').strip())


# Read or create Settings file
def read_settings(settings_file_name):
    app_settings = AppSettings()
    members = [attr for attr in dir(app_settings) if
               not callable(getattr(app_settings, attr)) and not attr.startswith("__")]

    # Create settings file
    if not os.path.exists(settings_file_name):
        print_header(f'Create new {settings_file_name}...')

        default = {}

        # Default json value
        for member in members:
            default[member] = getattr(app_settings, member)

        # Write default value to json file
        with open(settings_file_name, "w", encoding="utf-8") as f:
            json.dump(default, f, ensure_ascii=False, indent=2)

        # Return default value
        return False, app_settings

    with open(settings_file_name, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    for member in members:
        setattr(app_settings, member, cfg.get(member, getattr(app_settings, member)))

    return True, app_settings


# Create excel column
def create_excel_column():
    row = {"Num": 0, "Name": "Группа", "Width": 30, "CellName": CC.CATALOG}
    ListExcelColumn.append(row)
    row = {"Num": 1, "Name": "Ресурс (*.res)", "Width": 20, "CellName": CC.DEBUG_NAME}
    ListExcelColumn.append(row)
    row = {"Num": 2, "Name": "Наименование", "Width": 60, "CellName": CC.FORM_NAME}
    ListExcelColumn.append(row)
    row = {"Num": 3, "Name": "Разработчик", "Width": 25, "CellName": CC.DEVELOPER}
    ListExcelColumn.append(row)
    row = {"Num": 4, "Name": "Заказчик", "Width": 25, "CellName": CC.CLIENT}
    ListExcelColumn.append(row)
    row = {"Num": 5, "Name": "Путь запуска", "Width": 50, "CellName": CC.START_PATH}
    ListExcelColumn.append(row)
    row = {"Num": 6, "Name": "System Type Res", "Width": 20, "CellName": CC.SYSTEM_SOURCE_TYPE}
    ListExcelColumn.append(row)
    row = {"Num": 7, "Name": "Прототип LinkForm\\Имя", "Width": 25, "CellName": CC.PROTOTYPE}
    ListExcelColumn.append(row)
    row = {"Num": 8, "Name": "ДО (договор\\заявка)", "Width": 25, "CellName": CC.DO}
    ListExcelColumn.append(row)
    row = {"Num": 9, "Name": "Код", "Width": 15, "CellName": CC.CODE}
    ListExcelColumn.append(row)
    row = {"Num": 10, "Name": "Трудоемкость (н/час)", "Width": 10, "CellName": CC.HOUR}
    ListExcelColumn.append(row)
    row = {"Num": 11, "Name": "File VIP\\FCOM", "Width": 25, "CellName": CC.VIPFCOM_BASENAME}
    ListExcelColumn.append(row)
    row = {"Num": 12, "Name": "File DeffRes", "Width": 25, "CellName": CC.DEFRES_BASENAME}
    ListExcelColumn.append(row)
    row = {"Num": 13, "Name": "Path VIP\\FCOM", "Width": 90, "CellName": CC.VIPFCOM_FILE_PATH}
    ListExcelColumn.append(row)
    row = {"Num": 14, "Name": "Path DeffRes", "Width": 90, "CellName": CC.DEFRES_FILE_PATH}
    ListExcelColumn.append(row)
    row = {"Num": 15, "Name": "Примечание", "Width": 90, "CellName": CC.NOTE}
    ListExcelColumn.append(row)
    row = {"Num": 16, "Name": "Период", "Width": 20, "CellName": CC.RELEASE_DATE}
    ListExcelColumn.append(row)


#  Get release *.res name from name.bat
def get_res_public_name(p_project_path):
    res_name = 'debug.res'
    name_bat_path = os.path.join(p_project_path, 'name.bat')
    if os.path.exists(name_bat_path):
        with open(name_bat_path, 'r', encoding='cp866') as f:
            file_content = f.read()
            for line in file_content.splitlines():
                if line.upper().find('DEBUGNAME') != -1 and line.find('=') != -1 and line.upper().find('REM') == -1:
                    res_name = line.split("=")[1].strip()
    print_result(f'Name *.res: {res_name}')
    return res_name


# Parse line from file
def parse_line(p_line, p_item, p_file_type):
    # ------------------------------DEFRES in source code-------------------------------
    # Trim + upper (universal find)
    p_line = p_line.strip()
    line_upper = p_line.upper()

    if line_upper.find('#DEFRES') != -1:
        for cur_property in property_list: # Only function tag`s
            if cur_property in line_upper: # Tag in line
                p_item[cur_property] = p_line.replace('#DEFRES', '').strip()[len(cur_property):].strip()
                print(f'{cur_property} = {p_item[cur_property]}')
    # ----------------------------------------------------------------------------------
    # Interface Name: 'Interface GetSomeValue202;' -> 'GetSomeValue202'
    find_string = 'INTERFACE '
    if line_upper.find(find_string) == 0:
        p_item[CC.INTERFACE] = p_line[len(find_string):].strip().replace(';', '')

    # ТХОAPI Name: 'VipInterface txo_GetSomeValue2025 implements ObjTxoIdentPlugin;' -> 'GetSomeValue2025'
    if line_upper.find('VIPINTERFACE') != -1 and line_upper.find('IMPLEMENTS') != -1:
        p_item[CC.INTERFACE_NAME] = p_line.split(' ')[1].lower().replace('txo_', '').strip()
    # -------------------------Link form prototype-------------------------------
    # LinkForm
    if line_upper.find('.LINKFORM ') != -1 and line_upper.find(' PROTOTYPE ') != -1:
        x = line_upper.find(' IS ')
        p_item[CC.PROTOTYPE] = clear_form_name(p_line[x + 4:])
    # -------------------------Link form + ARD-------------------------------
    # LinkForm + ARD: '.NameInList 'Report name'' -> 'Report name'
    find_string = '.NAMEINLIST '
    if line_upper.find(find_string) != -1:
        p_item[CC.FORM_NAME] = clear_form_name(p_line[line_upper.find(find_string) + len(find_string):].strip())
        if p_item.get(CC.START_PATH) is None:
            p_item[CC.START_PATH] = f'LinkForm ({p_file_type})'
        if p_item.get(CC.SYSTEM_SOURCE_TYPE) is None:
            p_item[CC.SYSTEM_SOURCE_TYPE] = f'LinkForm ({p_file_type})'
    # -------------------------ARD-------------------------------
    # ARD: '.ard'
    if line_upper.find('.ARD') != -1:
        p_item[CC.START_PATH] = f'ARD ({p_file_type})'
        p_item[CC.SYSTEM_SOURCE_TYPE] = f'ARD ({p_file_type})'
    # -------------------------Auto form-----------------------------------------
    # Auto form: 'GetReportName := 'Report name'' -> 'Report name'
    if line_upper.find('GETREPORTNAME') != -1 and line_upper.find(':=') != -1:
        x = line_upper.find(':=')
        p_item[CC.FORM_NAME] = clear_form_name(p_line[x + 2:])
        p_item[CC.START_PATH] = 'UserReport (VIP)'
        p_item[CC.SYSTEM_SOURCE_TYPE] = p_item.get(CC.START_PATH)
    # -------------------------------ТХОAPI--------------------------------------
    # ТХОAPI+: '{GetInfo := 'TXOAPI. Name TXO';}' -> 'TXOAPI. Name TXO'
    find_string = 'GETINFO'
    if line_upper.find(find_string) != -1 and line_upper.find(':=') != -1:
        # Is 'GETINFO:' or '{GETINFO' or ' GETINFO :'
        process_step = False
        if line_upper.find(find_string) == 0 and (
                line_upper[len(find_string)] == ':' or line_upper[len(find_string)] == ' '):
            process_step = True
        elif not line_upper[line_upper.find(find_string) - 1].isalpha() and (
                line_upper[line_upper.find(find_string) + len(find_string)] == ':' or line_upper[
            line_upper.find(find_string) + len(find_string)] == ' '):
            process_step = True
        if process_step:
            x = line_upper.find(':=')
            p_item[CC.FORM_NAME] = clear_form_name(p_line[x + 2:])
            p_item[CC.START_PATH] = f'TXOAPI [{p_item.get(CC.INTERFACE_NAME)}]'
            p_item[CC.PROTOTYPE] = p_item.get(CC.INTERFACE_NAME)
            p_item[CC.SYSTEM_SOURCE_TYPE] = 'TXOApi (VIP)'
    # -----------------------------ALTER INTERFACE------------------------------
    # Alter: 'alter interface VSCHET;' -> 'VSCHET'
    if line_upper.find('ALTER') != -1 and line_upper.find('INTERFACE') != -1:
        x = line_upper.find('INTERFACE')
        p_item[CC.FORM_NAME] = p_line[x + 9:].replace('\';', '').replace('\'', '').replace(';', '').strip()
        p_item[CC.START_PATH] = 'AlterInterface (VIP)'
        p_item[CC.SYSTEM_SOURCE_TYPE] = p_item.get(CC.START_PATH)
    # -----------------------------TUNE-------------------------------
    # Tune: epDateModifyTune (TS: ObjStartCreateTune)
    if line_upper.find('OBJSTARTCREATETUNE') != -1:
        p_item[CC.START_PATH] = 'Настройки'
        p_item[CC.SYSTEM_SOURCE_TYPE] = 'Tune'
        p_item[CC.FORM_NAME] = "Настройки"
    if p_item.get(CC.START_PATH) == 'Настройки' and line_upper.find('TR.ADDTUNE') != -1:
        # No comment
        if line_upper.strip()[:2] != '//':
            x = line_upper.find(';')
            if x > 0:
                # Tune name: TR.AddTune (ttUSERTUNE, 'MR_TUNE.USER1', 'MR_TUNE', 'Show journal', ftList, '0', '', 1);  ->  'Show journal'
                note_split = line_upper[:x].strip()[10:-1].strip()[1:].strip().split(',')
                if len(note_split) >= 3:
                    print_value(note_split[3].strip()[1:-1])
                    if p_item.get(CC.NOTE) is None:
                        p_item[CC.NOTE] = note_split[3].strip()[1:-1] + ';'
                    else:
                        p_item[CC.NOTE] = p_item[CC.NOTE] + '\n' + note_split[3].strip()[1:-1].strip() + ';'
    # ---------------------------------------------------------------------------
    return p_item


# Create final list whit data
def create_list_from_source(p_process_dir, p_items_list, app_settings):
    file_ext_list = {'vip': 1, 'frm': 2, 'rtf': 3, 'slk': 4, 'gcd': 5, 'fr3': 6}  # Acceptable file extensions

    for _path, _sub_dirs, _files in os.walk(p_process_dir):
        for cur_file in _files:
            extension = cur_file.split('.')[-1].lower()
            if os.path.isfile(os.path.join(_path, cur_file)) and extension in file_ext_list \
                    and (((extension == 'vip' or extension == 'frm') and app_settings.ext_vip)
                         or (extension == 'slk' and app_settings.ext_slk)
                         or (extension == 'rtf' and app_settings.ext_rtf)
                         or (extension == 'gcd' and app_settings.ext_gcd)
                         or (extension == 'fr3' and app_settings.ext_fr3)):

                print_step1(f'Parse file: {cur_file}')

                _split_path = p_process_dir.split("\\")
                file_full_path = os.path.join(_path, cur_file)
                # All source file in one *.res file
                res_name = get_res_public_name(p_process_dir)

                # Создаем элемент для добавления в общий список
                new_item = {}
                new_item.clear()
                new_item[CC.CATALOG] = _split_path[len(_split_path) - 2]  # Folder prev v_last (project name)
                new_item[CC.VIPFCOM_BASENAME] = cur_file  # Source file
                new_item[CC.VIPFCOM_FILE_PATH] = file_full_path  # Source file full path
                new_item[CC.DEBUG_NAME] = res_name  # *.res name from name.bat

                # -----------------------------------VIP(FRM)-----------------------------------
                if (extension == 'vip' or extension == 'frm') and app_settings.ext_vip:
                    print_step2(f'Start VIP(FRM): {file_full_path}')
                    with open(file_full_path, 'r', encoding='cp866') as file:
                        for _line in file:
                            _str = str(_line)
                            new_item = parse_line(_str, new_item, 'FCOM')
                    print_value(f'End')

                # -----------------------------------RTF----------------------------------------
                if extension == 'rtf' and app_settings.ext_rtf:
                    print_step2(f'Start RTF: {file_full_path}')
                    try:
                        with open(file_full_path, "r") as f:
                            content_rtf = f.read()
                            text_rtf = rtf_to_text(content_rtf)
                            for i, line_rtf in enumerate(text_rtf.splitlines(), 1):
                                new_item = parse_line(line_rtf, new_item, 'RTF')
                    except Exception as e:
                        print_error(f'Error (rtf_to_text): {e}')
                        with open(file_full_path, 'r') as file:
                            _FileDataStr = file.read()
                            for _line in _FileDataStr.splitlines():
                                _str = str(_line).strip()
                                new_item = parse_line(_str, new_item, 'RTF')
                        if new_item.get(CC.FORM_NAME) is None:
                            print_error(f'Error, файл не содержит значимых данных')
                    print_value(f'End')

                # -----------------------------------SLK----------------------------------------
                if extension == 'slk' and app_settings.ext_slk:
                    print_step2(f'Start SLK: {file_full_path}')
                    try:
                        with open(file_full_path, 'r') as file:
                            _FileDataStr = file.read()
                            for _line in _FileDataStr.splitlines():
                                _str = str(_line.encode('UTF-8')).strip()
                                for key in a_dict:  # через словарь разбираем текст
                                    _str = _str.replace(key, a_dict[key])
                                new_item = parse_line(_str, new_item, 'SLK')
                    except Exception as e:
                        print_error(f'Error: {e}')
                    print_value(f'End')

                # -----------------------------------GCD----------------------------------------
                if extension == 'gcd' and app_settings.ext_gcd:
                    print_step2(f'Start GCD: {file_full_path}')
                    _TableInGDC = ''
                    new_item[CC.FORM_NAME] = 'VIP AltedDIC'
                    with open(file_full_path, 'r', encoding='cp866') as file:
                        for _line in file:
                            _str = str(_line)
                            if _str.upper().find('CREATE') != -1 and _str.upper().find('TABLE') != -1:
                                x = _str.upper().find('TABLE')
                                _TableInGDC = f'{_TableInGDC} {_str[x + 5:].strip()}'
                            if _str.upper().find('WITH') != -1 and _str.upper().find('TABLE_CODE') != -1:
                                x = _str.upper().find('=')
                                _TableInGDC = f'{_TableInGDC} ({_str[x + 1:].strip()});\n'
                            new_item = parse_line(_str, new_item, 'GCD')
                    new_item[CC.START_PATH] = 'VIP AltedDIC'
                    new_item[CC.SYSTEM_SOURCE_TYPE] = new_item.get(CC.START_PATH)
                    new_item[CC.NOTE] = _TableInGDC
                    print_value(f'End')
                # -----------------------------------FR3----------------------------------------
                # Only debug all *.fr3 must have deffres
                if extension == 'fr3' and app_settings.ext_fr3:
                    print_step2(f'Start FR3: {file_full_path}')
                    new_item[CC.FORM_NAME] = f'FR3 {cur_file}'
                    new_item[CC.START_PATH] = 'FR3'
                # ------------------------------ЗАПИСЬ В РЕЕСТР--------------------------------
                if new_item.get(CC.FORM_NAME) is not None:

                    # Update current property if empty
                    if new_item.get(CC.PROTOTYPE) is None:
                        new_item[CC.PROTOTYPE] = new_item.get(CC.INTERFACE)
                        if new_item.get(CC.SYSTEM_SOURCE_TYPE) is None:
                            new_item[CC.SYSTEM_SOURCE_TYPE] = 'Interface'

                    # Find unique item
                    _exist = False
                    for _el in p_items_list:
                        if _el.get(CC.CATALOG) == new_item.get(CC.CATALOG) and _el.get(CC.FORM_NAME) == new_item.get(
                                CC.FORM_NAME):
                            print_warning(
                                f'Update: {new_item.get(CC.CATALOG)}/{new_item.get(CC.FORM_NAME)}\nNEW: {new_item}\nOLD: {_el}')

                            # Update property from DEFRES
                            _el[CC.SYSTEM_SOURCE_TYPE] = new_item.get(CC.SYSTEM_SOURCE_TYPE)
                            _el[CC.VIPFCOM_BASENAME] = new_item.get(CC.VIPFCOM_BASENAME)
                            _el[CC.VIPFCOM_FILE_PATH] = new_item.get(CC.VIPFCOM_FILE_PATH)
                            _el[CC.PROTOTYPE] = new_item.get(CC.PROTOTYPE)
                            if _el.get(CC.NOTE) is None:
                                _el[CC.NOTE] = new_item.get(CC.NOTE)
                            else:
                                if new_item.get(CC.NOTE) is not None:
                                    _el[CC.NOTE] = _el[CC.NOTE] + '\n' + new_item[CC.NOTE]
                            _exist = True
                            print_warning(f'UPD: {_el}')

                    # Add unique item
                    if not _exist:
                        p_items_list.append(new_item)
                        print_result(f'Add: {new_item}')
                elif app_settings.show_skip:
                    print_error(f'Skip: can`t find form name\n{new_item}')
                print('\n')


# Create final list whit data
def create_list_from_def(p_process_dir, p_items_list):
    for _path, _sub_dirs, _files in os.walk(p_process_dir):
        _split_path = _path.split("\\")
        if _split_path[len(_split_path) - 1].casefold() == 'V_LAST'.casefold():
            print_step1(f'Parse folder: {_path}')
            res_name = get_res_public_name(_path)
            for cur_file in _files:
                if cur_file.upper().find('#DEFRES') != -1:  # File content '#DEFRES' in name
                    print_value(f'Parse file: {cur_file}')

                    # Create new Item
                    new_item = {}
                    new_item.clear()
                    new_item[CC.CATALOG] = _split_path[len(_split_path) - 2]  # Folder prev v_last (project name)
                    new_item[CC.DEFRES_BASENAME] = os.path.basename(cur_file)  # DEFRES file
                    new_item[CC.DEFRES_FILE_PATH] = os.path.join(_path, cur_file)  # DEFRES file full path
                    new_item[CC.DEBUG_NAME] = res_name  # *.res name from name.bat

                    # Read file and get tag
                    with open(os.path.join(_path, cur_file), 'r', encoding='cp866') as file:
                        file_content = file.read()

                        for line in file_content.splitlines():
                            line_upper = line.upper()
                            for cur_property in property_list: # Only function tag`s
                                if cur_property in line_upper: # Tag in line
                                    new_item[cur_property] = line.strip()[len(cur_property):].strip()

                    # Find unique item
                    _exist = False
                    for _el in p_items_list:
                        if _el.get(CC.CATALOG) == new_item.get(CC.CATALOG) and _el.get(CC.FORM_NAME) == new_item.get(
                                CC.FORM_NAME):
                            _exist = True
                            # System can`t contain double defres file to one source
                            print_error(
                                f'Error: {new_item.get(CC.CATALOG)}/{new_item.get(CC.FORM_NAME)}\n{new_item.get(CC.DEFRES_FILE_PATH)}')

                    # Add unique item
                    if not _exist:
                        print_result(f'Add: {new_item}')
                        p_items_list.append(new_item)

                    print('\n')


def main(app_settings):
    print_header('Starting process...')

    # Result list
    res_item_list = []

    # Process all subfolder
    for _path, _sub_dirs, _files in os.walk(app_settings.folder_process):
        if _path.endswith('v_last'):  # Process only 'v_last' folder, they content last version
            print_header(f'1. Work folder: {_path}')

            if app_settings.use_diff:
                print_header(f'1.1 File whit DEFF')
                create_list_from_def(_path, res_item_list)  # Parse #DEFRES file

            print_header(f'1.2 Source code file (update)')
            create_list_from_source(_path, res_item_list, app_settings)

    # Clear path (delete app folder 'c:\Client\Project1\File.vip' -> 'Project1\File.vip')
    for item in res_item_list:
        if item.get(CC.DEFRES_FILE_PATH) is not None and len(item[CC.DEFRES_FILE_PATH]) > len(app_settings.folder_process):
            item[CC.DEFRES_FILE_PATH] = item[CC.DEFRES_FILE_PATH][len(app_settings.folder_process):].lstrip('\\')
        if item.get(CC.VIPFCOM_FILE_PATH) is not None and len(item[CC.VIPFCOM_FILE_PATH]) > len(
                app_settings.folder_process):
            item[CC.VIPFCOM_FILE_PATH] = item[CC.VIPFCOM_FILE_PATH][len(app_settings.folder_process):].lstrip('\\')

    if settings_create_excel:
        print_header(f'2 Create Excel file in Result folder')
        # Create Excel file in Result folder
        cur_date = datetime.now()
        result_excel_file_path = os.path.join(app_settings.folder_result,
                                              f'{cur_date.strftime("%Y%m%d_%H-%M-%S")} Реестр ресурсов.xlsx')
        print_header(f'2.1 File: {result_excel_file_path}')
        workbook = xlsxwriter.Workbook(result_excel_file_path)
        worksheet = workbook.add_worksheet()

        #  Excel head
        create_excel_column()
        head_format = workbook.add_format({'bold': True, 'italic': True, 'align': 'center', 'text_wrap': True})
        for row in ListExcelColumn:
            row_num = row.get("Num")
            row_width = row.get("Width")
            row_name = row.get("Name")
            worksheet.write(0, row_num, row_name, head_format)
            worksheet.set_column(row_num, row_num, row_width)

        cell_format = workbook.add_format({'text_wrap': True})
        cell_format_off = workbook.add_format({'italic': True, 'font_color': 'Gray'})
        #  Excel row
        _ind = 1
        for item in sorted(res_item_list, key=lambda i: i[CC.CATALOG]):
            cur_format = cell_format

            flist = []

            # Folder (source)
            if item.get(CC.VIPFCOM_FILE_PATH) is not None:
                flist.append(item.get(CC.VIPFCOM_FILE_PATH))

            # Folder (def) (source can be empty)
            if item.get(CC.DEFRES_FILE_PATH) is not None:
                flist.append(item.get(CC.DEFRES_FILE_PATH))

            # All subfolder
            for fitem in flist:
                path = Path(fitem)
                folder_list = [parent.name for parent in path.parents]
                for folder in folder_list:
                    if folder[:4].lower() == 'off ' or folder.lower().find('отключено') != -1:
                        cur_format = cell_format_off

            for row in ListExcelColumn:
                row_num = row.get("Num")
                row_cell_name = row.get("CellName")
                worksheet.write(_ind, row_num, item.get(row_cell_name), cur_format)
            _ind += 1
        workbook.close()
    else:
        print_header(f'Skip create Excel')


if __name__ == "__main__":
    print_header('Create: Maxim Cherepanov masygreen@gmail.com (c), 05.2020-11.2025')
    ListExcelColumn = []

    # Settings file name
    settings_file_name = "config.json"

    # Read or create Settings
    is_exist, app_settings = read_settings(settings_file_name)

    if app_settings.folder_result is None or len(app_settings.folder_result) == 0:
        app_settings.folder_result = os.getcwd()

    if not is_exist:
        print_error(f'Create default setings: {settings_file_name}')
        print_result(f'\n*Press Space to Exit...')
        keyboard.wait("space")
        sys.exit(0)
    elif not os.path.isdir(app_settings.folder_process) or not os.path.isdir(app_settings.folder_result):
        print_error(f'*Please fill the configuration file: {settings_file_name}')
        if not os.path.isdir(app_settings.folder_process):
            print_error(f'*Not exist "folder: {app_settings.folder_process}"')
        if not os.path.isdir(app_settings.folder_result):
            print_error(f'*Not exist "folder: {app_settings.folder_result}"')
        print_result(f'\n*Press Space to Exit...')
        keyboard.wait("space")
        sys.exit(0)
    else:
        # Comment
        print_value(f'Folder (process): {app_settings.folder_process}')
        print_value(f'Folder (result): {app_settings.folder_result}')
        print_value(f'Process *.vip: {app_settings.ext_vip}')
        print_value(f'Process *.slk: {app_settings.ext_slk}')
        print_value(f'Process *.rtf: {app_settings.ext_rtf}')
        print_value(f'Process *.gcd: {app_settings.ext_gcd}')
        print_value(f'Process *.fr3: {app_settings.ext_fr3}')

        # Debug
        settings_create_excel = True
        print_header('Press Space to continue... (It the longest shortcut \\_(o0)_\\)')
        key_name = keyboard.read_key()

        if key_name == "space":
            try:
                main(app_settings)
                print_result(f'\n\nAll Process done.')
            except Exception as e:
                print_error(f'\n\nError: {e}')
            print_result(f'\n*Press any key to Exit...')
            keyboard.wait("space")
        else:
            print_result(f'Cancel parse...\n*Press Space to Exit...')
            keyboard.wait("space")