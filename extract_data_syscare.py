# -*- coding: utf-8 -*-
"""
Created on Wed Dec 27 21:32:43 2023

@author: bi
"""
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
from pyautogui import hotkey as hk
import time
from datetime import datetime
import win32com.client as win32
import pandas as pd
import os
from datetime import datetime
from PIL import ImageGrab
import win32com.client


from datetime import datetime, timedelta

class ExtraiRelatorio:
    def __init__(self, app):
        self.app = app

    def get_identifiers(self, app):
        try:
            main_window = app.window(
                title_re='.*SysCare - Sistema de Gestão de Home Care*')
            main_window.wait('visible', timeout=10)
            main_window.print_control_identifiers()
        except Exception as e:
            print("Erro ao localizar a janela:", e)

    def clicar_menu_item(self, main_window, **kwargs):
        menu_item = main_window.child_window(control_type='MenuItem', **kwargs)
        if menu_item.is_visible() and menu_item.is_enabled():
            menu_item.click_input()
        else:
            print("MenuItem não está visível ou habilitado.")

    def preencher_campo_edit(self, main_window, auto_id, texto, **kwargs):
        campo_edit = main_window.child_window(
            auto_id=auto_id, control_type="Edit", **kwargs)
        campo_edit.set_text(texto)

    def selecionar_item_combobox(self, main_window, auto_id, texto, **kwargs):
        combobox = main_window.child_window(
            auto_id=auto_id, control_type="ComboBox", **kwargs)
        combobox.select(texto)

    def clicar_botao_salvar(self, main_window):
        """
        Clica no botão 'Salvar' dentro da janela principal.

        :param main_window: A janela principal onde o botão 'Salvar' está localizado.
        """
        try:
            # Encontra o botão 'Salvar' pelo AutomationId e ClassName e clica nele
            botao_salvar = main_window.child_window(
                auto_id="1", control_type="Button", class_name="Button")
            if botao_salvar.is_visible() and botao_salvar.is_enabled():
                botao_salvar.click_input()
            else:
                print("Botão 'Salvar' não está visível ou habilitado.")
        except Exception as e:
            print(f"Não foi possível clicar no botão 'Salvar': {e}")

    def alterar_estado_checkbox(self, main_window, auto_id, marcar=True, **kwargs):
        checkbox = main_window.child_window(
            auto_id=auto_id, control_type="CheckBox", **kwargs)
        if marcar:
            if not checkbox.get_toggle_state():
                checkbox.toggle()
        else:
            if checkbox.get_toggle_state():
                checkbox.toggle()


def find_number_on_string(text):
    import re
    numbers = re.findall(r'\d+', text)
    if numbers:
        last_number = numbers[-1]
        print(last_number)
    else:
        print("No numbers found.")


def wait_for_window_visible(window, timeout=10):
    try:
        window.wait('visible', timeout=timeout)
        print("Window is now visible.")
    except TimeoutError:
        print(f"Window did not become visible within {timeout} seconds.")


def get_identifiers(app):
    try:
        main_window = app.window(
            title_re='.*SysCare - Sistema de Gestão de Home Care*')
        main_window.wait('visible', timeout=10)
        main_window.print_control_identifiers()
    except Exception as e:
        print("Erro ao localizar a janela:", e)


def clicar_menu_item(main_window, posicao=None, **kwargs):
    try:
        class_name = kwargs.pop('class_name', None)
        if posicao is None:
            menu_item = main_window.child_window(
                control_type="MenuItem", **kwargs)
            menu_item.wait("visible", timeout=60)
            if menu_item.is_visible() and menu_item.is_enabled():
                menu_item.click_input()
            else:
                print("MenuItem não está visível ou habilitado.")
        else:
            # Handle position-based menu item selection
            menu_items = main_window.children(control_type='MenuItem')
            if posicao < len(menu_items):
                menu_item = menu_items[posicao]
                menu_item.wait("visible", timeout=60)
                menu_item.click_input()
            else:
                print(
                    "Posição fornecida está fora do intervalo dos itens de menu disponíveis.")
    except Exception as e:
        print(f"Erro ao clicar no menu item: {e}")


def preencher_campo_edit(main_window, auto_id, texto, **kwargs):
    class_name = kwargs.pop('class_name', None)
    if class_name is not None:
        item = app.window(title_re=title_re, class_name=class_name).child_window(
            auto_id=auto_id, control_type="Edit", **kwargs)
    campo_edit = main_window.child_window(
        auto_id=auto_id, control_type="Edit", **kwargs)
    campo_edit.set_text(texto)


def generic_click(main_window, auto_id, **kwargs):
    class_name = kwargs.pop('class_name', None)
    if class_name is not None:
        item = app.window(title_re=title_re, class_name=class_name).child_window(
            auto_id=auto_id, **kwargs)
    item = main_window.child_window(auto_id=auto_id, **kwargs)
    item.click()


def generic_select(main_window, auto_id, texto, **kwargs):
    class_name = kwargs.pop('class_name', None)
    if class_name is not None:
        item = app.window(title_re=title_re, class_name=class_name).child_window(
            auto_id=auto_id, **kwargs)
    item = main_window.child_window(auto_id=auto_id, **kwargs)
    item.select(texto)


def selecionar_item_combobox(main_window, auto_id, texto, **kwargs):
    combobox = main_window.child_window(
        auto_id=auto_id, control_type="ComboBox", **kwargs)
    combobox.expand()
    # Pode ser necessário usar combobox.expand() antes, se o combobox estiver fechado
    combobox.select(texto)


def alterar_estado_checkbox(main_window, auto_id, marcar=True, **kwargs):
    checkbox = main_window.child_window(
        auto_id=auto_id, control_type="CheckBox", **kwargs)
    if marcar:
        if not checkbox.get_toggle_state():
            checkbox.toggle()
    else:
        if checkbox.get_toggle_state():
            checkbox.toggle()


def pega_legacy_value_de_elemento(main_window, auto_id, control_type):
    """
    Retrieves the LegacyIAccessible.Value property of a UI element identified by auto_id and control_type.

    Parameters:
    - window: A pywinauto window or application object.
    - auto_id (str): The Automation ID of the target UI element.
    - control_type (str): The Control Type of the target UI element.

    Returns:
    - The value of the LegacyIAccessible.Value property, or None if not found or inaccessible.
    """
    try:
        # Find the child window/control with the specified auto_id and control_type
        target_control = main_window.child_window(
            auto_id=auto_id, control_type=control_type)

        # Check if the target exists
        if target_control.exists():
            # Retrieve the LegacyIAccessible.Value property
            legacy_value = target_control.legacy_properties().get('Value', None)
            return legacy_value
        else:
            print(
                f"Control with auto_id='{auto_id}' and control_type='{control_type}' not found.")
            return None
    except ElementNotFoundError:
        print(
            f"Error finding element with auto_id='{auto_id}' and control_type='{control_type}'.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}")
        return None


def data_faturamento_prox_mes(day=30):
    current_date = datetime.now()
    # Check if today is before the 20th of the current month
    if current_date.day < day:
        # Return the 20th of the current month
        billing_date = current_date.replace(day=day)
    else:
        # Calculate the first day of the next month
        next_month = current_date.replace(day=28) + timedelta(days=4)
        first_day_next_month = next_month.replace(day=1)
        # Return the 20th of the next month
        billing_date = first_day_next_month.replace(day=day)
    return billing_date.strftime('%d/%m/%Y')


def faz_login(main_window):

    # app.SyscareSistemaDeGestãoDeHomeCare.print_control_identifiers()
    #pass_command = app.SyscareSistemaDeGestãoDeHomeCareoDeHomeCare
    text_usuario = main_window.child_window(
        title="Digite os números da imagem ao lado (CAPTCHA)", auto_id="106", control_type="Edit")
    text_usuario.type_keys('HQIROBOT')
    text_password = main_window.child_window(
        title="Usuário", auto_id="108", control_type="Edit")
    text_password.type_keys('123')
    confirma_button = main_window.child_window(
        title="Confirma", auto_id="101", control_type="Button")
    confirma_button.click()


def redefine_app(class_name=None, normal=False):
    if normal == False:
        if class_name is not None:
            main_window = app.window(title_re=title_re)
            main_window.wait("visible", timeout=40)
        else:
            main_window = app.window(title_re=title_re, class_name=class_name)
            main_window.wait("visible", timeout=40)
    else:
        main_window = app.window(title_re=title_re)
    return main_window


def monta_relat(main_window, app):
    relatorios_menu = main_window.child_window(
        title="Relatorios", control_type="MenuItem")
    relatorios_menu.click_input()
    # Obter todos os itens do menu como uma lista
    menu_items = relatorios_menu.children()

    clicar_menu_item(main_window, found_index=0, )
    clicar_menu_item(main_window, auto_id="20136", found_index=0)

    preencher_campo_edit(main_window, auto_id="160",
                         texto=data_faturamento_prox_mes(), title="De")
    selecionar_item_combobox(main_window, auto_id="204",
                             texto='N-Não', found_index=0)
    generic_click(main_window, title="Imprimir",
                  auto_id="130", control_type="Button")


def export_relat(main_window, app):
    #tela_de_carregamento = main_window.child_window(auto_id="TitleBar", control_type="TitleBar")
    #tela_de_carregamento.wait('visible', timeout=20)

    main_window = app.window(title_re="RELATÓRIO DE ORCAMENTOS COM VALOR")
    main_window.wait('visible', timeout=500000)
    '''legacy_value = tela_de_carregamento.legacy_properties()['Value']
    print(legacy_value)
    if legacy_value != None:
        wait_value = find_number_on_string(legacy_value)*5
        print(wait_value)
        if isinstance(wait_value, int):
            wait_for_window_visible(main_window, timeout=wait_value)

    print('Element Loading_Warning not visible.') '''
    clicar_menu_item(main_window, title="File", class_name="TWINDOW")
    clicar_menu_item(main_window, auto_id="20287",
                     found_index=0, class_name="#32768")
    generic_click(main_window, auto_id="1",
                  title="Salvar", control_type="Button")
    
    
def excel_dashboard_to_image(excel_path, sheet_name, range_string, image_path):
    # Initialize COM client
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # You can set this to False if you don't want the Excel application to be visible
    workbook = excel.Workbooks.Open(os.path.abspath(excel_path))

    # Update all connections
    for connection in workbook.Connections:
        connection.Refresh()

    # Wait for refresh to finish
    excel.CalculateUntilAsyncQueriesDone()

    # Get the specified sheet and range
    sheet = workbook.Sheets(sheet_name)
    range_cells = sheet.Range(range_string)
    range_cells.CopyPicture(Format=win32com.client.constants.xlBitmap)

    # Save the image from the clipboard
    image = ImageGrab.grabclipboard()
    image.save(image_path, 'PNG')
    
    workbook.Close(SaveChanges=False)
    excel.Quit()


def read_data_from_excel_sheet_and_paste_to_history_sheet():
    directory = r'C:\Users\bi\Documents\BotCase\Dash'
    source_excel_path = os.path.join(directory, 'RELATÓRIO DE ORCAMENTOS COM VALOR.csv')
    target_excel_path = os.path.join(directory, 'Dash Database.xlsx')
    

    # Read source data, adjust for decimal and thousands separators
    source_df = pd.read_csv(source_excel_path, skiprows=4, encoding='latin1', skipfooter=1, engine='python', decimal=',', sep=';', dtype=str)
    source_df['Valor'] = source_df['Valor'].str.replace('.', '').str.replace(',', '.').astype(float)

    # Create or update the historic dash database
    try:
        target_df = pd.read_excel(target_excel_path, sheet_name=0)
    except FileNotFoundError:
        target_df = pd.DataFrame(columns=source_df.columns)

    combined_df = pd.concat([target_df, source_df], ignore_index=True)
    combined_df.to_excel(target_excel_path, index=False)
    return target_excel_path


def create_updated_dash():
    target_excel_path = read_data_from_excel_sheet_and_paste_to_history_sheet()
    dash_filepath = os.path.join(os.path.dirname(target_excel_path), 'Dash.xlsm')
    image_path = os.path.join(os.path.dirname(target_excel_path), f'Dash History\Dashboard - {datetime.today().strftime("%y_%m_%d")}.png')
    excel_dashboard_to_image(dash_filepath, 'Dashboard', 'A1:V28', image_path)


def update_cell_in_excel_workbook(file_path, sheet_name, cell_reference, new_value):
    from openpyxl import load_workbook

    """
    Update a specific cell in an Excel workbook.

    Parameters:
    - file_path: Path to the Excel workbook (must be .xlsx).
    - sheet_name: Name of the sheet where the cell is located.
    - cell_reference: Cell reference in Excel format (e.g., 'A1').
    - new_value: New value to set in the specified cell.
    """
    # Load the workbook and select the specified sheet
    workbook = load_workbook(filename=file_path)
    sheet = workbook[sheet_name]

    # Update the cell with the new value
    sheet[cell_reference] = new_value

    # Save the workbook
    workbook.save(filename=file_path)
    print(
        f"Workbook updated: {file_path}, Sheet: {sheet_name}, Cell: {cell_reference}, New Value: {new_value}")


def refresh_excel_workbook(file_path):
    """
    Open an Excel workbook, refresh all data connections, and save it.
    """
    # Start an instance of Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    # Make Excel visible to the user (set to False to run in the background)
    excel.Visible = True

    # Open the workbook
    workbook = excel.Workbooks.Open(file_path)

    # Refresh all data connections
    workbook.RefreshAll()

    # Wait until Excel finishes refreshing all data
    excel.CalculateUntilAsyncQueriesDone()

    # Save the workbook
    workbook.Save()

    # Close the workbook and quit Excel
    workbook.Close()
    excel.Quit()

    print(f"Workbook refreshed and saved: {file_path}")


def extrai_relaorio_full_process():
    import os
    global main_window, app, title_re
    os.chdir(r'S:\See')
    title_re = '.*SysCare - Sistema de Gestão de Home Care*'
    app = Application(backend="uia").start(r'S:\See\cCase.exe')
    app.allow_magic_lookup = False
    main_window = app.window(title_re=title_re)
    main_window.wait("visible", timeout=40)

    faz_login(main_window)

    main_window = app.window(title_re=title_re)
    main_window.wait("visible", timeout=40)

    monta_relat(main_window, app)

    export_relat(main_window, app)

    create_updated_dash()
    
# operadoras Ana Costa, Prevent, Intermedica possuem datas específicas para a geralçao de relatorio; demais seguem primeiro dia até o ultimo dia, 
# criar input primeiro dia
