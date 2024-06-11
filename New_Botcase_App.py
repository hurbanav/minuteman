# -*- coding: utf-8 -*-
"""
Created on Wed Feb 21 21:38:58 2024

@author: Henrique Urbanavicius
"""

import sys
from upload_xml_v3 import UploadXML
from move_archives import organize_files
from extract_data_syscare import extrai_relaorio_full_process
from _tkinter import TclError
from tkinter_utils import TkinterUtils, TextRedirector




class FrontApps:
    def __init__(self, bg_color="#ededed",
                 font_size=20,
                 font_color="#1e313a",
                 padx=10, pady=5,
                 font_name='Segoe UI Bold'):
        self.bg_color = bg_color
        self.font_name = font_name
        self.padx = padx
        self.pady = pady
        self.font_size = font_size
        self.font_color = font_color
        self.Tk = TkinterUtils(
            title='Application Main Menu',
            bg_color=bg_color,
            font_size=font_size,
            font_color=font_color,
            padx=padx, pady=pady,
            font_name=font_name,
            width=800,
            height=800,
            title_color='#ffffff',
            hover_color='#a1622b'
        )
        self.app = self.Tk.app
        self.create_main_menu()  # Passa app como argumento para criar o menu principal

    def create_menu_button(self):
        self.Tk.add_button(text="Menu Principal", command=lambda: self.create_main_menu())

    def create_app_logo(self, scale=0.5):
        image_path = r'C:\Users\bi\Documents\BotCase\Dash\image.png'
        self.Tk.add_image(image_path, image_scale=scale)

    def create_organize_files_page(self):
        self.Tk.cleanup_page()

        self.create_menu_button()

        self.app.title('Organizador de Arquivos')

        self.create_app_logo()

        self.Tk.add_label(text="Selecione a pasta com os arquivos a serem organizados:")
        file_path_entry = self.Tk.add_entry()
        file_path_entry.insert(0, r'C:\Users\bi\Documents\BotCase\Dash\Teste Dir\tiss')

        self.Tk.add_button(text="Selecionar", command=lambda: self.Tk.open_file_dialog(file_path_entry))
        self.Tk.add_button(text="Organizar Arquivos", command=lambda: organize_files(file_path_entry.get()))

    def create_data_extract_page(self):
        self.Tk.cleanup_page()

        self.create_menu_button()

        self.app.title('Extração de Dados')

        self.create_app_logo()

        self.Tk.add_label(text="Extrair relatório de orçamentos aprovados:")

        self.Tk.add_button(text="Extrair Dados", command=lambda: extrai_relaorio_full_process())

    def run_main_function(self):
        self.username = self.Tk.check_entry_type(self.username_entry)
        self.password = self.Tk.check_entry_type(self.password_entry)
        self.file_path = self.Tk.check_entry_type(self.file_path_entry)
        upload_xml = UploadXML(self.username, self.password)
        self.Tk.start_task(target_function=upload_xml.main, variables=(self.file_path,))

    def create_xml_app(self):
        self.Tk.cleanup_page()
        app = self.app

        self.Tk.cleanup_page()

        self.create_menu_button()

        app.title('XML Upload')

        self.create_app_logo(scale=0.2)

        self.Tk.add_label(text="Usuário:")
        self.username_entry = self.Tk.add_entry()
        #self.username_entry = self.Tk.add_entry(insert='hurbanav')


        self.Tk.add_label(text="Senha:")
        self.password_entry = self.Tk.add_entry(show="*")
        #self.password_entry = self.Tk.add_entry(insert='Amarelinha12*', show="*")

        self.Tk.add_label(text="Selecione a pasta com os arquivos:")
        self.file_path_entry = self.Tk.add_entry()
        self.Tk.add_button(text="Selecionar", command=lambda: self.Tk.open_file_dialog(self.file_path_entry))
        self.Tk.add_button(text="Enviar Arquivos", command=lambda: self.run_main_function())


        output_box = self.Tk.add_text_output()
        redirector = TextRedirector(output_box)
        sys.stdout = redirector
        sys.stderr = redirector

    def create_main_menu(self):
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__

        self.Tk.cleanup_page()
        self.create_app_logo()
        self.app.title("BotCase Solutions")
        self.Tk.add_button(text="Envio de XMLs", command=lambda: self.create_xml_app())
        self.Tk.add_button(text="Organizar Arquivos", command=lambda: self.create_organize_files_page())
        self.Tk.add_button(text="Extrair Relatório", command=lambda: self.create_data_extract_page())

    def run(self):
        self.Tk.run()


def open_main_menu():
    FrontApps().Tk.cleanup_page()
    FrontApps().create_main_menu()


def create_app():
    front_app_instance.create_main_menu()


try:
    front_app_instance = FrontApps(bg_color="#ededed",
                                   font_size=20,
                                   font_color="#1e313a",
                                   padx=10, pady=5,
                                   font_name='Segoe UI',
                                   )
    front_app_instance.run()
    create_app()

except Exception as e:
    if 'application has been destroyed' in str(e):
        pass
    print(f'Exception: {e}')
