import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter import scrolledtext, filedialog, END
import threading
import sys
from _tkinter import TclError


def print_on_gui(text='', separator='*', size_separator=50):
    print('\n')
    print(f'{text}')
    print('\n')
    print(separator * size_separator)

def resource_path(relative_path):
    import os
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class TextRedirector:
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, message):
        if not self.widget:
            raise ValueError("TextRedirector widget is not initialized.")
        if "INFO" in message or "ERROR" in message:
            self.output_to_ide(message)
        else:
            self.output_to_widget(message)

    def output_to_ide(self, message):
        try:
            original_stdout = sys.__stdout__
            original_stdout.write(message)
            original_stdout.flush()
        except Exception:
            pass  # Silently ignore if there's no IDE

    def output_to_widget(self, message):
        self.widget.configure(state='normal')
        self.widget.insert('end', message)
        self.widget.configure(state='disabled')
        self.widget.yview('end')

    def flush(self):
        pass  # No action needed

class TkinterUtils(TextRedirector):
    def __init__(self,
                 title: str = "CustomTkinter App",
                 width: int = 400, height: int = 300,
                 bg_color: str = "#ededed",
                 font_size: int = 20,
                 font_color: str = "white",
                 font_name: str = 'Segoe UI',
                 title_color: str = "black",
                 hover_color: str = "black",
                 padx: int = 10, pady: int = 5):
        self.pady = pady
        self.padx = padx
        self.font_name = font_name  # Correct assignment here
        self.font_color = font_color
        self.font_size = font_size
        self.hover_color = hover_color
        self.title_color = title_color
        self.bg_color = bg_color
        self.app = ctk.CTk()
        self.app.title(title)
        self.app.geometry(f"{width}x{height}")
        if bg_color is not None:
            self.app.configure(bg=self.bg_color)

    def add_button(self, text="Click Me", command=None):
        button = ctk.CTkButton(master=self.app,
                               text=text,
                               command=command,
                               text_color=self.title_color,
                               font=(self.font_name, self.font_size),
                               hover_color=self.hover_color,
                               fg_color=self.font_color,
                               )
        button.pack(pady=self.pady, padx=self.padx)

    def add_entry(self, insert='', show=None):
        entry = ctk.CTkEntry(master=self.app, text_color=self.font_color, font=(self.font_name, self.font_size))
        if show is not None:
            entry.configure(show='*')
        entry.pack(pady=self.pady, padx=self.padx, fill='x')
        entry.insert(0, insert)
        return entry

    @staticmethod
    def check_entry_type(entry):
        if isinstance(entry, str):
            return entry
        if entry is None or entry == '':
            return None
        else:
            return entry.get()

    def add_text_output(self):
        output_box = scrolledtext.ScrolledText(self.app, height=10)
        output_box.pack(pady=self.pady, padx=self.padx, fill='both', expand=True)
        return output_box

    def add_label(self, text=""):
        label = ctk.CTkLabel(master=self.app,
                             text=text,
                             text_color=self.font_color,
                             font=(self.font_name, self.font_size))
        label.pack(pady=self.pady, padx=self.padx)
        return label

    def add_combobox(self, values=None):
        if values is None:
            values = []
        combobox = ctk.CTkComboBox(master=self.app,
                                   values=values,
                                   text_color=self.font_color,
                                   font=(self.font_name, self.font_size))
        combobox.pack(pady=self.pady, padx=self.padx)
        return combobox

    def add_image(self, image_path, image_scale=1):
        import os
        import sys

        image_path = resource_path(r"C:\Users\bi\Documents\BotCase\Dash\image.png")

        try:
            image = Image.open(resource_path(image_path))
            width, height = image.size
            width = int(width * image_scale)
            height = int(height * image_scale)
            image = image.resize((width, height), Image.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            label = ctk.CTkLabel(master=self.app,
                                 text='',
                                 image=photo,
                                 text_color=self.font_color)
            label.image = photo  # Keep a reference to avoid garbage collection
            label.pack(pady=self.pady, padx=self.padx)
            return label
        except Exception as e:
            print(f"Image loading failed: {e}")
            print_on_gui('Image loading failed.')

    def run(self):
        try:
            self.app.mainloop()
        except TclError as e:
            print(f"Error during mainloop: {e}")

    @staticmethod
    def open_file_dialog(entry_widget):
        import os
        # Use 'self.app' as the parent window for the 'askdirectory' dialog.
        file_path = filedialog.askopenfilename()
        entry_widget.delete(0, END)
        entry_widget.insert(0, os.path.dirname(file_path))

    def cleanup_page(self):
        try:
            if self.app.winfo_exists():
                for widget in self.app.winfo_children():
                    if widget.winfo_exists():
                        widget.destroy()
        except TclError as e:
            if 'application has been destroyed' in str(e):
                pass
            print(f"Error during cleanup: {e}")
            pass

    @staticmethod
    def start_task(target_function, variables):
        threading.Thread(target=target_function, args=variables, daemon=True).start()
