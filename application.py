import kivy

from os.path import expanduser, dirname

from kivy.app import App
from kivy.config import Config
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.image import Image
from kivy.uix.gridlayout import GridLayout
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.uix.popup import Popup, PopupException
from kivy.uix.label import Label

import kivy.resources
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.utils import platform

from get_info import ExcelInfo
from wordfile3cols import Word3Cols

Config.set('graphics', 'resizable', False)
Config.set('graphics', 'width', '500')
Config.set('graphics', 'height', '600')
Config.write()

INVERSE_MONTH = {
    "Janeiro": 1,
    "Fevereiro": 2,
    "Março": 3,
    "Abril": 4,
    "Maio": 5,
    "Junho": 6,
    "Julho": 7,
    "Agosto": 8,
    "Setembro": 9,
    "Outubro": 10,
    "Novembro": 11,
    "Dezembro": 12
}


class Application(ScreenManager):

    def __init__(self, **kwargs):
        super(Application, self).__init__(**kwargs)

        """Main Screen image"""
        self.screen_image = kivy.resources.resource_find("padariacentral.png")

        """Main screen"""
        self.main_screen = Screen(name='Padaria Central')
        self.layout = BoxLayout(orientation='vertical')
        self.logo = Image(source=self.screen_image, allow_stretch=True, size_hint_y=1)
        self.layout_down = GridLayout(cols=2, rows=1, size_hint_y=0.2)
        self.layout.add_widget(self.logo)
        self.layout.add_widget(self.layout_down)
        self.main_screen.add_widget(self.layout)

        """Define os and default path"""
        self.os = platform
        self.path = ""
        if self.os == 'win':
            self.path = dirname(expanduser("~"))
        else:
            self.path = expanduser("~")

        """ Main screen buttons"""
        self.generate_payments = Button(text='Tirar contas do mês', size_hint_y=0.1)
        self.generate_payments.bind(on_press=lambda x: self.file_chooser("pay"))

        self.send_warnings = Button(text='Enviar aviso para Clientes', size_hint_y=0.1)
        self.send_warnings.bind(on_press=lambda x: self.file_chooser("info"))

        self.layout_down.add_widget(self.generate_payments)
        self.layout_down.add_widget(self.send_warnings)

        # init default popup
        self.popup = Popup()

        """Screen Manager"""
        self.s_open_file = Screen(name='Selecionar Ficheiro')
        self.add_widget(self.main_screen)
        self.add_widget(self.s_open_file)

        self.s_save_file = Screen(name='Gravar Ficheiro')
        self.add_widget(self.s_save_file)

        """Init"""
        self.text_input = TextInput()
        self.text = ""

    def file_chooser(self, op):
        mp_layout = GridLayout(cols=1)

        info = Label(text="Selecionar ficheiro excel!")
        confirm = Button(text="Ok")

        mp_layout.add_widget(info)
        mp_layout.add_widget(confirm)

        confirm.bind(on_press=lambda x: self.browser(op))
        self.popup = Popup(title="Gerar ficheiro", separator_height=0, content=mp_layout,
                           size_hint=(None, None), size=(300, 150))
        self.popup.open()

    def browser(self, op):
        """ This function creates the file chooser to select image"""
        # Create Layout for popup
        try:
            self.popup.dismiss()
        except PopupException:
            pass
        self.current = 'Selecionar Ficheiro'
        b_main_lay = GridLayout(rows=2)
        action_layout = BoxLayout(orientation='horizontal', size_hint_y=0.1)
        file = FileChooserIconView(path=self.path, size_hint_y=0.9, multiselect=False)
        # this popup buttons and actions
        select = Button(text='Selecionar')
        select.bind(on_press=lambda x: self.open(file.selection, op))
        cancel = Button(text='Cancelar')

        cancel.bind(on_press=self.cancel_callback)
        action_layout.add_widget(select)
        action_layout.add_widget(cancel)
        b_main_lay.add_widget(file)
        b_main_lay.add_widget(action_layout)

        self.s_open_file.add_widget(b_main_lay)

    def cancel_callback(self, instance):
        self.current = 'Padaria Central'
        self.s_open_file.clear_widgets()

    def open(self, filename, op):
        self.current = 'Padaria Central'
        self.s_open_file.clear_widgets()
        if op == "pay":
            try:
                grid_s = GridLayout(rows=2)

                confirm = Button(text="Selecionar")

                spinner = Spinner(
                    # default value shown
                    text='Selecione o mês',
                    text_autoupdate=True,
                    # available values
                    values=('Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto',
                            'Setembro', 'Outubro', 'Novembro', 'Dezembro'),
                    # just for positioning in our example
                    size_hint=(None, None),
                    size=(275, 40),
                    pos_hint={'center_x': .5, 'center_y': .5})

                confirm.bind(on_press=lambda x: self.generate(spinner.text, filename, op, None))

                grid_s.add_widget(spinner)
                grid_s.add_widget(confirm)

                self.popup = Popup(title="Escolher Mês", separator_height=0, content=grid_s,
                                   size_hint=(None, None), size=(300, 150))
                self.popup.open()

            except FileExistsError as e:
                self.end_action(
                    "Oops! Algo correu mal!\nVerifique se selecionou o ficheiro correto")
        elif op == "info":
            try:
                grid_s = GridLayout(cols=1, rows=4)
                label = Label(text="Mensagem a enviar para clientes:", size_hint=(.5, .3))
                self.text_input = TextInput(text="", multiline=True, size_hint =(.5, .8))
                self.text_input.bind(text=self.set_text)
                confirm = Button(text="Seguinte", size_hint=(.5, .2))
                cancel = Button(text="Cancelar", size_hint=(.5, .2))

                cancel.bind(on_press=lambda x: self.popup.dismiss())
                confirm.bind(on_press=lambda x: self.generate(None, filename, op, self.text_input.text))

                grid_s.add_widget(label)
                grid_s.add_widget(self.text_input)
                grid_s.add_widget(confirm)
                grid_s.add_widget(cancel)

                self.popup = Popup(title="Mensagem", separator_height=0, content=grid_s,
                                   size_hint=(None, None), size=(400, 350))
                self.popup.open()
            except Exception as e:
                self.end_action("Erro, tente novamente!")
                print(e)

        else:
            self.end_action(
                "Oops! Algo correu mal!\nVerifique se selecionou o ficheiro correto")

    def end_action(self, text):
        self.current = 'Padaria Central'
        grid = GridLayout(rows=2)
        label = Label(text=text)
        dismiss = Button(text='OK', size_hint_y=None, size=(50, 30))
        grid.add_widget(label)
        grid.add_widget(dismiss)
        self.popup = Popup(title="Padaria Central", separator_height=0, content=grid, size_hint=(None, None),
                           size=(300, 200))
        self.popup.open()
        dismiss.bind(on_press=self.popup.dismiss)

    def show_selected_value(self, instance, text):
        """ Get current value from spinner """
        if text is not "Tipo de ficheiro" and not '':
            return text
        else:
            print("Invalid file extension")

    def generate(self, month, filename, op, infotext):
        try:
            self.popup.dismiss()

            mp_layout = GridLayout(cols=1)

            info = Label(text="Selecionar pasta e nome do ficheiro Word")
            confirm = Button(text="Ok")

            mp_layout.add_widget(info)
            mp_layout.add_widget(confirm)

            confirm.bind(on_press=lambda x: self.popup.dismiss())
            self.popup = Popup(title="Guardar Ficheiro", separator_height=0, content=mp_layout,
                               size_hint=(None, None), size=(400, 150))
            self.popup.open()
        except PopupException:
            pass

        try:
            self.current = 'Gravar Ficheiro'
            grid_s = GridLayout(rows=3)
            action_layout = BoxLayout(orientation='horizontal', size_hint_y=0.1)
            file_name_type_layout = BoxLayout(orientation='horizontal', size_hint_y=0.1)

            self.text_input = TextInput(text="", multiline=False)
            file_name_type_layout.add_widget(self.text_input)
            self.text_input.bind(text=self.set_text)

            save = Button(text="Guardar")
            cancel = Button(text="Cancelar")
            action_layout.add_widget(save)
            action_layout.add_widget(cancel)
            file = FileChooserIconView(path=self.path, )
            grid_s.add_widget(file)

            grid_s.add_widget(file_name_type_layout)
            grid_s.add_widget(action_layout)
            self.s_save_file.add_widget(grid_s)

            cancel.bind(on_press=self.cancel_callback)

            save.bind(on_press=lambda x: self.proc(month, filename, file.path, self.text_input, op, infotext))
        except:
            self.end_action("Algo correu mal, tente novamente!")

    def proc(self, month, input_file_path, output_file_path, output_file_name, op, infotext):

        filename = output_file_name.text

        if not filename and op == "pay":
            filename = "Contas" + month

        if not filename and op == "info":
            filename = "InfoClientes"

        if self.os == 'win':
            output_path = output_file_path + "\\" + filename + ".docx"
        else:
            output_path = output_file_path + "/" + filename + ".docx"

        try:
            if op == 'pay':
                clients = ExcelInfo(input_file_path[0]).get()
                Word3Cols(clients, INVERSE_MONTH[month], output_path, "pay", None).create()
                self.end_action("Ficheiro guardado com sucesso!")
            else:
                clients = ExcelInfo(input_file_path[0]).get()
                Word3Cols(clients, None, output_path, op, infotext).create()
                self.end_action("Ficheiro guardado com sucesso!")
        except Exception as e:
            self.end_action("ERRO, tente novamente!")
            print(e)

    def set_text(self, instance, input):
        """ Workaround to save input from textInput """
        self.text = input


class PadariaCentralApp(App):

    def build(self):
        self.title = 'Padaria Central'
        return Application()


if __name__ == '__main__':
    PadariaCentralApp().run()
