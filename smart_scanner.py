import time
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import datetime
import os
import getpass
import socket
import ctypes
import configparser
import segno
from openpyxl.drawing.image import Image
import xlwings as xw
import pyodbc
import logging


def to_eng():
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        tid = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
        current = ctypes.windll.user32.GetKeyboardLayout(tid) & 0xFFFF

        max_attempts = 9
        attempts = 0

        while (current & 0x00FF) != 0x09 and attempts < max_attempts:
            ctypes.windll.user32.PostMessageW(hwnd, 0x0050, 0, 0)
            time.sleep(0.05)

            hwnd = ctypes.windll.user32.GetForegroundWindow()
            tid = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
            current = ctypes.windll.user32.GetKeyboardLayout(tid) & 0xFFFF
            attempts += 1

        return (current & 0x00FF) == 0x09
    except:
        return False


def is_eng():
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
        layout = ctypes.windll.user32.GetKeyboardLayout(thread_id)
        layout_id = layout & 0xFFFF

        # английская раскладка - код 0x0409 или 0x09
        return (layout_id & 0x00FF) == 0x09 or layout_id == 0x0409
    except:
        return False


# логирование
log_path = 'logs/'
os.makedirs(log_path, exist_ok=True)
logging.basicConfig(
    filename=f'{log_path}logs.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# конфиг
config = configparser.ConfigParser()
if not os.path.exists('config.ini'):
    # конфиг по умолчанию
    config['settings'] = {
        'qr_path': '',
        'appearance_mode': 'dark'
    }
    config['database'] = {
        'sql_server': '',
        'sql_db': '',
        'sql_user': '',
        'sql_password': '',
    }
    with open('config.ini', 'w', encoding='utf-8') as f:
        config.write(f)

config.read('config.ini', encoding='utf-8')
qr_path = config['settings']['qr_path']
appearance_mode = config['settings']['appearance_mode']

# colors
border_color_base = '#808080'

border_color_green = '#2E8B57'
border_color_yellow = '#DAA520'
border_color_red = '#DC143C'

notification_color_green = '#2E8B57'
notification_color_yellow = '#DAA520'
notification_color_red = '#800000'

if 'light' in appearance_mode:
    fg_color_disable = '#b8b8b8'
    fg_color_enable = '#ffffff'
else:
    fg_color_disable = '#202121'
    fg_color_enable = '#343536'

# внешний вид
ctk.set_appearance_mode(appearance_mode)
ctk.set_default_color_theme('dark-blue')

# переключение на en с проверкой
if not is_eng():
    to_eng()


class DatabaseManager:
    def __init__(self):
        self.SQL_SERVER = config['database']['SQL_SERVER']
        self.SQL_DB = config['database']['SQL_DB']
        self.SQL_USER = config['database']['SQL_USER']
        self.SQL_PASSWORD = config['database']['SQL_PASSWORD']
        self.conn_str = f'DRIVER={{SQL Server}};SERVER={self.SQL_SERVER};DATABASE={self.SQL_DB};UID={self.SQL_USER};PWD={self.SQL_PASSWORD}'

    def check_connection(self):
        try:
            conn = pyodbc.connect(self.conn_str)
            conn.close()
            return True
        except pyodbc.Error as e:
            logging.error(f'Ошибка проверки подключения к бд: {e}')
            # print(f'Connection check failed: {e}')
            return False

    def get_connection(self):
        try:
            return pyodbc.connect(self.conn_str)
        except pyodbc.Error as e:
            logging.error(f'Ошибка подключения к бд: {e}')
            # print(f'Connection error: {e}')
            return None

    def execute_query(self, query, params=None, fetch_one=False, fetch_all=False, commit=False):
        conn = None
        cursor = None
        try:
            conn = self.get_connection()
            if not conn:
                return None
            cursor = conn.cursor()
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            if commit:
                conn.commit()
                return True
            if fetch_one:
                return cursor.fetchone()
            elif fetch_all:
                return cursor.fetchall()
            return None
        except pyodbc.Error as e:
            return None
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def add_record(self, barcode, excise, user_name, computer_name, qr_name, created_date=None):

        query = '''
            INSERT INTO barcodb (barcode, excise, user_name, computer_name, qr_name, created_date)
            VALUES (?, ?, ?, ?, ?, ?)
        '''

        params = (barcode, excise, user_name, computer_name, qr_name, created_date)
        return self.execute_query(query, params, commit=True)

    def get_data(self):
        query = 'SELECT * FROM barcodb'
        return self.execute_query(query, fetch_all=True)

    def check_exists(self, excise):
        query = 'SELECT 1 FROM barcodb WHERE excise = ?'
        result = self.execute_query(query, params=(excise,), fetch_one=True)

        return 1 if result else 0


class ReportGenerator:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title('Scanner')
        self.db = DatabaseManager()

        # задаем размеры окна
        self.width = 650
        self.height = 300
        self.root.geometry(f'{self.width}x{self.height}')

        self.root.resizable(True, False)
        self._notification_timer = None

        # ждем применения геометрии
        self.root.update_idletasks()

        # получаем реальные размеры окна после применения масштабирования
        actual_width = self.root.winfo_width()

        # находим dpi scaler
        scaler = round(actual_width / self.width, 2)

        # получаем размеры экрана
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # вычисляем позицию для центрирования
        x = (screen_width - self.width) // 2
        y = (screen_height - self.height) // 2

        # устанавливаем окончательную позицию
        self.root.geometry(f'+{round(x * scaler)}+{round(y * scaler)}')

        self.create_widgets()

        self.root.after(100, self.entry_barcode.focus)

    def create_widgets(self):
        # основной контейнер
        main_container = ctk.CTkFrame(self.root, fg_color='transparent')
        main_container.pack(fill='both', expand=True, padx=30, pady=10)

        # индикатор подключения к БД
        self.connection_indicator = ctk.CTkLabel(
            main_container,
            text='●',
            font=ctk.CTkFont(size=15)
        )
        self.connection_indicator.pack(anchor='w', pady=(0, 5), padx=10)
        self.update_connection_indicator()

        # рамка для баркода
        self.barcode_frame = ctk.CTkFrame(main_container, border_width=2, border_color=border_color_base,
                                          corner_radius=10)
        self.barcode_frame.pack(fill='x', pady=(0, 20))

        label_barcode = ctk.CTkLabel(self.barcode_frame, text='Баркод:', font=ctk.CTkFont(size=14, weight='bold'))
        label_barcode.pack(side='left', padx=(20, 10), pady=20)

        self.entry_barcode = ctk.CTkEntry(self.barcode_frame, width=300, height=35)
        self.entry_barcode.pack(side='left', padx=(0, 20), pady=20, fill='x', expand=True)

        # рамка для акциза
        self.excise_frame = ctk.CTkFrame(main_container, border_width=2, border_color=border_color_base,
                                         corner_radius=10)
        self.excise_frame.pack(fill='x', pady=(0, 10))

        label_excise = ctk.CTkLabel(self.excise_frame, text='Акциз:', font=ctk.CTkFont(size=14, weight='bold'))
        label_excise.pack(side='left', padx=(20, 10), pady=20)

        self.entry_excise = ctk.CTkEntry(self.excise_frame, width=300, height=35, state='disabled',
                                         fg_color=fg_color_disable)
        self.entry_excise.pack(side='left', padx=(0, 20), pady=20, fill='x', expand=True)

        # кнопка формирования отчёта
        self.button_generate = ctk.CTkButton(
            main_container,
            text='Сформировать отчет',
            command=self.generate_report,
            width=200,
            height=45,
            fg_color=border_color_green,
            text_color='white',
            font=ctk.CTkFont(size=15, weight='bold'),
            hover_color='#3CB371'
        )
        self.button_generate.pack(side='right')

        # метка для уведомлений
        self.notification_label = ctk.CTkLabel(
            main_container,
            text='',
            font=ctk.CTkFont(size=15),
            fg_color=border_color_green,
            text_color='white',
            corner_radius=6,
            padx=15,
            pady=5
        )

        # привязываем события ввода
        self.entry_barcode.bind('<KeyRelease>', self.on_barcode_change)

    def update_connection_indicator(self):
        (color, state) = ('#2E8B57', 'connected') if self.db.check_connection() else ('#DC143C', 'disconnected')
        self.connection_indicator.configure(text_color=color, text=f'●  {state}')

    def send_data(self, barcode, excise):
        try:
            # проверка уникальности
            if not self.db.check_connection():
                print('not self.db.check_connection()')
                self.update_connection_indicator()
                self.show_notification(f'Нет соединения с БД', label_bg=notification_color_red)
                return False
            if self.db.check_exists(excise):
                self.show_notification(f'Данный акциз уже существует', label_bg=notification_color_red)
                return 'exist'

            # получаем данные о пользователе и компьютере
            user_name = getpass.getuser()
            computer_name = socket.gethostname()
            created_date = datetime.datetime.now()
            # created_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # генерируем qr
            try:
                os.makedirs(qr_path, exist_ok=True)
            except:
                ...

            qr_name = round(time.time())
            qr = segno.make(str(excise), error='h')
            qr.save(rf'{qr_path}{qr_name}.png',
                    scale=10,  # размер модуля в пикселях
                    border=2,  # отступ вокруг QR
                    dark='black',  # цвет темных модулей
                    light='white')  # цвет светлых модулей

            # отправляем insert
            if self.db.add_record(barcode, excise, user_name, computer_name, qr_name, created_date):
                self.show_notification(f'Данные успешно добавлены', label_bg=notification_color_green)
            else:
                self.show_notification(f'Ошибка в insert запросе', label_bg=notification_color_red)

            # insert into db (barcode, excise, user_name, computer_name, qr_name, created_date)

            # выводим все данные
            print('\n')
            print(f'Время: {created_date}')
            print(f'Пользователь: {user_name}')
            print(f'Компьютер: {computer_name}')
            print(f'Баркод: {barcode}')
            print(f'Акциз: {excise}')
            print(f'qr: {qr_name}')
            print('\n')

            return True
        except Exception as e:
            logging.error(f'Ошибка в отправке данных: {e}')
            self.show_notification(f'Ошибка в отправке данных', label_bg=notification_color_red)
            return False

    def on_barcode_change(self, event=None):
        barcode = self.entry_barcode.get()
        barcode_len = len(barcode)

        if not is_eng():
            self.show_notification(f'Неверная раскладка. Переключите на EN', label_bg=notification_color_red)
            self.entry_barcode.delete(0, tk.END)
        else:
            # обновление рамки баркода
            if barcode_len == 0:
                self.barcode_frame.configure(border_color=border_color_base)

            elif barcode_len == 13 and barcode.isdigit():
                # self.show_notification(f'EAN: OK', 2000)
                self.barcode_frame.configure(border_color=border_color_green)
                # активируем поле акциза и ставим фокус
                self.entry_excise.configure(state='normal')
                self.entry_excise.configure(fg_color=fg_color_enable)
                self.entry_excise.focus()
                self.excise_frame.configure(border_color=border_color_yellow)

                # привязываем обработчик для акциза только когда поле активировано
                if not hasattr(self, '_excise_bound'):
                    self.entry_excise.bind('<KeyRelease>', self.on_excise_change)
                    self._excise_bound = True
            else:
                self.entry_barcode.delete(0, tk.END)
                if barcode.isdigit():
                    self.show_notification(f'Неверный EAN (длина {barcode_len} из 13)', label_bg=notification_color_red)
                else:
                    self.show_notification(f'В EAN должны быть только цифры', label_bg=notification_color_red)
                self.barcode_frame.configure(border_color=border_color_red)
                # деактивируем поле акциза и очищаем его
                self.entry_excise.delete(0, tk.END)
                self.entry_excise.configure(fg_color=fg_color_disable)
                self.entry_excise.configure(state='disabled')
                self.excise_frame.configure(border_color=border_color_base)

    def show_notification(self, message, duration=3000, label_bg=notification_color_green):
        if self._notification_timer is not None:
            try:
                self.root.after_cancel(self._notification_timer)
            except:
                pass
            self._notification_timer = None

        self.hide_notification()

        self.notification_label.configure(text=message, fg_color=label_bg)
        self.notification_label.pack(side='left', pady=(5, 0))

        self._notification_timer = self.root.after(duration, self.hide_notification)

    def hide_notification(self):
        self.notification_label.pack_forget()
        self.notification_label.configure(text='')

        if self._notification_timer is not None:
            try:
                self.root.after_cancel(self._notification_timer)
            except:
                pass
            self._notification_timer = None

    def on_excise_change(self, event=None):
        excise = self.entry_excise.get()
        excise_len = len(excise)

        barcode = self.entry_barcode.get()
        barcode_len = len(barcode)

        if not is_eng():
            self.show_notification(f'Неверная раскладка. Переключите на EN', label_bg=notification_color_red)
            self.entry_excise.delete(0, tk.END)
        else:

            # обновление рамки акциза
            if barcode_len == 0:
                self.excise_frame.configure(border_color=border_color_base)
                self.entry_excise.delete(0, tk.END)
                self.entry_excise.configure(state='disabled')
                # self.barcode.focus()
                self.entry_barcode.focus()
            elif excise_len == 0:
                self.excise_frame.configure(border_color=border_color_yellow)
            elif excise_len == 13 and excise.isdigit():
                self.entry_barcode.delete(0, tk.END)
                self.entry_barcode.insert(0, excise)
                self.entry_excise.delete(0, tk.END)
                self.excise_frame.configure(border_color=border_color_yellow)
                self.show_notification(f'Баркод обновлён', label_bg=notification_color_yellow)
            elif 0 < excise_len < 150:
                self.excise_frame.configure(border_color=border_color_red)
                self.show_notification(f'Неверный акциз (длина {excise_len} из 150)', label_bg=notification_color_red)
            else:  # больше n символов
                self.excise_frame.configure(border_color=border_color_green)

                # фильтр баркода
                if barcode_len == 13 and barcode.isdigit():
                    # отправляем данные
                    # self.show_notification(f'Отправляем данные..', label_bg=notification_color_yellow)
                    code = self.send_data(barcode, excise)

                    # очищаем оба поля
                    if code != 'exist':
                        self.entry_barcode.delete(0, tk.END)
                    self.entry_excise.delete(0, tk.END)

                    # деактивируем поле акциза
                    self.entry_excise.configure(state='disabled')

                    # сбрасываем цвета рамок
                    self.barcode_frame.configure(border_color=border_color_base)
                    self.excise_frame.configure(border_color=border_color_base)

                    # устанавливаем курсор на поле баркода
                    self.entry_barcode.focus()

    def generate_report(self):

        if not self.db.check_connection():
            self.update_connection_indicator()
            return self.show_notification(f'Нет соединения с БД', label_bg=notification_color_red)

        # формирование отчета в excel
        # select * from bd --> excel
        # + вставка qr

        data = []
        if self.db.check_connection():
            all_data = self.db.get_data()
            if all_data:
                for row in all_data:
                    # print(row)
                    data.append(row)

        report_time = time.time()
        filename = os.path.abspath(f'report_{round(report_time)}.xlsx')

        # создаем новую книгу Excel
        wb = xw.Book()
        ws = wb.sheets[0]
        ws.name = 'report'

        # заголовки
        headers = ['Баркод', 'Акциз', 'QR Акциза', 'Путь к QR', 'Компьютер', 'Пользователь', 'Дата добавления']
        for col, header in enumerate(headers, 1):
            ws.range((1, col)).value = header
            ws.range((1, col)).font.bold = True
            # центрируем заголовки по горизонтали
            ws.range((1, col)).api.HorizontalAlignment = -4108

        # устанавливаем текстовый формат для колонок A и B
        ws.range('A:A').api.NumberFormat = '@'
        ws.range('B:B').api.NumberFormat = '@'
        ws.range('D:D').api.NumberFormat = '@'
        ws.range('E:E').api.NumberFormat = '@'
        ws.range('F:F').api.NumberFormat = '@'
        ws.range('G:G').api.NumberFormat = '@'

        # центрируем столбец A (баркод) по горизонтали
        ws.range('A:A').api.HorizontalAlignment = -4108
        ws.range('D:D').api.HorizontalAlignment = -4108
        ws.range('E:E').api.HorizontalAlignment = -4108
        ws.range('F:F').api.HorizontalAlignment = -4108
        ws.range('G:G').api.HorizontalAlignment = -4108

        # включаем перенос текста и уменьшение шрифта для колонки B
        ws.range('B:B').api.WrapText = True
        ws.range('B:B').api.ShrinkToFit = True
        ws.range('D:D').api.WrapText = True

        # заполняем данные и вставляем QR
        for idx, row in enumerate(data, start=2):
            barcode, excise, user_name, computer_name, qr_name, created_date = row[1:]

            # преобразуем в абсолютный путь
            absolute_qr_path = os.path.abspath(qr_path + str(qr_name).strip() + '.png')
            # print(f"Проверяем путь: {absolute_qr_path}")
            # print(f"Файл существует: {os.path.exists(absolute_qr_path)}")

            # записываем данные
            ws.range((idx, 1)).value = str(barcode)
            ws.range((idx, 2)).value = str(excise)
            ws.range((idx, 4)).value = absolute_qr_path
            ws.range((idx, 5)).value = str(computer_name)
            ws.range((idx, 6)).value = str(user_name)
            ws.range((idx, 7)).value = str(created_date).split('.')[0]

            # вставляем готовый QR-код
            if os.path.exists(absolute_qr_path):
                cell = ws.range((idx, 3))

                pic = ws.pictures.add(absolute_qr_path,
                                      left=cell.left + 8,
                                      top=cell.top + 6,
                                      width=70,
                                      height=70)
                ws.range((idx, 3)).row_height = 80
                # print(f"QR вставлен для строки {idx}")
            else:
                ws.range((idx, 3)).value = 'QR не найден'
                # print(f"QR не найден: {absolute_qr_path}")

        # настраиваем ширину колонок
        ws.range('A:A').column_width = 20
        ws.range('B:B').column_width = 50
        ws.range('C:C').column_width = 15
        ws.range('D:D').column_width = 50
        ws.range('E:E').column_width = 15
        ws.range('F:F').column_width = 15
        ws.range('G:G').column_width = 20

        # сохраняем и закрываем
        wb.save(filename)
        wb.close()

        print(f"Отчет сохранен как {filename}")
        return filename

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    app = ReportGenerator()
    app.run()
