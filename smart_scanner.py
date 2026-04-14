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

config = configparser.ConfigParser()
if not os.path.exists('_config.ini'):
    # конфиг по умолчанию
    config['settings'] = {
        'qr_path': '',
        'appearance_mode': 'light'
    }
    with open('_config.ini', 'w', encoding='utf-8') as f:
        config.write(f)

config.read('_config.ini', encoding='utf-8')
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

if appearance_mode == 'light':
    fg_color_disable = '#b8b8b8'
    fg_color_enable = '#ffffff'
else:
    fg_color_disable = '#202121'
    fg_color_enable = '#343536'


# внешний вид
ctk.set_appearance_mode(appearance_mode)
ctk.set_default_color_theme('dark-blue')

print(is_eng())
to_eng()
print(is_eng())


class ReportGenerator:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title('Scanner')

        # self.root.geometry('650x300')
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
        # actual_height = self.root.winfo_height()
        # print(f'actual_width {actual_width}')
        # print(f'actual_height {actual_height}')

        # находим dpi scaler
        scaler = round(actual_width / self.width, 2)
        # print(f'scaler {scaler}')

        # получаем размеры экрана
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        # print(f'screen_width {screen_width}')
        # print(f'screen_height {screen_height}')

        # вычисляем позицию для центрирования
        x = (screen_width - self.width) // 2
        y = (screen_height - self.height) // 2
        # print(f'x {x}')
        # print(f'y {y}')
        # print(f'x*scaler {x*scaler}')
        # print(f'y*scaler {y*scaler}')

        # устанавливаем окончательную позицию
        self.root.geometry(f'+{round(x * scaler)}+{round(y * scaler)}')

        self.create_widgets()

        self.root.after(100, self.entry_barcode.focus)

    def create_widgets(self):
        # основной контейнер
        main_container = ctk.CTkFrame(self.root, fg_color='transparent')
        main_container.pack(fill='both', expand=True, padx=30, pady=30)

        # fg_color=fg_color_enable

        # рамка для баркода
        self.barcode_frame = ctk.CTkFrame(main_container, border_width=2, border_color=border_color_base, corner_radius=10)
        self.barcode_frame.pack(fill='x', pady=(0, 20))

        label_barcode = ctk.CTkLabel(self.barcode_frame, text='Баркод:', font=ctk.CTkFont(size=14, weight='bold'))
        label_barcode.pack(side='left', padx=(20, 10), pady=20)

        self.entry_barcode = ctk.CTkEntry(self.barcode_frame, width=300, height=35)
        self.entry_barcode.pack(side='left', padx=(0, 20), pady=20, fill='x', expand=True)

        # рамка для акциза
        self.excise_frame = ctk.CTkFrame(main_container, border_width=2, border_color=border_color_base, corner_radius=10)
        self.excise_frame.pack(fill='x', pady=(0, 30))

        label_excise = ctk.CTkLabel(self.excise_frame, text='Акциз:', font=ctk.CTkFont(size=14, weight='bold'))
        label_excise.pack(side='left', padx=(20, 10), pady=20)

        self.entry_excise = ctk.CTkEntry(self.excise_frame, width=300, height=35, state='disabled', fg_color=fg_color_disable)
        self.entry_excise.pack(side='left', padx=(0, 20), pady=20, fill='x', expand=True)

        # кнопка формирования отчёта
        self.button_generate = ctk.CTkButton(
            main_container,
            text='Сформировать отчет',
            command=self.generate_report,
            width=250,
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
            pady=8
        )

        # привязываем события ввода
        self.entry_barcode.bind('<KeyRelease>', self.on_barcode_change)

    def send_data(self, barcode, excise):
        try:
            # получаем данные о пользователе и компьютере
            username = getpass.getuser()
            computer_name = socket.gethostname()
            current_time = datetime.datetime.now().strftime('%d-%m-%y %H:%M:%S')

            # выводим все данные
            print(f'\n--- Отправка данных ---')
            print(f'Время: {current_time}')
            print(f'Пользователь: {username}')
            print(f'Компьютер: {computer_name}')
            print(f'Баркод: {barcode}')
            print(f'Акциз: {excise}')
            print(f'----------------------\n')

            return True
        except Exception as e:
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
                self.barcode_frame.configure(border_color=border_color_base)  # серый

            elif barcode_len == 13 and barcode.isdigit():
                self.show_notification(f'EAN: OK', 2000)
                self.barcode_frame.configure(border_color=border_color_green)  # зеленый
                # активируем поле акциза и ставим фокус
                self.entry_excise.configure(state='normal')
                self.entry_excise.configure(fg_color=fg_color_enable) # normal
                self.entry_excise.focus()
                self.excise_frame.configure(border_color=border_color_yellow) # желтый

                # привязываем обработчик для акциза только когда поле активировано
                if not hasattr(self, '_excise_bound'):
                    self.entry_excise.bind('<KeyRelease>', self.on_excise_change)
                    self._excise_bound = True
            else:
                if barcode.isdigit():
                    self.show_notification(f'Неверный EAN (длина {barcode_len} из 13)', label_bg=notification_color_red)
                else:
                    self.show_notification(f'В EAN должны быть только цифры', label_bg=notification_color_red)
                self.barcode_frame.configure(border_color=border_color_red) # красный
                # деактивируем поле акциза и очищаем его
                self.entry_excise.delete(0, tk.END)
                self.entry_excise.configure(fg_color=fg_color_disable)  # темный (черный)
                self.entry_excise.configure(state='disabled')
                self.excise_frame.configure(border_color=border_color_base) # серый

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
                self.excise_frame.configure(border_color=border_color_base)  # серый
                self.entry_excise.delete(0, tk.END)
                self.entry_excise.configure(state='disabled')
                # self.barcode.focus()
                self.entry_barcode.focus()
            elif excise_len == 0:
                self.excise_frame.configure(border_color=border_color_yellow)
            elif excise_len == 13:
                self.entry_barcode.delete(0, tk.END)
                self.entry_barcode.insert(0, excise)
                self.entry_excise.delete(0, tk.END)
                self.excise_frame.configure(border_color=border_color_yellow)
                self.show_notification(f'Баркод обновлён', label_bg=notification_color_yellow)
            elif 0 < excise_len <= 20:
                self.excise_frame.configure(border_color=border_color_red)  # красный
                self.show_notification(f'Неверный акциз (длина {excise_len} из 150)', label_bg=notification_color_red)
            else:  # больше n символов
                self.excise_frame.configure(border_color=border_color_green)  # зеленый

                # фильтр баркода
                if barcode_len == 13 and barcode.isdigit():
                    # отправляем данные
                    self.send_data(barcode, excise)

                    # показываем уведомление снизу
                    self.show_notification(f'Данные успешно отправлены')

                    # очищаем оба поля
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
        barcode = self.entry_barcode.get()
        excise = self.entry_excise.get()

        if not barcode or not excise:
            return

        # отправляем данные
        self.send_data(barcode, excise)

        # очищаем поля
        self.entry_barcode.delete(0, tk.END)
        self.entry_excise.delete(0, tk.END)
        self.entry_excise.configure(state='disabled')
        self.barcode_frame.configure(border_color=border_color_base)
        self.excise_frame.configure(border_color=border_color_base)
        self.entry_barcode.focus()

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    app = ReportGenerator()
    app.run()
