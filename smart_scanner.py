import time
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import datetime
import os
import getpass
import socket
import ctypes


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


# внешний вид
ctk.set_appearance_mode('dark')
ctk.set_default_color_theme('dark-blue')

print(is_eng())
to_eng()
print(is_eng())


class ReportGenerator:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title('Scanner')
        self.root.geometry('650x300')
        self.root.resizable(False, False)
        self._notification_timer = None

        self.create_widgets()

        self.root.after(100, self.entry_barcode.focus)

    # def set_focus_to_barcode(self):
    #     self.entry_barcode.focus()

    def create_widgets(self):
        # основной контейнер
        main_container = ctk.CTkFrame(self.root, fg_color='transparent')
        main_container.pack(fill='both', expand=True, padx=30, pady=30)

        # fg_color='#343536'

        # рамка для баркода
        self.barcode_frame = ctk.CTkFrame(main_container, border_width=2, border_color='#808080', corner_radius=10)
        self.barcode_frame.pack(fill='x', pady=(0, 20))

        label_barcode = ctk.CTkLabel(self.barcode_frame, text='Баркод:', font=ctk.CTkFont(size=14, weight='bold'))
        label_barcode.pack(side='left', padx=(20, 10), pady=20)

        self.entry_barcode = ctk.CTkEntry(self.barcode_frame, width=300, height=35)
        self.entry_barcode.pack(side='left', padx=(0, 20), pady=20, fill='x', expand=True)

        # рамка для акциза
        self.excise_frame = ctk.CTkFrame(main_container, border_width=2, border_color='#808080', corner_radius=10)
        self.excise_frame.pack(fill='x', pady=(0, 30))

        label_excise = ctk.CTkLabel(self.excise_frame, text='Акциз:', font=ctk.CTkFont(size=14, weight='bold'))
        label_excise.pack(side='left', padx=(20, 10), pady=20)

        self.entry_excise = ctk.CTkEntry(self.excise_frame, width=300, height=35, state='disabled', fg_color='#202121')
        self.entry_excise.pack(side='left', padx=(0, 20), pady=20, fill='x', expand=True)

        # кнопка формирования отчёта
        self.button_generate = ctk.CTkButton(
            main_container,
            text='Сформировать отчет',
            command=self.generate_report,
            width=250,
            height=45,
            fg_color='#2E8B57',
            text_color='white',
            font=ctk.CTkFont(size=15, weight='bold'),
            hover_color='#3CB371'
        )
        self.button_generate.pack(side='right')

        # метка для уведомлений
        self.notification_label = ctk.CTkLabel(
            main_container,
            text='',
            font=ctk.CTkFont(size=13),
            fg_color='#2E8B57',
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
            self.show_notification(f'Ошибка в отправке данных', label_bg='#800000')
            return False

    def on_barcode_change(self, event=None):
        barcode = self.entry_barcode.get()
        barcode_len = len(barcode)

        # обновление рамки баркода
        if barcode_len == 0:
            self.barcode_frame.configure(border_color='#808080')  # серый

        elif barcode_len == 13 and barcode.isdigit():
            # self.show_notification(f'EAN верный',2000)
            self.barcode_frame.configure(border_color='#2E8B57')  # зеленый
            # активируем поле акциза и ставим фокус
            self.entry_excise.configure(state='normal')
            self.entry_excise.configure(fg_color='#343536')
            self.entry_excise.focus()
            self.excise_frame.configure(border_color='#DAA520')

            # привязываем обработчик для акциза только когда поле активировано
            if not hasattr(self, '_excise_bound'):
                self.entry_excise.bind('<KeyRelease>', self.on_excise_change)
                self._excise_bound = True
        else:
            if barcode.isdigit():
                self.show_notification(f'Неверный EAN (длина {barcode_len} из 13)', label_bg='#800000')
            else:
                self.show_notification(f'В EAN должны быть только цифры', label_bg='#800000')
            self.barcode_frame.configure(border_color='#DC143C')
            # деактивируем поле акциза и очищаем его
            self.entry_excise.delete(0, tk.END)
            self.entry_excise.configure(fg_color='#202121')
            self.entry_excise.configure(state='disabled')
            self.excise_frame.configure(border_color='#808080')

    def show_notification(self, message, duration=3000, label_bg='#2E8B57'):
        if self._notification_timer is not None:
            try:
                self.root.after_cancel(self._notification_timer)
            except:
                pass
            self._notification_timer = None

        self.hide_notification()

        self.notification_label.configure(text=message, fg_color=label_bg)
        self.notification_label.pack(side='left', pady=(10, 0))

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
            self.show_notification(f'Неверная раскладка. Переключите на EN', label_bg='#800000')
            self.entry_excise.delete(0, tk.END)
        else:

            # обновление рамки акциза
            if barcode_len == 0:
                self.excise_frame.configure(border_color='#808080')  # серый
                self.entry_excise.delete(0, tk.END)
                self.entry_excise.configure(state='disabled')
                self.barcode.focus()
            elif excise_len == 0:
                pass
            elif 0 < excise_len <= 10:
                self.excise_frame.configure(border_color='#DC143C')  # красный
                self.show_notification(f'Неверный акциз (длина {excise_len} из 150)', label_bg='#800000')
            else:  # больше n символов
                self.excise_frame.configure(border_color='#2E8B57')  # зеленый

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
                    self.barcode_frame.configure(border_color='#808080')
                    self.excise_frame.configure(border_color='#808080')

                    # устанавливаем курсор на поле баркода
                    self.entry_barcode.focus()

    def generate_report(self):
        barcode = self.entry_barcode.get()
        excise = self.entry_excise.get()

        if not barcode or not excise:
            return

        # отправляем данные через ту же функцию
        self.send_data(barcode, excise)

        # очищаем поля
        self.entry_barcode.delete(0, tk.END)
        self.entry_excise.delete(0, tk.END)
        self.entry_excise.configure(state='disabled')
        self.barcode_frame.configure(border_color='#808080')
        self.excise_frame.configure(border_color='#808080')
        self.entry_barcode.focus()

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    app = ReportGenerator()
    app.run()
