import time
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import datetime
import os
import getpass
import socket
import ctypes

# внешний вид
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


def ensure_english_layout():
    try:
        # получаем текущее активное окно
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        # получаем идентификатор потока окна
        thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
        # получаем текущую раскладку
        layout = ctypes.windll.user32.GetKeyboardLayout(thread_id)
        layout_id = layout & 0xFFFF

        # проверяем, английская ли раскладка
        if (layout_id & 0x00FF) == 0x09 or layout_id == 0x0409:
            return True

        # 0x0050 - WM_INPUTLANGCHANGEREQUEST
        # 0x02 - INPUTLANGCHANGE_SYSCHARSET
        # 0x0409 - английская раскладка (LCID)
        result = ctypes.windll.user32.PostMessageW(hwnd, 0x0050, 0, 0x0409)

        # небольшая задержка для применения
        time.sleep(0.1)

        # проверяем, изменилась ли раскладка
        thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
        layout = ctypes.windll.user32.GetKeyboardLayout(thread_id)
        layout_id = layout & 0xFFFF

        return (layout_id & 0x00FF) == 0x09 or layout_id == 0x0409

    except Exception as e:
        print(f"Ошибка переключения раскладки: {e}")
        return False


class ReportGenerator:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Scanner")
        self.root.geometry("650x300")
        self.root.resizable(False, False)
        self._notification_timer = None
        self._skip_next_excise = False  # Флаг для пропуска обработки акциза

        self.create_widgets()

        # привязываем событие фокуса для проверки раскладки
        self.entry_barcode.bind('<FocusIn>', self.on_focus_barcode)
        self.entry_excise.bind('<FocusIn>', self.on_focus_excise)

        # устанавливаем фокус на поле баркода после запуска
        self.root.after(100, self.set_focus_to_barcode)

    def set_focus_to_barcode(self):
        self.entry_barcode.focus()
        # проверяем раскладку при запуске
        self.on_focus_barcode()

    def create_widgets(self):
        # основной контейнер
        main_container = ctk.CTkFrame(self.root, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=30, pady=30)

        # рамка для баркода
        self.barcode_frame = ctk.CTkFrame(main_container, border_width=2, border_color="#808080", corner_radius=10)
        self.barcode_frame.pack(fill="x", pady=(0, 20))

        label_barcode = ctk.CTkLabel(self.barcode_frame, text="Баркод:", font=ctk.CTkFont(size=14, weight="bold"))
        label_barcode.pack(side="left", padx=(20, 10), pady=20)

        self.entry_barcode = ctk.CTkEntry(self.barcode_frame, width=300, height=35)
        self.entry_barcode.pack(side="left", padx=(0, 20), pady=20, fill="x", expand=True)

        # рамка для акциза
        self.excise_frame = ctk.CTkFrame(main_container, border_width=2, border_color="#808080", corner_radius=10)
        self.excise_frame.pack(fill="x", pady=(0, 30))

        label_excise = ctk.CTkLabel(self.excise_frame, text="Акциз:", font=ctk.CTkFont(size=14, weight="bold"))
        label_excise.pack(side="left", padx=(20, 10), pady=20)

        self.entry_excise = ctk.CTkEntry(self.excise_frame, width=300, height=35, state="disabled")
        self.entry_excise.pack(side="left", padx=(0, 20), pady=20, fill="x", expand=True)

        # кнопка формирования отчёта
        self.button_generate = ctk.CTkButton(
            main_container,
            text="Сформировать отчет",
            command=self.generate_report,
            width=250,
            height=45,
            fg_color="#2E8B57",
            text_color="white",
            font=ctk.CTkFont(size=15, weight="bold"),
            hover_color="#3CB371"
        )
        self.button_generate.pack(side="right")

        # метка для уведомлений
        self.notification_label = ctk.CTkLabel(
            main_container,
            text="",
            font=ctk.CTkFont(size=13),
            fg_color="#2E8B57",
            text_color="white",
            corner_radius=6,
            padx=15,
            pady=8
        )

        # привязываем события ввода
        self.entry_barcode.bind('<KeyRelease>', self.on_barcode_change)

    def on_focus_barcode(self, event=None):
        """Проверяем раскладку при фокусе на поле баркода"""
        if not ensure_english_layout():
            # Если раскладка не английская, очищаем поле и показываем сообщение
            self.entry_barcode.delete(0, tk.END)
            self.show_notification("Ошибка: раскладка должна быть английской! Переключите на EN", 2500, label_bg='#800000')
            self.barcode_frame.configure(border_color="#DC143C")
        else:
            if len(self.entry_barcode.get()) == 0:
                self.barcode_frame.configure(border_color="#808080")

    def on_focus_excise(self, event=None):
        """Проверяем раскладку при фокусе на поле акциза"""
        if not ensure_english_layout():
            # Если раскладка не английская, очищаем поле и показываем сообщение
            self.entry_excise.delete(0, tk.END)
            self.show_notification("Ошибка: раскладка должна быть английской! Переключите на EN", 2500, label_bg='#800000')
            self.excise_frame.configure(border_color="#DC143C")
            # Устанавливаем флаг, чтобы пропустить следующую обработку акциза
            self._skip_next_excise = True
            # Возвращаем фокус на поле баркода
            self.entry_barcode.focus()
        else:
            if len(self.entry_excise.get()) == 0:
                self.excise_frame.configure(border_color="#DAA520")

    def send_data(self, barcode, excise):
        # получаем данные о пользователе и компьютере
        username = getpass.getuser()
        computer_name = socket.gethostname()
        current_time = datetime.datetime.now().strftime('%d-%m-%y %H:%M:%S')

        # выводим все данные
        print(f"\n--- Отправка данных ---")
        print(f"Время: {current_time}")
        print(f"Пользователь: {username}")
        print(f"Компьютер: {computer_name}")
        print(f"Баркод: {barcode}")
        print(f"Акциз: {excise}")
        print(f"----------------------\n")

        return True

    def on_barcode_change(self, event=None):
        barcode = self.entry_barcode.get()
        barcode_len = len(barcode)

        # обновляем цвет рамки баркода
        if barcode_len == 0:
            self.barcode_frame.configure(border_color="#808080")
        elif barcode_len == 13 and barcode.isdigit():
            self.show_notification('EAN верный', 2000)
            self.barcode_frame.configure(border_color="#2E8B57")
            self.entry_excise.configure(state="normal")
            self.entry_excise.focus()
            self.excise_frame.configure(border_color="#DAA520")
            if not hasattr(self, '_excise_bound'):
                self.entry_excise.bind('<KeyRelease>', self.on_excise_change)
                self._excise_bound = True
        else:
            if barcode.isdigit():
                self.show_notification(f'Неверный EAN (длина {barcode_len} из 13)', 1500, label_bg='#800000')
            else:
                self.show_notification('В EAN должны быть только цифры', 1500, label_bg='#800000')
            self.barcode_frame.configure(border_color="#DC143C")
            self.entry_excise.configure(state="disabled")
            self.entry_excise.delete(0, tk.END)
            self.excise_frame.configure(border_color="#808080")

    def show_notification(self, message, duration=2000, label_bg='#2E8B57'):
        if self._notification_timer is not None:
            try:
                self.root.after_cancel(self._notification_timer)
            except:
                pass
            self._notification_timer = None

        self.hide_notification()

        self.notification_label.configure(text=message, fg_color=label_bg)
        self.notification_label.pack(side="left", pady=(10, 0))

        self._notification_timer = self.root.after(duration, self.hide_notification)

    def hide_notification(self):
        self.notification_label.pack_forget()
        self.notification_label.configure(text="")

        if self._notification_timer is not None:
            try:
                self.root.after_cancel(self._notification_timer)
            except:
                pass
            self._notification_timer = None

    def on_excise_change(self, event=None):
        # Если установлен флаг пропуска, пропускаем обработку и сбрасываем флаг
        if self._skip_next_excise:
            self._skip_next_excise = False
            return

        excise = self.entry_excise.get()
        excise_len = len(excise)

        if excise_len == 0:
            pass
        elif excise_len <= 10:
            self.excise_frame.configure(border_color='#DC143C')
            self.show_notification(f'Неверный акциз (длина {excise_len} из 150)', 2000, label_bg='#800000')
        else:
            self.excise_frame.configure(border_color="#2E8B57")
            barcode = self.entry_barcode.get()
            if len(barcode) == 13 and barcode.isdigit():
                self.send_data(barcode, excise)
                self.show_notification("Данные успешно добавлены", 2000)

                self.entry_barcode.delete(0, tk.END)
                self.entry_excise.delete(0, tk.END)
                self.entry_excise.configure(state="disabled")
                self.barcode_frame.configure(border_color="#808080")
                self.excise_frame.configure(border_color="#808080")
                self.entry_barcode.focus()

    def generate_report(self):
        barcode = self.entry_barcode.get()
        excise = self.entry_excise.get()

        if not barcode or not excise:
            messagebox.showwarning("Предупреждение", "Заполните оба поля!")
            return

        self.send_data(barcode, excise)

        self.entry_barcode.delete(0, tk.END)
        self.entry_excise.delete(0, tk.END)
        self.entry_excise.configure(state="disabled")
        self.barcode_frame.configure(border_color="#808080")
        self.excise_frame.configure(border_color="#808080")
        self.entry_barcode.focus()

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    app = ReportGenerator()
    app.run()