# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                    90+ ПРОСРОЧКА АНАЛИЗАТОР v6.0                            ║
║                    Профессиональный банковский Dashboard                    ║
║                    Полная версия с прокруткой графиков                      ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import pandas as pd
import numpy as np
import os
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

import plotly.graph_objects as go
from plotly.subplots import make_subplots

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# =============================================================================
# КОНСТАНТЫ
# =============================================================================

THRESHOLD = 90

COLORS = {
    'primary': '#0D47A1',
    'secondary': '#1565C0',
    'on_date': '#1976D2',
    'entered': '#C62828',
    'exited': '#2E7D32',
    'insurance': '#E65100',
    'other': '#6A1B9A',
    'positive': '#D32F2F',
    'negative': '#388E3C',
    'neutral': '#455A64',
    'background': '#FAFAFA'
}

MONTH_NAMES_RU = {
    'jan': 'Январь', 'feb': 'Февраль', 'mar': 'Март', 'apr': 'Апрель',
    'may': 'Май', 'jun': 'Июнь', 'jul': 'Июль', 'aug': 'Август',
    'sep': 'Сентябрь', 'oct': 'Октябрь', 'nov': 'Ноябрь', 'dec': 'Декабрь'
}

MONTH_ORDER = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
}


# =============================================================================
# ФОРМАТИРОВАНИЕ
# =============================================================================

def format_number(value, decimals=0):
    """Форматирование числа: пробел - тысячи, запятая - дробная часть"""
    if pd.isna(value) or value == 0:
        return "0"
    
    if decimals > 0:
        formatted = f"{abs(value):,.{decimals}f}"
        parts = formatted.split('.')
        integer_part = parts[0].replace(',', ' ')
        decimal_part = parts[1] if len(parts) > 1 else '00'
        result = f"{integer_part},{decimal_part}"
        return f"-{result}" if value < 0 else result
    else:
        formatted = f"{abs(int(value)):,}"
        result = formatted.replace(',', ' ')
        return f"-{result}" if value < 0 else result


# =============================================================================
# GUI - ГЛАВНОЕ ОКНО
# =============================================================================

class MainApplication:
    """Главное окно приложения"""
    
    def __init__(self):
        self.data_files = {}
        self.insurance_file = None
        self.output_path = None
        self.analysis_mode = None
        self.should_run = False
        
    def run(self):
        """Запуск главного окна"""
        self.root = tk.Tk()
        self.root.title("🏦 Анализатор просрочки 90+ | Версия 6.0")
        self.root.geometry("1000x850")
        self.root.configure(bg='#F5F7FA')
        self.root.resizable(True, True)
        
        # Центрирование окна
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - 500
        y = (self.root.winfo_screenheight() // 2) - 425
        self.root.geometry(f'1000x850+{x}+{y}')
        
        self._setup_styles()
        self._create_widgets()
        
        self.root.mainloop()
        return self.should_run
    
    def _setup_styles(self):
        """Настройка стилей"""
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 22, 'bold'), 
                       background='#F5F7FA', 
                       foreground='#0D47A1')
        
        style.configure('Subtitle.TLabel', 
                       font=('Segoe UI', 11), 
                       background='#F5F7FA', 
                       foreground='#5D6D7E')
        
        style.configure('Header.TLabel', 
                       font=('Segoe UI', 11, 'bold'), 
                       background='#F5F7FA', 
                       foreground='#1A5276')
        
        style.configure('Info.TLabel', 
                       font=('Segoe UI', 10), 
                       background='#F5F7FA', 
                       foreground='#626567')
        
        style.configure('TLabelframe', 
                       background='#F5F7FA',
                       borderwidth=2,
                       relief='groove')
        
        style.configure('TLabelframe.Label', 
                       font=('Segoe UI', 11, 'bold'),
                       background='#F5F7FA', 
                       foreground='#0D47A1')
        
        style.configure('TButton',
                       font=('Segoe UI', 10),
                       padding=8)
        
        style.configure('TRadiobutton',
                       font=('Segoe UI', 10),
                       background='#F5F7FA')
    
    def _create_widgets(self):
        """Создание виджетов"""
        # Основной контейнер с прокруткой
        canvas = tk.Canvas(self.root, bg='#F5F7FA', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Привязка прокрутки мышью
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        main_frame = ttk.Frame(scrollable_frame, padding="30")
        main_frame.pack(fill='both', expand=True)
        
        # ═══════════════════════════════════════════════════════════════
        # ЗАГОЛОВОК
        # ═══════════════════════════════════════════════════════════════
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill='x', pady=(0, 25))
        
        ttk.Label(title_frame, 
                 text="🏦 Анализатор просрочки свыше 90 дней", 
                 style='Title.TLabel').pack()
        
        ttk.Label(title_frame, 
                 text="Профессиональный инструмент для анализа кредитного портфеля банка", 
                 style='Subtitle.TLabel').pack(pady=(8, 0))
        
        ttk.Label(title_frame, 
                 text="Версия 6.0 | Поддержка нескольких лет | Интерактивные графики с прокруткой", 
                 style='Info.TLabel').pack(pady=(4, 0))
        
        # ═══════════════════════════════════════════════════════════════
        # ФАЙЛЫ ДАННЫХ
        # ═══════════════════════════════════════════════════════════════
        data_frame = ttk.LabelFrame(main_frame, 
                                   text=" 📁 Шаг 1: Загрузка файлов данных ", 
                                   padding="15")
        data_frame.pack(fill='x', pady=12)
        
        ttk.Label(data_frame, 
                 text="Добавьте Excel-файлы с данными за каждый год (например: 2024.xlsx, 2025.xlsx)",
                 style='Info.TLabel').pack(anchor='w', pady=(0, 10))
        
        list_frame = ttk.Frame(data_frame)
        list_frame.pack(fill='x')
        
        self.files_listbox = tk.Listbox(
            list_frame, 
            height=5, 
            font=('Consolas', 11),
            selectmode=tk.SINGLE, 
            bg='white',
            relief='solid', 
            borderwidth=1,
            selectbackground='#0D47A1',
            selectforeground='white'
        )
        self.files_listbox.pack(side='left', fill='x', expand=True)
        
        scrollbar_list = ttk.Scrollbar(list_frame, orient='vertical', 
                                       command=self.files_listbox.yview)
        scrollbar_list.pack(side='right', fill='y')
        self.files_listbox.config(yscrollcommand=scrollbar_list.set)
        
        btn_frame = ttk.Frame(data_frame)
        btn_frame.pack(fill='x', pady=(12, 0))
        
        ttk.Button(btn_frame, text="➕ Добавить файл", 
                  command=self._add_data_file).pack(side='left', padx=4)
        ttk.Button(btn_frame, text="➖ Удалить выбранный", 
                  command=self._remove_data_file).pack(side='left', padx=4)
        ttk.Button(btn_frame, text="🗑️ Очистить список", 
                  command=self._clear_data_files).pack(side='left', padx=4)
        
        # ═══════════════════════════════════════════════════════════════
        # СТРАХОВКА
        # ═══════════════════════════════════════════════════════════════
        ins_frame = ttk.LabelFrame(main_frame, 
                                  text=" 🛡️ Шаг 2: Страховые погашения (опционально) ", 
                                  padding="15")
        ins_frame.pack(fill='x', pady=12)
        
        ttk.Label(ins_frame, 
                 text="Файл содержит список анкет, погашенных за счёт страхового возмещения",
                 style='Info.TLabel').pack(anchor='w', pady=(0, 10))
        
        ins_inner = ttk.Frame(ins_frame)
        ins_inner.pack(fill='x')
        
        self.insurance_var = tk.StringVar(value="Файл не выбран")
        ttk.Entry(ins_inner, textvariable=self.insurance_var, 
                 state='readonly', width=55, font=('Segoe UI', 10)).pack(side='left', fill='x', expand=True)
        ttk.Button(ins_inner, text="📂 Выбрать файл", 
                  command=self._select_insurance).pack(side='left', padx=(10, 4))
        ttk.Button(ins_inner, text="📋 Создать шаблон", 
                  command=self._create_template).pack(side='left', padx=4)
        
        # Формат файла
        format_frame = ttk.Frame(ins_frame)
        format_frame.pack(fill='x', pady=(12, 0))
        
        format_text = """📌 Формат файла страховки:
    • Колонка "dealid" — номер кредитной анкеты
    • Колонка "period" — период погашения (формат: 2024-01, 2024-02 и т.д.)
    • Суммы рассчитываются автоматически из основных данных
    • Дубликаты одной анкеты в одном периоде учитываются один раз"""
        
        ttk.Label(format_frame, text=format_text, style='Info.TLabel', 
                 justify='left').pack(anchor='w')
        
        # ═══════════════════════════════════════════════════════════════
        # ПАПКА РЕЗУЛЬТАТОВ
        # ═══════════════════════════════════════════════════════════════
        out_frame = ttk.LabelFrame(main_frame, 
                                  text=" 📂 Шаг 3: Папка для сохранения результатов ", 
                                  padding="15")
        out_frame.pack(fill='x', pady=12)
        
        out_inner = ttk.Frame(out_frame)
        out_inner.pack(fill='x')
        
        self.output_var = tk.StringVar(value="Папка не выбрана")
        ttk.Entry(out_inner, textvariable=self.output_var, 
                 state='readonly', width=65, font=('Segoe UI', 10)).pack(side='left', fill='x', expand=True)
        ttk.Button(out_inner, text="📂 Выбрать папку", 
                  command=self._select_output).pack(side='left', padx=(10, 0))
        
        # ═══════════════════════════════════════════════════════════════
        # РЕЖИМ АНАЛИЗА
        # ═══════════════════════════════════════════════════════════════
        mode_frame = ttk.LabelFrame(main_frame, 
                                   text=" 📊 Шаг 4: Выберите режим формирования отчёта ", 
                                   padding="15")
        mode_frame.pack(fill='x', pady=12)
        
        self.mode_var = tk.StringVar(value="combined")
        
        modes = [
            ("separate", "📄 Раздельные отчёты", 
             "Создать отдельный отчёт для каждого года (2024, 2025, 2026...)"),
            ("combined", "📋 Объединённый отчёт", 
             "Все данные в одном отчёте с общей таблицей и графиками за весь период"),
            ("both", "📄📋 Оба варианта", 
             "Создать и раздельные отчёты по годам, и общий объединённый отчёт")
        ]
        
        for value, title, description in modes:
            mode_item = ttk.Frame(mode_frame)
            mode_item.pack(fill='x', pady=6)
            
            rb = ttk.Radiobutton(mode_item, text=title, value=value, 
                                variable=self.mode_var, style='TRadiobutton')
            rb.pack(side='left')
            
            ttk.Label(mode_item, text=f"  —  {description}", 
                     style='Info.TLabel').pack(side='left', padx=(5, 0))
        
        # ═══════════════════════════════════════════════════════════════
        # СПРАВКА ПО МЕТОДОЛОГИИ
        # ═══════════════════════════════════════════════════════════════
        info_frame = ttk.LabelFrame(main_frame, 
                                   text=" ℹ️ Справка: Методология расчёта показателей ", 
                                   padding="15")
        info_frame.pack(fill='x', pady=12)
        
        methodology_text = """
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│  ПОКАЗАТЕЛЬ                │  ФОРМУЛА РАСЧЁТА                │  ОПИСАНИЕ               │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│  На отчётную дату          │  start_days > 90                │  Анкеты с просрочкой    │
│                            │                                 │  >90 дней на начало     │
│                            │                                 │  месяца                 │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│  Вошли в 90+               │  start_days ≤ 90 И              │  Анкеты, перешедшие     │
│                            │  max_days > 90                  │  порог 90 дней в        │
│                            │                                 │  течение месяца         │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│  Вышли из 90+              │  max_days > 90 И                │  Анкеты, полностью      │
│                            │  end_days = 0                   │  погашенные к концу     │
│                            │                                 │  месяца                 │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│  Из них страховка          │  Пересечение "Вышли из 90+"     │  Часть погашений за     │
│                            │  с файлом страховки             │  счёт страховки         │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│  Без страховки (прочие)    │  Вышли из 90+ − Страховка       │  Прочие источники       │
│                            │                                 │  погашения              │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│  Баланс за месяц           │  Вошли − Вышли                  │  Чистое изменение       │
│                            │                                 │  за месяц               │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│  Накопленный баланс        │  Σ(Баланс за месяц)             │  Нарастающий итог       │
│                            │  нарастающим итогом             │  с начала периода       │
└─────────────────────────────────────────────────────────────────────────────────────────┘

📈 Положительный накопленный баланс — портфель 90+ растёт (вошло больше, чем вышло)
📉 Отрицательный накопленный баланс — портфель 90+ сокращается (вышло больше, чем вошло)
"""
        
        info_text = tk.Text(info_frame, height=22, font=('Consolas', 9), 
                           bg='#FAFAFA', relief='flat', wrap='none')
        info_text.insert('1.0', methodology_text)
        info_text.config(state='disabled')
        info_text.pack(fill='x')
        
        # ═══════════════════════════════════════════════════════════════
        # КНОПКИ ДЕЙСТВИЙ
        # ═══════════════════════════════════════════════════════════════
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill='x', pady=25)
        
        # Кнопка "Сформировать отчёт"
        self.run_button = tk.Button(
            action_frame, 
            text="✅  СФОРМИРОВАТЬ ОТЧЁТ", 
            font=('Segoe UI', 14, 'bold'), 
            bg='#0D47A1', 
            fg='white',
            activebackground='#1565C0',
            activeforeground='white',
            padx=40, 
            pady=14, 
            cursor='hand2', 
            relief='flat',
            command=self._on_run
        )
        self.run_button.pack(side='left', padx=10)
        
        # Кнопка "Отмена"
        cancel_button = tk.Button(
            action_frame, 
            text="❌  Отмена", 
            font=('Segoe UI', 12), 
            bg='#757575', 
            fg='white',
            activebackground='#9E9E9E',
            activeforeground='white',
            padx=30, 
            pady=14, 
            cursor='hand2', 
            relief='flat',
            command=self._on_cancel
        )
        cancel_button.pack(side='left', padx=10)
        
        # Статус
        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(action_frame, textvariable=self.status_var, 
                                     style='Info.TLabel')
        self.status_label.pack(side='left', padx=20)
    
    def _add_data_file(self):
        """Добавление файла данных"""
        filepath = filedialog.askopenfilename(
            title="Выберите файл данных",
            filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if filepath:
            filename = os.path.basename(filepath)
            year_match = re.search(r'(20\d{2})', filename)
            
            if year_match:
                year = year_match.group(1)
            else:
                year = self._ask_year()
                if not year:
                    return
            
            if year in self.data_files:
                messagebox.showwarning("Предупреждение", 
                                      f"Файл за {year} год уже добавлен.\nОн будет заменён.")
                # Удаляем старую запись из listbox
                for i in range(self.files_listbox.size()):
                    if f"{year} год" in self.files_listbox.get(i):
                        self.files_listbox.delete(i)
                        break
            
            self.data_files[year] = filepath
            self.files_listbox.insert(tk.END, f"  📄 {year} год  →  {filename}")
            self.status_var.set(f"✅ Добавлен файл за {year} год")
    
    def _ask_year(self):
        """Диалог ввода года"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Укажите год данных")
        dialog.geometry("350x150")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#F5F7FA')
        
        # Центрирование
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 175
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 75
        dialog.geometry(f'+{x}+{y}')
        
        ttk.Label(dialog, text="Год не определён автоматически.\nВведите год данных:", 
                 style='Info.TLabel').pack(pady=15)
        
        year_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=year_var, width=15, font=('Segoe UI', 12))
        entry.pack()
        entry.focus()
        
        result = [None]
        
        def on_ok(event=None):
            val = year_var.get().strip()
            if val.isdigit() and len(val) == 4 and 2000 <= int(val) <= 2100:
                result[0] = val
                dialog.destroy()
            else:
                messagebox.showerror("Ошибка", "Введите корректный год (например: 2024)")
        
        entry.bind('<Return>', on_ok)
        ttk.Button(dialog, text="OK", command=on_ok).pack(pady=15)
        
        dialog.wait_window()
        return result[0]
    
    def _remove_data_file(self):
        """Удаление выбранного файла"""
        selection = self.files_listbox.curselection()
        if selection:
            idx = selection[0]
            text = self.files_listbox.get(idx)
            year_match = re.search(r'(20\d{2})', text)
            if year_match:
                year = year_match.group(1)
                if year in self.data_files:
                    del self.data_files[year]
            self.files_listbox.delete(idx)
            self.status_var.set("🗑️ Файл удалён из списка")
    
    def _clear_data_files(self):
        """Очистка всех файлов"""
        self.files_listbox.delete(0, tk.END)
        self.data_files.clear()
        self.status_var.set("🗑️ Список файлов очищен")
    
    def _select_insurance(self):
        """Выбор файла страховки"""
        filepath = filedialog.askopenfilename(
            title="Выберите файл страховых погашений",
            filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if filepath:
            self.insurance_file = filepath
            self.insurance_var.set(os.path.basename(filepath))
            self.status_var.set("✅ Файл страховки выбран")
    
    def _create_template(self):
        """Создание шаблона файла страховки"""
        folder = filedialog.askdirectory(title="Выберите папку для сохранения шаблона")
        if folder:
            template_df = pd.DataFrame({
                'dealid': [12345678, 23456789, 34567890, 45678901, 56789012],
                'period': ['2024-01', '2024-01', '2024-02', '2024-03', '2025-01']
            })
            filepath = os.path.join(folder, "ШАБЛОН_Страховка.xlsx")
            template_df.to_excel(filepath, index=False)
            messagebox.showinfo("Шаблон создан", 
                               f"Файл шаблона сохранён:\n{filepath}\n\n"
                               "Заполните колонки dealid и period,\n"
                               "затем выберите этот файл в программе.")
            self.status_var.set("📋 Шаблон страховки создан")
    
    def _select_output(self):
        """Выбор папки для результатов"""
        folder = filedialog.askdirectory(title="Выберите папку для сохранения отчётов")
        if folder:
            self.output_path = folder
            self.output_var.set(folder)
            self.status_var.set("✅ Папка для результатов выбрана")
    
    def _on_run(self):
        """Запуск формирования отчёта"""
        # Проверки
        if not self.data_files:
            messagebox.showerror("Ошибка", 
                                "Не добавлены файлы данных!\n\n"
                                "Добавьте хотя бы один Excel-файл с данными.")
            return
        
        if not self.output_path:
            messagebox.showerror("Ошибка", 
                                "Не выбрана папка для результатов!\n\n"
                                "Укажите папку, куда будут сохранены отчёты.")
            return
        
        # Подтверждение
        years = sorted(self.data_files.keys())
        years_str = ", ".join(years)
        mode_names = {
            'separate': 'Раздельные отчёты',
            'combined': 'Объединённый отчёт',
            'both': 'Оба варианта'
        }
        
        confirm_msg = f"""Параметры анализа:

📅 Годы данных: {years_str}
📊 Режим: {mode_names[self.mode_var.get()]}
🛡️ Страховка: {'Да' if self.insurance_file else 'Нет'}
📂 Папка: {self.output_path}

Начать формирование отчёта?"""
        
        if not messagebox.askyesno("Подтверждение", confirm_msg):
            return
        
        self.analysis_mode = self.mode_var.get()
        self.should_run = True
        self.root.quit()
        self.root.destroy()
    
    def _on_cancel(self):
        """Отмена"""
        self.root.quit()
        self.root.destroy()


# =============================================================================
# АНАЛИЗАТОР ДАННЫХ
# =============================================================================

class DataAnalyzer:
    """Анализатор данных просрочки 90+"""
    
    def __init__(self, data_files: dict, insurance_file: str = None):
        self.data_files = data_files
        self.insurance_file = insurance_file
        self.dataframes = {}
        self.insurance_by_period = {}
        self.results_by_year = {}
        self.combined_results = []
    
    def load_all_data(self):
        """Загрузка всех данных"""
        print("\n" + "═"*70)
        print("  📥 ЗАГРУЗКА ДАННЫХ")
        print("═"*70)
        
        for year, filepath in sorted(self.data_files.items()):
            print(f"\n  📄 {year} год: {os.path.basename(filepath)}")
            df = pd.read_excel(filepath)
            df.columns = [str(col).lower().strip() for col in df.columns]
            self.dataframes[year] = df
            print(f"     ✅ Загружено записей: {format_number(len(df))}")
        
        if self.insurance_file:
            self._load_insurance()
        
        print("\n  " + "─"*66)
        print(f"  📊 Всего загружено {len(self.dataframes)} файл(ов) данных")
    
    def _load_insurance(self):
        """Загрузка данных страховки"""
        print(f"\n  🛡️ Страховка: {os.path.basename(self.insurance_file)}")
        
        try:
            df_ins = pd.read_excel(self.insurance_file)
            df_ins.columns = [str(col).lower().strip() for col in df_ins.columns]
            
            # Поиск колонок
            dealid_col = None
            period_col = None
            
            for col in df_ins.columns:
                col_lower = col.lower()
                if dealid_col is None and any(x in col_lower for x in ['dealid', 'deal_id', 'анкета', 'id']):
                    dealid_col = col
                if period_col is None and any(x in col_lower for x in ['period', 'период', 'дата', 'месяц']):
                    period_col = col
            
            if dealid_col is None:
                dealid_col = df_ins.columns[0]
            if period_col is None and len(df_ins.columns) >= 2:
                period_col = df_ins.columns[1]
            
            df_ins = df_ins.rename(columns={dealid_col: 'dealid', period_col: 'period'})
            df_ins = df_ins[['dealid', 'period']].dropna()
            df_ins['dealid'] = pd.to_numeric(df_ins['dealid'], errors='coerce')
            df_ins = df_ins.dropna(subset=['dealid'])
            df_ins['dealid'] = df_ins['dealid'].astype(int)
            
            df_ins['period_parsed'] = df_ins['period'].apply(self._parse_period)
            df_ins = df_ins.dropna(subset=['period_parsed'])
            df_ins = df_ins.drop_duplicates(subset=['dealid', 'period_parsed'])
            
            for period in df_ins['period_parsed'].unique():
                mask = df_ins['period_parsed'] == period
                self.insurance_by_period[period] = set(df_ins.loc[mask, 'dealid'].tolist())
            
            print(f"     ✅ Загружено уникальных записей: {format_number(len(df_ins))}")
            print(f"     📅 Периоды: {', '.join(sorted(self.insurance_by_period.keys()))}")
            
        except Exception as e:
            print(f"     ⚠️ Ошибка загрузки: {str(e)}")
    
    def _parse_period(self, period_str):
        """Парсинг периода в формат YYYY-MM"""
        if pd.isna(period_str):
            return None
        period_str = str(period_str).strip()
        
        # Формат: 2024-01 или 2024/01
        match = re.match(r'(\d{4})[-/\.](\d{1,2})', period_str)
        if match:
            return f"{match.group(1)}-{int(match.group(2)):02d}"
        
        # Формат: 01.2024 или 01/2024
        match = re.match(r'(\d{1,2})[-./](\d{4})', period_str)
        if match:
            return f"{match.group(2)}-{int(match.group(1)):02d}"
        
        return None
    
    def _detect_months_in_df(self, df, base_year):
        """Определение доступных месяцев в DataFrame"""
        pattern = re.compile(r'^([a-z]{3})(\d{2})_start_days$')
        months = []
        
        for col in df.columns:
            match = pattern.match(str(col))
            if match:
                month_code = match.group(1)
                year_suffix = match.group(2)
                year_full = 2000 + int(year_suffix)
                
                # Включаем месяцы нужного года и январь следующего
                if str(year_full) == base_year or \
                   (str(year_full) == str(int(base_year) + 1) and month_code == 'jan'):
                    
                    month_num = MONTH_ORDER.get(month_code, 0)
                    month_name = MONTH_NAMES_RU.get(month_code, month_code)
                    prefix = f"{month_code}{year_suffix}"
                    
                    months.append({
                        'prefix': prefix,
                        'month_code': month_code,
                        'year': year_full,
                        'month_num': month_num,
                        'name_ru': f"{month_name} {year_full}",
                        'short_name': f"{month_name[:3]}'{str(year_full)[2:]}",
                        'period_key': f"{year_full}-{month_num:02d}",
                        'sort_key': year_full * 100 + month_num
                    })
        
        months.sort(key=lambda x: x['sort_key'])
        return months
    
    def _analyze_single_month(self, df, month_info: dict) -> dict:
        """Анализ одного месяца"""
        prefix = month_info['prefix']
        period_key = month_info['period_key']
        
        # Названия колонок
        start_days = f'{prefix}_start_days'
        max_days = f'{prefix}_max_days'
        end_days = f'{prefix}_end_days'
        start_rest = f'{prefix}_start_rest'
        
        # Поиск колонки max_rest
        max_rest_col = None
        for col_name in [f'{prefix}_max_rest_ref', f'{prefix}_max_rest']:
            if col_name in df.columns:
                max_rest_col = col_name
                break
        
        # Проверка обязательных колонок
        required = [start_days, max_days, end_days]
        for col in required:
            if col not in df.columns:
                return None
        
        data = df.copy()
        
        # Преобразование типов и заполнение пропусков
        for col in [start_days, max_days, end_days, start_rest]:
            if col in data.columns:
                data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0)
        
        if max_rest_col and max_rest_col in data.columns:
            data[max_rest_col] = pd.to_numeric(data[max_rest_col], errors='coerce').fillna(0)
        
        result = {
            'period': month_info['name_ru'],
            'short_period': month_info['short_name'],
            'prefix': prefix,
            'year': month_info['year'],
            'month_num': month_info['month_num'],
            'period_key': period_key,
            'sort_key': month_info['sort_key']
        }
        
        # ═══════════════════════════════════════════════════════════════
        # 1. НА ОТЧЁТНУЮ ДАТУ (начало месяца)
        # ═══════════════════════════════════════════════════════════════
        mask_on_date = data[start_days] > THRESHOLD
        result['on_date_count'] = int(mask_on_date.sum())
        result['on_date_sum'] = float(data.loc[mask_on_date, start_rest].sum()) if start_rest in data.columns else 0.0
        
        # ═══════════════════════════════════════════════════════════════
        # 2. ВОШЛИ В 90+
        # ═══════════════════════════════════════════════════════════════
        mask_entered = (data[start_days] <= THRESHOLD) & (data[max_days] > THRESHOLD)
        result['entered_count'] = int(mask_entered.sum())
        result['entered_sum'] = float(data.loc[mask_entered, max_rest_col].sum()) if max_rest_col else 0.0
        
        # ═══════════════════════════════════════════════════════════════
        # 3. ВЫШЛИ ИЗ 90+
        # ═══════════════════════════════════════════════════════════════
        mask_exited = (data[max_days] > THRESHOLD) & (data[end_days] == 0)
        exited_df = data[mask_exited].copy()
        exited_dealids = set(exited_df['dealid'].tolist())
        
        result['exited_count'] = int(mask_exited.sum())
        result['exited_sum'] = float(exited_df[max_rest_col].sum()) if max_rest_col else 0.0
        
        # ═══════════════════════════════════════════════════════════════
        # 4. ИЗ НИХ СТРАХОВКА
        # ═══════════════════════════════════════════════════════════════
        result['insurance_count'] = 0
        result['insurance_sum'] = 0.0
        
        if period_key in self.insurance_by_period:
            insurance_dealids = self.insurance_by_period[period_key]
            insurance_in_exited = exited_dealids.intersection(insurance_dealids)
            result['insurance_count'] = len(insurance_in_exited)
            
            if insurance_in_exited and max_rest_col:
                mask_ins = data['dealid'].isin(insurance_in_exited) & mask_exited
                result['insurance_sum'] = float(data.loc[mask_ins, max_rest_col].sum())
        
        # ═══════════════════════════════════════════════════════════════
        # 5. БЕЗ СТРАХОВКИ (ПРОЧИЕ)
        # ═══════════════════════════════════════════════════════════════
        result['other_count'] = result['exited_count'] - result['insurance_count']
        result['other_sum'] = result['exited_sum'] - result['insurance_sum']
        
        # ═══════════════════════════════════════════════════════════════
        # 6. БАЛАНС ЗА МЕСЯЦ
        # ═══════════════════════════════════════════════════════════════
        result['monthly_balance'] = result['entered_count'] - result['exited_count']
        result['monthly_balance_sum'] = result['entered_sum'] - result['exited_sum']
        
        return result
    
    def analyze_year(self, year: str):
        """Анализ одного года"""
        if year not in self.dataframes:
            return []
        
        df = self.dataframes[year]
        months = self._detect_months_in_df(df, year)
        results = []
        
        print(f"\n  📅 Анализ {year} года ({len(months)} месяцев):")
        
        for month_info in months:
            result = self._analyze_single_month(df, month_info)
            if result:
                results.append(result)
                print(f"     ✅ {result['period']}: "
                      f"на дату={format_number(result['on_date_count'])}, "
                      f"вошли={format_number(result['entered_count'])}, "
                      f"вышли={format_number(result['exited_count'])}")
        
        # Расчёт накопленного баланса для года
        cumulative = 0
        for r in results:
            cumulative += r['monthly_balance']
            r['cumulative_balance'] = cumulative
        
        return results
    
    def analyze_all(self):
        """Полный анализ всех данных"""
        print("\n" + "═"*70)
        print("  📊 АНАЛИЗ ДАННЫХ")
        print("═"*70)
        
        # Анализ по годам
        for year in sorted(self.dataframes.keys()):
            results = self.analyze_year(year)
            self.results_by_year[year] = results
            self.combined_results.extend(results)
        
        # Сортировка объединённых результатов по дате
        self.combined_results.sort(key=lambda x: x['sort_key'])
        
        # Пересчёт накопленного баланса для объединённых данных
        cumulative = 0
        for r in self.combined_results:
            cumulative += r['monthly_balance']
            r['cumulative_balance'] = cumulative
        
        print("\n  " + "─"*66)
        print(f"  📊 Всего проанализировано {len(self.combined_results)} месяцев")
        
        return self.combined_results


# =============================================================================
# ГЕНЕРАТОР ОТЧЁТОВ
# =============================================================================

class ReportGenerator:
    """Генератор HTML и Excel отчётов"""
    
    def __init__(self, analyzer: DataAnalyzer, output_path: str):
        self.analyzer = analyzer
        self.output_path = output_path
    
    def generate_separate_reports(self):
        """Генерация раздельных отчётов по годам"""
        print("\n  📄 Создание раздельных отчётов по годам...")
        
        paths = []
        for year, results in sorted(self.analyzer.results_by_year.items()):
            if results:
                df = pd.DataFrame(results)
                html_path = self._create_html_report(df, f"Отчёт_{year}", f"Отчёт за {year} год")
                self._create_excel_report(df, f"Отчёт_{year}")
                paths.append(html_path)
        
        return paths
    
    def generate_combined_report(self):
        """Генерация объединённого отчёта"""
        print("\n  📋 Создание объединённого отчёта...")
        
        df = pd.DataFrame(self.analyzer.combined_results)
        
        years = sorted(set(r['year'] for r in self.analyzer.combined_results))
        years_str = "-".join(str(y) for y in years)
        
        html_path = self._create_html_report(df, f"Отчёт_Объединённый_{years_str}", 
                                            f"Объединённый отчёт за {years_str} годы")
        self._create_excel_report(df, f"Отчёт_Объединённый_{years_str}")
        
        return html_path
    
    def _create_html_report(self, df, filename, title):
        """Создание HTML отчёта"""
        timestamp = datetime.now().strftime("%d.%m.%Y %H:%M")
        
        if len(df) == 0:
            return None
        
        # ═══════════════════════════════════════════════════════════════
        # РАСЧЁТ ИТОГОВЫХ ПОКАЗАТЕЛЕЙ
        # ═══════════════════════════════════════════════════════════════
        
        # На начало периода (только первое значение!)
        total_on_date_start = int(df['on_date_count'].iloc[0])
        total_on_date_sum_start = df['on_date_sum'].iloc[0] / 1e6
        
        # Суммы за весь период
        total_entered = int(df['entered_count'].sum())
        total_entered_sum = df['entered_sum'].sum() / 1e6
        
        total_exited = int(df['exited_count'].sum())
        total_exited_sum = df['exited_sum'].sum() / 1e6
        
        total_insurance = int(df['insurance_count'].sum())
        total_insurance_sum = df['insurance_sum'].sum() / 1e6
        
        total_other = int(df['other_count'].sum())
        total_other_sum = df['other_sum'].sum() / 1e6
        
        # Итоговый баланс (последнее значение накопленного)
        final_balance = int(df['cumulative_balance'].iloc[-1])
        
        # Период отчёта
        period_start = df['period'].iloc[0]
        period_end = df['period'].iloc[-1]
        num_months = len(df)
        
        # Расчёт ширины графиков (минимум 50px на столбец, но не менее 100%)
        chart_width = max(100, num_months * 70)
        
        # ═══════════════════════════════════════════════════════════════
        # СОЗДАНИЕ ГРАФИКОВ
        # ═══════════════════════════════════════════════════════════════
        
        chart1_json = self._create_count_chart(df).to_json()
        chart2_json = self._create_sum_chart(df).to_json()
        chart3_json = self._create_waterfall_chart(df).to_json()
        chart4_json = self._create_pie_chart(df).to_json()
        chart5_json = self._create_exit_breakdown_chart(df).to_json()
        chart6_json = self._create_balance_chart(df).to_json()
        
        # ═══════════════════════════════════════════════════════════════
        # HTML ШАБЛОН
        # ═══════════════════════════════════════════════════════════════
        
        html_content = f'''<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title} | Анализ просрочки 90+</title>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;600;700&family=Roboto+Mono:wght@400;500&display=swap" rel="stylesheet">
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Roboto', -apple-system, BlinkMacSystemFont, sans-serif;
            background: linear-gradient(135deg, #1565C0 0%, #0D47A1 50%, #0A3D91 100%);
            background-attachment: fixed;
            min-height: 100vh;
            padding: 20px;
            color: #212121;
            line-height: 1.6;
        }}
        
        .container {{
            max-width: 1800px;
            margin: 0 auto;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* HEADER */
        /* ═══════════════════════════════════════════════════════════════ */
        .header {{
            background: linear-gradient(135deg, #FFFFFF 0%, #F8F9FA 100%);
            border-radius: 16px;
            padding: 28px 36px;
            margin-bottom: 20px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.12);
            border-left: 6px solid #0D47A1;
        }}
        
        .header-content {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 20px;
        }}
        
        .header h1 {{
            font-size: 26px;
            font-weight: 700;
            color: #0D47A1;
            margin-bottom: 4px;
        }}
        
        .header .subtitle {{
            color: #546E7A;
            font-size: 14px;
            font-weight: 400;
        }}
        
        .header-info {{
            text-align: right;
        }}
        
        .header-info .date {{
            color: #78909C;
            font-size: 13px;
            margin-bottom: 8px;
        }}
        
        .period-badge {{
            background: linear-gradient(135deg, #0D47A1, #1565C0);
            color: white;
            padding: 10px 20px;
            border-radius: 8px;
            font-weight: 500;
            font-size: 13px;
            display: inline-block;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* METRICS GRID */
        /* ═══════════════════════════════════════════════════════════════ */
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 16px;
            margin-bottom: 20px;
        }}
        
        @media (max-width: 1400px) {{
            .metrics-grid {{
                grid-template-columns: repeat(3, 1fr);
            }}
        }}
        
        @media (max-width: 900px) {{
            .metrics-grid {{
                grid-template-columns: repeat(2, 1fr);
            }}
        }}
        
        .metric-card {{
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 4px 16px rgba(0,0,0,0.06);
            transition: all 0.25s ease;
            border-left: 4px solid;
            position: relative;
        }}
        
        .metric-card:hover {{
            transform: translateY(-4px);
            box-shadow: 0 8px 24px rgba(0,0,0,0.1);
        }}
        
        .metric-card.blue {{ border-color: #1976D2; }}
        .metric-card.red {{ border-color: #C62828; }}
        .metric-card.green {{ border-color: #2E7D32; }}
        .metric-card.orange {{ border-color: #E65100; }}
        .metric-card.purple {{ border-color: #6A1B9A; }}
        .metric-card.gray {{ border-color: #455A64; }}
        
        .metric-icon {{
            font-size: 26px;
            margin-bottom: 8px;
        }}
        
        .metric-value {{
            font-family: 'Roboto Mono', monospace;
            font-size: 26px;
            font-weight: 700;
            color: #212121;
            line-height: 1.2;
        }}
        
        .metric-value.positive {{ color: #C62828; }}
        .metric-value.negative {{ color: #2E7D32; }}
        
        .metric-label {{
            font-size: 11px;
            color: #78909C;
            margin-top: 6px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            font-weight: 600;
        }}
        
        .metric-sub {{
            font-family: 'Roboto Mono', monospace;
            font-size: 12px;
            color: #90A4AE;
            margin-top: 4px;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* CARDS */
        /* ═══════════════════════════════════════════════════════════════ */
        .card {{
            background: white;
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 4px 16px rgba(0,0,0,0.06);
        }}
        
        .card-title {{
            font-size: 15px;
            font-weight: 600;
            color: #0D47A1;
            margin-bottom: 16px;
            padding-bottom: 12px;
            border-bottom: 2px solid #E3F2FD;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .card-subtitle {{
            font-size: 12px;
            color: #78909C;
            font-weight: 400;
            margin-left: auto;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* LEGEND */
        /* ═══════════════════════════════════════════════════════════════ */
        .legend-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 12px;
        }}
        
        .legend-item {{
            display: flex;
            align-items: flex-start;
            gap: 12px;
            padding: 14px 16px;
            background: #FAFAFA;
            border-radius: 8px;
            border: 1px solid #ECEFF1;
            transition: background 0.2s;
        }}
        
        .legend-item:hover {{
            background: #F5F5F5;
        }}
        
        .legend-color {{
            width: 20px;
            height: 20px;
            border-radius: 4px;
            flex-shrink: 0;
            margin-top: 2px;
        }}
        
        .legend-text strong {{
            color: #37474F;
            font-size: 13px;
            font-weight: 600;
            display: block;
            margin-bottom: 4px;
        }}
        
        .legend-text span {{
            font-size: 11px;
            color: #78909C;
            line-height: 1.5;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* CHART SCROLL CONTAINER */
        /* ═══════════════════════════════════════════════════════════════ */
        .chart-scroll-container {{
            overflow-x: auto;
            overflow-y: hidden;
            padding-bottom: 10px;
        }}
        
        .chart-scroll-container::-webkit-scrollbar {{
            height: 10px;
        }}
        
        .chart-scroll-container::-webkit-scrollbar-track {{
            background: #ECEFF1;
            border-radius: 5px;
        }}
        
        .chart-scroll-container::-webkit-scrollbar-thumb {{
            background: #90A4AE;
            border-radius: 5px;
        }}
        
        .chart-scroll-container::-webkit-scrollbar-thumb:hover {{
            background: #607D8B;
        }}
        
        .chart-inner {{
            min-width: {chart_width}%;
        }}
        
        .charts-row {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 20px;
        }}
        
        @media (max-width: 1200px) {{
            .charts-row {{
                grid-template-columns: 1fr;
            }}
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* TABLE */
        /* ═══════════════════════════════════════════════════════════════ */
        .table-scroll-container {{
            overflow-x: auto;
            padding-bottom: 10px;
        }}
        
        .table-scroll-container::-webkit-scrollbar {{
            height: 10px;
        }}
        
        .table-scroll-container::-webkit-scrollbar-track {{
            background: #ECEFF1;
            border-radius: 5px;
        }}
        
        .table-scroll-container::-webkit-scrollbar-thumb {{
            background: #90A4AE;
            border-radius: 5px;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            min-width: 1300px;
        }}
        
        th {{
            background: #0D47A1;
            color: white;
            padding: 14px 10px;
            text-align: center;
            font-weight: 500;
            font-size: 10px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            white-space: nowrap;
            position: sticky;
            top: 0;
        }}
        
        th:first-child {{
            border-radius: 8px 0 0 0;
            position: sticky;
            left: 0;
            z-index: 2;
        }}
        
        th:last-child {{
            border-radius: 0 8px 0 0;
        }}
        
        th small {{
            display: block;
            font-weight: 400;
            font-size: 9px;
            opacity: 0.85;
            margin-top: 2px;
            text-transform: none;
        }}
        
        td {{
            padding: 12px 10px;
            text-align: right;
            border-bottom: 1px solid #ECEFF1;
            font-family: 'Roboto Mono', monospace;
            font-size: 11px;
            white-space: nowrap;
        }}
        
        td:first-child {{
            text-align: left;
            font-family: 'Roboto', sans-serif;
            font-weight: 500;
            position: sticky;
            left: 0;
            background: white;
            z-index: 1;
        }}
        
        tr:hover td {{
            background: #F5F5F5;
        }}
        
        tr:hover td:first-child {{
            background: #F5F5F5;
        }}
        
        .total-row {{
            background: #E3F2FD !important;
        }}
        
        .total-row td {{
            font-weight: 700;
            border-top: 2px solid #0D47A1;
            color: #0D47A1;
            background: #E3F2FD !important;
        }}
        
        .total-row:hover td {{
            background: #E3F2FD !important;
        }}
        
        .positive {{ color: #C62828; }}
        .negative {{ color: #2E7D32; }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* FOOTER */
        /* ═══════════════════════════════════════════════════════════════ */
        .footer {{
            background: white;
            border-radius: 12px;
            padding: 18px 28px;
            text-align: center;
            box-shadow: 0 4px 16px rgba(0,0,0,0.06);
        }}
        
        .footer p {{
            color: #78909C;
            font-size: 12px;
        }}
        
        .footer strong {{
            color: #0D47A1;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* SCROLL HINT */
        /* ═══════════════════════════════════════════════════════════════ */
        .scroll-hint {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            padding: 8px 16px;
            background: #FFF3E0;
            border-radius: 8px;
            margin-bottom: 12px;
            font-size: 12px;
            color: #E65100;
        }}
        
        .scroll-hint-icon {{
            animation: bounce 1.5s infinite;
        }}
        
        @keyframes bounce {{
            0%, 100% {{ transform: translateX(0); }}
            50% {{ transform: translateX(5px); }}
        }}
    </style>
</head>
<body>
    <div class="container">
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- HEADER -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="header">
            <div class="header-content">
                <div>
                    <h1>🏦 {title}</h1>
                    <p class="subtitle">Анализ просроченной задолженности свыше 90 дней</p>
                </div>
                <div class="header-info">
                    <div class="date">Дата формирования: {timestamp}</div>
                    <div class="period-badge">📅 {period_start} — {period_end} ({num_months} мес.)</div>
                </div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="metrics-grid">
            <div class="metric-card blue">
                <div class="metric-icon">📊</div>
                <div class="metric-value">{format_number(total_on_date_start)}</div>
                <div class="metric-label">На начало периода</div>
                <div class="metric-sub">{format_number(total_on_date_sum_start, 2)} млн сум</div>
            </div>
            
            <div class="metric-card red">
                <div class="metric-icon">📈</div>
                <div class="metric-value">{format_number(total_entered)}</div>
                <div class="metric-label">Вошли в 90+</div>
                <div class="metric-sub">{format_number(total_entered_sum, 2)} млн сум</div>
            </div>
            
            <div class="metric-card green">
                <div class="metric-icon">📉</div>
                <div class="metric-value">{format_number(total_exited)}</div>
                <div class="metric-label">Вышли из 90+</div>
                <div class="metric-sub">{format_number(total_exited_sum, 2)} млн сум</div>
            </div>
            
            <div class="metric-card orange">
                <div class="metric-icon">🛡️</div>
                <div class="metric-value">{format_number(total_insurance)}</div>
                <div class="metric-label">Из них страховка</div>
                <div class="metric-sub">{format_number(total_insurance_sum, 2)} млн сум</div>
            </div>
            
            <div class="metric-card purple">
                <div class="metric-icon">💼</div>
                <div class="metric-value">{format_number(total_other)}</div>
                <div class="metric-label">Без страховки</div>
                <div class="metric-sub">{format_number(total_other_sum, 2)} млн сум</div>
            </div>
            
            <div class="metric-card gray">
                <div class="metric-icon">📊</div>
                <div class="metric-value {'positive' if final_balance > 0 else 'negative' if final_balance < 0 else ''}">{'+' if final_balance > 0 else ''}{format_number(final_balance)}</div>
                <div class="metric-label">Итоговый баланс</div>
                <div class="metric-sub">Вошли − Вышли за период</div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- МЕТОДОЛОГИЯ -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="card">
            <div class="card-title">
                📖 Методология расчёта показателей
                <span class="card-subtitle">Описание формул и логики анализа</span>
            </div>
            <div class="legend-grid">
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['on_date']};"></div>
                    <div class="legend-text">
                        <strong>На отчётную дату</strong>
                        <span>Количество и сумма анкет с просрочкой более 90 дней на начало отчётного месяца. Условие: start_days &gt; 90</span>
                    </div>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['entered']};"></div>
                    <div class="legend-text">
                        <strong>Вошли в 90+</strong>
                        <span>Анкеты, у которых просрочка превысила 90 дней в течение месяца. Условие: start_days ≤ 90 И max_days &gt; 90</span>
                    </div>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['exited']};"></div>
                    <div class="legend-text">
                        <strong>Вышли из 90+</strong>
                        <span>Анкеты 90+, полностью погашенные к концу месяца. Условие: max_days &gt; 90 И end_days = 0</span>
                    </div>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['insurance']};"></div>
                    <div class="legend-text">
                        <strong>Из них страховка</strong>
                        <span>Часть погашенных анкет, которые присутствуют в файле страховых возмещений за соответствующий период</span>
                    </div>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['other']};"></div>
                    <div class="legend-text">
                        <strong>Без страховки (прочие)</strong>
                        <span>Погашения за счёт собственных средств заёмщика, реструктуризации и др. Расчёт: Вышли из 90+ − Страховка</span>
                    </div>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['neutral']};"></div>
                    <div class="legend-text">
                        <strong>Накопленный баланс</strong>
                        <span>Сумма (Вошли − Вышли) нарастающим итогом с начала периода. Показывает чистое изменение размера портфеля 90+</span>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- ГРАФИК 1: КОЛИЧЕСТВО -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="card">
            <div class="card-title">
                📊 Динамика количества анкет по месяцам
                <span class="card-subtitle">Единица измерения: штуки</span>
            </div>
            {f'<div class="scroll-hint"><span class="scroll-hint-icon">👉</span> Прокрутите график вправо для просмотра всех месяцев</div>' if num_months > 12 else ''}
            <div class="chart-scroll-container">
                <div class="chart-inner" id="chart1"></div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- ГРАФИК 2: СУММЫ -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="card">
            <div class="card-title">
                💰 Динамика сумм по месяцам
                <span class="card-subtitle">Единица измерения: млн сум</span>
            </div>
            {f'<div class="scroll-hint"><span class="scroll-hint-icon">👉</span> Прокрутите график вправо для просмотра всех месяцев</div>' if num_months > 12 else ''}
            <div class="chart-scroll-container">
                <div class="chart-inner" id="chart2"></div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- ГРАФИКИ 3-4: WATERFALL И PIE -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="charts-row">
            <div class="card">
                <div class="card-title">
                    🌊 Движение портфеля 90+ за весь период
                    <span class="card-subtitle">Waterfall-диаграмма</span>
                </div>
                <div id="chart3"></div>
            </div>
            <div class="card">
                <div class="card-title">
                    🎯 Структура погашений "Вышли из 90+"
                    <span class="card-subtitle">Доля страховки</span>
                </div>
                <div id="chart4"></div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- ГРАФИК 5: ДЕТАЛИЗАЦИЯ ПОГАШЕНИЙ -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="card">
            <div class="card-title">
                🛡️ Детализация погашений: Страховка vs Без страховки
                <span class="card-subtitle">Единица измерения: млн сум</span>
            </div>
            {f'<div class="scroll-hint"><span class="scroll-hint-icon">👉</span> Прокрутите график вправо для просмотра всех месяцев</div>' if num_months > 12 else ''}
            <div class="chart-scroll-container">
                <div class="chart-inner" id="chart5"></div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- ГРАФИК 6: БАЛАНС -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="card">
            <div class="card-title">
                📈 Баланс: Вошли vs Вышли и накопленный итог
                <span class="card-subtitle">Единица измерения: штуки</span>
            </div>
            {f'<div class="scroll-hint"><span class="scroll-hint-icon">👉</span> Прокрутите график вправо для просмотра всех месяцев</div>' if num_months > 12 else ''}
            <div class="chart-scroll-container">
                <div class="chart-inner" id="chart6"></div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- СВОДНАЯ ТАБЛИЦА -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="card">
            <div class="card-title">
                📋 Сводная таблица по месяцам
                <span class="card-subtitle">Все показатели • Суммы в млн сум</span>
            </div>
            {f'<div class="scroll-hint"><span class="scroll-hint-icon">👉</span> Прокрутите таблицу вправо для просмотра всех колонок</div>'}
            <div class="table-scroll-container">
                {self._create_html_table(df)}
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- FOOTER -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="footer">
            <p><strong>Анализатор просрочки 90+ v6.0</strong> | Период: {period_start} — {period_end} | Всего месяцев: {num_months}</p>
        </div>
        
    </div>
    
    <script>
        const config = {{
            responsive: true,
            displayModeBar: true,
            displaylogo: false,
            modeBarButtonsToRemove: ['lasso2d', 'select2d', 'autoScale2d']
        }};
        
        Plotly.newPlot('chart1', {chart1_json}.data, {chart1_json}.layout, config);
        Plotly.newPlot('chart2', {chart2_json}.data, {chart2_json}.layout, config);
        Plotly.newPlot('chart3', {chart3_json}.data, {chart3_json}.layout, config);
        Plotly.newPlot('chart4', {chart4_json}.data, {chart4_json}.layout, config);
        Plotly.newPlot('chart5', {chart5_json}.data, {chart5_json}.layout, config);
        Plotly.newPlot('chart6', {chart6_json}.data, {chart6_json}.layout, config);
    </script>
</body>
</html>'''
        
        filepath = os.path.join(self.output_path, f"{filename}.html")
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"     ✅ HTML: {filepath}")
        return filepath
    
    def _create_count_chart(self, df):
        """График количества анкет"""
        fig = go.Figure()
        
        periods = df['short_period'].tolist()
        
        fig.add_trace(go.Bar(
            name='На отчётную дату',
            x=periods,
            y=df['on_date_count'],
            marker_color=COLORS['on_date'],
            text=[format_number(x) for x in df['on_date_count']],
            textposition='outside',
            textfont=dict(size=9),
            hovertemplate='<b>%{x}</b><br>На отчётную дату: %{y:,} шт<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            name='Вошли в 90+',
            x=periods,
            y=df['entered_count'],
            marker_color=COLORS['entered'],
            text=[format_number(x) for x in df['entered_count']],
            textposition='outside',
            textfont=dict(size=9),
            hovertemplate='<b>%{x}</b><br>Вошли в 90+: %{y:,} шт<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            name='Вышли (страховка)',
            x=periods,
            y=df['insurance_count'],
            marker_color=COLORS['insurance'],
            text=[format_number(x) for x in df['insurance_count']],
            textposition='outside',
            textfont=dict(size=9),
            hovertemplate='<b>%{x}</b><br>Страховка: %{y:,} шт<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            name='Вышли (без страховки)',
            x=periods,
            y=df['other_count'],
            marker_color=COLORS['other'],
            text=[format_number(x) for x in df['other_count']],
            textposition='outside',
            textfont=dict(size=9),
            hovertemplate='<b>%{x}</b><br>Без страховки: %{y:,} шт<extra></extra>'
        ))
        
        fig.update_layout(
            barmode='group',
            xaxis_tickangle=-45,
            xaxis_title='Отчётный период',
            yaxis_title='Количество анкет (шт)',
            legend=dict(
                orientation='h',
                yanchor='bottom',
                y=1.02,
                xanchor='center',
                x=0.5,
                font=dict(size=11)
            ),
            margin=dict(l=60, r=40, t=80, b=120),
            height=520,
            hovermode='x unified',
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=dict(family='Roboto, sans-serif', size=11)
        )
        
        fig.update_xaxes(gridcolor='#ECEFF1', tickfont=dict(size=10))
        fig.update_yaxes(gridcolor='#ECEFF1', tickformat=',')
        
        return fig
    
    def _create_sum_chart(self, df):
        """График сумм"""
        fig = go.Figure()
        
        periods = df['short_period'].tolist()
        
        fig.add_trace(go.Scatter(
            name='На отчётную дату',
            x=periods,
            y=df['on_date_sum'] / 1e6,
            mode='lines+markers',
            line=dict(color=COLORS['on_date'], width=3),
            marker=dict(size=8),
            hovertemplate='<b>%{x}</b><br>На отчётную дату: %{y:,.2f} млн<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            name='Вошли в 90+',
            x=periods,
            y=df['entered_sum'] / 1e6,
            mode='lines+markers',
            line=dict(color=COLORS['entered'], width=3),
            marker=dict(size=8),
            hovertemplate='<b>%{x}</b><br>Вошли в 90+: %{y:,.2f} млн<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            name='Вышли из 90+ (всего)',
            x=periods,
            y=df['exited_sum'] / 1e6,
            mode='lines+markers',
            line=dict(color=COLORS['exited'], width=3),
            marker=dict(size=8),
            hovertemplate='<b>%{x}</b><br>Вышли из 90+: %{y:,.2f} млн<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            name='Из них страховка',
            x=periods,
            y=df['insurance_sum'] / 1e6,
            mode='lines+markers',
            line=dict(color=COLORS['insurance'], width=2, dash='dash'),
            marker=dict(size=6, symbol='diamond'),
            hovertemplate='<b>%{x}</b><br>Страховка: %{y:,.2f} млн<extra></extra>'
        ))
        
        fig.update_layout(
            xaxis_tickangle=-45,
            xaxis_title='Отчётный период',
            yaxis_title='Сумма (млн сум)',
            legend=dict(
                orientation='h',
                yanchor='bottom',
                y=1.02,
                xanchor='center',
                x=0.5,
                font=dict(size=11)
            ),
            margin=dict(l=60, r=40, t=80, b=120),
            height=520,
            hovermode='x unified',
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=dict(family='Roboto, sans-serif', size=11)
        )
        
        fig.update_xaxes(gridcolor='#ECEFF1', tickfont=dict(size=10))
        fig.update_yaxes(gridcolor='#ECEFF1', tickformat=',.2f')
        
        return fig
    
    def _create_waterfall_chart(self, df):
        """Waterfall диаграмма"""
        start_value = int(df['on_date_count'].iloc[0])
        entered_total = int(df['entered_count'].sum())
        exited_total = int(df['exited_count'].sum())
        end_value = start_value + entered_total - exited_total
        
        fig = go.Figure(go.Waterfall(
            orientation='v',
            measure=['absolute', 'relative', 'relative', 'total'],
            x=['На начало<br>периода', 'Вошли<br>в 90+', 'Вышли<br>из 90+', 'Расчётный<br>итог'],
            y=[start_value, entered_total, -exited_total, end_value],
            text=[format_number(start_value), f'+{format_number(entered_total)}',
                  f'-{format_number(exited_total)}', format_number(end_value)],
            textposition='outside',
            textfont=dict(size=14, family='Roboto Mono'),
            connector={'line': {'color': '#0D47A1', 'width': 2, 'dash': 'dot'}},
            increasing={'marker': {'color': COLORS['entered']}},
            decreasing={'marker': {'color': COLORS['exited']}},
            totals={'marker': {'color': COLORS['on_date']}}
        ))
        
        fig.update_layout(
            showlegend=False,
            margin=dict(l=50, r=50, t=40, b=60),
            height=420,
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=dict(family='Roboto, sans-serif')
        )
        
        fig.update_yaxes(gridcolor='#ECEFF1', tickformat=',')
        
        return fig
    
    def _create_pie_chart(self, df):
        """Круговая диаграмма структуры погашений"""
        insurance_total = int(df['insurance_count'].sum())
        other_total = int(df['other_count'].sum())
        total = insurance_total + other_total
        
        fig = go.Figure(data=[go.Pie(
            labels=['Страховка', 'Без страховки (прочие)'],
            values=[insurance_total, other_total],
            hole=0.55,
            marker_colors=[COLORS['insurance'], COLORS['other']],
            textinfo='label+percent',
            texttemplate='%{label}<br>%{value:,} шт<br>(%{percent})',
            textfont=dict(size=11),
            hovertemplate='<b>%{label}</b><br>Количество: %{value:,} шт<br>Доля: %{percent}<extra></extra>',
            pull=[0.02, 0]
        )])
        
        fig.update_layout(
            annotations=[dict(
                text=f'<b>Всего</b><br>{format_number(total)} шт',
                x=0.5, y=0.5,
                font_size=14,
                showarrow=False,
                font=dict(family='Roboto')
            )],
            margin=dict(l=20, r=20, t=40, b=20),
            height=420,
            paper_bgcolor='white'
        )
        
        return fig
    
    def _create_exit_breakdown_chart(self, df):
        """Детализация погашений по месяцам"""
        fig = go.Figure()
        
        periods = df['short_period'].tolist()
        
        fig.add_trace(go.Bar(
            name='Страховка',
            x=periods,
            y=df['insurance_sum'] / 1e6,
            marker_color=COLORS['insurance'],
            text=[format_number(x / 1e6, 1) for x in df['insurance_sum']],
            textposition='inside',
            textfont=dict(size=9, color='white'),
            hovertemplate='<b>%{x}</b><br>Страховка: %{y:,.2f} млн<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            name='Без страховки (прочие)',
            x=periods,
            y=df['other_sum'] / 1e6,
            marker_color=COLORS['other'],
            text=[format_number(x / 1e6, 1) for x in df['other_sum']],
            textposition='inside',
            textfont=dict(size=9, color='white'),
            hovertemplate='<b>%{x}</b><br>Без страховки: %{y:,.2f} млн<extra></extra>'
        ))
        
        fig.update_layout(
            barmode='stack',
            xaxis_tickangle=-45,
            xaxis_title='Отчётный период',
            yaxis_title='Сумма погашений (млн сум)',
            legend=dict(
                orientation='h',
                yanchor='bottom',
                y=1.02,
                xanchor='center',
                x=0.5,
                font=dict(size=11)
            ),
            margin=dict(l=60, r=40, t=80, b=120),
            height=480,
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=dict(family='Roboto, sans-serif', size=11)
        )
        
        fig.update_xaxes(gridcolor='#ECEFF1', tickfont=dict(size=10))
        fig.update_yaxes(gridcolor='#ECEFF1', tickformat=',.2f')
        
        return fig
    
    def _create_balance_chart(self, df):
        """График баланса с накопленным итогом"""
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        periods = df['short_period'].tolist()
        
        # Вошли (положительные столбцы)
        fig.add_trace(go.Bar(
            name='Вошли в 90+',
            x=periods,
            y=df['entered_count'],
            marker_color=COLORS['entered'],
            hovertemplate='<b>%{x}</b><br>Вошли: +%{y:,} шт<extra></extra>'
        ), secondary_y=False)
        
        # Вышли (отрицательные столбцы)
        fig.add_trace(go.Bar(
            name='Вышли из 90+',
            x=periods,
            y=-df['exited_count'],
            marker_color=COLORS['exited'],
            customdata=df['exited_count'],
            hovertemplate='<b>%{x}</b><br>Вышли: -%{customdata:,} шт<extra></extra>'
        ), secondary_y=False)
        
        # Накопленный баланс (линия)
        fig.add_trace(go.Scatter(
            name='Накопленный баланс',
            x=periods,
            y=df['cumulative_balance'],
            mode='lines+markers+text',
            line=dict(color=COLORS['neutral'], width=3),
            marker=dict(size=8),
            text=[format_number(x) for x in df['cumulative_balance']],
            textposition='top center',
            textfont=dict(size=9),
            hovertemplate='<b>%{x}</b><br>Накоплено: %{y:,} шт<extra></extra>'
        ), secondary_y=True)
        
        fig.update_layout(
            barmode='relative',
            xaxis_tickangle=-45,
            xaxis_title='Отчётный период',
            legend=dict(
                orientation='h',
                yanchor='bottom',
                y=1.02,
                xanchor='center',
                x=0.5,
                font=dict(size=11)
            ),
            margin=dict(l=60, r=80, t=80, b=120),
            height=500,
            plot_bgcolor='white',
            paper_bgcolor='white',
            hovermode='x unified',
            font=dict(family='Roboto, sans-serif', size=11)
        )
        
        fig.update_xaxes(gridcolor='#ECEFF1', tickfont=dict(size=10))
        fig.update_yaxes(
            title_text='Изменение за месяц (шт)',
            gridcolor='#ECEFF1',
            tickformat=',',
            zeroline=True,
            zerolinecolor=COLORS['neutral'],
            zerolinewidth=2,
            secondary_y=False
        )
        fig.update_yaxes(
            title_text='Накопленный баланс (шт)',
            tickformat=',',
            showgrid=False,
            secondary_y=True
        )
        
        return fig
    
    def _create_html_table(self, df):
        """Создание HTML таблицы с правильными итогами"""
        
        # Расчёт итогов
        total_on_date = int(df['on_date_count'].iloc[0])
        total_on_date_sum = df['on_date_sum'].iloc[0]
        total_entered = int(df['entered_count'].sum())
        total_entered_sum = df['entered_sum'].sum()
        total_exited = int(df['exited_count'].sum())
        total_exited_sum = df['exited_sum'].sum()
        total_insurance = int(df['insurance_count'].sum())
        total_insurance_sum = df['insurance_sum'].sum()
        total_other = int(df['other_count'].sum())
        total_other_sum = df['other_sum'].sum()
        total_balance = total_entered - total_exited
        final_cumulative = int(df['cumulative_balance'].iloc[-1])
        
        html = '''<table>
<thead>
<tr>
<th rowspan="2">Период</th>
<th colspan="2">На отчётную дату</th>
<th colspan="2">Вошли в 90+</th>
<th colspan="2">Вышли из 90+</th>
<th colspan="2">Из них страховка</th>
<th colspan="2">Без страховки</th>
<th>Баланс</th>
<th>Накоплено</th>
</tr>
<tr>
<th>шт</th><th>млн сум</th>
<th>шт</th><th>млн сум</th>
<th>шт</th><th>млн сум</th>
<th>шт</th><th>млн сум</th>
<th>шт</th><th>млн сум</th>
<th>шт</th>
<th>шт</th>
</tr>
</thead>
<tbody>'''
        
        for _, row in df.iterrows():
            balance = int(row['monthly_balance'])
            cumulative = int(row['cumulative_balance'])
            
            balance_class = 'positive' if balance > 0 else 'negative' if balance < 0 else ''
            cumulative_class = 'positive' if cumulative > 0 else 'negative' if cumulative < 0 else ''
            
            balance_sign = '+' if balance > 0 else ''
            cumulative_sign = '+' if cumulative > 0 else ''
            
            html += f'''<tr>
<td>{row['period']}</td>
<td>{format_number(row['on_date_count'])}</td>
<td>{format_number(row['on_date_sum'] / 1e6, 2)}</td>
<td>{format_number(row['entered_count'])}</td>
<td>{format_number(row['entered_sum'] / 1e6, 2)}</td>
<td>{format_number(row['exited_count'])}</td>
<td>{format_number(row['exited_sum'] / 1e6, 2)}</td>
<td>{format_number(row['insurance_count'])}</td>
<td>{format_number(row['insurance_sum'] / 1e6, 2)}</td>
<td>{format_number(row['other_count'])}</td>
<td>{format_number(row['other_sum'] / 1e6, 2)}</td>
<td class="{balance_class}">{balance_sign}{format_number(balance)}</td>
<td class="{cumulative_class}">{cumulative_sign}{format_number(cumulative)}</td>
</tr>'''
        
        # Итоговая строка
        total_balance_class = 'positive' if total_balance > 0 else 'negative' if total_balance < 0 else ''
        final_cumulative_class = 'positive' if final_cumulative > 0 else 'negative' if final_cumulative < 0 else ''
        
        total_balance_sign = '+' if total_balance > 0 else ''
        final_cumulative_sign = '+' if final_cumulative > 0 else ''
        
        html += f'''<tr class="total-row">
<td><strong>ИТОГО</strong></td>
<td><strong>{format_number(total_on_date)}</strong></td>
<td><strong>{format_number(total_on_date_sum / 1e6, 2)}</strong></td>
<td><strong>{format_number(total_entered)}</strong></td>
<td><strong>{format_number(total_entered_sum / 1e6, 2)}</strong></td>
<td><strong>{format_number(total_exited)}</strong></td>
<td><strong>{format_number(total_exited_sum / 1e6, 2)}</strong></td>
<td><strong>{format_number(total_insurance)}</strong></td>
<td><strong>{format_number(total_insurance_sum / 1e6, 2)}</strong></td>
<td><strong>{format_number(total_other)}</strong></td>
<td><strong>{format_number(total_other_sum / 1e6, 2)}</strong></td>
<td class="{total_balance_class}"><strong>{total_balance_sign}{format_number(total_balance)}</strong></td>
<td class="{final_cumulative_class}"><strong>{final_cumulative_sign}{format_number(final_cumulative)}</strong></td>
</tr>'''
        
        html += '</tbody></table>'
        return html
    
    def _create_excel_report(self, df, filename):
        """Создание Excel отчёта"""
        filepath = os.path.join(self.output_path, f"{filename}.xlsx")
        
        export_df = df.copy()
        
        # Переименование колонок
        export_df = export_df.rename(columns={
            'period': 'Период',
            'on_date_count': 'На отчётную дату (шт)',
            'on_date_sum': 'На отчётную дату (сумма)',
            'entered_count': 'Вошли в 90+ (шт)',
            'entered_sum': 'Вошли в 90+ (сумма)',
            'exited_count': 'Вышли из 90+ (шт)',
            'exited_sum': 'Вышли из 90+ (сумма)',
            'insurance_count': 'Страховка (шт)',
            'insurance_sum': 'Страховка (сумма)',
            'other_count': 'Без страховки (шт)',
            'other_sum': 'Без страховки (сумма)',
            'monthly_balance': 'Баланс за месяц',
            'cumulative_balance': 'Накопленный баланс'
        })
        
        # Удаление служебных колонок
        drop_cols = ['short_period', 'prefix', 'year', 'month_num', 'period_key', 
                    'sort_key', 'monthly_balance_sum']
        export_df = export_df.drop(columns=[c for c in drop_cols if c in export_df.columns], 
                                   errors='ignore')
        
        export_df.to_excel(filepath, index=False, sheet_name='Данные')
        print(f"     ✅ Excel: {filepath}")
        return filepath


# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# =============================================================================

def main():
    """Точка входа в приложение"""
    print("\n" + "═"*70)
    print("  🏦 АНАЛИЗАТОР ПРОСРОЧКИ 90+ | Версия 6.0")
    print("  Профессиональный инструмент для банковской аналитики")
    print("═"*70)
    
    # Запуск GUI
    app = MainApplication()
    
    if not app.run():
        print("\n  ❌ Операция отменена пользователем")
        return
    
    try:
        # Анализ данных
        analyzer = DataAnalyzer(
            data_files=app.data_files,
            insurance_file=app.insurance_file
        )
        
        analyzer.load_all_data()
        analyzer.analyze_all()
        
        # Генерация отчётов
        reporter = ReportGenerator(analyzer, app.output_path)
        
        html_path = None
        
        if app.analysis_mode == 'separate':
            print("\n  📄 Режим: Раздельные отчёты по годам")
            paths = reporter.generate_separate_reports()
            html_path = paths[0] if paths else None
            
        elif app.analysis_mode == 'combined':
            print("\n  📋 Режим: Объединённый отчёт")
            html_path = reporter.generate_combined_report()
            
        elif app.analysis_mode == 'both':
            print("\n  📄📋 Режим: Раздельные + Объединённый отчёты")
            reporter.generate_separate_reports()
            html_path = reporter.generate_combined_report()
        
        # Открытие отчёта в браузере
        if html_path:
            import webbrowser
            webbrowser.open(f'file://{os.path.abspath(html_path)}')
        
        print("\n" + "═"*70)
        print("  ✅ ОТЧЁТЫ УСПЕШНО СФОРМИРОВАНЫ!")
        print("═"*70)
        print(f"\n  📂 Результаты сохранены в: {app.output_path}")
        
    except Exception as e:
        print(f"\n  ❌ Ошибка: {str(e)}")
        import traceback
        traceback.print_exc()
        messagebox.showerror("Ошибка выполнения", 
                            f"Произошла ошибка при формировании отчёта:\n\n{str(e)}")


if __name__ == "__main__":
    main()
