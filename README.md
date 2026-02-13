# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                    90+ ПРОСРОЧКА АНАЛИЗАТОР v3.0                            ║
║                    Интерактивный Dashboard для банка                        ║
╚══════════════════════════════════════════════════════════════════════════════╝

Описание: Анализ просроченных кредитов свыше 90 дней
Версия: 3.0 - исправлена загрузка страховки, улучшен дизайн
"""

import pandas as pd
import numpy as np
import os
import re
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Интерактивные графики
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# Excel
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# GUI
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# =============================================================================
# КОНСТАНТЫ
# =============================================================================

THRESHOLD = 90  # Порог просрочки

# Цветовая схема (банковская, профессиональная)
COLORS = {
    'primary': '#1E3A5F',       # Темно-синий (основной)
    'secondary': '#3D5A80',     # Синий
    'start': '#2196F3',         # Голубой - начало месяца
    'new': '#F44336',           # Красный - новые 90+
    'closed': '#4CAF50',        # Зеленый - погашено
    'insurance': '#FF9800',     # Оранжевый - страховка
    'other': '#9C27B0',         # Фиолетовый - прочие
    'accent': '#00BCD4',        # Бирюзовый
    'background': '#F5F7FA',    # Светлый фон
    'card': '#FFFFFF',          # Белый для карточек
    'text': '#333333',          # Темный текст
    'text_light': '#666666',    # Светлый текст
    'border': '#E0E6ED',        # Границы
    'success': '#28A745',
    'warning': '#FFC107',
    'danger': '#DC3545'
}

# Названия месяцев
MONTH_NAMES_RU = {
    'jan': 'Январь', 'feb': 'Февраль', 'mar': 'Март', 'apr': 'Апрель',
    'may': 'Май', 'jun': 'Июнь', 'jul': 'Июль', 'aug': 'Август',
    'sep': 'Сентябрь', 'oct': 'Октябрь', 'nov': 'Ноябрь', 'dec': 'Декабрь'
}


# =============================================================================
# ШАБЛОН СТРАХОВКИ
# =============================================================================

def create_insurance_template(output_path: str):
    """Создание шаблона файла страховки"""
    template_df = pd.DataFrame({
        'dealid': [12345678, 23456789, 34567890, 45678901, 56789012],
        'period': ['2024-01', '2024-01', '2024-02', '2024-03', '2025-01']
    })
    
    filepath = os.path.join(output_path, "ШАБЛОН_Страховка.xlsx")
    template_df.to_excel(filepath, index=False)
    
    return filepath


# =============================================================================
# КЛАСС ВЫБОРА ФАЙЛОВ (GUI)
# =============================================================================

class FileSelector:
    """GUI для выбора файлов"""
    
    def __init__(self):
        self.data_files = []
        self.insurance_file = None
        self.output_path = None
        self.result = False
        
    def select_files(self):
        """Открыть диалог выбора файлов"""
        self.root = tk.Tk()
        self.root.title("90+ Просрочка Анализатор v3.0")
        self.root.geometry("850x650")
        self.root.configure(bg='#F5F7FA')
        
        # Центрирование
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (425)
        y = (self.root.winfo_screenheight() // 2) - (325)
        self.root.geometry(f'850x650+{x}+{y}')
        
        # Стили
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'), 
                       background='#F5F7FA', foreground='#1E3A5F')
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'), 
                       background='#F5F7FA', foreground='#1E3A5F')
        style.configure('Info.TLabel', font=('Segoe UI', 9), 
                       background='#F5F7FA', foreground='#666666')
        style.configure('TButton', font=('Segoe UI', 10), padding=8)
        style.configure('Action.TButton', font=('Segoe UI', 12, 'bold'), padding=12)
        style.configure('TLabelframe', background='#F5F7FA')
        style.configure('TLabelframe.Label', font=('Segoe UI', 10, 'bold'),
                       background='#F5F7FA', foreground='#1E3A5F')
        
        # Главный контейнер
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.pack(fill='both', expand=True)
        
        # Заголовок
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill='x', pady=(0, 25))
        
        ttk.Label(title_frame, text="🏦 Анализатор просрочки 90+", 
                 style='Title.TLabel').pack()
        ttk.Label(title_frame, text="Профессиональный анализ кредитного портфеля", 
                 style='Info.TLabel').pack(pady=(5, 0))
        
        # ═══════════════════════════════════════════════════════════════
        # Секция файлов данных
        # ═══════════════════════════════════════════════════════════════
        data_frame = ttk.LabelFrame(main_frame, text=" 📁 Файлы данных ", padding="15")
        data_frame.pack(fill='x', pady=10)
        
        list_frame = ttk.Frame(data_frame)
        list_frame.pack(fill='x')
        
        self.files_listbox = tk.Listbox(list_frame, height=5, font=('Consolas', 10),
                                        selectmode=tk.SINGLE, bg='white',
                                        relief='flat', borderwidth=1,
                                        highlightthickness=1, highlightcolor='#2196F3')
        self.files_listbox.pack(side='left', fill='x', expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', 
                                 command=self.files_listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.files_listbox.config(yscrollcommand=scrollbar.set)
        
        btn_frame = ttk.Frame(data_frame)
        btn_frame.pack(fill='x', pady=(10, 0))
        
        ttk.Button(btn_frame, text="➕ Добавить", 
                  command=self._add_data_file).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="➖ Удалить", 
                  command=self._remove_data_file).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="🗑️ Очистить", 
                  command=self._clear_data_files).pack(side='left', padx=5)
        
        # ═══════════════════════════════════════════════════════════════
        # Секция страховки
        # ═══════════════════════════════════════════════════════════════
        insurance_frame = ttk.LabelFrame(main_frame, text=" 🛡️ Страховые погашения ", 
                                        padding="15")
        insurance_frame.pack(fill='x', pady=10)
        
        ins_inner = ttk.Frame(insurance_frame)
        ins_inner.pack(fill='x')
        
        self.insurance_var = tk.StringVar(value="Не выбран (опционально)")
        ttk.Entry(ins_inner, textvariable=self.insurance_var, 
                 state='readonly', width=55, font=('Segoe UI', 9)).pack(side='left', fill='x', expand=True)
        ttk.Button(ins_inner, text="📂 Выбрать", 
                  command=self._select_insurance).pack(side='left', padx=(10, 5))
        ttk.Button(ins_inner, text="📋 Шаблон", 
                  command=self._create_template).pack(side='left')
        
        # Описание формата
        format_text = "Формат: dealid | period (2024-01, 2024-02 и т.д.) • Суммы берутся из основных данных (max_rest)"
        ttk.Label(insurance_frame, text=format_text, style='Info.TLabel').pack(anchor='w', pady=(10, 0))
        
        # ═══════════════════════════════════════════════════════════════
        # Секция вывода
        # ═══════════════════════════════════════════════════════════════
        output_frame = ttk.LabelFrame(main_frame, text=" 📂 Папка результатов ", padding="15")
        output_frame.pack(fill='x', pady=10)
        
        out_inner = ttk.Frame(output_frame)
        out_inner.pack(fill='x')
        
        self.output_var = tk.StringVar(value="Не выбрана")
        ttk.Entry(out_inner, textvariable=self.output_var, 
                 state='readonly', width=65, font=('Segoe UI', 9)).pack(side='left', fill='x', expand=True)
        ttk.Button(out_inner, text="📂 Выбрать", 
                  command=self._select_output).pack(side='left', padx=(10, 0))
        
        # ═══════════════════════════════════════════════════════════════
        # Информация
        # ═══════════════════════════════════════════════════════════════
        info_frame = ttk.LabelFrame(main_frame, text=" ℹ️ Информация ", padding="15")
        info_frame.pack(fill='x', pady=10)
        
        info_text = """• Поддерживаются любые годы (2024, 2025, 2026 и т.д.)
• Файл страховки: только dealid и period — суммы берутся автоматически из основных данных
• Дубликаты dealid в страховке учитываются только од��н раз
• Результат: интерактивный HTML-отчет + Excel таблица"""
        
        ttk.Label(info_frame, text=info_text, style='Info.TLabel', 
                 justify='left').pack(anchor='w')
        
        # ═══════════════════════════════════════════════════════════════
        # Кнопки действий
        # ═══════════════════════════════════════════════════════════════
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill='x', pady=20)
        
        ttk.Button(action_frame, text="✅ Запустить анализ", 
                  style='Action.TButton', command=self._on_submit).pack(side='left', padx=10)
        ttk.Button(action_frame, text="❌ Отмена", 
                  command=self._on_cancel).pack(side='left', padx=10)
        
        self.root.mainloop()
        return self.result
    
    def _add_data_file(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл данных",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filepath and filepath not in self.data_files:
            self.data_files.append(filepath)
            self.files_listbox.insert(tk.END, f"📄 {os.path.basename(filepath)}")
    
    def _remove_data_file(self):
        selection = self.files_listbox.curselection()
        if selection:
            idx = selection[0]
            self.files_listbox.delete(idx)
            del self.data_files[idx]
    
    def _clear_data_files(self):
        self.files_listbox.delete(0, tk.END)
        self.data_files.clear()
    
    def _select_insurance(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл страховых погашений",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filepath:
            self.insurance_file = filepath
            self.insurance_var.set(os.path.basename(filepath))
    
    def _create_template(self):
        folder = filedialog.askdirectory(title="Папка для шаблона")
        if folder:
            filepath = create_insurance_template(folder)
            messagebox.showinfo("Готово", f"Шаблон создан:\n{filepath}")
    
    def _select_output(self):
        folder = filedialog.askdirectory(title="Папка для результатов")
        if folder:
            self.output_path = folder
            self.output_var.set(folder)
    
    def _on_submit(self):
        if not self.data_files:
            messagebox.showerror("Ошибка", "Добавьте хотя бы один файл данных!")
            return
        if not self.output_path:
            messagebox.showerror("Ошибка", "Выберите папку для результатов!")
            return
        
        self.result = True
        self.root.quit()
        self.root.destroy()
    
    def _on_cancel(self):
        self.root.quit()
        self.root.destroy()


# =============================================================================
# КЛАСС АНАЛИЗА ДАННЫХ
# =============================================================================

class Prosrochka90Analyzer:
    """Универсальный анализатор просрочки 90+"""
    
    def __init__(self, data_files: list, insurance_file: str = None):
        self.data_files = data_files
        self.insurance_file = insurance_file
        
        self.df_combined = None
        self.df_insurance = None
        self.insurance_by_period = {}  # {period: set(dealids)}
        self.results = []
        self.months_order = []
        
    def load_data(self):
        """Загрузка всех файлов данных"""
        print("\n" + "="*70)
        print("📥 ЗАГРУЗКА ДАННЫХ")
        print("="*70)
        
        all_dfs = []
        
        for filepath in self.data_files:
            print(f"\n📄 Файл: {os.path.basename(filepath)}")
            df = pd.read_excel(filepath)
            df.columns = [str(col).lower().strip() for col in df.columns]
            all_dfs.append(df)
            print(f"   ✅ Загружено: {len(df):,} записей")
        
        # Объединение
        if len(all_dfs) > 1:
            self.df_combined = all_dfs[0]
            for df in all_dfs[1:]:
                self.df_combined = self.df_combined.merge(
                    df, on='dealid', how='outer', suffixes=('', '_dup')
                )
                dup_cols = [c for c in self.df_combined.columns if c.endswith('_dup')]
                self.df_combined.drop(columns=dup_cols, inplace=True, errors='ignore')
        else:
            self.df_combined = all_dfs[0]
        
        print(f"\n📊 Всего уникальных анкет: {len(self.df_combined):,}")
        
        # Загрузка страховки
        if self.insurance_file:
            self._load_insurance()
        
        # Определение месяцев
        self._detect_months()
    
    def _load_insurance(self):
        """Загрузка данных страховки с улучшенной обработкой"""
        print(f"\n🛡️ Загрузка страховки: {os.path.basename(self.insurance_file)}")
        
        try:
            # Пробуем разные варианты загрузки
            df_ins = None
            
            # Вариант 1: Обычная загрузка
            try:
                df_ins = pd.read_excel(self.insurance_file)
                df_ins.columns = [str(col).lower().strip() for col in df_ins.columns]
            except:
                pass
            
            # Вариант 2: Если данные начинаются с 3-й строки (шаблон)
            if df_ins is None or len(df_ins) == 0:
                df_ins = pd.read_excel(self.insurance_file, skiprows=3)
                df_ins.columns = [str(col).lower().strip() for col in df_ins.columns]
            
            # Поиск нужных колонок
            dealid_col = None
            period_col = None
            
            # Возможные названия колонок
            dealid_variants = ['dealid', 'deal_id', 'анкета', 'id', 'deal', 'номер']
            period_variants = ['period', 'период', 'дата', 'месяц', 'date', 'month']
            
            for col in df_ins.columns:
                col_lower = col.lower().strip()
                
                # Поиск dealid
                if dealid_col is None:
                    for variant in dealid_variants:
                        if variant in col_lower:
                            dealid_col = col
                            break
                
                # Поиск period
                if period_col is None:
                    for variant in period_variants:
                        if variant in col_lower:
                            period_col = col
                            break
            
            # Если колонки не найдены, пробуем по позиции
            if dealid_col is None and len(df_ins.columns) >= 1:
                dealid_col = df_ins.columns[0]
                print(f"   ⚠️ Колонка dealid не найдена, используется первая колонка: {dealid_col}")
            
            if period_col is None and len(df_ins.columns) >= 2:
                period_col = df_ins.columns[1]
                print(f"   ⚠️ Колонка period не найдена, используется вторая колонка: {period_col}")
            
            if dealid_col is None or period_col is None:
                print("   ❌ Не удалось определить структуру файла страховки")
                self.df_insurance = None
                return
            
            # Переименование колонок
            df_ins = df_ins.rename(columns={dealid_col: 'dealid', period_col: 'period'})
            
            # Очистка данных
            df_ins = df_ins[['dealid', 'period']].dropna()
            df_ins['dealid'] = pd.to_numeric(df_ins['dealid'], errors='coerce')
            df_ins = df_ins.dropna(subset=['dealid'])
            df_ins['dealid'] = df_ins['dealid'].astype(int)
            
            # Парсинг периода
            df_ins['period_parsed'] = df_ins['period'].apply(self._parse_period)
            df_ins = df_ins.dropna(subset=['period_parsed'])
            
            # Удаление дубликатов (один dealid на период учитывается один раз)
            df_ins = df_ins.drop_duplicates(subset=['dealid', 'period_parsed'])
            
            self.df_insurance = df_ins
            
            # Группировка по периодам
            for period in df_ins['period_parsed'].unique():
                mask = df_ins['period_parsed'] == period
                self.insurance_by_period[period] = set(df_ins.loc[mask, 'dealid'].tolist())
            
            print(f"   ✅ Загружено: {len(df_ins):,} уникальных записей")
            print(f"   📅 Периоды: {', '.join(sorted(self.insurance_by_period.keys()))}")
            
        except Exception as e:
            print(f"   ❌ Ошибка загрузки страховки: {str(e)}")
            self.df_insurance = None
    
    def _parse_period(self, period_str):
        """Парсинг периода в формат 'YYYY-MM'"""
        if pd.isna(period_str):
            return None
        
        period_str = str(period_str).strip()
        
        # Формат: 2024-01
        match = re.match(r'(\d{4})[-/\.](\d{1,2})', period_str)
        if match:
            return f"{match.group(1)}-{int(match.group(2)):02d}"
        
        # Формат: 01.2024
        match = re.match(r'(\d{1,2})[-./](\d{4})', period_str)
        if match:
            return f"{match.group(2)}-{int(match.group(1)):02d}"
        
        # Формат: 01-2024
        match = re.match(r'(\d{1,2})[-](\d{4})', period_str)
        if match:
            return f"{match.group(2)}-{int(match.group(1)):02d}"
        
        return None
    
    def _detect_months(self):
        """Определение доступных месяцев"""
        print("\n🔍 Определение периодов...")
        
        pattern = re.compile(r'^([a-z]{3})(\d{2})_start_days$')
        months_found = []
        
        for col in self.df_combined.columns:
            match = pattern.match(str(col))
            if match:
                month_code = match.group(1)
                year_code = match.group(2)
                year_full = 2000 + int(year_code)
                
                month_order = {
                    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4,
                    'may': 5, 'jun': 6, 'jul': 7, 'aug': 8,
                    'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
                }
                
                month_num = month_order.get(month_code, 0)
                month_name = MONTH_NAMES_RU.get(month_code, month_code)
                prefix = f"{month_code}{year_code}"
                
                months_found.append({
                    'prefix': prefix,
                    'month_code': month_code,
                    'year': year_full,
                    'month_num': month_num,
                    'name_ru': f"{month_name} {year_full}",
                    'period_key': f"{year_full}-{month_num:02d}",
                    'sort_key': year_full * 100 + month_num
                })
        
        months_found.sort(key=lambda x: x['sort_key'])
        self.months_order = months_found
        
        print(f"   ✅ Найдено {len(months_found)} периодов")
    
    def analyze_month(self, month_info: dict) -> dict:
        """Анализ одного месяца"""
        prefix = month_info['prefix']
        month_name = month_info['name_ru']
        period_key = month_info['period_key']
        
        # Колонки
        start_days = f'{prefix}_start_days'
        max_days = f'{prefix}_max_days'
        end_days = f'{prefix}_end_days'
        start_rest = f'{prefix}_start_rest'
        
        # Поиск колонки max_rest
        max_rest_col = None
        for col_name in [f'{prefix}_max_rest_ref', f'{prefix}_max_rest']:
            if col_name in self.df_combined.columns:
                max_rest_col = col_name
                break
        
        # Проверка колонок
        required = [start_days, max_days, end_days]
        for col in required:
            if col not in self.df_combined.columns:
                return None
        
        df = self.df_combined.copy()
        
        # Заполнение NaN
        for col in [start_days, max_days, end_days, start_rest]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        if max_rest_col and max_rest_col in df.columns:
            df[max_rest_col] = pd.to_numeric(df[max_rest_col], errors='coerce').fillna(0)
        
        result = {
            'period': month_name,
            'prefix': prefix,
            'year': month_info['year'],
            'month_num': month_info['month_num'],
            'period_key': period_key,
            'sort_key': month_info['sort_key']
        }
        
        # ═══════════════════════════════════════════════════════════════
        # 1. НАЧАЛО МЕСЯЦА: start_days > 90
        # ═══════════════════════════════════════════════════════════════
        mask_start = df[start_days] > THRESHOLD
        result['start_count'] = int(mask_start.sum())
        result['start_sum'] = float(df.loc[mask_start, start_rest].sum()) if start_rest in df.columns else 0.0
        
        # ═══════════════════════════════════════════════════════════════
        # 2. НОВЫЕ 90+: start_days <= 90 AND max_days > 90
        # ═══════════════════════════════════════════════════════════════
        mask_new = (df[start_days] <= THRESHOLD) & (df[max_days] > THRESHOLD)
        result['new_count'] = int(mask_new.sum())
        if max_rest_col and max_rest_col in df.columns:
            result['new_sum'] = float(df.loc[mask_new, max_rest_col].sum())
        else:
            result['new_sum'] = 0.0
        
        # ═══════════════════════════════════════════════════════════════
        # 3. ПОГАШЕНО: max_days > 90 AND end_days == 0
        # ═══════════════════════════════════════════════════════════════
        mask_closed = (df[max_days] > THRESHOLD) & (df[end_days] == 0)
        closed_df = df[mask_closed].copy()
        closed_dealids = set(closed_df['dealid'].tolist())
        
        result['closed_count'] = int(mask_closed.sum())
        if max_rest_col and max_rest_col in df.columns:
            result['closed_sum'] = float(closed_df[max_rest_col].sum())
        else:
            result['closed_sum'] = 0.0
        
        # ═══════════════════════════════════════════════════════════════
        # 4. СТРАХОВКА
        # ═══════════════════════════════════════════════════════════════
        result['insurance_count'] = 0
        result['insurance_sum'] = 0.0
        
        if period_key in self.insurance_by_period:
            insurance_dealids = self.insurance_by_period[period_key]
            
            # Пересечение: погашенные + в страховке
            insurance_in_closed = closed_dealids.intersection(insurance_dealids)
            
            result['insurance_count'] = len(insurance_in_closed)
            
            # Сумма страховки из основных данных (max_rest)
            if insurance_in_closed and max_rest_col and max_rest_col in df.columns:
                mask_insurance = df['dealid'].isin(insurance_in_closed) & mask_closed
                result['insurance_sum'] = float(df.loc[mask_insurance, max_rest_col].sum())
        
        # ═══════════════════════════════════════════════════════════════
        # 5. ПРОЧИЕ ПОГАШЕНИЯ
        # ═══════════════════════════════════════════════════════════════
        result['other_closed_count'] = result['closed_count'] - result['insurance_count']
        result['other_closed_sum'] = result['closed_sum'] - result['insurance_sum']
        
        return result
    
    def analyze_all(self):
        """Анализ всех периодов"""
        print("\n" + "="*70)
        print("📊 АНАЛИЗ ДАННЫХ")
        print("="*70)
        
        for month_info in self.months_order:
            result = self.analyze_month(month_info)
            if result:
                self.results.append(result)
                
                print(f"\n✅ {result['period']}:")
                print(f"   Начало 90+:     {result['start_count']:>6,} шт | {result['start_sum']/1e6:>10,.2f} млн")
                print(f"   Новые 90+:      {result['new_count']:>6,} шт | {result['new_sum']/1e6:>10,.2f} млн")
                print(f"   Погашено:       {result['closed_count']:>6,} шт | {result['closed_sum']/1e6:>10,.2f} млн")
                print(f"    ├─ Страховка:  {result['insurance_count']:>6,} шт | {result['insurance_sum']/1e6:>10,.2f} млн")
                print(f"    └─ Прочие:     {result['other_closed_count']:>6,} шт | {result['other_closed_sum']/1e6:>10,.2f} млн")
        
        return pd.DataFrame(self.results)


# =============================================================================
# ГЕНЕРАТОР ОТЧЕТОВ
# =============================================================================

class InteractiveReportGenerator:
    """Генератор интерактивных HTML отчетов"""
    
    def __init__(self, analyzer: Prosrochka90Analyzer, output_path: str):
        self.analyzer = analyzer
        self.output_path = output_path
        self.df = pd.DataFrame(analyzer.results)
        
    def generate_all(self):
        """Генерация всех отчетов"""
        print("\n" + "="*70)
        print("📝 ГЕНЕРАЦИЯ ОТЧЕТОВ")
        print("="*70)
        
        html_path = self._generate_html_dashboard()
        excel_path = self._generate_excel_report()
        
        print("\n" + "="*70)
        print("✅ ОТЧЕТЫ СОЗДАНЫ!")
        print("="*70)
        print(f"\n🌐 HTML: {html_path}")
        print(f"📊 Excel: {excel_path}")
        
        return html_path
    
    def _generate_html_dashboard(self):
        """Генерация HTML dashboard"""
        print("\n🌐 Создание HTML dashboard...")
        
        timestamp = datetime.now().strftime("%d.%m.%Y %H:%M")
        
        # Метрики
        total_start = self.df['start_count'].iloc[0] if len(self.df) > 0 else 0
        total_new = self.df['new_count'].sum()
        total_closed = self.df['closed_count'].sum()
        total_insurance = self.df['insurance_count'].sum()
        total_other = self.df['other_closed_count'].sum()
        
        total_new_sum = self.df['new_sum'].sum() / 1e6
        total_closed_sum = self.df['closed_sum'].sum() / 1e6
        total_insurance_sum = self.df['insurance_sum'].sum() / 1e6
        
        # Период
        period_start = self.df['period'].iloc[0] if len(self.df) > 0 else ""
        period_end = self.df['period'].iloc[-1] if len(self.df) > 0 else ""
        
        # Генерация графиков
        chart1_json = self._create_main_bar_chart().to_json()
        chart2_json = self._create_sum_chart().to_json()
        chart3_json = self._create_waterfall_chart().to_json()
        chart4_json = self._create_pie_chart().to_json()
        chart5_json = self._create_insurance_stack_chart().to_json()
        chart6_json = self._create_monthly_trend_chart().to_json()
        
        html_content = f'''<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Анализ просрочки 90+ | Dashboard</title>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            background-attachment: fixed;
            min-height: 100vh;
            padding: 20px;
            color: #333;
        }}
        
        .container {{
            max-width: 1600px;
            margin: 0 auto;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* HEADER */
        /* ═══════════════════════════════════════════════════════════════ */
        .header {{
            background: linear-gradient(135deg, #1E3A5F 0%, #2C5282 100%);
            border-radius: 20px;
            padding: 30px 40px;
            margin-bottom: 25px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            color: white;
        }}
        
        .header-content {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 20px;
        }}
        
        .header h1 {{
            font-size: 2.2em;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 15px;
        }}
        
        .header h1 span {{
            font-size: 1.5em;
        }}
        
        .header-info {{
            text-align: right;
        }}
        
        .header-info p {{
            opacity: 0.9;
            font-size: 0.95em;
        }}
        
        .header-info .period {{
            font-size: 1.1em;
            font-weight: 600;
            margin-top: 5px;
            background: rgba(255,255,255,0.2);
            padding: 8px 16px;
            border-radius: 8px;
            display: inline-block;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* METRIC CARDS */
        /* ═══════════════════════════════════════════════════════════════ */
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 20px;
            margin-bottom: 25px;
        }}
        
        .metric-card {{
            background: white;
            border-radius: 16px;
            padding: 25px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }}
        
        .metric-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 20px 60px rgba(0,0,0,0.15);
        }}
        
        .metric-card::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 5px;
            height: 100%;
        }}
        
        .metric-card.blue::before {{ background: linear-gradient(180deg, #2196F3, #1976D2); }}
        .metric-card.red::before {{ background: linear-gradient(180deg, #F44336, #D32F2F); }}
        .metric-card.green::before {{ background: linear-gradient(180deg, #4CAF50, #388E3C); }}
        .metric-card.orange::before {{ background: linear-gradient(180deg, #FF9800, #F57C00); }}
        .metric-card.purple::before {{ background: linear-gradient(180deg, #9C27B0, #7B1FA2); }}
        
        .metric-icon {{
            font-size: 2.5em;
            margin-bottom: 10px;
        }}
        
        .metric-value {{
            font-size: 2.2em;
            font-weight: 700;
            color: #1E3A5F;
            line-height: 1.2;
        }}
        
        .metric-label {{
            font-size: 0.9em;
            color: #666;
            margin-top: 8px;
            font-weight: 500;
        }}
        
        .metric-sub {{
            font-size: 0.85em;
            color: #999;
            margin-top: 4px;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* LEGEND BOX */
        /* ═══════════════════════════════════════════════════════════════ */
        .legend-card {{
            background: white;
            border-radius: 16px;
            padding: 25px 30px;
            margin-bottom: 25px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
        }}
        
        .legend-title {{
            font-size: 1.2em;
            font-weight: 600;
            color: #1E3A5F;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .legend-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 15px;
        }}
        
        .legend-item {{
            display: flex;
            align-items: flex-start;
            gap: 12px;
            padding: 12px;
            background: #F8FAFC;
            border-radius: 10px;
            transition: background 0.2s;
        }}
        
        .legend-item:hover {{
            background: #EDF2F7;
        }}
        
        .legend-color {{
            width: 24px;
            height: 24px;
            border-radius: 6px;
            flex-shrink: 0;
            margin-top: 2px;
        }}
        
        .legend-text strong {{
            color: #1E3A5F;
            display: block;
            margin-bottom: 4px;
        }}
        
        .legend-text span {{
            font-size: 0.85em;
            color: #666;
            line-height: 1.4;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* CHART CARDS */
        /* ═══════════════════════════════════════════════════════════════ */
        .chart-card {{
            background: white;
            border-radius: 16px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
        }}
        
        .chart-title {{
            font-size: 1.15em;
            font-weight: 600;
            color: #1E3A5F;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 2px solid #EDF2F7;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .charts-row {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 25px;
        }}
        
        @media (max-width: 1100px) {{
            .charts-row {{
                grid-template-columns: 1fr;
            }}
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* TABLE */
        /* ═══════════════════════════════════════════════════════════════ */
        .table-card {{
            background: white;
            border-radius: 16px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            overflow-x: auto;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9em;
        }}
        
        thead {{
            position: sticky;
            top: 0;
        }}
        
        th {{
            background: linear-gradient(135deg, #1E3A5F 0%, #2C5282 100%);
            color: white;
            padding: 14px 12px;
            text-align: center;
            font-weight: 600;
            font-size: 0.85em;
            white-space: nowrap;
        }}
        
        th:first-child {{
            border-radius: 10px 0 0 0;
        }}
        
        th:last-child {{
            border-radius: 0 10px 0 0;
        }}
        
        td {{
            padding: 12px;
            text-align: center;
            border-bottom: 1px solid #EDF2F7;
        }}
        
        tr:hover td {{
            background: #F8FAFC;
        }}
        
        tr:last-child td:first-child {{
            border-radius: 0 0 0 10px;
        }}
        
        tr:last-child td:last-child {{
            border-radius: 0 0 10px 0;
        }}
        
        .total-row {{
            background: linear-gradient(135deg, #EBF8FF 0%, #E6FFFA 100%) !important;
            font-weight: 600;
        }}
        
        .total-row td {{
            border-top: 2px solid #1E3A5F;
        }}
        
        /* ═══════════════════════════════════════════════════════════════ */
        /* FOOTER */
        /* ═══════════════════════════════════════════════════════════════ */
        .footer {{
            background: white;
            border-radius: 16px;
            padding: 20px 30px;
            text-align: center;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
        }}
        
        .footer p {{
            color: #666;
            font-size: 0.9em;
        }}
        
        .footer strong {{
            color: #1E3A5F;
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
                <h1><span>🏦</span> Анализ просрочки 90+ дней</h1>
                <div class="header-info">
                    <p>Сформировано: {timestamp}</p>
                    <div class="period">📅 {period_start} — {period_end}</div>
                </div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- METRICS -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="metrics-grid">
            <div class="metric-card blue">
                <div class="metric-icon">📊</div>
                <div class="metric-value">{total_start:,}</div>
                <div class="metric-label">90+ на начало периода</div>
                <div class="metric-sub">Стартовая база</div>
            </div>
            
            <div class="metric-card red">
                <div class="metric-icon">📈</div>
                <div class="metric-value">{total_new:,}</div>
                <div class="metric-label">Новых 90+ за период</div>
                <div class="metric-sub">{total_new_sum:,.1f} млн сум</div>
            </div>
            
            <div class="metric-card green">
                <div class="metric-icon">✅</div>
                <div class="metric-value">{total_closed:,}</div>
                <div class="metric-label">Погашено всего</div>
                <div class="metric-sub">{total_closed_sum:,.1f} млн сум</div>
            </div>
            
            <div class="metric-card orange">
                <div class="metric-icon">🛡️</div>
                <div class="metric-value">{total_insurance:,}</div>
                <div class="metric-label">Из них страховка</div>
                <div class="metric-sub">{total_insurance_sum:,.1f} млн сум</div>
            </div>
            
            <div class="metric-card purple">
                <div class="metric-icon">💼</div>
                <div class="metric-value">{total_other:,}</div>
                <div class="metric-label">Прочие погашения</div>
                <div class="metric-sub">{total_closed_sum - total_insurance_sum:,.1f} млн сум</div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- LEGEND -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="legend-card">
            <div class="legend-title">📖 Описание показателей</div>
            <div class="legend-grid">
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['start']};"></div>
                    <div class="legend-text">
                        <strong>Начало 90+</strong>
                        <span>Анкеты с просрочкой более 90 дней на начало месяца (start_days > 90)</span>
                    </div>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['new']};"></div>
                    <div class="legend-text">
                        <strong>Новые 90+</strong>
                        <span>Анкеты, перешедшие порог 90 дней в течение месяца (start ≤ 90, max > 90)</span>
                    </div>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['closed']};"></div>
                    <div class="legend-text">
                        <strong>Погашено всего</strong>
                        <span>Анкеты 90+, полностью закрытые к концу месяца (max > 90, end = 0)</span>
                    </div>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['insurance']};"></div>
                    <div class="legend-text">
                        <strong>Страховка</strong>
                        <span>Погашения за счёт страхового возмещения (из файла страховки)</span>
                    </div>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: {COLORS['other']};"></div>
                    <div class="legend-text">
                        <strong>Прочие погашения</strong>
                        <span>Погашения без учёта страховки (собственные средства, реструктуризация и др.)</span>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- CHART 1: Main Bar Chart -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="chart-card">
            <div class="chart-title">📊 Динамика количества анкет по месяцам</div>
            <div id="chart1"></div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- CHART 2: Sum Chart -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="chart-card">
            <div class="chart-title">💰 Динамика сумм (миллионы)</div>
            <div id="chart2"></div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- CHARTS ROW: Waterfall + Pie -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="charts-row">
            <div class="chart-card">
                <div class="chart-title">🌊 Движение портфеля (Waterfall)</div>
                <div id="chart3"></div>
            </div>
            <div class="chart-card">
                <div class="chart-title">🎯 Структура погашений</div>
                <div id="chart4"></div>
            </div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- CHART 5: Insurance Stack -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="chart-card">
            <div class="chart-title">🛡️ Погашения: Страховка vs Прочие (по месяцам)</div>
            <div id="chart5"></div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- CHART 6: Trend -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="chart-card">
            <div class="chart-title">📈 Тренд: Новые vs Погашенные</div>
            <div id="chart6"></div>
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- TABLE -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="table-card">
            <div class="chart-title">📋 Сводная таблица</div>
            {self._create_html_table()}
        </div>
        
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <!-- FOOTER -->
        <!-- ═══════════════════════════════════════════════════════════════ -->
        <div class="footer">
            <p><strong>90+ Просрочка Анализатор v3.0</strong> | Данные за {period_start} — {period_end}</p>
        </div>
        
    </div>
    
    <script>
        const config = {{responsive: true, displayModeBar: true, displaylogo: false}};
        
        Plotly.newPlot('chart1', {chart1_json}.data, {chart1_json}.layout, config);
        Plotly.newPlot('chart2', {chart2_json}.data, {chart2_json}.layout, config);
        Plotly.newPlot('chart3', {chart3_json}.data, {chart3_json}.layout, config);
        Plotly.newPlot('chart4', {chart4_json}.data, {chart4_json}.layout, config);
        Plotly.newPlot('chart5', {chart5_json}.data, {chart5_json}.layout, config);
        Plotly.newPlot('chart6', {chart6_json}.data, {chart6_json}.layout, config);
    </script>
</body>
</html>'''
        
        filepath = os.path.join(self.output_path, "Dashboard_90plus.html")
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"   ✅ Сохранено: {filepath}")
        return filepath
    
    def _create_main_bar_chart(self):
        """Основной график"""
        fig = go.Figure()
        
        periods = self.df['period'].tolist()
        
        fig.add_trace(go.Bar(
            name='Начало 90+', x=periods, y=self.df['start_count'],
            marker_color=COLORS['start'], text=self.df['start_count'],
            textposition='auto', hovertemplate='%{x}<br>Начало: %{y:,}<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            name='Новые 90+', x=periods, y=self.df['new_count'],
            marker_color=COLORS['new'], text=self.df['new_count'],
            textposition='auto', hovertemplate='%{x}<br>Новые: %{y:,}<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            name='Страховка', x=periods, y=self.df['insurance_count'],
            marker_color=COLORS['insurance'], text=self.df['insurance_count'],
            textposition='auto', hovertemplate='%{x}<br>Страховка: %{y:,}<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            name='Прочие погашения', x=periods, y=self.df['other_closed_count'],
            marker_color=COLORS['other'], text=self.df['other_closed_count'],
            textposition='auto', hovertemplate='%{x}<br>Прочие: %{y:,}<extra></extra>'
        ))
        
        fig.update_layout(
            barmode='group',
            xaxis_tickangle=-45,
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
            margin=dict(l=60, r=40, t=60, b=100),
            height=500,
            hovermode='x unified',
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        fig.update_xaxes(gridcolor='#EDF2F7')
        fig.update_yaxes(gridcolor='#EDF2F7')
        
        return fig
    
    def _create_sum_chart(self):
        """График сумм"""
        fig = go.Figure()
        
        periods = self.df['period'].tolist()
        
        fig.add_trace(go.Scatter(
            name='Начало 90+', x=periods, y=self.df['start_sum'] / 1e6,
            mode='lines+markers', line=dict(color=COLORS['start'], width=3),
            marker=dict(size=10), hovertemplate='%{x}<br>Начало: %{y:,.1f} млн<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            name='Новые 90+', x=periods, y=self.df['new_sum'] / 1e6,
            mode='lines+markers', line=dict(color=COLORS['new'], width=3),
            marker=dict(size=10), hovertemplate='%{x}<br>Новые: %{y:,.1f} млн<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            name='Погашено всего', x=periods, y=self.df['closed_sum'] / 1e6,
            mode='lines+markers', line=dict(color=COLORS['closed'], width=3),
            marker=dict(size=10), hovertemplate='%{x}<br>Погашено: %{y:,.1f} млн<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            name='Страховка', x=periods, y=self.df['insurance_sum'] / 1e6,
            mode='lines+markers', line=dict(color=COLORS['insurance'], width=3, dash='dash'),
            marker=dict(size=10, symbol='diamond'),
            hovertemplate='%{x}<br>Страховка: %{y:,.1f} млн<extra></extra>'
        ))
        
        fig.update_layout(
            xaxis_tickangle=-45,
            yaxis_title='Сумма (млн)',
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
            margin=dict(l=60, r=40, t=60, b=100),
            height=500,
            hovermode='x unified',
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        fig.update_xaxes(gridcolor='#EDF2F7')
        fig.update_yaxes(gridcolor='#EDF2F7')
        
        return fig
    
    def _create_waterfall_chart(self):
        """Waterfall диаграмма"""
        total_start = self.df['start_count'].iloc[0] if len(self.df) > 0 else 0
        total_new = self.df['new_count'].sum()
        total_closed = self.df['closed_count'].sum()
        calculated_end = total_start + total_new - total_closed
        
        fig = go.Figure(go.Waterfall(
            orientation='v',
            measure=['absolute', 'relative', 'relative', 'total'],
            x=['Начало<br>периода', 'Новые<br>90+', 'Погашено', 'Расчётный<br>итог'],
            y=[total_start, total_new, -total_closed, calculated_end],
            text=[f'{total_start:,}', f'+{total_new:,}', f'-{total_closed:,}', f'{calculated_end:,}'],
            textposition='outside',
            textfont=dict(size=14, color='#1E3A5F'),
            connector={'line': {'color': '#1E3A5F', 'width': 2}},
            increasing={'marker': {'color': COLORS['new']}},
            decreasing={'marker': {'color': COLORS['closed']}},
            totals={'marker': {'color': COLORS['start']}}
        ))
        
        fig.update_layout(
            showlegend=False,
            margin=dict(l=40, r=40, t=40, b=60),
            height=400,
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        fig.update_yaxes(gridcolor='#EDF2F7')
        
        return fig
    
    def _create_pie_chart(self):
        """Круговая диаграмма"""
        total_insurance = self.df['insurance_count'].sum()
        total_other = self.df['other_closed_count'].sum()
        
        fig = go.Figure(data=[go.Pie(
            labels=['Страховка', 'Прочие'],
            values=[total_insurance, total_other],
            hole=0.5,
            marker_colors=[COLORS['insurance'], COLORS['other']],
            textinfo='label+percent+value',
            texttemplate='%{label}<br>%{value:,}<br>(%{percent})',
            textfont=dict(size=13),
            hovertemplate='<b>%{label}</b><br>Количество: %{value:,}<br>Доля: %{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            annotations=[dict(
                text=f'Всего<br><b>{total_insurance + total_other:,}</b>',
                x=0.5, y=0.5, font_size=16, showarrow=False
            )],
            margin=dict(l=20, r=20, t=40, b=20),
            height=400,
            paper_bgcolor='white'
        )
        
        return fig
    
    def _create_insurance_stack_chart(self):
        """Стэк страховки"""
        fig = go.Figure()
        
        periods = self.df['period'].tolist()
        
        fig.add_trace(go.Bar(
            name='Страховка', x=periods, y=self.df['insurance_sum'] / 1e6,
            marker_color=COLORS['insurance'],
            text=[f'{x/1e6:.1f}' for x in self.df['insurance_sum']],
            textposition='inside',
            hovertemplate='%{x}<br>Страховка: %{y:,.1f} млн<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            name='Прочие', x=periods, y=self.df['other_closed_sum'] / 1e6,
            marker_color=COLORS['other'],
            text=[f'{x/1e6:.1f}' for x in self.df['other_closed_sum']],
            textposition='inside',
            hovertemplate='%{x}<br>Прочие: %{y:,.1f} млн<extra></extra>'
        ))
        
        fig.update_layout(
            barmode='stack',
            xaxis_tickangle=-45,
            yaxis_title='Сумма (млн)',
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
            margin=dict(l=60, r=40, t=60, b=100),
            height=450,
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        fig.update_xaxes(gridcolor='#EDF2F7')
        fig.update_yaxes(gridcolor='#EDF2F7')
        
        return fig
    
    def _create_monthly_trend_chart(self):
        """Тренд: Новые vs Погашенные"""
        fig = go.Figure()
        
        periods = self.df['period'].tolist()
        
        # Новые (отрицательные для визуализации)
        fig.add_trace(go.Bar(
            name='Новые 90+ (приток)', x=periods, y=self.df['new_count'],
            marker_color=COLORS['new'], 
            hovertemplate='%{x}<br>Новые: +%{y:,}<extra></extra>'
        ))
        
        # Погашенные
        fig.add_trace(go.Bar(
            name='Погашено (отток)', x=periods, y=-self.df['closed_count'],
            marker_color=COLORS['closed'],
            hovertemplate='%{x}<br>Погашено: %{customdata:,}<extra></extra>',
            customdata=self.df['closed_count']
        ))
        
        # Линия баланса
        balance = self.df['new_count'] - self.df['closed_count']
        fig.add_trace(go.Scatter(
            name='Баланс (Новые - Погашено)', x=periods, y=balance.cumsum(),
            mode='lines+markers', line=dict(color=COLORS['primary'], width=3),
            marker=dict(size=8),
            hovertemplate='%{x}<br>Накопленный баланс: %{y:,}<extra></extra>'
        ))
        
        fig.update_layout(
            barmode='relative',
            xaxis_tickangle=-45,
            yaxis_title='Количество анкет',
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
            margin=dict(l=60, r=40, t=60, b=100),
            height=450,
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        fig.update_xaxes(gridcolor='#EDF2F7')
        fig.update_yaxes(gridcolor='#EDF2F7', zeroline=True, zerolinecolor='#1E3A5F', zerolinewidth=2)
        
        return fig
    
    def _create_html_table(self):
        """HTML таблица"""
        html = '<table><thead><tr>'
        
        columns = [
            ('period', 'Период'),
            ('start_count', 'Начало<br>(шт)'),
            ('start_sum', 'Начало<br>(сумма)'),
            ('new_count', 'Новые<br>(шт)'),
            ('new_sum', 'Новые<br>(сумма)'),
            ('closed_count', 'Погашено<br>(шт)'),
            ('closed_sum', 'Погашено<br>(сумма)'),
            ('insurance_count', 'Страховка<br>(шт)'),
            ('insurance_sum', 'Страховка<br>(сумма)'),
            ('other_closed_count', 'Прочие<br>(шт)'),
            ('other_closed_sum', 'Прочие<br>(сумма)')
        ]
        
        for _, header in columns:
            html += f'<th>{header}</th>'
        html += '</tr></thead><tbody>'
        
        for _, row in self.df.iterrows():
            html += '<tr>'
            for col, _ in columns:
                val = row[col]
                if 'sum' in col:
                    formatted = f'{val/1e6:,.2f} млн'
                elif col == 'period':
                    formatted = val
                else:
                    formatted = f'{int(val):,}'
                html += f'<td>{formatted}</td>'
            html += '</tr>'
        
        # Итого
        html += '<tr class="total-row"><td><strong>ИТОГО</strong></td>'
        for col, _ in columns[1:]:
            if col == 'start_count':
                val = self.df[col].iloc[0]
            elif col == 'start_sum':
                val = self.df[col].iloc[0]
            else:
                val = self.df[col].sum()
            
            if 'sum' in col:
                formatted = f'<strong>{val/1e6:,.2f} млн</strong>'
            else:
                formatted = f'<strong>{int(val):,}</strong>'
            html += f'<td>{formatted}</td>'
        
        html += '</tr></tbody></table>'
        return html
    
    def _generate_excel_report(self):
        """Excel отчет"""
        print("\n📊 Создание Excel...")
        
        filepath = os.path.join(self.output_path, "Отчет_90plus.xlsx")
        
        export_df = self.df.copy()
        export_df = export_df.rename(columns={
            'period': 'Период',
            'start_count': 'Начало 90+ (шт)',
            'start_sum': 'Начало 90+ (сумма)',
            'new_count': 'Новые 90+ (шт)',
            'new_sum': 'Новые 90+ (сумма)',
            'closed_count': 'Погашено всего (шт)',
            'closed_sum': 'Погашено всего (сумма)',
            'insurance_count': 'Страховка (шт)',
            'insurance_sum': 'Страховка (сумма)',
            'other_closed_count': 'Прочие (шт)',
            'other_closed_sum': 'Прочие (сумма)'
        })
        
        drop_cols = ['prefix', 'year', 'month_num', 'period_key', 'sort_key']
        export_df = export_df.drop(columns=[c for c in drop_cols if c in export_df.columns])
        
        export_df.to_excel(filepath, index=False, sheet_name='Данные')
        
        print(f"   ✅ Сохранено: {filepath}")
        return filepath


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("\n" + "="*70)
    print("   🏦 90+ ПРОСРОЧКА АНАЛИЗАТОР v3.0")
    print("   Профессиональный анализ кредитного портфеля")
    print("="*70)
    
    selector = FileSelector()
    
    if not selector.select_files():
        print("\n❌ Отменено")
        return
    
    try:
        analyzer = Prosrochka90Analyzer(
            data_files=selector.data_files,
            insurance_file=selector.insurance_file
        )
        
        analyzer.load_data()
        analyzer.analyze_all()
        
        reporter = InteractiveReportGenerator(analyzer, selector.output_path)
        html_path = reporter.generate_all()
        
        import webbrowser
        webbrowser.open(f'file://{os.path.abspath(html_path)}')
        
        print("\n" + "="*70)
        print("   ✅ ГОТОВО!")
        print("="*70)
        
    except Exception as e:
        print(f"\n❌ Ошибка: {str(e)}")
        import traceback
        traceback.print_exc()
        messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    main()
