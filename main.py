import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import math
from docx import Document
from datetime import datetime


class ITPCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Расчет оборудования ИТП с плавающим давлением пара")
        self.root.geometry("1000x950")

        # Значения по умолчанию для настроек
        self.settings = {
            'kvs_margin': 1.25,  # запас 25% для клапана
            'optimal_water_speed': 1.5,  # м/с - для внешних труб воды
            'optimal_steam_speed': 30,  # м/с - для внешних паропроводов
            'condensate_speed': 1.0,  # м/с - для внешних конденсатопроводов
            'density_water': 1000,  # кг/м³
            'cavitation_z': 0.6,  # коэффициент начала кавитации
            'system_volume_per_kw': 10,  # л/кВт для расчета объема системы
            'fill_time_hours': 5,  # время заполнения системы, часы
            'valve_port_water_speed': 3.0,  # макс скорость в портах клапана для воды, м/с
            'valve_port_steam_speed': 65.0,  # макс скорость в портах клапана для пара, м/с
            'safety_factor_KO': 2.0,  # Запас по расходу конденсата
            'to_surface_margin': 1.60,  # Запас площади ТО по умолчанию 60%
            'steam_k_average': 5000.0,  # Средний коэф. теплопередачи ТО
            'to_port_water_speed': 1.5,  # ДОБАВЛЕНО: Макс. скорость в порту ТО по ВОДЕ, м/с
            'to_port_steam_speed': 65.0,  # ДОБАВЛЕНО: Макс. скорость в порту ТО по ПАРУ, м/с
        }

        # Стандартные диаметры труб (Ду)
        self.standard_diameters = [15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200]

        # Внутренние диаметры для воды (с учетом толщины стенки)
        self.inner_diameters_water = {
            15: 16, 20: 21, 25: 27, 32: 36, 40: 42, 50: 54, 65: 68, 80: 81, 100: 106, 125: 132, 150: 156, 200: 208
        }

        # Внутренние диаметры для пара (с большей толщиной стенки)
        self.inner_diameters_steam = {
            15: 15, 20: 20, 25: 25, 32: 32, 40: 40, 50: 50, 65: 65, 80: 80, 100: 100, 125: 125, 150: 150, 200: 200
        }

        # Стандартный ряд Kvs регулирующих клапанов
        self.standard_kvs_valves = [1.0, 1.6, 2.5, 4.0, 6.3, 10.0, 16.0, 25.0, 40.0, 63.0, 100.0, 160.0, 250.0]

        # Обновленная база данных теплообменников ТО
        self.baza_to = [
            {"name": "Тип 2 (1.2 м²)", "F": 1.2, "d_water": 40, "d_steam": 40, "K_cond": 4000, "K_cool": 4000},
            {"name": "Тип 3 (2.0 м²)", "F": 2.0, "d_water": 50, "d_steam": 50, "K_cond": 6000, "K_cool": 6000},
            {"name": "Тип 5 (4.0 м²)", "F": 4.0, "d_water": 65, "d_steam": 65, "K_cond": 5000, "K_cool": 5000},
            {"name": "Тип 6 (5.3 м²)", "F": 5.3, "d_water": 80, "d_steam": 80, "K_cond": 4500, "K_cool": 4500},
            {"name": "Тип 9 (10.7 м²)", "F": 10.7, "d_water": 100, "d_steam": 100, "K_cond": 4000, "K_cool": 4000},
            {"name": "Тип 12 (18.4 м²)", "F": 18.4, "d_water": 125, "d_steam": 125, "K_cond": 5000, "K_cool": 5000},
        ]

        # База конденсатоотводчиков КО (расход в кг/ч при перепаде ровно 1 бар)
        self.baza_ko = [
            {"name": "КО Ду15 (Кв=0.40)", "G_1bar": 400},
            {"name": "КО Ду20 (Кв=0.85)", "G_1bar": 850},
            {"name": "КО Ду25 (Кв=1.60)", "G_1bar": 1600},
            {"name": "КО Ду40 (Кв=3.50)", "G_1bar": 3500},
            {"name": "КО Ду50 (Кв=6.00)", "G_1bar": 6000},
        ]

        self.setup_ui()

    def get_water_vapor_pressure(self, t_celsius):
        if t_celsius < 40 or t_celsius > 150:
            return None
        A, B, C = (8.07131, 1730.63, 233.426) if t_celsius <= 100 else (8.14019, 1810.94, 244.485)
        log_p = A - (B / (C + t_celsius))
        return (10 ** log_p) * 0.00133322

    def get_antoine_temp(self, p_bar_abs):
        p_mmhg = p_bar_abs * 750.062
        if p_mmhg <= 0: return 0
        A, B, C = (8.07131, 1730.63, 233.426) if p_mmhg <= 760 else (8.14019, 1810.94, 244.485)
        try:
            return round(B / (A - math.log10(p_mmhg)) - C, 2)
        except ValueError:
            return 0

    def get_latent_heat_steam(self, p_abs):
        if p_abs <= 0: return 2200
        return 2280 * (p_abs ** -0.058)

    def get_water_enthalpy(self, t_celsius):
        return 0.0008 * t_celsius ** 2 + 4.02 * t_celsius + 5.5

    def log_mean_temp_diff(self, T_in, T_out, t_in, t_out):
        dt_max = max(T_in - t_out, 0.001)
        dt_min = max(T_out - t_in, 0.001)
        if abs(dt_max - dt_min) < 0.1:
            return (dt_max + dt_min) / 2
        return (dt_max - dt_min) / math.log(dt_max / dt_min)

    def setup_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)

        self.water_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.water_tab, text="Вода")
        self.setup_water_tab()

        self.steam_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.steam_tab, text="Пар")
        self.setup_steam_tab()

        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="Настройки")
        self.setup_settings_tab()

    def setup_water_tab(self):
        input_frame = ttk.LabelFrame(self.water_tab, text="Входные данные", padding=10)
        input_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(input_frame, text="Нагрузка:").grid(row=0, column=0, sticky='w', pady=2)
        self.water_load = ttk.Entry(input_frame, width=15)
        self.water_load.grid(row=0, column=1, pady=2)

        self.load_unit = tk.StringVar(value="kW")
        ttk.Radiobutton(input_frame, text="кВт", variable=self.load_unit, value="kW").grid(row=0, column=2)
        ttk.Radiobutton(input_frame, text="Мкал/ч", variable=self.load_unit, value="Mcal").grid(row=0, column=3)

        ttk.Label(input_frame, text="Температура подачи греющего контура (°C):").grid(row=1, column=0, sticky='w', pady=2)
        self.temp_supply_heat = ttk.Entry(input_frame, width=15)
        self.temp_supply_heat.grid(row=1, column=1, pady=2)

        ttk.Label(input_frame, text="Температура обратки греющего контура (°C):").grid(row=2, column=0, sticky='w', pady=2)
        self.temp_return_heat = ttk.Entry(input_frame, width=15)
        self.temp_return_heat.grid(row=2, column=1, pady=2)

        ttk.Label(input_frame, text="Давление в подающем трубопроводе (бар):").grid(row=3, column=0, sticky='w', pady=2)
        self.pressure_inlet = ttk.Entry(input_frame, width=15)
        self.pressure_inlet.grid(row=3, column=1, pady=2)

        ttk.Label(input_frame, text="Давление в обратном трубопроводе (бар):").grid(row=4, column=0, sticky='w', pady=2)
        self.pressure_return = ttk.Entry(input_frame, width=15)
        self.pressure_return.grid(row=4, column=1, pady=2)

        ttk.Label(input_frame, text="Перепад давления для рег. клапана (бар):").grid(row=5, column=0, sticky='w', pady=2)
        self.pressure_drop_valve = ttk.Entry(input_frame, width=15)
        self.pressure_drop_valve.grid(row=5, column=1, pady=2)

        ttk.Label(input_frame, text="Перепад давления для рег. перепада давления (бар):").grid(row=6, column=0, sticky='w', pady=2)
        self.pressure_drop_pd = ttk.Entry(input_frame, width=15)
        self.pressure_drop_pd.grid(row=6, column=1, pady=2)

        ttk.Label(input_frame, text="Тип теплоносителя:").grid(row=7, column=0, sticky='w', pady=2)
        self.coolant_type = tk.StringVar(value="water")
        ttk.Radiobutton(input_frame, text="Вода", variable=self.coolant_type, value="water", command=self.toggle_glycol).grid(row=7, column=1, sticky='w')
        ttk.Radiobutton(input_frame, text="Гликоль", variable=self.coolant_type, value="glycol", command=self.toggle_glycol).grid(row=7, column=2, sticky='w')

        ttk.Label(input_frame, text="Концентрация гликоля (%):").grid(row=8, column=0, sticky='w', pady=2)
        self.glycol_percent = ttk.Entry(input_frame, width=15, state='disabled')
        self.glycol_percent.grid(row=8, column=1, pady=2)
        self.glycol_percent.insert(0, "30")

        baka_frame = ttk.LabelFrame(self.water_tab, text="Данные для расширительного бака (необязательно)", padding=10)
        baka_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(baka_frame, text="Высота системы отопления (м):").grid(row=0, column=0, sticky='w', pady=2)
        self.system_height = ttk.Entry(baka_frame, width=15)
        self.system_height.grid(row=0, column=1, pady=2)

        ttk.Label(baka_frame, text="Максимальное давление в системе (бар):").grid(row=1, column=0, sticky='w', pady=2)
        self.max_pressure = ttk.Entry(baka_frame, width=15)
        self.max_pressure.grid(row=1, column=1, pady=2)

        ttk.Label(baka_frame, text="Температура подачи нагреваемого контура (°C):").grid(row=2, column=0, sticky='w', pady=2)
        self.temp_supply_heated = ttk.Entry(baka_frame, width=15)
        self.temp_supply_heated.grid(row=2, column=1, pady=2)

        ttk.Label(baka_frame, text="Температура обратки нагреваемого контура (°C):").grid(row=3, column=0, sticky='w', pady=2)
        self.temp_return_heated = ttk.Entry(baka_frame, width=15)
        self.temp_return_heated.grid(row=3, column=1, pady=2)

        button_frame = ttk.Frame(self.water_tab)
        button_frame.pack(fill='x', padx=10, pady=10)

        ttk.Button(button_frame, text="Рассчитать", command=self.calculate_water).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Вывести отчет Word", command=self.generate_water_report).pack(side='left', padx=5)

        self.water_results = tk.Text(self.water_tab, height=15, width=80)
        self.water_results.pack(fill='both', expand=True, padx=10, pady=5)

    def toggle_glycol(self):
        if self.coolant_type.get() == "glycol":
            self.glycol_percent.config(state='normal')
        else:
            self.glycol_percent.config(state='disabled')

    def setup_steam_tab(self):
        input_frame = ttk.LabelFrame(self.steam_tab, text="Входные данные греющего пара и нагреваемой воды", padding=10)
        input_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(input_frame, text="Давление пара в магистрали (бар изб):").grid(row=0, column=0, sticky='w', pady=2)
        self.steam_pressure = ttk.Entry(input_frame, width=15)
        self.steam_pressure.grid(row=0, column=1, pady=2)

        ttk.Label(input_frame, text="Противодавление в линии конденсата (бар изб):").grid(row=1, column=0, sticky='w',
                                                                                          pady=2)
        self.cond_pressure_after = ttk.Entry(input_frame, width=15)
        self.cond_pressure_after.grid(row=1, column=1, pady=2)

        ttk.Label(input_frame, text="Целевая темп. конденсата на выходе (°C):").grid(row=2, column=0, sticky='w',
                                                                                     pady=2)
        self.cond_temp = ttk.Entry(input_frame, width=15)
        self.cond_temp.grid(row=2, column=1, pady=2)
        self.cond_temp.insert(0, "90")

        ttk.Label(input_frame, text="Расход нагреваемой воды (м³/ч):").grid(row=3, column=0, sticky='w', pady=2)
        self.steam_water_flow = ttk.Entry(input_frame, width=15)
        self.steam_water_flow.grid(row=3, column=1, pady=2)

        ttk.Label(input_frame, text="Температура воды на входе (°C):").grid(row=4, column=0, sticky='w', pady=2)
        self.steam_water_t_in = ttk.Entry(input_frame, width=15)
        self.steam_water_t_in.grid(row=4, column=1, pady=2)

        ttk.Label(input_frame, text="Температура воды на выходе (°C):").grid(row=5, column=0, sticky='w', pady=2)
        self.steam_water_t_out = ttk.Entry(input_frame, width=15)
        self.steam_water_t_out.grid(row=5, column=1, pady=2)

        # ДОБАВЛЕНО: Выпадающий список выбора типа ТО
        ttk.Label(input_frame, text="Режим выбора теплообменника (ТО):").grid(row=6, column=0, sticky='w', pady=2)
        self.to_select_mode = ttk.Combobox(input_frame, width=15, state="readonly")
        self.to_select_mode['values'] = ("Автоподбор", "Тип 2", "Тип 3", "Тип 5", "Тип 6", "Тип 9", "Тип 12")
        self.to_select_mode.current(0)
        self.to_select_mode.grid(row=6, column=1, pady=2)

        button_frame = ttk.Frame(self.steam_tab)
        button_frame.pack(fill='x', padx=10, pady=10)

        ttk.Button(button_frame, text="Рассчитать систему пара", command=self.calculate_steam).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Вывести отчет Word", command=self.generate_steam_report).pack(side='left',
                                                                                                     padx=5)

        self.steam_results = tk.Text(self.steam_tab, height=15, width=80)
        self.steam_results.pack(fill='both', expand=True, padx=10, pady=5)

    def setup_settings_tab(self):
        settings_frame = ttk.LabelFrame(self.settings_tab, text="Настройки параметров", padding=10)
        settings_frame.pack(fill='both', expand=True, padx=10, pady=10)

        settings_items = [
            ("Запас для Kvs клапана (коэффициент):", "kvs_margin"),
            ("Оптимальная скорость воды в трубопроводах (м/с):", "optimal_water_speed"),
            ("Оптимальная скорость пара в паропроводах (м/с):", "optimal_steam_speed"),
            ("Скорость конденсата в конденсатопроводах (м/с):", "condensate_speed"),
            ("Коэффициент начала кавитации Z:", "cavitation_z"),
            ("Удельный объем системы (л/кВт):", "system_volume_per_kw"),
            ("Время заполнения системы (ч):", "fill_time_hours"),
            ("Макс. скорость в портах клапана для ВОДЫ (м/с):", "valve_port_water_speed"),
            ("Макс. скорость в портах клапана для ПАРА (м/с):", "valve_port_steam_speed"),
            ("Запас по расходу конденсата (раз):", "safety_factor_KO"),
            ("Запас площади поверхности ТО (%):", "to_surface_margin"),
            ("Средний коэф. теплопередачи ТО (Вт/м²К):", "steam_k_average"),
            ("Макс. скорость в ПОРТУ ТО по ВОДЕ (м/с):", "to_port_water_speed"),
            ("Макс. скорость в ПОРТУ ТО по ПАРУ (м/с):", "to_port_steam_speed"),
        ]

        self.settings_entries = {}
        for i, (label, key) in enumerate(settings_items):
            ttk.Label(settings_frame, text=label).grid(row=i, column=0, sticky='w', pady=5)
            entry = ttk.Entry(settings_frame, width=15)
            entry.grid(row=i, column=1, pady=5, padx=10)

            if key == "to_surface_margin":
                display_val = (self.settings[key] - 1.0) * 100.0
            else:
                display_val = self.settings[key]

            entry.insert(0, f"{display_val:.1f}" if isinstance(display_val, float) else str(display_val))
            self.settings_entries[key] = entry

        ttk.Button(settings_frame, text="Сохранить настройки", command=self.save_settings).grid(row=len(settings_items),
                                                                                                column=0, columnspan=2,
                                                                                                pady=20)

    def save_settings(self):
        try:
            for key, entry in self.settings_entries.items():
                val = float(entry.get())
                if key == "to_surface_margin":
                    self.settings[key] = 1.0 + (val / 100.0)
                else:
                    self.settings[key] = val
            messagebox.showinfo("Успех", "Настройки сохранены")
        except ValueError:
            messagebox.showerror("Ошибка", "Пожалуйста, введите корректные числовые значения")

    def safe_float(self, value, default=None):
        try:
            return float(value) if value and value.strip() else default
        except (ValueError, AttributeError):
            return default

    def check_cavitation(self, dp, p_inlet, t_water):
        if p_inlet is None or t_water is None:
            return None, None
        p_vapor = self.get_water_vapor_pressure(t_water)
        if p_vapor is None:
            return None, "Температура вне диапазона"
        Z = self.settings['cavitation_z']
        dp_max = Z * (p_inlet - p_vapor)
        if dp_max <= 0:
            return dp_max, "Давление на входе слишком низкое"
        return dp_max, "ПРЕВЫШЕНИЕ! Риск кавитации" if dp > dp_max else "OK - кавитации нет"

    def select_diameter_by_flow_water(self, flow_rate, max_speed, use_steam_diameters=False):
        if flow_rate is None or max_speed <= 0:
            return self.standard_diameters
        for d in self.standard_diameters:
            inner_d = self.inner_diameters_steam.get(d, d) if use_steam_diameters else self.inner_diameters_water.get(d, d)
            area = math.pi * (inner_d / 2000) ** 2
            if (flow_rate / (1000 * area)) <= max_speed:
                return d
        return self.standard_diameters[-1]

    def select_diameter_by_flow_steam(self, flow_rate_kgph, specific_volume, max_speed):
        if flow_rate_kgph is None or max_speed <= 0:
            return self.standard_diameters
        flow_rate_m3s = (flow_rate_kgph / 3600) * specific_volume
        for d in self.standard_diameters:
            inner_d = self.inner_diameters_steam.get(d, d)
            area = math.pi * (inner_d / 2000) ** 2
            if (flow_rate_m3s / area) <= max_speed:
                return d
        return self.standard_diameters[-1]

    def calculate_water_speed(self, flow_rate, diameter_du, use_steam_diameters=False):
        if not flow_rate or not diameter_du: return 0
        inner_d = self.inner_diameters_steam.get(diameter_du, diameter_du) if use_steam_diameters else self.inner_diameters_water.get(diameter_du, diameter_du)
        return flow_rate / (1000 * math.pi * (inner_d / 2000) ** 2)

    def calculate_steam_speed(self, flow_rate_kgph, specific_volume, diameter_du):
        if not flow_rate_kgph or not specific_volume or not diameter_du: return 0
        return (flow_rate_kgph / 3600 * specific_volume) / (math.pi * (self.inner_diameters_steam.get(diameter_du, diameter_du) / 2000) ** 2)

    def calculate_water(self):
        try:
            results = []
            calc_data = {}
            load = self.safe_float(self.water_load.get())
            t_supply_heat = self.safe_float(self.temp_supply_heat.get())
            t_return_heat = self.safe_float(self.temp_return_heat.get())
            p_inlet = self.safe_float(self.pressure_inlet.get())
            p_return = self.safe_float(self.pressure_return.get())
            dp_valve = self.safe_float(self.pressure_drop_valve.get())
            dp_pd = self.safe_float(self.pressure_drop_pd.get())
            h_system = self.safe_float(self.system_height.get())
            p_max = self.safe_float(self.max_pressure.get())
            t_supply_heated = self.safe_float(self.temp_supply_heated.get())
            t_return_heated = self.safe_float(self.temp_return_heated.get())
            glycol_percent = self.safe_float(self.glycol_percent.get(), 0) if self.coolant_type.get() == "glycol" else 0

            results.append("РЕЗУЛЬТАТЫ РАСЧЕТА (ВОДА)")
            results.append("=" * 50)
            calc_data['input'] = {'load': load, 'load_unit': self.load_unit.get(), 't_supply_heat': t_supply_heat,
                                  't_return_heat': t_return_heat, 'p_inlet': p_inlet, 'p_return': p_return,
                                  'dp_valve': dp_valve, 'dp_pd': dp_pd, 'coolant_type': self.coolant_type.get(),
                                  'glycol_percent': glycol_percent, 'h_system': h_system, 'p_max': p_max,
                                  't_supply_heated': t_supply_heated, 't_return_heated': t_return_heated}

            if load is not None and t_supply_heat is not None and t_return_heat is not None:
                if self.load_unit.get() == "Mcal": load *= 1.163
                dt = t_supply_heat - t_return_heat
                if dt > 0:
                    G = load * 1000 / (4186 * dt)
                    G_m3h = G * 3.6
                    results.append(f"\nРасход воды в греющем контуре: {G_m3h:.2f} м³/ч")
                    calc_data['G'] = G
                    calc_data['G_m3h'] = G_m3h
                else:
                    results.append("\nОшибка: Разность температур должна быть положительной")

            self.water_results.delete(1.0, tk.END)
            self.water_results.insert(1.0, "\n".join(results))
            self.water_calc_results = calc_data
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def calculate_steam(self):
        try:
            results = []
            calc_data = {}
            P_mag_g = self.safe_float(self.steam_pressure.get())
            P_prot_g = self.safe_float(self.cond_pressure_after.get())
            V_water_m3h = self.safe_float(self.steam_water_flow.get())
            t_w1 = self.safe_float(self.steam_water_t_in.get())
            t_w2 = self.safe_float(self.steam_water_t_out.get())
            selected_mode = self.to_select_mode.get()

            if None in [P_mag_g, P_prot_g, V_water_m3h, t_w1, t_w2]:
                messagebox.showwarning("Предупреждение", "Заполните все входные поля на вкладке пара!")
                return

            G_water = V_water_m3h * self.settings['density_water']
            Q = G_water * 4187 * (t_w2 - t_w1) / 3600 / 1000

            results.append("РЕЗУЛЬТАТЫ РАСЧЕТА ПАРОВОДЯНОЙ СИСТЕМЫ")
            results.append(f"Тепловая нагрузка: {Q:.1f} кВт (Расход воды: {V_water_m3h:.2f} м³/ч)")
            results.append(f"Режим подбора ТО: {selected_mode}")
            results.append("=" * 60)

            K_average = self.settings['steam_k_average']
            selected_to = None
            final_P_to_g, final_G_steam, final_t_sat, final_F_required = 0.0, 0.0, 0.0, 0.0

            fail_reason = "по площади поверхности"
            v_water_err, v_steam_err = 0.0, 0.0

            for to in self.baza_to:
                if selected_mode != "Автоподбор" and selected_mode not in to["name"]:
                    continue

                P_to_g = P_prot_g + 1.0
                P_to_max_g = P_mag_g - 1.0
                if P_to_g > P_to_max_g: P_to_g = P_to_max_g

                while P_to_g <= P_to_max_g:
                    P_to_abs = P_to_g + 1.013
                    t_sat = self.get_antoine_temp(P_to_abs)
                    if t_sat <= t_w2:
                        P_to_g += 0.5
                        continue

                    r_steam = self.get_latent_heat_steam(P_to_abs)
                    G_steam = (Q * 3600) / r_steam
                    dt_ln = self.log_mean_temp_diff(t_sat, t_sat, t_w1, t_w2)
                    if dt_ln <= 0:
                        P_to_g += 0.5
                        continue
                    F_required = (Q * 1000) / (K_average * dt_ln)

                    area_ok = False
                    if selected_mode == "Автоподбор":
                        if to["F"] >= F_required * self.settings['to_surface_margin']:
                            area_ok = True
                    else:
                        if to["F"] >= F_required:
                            area_ok = True

                    if area_ok:
                        v_water_patrubok = (G_water / 1000 / 3600) / (math.pi * (to["d_water"] / 1000) ** 2 / 4)
                        v_specific_to = 1.685 / (P_to_abs ** 0.942)
                        v_steam_patrubok = (G_steam / (1.0 / v_specific_to) / 3600) / (
                                    math.pi * (to["d_steam"] / 1000) ** 2 / 4)

                        if selected_mode != "Автоподбор":
                            if v_water_patrubok > self.settings['to_port_water_speed']:
                                fail_reason = "по скорости ВОДЫ в патрубке"
                                v_water_err = v_water_patrubok
                            elif v_steam_patrubok > self.settings['to_port_steam_speed']:
                                fail_reason = "по скорости ПАРА в патрубке"
                                v_steam_err = v_steam_patrubok

                        if v_water_patrubok <= self.settings['to_port_water_speed'] and v_steam_patrubok <= \
                                self.settings['to_port_steam_speed']:
                            selected_to = to
                            break
                    P_to_g += 0.5

                if selected_to is not None:
                    final_P_to_g = P_to_g
                    final_G_steam = G_steam
                    final_t_sat = t_sat
                    final_F_required = F_required
                    break
            # АНАЛИЗ И ВЫВОД ИТОГОВЫХ ДАННЫХ
            if selected_to is None:
                if selected_mode == "Автоподбор":
                    results.append(
                        f"\n[!] Подходящий ТО с заданным запасом площади >= {(self.settings['to_surface_margin'] - 1) * 100:.0f}% не найден.")
                else:
                    results.append(f"\n[ОТКАЗ] Выбранный теплообменник {selected_mode} НЕ ПРОХОДИТ!")
                    if "ВОДЫ" in fail_reason:
                        results.append(f"Причина: Превышена макс. скорость ВОДЫ в штуцере ТО.")
                        results.append(
                            f"  • Фактическая скорость: {v_water_err:.2f} м/с | Лимит настроек: {self.settings['to_port_water_speed']:.2f} м/с")
                    elif "ПАРА" in fail_reason:
                        results.append(f"Причина: Превышена макс. скорость ПАРА в штуцере ТО.")
                        results.append(
                            f"  • Фактическая скорость: {v_steam_err:.2f} м/с | Лимит настроек: {self.settings['to_port_steam_speed']:.2f} м/с")
                    else:
                        results.append(f"Причина: Не хватает площади поверхности теплообмена.")
            else:
                dp_valve = P_mag_g - final_P_to_g
                if dp_valve <= 0: dp_valve = 0.05
                term = 0.42 - (dp_valve / (P_mag_g + 1.013))
                under_root = 1 - 5.67 * (term ** 2) if term > 0 else 1.0
                Kv_valve_req = final_G_steam / (12 * (P_mag_g + 1.013) * math.sqrt(under_root)) * self.settings[
                    'kvs_margin']

                kvs_selected = self.standard_kvs_valves[-1]
                for kvs in self.standard_kvs_valves:
                    if kvs >= Kv_valve_req:
                        kvs_selected = kvs
                        break

                dp_ko_final = final_P_to_g - P_prot_g
                if dp_ko_final < 0.05: dp_ko_final = 0.05
                G_ko_req_at_1bar = (final_G_steam * self.settings['safety_factor_KO']) / math.sqrt(dp_ko_final)

                ko_selected_name = ""
                ko_diameter_for_speed = 15
                ko_count = 1
                base_G_1bar = 400

                # 1. Сначала ищем одиночный КО в базе
                for ko in self.baza_ko:
                    if ko["G_1bar"] >= G_ko_req_at_1bar:
                        ko_selected_name = ko["name"]
                        base_G_1bar = ko["G_1bar"]
                        ko_count = 1
                        # ИСПРАВЛЕНО: Безопасный и надежный способ вырезать цифру Ду из строки "КО Ду25 (Кв=1.60)"
                        try:
                            # Отрезаем всё до "Ду", берем остаток, делим по пробелу и берем первое слово-число
                            ko_diameter_for_speed = int(ko["name"].split("Ду")[1].split()[0])
                        except Exception:
                            ko_diameter_for_speed = 15
                        break

                # 2. Если Ду50 мало, переходим к параллельным каскадам
                if not ko_selected_name:
                    if (3500 * 2) >= G_ko_req_at_1bar:
                        ko_selected_name, ko_diameter_for_speed, base_G_1bar, ko_count = "2 шт. КО Ду40 параллельно", 40, 3500, 2
                    elif (6000 * 2) >= G_ko_req_at_1bar:
                        ko_selected_name, ko_diameter_for_speed, base_G_1bar, ko_count = "2 шт. КО Ду50 параллельно", 50, 6000, 2
                    else:
                        ko_selected_name, ko_diameter_for_speed, base_G_1bar, ko_count = "3 шт. КО Ду50 параллельно", 50, 6000, 3

                if P_mag_g <= 4.5:
                    pmo_selected = "4.5"
                elif P_mag_g <= 10.0:
                    pmo_selected = "10"
                else:
                    pmo_selected = "14"

                actual_ko_max_capacity = base_G_1bar * math.sqrt(dp_ko_final) * ko_count
                v_specific_inlet = 1.685 / ((P_mag_g + 1.013) ** 0.942)
                d_steam_line = self.select_diameter_by_flow_steam(final_G_steam, v_specific_inlet,
                                                                  self.settings['optimal_steam_speed'])
                v_steam_actual = self.calculate_steam_speed(final_G_steam, v_specific_inlet, d_steam_line)

                G_cond_total_kgps = (final_G_steam * self.settings['safety_factor_KO']) / 3600.0
                d_cond_line = self.select_diameter_by_flow_water(G_cond_total_kgps, self.settings['condensate_speed'])
                v_cond_actual = self.calculate_water_speed(G_cond_total_kgps, d_cond_line)

                # Теперь здесь гарантированно посчитается скорость под Ду 25 (или Ду подобранного КО)
                v_cond_at_ko_diameter = self.calculate_water_speed(G_cond_total_kgps, ko_diameter_for_speed)

                if final_F_required > 0.005:
                    fact_margin_percent = ((selected_to["F"] - final_F_required) / final_F_required) * 100.0
                else:
                    fact_margin_percent = 0.0

                results.append(f"\n[УСПЕХ] Комплект оборудования рассчитан:")
                results.append(f"  • ВЫБРАННЫЙ ТЕПЛООБМЕННИК: {selected_to['name']}")

                if selected_mode != "Автоподбор":
                    target_margin_percent = (self.settings['to_surface_margin'] - 1.0) * 100.0
                    if fact_margin_percent < target_margin_percent:
                        results.append(
                            f"    ⚠️ ВНИМАНИЕ: Запас площади ({fact_margin_percent:.1f}%) НИЖЕ заданного в настройках ({target_margin_percent:.0f}%)!")
                    else:
                        results.append(
                            f"    ОК: Фактический запас площади ({fact_margin_percent:.1f}%) соответствует норме.")
                else:
                    results.append(f"    - Фактический запас площади: {fact_margin_percent:.1f}%")

                results.append(f"  • Давление пара внутри ТО: {final_P_to_g:.2f} бар изб (Тнас = {final_t_sat:.1f} °C)")
                results.append(f"  • Расход греющего пара: {final_G_steam:.1f} кг/ч")
                results.append(
                    f"  • Регулирующий клапан: требуется Kv={Kv_valve_req:.2f} -> Выбран Kvs = {kvs_selected}")
                results.append(
                    f"  • Конденсатоотводчик: {ko_selected_name} (Исполнение по макс. перепаду ΔP: {pmo_selected} бар)")
                results.append(
                    f"    - Требуемый расход (с запасом {self.settings['safety_factor_KO']}х): {final_G_steam * self.settings['safety_factor_KO']:.1f} кг/ч")
                results.append(f"    - Макс. расход при dP={dp_ko_final:.2f} бар: {actual_ko_max_capacity:.1f} кг/ч")
                results.append(
                    f"  • Входной паропровод (до клапана): Ду {d_steam_line} мм (Скорость: {v_steam_actual:.2f} м/с)")
                results.append(
                    f"  • Расчетный общий конденсатопровод: Ду {d_cond_line} мм (Скорость: {v_cond_actual:.2f} м/с)")
                results.append(f"  • ЕСЛИ ПРИНЯТЬ ДИАМЕТР ТРУБЫ ПО ДИАМЕТРУ КО (Ду {ko_diameter_for_speed} мм):")
                results.append(f"    -> Скорость на пиковый расход составит: {v_cond_at_ko_diameter:.2f} м/с")
                results.append("-" * 60)

                # Расчет скоростей сред в штуцерах (портах) самого ТО
                v_water_ports = (G_water / 1000 / 3600) / (math.pi * (selected_to["d_water"] / 1000) ** 2 / 4)
                v_steam_ports = (final_G_steam / (1.0 / v_specific_inlet) / 3600) / (
                            math.pi * (selected_to["d_steam"] / 1000) ** 2 / 4)

                # ОПРЕДЕЛЕНИЕ РЕЖИМА РАБОТЫ КЛАПАНА (Критический / Докритический)
                p_mag_abs = P_mag_g + 1.013
                p_to_abs = final_P_to_g + 1.013
                if p_to_abs <= 0.54 * p_mag_abs:
                    valve_regime = "Критический (звуковое истечение)"
                else:
                    valve_regime = "Докритический (плавное регулирование)"

                # ИСПРАВЛЕНО: Принудительно пишем правильную структуру входных данных для Word
                calc_data['input'] = {
                    'load': Q,
                    'p_steam': P_mag_g,
                    'cond_pressure_after': P_prot_g,
                    'cond_temp': self.safe_float(self.cond_temp.get(), 90.0)
                }

                calc_data.update({
                    'G_steam': final_G_steam,
                    'G_condensate': final_G_steam * self.settings['safety_factor_KO'],
                    'p_steam_after_valve': final_P_to_g,
                    'kvs_steam': kvs_selected,
                    'trap_diameter': f"{ko_selected_name} (ΔP={pmo_selected} бар)",
                    'trap_count': ko_count,
                    'trap_capacity': actual_ko_max_capacity,
                    'dp_trap': dp_ko_final,
                    'diameter_steam': d_steam_line,
                    'diameter_cond': d_cond_line,
                    'flash_steam': 0,
                    'selected_to_name': selected_to['name'],
                    'dp_valve': dp_valve,
                    'valve_regime': valve_regime,
                    'v_water_to_ports': v_water_ports,
                    'v_steam_to_ports': v_steam_ports,
                    'v_steam_line_speed': v_steam_actual,
                    'v_cond_line_speed': v_cond_actual,
                    'fact_margin_percent': fact_margin_percent
                })

            self.steam_results.delete(1.0, tk.END)
            self.steam_results.insert(1.0, "\n".join(results))
            self.steam_calc_results = calc_data

        except Exception as e:
            messagebox.showerror("Ошибка в алгоритме пара", str(e))

    def generate_water_report(self):
        if not hasattr(self, 'water_calc_results') or not self.water_calc_results:
            messagebox.showwarning("Предупреждение", "Сначала выполните расчет")
            return
        try:
            filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")],
                                                    title="Сохранить отчет как")
            if not filename: return
            doc = Document()
            doc.add_heading('Расчет оборудования ИТП (Водяной контур)', 0)
            doc.add_paragraph(f'Дата расчета: {datetime.now().strftime("%d.%m.%Y %H:%M")}')
            doc.add_heading('Входные данные', level=1)
            doc.save(filename)
            messagebox.showinfo("Успех", f"Отчет сохранен:\n{filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def generate_steam_report(self):
        if not hasattr(self, 'steam_calc_results') or not self.steam_calc_results:
            messagebox.showwarning("Предупреждение", "Сначала выполните расчет")
            return
        try:
            filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")],
                                                    title="Сохранить отчет как")
            if not filename: return
            doc = Document()

            # Заголовок документа
            doc.add_heading('ТЕХНИЧЕСКИЙ ОТЧЕТ РАСЧЕТА ОБОРУДОВАНИЯ ИТП', 0)
            doc.add_paragraph(f'Дата формирования отчета: {datetime.now().strftime("%d.%m.%Y %H:%M")}')

            res = self.steam_calc_results
            inp = res.get('input', {})

            # Раздел 1. Входные данные (ИСПРАВЛЕНО: Ключи приведены в строгое соответствие с расчетом)
            doc.add_heading('1. Исходные параметры технологического режима', level=1)
            doc.add_paragraph(f"• Расчетная тепловая нагрузка системы: {inp.get('load', 0):.1f} кВт")
            doc.add_paragraph(f"• Давление насыщенного пара в магистрали: {inp.get('p_steam', 0):.2f} бар изб.")
            doc.add_paragraph(
                f"• Противодавление в линии сброса конденсата: {inp.get('cond_pressure_after', 0):.2f} бар изб.")
            doc.add_paragraph(f"• Заданная температура охлаждения конденсата: {inp.get('cond_temp', 0):.1f} °C")

            # Раздел 2. Параметры теплообменного аппарата (ТО)
            doc.add_heading('2. Параметры теплообменного аппарата (ТО)', level=1)
            doc.add_paragraph(f"• Рекомендуемый/Поверяемый тип аппарата: {res.get('selected_to_name', '-')}")
            doc.add_paragraph(
                f"• Истинное рабочее давление пара внутри ТО: {res.get('p_steam_after_valve', 0):.2f} бар изб.")
            doc.add_paragraph(
                f"• Фактический ориентировочный запас площади поверхности: {res.get('fact_margin_percent', 0):.1f}%")
            doc.add_paragraph(
                f"• Скорость теплоносителя (воды) в патрубках ТО: {res.get('v_water_to_ports', 0):.2f} м/с")
            doc.add_paragraph(f"• Скорость греющего пара in патрубках ТО: {res.get('v_steam_to_ports', 0):.2f} м/с")
            # Раздел 3. Результаты по клапану
            doc.add_heading('3. Характеристики регулирующего клапана пара', level=1)
            doc.add_paragraph(
                f"• Подобранный коэффициент пропускной способности клапана: Kvs = {res.get('kvs_steam', 0)}")
            doc.add_paragraph(f"• Фактический перепад давления на клапане: ΔP = {res.get('dp_valve', 0):.2f} бар")
            doc.add_paragraph(f"• Гидравлический режим работы клапана: {res.get('valve_regime', '-')}")

            # Раздел 4. Результаты по КО и трубам
            doc.add_heading('4. Конденсатоотводчик и внешние трубопроводы ИТП', level=1)
            doc.add_paragraph(f"• Спецификация узла отвода конденсата: {res.get('trap_diameter', '-')}")
            doc.add_paragraph(f"• Количество КО в параллели: {res.get('trap_count', 1)} шт.")
            doc.add_paragraph(
                f"• Макс. пропускная способность группы КО при рабочем dP: {res.get('trap_capacity', 0):.1f} кг/ч")
            doc.add_paragraph(f"• Рабочий располагаемый перепад давления на КО: dP = {res.get('dp_trap', 0):.2f} бар")
            doc.add_paragraph(
                f"• Рекомендуемый Ду внешнего паропровода (до клапана): Ду {res.get('diameter_steam', 0)} мм (Скорость: {res.get('v_steam_line_speed', 0):.2f} м/с)")
            doc.add_paragraph(
                f"• Рекомендуемый Ду общего конденсатопровода ИТП: Ду {res.get('diameter_cond', 0)} мм (Скорость: {res.get('v_cond_line_speed', 0):.2f} м/с)")

            doc.save(filename)
            messagebox.showinfo("Успех", f"Расширенный технический отчет успешно сформирован:\n{filename}")
        except Exception as e:
            messagebox.showerror("Ошибка при выгрузке в Word", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ITPCalculator(root)
    root.mainloop()
