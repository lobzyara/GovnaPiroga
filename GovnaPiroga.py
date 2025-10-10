# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
from datetime import datetime
import sys
import ctypes
from typing import List, Optional

class COXOproScan:
    def __init__(self, root):
        self.root = root
        self.params = {}
        self.cyclic_params = {
            "depth_step": tk.DoubleVar(value=0.02),
            "copies_count": tk.IntVar(value=5)
        }
        self.peak_depth_warning_shown = False
        self.cyclic_warning_shown = False
        self.setup_ui()
        
        # Установка начальных значений
        self.params["retract"].set(5.0)
        self.params["speed"].set(300.0)
        self.params["main_zone_step"].set(1.0)
        self.params["probe_depth"].set(20.0)
        self.set_icon()

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def set_icon(self):
        if sys.platform == 'win32':
            try:
                icon_path = self.resource_path('IMG_8084.ico')
                self.root.iconbitmap(icon_path)
            except Exception as e:
                print(f"Ошибка загрузки иконки: {e}")

    def setup_ui(self):
        self.root.title("COXOproScan v3.4")
        self.root.geometry("500x500")
        self.root.resizable(False, False)
        self.center_window()
        
        style = ttk.Style()
        style.configure(".", padding=1)
        style.configure("TLabelFrame", font=('Arial', 9, 'bold'), padding=3)
        style.configure("TEntry", padding=1, font=('Arial', 8))
        style.configure("TCheckbutton", font=('Arial', 8))
        style.configure("TLabel", font=('Arial', 8))
        style.configure("Custom.TButton", font=('Arial', 9, 'bold'), padding=2, borderwidth=1)

        main_container = ttk.Frame(self.root, padding=1)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Основные параметры
        main_frame = ttk.LabelFrame(main_container, text="ОСНОВНЫЕ ПАРАМЕТРЫ", padding=3)
        main_frame.pack(fill=tk.X, pady=1)
        
        self.add_param(main_frame, "scan_length", "Длина (мм):", 0)
        self.add_param(main_frame, "retract", "Отвод (мм):", 1)
        self.add_param(main_frame, "speed", "Скорость (мм/мин):", 2)
        
        # Пиковая глубина
        self.peak_depth_visible = False
        self.peak_depth_frame = ttk.Frame(main_container)
        self.peak_depth_frame.pack(fill=tk.X, pady=0)
        
        self.peak_depth_header = ttk.Label(
            self.peak_depth_frame, 
            text="▶ ПИКОВАЯ ГЛУБИНА ▶",
            padding=2,
            font=('Arial', 8, 'bold'),
            relief="flat",
            cursor="hand2"
        )
        self.peak_depth_header.pack(fill=tk.X)
        self.peak_depth_header.bind("<Button-1>", self.toggle_peak_depth)
        
        self.peak_depth_container = ttk.Frame(self.peak_depth_frame)
        self.add_param(self.peak_depth_container, "probe_depth", "Глубина (мм):", 0)

        # Зоны сканирования
        start_frame = ttk.LabelFrame(main_container, text="СТАРТОВАЯ ЗОНА", padding=3)
        start_frame.pack(fill=tk.X, pady=1)
        
        self.add_checkbox(start_frame, "use_start_zone", "Активировать", 0)
        self.add_param(start_frame, "start_zone_length", "Длина (мм):", 1, "use_start_zone")
        self.add_param(start_frame, "start_zone_step", "Шаг (мм):", 2, "use_start_zone")

        main_zone_frame = ttk.LabelFrame(main_container, text="ОСНОВНАЯ ЗОНА", padding=3)
        main_zone_frame.pack(fill=tk.X, pady=1)
        self.add_param(main_zone_frame, "main_zone_step", "Шаг (мм):", 0)

        end_frame = ttk.LabelFrame(main_container, text="КОНЕЧНАЯ ЗОНА", padding=3)
        end_frame.pack(fill=tk.X, pady=1)
        
        self.add_checkbox(end_frame, "use_end_zone", "Активировать", 0)
        self.add_param(end_frame, "end_zone_length", "Длина (мм):", 1, "use_end_zone")
        self.add_param(end_frame, "end_zone_step", "Шаг (мм):", 2, "use_end_zone")

        # Циклическая обработка
        self.cyclic_frame = ttk.Frame(main_container)
        self.cyclic_frame.pack(fill=tk.X, pady=0)
        
        self.cyclic_header = ttk.Label(
            self.cyclic_frame, 
            text="ЦИКЛИЧЕСКАЯ ОБРАБОТКА",
            padding=2,
            font=('Arial', 8, 'bold'),
            relief="flat"
        )
        self.cyclic_header.pack(fill=tk.X)
        
        # Контейнер для параметров циклической обработки
        self.cyclic_params_container = ttk.Frame(self.cyclic_frame)
        
        # Параметры
        self.add_cyclic_param(self.cyclic_params_container, "depth_step", "Шаг углубления (мм):", 0)
        self.add_cyclic_param(self.cyclic_params_container, "copies_count", "Количество проходов:", 1)
        
        # Кнопки управления
        btn_frame_cyclic = ttk.Frame(self.cyclic_params_container)
        btn_frame_cyclic.pack(fill=tk.X, pady=(5,0))
        
        ttk.Button(
            btn_frame_cyclic,
            text="ОТКРЫТЬ УП",
            style="Custom.TButton",
            command=self.open_tap_file
        ).pack(side=tk.LEFT, expand=True, padx=1)

        ttk.Button(
            btn_frame_cyclic,
            text="СОХРАНИТЬ УП",
            style="Custom.TButton",
            command=self.process_tap_file
        ).pack(side=tk.LEFT, expand=True, padx=1)

        # Основные кнопки
        btn_frame = ttk.Frame(main_container)
        btn_frame.pack(fill=tk.X, pady=(3,0))
        
        ttk.Button(
            btn_frame,
            text="G-КОД",
            style="Custom.TButton",
            command=self.generate_gcode
        ).pack(side=tk.LEFT, expand=True, padx=1)

        ttk.Button(
            btn_frame,
            text="ВЕКТОР ИЗ СКАНА",
            style="Custom.TButton",
            command=self.create_artcam_file
        ).pack(side=tk.LEFT, expand=True, padx=1)

        # Показываем параметры циклической обработки по умолчанию
        self.cyclic_params_container.pack(fill=tk.X)

    def add_cyclic_param(self, frame, param_name, label_text, row):
        container = ttk.Frame(frame)
        container.pack(fill=tk.X, pady=0)
        ttk.Label(container, text=label_text, width=20, anchor="w").pack(side=tk.LEFT)
        if param_name == "copies_count":
            entry = ttk.Entry(container, textvariable=self.cyclic_params[param_name], width=5)
        else:
            entry = ttk.Entry(container, textvariable=self.cyclic_params[param_name], width=10)
        entry.pack(side=tk.LEFT, padx=1)
        
        # Добавляем обработчик для показа предупреждения
        self.cyclic_params[param_name].trace("w", self.show_cyclic_warning)

    def show_cyclic_warning(self, *args):
        """Показывает предупреждение при изменении параметров циклической обработки"""
        if not self.cyclic_warning_shown:
            messagebox.showwarning(
                "Внимание!",
                "Убедитесь в том, что последняя точка вектора находится выше всей обрабатываемой поверхности!"
            )
            self.cyclic_warning_shown = True

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def add_param(self, frame, param_name, label_text, row, dependency=None):
        container = ttk.Frame(frame)
        container.pack(fill=tk.X, pady=0)
        ttk.Label(container, text=label_text, width=16, anchor="w").pack(side=tk.LEFT)
        var = tk.DoubleVar(value=0.0)
        entry = ttk.Entry(container, textvariable=var, width=10)
        entry.pack(side=tk.LEFT, padx=1)
        self.params[param_name] = var
        if dependency:
            self.toggle_dependency(entry, dependency)
        return var

    def add_checkbox(self, frame, param_name, text, row):
        var = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(frame, text=text, variable=var)
        cb.pack(anchor="w", pady=0)
        self.params[param_name] = var
        return var

    def toggle_dependency(self, widget, dependency):
        widget.state(["!disabled" if self.params[dependency].get() else "disabled"])
        self.params[dependency].trace("w", lambda *_, w=widget, d=dependency: 
            w.state(["!disabled" if self.params[d].get() else "disabled"]))

    def toggle_peak_depth(self, event=None):
        """Переключает видимость параметра пиковой глубины"""
        self.peak_depth_visible = not self.peak_depth_visible
        if self.peak_depth_visible:
            self.peak_depth_container.pack(fill=tk.X)
            self.peak_depth_header.config(text="▼ ПИКОВАЯ ГЛУБИНА ▼")
            
            # Показываем предупреждение при первом открытии
            if not self.peak_depth_warning_shown:
                messagebox.showwarning(
                    "Внимание! Безопасность!",
                    "Параметр 'ПИКОВАЯ ГЛУБИНА' отвечает за безопасность!\n"
                    "Использовать только после ознакомления с FAQ!"
                )
                self.peak_depth_warning_shown = True
        else:
            self.peak_depth_container.pack_forget()
            self.peak_depth_header.config(text="▶ ПИКОВАЯ ГЛУБИНА ▶")
        self.root.update_idletasks()

    def get_order_number(self):
        """Запрашивает номер завода у пользователя"""
        while True:
            order_number = simpledialog.askstring(
                "Номер заказа", 
                "Введите номер заказа:",
                parent=self.root
            )
            
            if order_number is None:
                return None
                
            order_number = order_number.strip()
            if order_number:
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    order_number = order_number.replace(char, '')
                return order_number
            else:
                messagebox.showwarning("Пустой номер", "Номер заказа не может быть пустым!")

    def generate_gcode(self):
        try:
            params = {name: var.get() for name, var in self.params.items()}
            
            if not params["use_start_zone"]:
                params["start_zone_length"] = 0.0
            if not params["use_end_zone"]:
                params["end_zone_length"] = 0.0
            
            errors = []
            if params["scan_length"] <= 0: errors.append("Длина сканирования должна быть > 0")
            if params["probe_depth"] <= 0: errors.append("Глубина зондирования должна быть > 0")
            if params["use_start_zone"] and params["start_zone_step"] <= 0: errors.append("Шаг стартовой зоны должен быть > 0")
            if params["main_zone_step"] <= 0: errors.append("Шаг основной зоны должен быть > 0")
            if params["use_end_zone"] and params["end_zone_step"] <= 0: errors.append("Шаг конечной зоны должен быть > 0")
            if params["start_zone_length"] >= params["scan_length"]: errors.append("Длина стартовой зоны должна быть меньше общей длины")
            if params["end_zone_length"] >= params["scan_length"]: errors.append("Длина конечной зоны должна быть меньше общей длины")
            if (params["start_zone_length"] + params["end_zone_length"]) >= params["scan_length"]: errors.append("Сумма длин стартовой и конечной зон должна быть меньше общей длины")
            
            if errors: 
                raise ValueError("\n".join(errors))

            order_number = self.get_order_number()
            if order_number is None:
                return

            gcode = [
                "(*** сканирование ***)",
                "M40",
                f"F{params['speed']}",
                "(установите щуп у края диска, затем нажмите СТАРТ!)",
                "M00",
                "G91"
            ]

            y_pos = 0.0
            total_length = params["scan_length"]

            # Стартовая зона
            if params["use_start_zone"] and params["start_zone_length"] > 0:
                start_steps = int(params["start_zone_length"] / params["start_zone_step"])
                start_remainder = params["start_zone_length"] % params["start_zone_step"]
                
                for _ in range(start_steps):
                    gcode.append(f"G31X{params['probe_depth']}")
                    gcode.append(f"G0X-{params['retract']}")
                    gcode.append(f"G0Y{params['start_zone_step']}")
                    y_pos += params["start_zone_step"]
                
                if start_remainder > 0:
                    gcode.append(f"G31X{params['probe_depth']}")
                    gcode.append(f"G0X-{params['retract']}")
                    gcode.append(f"G0Y{start_remainder}")
                    y_pos += start_remainder

            # Основная зона
            main_zone_length = total_length - params["start_zone_length"] - params["end_zone_length"]
            main_steps = int(main_zone_length / params["main_zone_step"])
            main_remainder = main_zone_length % params["main_zone_step"]
            
            for _ in range(main_steps):
                gcode.append(f"G31X{params['probe_depth']}")
                gcode.append(f"G0X-{params['retract']}")
                gcode.append(f"G0Y{params['main_zone_step']}")
                y_pos += params["main_zone_step"]
            
            if main_remainder > 0:
                gcode.append(f"G31X{params['probe_depth']}")
                gcode.append(f"G0X-{params['retract']}")
                gcode.append(f"G0Y{main_remainder}")
                y_pos += main_remainder

            # Конечная зона
            if params["use_end_zone"] and params["end_zone_length"] > 0:
                end_steps = int(params["end_zone_length"] / params["end_zone_step"])
                end_remainder = params["end_zone_length"] % params["end_zone_step"]
                
                for _ in range(end_steps):
                    gcode.append(f"G31X{params['probe_depth']}")
                    gcode.append(f"G0X-{params['retract']}")
                    gcode.append(f"G0Y{params['end_zone_step']}")
                    y_pos += params["end_zone_step"]
                
                if end_remainder > 0:
                    gcode.append(f"G31X{params['probe_depth']}")
                    gcode.append(f"G0X-{params['retract']}")
                    gcode.append(f"G0Y{end_remainder}")
                    y_pos += end_remainder

            gcode.extend([
                "G90",
                "G0X0",
                "G0Y0",
                "(* выберете название файла сканирования *)",
                "M30",
                ""
            ])

            filename = f"{order_number}scan.tap"
            filepath = os.path.join(os.path.expanduser("~"), "Desktop", "COXOproScan", filename)
            
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            with open(filepath, "w", encoding='cp1251') as f:
                f.write("\n".join(gcode))

            messagebox.showinfo(
                "Готово!",
                f"G-код сохранён в:\n{filepath}\n"
                f"Номер заказа: {order_number}\n"
                f"Общая длина: {y_pos:.2f} мм\n"
                f"Шагов: {len(gcode) - 7}\n"
                f"Стартовая зона: {params['start_zone_length']} мм\n"
                f"Основная зона: {main_zone_length:.2f} мм\n"
                f"Конечная зона: {params['end_zone_length']} мм"
            )
        except ValueError as ve:
            messagebox.showerror("Ошибка ввода", str(ve))
        except IOError as ioe:
            messagebox.showerror("Ошибка файла", f"Не удалось сохранить файл:\n{str(ioe)}")
        except Exception as e:
            messagebox.showerror("Неизвестная ошибка", f"Произошла непредвиденная ошибка:\n{str(e)}")

    def create_artcam_file(self):
        try:
            points_file = filedialog.askopenfilename(
                initialdir=os.path.join(os.path.expanduser("~"), "Desktop", "COXOproScan"),
                title="Выберите файл точек из Mach3",
                filetypes=(("Текстовые файлы", "*.txt"), ("Все файлы", "*.*"))
            )
            
            if not points_file:
                return

            original_filename = os.path.splitext(os.path.basename(points_file))[0]
            new_filename = f"{original_filename}vector.dxf"
            filepath = os.path.join(os.path.expanduser("~"), "Desktop", "COXOproScan", new_filename)

            points = []
            with open(points_file, 'r', encoding='cp1251') as f:
                for line in f:
                    line = line.strip()
                    if line:
                        parts = [part.strip() for part in line.split(',')]
                        if len(parts) >= 2:
                            try:
                                x = float(parts[0])
                                y = float(parts[1])
                                z = float(parts[2]) if len(parts) >= 3 else 0.0
                                points.append((x, y, z))
                            except ValueError:
                                continue

            if len(points) < 2:
                raise ValueError("Необходимо минимум 2 точки для создания полилинии")
            
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            
            with open(filepath, 'w', encoding='cp1251') as f:
                f.write("  0\nSECTION\n  2\nENTITIES\n")
                f.write("  0\nPOLYLINE\n  8\n0\n")
                
                for x, y, z in points:
                    f.write(f"  0\nVERTEX\n  8\n0\n")
                    f.write(f" 10\n{x:.5f}\n")
                    f.write(f" 20\n{y:.5f}\n")
                    f.write(f" 30\n{z:.5f}\n")
                    f.write(" 70\n    32\n")
                
                f.write("  0\nSEQEND\n")
                f.write("  0\nENDSEC\n")
                f.write("  0\nEOF")

            messagebox.showinfo(
                "Готово!",
                f"DXF файл успешно создан:\n{filepath}\n"
                f"Исходный файл: {os.path.basename(points_file)}\n"
                f"Количество точек: {len(points)}"
            )
        except ValueError as ve:
            messagebox.showerror("Ошибка данных", str(ve))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")

    def open_tap_file(self):
        self.tap_file_path = filedialog.askopenfilename(
            title="Выберите файл .tap",
            filetypes=(("TAP files", "*.tap"), ("Все файлы", "*.*")),
            initialdir=os.path.join(os.path.expanduser("~"), "Desktop", "COXOproScan")
        )
        if self.tap_file_path:
            messagebox.showinfo("Файл выбран", f"Выбран файл:\n{self.tap_file_path}")

    def _generate_cyclic_passes(self, original_block: List[str], depth_step: float, copies_count: int) -> List[str]:
        """Генерирует циклические проходы с сохранением смещения по X"""
        all_passes = []
        
        # Первый проход - оригинальный
        all_passes.extend(original_block)
        
        # Если нужен только один проход - возвращаем как есть
        if copies_count <= 1:
            return all_passes
        
        # Извлекаем первую Y координату для возврата
        first_y = self._extract_first_y_coordinate(original_block)
        
        # Генерируем дополнительные проходы
        for pass_num in range(1, copies_count):
            current_depth_offset = depth_step * pass_num
            
            # Добавляем переход к началу
            all_passes.append("(*** ЦИКЛ ПРОХОД {} ***)".format(pass_num + 1))
            
            # Перемещаемся к началу по Y, сохраняя текущую X координату
            if first_y is not None:
                all_passes.append(f"G0 Y{first_y:.4f}")
            
            # Дублируем весь рабочий блок с НАКОПЛЕННЫМ смещением глубины
            deepened_block = self._apply_depth_offset(original_block, current_depth_offset)
            all_passes.extend(deepened_block)
        
        return all_passes

    def _extract_first_y_coordinate(self, block: List[str]) -> Optional[float]:
        """Извлекает Y-координату из первой команды перемещения в блоке"""
        for line in block:
            if 'Y' in line and not line.startswith('('):
                try:
                    # Ищем Y-координату в строке
                    parts = line.split()
                    for part in parts:
                        if part.startswith('Y'):
                            return float(part[1:])
                except ValueError:
                    continue
        return None

    def _apply_depth_offset(self, block: List[str], offset: float) -> List[str]:
        """Применяет смещение глубины ко всем X-координатам в блоке"""
        processed_block = []
        
        for line in block:
            if 'X' in line and not line.startswith('('):
                # Обрабатываем только строки с X-координатами
                processed_line = self._offset_x_coordinate(line, offset)
                processed_block.append(processed_line)
            else:
                # Оставляем без изменений комментарии и другие команды
                processed_block.append(line)
        
        return processed_block

    def _offset_x_coordinate(self, line: str, offset: float) -> str:
        """Применяет смещение к X-координате в строке G-кода"""
        parts = line.split()
        
        for i, part in enumerate(parts):
            if part.startswith('X'):
                try:
                    x_val = float(part[1:])
                    new_x = x_val + offset
                    parts[i] = f"X{new_x:.4f}"
                except ValueError:
                    continue
        
        return " ".join(parts)

    def process_tap_file(self):
        """Обрабатывает .tap файл с циклической обработкой"""
        try:
            if not hasattr(self, 'tap_file_path') or not self.tap_file_path:
                raise ValueError("Сначала выберите файл .tap (кнопка ОТКРЫТЬ УП)")

            depth_step = self.cyclic_params["depth_step"].get()
            copies_count = self.cyclic_params["copies_count"].get()

            if depth_step <= 0:
                raise ValueError("Шаг углубления должен быть > 0")
            if copies_count < 1:
                raise ValueError("Количество проходов должно быть >= 1")

            # Вычисляем общую глубину заглубления (УЧИТЫВАЕМ ПЕРВЫЙ ПРОХОД!)
            total_depth = depth_step * copies_count
            total_depth_str = f"{total_depth:.2f}".replace('.', '')

            # Чтение файла
            with open(self.tap_file_path, 'r', encoding='cp1251') as f:
                original_lines = [line.strip() for line in f if line.strip()]

            # Находим разделитель - строку ПЕРЕД G0 X0 Y0
            split_index = -1
            for i in range(len(original_lines) - 1):
                if original_lines[i+1].startswith("G0 X0") or original_lines[i+1].startswith("G0Y0"):
                    split_index = i
                    break
            
            if split_index == -1:
                raise ValueError("Не найдена точка вставки циклов (перед G0 X0 Y0)")

            # Разделяем файл
            header_and_working = original_lines[:split_index + 1]
            footer = original_lines[split_index + 1:]
            
            # Извлекаем рабочий блок (после G21, но саму G21 не включаем)
            working_block_start = -1
            for i, line in enumerate(header_and_working):
                if "G21" in line:
                    working_block_start = i
                    break
            
            if working_block_start == -1:
                header = header_and_working
                working_block = []
            else:
                header = header_and_working[:working_block_start + 1]
                working_block = header_and_working[working_block_start + 1:]
            
            # Генерируем все проходы
            all_passes = self._generate_cyclic_passes(working_block, depth_step, copies_count)
            
            # Собираем новый файл
            new_content = header + all_passes + footer
            
            # Формируем имя файла
            original_filename = os.path.basename(self.tap_file_path)
            name_without_ext = os.path.splitext(original_filename)[0]
            new_filename = f"{name_without_ext}cycle{total_depth_str}.tap"
            
            # Сохраняем
            new_filepath = os.path.join(os.path.dirname(self.tap_file_path), new_filename)
            
            with open(new_filepath, 'w', encoding='cp1251') as f:
                f.write("\n".join(new_content) + "\n")

            messagebox.showinfo(
                "Готово!",
                f"Файл успешно обработан:\n{new_filepath}\n"
                f"Параметры обработки:\n"
                f"- Шаг углубления: {depth_step} мм\n"
                f"- Количество проходов: {copies_count}\n"
                f"- Общая глубина заглубления: {total_depth:.2f} мм"
            )
            
        except ValueError as ve:
            messagebox.showerror("Ошибка ввода", str(ve))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки файла:\n{str(e)}")

if __name__ == "__main__":
    if sys.platform == 'win32':
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('COXOproScan.1.0')
    
    root = tk.Tk()
    app = COXOproScan(root)
    root.mainloop()