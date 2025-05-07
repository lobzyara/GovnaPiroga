import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import csv
from datetime import datetime
import sys

class IvanorProScan:
    def __init__(self, root):
        self.root = root
        self.params = {}
        self.initialize_ui()
        self.setup_workspace()
        
    def setup_workspace(self):
        """Инициализация рабочей папки с обработкой ошибок"""
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            self.work_folder = os.path.join(desktop, "IvanorProScan")
            os.makedirs(self.work_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать рабочую папку:\n{str(e)}")
            sys.exit(1)

    def initialize_ui(self):
        """Оптимизированная инициализация интерфейса"""
        self.root.title("IvanorProScan v3.5")
        self.root.geometry("850x460")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Очистка возможных предыдущих элементов
        for widget in self.root.winfo_children():
            widget.destroy()

        self.setup_styles()
        self.create_main_frame()
        self.center_window()

    def create_main_frame(self):
        """Создание основного фрейма с элементами управления"""
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Основные параметры
        self.create_parameter_section(main_frame)
        
        # Зоны сканирования
        self.create_scan_zones(main_frame)
        
        # Кнопки управления
        self.create_control_buttons(main_frame)
        
        # Подпись
        ttk.Label(
            main_frame,
            text="COXO inc | v3.5 | Память: стабильная",
            font=('Arial', 7),
            foreground="gray"
        ).pack(side=tk.BOTTOM, pady=2)

    def create_parameter_section(self, parent):
        """Секция основных параметров"""
        frame = ttk.LabelFrame(parent, text="ОСНОВНЫЕ ПАРАМЕТРЫ", padding=5)
        frame.pack(fill=tk.X, pady=3)
        
        grid_frame = ttk.Frame(frame)
        grid_frame.pack(fill=tk.X)
        
        params = [
            ("scan_length", "Длина сканирования (мм):", 0, 0),
            ("main_step", "Шаг основной зоны (мм):", 0, 1),
            ("probe_depth", "Глубина зондирования (мм):", 1, 0),
            ("retract", "Величина отвода (мм):", 1, 1),
            ("speed", "Скорость (мм/мин):", 2, 0)
        ]
        
        for name, label, row, col in params:
            self.add_compact_param(grid_frame, name, label, row, col)

    def create_scan_zones(self, parent):
        """Секция зон сканирования"""
        zones_frame = ttk.Frame(parent)
        zones_frame.pack(fill=tk.X, pady=3)

        # Стартовая зона
        start_frame = ttk.LabelFrame(zones_frame, text="СТАРТОВАЯ ЗОНА", padding=5)
        start_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.add_checkbox(start_frame, "use_start_zone", "Активировать", 0)
        self.add_compact_param(start_frame, "start_zone", "Длина (мм):", 1, 0, "use_start_zone")
        self.add_compact_param(start_frame, "start_step", "Шаг (мм):", 1, 1, "use_start_zone")

        # Конечная зона
        end_frame = ttk.LabelFrame(zones_frame, text="КОНЕЧНАЯ ЗОНА", padding=5)
        end_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.add_checkbox(end_frame, "use_end_zone", "Активировать", 0)
        self.add_compact_param(end_frame, "end_zone", "Длина (мм):", 1, 0, "use_end_zone")
        self.add_compact_param(end_frame, "end_step", "Шаг (мм):", 1, 1, "use_end_zone")

    def create_control_buttons(self, parent):
        """Кнопки управления"""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=8)

        buttons = [
            ("СБРОС", "Secondary.TButton", self.reset_settings),
            ("G-КОД", "Accent.TButton", self.generate_gcode),
            ("ARTCAM", "Accent.TButton", self.create_artcam_file)
        ]
        
        for text, style, command in buttons:
            btn = ttk.Button(
                btn_frame,
                text=text,
                style=style,
                command=command,
                width=10
            )
            btn.pack(side=tk.LEFT, expand=True, padx=3)
            btn.bind("<Button-1>", lambda e: self.root.update())  # Принудительное обновление UI

    def setup_styles(self):
        """Настройка стилей с оптимизацией"""
        style = ttk.Style()
        style.theme_use('clam')  # Более стабильная тема
        
        # Очистка предыдущих стилей
        style._style_names = []
        
        style.configure("TLabelFrame", font=('Arial', 8, 'bold'))
        style.configure("Accent.TButton", foreground="white", background="#4CAF50", font=('Arial', 9))
        style.configure("Secondary.TButton", font=('Arial', 9))
        style.map("Accent.TButton", background=[('active', '#45a049')])

    def add_compact_param(self, frame, param_name, label_text, row, col, dependency=None):
        """Оптимизированное добавление параметра"""
        container = ttk.Frame(frame)
        container.grid(row=row, column=col, sticky="ew", padx=2, pady=1)
        
        label = ttk.Label(container, text=label_text, width=14, anchor="e")
        label.pack(side=tk.LEFT)
        
        var = tk.DoubleVar(value=0.0)
        entry = ttk.Entry(container, textvariable=var, width=10)
        entry.pack(side=tk.LEFT)
        
        self.params[param_name] = var
        
        if dependency:
            entry.state(["!disabled" if self.params[dependency].get() else "disabled"])
            self.params[dependency].trace_add(
                "write", 
                lambda *_, e=entry, d=self.params[dependency]: e.state(["!disabled" if d.get() else "disabled"])
            )
        
        return var

    def add_checkbox(self, frame, param_name, text, row):
        """Оптимизированное добавление чекбокса"""
        var = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(frame, text=text, variable=var, style='Toolbutton')
        cb.grid(row=row, column=0, sticky="w", pady=1, padx=5)
        self.params[param_name] = var
        return var

    def reset_settings(self):
        """Сброс настроек с очисткой памяти"""
        for var in self.params.values():
            if isinstance(var, tk.DoubleVar):
                var.set(0.0)
            elif isinstance(var, tk.BooleanVar):
                var.set(False)
        self.root.update()

    def generate_gcode(self):
        """Генерация G-кода с обработкой памяти"""
        try:
            # Очистка памяти перед генерацией
            self.root.update()
            
            params = {name: var.get() for name, var in self.params.items()}
            
            # Валидация параметров
            if params["scan_length"] <= 0:
                raise ValueError("Длина сканирования должна быть > 0")
            if params["probe_depth"] <= 0:
                raise ValueError("Глубина зондирования должна быть > 0")

            # Генерация G-кода
            gcode = [
                "%",
                "G21 G90 G17 G40 G49 G80",
                "M5 M9",
                "(*** СКАНИРОВАНИЕ ***)",
                f"F{params['speed']}",
                "M00 (Подведите щуп к краю диска)",
                "G91 G94"
            ]

            # Оптимизированный алгоритм генерации
            gcode.extend(self.generate_scan_moves(params))
            
            gcode.extend([
                "G90",
                "G0 X0 Y0",
                "M30",
                "%"
            ])

            # Сохранение с контролем памяти
            filename = self.save_gcode_file(gcode)
            
            messagebox.showinfo(
                "Готово",
                f"G-код сохранен в:\n{filename}\n"
                f"Общая длина: {self.calculate_scan_length(params):.2f} мм"
            )
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.root.update()

    def generate_scan_moves(self, params):
        """Генерация движений сканирования"""
        moves = []
        y_pos = 0.0
        
        # Стартовая зона
        if params["use_start_zone"] and params["start_zone"] > 0:
            moves.extend(self.generate_zone_moves(
                params["start_zone"],
                params["start_step"],
                params["probe_depth"],
                params["retract"],
                y_pos
            ))
            y_pos += params["start_zone"]

        # Основная зона
        main_zone_end = params["scan_length"] - (params["end_zone"] if params["use_end_zone"] else 0)
        if main_zone_end > y_pos:
            moves.extend(self.generate_zone_moves(
                main_zone_end - y_pos,
                params["main_step"],
                params["probe_depth"],
                params["retract"],
                y_pos
            ))
            y_pos = main_zone_end

        # Конечная зона
        if params["use_end_zone"] and params["end_zone"] > 0:
            moves.extend(self.generate_zone_moves(
                params["end_zone"],
                params["end_step"],
                params["probe_depth"],
                params["retract"],
                y_pos
            ))
            
        return moves

    def generate_zone_moves(self, length, step, depth, retract, start_pos):
        """Генерация движений для одной зоны"""
        moves = []
        pos = 0.0
        while pos < length:
            moves.append(f"G31 X{depth} (Зонд вниз)")
            moves.append(f"G0 X-{retract} (Отвод)")
            if pos + step <= length:
                moves.append(f"G0 Y{step} (Шаг {start_pos + pos:.1f}-{start_pos + pos + step:.1f} мм)")
                pos += step
            else:
                last_step = length - pos
                moves.append(f"G0 Y{last_step} (Финальный шаг)")
                pos = length
        return moves

    def calculate_scan_length(self, params):
        """Расчет общей длины сканирования"""
        total = 0.0
        if params["use_start_zone"]:
            total += params["start_zone"]
        if params["use_end_zone"]:
            total += params["end_zone"]
        total += max(0, params["scan_length"] - (params["start_zone"] if params["use_start_zone"] else 0) 
                              - (params["end_zone"] if params["use_end_zone"] else 0))
        return total

    def save_gcode_file(self, gcode):
        """Сохранение G-кода с обработкой ошибок"""
        filename = f"scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.tap"
        filepath = os.path.join(self.work_folder, filename)
        
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write("\n".join(gcode))
            return filename
        except Exception as e:
            raise Exception(f"Ошибка сохранения файла: {str(e)}")

    def create_artcam_file(self):
        """Создание файла для ArtCAM с оптимизацией памяти"""
        try:
            self.root.update()  # Освобождение ресурсов
            
            points_file = filedialog.askopenfilename(
                initialdir=self.work_folder,
                title="Выберите файл точек",
                filetypes=(("Текстовые файлы", "*.txt"), ("Все файлы", "*.*"))
            
            if not points_file:
                return

            # Оптимизированное чтение точек
            points = self.read_points_file(points_file)
            
            if len(points) < 2:
                raise ValueError("Недостаточно точек для создания файла")
            
            # Генерация файла
            filename = f"scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            self.save_artcam_file(points, filename)
            
            messagebox.showinfo("Готово", f"Файл создан:\n{filename}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.root.update()

    def read_points_file(self, filepath):
        """Оптимизированное чтение точек из файла"""
        points = []
        current_x, current_y = 0.0, 0.0
        try:
            with open(filepath, 'r') as f:
                for line in f:
                    parts = line.strip().split()
                    if len(parts) >= 2:
                        current_x += float(parts[0])
                        current_y += float(parts[1])
                        points.append((current_x, current_y))
                        # Периодическое обновление UI для больших файлов
                        if len(points) % 100 == 0:
                            self.root.update()
        except Exception as e:
            raise Exception(f"Ошибка чтения файла точек: {str(e)}")
        return points

    def save_artcam_file(self, points, filename):
        """Сохранение файла для ArtCAM"""
        filepath = os.path.join(self.work_folder, filename)
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("0\nSECTION\n2\nENTITIES\n0\nPOLYLINE\n8\n0\n")
                
                for i, (x, y) in enumerate(points):
                    f.write(f"0\nVERTEX\n8\n0\n10\n{x:.5f}\n20\n{y:.5f}\n30\n0.00000\n70\n32\n0\n")
                    if i % 50 == 0:  # Периодическое обновление UI
                        self.root.update()
                
                f.write("0\nSEQEND\n")
        except Exception as e:
            raise Exception(f"Ошибка сохранения файла: {str(e)}")

    def center_window(self):
        """Центрирование окна с обработкой ошибок"""
        try:
            self.root.update_idletasks()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            x = (self.root.winfo_screenwidth() // 2) - (width // 2)
            y = (self.root.winfo_screenheight() // 2) - (height // 2)
            self.root.geometry(f'{width}x{height}+{x}+{y}')
        except:
            pass

    def on_close(self):
        """Обработчик закрытия окна с освобождением ресурсов"""
        try:
            self.root.quit()
            self.root.destroy()
        except:
            pass
        finally:
            sys.exit(0)

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = IvanorProScan(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Критическая ошибка", f"Программа завершена из-за ошибки:\n{str(e)}")
        sys.exit(1)
