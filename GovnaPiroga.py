# -*- coding: utf-8 -*-
import os
import sys
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ezdxf

# Проверка версии Python
if sys.version_info < (3, 4):
    raise RuntimeError("Требуется Python 3.4 или выше (для Windows 7 x86)")

# Настройки папки
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
WORK_FOLDER = os.path.join(DESKTOP, "COXOproScan")
if not os.path.exists(WORK_FOLDER):
    try:
        os.makedirs(WORK_FOLDER)
    except OSError as e:
        raise RuntimeError(f"Не удалось создать рабочую папку: {str(e)}")

class COXOproScan:
    def __init__(self, root):
        self.root = root
        self.params = {}
        self.setup_ui()
        self.reset_settings()
        
    def setup_ui(self):
        self.root.title("COXOproScan v3.4 (Win7 x86)")
        self.root.geometry("850x600")
        self.root.resizable(False, False)
        self.center_window()
        
        # Стили
        style = ttk.Style()
        style.configure("TLabelFrame", font=('Arial', 9, 'bold'))
        style.configure("Accent.TButton", foreground="white", background="#4CAF50", font=('Arial', 10, 'bold'))
        
        # Основные параметры
        main_frame = ttk.LabelFrame(self.root, text="ОСНОВНЫЕ ПАРАМЕТРЫ", padding=10)
        main_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.add_param(main_frame, "scan_length", "Общая длина сканирования (мм):", 0)
        self.add_param(main_frame, "probe_depth", "Глубина зондирования (мм):", 1)
        self.add_param(main_frame, "retract", "Величина отвода (мм):", 2)
        self.add_param(main_frame, "speed", "Скорость (мм/мин):", 3)

        # Стартовая зона
        start_frame = ttk.LabelFrame(self.root, text="СТАРТОВАЯ ЗОНА", padding=10)
        start_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.add_checkbox(start_frame, "use_start_zone", "Активировать стартовую зону", 0)
        self.add_param(start_frame, "start_zone_length", "Длина зоны (мм):", 1, "use_start_zone")
        self.add_param(start_frame, "start_zone_step", "Шаг (мм):", 2, "use_start_zone")

        # Основная зона
        main_zone_frame = ttk.LabelFrame(self.root, text="ОСНОВНАЯ ЗОНА", padding=10)
        main_zone_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.add_param(main_zone_frame, "main_zone_step", "Шаг (мм):", 0)

        # Конечная зона
        end_frame = ttk.LabelFrame(self.root, text="КОНЕЧНАЯ ЗОНА", padding=10)
        end_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.add_checkbox(end_frame, "use_end_zone", "Активировать конечную зону", 0)
        self.add_param(end_frame, "end_zone_length", "Длина зоны (мм):", 1, "use_end_zone")
        self.add_param(end_frame, "end_zone_step", "Шаг (мм):", 2, "use_end_zone")

        # Кнопки
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(
            btn_frame,
            text="СБРОСИТЬ НАСТРОЙКИ",
            style="Accent.TButton",
            command=self.reset_settings
        ).pack(side=tk.LEFT, expand=True, padx=5)

        ttk.Button(
            btn_frame,
            text="ГЕНЕРИРОВАТЬ G-КОД",
            style="Accent.TButton",
            command=self.generate_gcode
        ).pack(side=tk.LEFT, expand=True, padx=5)

        ttk.Button(
            self.root,
            text="СОЗДАТЬ DXF ДЛЯ ARTCAM",
            style="Accent.TButton",
            command=self.create_artcam_file
        ).pack(fill=tk.X, padx=10, pady=5)

    def add_param(self, frame, param_name, label_text, row, dependency=None):
        container = ttk.Frame(frame)
        container.grid(row=row, column=0, sticky="ew", pady=3)
        ttk.Label(container, text=label_text, width=25, anchor="e").pack(side=tk.LEFT)
        var = tk.DoubleVar(value=0.0)
        entry = ttk.Entry(container, textvariable=var, width=10)
        entry.pack(side=tk.LEFT)
        self.params[param_name] = var
        
        if dependency:
            self.toggle_dependency(entry, dependency)
        return var

    def add_checkbox(self, frame, param_name, text, row):
        var = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(frame, text=text, variable=var)
        cb.grid(row=row, column=0, sticky="w", pady=2, padx=5)
        self.params[param_name] = var
        return var

    def toggle_dependency(self, widget, dependency):
        widget.state(["!disabled" if self.params[dependency].get() else "disabled"])
        self.params[dependency].trace("w", lambda *_, w=widget, d=dependency: 
            w.state(["!disabled" if self.params[d].get() else "disabled"]))

    def reset_settings(self):
        defaults = {
            "scan_length": 100.0,
            "probe_depth": 20.0,
            "retract": 5.0,
            "speed": 300.0,
            "use_start_zone": True,
            "start_zone_length": 10.0,
            "start_zone_step": 0.5,
            "main_zone_step": 2.0,
            "use_end_zone": True,
            "end_zone_length": 10.0,
            "end_zone_step": 0.5
        }
        for name, var in self.params.items():
            if name in defaults:
                var.set(defaults[name])

    def generate_gcode(self):
        try:
            params = {name: var.get() for name, var in self.params.items()}
            
            if not params["use_start_zone"]:
                params["start_zone_length"] = 0.0
            if not params["use_end_zone"]:
                params["end_zone_length"] = 0.0
            
            errors = []
            if params["scan_length"] <= 0: 
                errors.append("Общая длина сканирования должна быть > 0")
            if params["probe_depth"] <= 0: 
                errors.append("Глубина зондирования должна быть > 0")
            if params["use_start_zone"] and params["start_zone_step"] <= 0: 
                errors.append("Шаг стартовой зоны должен быть > 0")
            if params["main_zone_step"] <= 0: 
                errors.append("Шаг основной зоны должен быть > 0")
            if params["use_end_zone"] and params["end_zone_step"] <= 0: 
                errors.append("Шаг конечной зоны должен быть > 0")
            
            if params["start_zone_length"] >= params["scan_length"]:
                errors.append("Длина стартовой зоны должна быть меньше общей длины сканирования")
            if params["end_zone_length"] >= params["scan_length"]:
                errors.append("Длина конечной зоны должна быть меньше общей длины сканирования")
            if (params["start_zone_length"] + params["end_zone_length"]) >= params["scan_length"]:
                errors.append("Сумма длин стартовой и конечной зон должна быть меньше общей длины сканирования")
            
            if errors: 
                raise ValueError("\n".join(errors))

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

            # Завершение
            gcode.extend([
                "G90",
                "G0X0",
                "G0Y0",
                "(* выберете название файла сканирования *)",
                "M30",
                ""  # Добавленная пустая строка после M30
            ])

            filename = f"scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.tap"
            filepath = os.path.join(WORK_FOLDER, filename)
            
            with open(filepath, "w", encoding='cp1251') as f:
                f.write("\n".join(gcode))

            messagebox.showinfo(
                "Готово!",
                f"G-код сохранён в:\n{filepath}\n"
                f"Общая длина: {y_pos:.2f} мм\n"
                f"Шагов: {len(gcode) - 7}\n"
                f"Стартовая зона: {params['start_zone_length']} мм (шаг: {params['start_zone_step']} мм)\n"
                f"Основная зона: {main_zone_length:.2f} мм (шаг: {params['main_zone_step']} мм)\n"
                f"Конечная зона: {params['end_zone_length']} мм (шаг: {params['end_zone_step']} мм)"
            )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def create_artcam_file(self):
        try:
            points_file = filedialog.askopenfilename(
                initialdir=WORK_FOLDER,
                title="Выберите файл точек из Mach3",
                filetypes=(("Текстовые файлы", "*.txt"), ("Все файлы", "*.*"))
            )
            
            if not points_file:
                return

            points = []
            with open(points_file, 'r', encoding='cp1251') as f:
                for line in f:
                    line = line.strip()
                    if line:
                        parts = line.split(',')
                        if len(parts) >= 2:
                            try:
                                x = float(parts[0].strip())
                                y = float(parts[1].strip())
                                z = float(parts[2].strip()) if len(parts) >= 3 else 0.0
                                points.append((x, y, z))
                            except ValueError:
                                continue

            if len(points) < 2:
                raise ValueError("Необходимо минимум 2 точки")

            # Создаем DXF документ
            doc = ezdxf.new('R2010')
            msp = doc.modelspace()

            # Добавляем полилинию
            polyline = msp.add_polyline2d(points)

            filename = f"artcam_{datetime.now().strftime('%Y%m%d_%H%M%S')}.dxf"
            filepath = os.path.join(WORK_FOLDER, filename)
            doc.saveas(filepath)

            messagebox.showinfo(
                "Готово!",
                f"DXF создан:\n{filepath}\n"
                f"Точек: {len(points)}"
            )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

if __name__ == "__main__":
    root = tk.Tk()
    app = COXOproScan(root)
    root.mainloop()
