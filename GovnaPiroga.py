import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import csv
from datetime import datetime

# Настройки папки
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
WORK_FOLDER = os.path.join(DESKTOP, "IvanorProScan")
os.makedirs(WORK_FOLDER, exist_ok=True)

class IvanorProScan:
    def __init__(self, root):
        self.root = root
        self.params = {}
        self.setup_ui()
        self.reset_settings()
       
    def setup_ui(self):
        self.root.title("IvanorProScan v3.4")
        self.root.geometry("850x460")
        self.root.resizable(False, False)
        self.center_window()
        
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        self.setup_styles()

        # Основные параметры в 2 колонки
        params_frame = ttk.LabelFrame(main_frame, text="ОСНОВНЫЕ ПАРАМЕТРЫ", padding=5)
        params_frame.pack(fill=tk.X, pady=3)
        
        param_grid = ttk.Frame(params_frame)
        param_grid.pack(fill=tk.X)
        
        self.add_compact_param(param_grid, "scan_length", "Длина сканирования (мм):", 0, 0)
        self.add_compact_param(param_grid, "main_step", "Шаг основной зоны (мм):", 0, 1)
        self.add_compact_param(param_grid, "probe_depth", "Глубина зондирования (мм):", 1, 0)
        self.add_compact_param(param_grid, "retract", "Величина отвода (мм):", 1, 1)
        self.add_compact_param(param_grid, "speed", "Скорость (мм/мин):", 2, 0)

        # Зоны сканирования в одной строке
        zones_frame = ttk.Frame(main_frame)
        zones_frame.pack(fill=tk.X, pady=3)

        # Стартовая зона
        start_zone_frame = ttk.LabelFrame(zones_frame, text="СТАРТОВАЯ ЗОНА", padding=5)
        start_zone_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.add_checkbox(start_zone_frame, "use_start_zone", "Активировать", 0)
        self.add_compact_param(start_zone_frame, "start_zone", "Длина (мм):", 1, 0, "use_start_zone")
        self.add_compact_param(start_zone_frame, "start_step", "Шаг (мм):", 1, 1, "use_start_zone")

        # Конечная зона
        end_zone_frame = ttk.LabelFrame(zones_frame, text="КОНЕЧНАЯ ЗОНА", padding=5)
        end_zone_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.add_checkbox(end_zone_frame, "use_end_zone", "Активировать", 0)
        self.add_compact_param(end_zone_frame, "end_zone", "Длина (мм):", 1, 0, "use_end_zone")
        self.add_compact_param(end_zone_frame, "end_step", "Шаг (мм):", 1, 1, "use_end_zone")

        # Кнопки управления
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=8)

        ttk.Button(
            btn_frame,
            text="СБРОС",
            style="Secondary.TButton",
            command=self.reset_settings,
            width=10
        ).pack(side=tk.LEFT, expand=True, padx=3)

        ttk.Button(
            btn_frame,
            text="G-КОД",
            style="Accent.TButton",
            command=self.generate_gcode,
            width=10
        ).pack(side=tk.LEFT, expand=True, padx=3)

        ttk.Button(
            btn_frame,
            text="ARTCAM",
            style="Accent.TButton",
            command=self.create_artcam_file,
            width=10
        ).pack(side=tk.LEFT, expand=True, padx=3)

        # Подпись
        ttk.Label(
            main_frame,
            text="COXO inc | v3.4",
            font=('Arial', 7),
            foreground="gray"
        ).pack(side=tk.BOTTOM, pady=2)

    def add_compact_param(self, frame, param_name, label_text, row, col, dependency=None):
        container = ttk.Frame(frame)
        container.grid(row=row, column=col, sticky="ew", padx=2, pady=1)
        
        ttk.Label(container, text=label_text, width=14, anchor="e").pack(side=tk.LEFT)
        
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

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def setup_styles(self):
        style = ttk.Style()
        style.configure("TLabelFrame", font=('Arial', 8, 'bold'))
        style.configure("Accent.TButton", foreground="white", background="#4CAF50", font=('Arial', 9))
        style.configure("Secondary.TButton", font=('Arial', 9))
        style.map("Accent.TButton", background=[('active', '#45a049')])

    def add_checkbox(self, frame, param_name, text, row):
        var = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(frame, text=text, variable=var, style='Toolbutton')
        cb.grid(row=row, column=0, sticky="w", pady=1, padx=5)
        self.params[param_name] = var
        return var

    def reset_settings(self):
        for var in self.params.values():
            if isinstance(var, tk.DoubleVar):
                var.set(0.0)
            elif isinstance(var, tk.BooleanVar):
                var.set(False)

    def generate_gcode(self):
        try:
            params = {name: var.get() for name, var in self.params.items()}

            if params["scan_length"] <= 0:
                raise ValueError("Длина сканирования должна быть > 0")
            if params["probe_depth"] <= 0:
                raise ValueError("Глубина зондирования должна быть > 0")

            gcode = [
                "(*** СКАНИРОВАНИЕ ***)",
                "M40",
                f"F{params['speed']}",
                "(Подведите щуп к краю диска и нажмите СТАРТ)",
                "M00",
                "G91"
            ]

            y_pos = 0.0
            if params["use_start_zone"] and params["start_zone"] > 0:
                while y_pos < params["start_zone"]:
                    gcode.append(f"G31X{params['probe_depth']}")
                    gcode.append(f"G0X-{params['retract']}")
                    if y_pos + params["start_step"] <= params["start_zone"]:
                        gcode.append(f"G0Y{params['start_step']}")
                        y_pos += params["start_step"]

            main_zone_end = params["scan_length"] - (params["end_zone"] if params["use_end_zone"] else 0)
            while y_pos < main_zone_end:
                gcode.append(f"G31X{params['probe_depth']}")
                gcode.append(f"G0X-{params['retract']}")
                if y_pos + params["main_step"] <= main_zone_end:
                    gcode.append(f"G0Y{params['main_step']}")
                    y_pos += params["main_step"]

            if params["use_end_zone"] and params["end_zone"] > 0:
                while y_pos < params["scan_length"]:
                    gcode.append(f"G31X{params['probe_depth']}")
                    gcode.append(f"G0X-{params['retract']}")
                    if y_pos + params["end_step"] <= params["scan_length"]:
                        gcode.append(f"G0Y{params['end_step']}")
                        y_pos += params["end_step"]

            gcode.extend([
                "G90",
                "G0X0",
                "G0Y0",
                "M30~"
            ])

            filename = f"scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.tap"
            with open(os.path.join(WORK_FOLDER, filename), "w", encoding="utf-8") as f:
                f.write("\n".join(gcode))

            messagebox.showinfo(
                "Готово",
                f"G-код сохранен в:\n{filename}\n\n"
                f"Тип файла: .tap\n"
                f"Общая длина сканирования: {y_pos:.2f} мм"
            )
        except Exception as e:
            messagebox.showerror("Ошибка генерации", str(e))

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
            current_x, current_y = 0.0, 0.0
            with open(points_file, 'r') as f:
                for line in f:
                    parts = line.strip().split()
                    if len(parts) >= 2:
                        current_x += float(parts[0])
                        current_y += float(parts[1])
                        points.append((current_x, current_y))

            if len(points) < 2:
                raise ValueError("Недостаточно точек для создания файла")
            
            filename = f"scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            output_file = os.path.join(WORK_FOLDER, filename)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write("0\nSECTION\n")
                f.write("2\nENTITIES\n")
                f.write("0\nPOLYLINE\n")
                f.write("8\n0\n")
                
                for x, y in points:
                    f.write("0\nVERTEX\n")
                    f.write("8\n0\n")
                    f.write(f"10\n{x:.5f}\n")
                    f.write(f"20\n{y:.5f}\n")
                    f.write("30\n0.00000\n")
                    f.write("70\n32\n")
                    f.write("0\n")
                
                f.write("0\nSEQEND\n")

            messagebox.showinfo(
                "Готово",
                f"Файл создан:\n{filename}"
            )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = IvanorProScan(root)
    root.mainloop()
