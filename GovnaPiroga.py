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
        self.root.geometry("850x800")
        self.root.resizable(False, False)
        self.center_window()
        
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        self.setup_styles()

        # Основные параметры
        params_frame = ttk.LabelFrame(main_frame, text="ОСНОВНЫЕ ПАРАМЕТРЫ СКАНИРОВАНИЯ", padding=10)
        params_frame.pack(fill=tk.X, pady=5)
        
        self.add_param(params_frame, "scan_length", "Общая длина сканирования (мм):", 0)
        self.add_param(params_frame, "main_step", "Шаг основной зоны (мм):", 1)
        self.add_param(params_frame, "probe_depth", "Глубина зондирования (мм):", 2)
        self.add_param(params_frame, "retract", "Величина отвода (мм):", 3)
        self.add_param(params_frame, "speed", "Скорость сканирования (мм/мин):", 4)

        # Стартовая зона
        start_zone_frame = ttk.LabelFrame(main_frame, text="СТАРТОВАЯ ЗОНА", padding=10)
        start_zone_frame.pack(fill=tk.X, pady=5)

        self.add_checkbox(start_zone_frame, "use_start_zone", "Активировать стартовую зону", 0)
        self.add_param(start_zone_frame, "start_zone", "Длина стартовой зоны (мм):", 1, "use_start_zone")
        self.add_param(start_zone_frame, "start_step", "Шаг в стартовой зоне (мм):", 2, "use_start_zone")

        # Конечная зона
        end_zone_frame = ttk.LabelFrame(main_frame, text="КОНЕЧНАЯ ЗОНА", padding=10)
        end_zone_frame.pack(fill=tk.X, pady=5)

        self.add_checkbox(end_zone_frame, "use_end_zone", "Активировать конечную зону", 0)
        self.add_param(end_zone_frame, "end_zone", "Длина конечной зоны (мм):", 1, "use_end_zone")
        self.add_param(end_zone_frame, "end_step", "Шаг в конечной зоне (мм):", 2, "use_end_zone")

        # Кнопки управления
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=15)

        ttk.Button(
            btn_frame,
            text="СБРОСИТЬ НАСТРОЙКИ",
            style="Secondary.TButton",
            command=self.reset_settings
        ).pack(side=tk.LEFT, expand=True, padx=5)

        ttk.Button(
            btn_frame,
            text="ГЕНЕРИРОВАТЬ G-КОД",
            style="Accent.TButton",
            command=self.generate_gcode
        ).pack(side=tk.LEFT, expand=True, padx=5)

        # Кнопка для ArtCAM
        ttk.Button(
            main_frame,
            text="СОЗДАТЬ ФАЙЛ ДЛЯ ARTCAM",
            style="Accent.TButton",
            command=self.create_artcam_file
        ).pack(fill=tk.X, pady=10)

        # Подпись
        ttk.Label(
            main_frame,
            text="ivanor inc. | Версия 3.4 | Чистый G-код без комментариев",
            font=('Arial', 8),
            foreground="gray"
        ).pack(side=tk.BOTTOM, pady=5)

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def setup_styles(self):
        style = ttk.Style()
        style.configure("TLabelFrame", font=('Arial', 9, 'bold'))
        style.configure("Accent.TButton", foreground="white", background="#4CAF50", font=('Arial', 10, 'bold'))
        style.configure("Secondary.TButton", background="#f0f0f0", font=('Arial', 10))
        style.map("Accent.TButton", background=[('active', '#45a049')])

    def add_param(self, frame, param_name, label_text, row, dependency=None):
        container = ttk.Frame(frame)
        container.grid(row=row, column=0, sticky="ew", pady=3)
        
        ttk.Label(container, text=label_text, width=28, anchor="e").pack(side=tk.LEFT, padx=5)
        
        var = tk.DoubleVar(value=0.0)
        entry = ttk.Entry(container, textvariable=var, width=12)
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
        var = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(frame, text=text, variable=var)
        cb.grid(row=row, column=0, sticky="w", pady=2, padx=5)
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

            # Валидация параметров
            if params["scan_length"] <= 0:
                raise ValueError("Длина сканирования должна быть > 0")
            if params["probe_depth"] <= 0:
                raise ValueError("Глубина зондирования должна быть > 0")
            if params["use_start_zone"] and params["start_step"] <= 0:
                raise ValueError("Шаг стартовой зоны должен быть > 0")
            if params["use_end_zone"] and params["end_step"] <= 0:
                raise ValueError("Шаг конечной зоны должен быть > 0")

            # Генерация чистого G-кода без комментариев
            gcode = [
                "M40",
                f"F{params['speed']}",
                "M00",
                "G91"
            ]

            y_pos = 0.0

            # Стартовая зона
            if params["use_start_zone"] and params["start_zone"] > 0:
                while y_pos < params["start_zone"]:
                    gcode.append(f"G31X{params['probe_depth']}")
                    gcode.append(f"G0X-{params['retract']}")
                    if y_pos + params["start_step"] <= params["start_zone"]:
                        gcode.append(f"G0Y{params['start_step']}")
                        y_pos += params["start_step"]

            # Основная зона
            main_zone_end = params["scan_length"] - (params["end_zone"] if params["use_end_zone"] else 0)
            while y_pos < main_zone_end:
                gcode.append(f"G31X{params['probe_depth']}")
                gcode.append(f"G0X-{params['retract']}")
                if y_pos + params["main_step"] <= main_zone_end:
                    gcode.append(f"G0Y{params['main_step']}")
                    y_pos += params["main_step"]

            # Конечная зона
            if params["use_end_zone"] and params["end_zone"] > 0:
                while y_pos < params["scan_length"]:
                    gcode.append(f"G31X{params['probe_depth']}")
                    gcode.append(f"G0X-{params['retract']}")
                    if y_pos + params["end_step"] <= params["scan_length"]:
                        gcode.append(f"G0Y{params['end_step']}")
                        y_pos += params["end_step"]

            # Завершение программы
            gcode.extend([
                "G90",
                "G0X0Y0",
                "M30~"
            ])

            # Сохранение файла
            filename = f"scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.nc"
            with open(os.path.join(WORK_FOLDER, filename), "w", encoding="utf-8") as f:
                f.write("\n".join(gcode))

            messagebox.showinfo(
                "Готово",
                f"G-код сохранен в:\n{filename}\n\n"
                f"Тип файла: Чистый G-код без комментариев\n"
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
            
            filename = f"artcam_scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            output_file = os.path.join(WORK_FOLDER, filename)
            
            with open(output_file, 'w', newline='') as f:
                writer = csv.writer(f, delimiter=',')
                writer.writerow(["Type", "X", "Y", "Z", "Bulge", "Weight", "Layer"])
                writer.writerow(["Point", f"{points[0][0]:.4f}", f"{points[0][1]:.4f}", "0.0", "0.0", "1", "0"])
                for point in points[1:]:
                    writer.writerow(["LineTo", f"{point[0]:.4f}", f"{point[1]:.4f}", "0.0", "0.0", "1", "0"])

            messagebox.showinfo(
                "Готово",
                f"Файл для ArtCAM создан:\n{filename}\n\n"
                "Как использовать:\n"
                "1. Откройте файл в ArtCAM\n"
                "2. Точки будут соединены автоматически"
            )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = IvanorProScan(root)
    root.mainloop()