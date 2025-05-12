# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from datetime import datetime
import ezdxf

class COXOproScan:
    def __init__(self, root):
        self.root = root
        self.params = {}
        self.setup_ui()
        self.reset_settings()
        
    def setup_ui(self):
        self.root.title("COXOproScan v3.4")
        self.root.geometry("900x700")
        self.root.resizable(False, False)
        self.center_window()
        
        # Настройка стилей
        style = ttk.Style()
        style.configure("TLabelFrame", font=('Arial', 9, 'bold'), padding=5)
        style.configure("Custom.TButton", 
                      foreground="black",
                      background="#f0f0f0",
                      font=('Arial', 10, 'bold'),
                      padding=3,
                      borderwidth=1)
        style.map("Custom.TButton",
                 foreground=[('active', 'black'), ('disabled', 'gray')],
                 background=[('active', '#e0e0e0'), ('disabled', '#cccccc')])

        # Основной контейнер
        main_container = ttk.Frame(self.root, padding=3)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Верхняя часть (параметры + изображение)
        top_panel = ttk.Frame(main_container)
        top_panel.pack(fill=tk.BOTH, expand=True, pady=(0,5))

        # Левая панель параметров
        left_panel = ttk.Frame(top_panel, width=400)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0,5))

        # Основные параметры
        main_frame = ttk.LabelFrame(left_panel, text="ОСНОВНЫЕ ПАРАМЕТРЫ", padding=5)
        main_frame.pack(fill=tk.X, pady=3)
        
        self.add_param(main_frame, "scan_length", "Общая длина (мм):", 0)
        self.add_param(main_frame, "retract", "Отвод (мм):", 1)
        self.add_param(main_frame, "speed", "Скорость (мм/мин):", 2)
        
        # Дополнительные параметры
        self.additional_params_visible = True
        self.probe_depth_frame = ttk.Frame(left_panel)
        self.probe_depth_frame.pack(fill=tk.X, pady=0)
        
        self.additional_params_header = ttk.Label(
            self.probe_depth_frame, 
            text="▼ ДОПОЛНИТЕЛЬНЫЕ ПАРАМЕТРЫ ▼",
            padding=3,
            font=('Arial', 9, 'bold'),
            relief="groove",
            cursor="hand2"
        )
        self.additional_params_header.pack(fill=tk.X)
        self.additional_params_header.bind("<Button-1>", self.toggle_additional_params)
        
        self.additional_params_container = ttk.Frame(self.probe_depth_frame)
        self.additional_params_container.pack(fill=tk.X)
        self.add_param(self.additional_params_container, "probe_depth", "Глубина (мм):", 0)

        # Зоны сканирования
        start_frame = ttk.LabelFrame(left_panel, text="СТАРТОВАЯ ЗОНА", padding=5)
        start_frame.pack(fill=tk.X, pady=3)
        
        self.add_checkbox(start_frame, "use_start_zone", "Активировать", 0)
        self.add_param(start_frame, "start_zone_length", "Длина (мм):", 1, "use_start_zone")
        self.add_param(start_frame, "start_zone_step", "Шаг (мм):", 2, "use_start_zone")

        main_zone_frame = ttk.LabelFrame(left_panel, text="ОСНОВНАЯ ЗОНА", padding=5)
        main_zone_frame.pack(fill=tk.X, pady=3)
        self.add_param(main_zone_frame, "main_zone_step", "Шаг (мм):", 0)

        end_frame = ttk.LabelFrame(left_panel, text="КОНЕЧНАЯ ЗОНА", padding=5)
        end_frame.pack(fill=tk.X, pady=3)
        
        self.add_checkbox(end_frame, "use_end_zone", "Активировать", 0)
        self.add_param(end_frame, "end_zone_length", "Длина (мм):", 1, "use_end_zone")
        self.add_param(end_frame, "end_zone_step", "Шаг (мм):", 2, "use_end_zone")

        # Правая панель изображения
        right_panel = ttk.Frame(top_panel)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Заглушка вместо изображения
        img_placeholder = ttk.Label(
            right_panel, 
            text="Здесь будет изображение схемы сканирования",
            background="#f0f0f0",
            justify="center",
            font=('Arial', 10)
        )
        img_placeholder.pack(fill=tk.BOTH, expand=True)

        # Нижняя панель кнопок
        btn_frame = ttk.Frame(main_container)
        btn_frame.pack(fill=tk.X, pady=(5,0))
        
        ttk.Button(
            btn_frame,
            text="СБРОСИТЬ",
            style="Custom.TButton",
            command=self.reset_settings
        ).pack(side=tk.LEFT, expand=True, padx=2)

        ttk.Button(
            btn_frame,
            text="G-КОД",
            style="Custom.TButton",
            command=self.generate_gcode
        ).pack(side=tk.LEFT, expand=True, padx=2)

        ttk.Button(
            btn_frame,
            text="DXF ДЛЯ ARTCAM",
            style="Custom.TButton",
            command=self.create_artcam_file
        ).pack(side=tk.LEFT, expand=True, padx=2)

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def add_param(self, frame, param_name, label_text, row, dependency=None):
        container = ttk.Frame(frame)
        container.pack(fill=tk.X, pady=1)
        ttk.Label(container, text=label_text, width=20, anchor="w").pack(side=tk.LEFT)
        var = tk.DoubleVar(value=0.0)
        entry = ttk.Entry(container, textvariable=var, width=10)
        entry.pack(side=tk.LEFT, padx=2)
        self.params[param_name] = var
        if dependency:
            self.toggle_dependency(entry, dependency)
        return var

    def add_checkbox(self, frame, param_name, text, row):
        var = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(frame, text=text, variable=var)
        cb.pack(anchor="w", pady=1)
        self.params[param_name] = var
        return var

    def toggle_dependency(self, widget, dependency):
        widget.state(["!disabled" if self.params[dependency].get() else "disabled"])
        self.params[dependency].trace("w", lambda *_, w=widget, d=dependency: 
            w.state(["!disabled" if self.params[d].get() else "disabled"]))

    def toggle_additional_params(self, event=None):
        self.additional_params_visible = not self.additional_params_visible
        if self.additional_params_visible:
            self.additional_params_container.pack(fill=tk.X)
            self.additional_params_header.config(text="▼ ДОПОЛНИТЕЛЬНЫЕ ПАРАМЕТРЫ ▼")
        else:
            self.additional_params_container.pack_forget()
            self.additional_params_header.config(text="▶ ДОПОЛНИТЕЛЬНЫЕ ПАРАМЕТРЫ ▶")
        self.root.update_idletasks()

    def reset_settings(self):
        current_probe_depth = self.params["probe_depth"].get()
        defaults = {
            "scan_length": 0.0, "retract": 0.0, "speed": 0.0,
            "use_start_zone": False, "start_zone_length": 0.0, "start_zone_step": 0.0,
            "main_zone_step": 0.0, "use_end_zone": False, "end_zone_length": 0.0,
            "end_zone_step": 0.0, "probe_depth": current_probe_depth
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
                errors.append("Длина стартовой зоны должна быть меньше общей длины")
            if params["end_zone_length"] >= params["scan_length"]:
                errors.append("Длина конечной зоны должна быть меньше общей длины")
            if (params["start_zone_length"] + params["end_zone_length"]) >= params["scan_length"]:
                errors.append("Сумма длин стартовой и конечной зон должна быть меньше общей длины")
            
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

            gcode.extend([
                "G90",
                "G0X0",
                "G0Y0",
                "(* выберете название файла сканирования *)",
                "M30",
                ""
            ])

            filename = f"scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.tap"
            filepath = os.path.join(os.path.expanduser("~"), "Desktop", "COXOproScan", filename)
            
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            with open(filepath, "w", encoding='cp1251') as f:
                f.write("\n".join(gcode))

            messagebox.showinfo(
                "Готово!",
                f"G-код сохранён в:\n{filepath}\n"
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

            # Чтение и обработка точек
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

            # Создание чистого DXF
            doc = ezdxf.new('R2010')
            msp = doc.modelspace()
            
            # Добавляем 3D полилинию
            polyline = msp.add_polyline3d(points)
            
            # Упрощаем файл - оставляем только необходимые данные
            doc.header['$ACADVER'] = 'AC1021'  # Версия AutoCAD 2010
            doc.header['$INSBASE'] = (0, 0, 0)
            doc.header['$EXTMIN'] = (min(p[0] for p in points), min(p[1] for p in points), min(p[2] for p in points))
            doc.header['$EXTMAX'] = (max(p[0] for p in points), max(p[1] for p in points), max(p[2] for p in points))
            
            # Удаляем все ненужные слои
            for layer in list(doc.layers):
                if layer.dxf.name != '0':
                    doc.layers.remove(layer.dxf.name)
            
            filename = f"artcam_{datetime.now().strftime('%Y%m%d_%H%M%S')}.dxf"
            filepath = os.path.join(os.path.expanduser("~"), "Desktop", "COXOproScan", filename)
            doc.saveas(filepath)

            messagebox.showinfo(
                "Готово!",
                f"DXF файл успешно создан:\n{filepath}\n"
                f"Количество точек: {len(points)}"
            )
        except ValueError as ve:
            messagebox.showerror("Ошибка данных", str(ve))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = COXOproScan(root)
    app.params["probe_depth"].set(20.0)  # Установка глубины зондирования по умолчанию
    root.mainloop()
