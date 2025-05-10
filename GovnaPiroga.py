# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import base64
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
        self.root.geometry("900x750")
        self.root.resizable(False, False)
        self.center_window()
        
        # Настройка стилей
        style = ttk.Style()
        style.configure("TLabelFrame", font=('Arial', 9, 'bold'))
        style.configure("Custom.TButton", 
                      foreground="black",
                      background="#f0f0f0",
                      font=('Arial', 10, 'bold'),
                      padding=5,
                      borderwidth=1)
        style.map("Custom.TButton",
                 foreground=[('active', 'black'), ('disabled', 'gray')],
                 background=[('active', '#e0e0e0'), ('disabled', '#cccccc')])

        # Основной контейнер
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Левая панель (параметры)
        left_panel = ttk.Frame(main_container, width=400)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)
        
        # Основные параметры
        main_frame = ttk.LabelFrame(left_panel, text="ОСНОВНЫЕ ПАРАМЕТРЫ", padding=10)
        main_frame.pack(fill=tk.X, pady=5, anchor="w")
        
        self.add_param(main_frame, "scan_length", "Общая длина сканирования (мм):", 0)
        self.add_param(main_frame, "retract", "Величина отвода (мм):", 1)
        self.add_param(main_frame, "speed", "Скорость (мм/мин):", 2)
        
        # Дополнительные параметры (сворачиваемые)
        self.additional_params_visible = True
        self.probe_depth_frame = ttk.Frame(left_panel)
        self.probe_depth_frame.pack(fill=tk.X, pady=0, anchor="w")
        
        self.additional_params_header = ttk.Label(
            self.probe_depth_frame, 
            text="▼ ДОПОЛНИТЕЛЬНЫЕ ПАРАМЕТРЫ ▼",
            padding=5,
            font=('Arial', 9, 'bold'),
            relief="groove",
            cursor="hand2"
        )
        self.additional_params_header.pack(fill=tk.X)
        self.additional_params_header.bind("<Button-1>", self.toggle_additional_params)
        
        self.additional_params_container = ttk.Frame(self.probe_depth_frame)
        self.additional_params_container.pack(fill=tk.X)
        self.add_param(self.additional_params_container, "probe_depth", "Глубина зондирования (мм):", 0)

        # Стартовая зона
        start_frame = ttk.LabelFrame(left_panel, text="СТАРТОВАЯ ЗОНА", padding=10)
        start_frame.pack(fill=tk.X, pady=5, anchor="w")
        
        self.add_checkbox(start_frame, "use_start_zone", "Активировать стартовую зону", 0)
        self.add_param(start_frame, "start_zone_length", "Длина зоны (мм):", 1, "use_start_zone")
        self.add_param(start_frame, "start_zone_step", "Шаг (мм):", 2, "use_start_zone")

        # Основная зона
        main_zone_frame = ttk.LabelFrame(left_panel, text="ОСНОВНАЯ ЗОНА", padding=10)
        main_zone_frame.pack(fill=tk.X, pady=5, anchor="w")
        
        self.add_param(main_zone_frame, "main_zone_step", "Шаг (мм):", 0)

        # Конечная зона
        end_frame = ttk.LabelFrame(left_panel, text="КОНЕЧНАЯ ЗОНА", padding=10)
        end_frame.pack(fill=tk.X, pady=5, anchor="w")
        
        self.add_checkbox(end_frame, "use_end_zone", "Активировать конечную зону", 0)
        self.add_param(end_frame, "end_zone_length", "Длина зоны (мм):", 1, "use_end_zone")
        self.add_param(end_frame, "end_zone_step", "Шаг (мм):", 2, "use_end_zone")

        # Кнопки
        btn_frame = ttk.Frame(left_panel)
        btn_frame.pack(fill=tk.X, pady=10, anchor="w")
        
        ttk.Button(
            btn_frame,
            text="СБРОСИТЬ НАСТРОЙКИ",
            style="Custom.TButton",
            command=self.reset_settings
        ).pack(side=tk.TOP, fill=tk.X, pady=5)

        ttk.Button(
            btn_frame,
            text="ГЕНЕРИРОВАТЬ G-КОД",
            style="Custom.TButton",
            command=self.generate_gcode
        ).pack(side=tk.TOP, fill=tk.X, pady=5)

        ttk.Button(
            btn_frame,
            text="СОЗДАТЬ DXF ДЛЯ ARTCAM",
            style="Custom.TButton",
            command=self.create_artcam_file
        ).pack(side=tk.TOP, fill=tk.X, pady=5)

        # Правая панель (изображение GIF в Base64)
        right_panel = ttk.Frame(main_container)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10)

        #######################################################
        # ВСТАВЬТЕ СВОЙ BASE64-GIF КОД НИЖЕ (ЗАМЕНИТЕ ЭТУ СТРОКУ)
        gif_base64 = "R0lGODlhwgHCAfAAAAAAAAAAACH5BAAAAAAAIf8LTUdLOEJJTTAwMDD/OEJJTQQlAAAAAAAQAAAAAAAAAAAAAAAAAAAAADhCSU0EOgAAAAAA9wAAABAAAAABAAAAAAALcHJpbnRPdXRwdXQAAAAFAAAAAFBzdFNib29sAQAAAABJbnRlZW51bQAAAABJbnRlAAAAAENscm0AAAAPcHJpbnRTaXh0ZWVuQml0Ym9vbAAAAAALcHJpbnRlck5hbWVURVhUAAAAAQAAAAAAD3ByaW50UHJvb2ZTZXR1cE9iamMAAAAVBB8EMARABDAEPAQ1BEIEQARLACAERgQyBDUEQgQ+BD8EQAQ+BDEESwAAAAAACnByb29mU2V0dXAAAAABAAAAAEJsdG5l/251bQAAAAxidWlsdGluUHJvb2YAAAAJcHJvb2ZDTVlLADhCSU0EOwAAAAACLQAAABAAAAABAAAAAAAScHJpbnRPdXRwdXRPcHRpb25zAAAAFwAAAABDcHRuYm9vbAAAAAAAQ2xicmJvb2wAAAAAAFJnc01ib29sAAAAAABDcm5DYm9vbAAAAAAAQ250Q2Jvb2wAAAAAAExibHNib29sAAAAAABOZ3R2Ym9vbAAAAAAARW1sRGJvb2wAAAAAAEludHJib29sAAAAAABCY2tnT2JqYwAAAAEAAAAAAABSR0JDAAAAAwAAAABSZCAgZG91YkBv4AAAAAAAAAAAAEdybv8gZG91YkBv4AAAAAAAAAAAAEJsICBkb3ViQG/gAAAAAAAAAAAAQnJkVFVudEYjUmx0AAAAAAAAAAAAAAAAQmxkIFVudEYjUmx0AAAAAAAAAAAAAAAAUnNsdFVudEYjUHhsQHLAAAAAAAAAAAAKdmVjdG9yRGF0YWJvb2wBAAAAAFBnUHNlbnVtAAAAAFBnUHMAAAAAUGdQQwAAAABMZWZ0VW50RiNSbHQAAAAAAAAAAAAAAABUb3AgVW50RiNSbHQAAAAAAAAAAAAAAABTY2wgVW50RiNQcmNAWQAAAAAAAAAAABBjcm9wV2hlblByaW50aW5nYm9vbAAAAAAOY3L/b3BSZWN0Qm90dG9tbG9uZwAAAAAAAAAMY3JvcFJlY3RMZWZ0bG9uZwAAAAAAAAANY3JvcFJlY3RSaWdodGxvbmcAAAAAAAAAC2Nyb3BSZWN0VG9wbG9uZwAAAAAAOEJJTQPtAAAAAAAQASwAAAABAAIBLAAAAAEAAjhCSU0EJgAAAAAADgAAAAAAAAAAAAA/gAAAOEJJTQQNAAAAAAAEAAAAWjhCSU0EGQAAAAAABAAAAB44QklNA/MAAAAAAAkAAAAAAAAAAAEAOEJJTScQAAAAAAAKAAEAAAAAAAAAAjhCSU0D9QAAAAAASAAvZmYAAQBsZmYABgAAAAAAAQAv/2ZmAAEAoZmaAAYAAAAAAAEAMgAAAAEAWgAAAAYAAAAAAAEANQAAAAEALQAAAAYAAAAAAAE4QklNA/gAAAAAAHAAAP////////////////////////////8D6AAAAAD/////////////////////////////A+gAAAAA/////////////////////////////wPoAAAAAP////////////////////////////8D6AAAOEJJTQQIAAAAAAAQAAAAAQAAAkAAAAJAAAAAADhCSU0EHgAAAAAABAAAAAA4QklNBBoAAAAAA0sAAAAGAAAAAAAAAAAAAAHCAAABwgAAAP8LBBEENQQ3ACAEOAQ8BDUEPQQ4AC0AMQAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAABwgAAAcIAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAQAAAAAAAG51bGwAAAACAAAABmJvdW5kc09iamMAAAABAAAAAAAAUmN0MQAAAAQAAAAAVG9wIGxvbmcAAAAAAAAAAExlZnRsb25nAAAAAAAAAABCdG9tbG9uZwAAAcIAAAAAUmdodGxvbmcAAAHCAAAABnNsaWNlc1ZsTHMAAAABT2JqYwAAAAEAAAAAAAVzbGljZQAAABL/AAAAB3NsaWNlSURsb25nAAAAAAAAAAdncm91cElEbG9uZwAAAAAAAAAGb3JpZ2luZW51bQAAAAxFU2xpY2VPcmlnaW4AAAANYXV0b0dlbmVyYXRlZAAAAABUeXBlZW51bQAAAApFU2xpY2VUeXBlAAAAAEltZyAAAAAGYm91bmRzT2JqYwAAAAEAAAAAAABSY3QxAAAABAAAAABUb3AgbG9uZwAAAAAAAAAATGVmdGxvbmcAAAAAAAAAAEJ0b21sb25nAAABwgAAAABSZ2h0bG9uZwAAAcIAAAADdXJs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"  # Ваш Base64-GIF
        #######################################################

        try:
            gif_data = base64.b64decode(gif_base64)
            self.tk_image = tk.PhotoImage(data=gif_data)
            img_label = ttk.Label(right_panel, image=self.tk_image)
            img_label.pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            ttk.Label(
                right_panel, 
                text="Ошибка загрузки изображения\n" + str(e),
                background="#f0f0f0",
                justify="center"
            ).pack(fill=tk.BOTH, expand=True)

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def add_param(self, frame, param_name, label_text, row, dependency=None):
        container = ttk.Frame(frame)
        container.pack(fill=tk.X, pady=2, anchor="w")
        ttk.Label(container, text=label_text, width=30, anchor="w").pack(side=tk.LEFT)
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
        cb.pack(anchor="w", pady=2)
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
                f"Стартовая зона: {params['start_zone_length']} мм (шаг: {params['start_zone_step']} мм)\n"
                f"Основная зона: {main_zone_length:.2f} мм (шаг: {params['main_zone_step']} мм)\n"
                f"Конечная зона: {params['end_zone_length']} мм (шаг: {params['end_zone_step']} мм)"
            )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def create_artcam_file(self):
        try:
            points_file = filedialog.askopenfilename(
                initialdir=os.path.join(os.path.expanduser("~"), "Desktop", "COXOproScan"),
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

            doc = ezdxf.new('R2010')
            msp = doc.modelspace()
            polyline = msp.add_polyline2d(points)

            filename = f"artcam_{datetime.now().strftime('%Y%m%d_%H%M%S')}.dxf"
            filepath = os.path.join(os.path.expanduser("~"), "Desktop", "COXOproScan", filename)
            doc.saveas(filepath)

            messagebox.showinfo(
                "Готово!",
                f"DXF создан:\n{filepath}\n"
                f"Точек: {len(points)}"
            )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = COXOproScan(root)
    app.params["probe_depth"].set(20.0)
    root.mainloop()
