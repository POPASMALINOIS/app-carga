#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LogiDesk Win v1.2 - Import Wizard + UI más intuitiva
"""
import os, sys, glob, json, datetime as dt, tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd

APP_NAME = "LogiDesk Win"
BASE_DIR = os.path.abspath(".")
HISTORY_DIR = os.path.join(BASE_DIR, "history")
REPORT_DIR  = os.path.join(BASE_DIR, "reports")
SESSION     = os.path.join(BASE_DIR, "current_session.csv")
PROFILES    = os.path.join(BASE_DIR, "import_profiles.json")

REQUIRED_COLS = [
    "TRANSPORTISTA","MATRICULA","MUELLE","ESTADO","DESTINO",
    "LLEGADA","LLEGADA REAL","SALIDA REAL","SALIDA TOPE","OBSERVACIONES","INCIDENCIAS"
]

ALIASES = {
    "TRANSPORTISTA": ["TRANSPORTISTA"],
    "MATRICULA": ["MATRICULA","MATRÍCULA","MAT."],
    "MUELLE": ["MUELLE","DOCK","RAMPA"],
    "ESTADO": ["ESTADO","STATUS"],
    "DESTINO": ["DESTINO","DEST."],
    "LLEGADA": ["LLEGADA","HORA LLEGADA","ETA"],
    "LLEGADA REAL": ["LLEGADA REAL","LLEGADA_REAL","CHECK-IN","ENTRADA REAL"],
    "SALIDA REAL": ["SALIDA REAL","SALIDA_REAL","CHECK-OUT"],
    "SALIDA TOPE": ["SALIDA TOPE","SALIDA_TOPE","CUT-OFF","HORA TOPE"],
    "OBSERVACIONES": ["OBSERVACIONES","OBS.","NOTAS"],
    "INCIDENCIAS": ["INCIDENCIAS","INC.","EVENTOS"],
}

def load_profiles():
    if os.path.exists(PROFILES):
        try:
            with open(PROFILES, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_profiles(p):
    try:
        with open(PROFILES, "w", encoding="utf-8") as f:
            json.dump(p, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

class ImportWizard(tk.Toplevel):
    def __init__(self, master, path):
        super().__init__(master)
        self.title("Asistente de importación")
        self.geometry("700x520")
        self.resizable(False, False)
        self.path = path
        self.result_df = None
        self.profiles = load_profiles()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill="both", expand=True)

        self.var_sheet = tk.StringVar()
        sheets = []
        ext = os.path.splitext(path)[1].lower()
        if ext in [".xlsx",".xls"]:
            try:
                xl = pd.ExcelFile(path)
                sheets = xl.sheet_names
            except Exception as e:
                messagebox.showerror(APP_NAME, f"No se pudo leer el Excel: {e}")
        ttk.Label(frame, text="1) Hoja:").grid(row=0, column=0, sticky="w")
        self.cmb_sheet = ttk.Combobox(frame, textvariable=self.var_sheet, values=sheets, state="readonly", width=40)
        self.cmb_sheet.grid(row=0, column=1, sticky="w")
        if sheets:
            self.cmb_sheet.current(0)

        ttk.Label(frame, text="2) Fila de encabezados (1,2,3…):").grid(row=1, column=0, sticky="w", pady=(10,0))
        self.ent_header = ttk.Entry(frame, width=10); self.ent_header.insert(0, "1"); self.ent_header.grid(row=1, column=1, sticky="w", pady=(10,0))
        ttk.Label(frame, text="3) Fila donde empiezan los datos:").grid(row=2, column=0, sticky="w")
        self.ent_start = ttk.Entry(frame, width=10); self.ent_start.insert(0, "2"); self.ent_start.grid(row=2, column=1, sticky="w")

        ttk.Button(frame, text="Previsualizar columnas", command=self.preview).grid(row=3, column=0, columnspan=2, pady=10)

        self.list_cols = tk.Listbox(frame, height=8, width=60); self.list_cols.grid(row=4, column=0, columnspan=2, sticky="w")

        ttk.Label(frame, text="4) Mapeo de columnas a estándar:").grid(row=5, column=0, sticky="w", pady=(10,0))
        self.map_frame = ttk.Frame(frame); self.map_frame.grid(row=6, column=0, columnspan=2, sticky="w")

        btns = ttk.Frame(frame); btns.grid(row=7, column=0, columnspan=2, pady=12, sticky="e")
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right", padx=6)
        ttk.Button(btns, text="Importar", command=self.finish).pack(side="right")

        helpbox = ttk.LabelFrame(self, text="Ayuda", padding=10)
        helpbox.pack(fill="x", padx=10, pady=(0,10))
        ttk.Label(helpbox, text=(
            "• Si tu Excel tiene títulos en la fila 2 o 3, pon ese número en 'Fila de encabezados'.\n"
            "• 'Fila donde empiezan los datos' es la primera fila REAL con registros.\n"
            "• Las columnas 'Unnamed' o vacías se descartan automáticamente.\n"
            "• Asigna las columnas del archivo a TRANSPORTISTA, MATRICULA, etc.\n"
            "• El mapeo se guarda y se reutilizará la próxima vez para ese archivo/hoja."
        )).pack(anchor="w")

    def _read_df(self):
        header_visible = int(self.ent_header.get())
        start_visible  = int(self.ent_start.get())
        header_idx = max(header_visible - 1, 0)
        skiprows = list(range(0, start_visible - 1)) if start_visible - 1 > 0 else None
        ext = os.path.splitext(self.path)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(self.path, sep=None, engine="python", header=header_idx, skiprows=skiprows)
        else:
            sheet = self.var_sheet.get() if self.var_sheet.get() else 0
            df = pd.read_excel(self.path, sheet_name=sheet, header=header_idx, skiprows=skiprows)
        # limpiar columnas unnamed
        new_cols = []
        for c in df.columns:
            name = str(c).strip()
            if name.startswith("Unnamed"):
                name = ""
            new_cols.append(name)
        df.columns = new_cols
        df = df.loc[:, [c for c in df.columns if str(c).strip() != ""]]
        df = df.dropna(how="all")
        return df

    def preview(self):
        try:
            df = self._read_df()
        except Exception as e:
            messagebox.showerror(APP_NAME, f"No se pudo previsualizar: {e}")
            return
        self.list_cols.delete(0, tk.END)
        for c in df.columns:
            self.list_cols.insert(tk.END, str(c))
        for child in self.map_frame.winfo_children():
            child.destroy()
        cols = [str(c) for c in df.columns]
        self.cmb_map = {}
        for i, std in enumerate(REQUIRED_COLS):
            ttk.Label(self.map_frame, text=std, width=16).grid(row=i, column=0, sticky="w")
            cmb = ttk.Combobox(self.map_frame, values=["--No importar--"] + cols, width=40, state="readonly")
            # autoselección simple por nombre exacto
            chosen = None
            for c in cols:
                if c.strip().upper() == std:
                    chosen = c; break
            cmb.set(chosen if chosen else "--No importar--")
            cmb.grid(row=i, column=1, sticky="w")
            self.cmb_map[std] = cmb

    def finish(self):
        try:
            df = self._read_df()
        except Exception as e:
            messagebox.showerror(APP_NAME, f"No se pudo leer el archivo: {e}")
            return
        out = {c: [] for c in REQUIRED_COLS}
        n = len(df)
        for std in REQUIRED_COLS:
            src = self.cmb_map[std].get()
            if src and src != "--No importar--" and src in df.columns:
                out[std] = df[src].astype(str).fillna("")
            else:
                out[std] = pd.Series([""]*n)
        self.result_df = pd.DataFrame(out)
        self.destroy()

class LogiDeskApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v1.2")
        self.geometry("1320x820")
        self.minsize(1100, 660)
        try:
            style = ttk.Style(self)
            if "vista" in style.theme_names():
                style.theme_use("vista")
        except Exception:
            pass

        self.df_orig = None
        self.df_view = None

        os.makedirs(HISTORY_DIR, exist_ok=True)
        os.makedirs(REPORT_DIR,  exist_ok=True)

        self._build_menu()
        self._build_toolbar()
        self._build_filterbar()
        self._build_table()
        self._build_help_status()

        self.protocol("WM_DELETE_WINDOW", self.on_exit)

    def _build_menu(self):
        m = tk.Menu(self)
        filem = tk.Menu(m, tearoff=0)
        filem.add_command(label="Cargar Excel/CSV (Asistente)", command=self.load_file)
        filem.add_separator()
        filem.add_command(label="Guardar sesión", command=self.save_session)
        filem.add_command(label="Cerrar día (a histórico)", command=self.close_day)
        filem.add_command(label="Exportar reporte del día (Excel)", command=self.export_daily_report)
        m.add_cascade(label="Archivo", menu=filem)
        actm = tk.Menu(m, tearoff=0)
        actm.add_command(label="➕ LLEGADA REAL", command=self.mark_llegada_real)
        actm.add_command(label="➕ SALIDA REAL", command=self.mark_salida_real)
        m.add_cascade(label="Acciones", menu=actm)
        self.config(menu=m)

    def _build_toolbar(self):
        tb = ttk.Frame(self, padding=(10,8)); tb.pack(fill="x")
        ttk.Button(tb, text="Cargar Excel/CSV (Asistente)", command=self.load_file).pack(side="left", padx=4)
        ttk.Button(tb, text="Guardar sesión", command=self.save_session).pack(side="left", padx=4)
        ttk.Button(tb, text="Cerrar día", command=self.close_day).pack(side="left", padx=4)
        ttk.Button(tb, text="Exportar Excel", command=self.export_daily_report).pack(side="left", padx=4)
        ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=12)
        ttk.Button(tb, text="➕ LLEGADA REAL", command=self.mark_llegada_real).pack(side="left", padx=4)
        ttk.Button(tb, text="➕ SALIDA REAL", command=self.mark_salida_real).pack(side="left", padx=4)

    def _build_filterbar(self):
        filt = ttk.LabelFrame(self, text="Filtros"); filt.pack(fill="x", padx=10, pady=(0,6))
        ttk.Label(filt, text="Columna:").pack(side="left", padx=(8,4))
        self.cmb_column = ttk.Combobox(filt, state="readonly", values=[]); self.cmb_column.pack(side="left", padx=4, ipadx=30)
        ttk.Label(filt, text="Valor contiene:").pack(side="left", padx=(12,4))
        self.ent_value = ttk.Entry(filt, width=30); self.ent_value.pack(side="left", padx=4)
        ttk.Button(filt, text="Aplicar", command=self.apply_filter).pack(side="left", padx=6)
        ttk.Button(filt, text="Limpiar", command=self.clear_filter).pack(side="left", padx=6)

    def _build_table(self):
        table = ttk.Frame(self); table.pack(fill="both", expand=True, padx=10, pady=(0,10))
        self.tree = ttk.Treeview(table, show="headings", selectmode="extended"); self.tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview); vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview); hsb.pack(fill="x", padx=10, pady=(0,10))
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.bind("<Double-1>", self._on_double_click_cell)
        import pandas as pd
        df = pd.DataFrame(columns=REQUIRED_COLS); self.df_orig=df.copy(); self.df_view=df.copy()
        self._refresh_columns_combobox(); self._populate_table(self.df_view)

    def _build_help_status(self):
        helpf = ttk.LabelFrame(self, text="Ayuda rápida"); helpf.pack(fill="x", padx=10, pady=(0,8))
        ttk.Label(helpf, text=(
            "1) Cargar Excel/CSV (Asistente) → elige hoja, pon fila de títulos y de datos; mapea columnas.\n"
            "2) Usa ➕ LLEGADA REAL / ➕ SALIDA REAL para sellar hora actual en las filas seleccionadas.\n"
            "3) Cerrar día → guarda snapshot (30 días). Exportar Excel → KPIs y resúmenes."
        )).pack(anchor="w")
        self.status = tk.StringVar(value="Listo."); ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w").pack(fill="x")

    # import
    def load_file(self):
        path = filedialog.askopenfilename(title="Selecciona un archivo", filetypes=[("Excel","*.xlsx *.xls"),("CSV","*.csv"),("Todos los archivos","*.*")])
        if not path: return
        wiz = ImportWizard(self, path); self.wait_window(wiz)
        if wiz.result_df is None: self.status.set("Importación cancelada."); return
        df = wiz.result_df
        for col in REQUIRED_COLS:
            if col not in df.columns: df[col] = ""
        df = df[REQUIRED_COLS]
        self.df_orig = df.copy(); self.df_view = df.copy()
        self._refresh_columns_combobox(); self._populate_table(self.df_view)
        self.status.set(f"Importado {os.path.basename(path)}")

    def _refresh_columns_combobox(self):
        if self.df_view is not None:
            self.cmb_column["values"] = list(self.df_view.columns)
            if len(self.df_view.columns) > 0: self.cmb_column.current(0)

    def _populate_table(self, df):
        for c in self.tree["columns"]: self.tree.heading(c, text="")
        self.tree.delete(*self.tree.get_children())
        cols = list(df.columns); self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c); self.tree.column(c, width=150 if c in REQUIRED_COLS else 120, stretch=True, anchor="w")
        for i, row in df.iterrows():
            self.tree.insert("", "end", iid=str(i), values=[row[c] for c in cols])

    def _sync_view_from_tree(self):
        if self.df_view is None: return
        cols = list(self.df_view.columns); new_rows = []
        for iid in self.tree.get_children(): new_rows.append(self.tree.item(iid, "values"))
        import pandas as pd
        self.df_view = pd.DataFrame(new_rows, columns=cols)

    def _on_double_click_cell(self, event): self.edit_selected_cell()
    def edit_selected_cell(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo(APP_NAME, "Seleccione una fila para editar."); return
        iid = sel[0]; col_ids = self.tree["columns"]
        col = simpledialog.askstring(APP_NAME, "¿Qué columna desea editar? (nombre exacto)")
        if not col or col not in col_ids: return
        idx = list(col_ids).index(col); cur = self.tree.item(iid, "values")[idx]
        new = simpledialog.askstring(APP_NAME, f"Nuevo valor para '{col}':", initialvalue=cur)
        if new is None: return
        vals = list(self.tree.item(iid, "values")); vals[idx] = new; self.tree.item(iid, values=vals); self._sync_view_from_tree()

    def apply_filter(self):
        if self.df_orig is None: return
        column = self.cmb_column.get().strip(); needle = self.ent_value.get().strip()
        df = self.df_orig.copy()
        if column and needle:
            if column not in df.columns: messagebox.showwarning(APP_NAME, "La columna indicada no existe en los datos."); return
            df = df[df[column].astype(str).str.contains(needle, case=False, na=False)]
        self.df_view = df; self._populate_table(df); self.status.set(f"Filtro aplicado: {column} contiene '{needle}'")

    def clear_filter(self):
        if self.df_orig is None: return
        self.ent_value.delete(0, tk.END); self.df_view = self.df_orig.copy(); self._populate_table(self.df_view); self.status.set("Filtros limpiados.")

    def _set_timestamp_for_selected(self, target):
        if self.df_view is None: return
        sels = self.tree.selection()
        if not sels: messagebox.showinfo(APP_NAME, "Seleccione al menos una fila."); return
        now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try: idx = list(self.tree["columns"]).index(target)
        except ValueError: messagebox.showwarning(APP_NAME, f"No existe la columna '{target}'."); return
        for iid in sels:
            vals = list(self.tree.item(iid, "values")); vals[idx] = now; self.tree.item(iid, values=vals)
        self._sync_view_from_tree(); self.status.set(f"{target} sellada para {len(sels)} fila(s).")
    def mark_llegada_real(self): self._set_timestamp_for_selected("LLEGADA REAL")
    def mark_salida_real(self): self._set_timestamp_for_selected("SALIDA REAL")

    def export_daily_report(self):
        if self.df_view is None or self.df_view.empty: messagebox.showwarning(APP_NAME, "No hay datos para reportar."); return
        import pandas as pd
        def parse_dt(s):
            if pd.isna(s) or str(s).strip() == "": return pd.NaT
            for fmt in ("%Y-%m-%d %H:%M:%S","%d/%m/%Y %H:%M","%d/%m/%Y %H:%M:%S","%H:%M","%H:%M:%S"):
                try:
                    if fmt in ("%H:%M","%H:%M:%S"):
                        t = pd.to_datetime(str(s), format=fmt, errors="coerce")
                        if pd.isna(t): continue
                        today = pd.Timestamp.today().normalize()
                        return today + pd.to_timedelta(t.strftime("%H:%M:%S"))
                    return pd.to_datetime(str(s), format=fmt, errors="coerce")
                except Exception: continue
            return pd.to_datetime(str(s), errors="coerce", dayfirst=True)
        tmp = self.df_view.copy()
        tmp["LR_dt"] = tmp["LLEGADA REAL"].apply(parse_dt)
        tmp["SR_dt"] = tmp["SALIDA REAL"].apply(parse_dt)
        tmp["ST_dt"] = tmp["SALIDA TOPE"].apply(parse_dt)
        total = len(tmp); con_lr = tmp["LR_dt"].notna().sum(); con_sr = tmp["SR_dt"].notna().sum()
        retrasos = ((tmp["SR_dt"].notna()) & (tmp["ST_dt"].notna()) & (tmp["SR_dt"] > tmp["ST_dt"])).sum()
        estancias = (tmp["SR_dt"] - tmp["LR_dt"]).dropna(); media = str(estancias.mean()) if not estancias.empty else ""
        kpis = {"Total filas": total, "% con LLEGADA REAL": (con_lr/total*100) if total else 0.0, "% con SALIDA REAL": (con_sr/total*100) if total else 0.0, "Retrasos vs SALIDA TOPE (nº)": int(retrasos), "Tiempo medio de estancia (hh:mm:ss)": media}
        ts = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S"); out_path = os.path.join(REPORT_DIR, f"{ts}_report.xlsx")
        try:
            with pd.ExcelWriter(out_path, engine="openpyxl") as w:
                pd.DataFrame([kpis]).T.rename(columns={0:"Valor"}).to_excel(w, index=True, header=True, sheet_name="KPIs")
                self.df_view.to_excel(w, index=False, sheet_name="Datos del día")
                tmp.groupby("TRANSPORTISTA", dropna=False).agg(total=("TRANSPORTISTA","size"),
                    con_llegada_real=("LR_dt", lambda s: s.notna().sum()),
                    con_salida_real=("SR_dt", lambda s: s.notna().sum())).reset_index().to_excel(w, index=False, sheet_name="Por transportista")
                tmp.groupby("MUELLE", dropna=False).agg(total=("MUELLE","size"),
                    con_llegada_real=("LR_dt", lambda s: s.notna().sum()),
                    con_salida_real=("SR_dt", lambda s: s.notna().sum())).reset_index().to_excel(w, index=False, sheet_name="Por muelle")
                tmp.loc[tmp["INCIDENCIAS"].astype(str).str.strip() != "", ["TRANSPORTISTA","MATRICULA","MUELLE","ESTADO","DESTINO","INCIDENCIAS"]].to_excel(w, index=False, sheet_name="Incidencias")
            messagebox.showinfo(APP_NAME, f"Reporte creado:\n{out_path}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"No se pudo crear el reporte.\n\n{e}")

    def save_session(self):
        if self.df_view is None: messagebox.showwarning(APP_NAME, "No hay datos para guardar."); return
        try: self.df_view.to_csv(SESSION, index=False, encoding="utf-8"); messagebox.showinfo(APP_NAME, f"Sesión guardada en {SESSION}")
        except Exception as e: messagebox.showerror(APP_NAME, f"No se pudo guardar la sesión.\n\n{e}")

    def close_day(self):
        if self.df_view is None or self.df_view.empty: messagebox.showwarning(APP_NAME, "No hay datos a guardar en el histórico."); return
        today = dt.datetime.now().strftime("%Y-%m-%d"); ts = dt.datetime.now().strftime("%H%M%S"); os.makedirs(HISTORY_DIR, exist_ok=True); fname = os.path.join(HISTORY_DIR, f"{today}_{ts}.csv")
        try: self.df_view.to_csv(fname, index=False, encoding="utf-8"); messagebox.showinfo(APP_NAME, f"Histórico guardado: {fname}")
        except Exception as e: messagebox.showerror(APP_NAME, f"No se pudo guardar el histórico.\n\n{e}")

    def _refresh_columns_combobox(self):
        if self.df_view is not None:
            self.cmb_column["values"] = list(self.df_view.columns)
            if len(self.df_view.columns) > 0: self.cmb_column.current(0)

    def _on_double_click_cell(self, event): pass  # simplificado

    def on_exit(self): self.destroy()

if __name__ == "__main__":
    app = LogiDeskApp(); app.mainloop()
