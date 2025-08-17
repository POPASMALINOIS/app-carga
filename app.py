#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LogiDesk Win v1.1 - Portable Windows app (offline, plug & play)
"""
import os, sys, glob, datetime as dt, tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd

APP_NAME = "LogiDesk Win"
REQUIRED_COLS = ["TRANSPORTISTA","MATRICULA","MUELLE","ESTADO","DESTINO","LLEGADA","LLEGADA REAL","SALIDA REAL","SALIDA TOPE","OBSERVACIONES","INCIDENCIAS"]
ALIASES = {"TRANSPORTISTA":["TRANSPORTISTA"],"MATRICULA":["MATRICULA","MATRÍCULA","MAT."],"MUELLE":["MUELLE","DOCK","RAMPA"],"ESTADO":["ESTADO","STATUS"],"DESTINO":["DESTINO","DEST."],"LLEGADA":["LLEGADA","HORA LLEGADA","ETA"],"LLEGADA REAL":["LLEGADA REAL","LLEGADA_REAL","CHECK-IN","ENTRADA REAL"],"SALIDA REAL":["SALIDA REAL","SALIDA_REAL","CHECK-OUT"],"SALIDA TOPE":["SALIDA TOPE","SALIDA_TOPE","CUT-OFF","HORA TOPE"],"OBSERVACIONES":["OBSERVACIONES","OBS.","NOTAS"],"INCIDENCIAS":["INCIDENCIAS","INC.","EVENTOS"]}
def canonical_name(col:str)->str:
    up=str(col).strip().upper()
    if up in REQUIRED_COLS: return up
    for t,als in ALIASES.items():
        for a in als:
            if up==a.upper(): return t
    return up
def app_data_dir()->str:
    base=os.path.abspath(".")
    try:
        os.makedirs(base,exist_ok=True)
        with open(os.path.join(base,".write_test"),"w",encoding="utf-8") as f: f.write("ok")
        os.remove(os.path.join(base,".write_test"))
        return base
    except Exception:
        fb=os.path.join(os.path.expanduser("~"),"LogiDeskWin"); os.makedirs(fb,exist_ok=True); return fb
BASE=app_data_dir(); HISTORY_DIR=os.path.join(BASE,"history"); REPORT_DIR=os.path.join(BASE,"reports"); SESSION=os.path.join(BASE,"current_session.csv")

class LogiDeskApp(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(f"{APP_NAME} v1.1"); self.geometry("1280x800"); self.minsize(1100,660)
        try:
            st=ttk.Style(self); 
            if "vista" in st.theme_names(): st.theme_use("vista")
        except Exception: pass
        self.df_orig=None; self.df_view=None; self.current_path=None
        os.makedirs(HISTORY_DIR,exist_ok=True); os.makedirs(REPORT_DIR,exist_ok=True)
        self._build_menu(); self._build_toolbar(); self._build_filterbar(); self._build_table()
        self._load_session_if_exists()
        self.protocol("WM_DELETE_WINDOW", self.on_exit)

    def _build_menu(self):
        m=tk.Menu(self)
        filem=tk.Menu(m,tearoff=0)
        filem.add_command(label="Cargar Excel/CSV\tCtrl+O",command=self.load_file,accelerator="Ctrl+O")
        filem.add_separator()
        filem.add_command(label="Guardar sesión\tCtrl+S",command=self.save_session,accelerator="Ctrl+S")
        filem.add_command(label="Cerrar día (a histórico)\tCtrl+D",command=self.close_day,accelerator="Ctrl+D")
        filem.add_command(label="Exportar reporte del día (Excel)\tCtrl+R",command=self.export_daily_report,accelerator="Ctrl+R")
        filem.add_separator()
        filem.add_command(label="Abrir carpeta histórico",command=lambda:self._open_folder(HISTORY_DIR))
        filem.add_command(label="Abrir carpeta de reportes",command=lambda:self._open_folder(REPORT_DIR))
        filem.add_separator(); filem.add_command(label="Salir\tAlt+F4",command=self.on_exit)
        m.add_cascade(label="Archivo",menu=filem)
        actm=tk.Menu(m,tearoff=0)
        actm.add_command(label="➕ LLEGADA REAL\tCtrl+L",command=self.mark_llegada_real,accelerator="Ctrl+L")
        actm.add_command(label="➕ SALIDA REAL\tCtrl+E",command=self.mark_salida_real,accelerator="Ctrl+E")
        actm.add_separator(); actm.add_command(label="Nueva fila\tIns",command=self.add_row,accelerator="Ins")
        actm.add_command(label="Eliminar fila(s)\tDel",command=self.delete_rows,accelerator="Del")
        actm.add_separator(); actm.add_command(label="Editar celda seleccionada…",command=self.edit_selected_cell)
        m.add_cascade(label="Acciones",menu=actm)
        viewm=tk.Menu(m,tearoff=0); viewm.add_command(label="Aplicar filtro\tCtrl+F",command=self.apply_filter,accelerator="Ctrl+F")
        viewm.add_command(label="Limpiar filtros\tCtrl+Shift+F",command=self.clear_filter,accelerator="Ctrl+Shift+F")
        m.add_cascade(label="Ver",menu=viewm)
        helpm=tk.Menu(m,tearoff=0); helpm.add_command(label="Acerca de",command=self.show_about); m.add_cascade(label="Ayuda",menu=helpm)
        self.config(menu=m)
        self.bind_all("<Control-o>",lambda e:self.load_file()); self.bind_all("<Control-s>",lambda e:self.save_session())
        self.bind_all("<Control-d>",lambda e:self.close_day()); self.bind_all("<Control-r>",lambda e:self.export_daily_report())
        self.bind_all("<Control-l>",lambda e:self.mark_llegada_real()); self.bind_all("<Control-e>",lambda e:self.mark_salida_real())
        self.bind_all("<Insert>",lambda e:self.add_row()); self.bind_all("<Delete>",lambda e:self.delete_rows())
        self.bind_all("<Control-f>",lambda e:self.apply_filter()); self.bind_all("<Control-Shift-F>",lambda e:self.clear_filter())

    def _build_toolbar(self):
        tb=ttk.Frame(self,padding=(10,8)); tb.pack(fill="x")
        for text,cmd in [("Cargar Excel/CSV",self.load_file),("Guardar sesión",self.save_session),("Cerrar día",self.close_day),("Exportar Excel",self.export_daily_report)]: ttk.Button(tb,text=text,command=cmd).pack(side="left",padx=4)
        ttk.Separator(tb,orient="vertical").pack(side="left",fill="y",padx=12)
        for text,cmd in [("➕ LLEGADA REAL",self.mark_llegada_real),("➕ SALIDA REAL",self.mark_salida_real)]: ttk.Button(tb,text=text,command=cmd).pack(side="left",padx=4)
        ttk.Separator(tb,orient="vertical").pack(side="left",fill="y",padx=12)
        ttk.Button(tb,text="Nueva fila",command=self.add_row).pack(side="left",padx=4); ttk.Button(tb,text="Eliminar fila(s)",command=self.delete_rows).pack(side="left",padx=4)

    def _build_filterbar(self):
        filt=ttk.LabelFrame(self,text="Filtros"); filt.pack(fill="x",padx=10,pady=(0,6))
        ttk.Label(filt,text="Columna:").pack(side="left",padx=(8,4)); self.cmb_column=ttk.Combobox(filt,state="readonly",values=[]); self.cmb_column.pack(side="left",padx=4,ipadx=30)
        ttk.Label(filt,text="Valor contiene:").pack(side="left",padx=(12,4)); self.ent_value=ttk.Entry(filt,width=30); self.ent_value.pack(side="left",padx=4)
        ttk.Button(filt,text="Aplicar",command=self.apply_filter).pack(side="left",padx=6); ttk.Button(filt,text="Limpiar",command=self.clear_filter).pack(side="left",padx=6)

    def _build_table(self):
        table=ttk.Frame(self); table.pack(fill="both",expand=True, padx=10,pady=(0,10))
        self.tree=ttk.Treeview(table,show="headings",selectmode="extended"); self.tree.pack(side="left",fill="both",expand=True)
        vsb=ttk.Scrollbar(table,orient="vertical",command=self.tree.yview); vsb.pack(side="right",fill="y")
        hsb=ttk.Scrollbar(self,orient="horizontal",command=self.tree.xview); hsb.pack(fill="x",padx=10,pady=(0,10))
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set); self.tree.bind("<Double-1>",self._on_double_click_cell)
        df=pd.DataFrame(columns=REQUIRED_COLS); self.df_orig=df.copy(); self.df_view=df.copy(); self._refresh_columns_combobox(); self._populate_table(self.df_view)

    def load_file(self):
        path=filedialog.askopenfilename(title="Selecciona un archivo",filetypes=[("Excel","*.xlsx *.xls"),("CSV","*.csv"),("Todos","*.*")]); 
        if not path: return
        try: self._load_path(path); self.current_path=path; messagebox.showinfo(APP_NAME,f"Archivo cargado: {os.path.basename(path)}")
        except Exception as e: messagebox.showerror(APP_NAME,f"No se pudo cargar el archivo.\n\n{e}")

    def _load_path(self,path:str):
        ext=os.path.splitext(path)[1].lower()
        if ext==".csv": df=pd.read_csv(path,sep=None,engine="python")
        elif ext in {".xlsx",".xls"}:
            try: df=pd.read_excel(path)
            except Exception:
                if ext==".xlsx": df=pd.read_excel(path,engine="openpyxl")
                else: df=pd.read_excel(path,engine="xlrd")
        else: raise ValueError("Extensión no soportada. Usa .csv, .xlsx o .xls")
        df.columns=[canonical_name(c) for c in df.columns]
        for c in REQUIRED_COLS:
            if c not in df.columns: df[c]=""
        other=[c for c in df.columns if c not in REQUIRED_COLS]
        df=df[REQUIRED_COLS+other]
        self.df_orig=df.copy(); self.df_view=df.copy(); self._refresh_columns_combobox(); self._populate_table(self.df_view)

    def _refresh_columns_combobox(self):
        if self.df_view is not None:
            self.cmb_column["values"]=list(self.df_view.columns)
            if len(self.df_view.columns)>0: self.cmb_column.current(0)

    def _populate_table(self,df:pd.DataFrame):
        for c in self.tree["columns"]: self.tree.heading(c,text="")
        self.tree.delete(*self.tree.get_children())
        cols=list(df.columns); self.tree["columns"]=cols
        for c in cols:
            self.tree.heading(c,text=c); self.tree.column(c,width=150 if c in REQUIRED_COLS else 120,stretch=True,anchor="w")
        for i,row in df.iterrows(): self.tree.insert("", "end", iid=str(i), values=[row[c] for c in cols])

    def _sync_view_from_tree(self):
        if self.df_view is None: return
        cols=list(self.df_view.columns); new_rows=[]
        for iid in self.tree.get_children(): new_rows.append(self.tree.item(iid,"values"))
        self.df_view=pd.DataFrame(new_rows,columns=cols)

    def _on_double_click_cell(self,event): self.edit_selected_cell()

    def edit_selected_cell(self):
        sel=self.tree.selection()
        if not sel: messagebox.showinfo(APP_NAME,"Seleccione una fila para editar."); return
        iid=sel[0]; col_ids=self.tree["columns"]
        col=simpledialog.askstring(APP_NAME,"¿Qué columna desea editar? (nombre exacto)")
        if not col or col not in col_ids: return
        idx=list(col_ids).index(col); cur=self.tree.item(iid,"values")[idx]
        new=simpledialog.askstring(APP_NAME,f"Nuevo valor para '{col}':",initialvalue=cur)
        if new is None: return
        vals=list(self.tree.item(iid,"values")); vals[idx]=new; self.tree.item(iid,values=vals); self._sync_view_from_tree()

    def add_row(self):
        cols=list(self.df_view.columns); new=["" for _ in cols]; new_iid=str(len(self.tree.get_children())); self.tree.insert("", "end", iid=new_iid, values=new); self._sync_view_from_tree()

    def delete_rows(self):
        sels=self.tree.selection()
        if not sels: messagebox.showinfo(APP_NAME,"Seleccione al menos una fila a eliminar."); return
        for iid in sels: self.tree.delete(iid); self._sync_view_from_tree()

    def apply_filter(self):
        if self.df_orig is None: return
        column=self.cmb_column.get().strip(); needle=self.ent_value.get().strip(); df=self.df_orig.copy()
        if column and needle:
            if column not in df.columns: messagebox.showwarning(APP_NAME,"La columna indicada no existe."); return
            df=df[df[column].astype(str).str.contains(needle,case=False,na=False)]
        self.df_view=df; self._populate_table(df)

    def clear_filter(self):
        if self.df_orig is None: return
        self.ent_value.delete(0,tk.END); self.df_view=self.df_orig.copy(); self._populate_table(self.df_view)

    def _set_ts(self,target:str):
        if self.df_view is None: return
        sels=self.tree.selection()
        if not sels: messagebox.showinfo(APP_NAME,"Seleccione al menos una fila."); return
        now=dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try: idx=list(self.tree["columns"]).index(target)
        except ValueError: messagebox.showwarning(APP_NAME,f"No existe la columna '{target}'."); return
        for iid in sels:
            vals=list(self.tree.item(iid,"values")); vals[idx]=now; self.tree.item(iid,values=vals)
        self._sync_view_from_tree()

    def mark_llegada_real(self): self._set_ts("LLEGADA REAL")
    def mark_salida_real(self): self._set_ts("SALIDA REAL")

    def _parse_dt(self,s):
        import pandas as pd
        if pd.isna(s) or str(s).strip()=="": return pd.NaT
        for fmt in ("%Y-%m-%d %H:%M:%S","%d/%m/%Y %H:%M","%d/%m/%Y %H:%M:%S","%H:%M","%H:%M:%S"):
            try:
                if fmt in ("%H:%M","%H:%M:%S"):
                    t=pd.to_datetime(str(s),format=fmt,errors="coerce")
                    if pd.isna(t): continue
                    today=pd.Timestamp.today().normalize()
                    return today+pd.to_timedelta(t.strftime("%H:%M:%S"))
                return pd.to_datetime(str(s),format=fmt,errors="coerce")
            except Exception: continue
        return pd.to_datetime(str(s),errors="coerce",dayfirst=True)

    def _compute_kpis(self,df):
        import pandas as pd
        tmp=df.copy()
        tmp["LR_dt"]=tmp["LLEGADA REAL"].apply(self._parse_dt)
        tmp["SR_dt"]=tmp["SALIDA REAL"].apply(self._parse_dt)
        tmp["ST_dt"]=tmp["SALIDA TOPE"].apply(self._parse_dt)
        total=len(tmp); con_lr=tmp["LR_dt"].notna().sum(); con_sr=tmp["SR_dt"].notna().sum()
        retrasos=((tmp["SR_dt"].notna())&(tmp["ST_dt"].notna())&(tmp["SR_dt"]>tmp["ST_dt"])).sum()
        estancias=(tmp["SR_dt"]-tmp["LR_dt"]).dropna(); media=str(estancias.mean()) if not estancias.empty else ""
        return {"Total filas":total,"% con LLEGADA REAL":(con_lr/total*100) if total else 0.0,"% con SALIDA REAL":(con_sr/total*100) if total else 0.0,"Retrasos vs SALIDA TOPE (nº)":int(retrasos),"Tiempo medio de estancia (hh:mm:ss)":media}, tmp

    def export_daily_report(self):
        import pandas as pd, os
        if self.df_view is None or self.df_view.empty: messagebox.showwarning(APP_NAME,"No hay datos para reportar."); return
        kpis,tmp=self._compute_kpis(self.df_view)
        by_tr=tmp.groupby("TRANSPORTISTA",dropna=False).agg(total=("TRANSPORTISTA","size"),con_llegada_real=("LR_dt",lambda s:s.notna().sum()),con_salida_real=("SR_dt",lambda s:s.notna().sum())).reset_index()
        by_mu=tmp.groupby("MUELLE",dropna=False).agg(total=("MUELLE","size"),con_llegada_real=("LR_dt",lambda s:s.notna().sum()),con_salida_real=("SR_dt",lambda s:s.notna().sum())).reset_index()
        incid=tmp.loc[tmp["INCIDENCIAS"].astype(str).str.strip()!="",["TRANSPORTISTA","MATRICULA","MUELLE","ESTADO","DESTINO","INCIDENCIAS"]]
        ts=dt.datetime.now().strftime("%Y-%m-%d_%H%M%S"); out=os.path.join(REPORT_DIR,f"{ts}_report.xlsx")
        try:
            with pd.ExcelWriter(out,engine="openpyxl") as w:
                pd.DataFrame([kpis]).T.rename(columns={0:"Valor"}).to_excel(w,index=True,header=True,sheet_name="KPIs")
                self.df_view.to_excel(w,index=False,sheet_name="Datos del día")
                by_tr.to_excel(w,index=False,sheet_name="Por transportista")
                by_mu.to_excel(w,index=False,sheet_name="Por muelle")
                incid.to_excel(w,index=False,sheet_name="Incidencias")
            messagebox.showinfo(APP_NAME,f"Reporte creado:\n{out}")
        except Exception as e:
            messagebox.showerror(APP_NAME,f"No se pudo crear el reporte.\n\n{e}")

    def save_session(self):
        try:
            if self.df_view is None: messagebox.showwarning(APP_NAME,"No hay datos para guardar."); return
            self.df_view.to_csv(SESSION,index=False,encoding="utf-8"); messagebox.showinfo(APP_NAME,f"Sesión guardada en {SESSION}")
        except Exception as e: messagebox.showerror(APP_NAME,f"No se pudo guardar la sesión.\n\n{e}")

    def _load_session_if_exists(self):
        import pandas as pd, os
        if os.path.exists(SESSION):
            try:
                df=pd.read_csv(SESSION,encoding="utf-8"); df.columns=[canonical_name(c) for c in df.columns]
                for c in REQUIRED_COLS:
                    if c not in df.columns: df[c]=""
                other=[c for c in df.columns if c not in REQUIRED_COLS]; df=df[REQUIRED_COLS+other]
                self.df_orig=df.copy(); self.df_view=df.copy(); self._refresh_columns_combobox(); self._populate_table(self.df_view)
            except Exception as e: messagebox.showwarning(APP_NAME,f"No se pudo cargar la sesión previa: {e}")

    def close_day(self):
        import os
        if self.df_view is None or self.df_view.empty: messagebox.showwarning(APP_NAME,"No hay datos a guardar en el histórico."); return
        today=dt.datetime.now().strftime("%Y-%m-%d"); ts=dt.datetime.now().strftime("%H%M%S"); os.makedirs(HISTORY_DIR,exist_ok=True); fname=os.path.join(HISTORY_DIR,f"{today}_{ts}.csv")
        try:
            self.df_view.to_csv(fname,index=False,encoding="utf-8"); self._enforce_history_retention(days=30); messagebox.showinfo(APP_NAME,f"Histórico guardado: {fname}")
        except Exception as e: messagebox.showerror(APP_NAME,f"No se pudo guardar el histórico.\n\n{e}")

    def _enforce_history_retention(self,days:int=30):
        import os, pandas as pd
        files=sorted(glob.glob(os.path.join(HISTORY_DIR,"*.csv"))); 
        if not files: return
        cutoff=dt.datetime.now()-dt.timedelta(days=days)
        for f in files:
            try:
                mtime=dt.datetime.fromtimestamp(os.path.getmtime(f))
                if mtime<cutoff: os.remove(f)
            except Exception: pass
        by_date={}
        for f in sorted(glob.glob(os.path.join(HISTORY_DIR,"*.csv"))):
            base=os.path.basename(f); date_part=base.split("_")[0]; by_date.setdefault(date_part,[]).append(f)
        dates=sorted(by_date.keys())
        if len(dates)>30:
            overflow=dates[:len(dates)-30]
            for d in overflow:
                for f in by_date.get(d,[]): 
                    try: os.remove(f)
                    except Exception: pass

    def _open_folder(self,path:str):
        try:
            if sys.platform.startswith("win"): os.startfile(path)  # type: ignore
            elif sys.platform=="darwin": os.system(f'open "{path}"')
            else: os.system(f'xdg-open "{path}"')
        except Exception: messagebox.showinfo(APP_NAME,f"Carpeta: {path}")

    def show_about(self): messagebox.showinfo(APP_NAME,f"{APP_NAME} v1.1\n\nPortable, offline, sin Python en el PC destino (usando .exe compilado).")

    def on_exit(self):
        try:
            if self.df_view is not None: self.df_view.to_csv(SESSION,index=False,encoding="utf-8")
        except Exception: pass
        self.destroy()

if __name__=="__main__":
    app=LogiDeskApp(); app.mainloop()
