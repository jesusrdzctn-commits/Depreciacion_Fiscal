"""
GUI flow (Tkinter):
1) Selecciona el PRIMER archivo (nombre variable) ➜ selecciona la hoja.
2) Selecciona SOLO las columnas para construir la LLAVE_1 (pre‑marcadas: "Soc.", "Activo fijo", "SNº").
3) Se construye la LLAVE (concatenación de las columnas elegidas del archivo 1).
4) Selecciona el SEGUNDO archivo ➜ selecciona la hoja.
5) Selecciona SOLO las columnas para construir la LLAVE_2 (pre‑marcadas: "Sociedad", "Activo fijo", "Subnúmero").
6) Ventana para elegir QUÉ CAMPOS exportar de cada archivo:
   - Pre‑seleccionar TODOS los campos del segundo archivo.
   - En el primero, pre‑seleccionar únicamente "   Ingr.por baja" (si existe; respeta espacios iniciales).
7) Merge por LLAVE (inner por defecto; configurable "left/right/outer").
8) Post‑merge, pre‑exportación:
   8.1) Pedir por ventana el valor numérico de INPC diferido (mensaje: "Colocar último INPC conocido").
   8.2) Recalcular Factor de Actualización2 con la lógica:
        ratio = INPC_diferido / INPC_de_compra.
        Si ratio > 1 entonces:
          • Si INPC_diferido > 300 → 1
          • Si no → truncar a 4 decimales: Fix(ratio*10000)/10000
        En caso contrario → 1
   8.3) Depreciación Fiscal2 = (Depreciación del Ejercicio Actual) * (Factor de Actualización2)
   8.4) Saldo por redimir actualizado2 = (Valor en libros al cierre) * (Factor de Actualización2)
   8.5) Crear Utilidad = max(0, Ingr.por baja − Saldo actualizado) y Pérdida = max(0, Saldo actualizado − Ingr.por baja)
   8.6) Colocar "   Ingr.por baja" antes de Utilidad en el orden final.
9) Exportar a Excel.

Notas:
- Lee todo como texto (dtype=object) para no perder ceros a la izquierda.
- Soporta .xlsx/.xls/.xlsm (pandas + openpyxl).
- Si hay nombres duplicados entre archivos (aparte de LLAVE), las del segundo se renombran con sufijo _2.
- Cambia el tipo de merge en MERGE_HOW si lo necesitas ("inner", "left", "right", "outer").
"""
from __future__ import annotations

import math
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

# ===================== Configuración =====================
SUGGESTED_KEYS_FILE1 = ["Soc.", "Activo fijo", "SNº"]
SUGGESTED_KEYS_FILE2 = ["Sociedad", "Activo fijo", "Subnúmero"]
PREFILL_ONLY_FROM_FILE1 = ["   Ingr.por baja"]  # respeta espacios como vienen

REQUIRED_FILE2_COLS = [
    "INPC de compra",
    "Depreciación del Ejercicio Actual",
    "Valor en libros al cierre",
]

MERGE_HOW = "inner"  # opciones: "inner", "left", "right", "outer"

# ===================== Utilidades de archivo/hoja =====================

def pick_file(title: str) -> Path | None:
    root = tk.Tk()
    root.withdraw()
    filetypes = [("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
    path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return Path(path) if path else None


def pick_sheet(xls: pd.ExcelFile, title: str) -> str | None:
    sheets = xls.sheet_names
    if not sheets:
        messagebox.showerror("Sin hojas", "El archivo no contiene hojas de cálculo.")
        return None

    win = tk.Tk()
    win.title(title)
    win.geometry("420x160")
    win.resizable(False, False)

    ttk.Label(win, text=title, font=("Segoe UI", 10, "bold")).pack(pady=(12, 4))

    choice = tk.StringVar(value=sheets[0])
    combo = ttk.Combobox(win, textvariable=choice, values=sheets, state="readonly")
    combo.pack(pady=6, padx=12, fill="x")

    result: dict[str, str | None] = {"sheet": None}

    def confirm() -> None:
        result["sheet"] = choice.get()
        win.destroy()

    ttk.Button(win, text="Continuar", command=confirm).pack(pady=14)
    win.attributes("-topmost", True)
    win.mainloop()

    return result["sheet"]

# ===================== Selectores de columnas =====================

def pick_columns_for_key(
    columns: List[str],
    suggestions: List[str],
    title: str,
    note: str,
) -> List[str] | None:
    """Selector multi-columna (solo para construir LLAVE). Preselecciona 'suggestions' si existen."""
    win = tk.Tk()
    win.title(title)
    win.geometry("580x460")
    win.resizable(True, True)

    info = (
        f"{note}"
        f"Recomendadas (preseleccionadas si existen): {', '.join(suggestions)}"
    )
    ttk.Label(win, text=info, wraplength=540, justify="left").pack(padx=12, pady=(12, 8), anchor="w")

    frame = ttk.Frame(win)
    frame.pack(fill="both", expand=True, padx=12, pady=8)

    scrollbar = ttk.Scrollbar(frame, orient="vertical")
    lst = tk.Listbox(frame, selectmode="extended", exportselection=False)
    lst.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=lst.yview)

    lst.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    for col in columns:
        lst.insert("end", col)

    sugg = {s for s in suggestions if s in columns}
    for i, col in enumerate(columns):
        if col in sugg:
            lst.selection_set(i)

    result: dict[str, List[str] | None] = {"cols": None}

    def confirm() -> None:
        idx = lst.curselection()
        if not idx:
            messagebox.showwarning("Sin selección", "Selecciona al menos una columna para la LLAVE.")
            return
        result["cols"] = [columns[i] for i in idx]
        win.destroy()

    ttk.Button(win, text="Usar para LLAVE", command=confirm).pack(pady=10)
    win.attributes("-topmost", True)
    win.mainloop()

    return result["cols"]


def pick_export_columns(columns1: List[str], columns2: List[str]) -> Tuple[List[str], List[str]] | None:
    """Ventana dual para elegir qué columnas exportar de cada archivo.
    Preselecciona TODO en archivo 2 y solo las indicadas en PREFILL_ONLY_FROM_FILE1 en archivo 1 (si existen).
    """
    win = tk.Tk()
    win.title("Selecciona columnas a exportar de cada archivo")
    win.geometry("900x520")

    lbl = ttk.Label(
        win,
        text=(
            "Elige las columnas que aparecerán en el archivo final."
            "LLAVE irá siempre primero."
        ),
        wraplength=860,
    )
    lbl.pack(padx=12, pady=(10, 4), anchor="w")

    container = ttk.Frame(win)
    container.pack(fill="both", expand=True, padx=12, pady=8)

    # Panel archivo 1
    left = ttk.Frame(container)
    left.pack(side="left", fill="both", expand=True, padx=(0, 6))

    ttk.Label(left, text="Archivo 1 – columnas a exportar", font=("Segoe UI", 10, "bold")).pack(anchor="w")

    sb1 = ttk.Scrollbar(left, orient="vertical")
    lst1 = tk.Listbox(left, selectmode="extended", exportselection=False)
    lst1.config(yscrollcommand=sb1.set)
    sb1.config(command=lst1.yview)

    lst1.pack(side="left", fill="both", expand=True, pady=(6, 0))
    sb1.pack(side="left", fill="y")

    for c in columns1:
        lst1.insert("end", c)

    pre1 = {p for p in PREFILL_ONLY_FROM_FILE1 if p in columns1}
    for i, c in enumerate(columns1):
        if c in pre1:
            lst1.selection_set(i)

    # Panel archivo 2
    right = ttk.Frame(container)
    right.pack(side="left", fill="both", expand=True, padx=(6, 0))

    ttk.Label(right, text="Archivo 2 – columnas a exportar (preseleccionadas todas)", font=("Segoe UI", 10, "bold")).pack(anchor="w")

    sb2 = ttk.Scrollbar(right, orient="vertical")
    lst2 = tk.Listbox(right, selectmode="extended", exportselection=False)
    lst2.config(yscrollcommand=sb2.set)
    sb2.config(command=lst2.yview)

    lst2.pack(side="left", fill="both", expand=True, pady=(6, 0))
    sb2.pack(side="left", fill="y")

    for c in columns2:
        lst2.insert("end", c)

    if columns2:
        lst2.selection_set(0, "end")

    result: dict[str, List[str] | None] = {"c1": None, "c2": None}

    def confirm() -> None:
        idx1 = lst1.curselection()
        idx2 = lst2.curselection()
        result["c1"] = [columns1[i] for i in idx1]
        result["c2"] = [columns2[i] for i in idx2]
        win.destroy()

    ttk.Button(win, text="Continuar", command=confirm).pack(pady=10)
    win.attributes("-topmost", True)
    win.mainloop()

    if result["c1"] is None and result["c2"] is None:
        return None

    return result["c1"], result["c2"]

# ===================== LLAVE, helpers numéricos y exportación =====================

def build_llave_from_cols(df: pd.DataFrame, key_cols: List[str]) -> pd.Series:
    parts: List[pd.Series] = []
    for c in key_cols:
        if c in df.columns:
            parts.append(df[c].astype(str).fillna("").str.strip())
        else:
            parts.append(pd.Series([""] * len(df), index=df.index))

    if not parts:
        return pd.Series([""] * len(df), index=df.index)

    out = parts[0]
    for p in parts[1:]:
        out = out + p  # cambia a out + '|' + p si quieres separador
    return out


def coerce_numeric(series: pd.Series) -> pd.Series:
    """Convierte a número de forma tolerante: limpia NBSP/comas/espacios y coacciona a float (NaN si no se puede)."""
    s = series.astype(str).str.replace(" ", " ")  # NBSP → espacio normal
    s = s.str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce")


def prompt_inpc_diferido() -> float | None:
    while True:
        val = simpledialog.askstring("INPC diferido", "Colocar último INPC conocido:")
        if val is None:
            return None
        try:
            num = float(str(val).replace(",", "."))
            return num
        except ValueError:
            messagebox.showerror("Valor inválido", "Ingresa un número válido para INPC diferido.")


def export_excel(df: pd.DataFrame, default_name: str, initial_dir: Path) -> None:
    save_path = filedialog.asksaveasfilename(
        title="Guardar Excel",
        defaultextension=".xlsx",
        initialdir=str(initial_dir),
        initialfile=default_name,
        filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx")],
    )
    if not save_path:
        messagebox.showinfo("Cancelado", "No se guardó ningún archivo.")
        return

    try:
        df.to_excel(save_path, index=False)
    except Exception as e:
        messagebox.showerror(
            "Error al guardar",
            f"No se pudo guardar el archivo: {e}"
        )
        return

    messagebox.showinfo(
        "Listo",
        f"Archivo exportado con éxito: {save_path}"
    )

# ===================== Flujo principal =====================

def main() -> None:
    # --- Archivo 1 ---
    path1 = pick_file("Selecciona el PRIMER archivo de Excel")
    if not path1:
        return

    try:
        xls1 = pd.ExcelFile(path1)
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"No se pudo abrir el Excel 1:{e}"
        )
        return

    sheet1 = pick_sheet(xls1, "Selecciona la hoja del PRIMER archivo")
    if not sheet1:
        return

    try:
        df1 = pd.read_excel(xls1, sheet_name=sheet1, dtype=object)
        df1.columns = [str(c) for c in df1.columns]
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"No se pudo leer la hoja '{sheet1}' del archivo 1: {e}"
        )
        return

    key_cols1 = pick_columns_for_key(
        list(df1.columns),
        SUGGESTED_KEYS_FILE1,
        "Archivo 1: columnas para LLAVE",
        "Selecciona las columnas del Archivo 1 que formarán la LLAVE."
    )
    if not key_cols1:
        return

    df1 = df1.copy()
    df1["LLAVE"] = build_llave_from_cols(df1, key_cols1)

    # --- Archivo 2 ---
    path2 = pick_file("Selecciona el SEGUNDO archivo de Excel")
    if not path2:
        return

    try:
        xls2 = pd.ExcelFile(path2)
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"No se pudo abrir el Excel 2: {e}"
        )
        return

    sheet2 = pick_sheet(xls2, "Selecciona la hoja del SEGUNDO archivo")
    if not sheet2:
        return

    try:
        df2 = pd.read_excel(xls2, sheet_name=sheet2, dtype=object)
        df2.columns = [str(c) for c in df2.columns]
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"No se pudo leer la hoja '{sheet2}' del archivo 2: {e}"
        )
        return

    key_cols2 = pick_columns_for_key(
        list(df2.columns),
        SUGGESTED_KEYS_FILE2,
        "Archivo 2: columnas para LLAVE",
        "Selecciona las columnas del Archivo 2 que formarán la LLAVE."
    )
    if not key_cols2:
        return

    df2 = df2.copy()
    df2["LLAVE"] = build_llave_from_cols(df2, key_cols2)

    # --- Selección de columnas a exportar de cada archivo ---
    cols1_to_export, cols2_to_export = pick_export_columns(list(df1.columns), list(df2.columns))
    if cols1_to_export is None or cols2_to_export is None:
        return

    # Evitar duplicar LLAVE
    cols1_to_export = [c for c in cols1_to_export if c != "LLAVE"]
    cols2_to_export = [c for c in cols2_to_export if c != "LLAVE"]

    # Forzar columnas requeridas del archivo 2
    for req in REQUIRED_FILE2_COLS:
        if req not in cols2_to_export and req in df2.columns:
            cols2_to_export.append(req)

    # --- Merge por LLAVE ---
    overlap = set(cols1_to_export).intersection(set(cols2_to_export))
    df1_export = df1[["LLAVE"] + cols1_to_export].copy()
    df2_export = df2[["LLAVE"] + cols2_to_export].copy()

    rename2 = {c: f"{c}_2" for c in overlap}
    if rename2:
        df2_export = df2_export.rename(columns=rename2)

    merged = pd.merge(df1_export, df2_export, on="LLAVE", how=MERGE_HOW)

    # --- Post-merge: cálculos ---
    def col2(name: str) -> str:
        if name in merged.columns:
            return name
        alt = name + "_2"
        return alt if alt in merged.columns else name

    inpc_diferido = prompt_inpc_diferido()
    if inpc_diferido is None:
        messagebox.showinfo("Cancelado", "Proceso cancelado antes de exportar.")
        return

    merged["INPC diferido"] = inpc_diferido

    inpc_compra_col = col2("INPC de compra")
    inpc_compra = coerce_numeric(merged.get(inpc_compra_col, pd.Series([float("nan")] * len(merged))))
    inpc_dif = float(inpc_diferido)

    def calc_factor(row_inpc_compra: float) -> float:
        try:
            denom = float(row_inpc_compra)
        except (TypeError, ValueError):
            denom = float("nan")
        if not (denom and denom > 0):
            return 1.0
        ratio = inpc_dif / denom
        if ratio > 1:
            if inpc_dif > 300:
                return 1.0
            return math.trunc(ratio * 10000.0) / 10000.0
        return 1.0

    merged["Factor de Actualización2"] = inpc_compra.apply(calc_factor)

    dep_ej_col = col2("Depreciación del Ejercicio Actual")
    dep_ej = coerce_numeric(merged.get(dep_ej_col, pd.Series([0.0] * len(merged))))
    merged["Depreciacion Fiscal2"] = dep_ej * merged["Factor de Actualización2"]

    val_lib_col = col2("Valor en libros al cierre")
    val_lib = coerce_numeric(merged.get(val_lib_col, pd.Series([0.0] * len(merged))))
    merged["Saldo por redimir actualizado2"] = val_lib * merged["Factor de Actualización2"]

    ingr_col = "   Ingr.por baja"  # del archivo 1
    if ingr_col not in merged.columns:
        merged[ingr_col] = pd.NA

    ingr = coerce_numeric(merged[ingr_col])
    saldo_act = coerce_numeric(merged["Saldo por redimir actualizado2"])

    merged["Utilidad"] = (ingr - saldo_act).clip(lower=0)
    merged["Pérdida"] = (saldo_act - ingr).clip(lower=0)

    # --- Orden final de columnas ---
    if ingr_col not in cols1_to_export:
        cols1_to_export = cols1_to_export + [ingr_col]

    cols2_real = [col2(c) for c in cols2_to_export]

    ordered_cols: List[str] = ["LLAVE"]

    cols1_no_ingr = [c for c in cols1_to_export if c != ingr_col and c in merged.columns]
    ordered_cols += cols1_no_ingr

    if ingr_col in merged.columns:
        ordered_cols.append(ingr_col)

    calc_cols = [
        "INPC diferido",
        "Factor de Actualización2",
        "Depreciacion Fiscal2",
        "Saldo por redimir actualizado2",
        "Utilidad",
        "Pérdida",
    ]

    for c in cols2_real:
        if c not in ["LLAVE", ingr_col] and c not in calc_cols and c in merged.columns:
            ordered_cols.append(c)

    for c in calc_cols:
        if c in merged.columns and c not in ordered_cols:
            ordered_cols.append(c)

    seen: set[str] = set()
    final_cols: List[str] = []
    for c in ordered_cols:
        if c in merged.columns and c not in seen:
            seen.add(c)
            final_cols.append(c)

    merged_out = merged[final_cols].copy()

    # --- Export ---
    default_name = f"{Path(path1).stem}__{Path(path2).stem}_merge.xlsx"
    export_excel(merged_out, default_name, initial_dir=path1.parent)


if __name__ == "__main__":
    main()
