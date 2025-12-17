import argparse
import os
import sys
import csv


def missatges(titul, msg, kind='info'):
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()

        # Forçar que la finestra sigui la superior i guanyi el focus
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass
        root.update()
        root.lift()
        root.focus_force()

        # Mostra el quadre amb parent perquè sigui modal sobre root
        if kind == 'error':
            messagebox.showerror(titul, msg, parent=root)
        elif kind == 'warning':
            messagebox.showwarning(titul, msg, parent=root)
        else:
            messagebox.showinfo(titul, msg, parent=root)

        # Opcional: treure topmost abans de destruir
        # try:
        #     root.attributes("-topmost", False)
        # except Exception:
        #     pass
        # root.destroy()
    except Exception:
        # fallback a consola si no hi ha GUI disponible
        print(f"{titul}: {msg}")


def crear_ruta_unica(path):
    base, ext = os.path.splitext(path)
    counter = 1
    candidate = path
    while os.path.exists(candidate):
        candidate = f"{base}_{counter}{ext}"
        counter += 1
    return candidate


def nom_fulla(fulla):
    return "".join(c if c.isalnum() or c in " .-_()" else "_" for c in fulla).strip() or "sheet"


def formatejar_columna_data(df, posicions=(7, 17, 18), format_data='%d/%m/%Y'):
    import pandas as pd
    for pos in posicions:
        if pos < df.shape[1]:
            col = df.columns[pos]
            parsed = pd.to_datetime(df[col], errors='coerce')
            df[col] = parsed.dt.strftime(format_data)
            df[col] = df[col].fillna('')
    return df


def fmt_num(x):
    import pandas as pd
    if pd.isnull(x):
        return ''
    s = "{:.15g}".format(x)
    return s.replace('.', ',')



def conversio_excel_a_csv(file_path, out_dir=".", overwrite=False, skiprows=10):
    import pandas as pd
    if not os.path.exists(file_path):
        missatges("Error", f"Arxiu `{file_path}` no trobat", kind='error')
        return

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    try:
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        missatges("Error lectura", f"  -> Error lectura arxiu `{file_path}`: {e}", kind='error')
        return

    for fulla in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=fulla, skiprows=skiprows+1, header=None)
        except Exception as e:
            missatges("Error lectura", f"  -> Error lectura pagina `{fulla}`: {e}", kind='error')
            continue

        if df.shape[1] > 0:
            df = df.iloc[:, 1:]

        df = df.replace({r'[\r\n]+': ' '}, regex=True)
        df = formatejar_columna_data(df, posicions=(7, 17, 18), format_data='%d/%m/%Y')
        obj_cols = df.select_dtypes(include=['object']).columns
        if len(obj_cols) > 0:
            df[obj_cols] = df[obj_cols].apply(lambda s: s.str.strip().str.rstrip(', '))

        num_cols = df.select_dtypes(include=['number']).columns
        if len(num_cols) > 0:

            df[num_cols] = df[num_cols].apply(lambda col: col.map(fmt_num))

        safe_sheet = nom_fulla(fulla)
        out_name = f"{base_name}_{safe_sheet}.csv"
        out_path = os.path.join(out_dir, out_name)

        if not overwrite:
            out_path = crear_ruta_unica(out_path)

        try:
            with open(out_path, "w", newline="", encoding="utf-8") as f:
                df.to_csv(
                    f,
                    sep=";",
                    header=False,
                    index=False,
                    quoting=csv.QUOTE_MINIMAL,
                    quotechar='"',
                    doublequote=True,
                    lineterminator="\n"
                )
            missatges("Conversió correcta", f"`{file_path}` :: `{fulla}` -> `{out_path}`", kind='info')
        except Exception as e:
            missatges("Error escriptura", f"  -> Error writing `{out_path}`: {e}", kind='error')

def escollir_arxiu():
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return []
    root = tk.Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(
        title="Seleccionar arxius",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    root.destroy()
    return list(files)

def parse_args(argv):
    p = argparse.ArgumentParser(
        description="Converteix fitxers Excel a CSV. Elimina les primeres 10 files i primera columna per defecte."
    )
    p.add_argument("files", nargs="*", help="Fitxer(s) Excel a convertir")
    p.add_argument("--out-dir", default=".", help="Directori de sortida (per defecte: directori actual)")
    p.add_argument("--overwrite", action="store_true", help="Permet sobreescriure fitxers existents")
    p.add_argument("--skiprows", type=int, default=10, help="Files a ometre a l'inici (per defecte: 10)")
    return p.parse_args(argv)

def main():
    args = parse_args(sys.argv[1:])
    files = list(args.files)

    if not files:
        files = escollir_arxiu()
        if not files:
            missatges("Info", "No s'han seleccionat fitxers. Tancant.", kind='info')
            return 0

    os.makedirs(args.out_dir, exist_ok=True)

    for f in files:
        conversio_excel_a_csv(
            f,
            out_dir=args.out_dir,
            overwrite=args.overwrite,
            skiprows=args.skiprows
        )

    return 0

if __name__ == "__main__":
    sys.exit(main())