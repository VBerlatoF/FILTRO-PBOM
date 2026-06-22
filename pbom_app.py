from io import BytesIO
from tkinter import BOTH, END, EXTENDED, LEFT, RIGHT, SINGLE, X, Y, filedialog, messagebox, ttk
import tkinter as tk

import pandas as pd


DEFAULT_FILE = (
    "//cbrsor01clfs050/public$/COMPOSI\u00c7\u00c3O/0 - BOM COMBINES/"
    "SAP HANA/WK25/PBOM COMPLETO 16.06.26.xlsb"
)


def excel_column_to_index(column_name: str) -> int:
    value = 0
    for char in column_name.upper():
        value = value * 26 + ord(char) - ord("A") + 1
    return value - 1


def normalize_marker(value) -> bool:
    return str(value).strip().upper() == "X"


def type_key(column) -> str:
    if pd.isna(column):
        return ""
    if isinstance(column, float) and column.is_integer():
        return str(int(column))
    label = str(column).strip()
    if label.endswith(".0") and label[:-2].isdigit():
        return label[:-2]
    return label


def build_type_groups(type_columns):
    groups = {}
    for column in type_columns:
        key = type_key(column)
        groups.setdefault(key, []).append(column)
    return groups


def build_result(df: pd.DataFrame, component_column: str, type_columns, type_groups, selected_type_keys):
    selected_type_keys = list(selected_type_keys)
    selected_columns = [
        column
        for key in selected_type_keys
        for column in type_groups.get(key, [])
    ]
    unselected_columns = [column for column in type_columns if column not in selected_columns]

    markers = df[type_columns].apply(lambda column: column.map(normalize_marker))
    component_key = "__component_key"
    markers[component_key] = df[component_column].astype(str).str.strip()
    markers = markers[markers[component_key] != ""]
    markers = markers.groupby(component_key, as_index=False)[type_columns].any()

    has_selected_type = markers[selected_columns].any(axis=1)
    has_unselected_type = markers[unselected_columns].any(axis=1) if unselected_columns else False
    exclusive_mask = has_selected_type & ~has_unselected_type

    result_markers = markers.loc[exclusive_mask, [component_key] + selected_columns].copy()
    summary = pd.DataFrame({component_key: result_markers[component_key]})
    summary["Tipos tecnicos selecionados"] = result_markers[selected_columns].apply(
        lambda row: ", ".join(
            key for key in selected_type_keys if any(row[column] for column in type_groups.get(key, []))
        ),
        axis=1,
    )
    summary["Colunas tecnicas com X"] = result_markers[selected_columns].apply(
        lambda row: ", ".join(str(column) for column in selected_columns if row[column]),
        axis=1,
    )

    bom_rows = df.copy()
    bom_rows[component_key] = bom_rows[component_column].astype(str).str.strip()
    bom_rows = bom_rows[bom_rows[component_key].isin(summary[component_key])]
    bom_rows = bom_rows.drop_duplicates(subset=[component_key], keep="first")
    bom_rows[component_column] = bom_rows[component_key]

    result = summary.merge(bom_rows, on=component_key, how="left").drop(columns=[component_key])
    leading_columns = ["Tipos tecnicos selecionados", "Colunas tecnicas com X"]
    bom_columns = [column for column in bom_rows.columns if column != component_key]
    result = result[leading_columns + bom_columns]
    result = result.sort_values(component_column)
    return result.reset_index(drop=True)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Componentes")
    return output.getvalue()


class PbomFilterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Filtro de PNs Exclusivos")
        self.geometry("1120x720")
        self.minsize(900, 560)

        self.file_path = tk.StringVar(value=DEFAULT_FILE)
        self.sheet_name = tk.StringVar()
        self.status = tk.StringVar(value="Selecione ou carregue o arquivo PBOM.")

        self.excel_file = None
        self.df = None
        self.component_column = None
        self.type_columns = []
        self.type_groups = {}
        self.type_keys = []
        self.result = pd.DataFrame()

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=14)
        root.pack(fill=BOTH, expand=True)

        file_frame = ttk.LabelFrame(root, text="Arquivo")
        file_frame.pack(fill=X)

        file_entry = ttk.Entry(file_frame, textvariable=self.file_path)
        file_entry.pack(side=LEFT, fill=X, expand=True, padx=(10, 8), pady=10)

        ttk.Button(file_frame, text="Procurar", command=self.choose_file).pack(side=LEFT, padx=4)
        ttk.Button(file_frame, text="Carregar arquivo", command=self.load_file).pack(side=LEFT, padx=(4, 10))

        sheet_frame = ttk.LabelFrame(root, text="Aba")
        sheet_frame.pack(fill=X, pady=(12, 0))

        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.sheet_name, state="readonly")
        self.sheet_combo.pack(side=LEFT, fill=X, expand=True, padx=(10, 8), pady=10)
        ttk.Button(sheet_frame, text="Carregar aba", command=self.load_sheet).pack(side=LEFT, padx=(4, 10))

        body = ttk.Frame(root)
        body.pack(fill=BOTH, expand=True, pady=(12, 0))

        selector_frame = ttk.LabelFrame(body, text="Tipos tecnicos finais AA ate BG")
        selector_frame.pack(side=LEFT, fill=Y)

        selector_tools = ttk.Frame(selector_frame)
        selector_tools.pack(fill=X, padx=10, pady=(10, 4))
        ttk.Button(selector_tools, text="Selecionar todos", command=self.select_all_types).pack(side=LEFT)
        ttk.Button(selector_tools, text="Limpar", command=self.clear_types).pack(side=LEFT, padx=(8, 0))

        list_frame = ttk.Frame(selector_frame)
        list_frame.pack(fill=BOTH, expand=True, padx=10, pady=(4, 10))

        self.types_list = tk.Listbox(list_frame, selectmode=EXTENDED, width=26, exportselection=False)
        self.types_list.pack(side=LEFT, fill=BOTH, expand=True)
        types_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.types_list.yview)
        types_scroll.pack(side=RIGHT, fill=Y)
        self.types_list.configure(yscrollcommand=types_scroll.set)

        action_frame = ttk.Frame(selector_frame)
        action_frame.pack(fill=X, padx=10, pady=(0, 10))
        ttk.Button(action_frame, text="Filtrar componentes", command=self.filter_components).pack(fill=X)
        ttk.Button(action_frame, text="Exportar Excel", command=self.export_result).pack(fill=X, pady=(8, 0))

        result_frame = ttk.LabelFrame(body, text="Resultado")
        result_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(12, 0))

        self.summary_label = ttk.Label(result_frame, text="0 componentes")
        self.summary_label.pack(anchor="w", padx=10, pady=(10, 4))

        tree_frame = ttk.Frame(result_frame)
        tree_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

        self.tree = ttk.Treeview(tree_frame, show="headings", selectmode=SINGLE)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_y.pack(side=RIGHT, fill=Y)
        self.tree.configure(yscrollcommand=tree_scroll_y.set)

        status_label = ttk.Label(root, textvariable=self.status, anchor="w")
        status_label.pack(fill=X, pady=(10, 0))

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Selecione o arquivo PBOM",
            filetypes=[("Excel Binary Workbook", "*.xlsb"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.file_path.set(path)

    def load_file(self):
        path = self.file_path.get().strip()
        if not path:
            messagebox.showwarning("Arquivo", "Informe o caminho do arquivo PBOM.")
            return

        try:
            self.excel_file = pd.ExcelFile(path, engine="pyxlsb")
        except Exception as exc:
            messagebox.showerror("Erro ao abrir arquivo", str(exc))
            return

        self.sheet_combo["values"] = self.excel_file.sheet_names
        if self.excel_file.sheet_names:
            self.sheet_name.set(self.excel_file.sheet_names[0])
        self.status.set(f"Arquivo carregado. {len(self.excel_file.sheet_names)} aba(s) encontrada(s).")

    def load_sheet(self):
        if self.excel_file is None:
            self.load_file()
            if self.excel_file is None:
                return

        if not self.sheet_name.get():
            messagebox.showwarning("Aba", "Selecione uma aba da planilha.")
            return

        try:
            self.df = pd.read_excel(self.file_path.get().strip(), sheet_name=self.sheet_name.get(), engine="pyxlsb")
        except Exception as exc:
            messagebox.showerror("Erro ao carregar aba", str(exc))
            return

        candidates = [column for column in self.df.columns if str(column).strip().lower() == "componente"]
        if not candidates:
            messagebox.showerror("Coluna nao encontrada", "Nao encontrei a coluna chamada Componente.")
            return

        self.component_column = candidates[0]
        start_index = excel_column_to_index("AA")
        end_index = excel_column_to_index("BG")
        self.type_columns = list(self.df.columns[start_index : end_index + 1])
        self.type_groups = build_type_groups(self.type_columns)
        self.type_keys = list(self.type_groups.keys())

        self.types_list.delete(0, END)
        for key in self.type_keys:
            count = len(self.type_groups[key])
            self.types_list.insert(END, f"{key} ({count} coluna{'s' if count != 1 else ''})")

        self.clear_result()
        self.status.set(
            f"Aba carregada. Coluna {self.component_column} encontrada e "
            f"{len(self.type_columns)} colunas / {len(self.type_keys)} finais tecnicos detectados."
        )

    def select_all_types(self):
        self.types_list.select_set(0, END)

    def clear_types(self):
        self.types_list.select_clear(0, END)

    def selected_types(self):
        indexes = self.types_list.curselection()
        return [self.type_keys[index] for index in indexes]

    def filter_components(self):
        if self.df is None:
            messagebox.showwarning("Planilha", "Carregue uma aba antes de filtrar.")
            return

        selected_types = self.selected_types()
        if not selected_types:
            messagebox.showwarning("Tipos tecnicos", "Selecione pelo menos um tipo tecnico.")
            return

        try:
            self.result = build_result(
                self.df,
                self.component_column,
                self.type_columns,
                self.type_groups,
                selected_types,
            )
        except Exception as exc:
            messagebox.showerror("Erro ao filtrar", str(exc))
            return

        self.render_result(self.result)
        self.status.set("Filtro aplicado com exclusividade e componentes repetidos removidos.")

    def clear_result(self):
        self.result = pd.DataFrame()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = []
        self.summary_label.configure(text="0 componentes")

    def render_result(self, result: pd.DataFrame):
        for item in self.tree.get_children():
            self.tree.delete(item)

        columns = list(result.columns)
        self.tree["columns"] = columns

        for column in columns:
            self.tree.heading(column, text=str(column))
            width = 170 if column == self.component_column else 110
            self.tree.column(column, width=width, minwidth=80, anchor="w")

        for _, row in result.iterrows():
            values = [row[column] for column in columns]
            self.tree.insert("", END, values=values)

        self.summary_label.configure(text=f"{len(result)} componentes exclusivos")

    def export_result(self):
        if self.result.empty:
            messagebox.showwarning("Exportar", "Filtre os componentes antes de exportar.")
            return

        path = filedialog.asksaveasfilename(
            title="Salvar resultado",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="componentes_pbom_filtrados.xlsx",
        )
        if not path:
            return

        try:
            with open(path, "wb") as file:
                file.write(to_excel_bytes(self.result))
        except Exception as exc:
            messagebox.showerror("Erro ao exportar", str(exc))
            return

        self.status.set(f"Resultado exportado: {path}")


if __name__ == "__main__":
    app = PbomFilterApp()
    app.mainloop()
