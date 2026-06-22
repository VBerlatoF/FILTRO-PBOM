from io import BytesIO

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
