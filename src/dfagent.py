from src.memo import TableMemo, Modification, TableSnapShot, ModificationType, build_item_info
from src.observer import Observer
import pandas as pd


class DataFrameAgent(Observer):
    def __init__(self):
        self.df = None

    def load(self, excel_file, header=0):
        # On lit le DataFrame
        df = pd.read_excel(excel_file, header=header)

        # --- RÉPARATION DES EN-TÊTES FUSIONNÉS POUR L'IA ---
        # Si une colonne s'appelle "Unnamed: ...", on récupère le nom de la colonne à sa gauche
        new_columns = []
        last_valid_name = None
        for col in df.columns:
            col_str = str(col)
            if "Unnamed" in col_str:
                if last_valid_name is not None:
                    new_columns.append(last_valid_name)
                else:
                    new_columns.append(col_str)
            else:
                new_columns.append(col_str)
                last_valid_name = col_str

        df.columns = new_columns
        self.df = df

    def _update(self, subject):
        modifications = subject.modification
        mtype = modifications.mtype
        if mtype == ModificationType.NEW_TABLE:
            self.df = modifications.df

        # Récupération de l'offset pour les mises à jour individuelles
        self.header_offset = getattr(subject, 'header_row_idx', 1)

        if mtype != ModificationType.RESET:
            return

    def insert_row(self, item_infos):
        if len(item_infos) == 0:
            return
        new_row_data = {}
        for col_name, item_info in zip(self.df.columns, item_infos):
            new_row_data[col_name] = item_info.dtype(item_info.text)

        row_index = item_infos[0].index[0]
        # Ajuster pour le préambule (index absolu QTable -> index DataFrame)
        df_row_index = row_index - self.header_offset

        if df_row_index >= len(self.df):
            self.df.loc[len(self.df)] = new_row_data
        elif df_row_index >= 0:
            self.df = pd.concat([self.df.iloc[:df_row_index], pd.DataFrame([new_row_data]), self.df.iloc[df_row_index:]]).reset_index(
                drop=True)

    def insert_column(self, item_infos):
        new_column_data = []
        for item_info in item_infos:
            new_column_data.append(item_info.dtype(item_info.text))
        new_column_name = item_infos[0].column_name
        insert_position = item_infos[0].index[1]
        self.df.insert(insert_position, new_column_name, new_column_data)

    def delete_row(self, item_infos):
        if len(item_infos):
            item_info = item_infos[0]
            row, column = item_info.index
            # Ajuster pour le préambule
            df_row = row - self.header_offset
            if df_row >= 0 and df_row < len(self.df):
                self.df = self.df.drop(self.df.index[df_row]).reset_index(drop=True)

    def delete_column(self, item_infos):
        if len(item_infos) == 0:
            return
        item_info = item_infos[0]
        column_name_to_delete = item_info.column_name
        del self.df[column_name_to_delete]

    def update_inplace(self, item_infos):
        for item_info in item_infos:
            row, column = item_info.index
            value = item_info.text
            dtype = item_info.dtype
            # Ajuster l'index pour Pandas (qui ne contient pas le préambule)
            # row dans le tableau est 0-based, header_offset est 1-based (Excel)
            df_row = row - self.header_offset
            if df_row >= 0:
                self.df.iloc[df_row, column] = dtype(value)

    @property
    def shape(self):
        if self.df is not None:
            return self.df.shape

    @property
    def dtypes(self):
        if self.df is not None:
            return self.df.dtypes

    def head(self, n=3):
        if self.df is not None:
            return self.df.head(n)
