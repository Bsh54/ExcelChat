from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, \
    QMenu, QAction, QShortcut, \
    QColorDialog, QFontDialog, QFileDialog
from PyQt5.QtCore import Qt, QPoint
from src.memo import TableMemo, Modification, TableSnapShot, ModificationType, build_item_info
from src.utis import withoutconnect, hex_to_rgb
import pandas as pd
from src.logger import logger
from PyQt5.QtGui import QKeySequence, QIcon, QFont, QColor
from src.excelio import ExcelAgent
from src.dfagent import DataFrameAgent
import tempfile
import shutil


class ResultTable(QTableWidget):
    def __init__(self, parent=None):
        super(ResultTable, self).__init__(parent)
        self.setAlternatingRowColors(True)
        self.setStyleSheet("QTableWidget { alternate-background-color: #f2f2f2; background-color: white; }")
        self.setEditTriggers(QTableWidget.NoEditTriggers) # Lecture seule pour les résultats

    def _update(self, subject):
        modifications = subject.modification
        mtype = modifications.mtype
        if mtype not in [ModificationType.NEW_TABLE]:
            return
        self.clear() # clear() au lieu de clearContents pour réinitialiser aussi les headers
        df = modifications.df
        if df is None:
            return

        self.setRowCount(len(df))
        self.setColumnCount(len(df.columns))
        self.setHorizontalHeaderLabels(list(df.columns))

        for i in range(len(df)):
            for j in range(len(df.columns)):
                val = df.iloc[i, j]
                display_val = "" if pd.isna(val) else str(val)
                item = QTableWidgetItem(display_val)

                # Alignement au centre pour les chiffres
                if isinstance(val, (int, float)) and not pd.isna(val):
                    item.setTextAlignment(Qt.AlignCenter)

                self.setItem(i, j, item)

        self.resizeColumnsToContents()


class EnhancedTable(QTableWidget):
    _observers = []

    def __init__(self, fig_dir=None):
        super(EnhancedTable, self).__init__()
        self.itemChanged.connect(self.handle_item_changed)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        self.setSelectionMode(QTableWidget.MultiSelection)
        self.selectionModel().selectionChanged.connect(self.selection_changed)
        self.cellClicked.connect(self.cell_clicked)
        self.selected_indexes = set()
        self.dataframe = None
        self.loaded = False
        self.init_table()
        self.snapshot = TableSnapShot()
        self.excel_agent = ExcelAgent()
        self.df_agent = DataFrameAgent()
        self.attach(self.excel_agent)
        self.attach(self.df_agent)
        self.recoder = TableMemo(self)
        self._modification = None
        self.fig_dir = fig_dir
        self.header_row_idx = 1 # Par défaut

    @property
    def modification(self):
        return self._modification

    @modification.setter
    def modification(self, data):
        self._modification = data

    def show_context_menu(self, pos):
        item = self.itemAt(pos)
        if item:
            row, column = item.row(), item.column()
            if len(self.selected_indexes) <= 1:
                self.clearSelection()
                self.selected_indexes = {(row, column)}
                self.setCurrentCell(row, column)
            main_menu = QMenu(self)
            style_menu = QMenu("Paramètres du style de cellule", self)

            font_action = QAction('Police', self)
            font_action.triggered.connect(self.config_font_style)
            style_menu.addAction(font_action)

            font_color_action = QAction('Couleur du texte', self)
            font_color_action.triggered.connect(self.config_font_color)
            style_menu.addAction(font_color_action)
            main_menu.addMenu(style_menu)

            bg_color_action = QAction('Couleur de fond', self)
            bg_color_action.triggered.connect(self.config_bg_color)
            style_menu.addAction(bg_color_action)

            edit_menu = QMenu("Édition Structure", self)

            insert_row_action = QAction('Insérer une ligne au-dessus', self)
            insert_row_action.triggered.connect(self.insert_new_row)
            edit_menu.addAction(insert_row_action)

            insert_col_action = QAction('Insérer une colonne à gauche', self)
            insert_col_action.triggered.connect(self.insert_new_column)
            edit_menu.addAction(insert_col_action)

            edit_menu.addSeparator()

            delete_row_action = QAction('Supprimer la ligne', self)
            delete_row_action.triggered.connect(self.remove_selected_rows)
            edit_menu.addAction(delete_row_action)

            delete_column_action = QAction('Supprimer la colonne', self)
            delete_column_action.triggered.connect(self.remove_selected_columns)
            edit_menu.addAction(delete_column_action)
            main_menu.addMenu(edit_menu)

            main_menu.exec_(self.mapToGlobal(QPoint(pos.x() + 100, pos.y())))

    @withoutconnect
    def insert_new_row(self):
        current_row = self.currentRow()
        if current_row < 0: current_row = self.rowCount()

        # 1. UI
        self.insertRow(current_row)
        for j in range(self.columnCount()):
            self.setItem(current_row, j, QTableWidgetItem(""))
            # Copie style du dessus
            if current_row > 0:
                prev = self.item(current_row - 1, j)
                if prev: self.item(current_row, j).setBackground(prev.background())

        # 2. DataFrame
        df_idx = current_row - self.header_row_idx
        if self.dataframe is not None:
            new_row = pd.Series([None] * len(self.dataframe.columns), index=self.dataframe.columns)
            if df_idx <= 0:
                self.dataframe = pd.concat([pd.DataFrame([new_row]), self.dataframe]).reset_index(drop=True)
            elif df_idx >= len(self.dataframe):
                self.dataframe.loc[len(self.dataframe)] = None
            else:
                self.dataframe = pd.concat([self.dataframe.iloc[:df_idx], pd.DataFrame([new_row]), self.dataframe.iloc[df_idx:]]).reset_index(drop=True)

        # 3. Notify
        self.notify_structure_change()

    @withoutconnect
    def insert_new_column(self):
        current_col = self.currentColumn()
        if current_col < 0: current_col = self.columnCount()

        # 1. UI
        self.insertColumn(current_col)
        new_name = f"NouvCol{self.columnCount()}"

        # 2. DataFrame
        if self.dataframe is not None:
            self.dataframe.insert(max(0, current_col), new_name, None)

        # 3. Notify
        self.notify_structure_change()

    def notify_structure_change(self):
        self.modification = Modification(ModificationType.NEW_TABLE, [])
        self.modification.df = self.dataframe
        self.notify()
        self.save_checkpoint()

    @withoutconnect
    def remove_selected_rows(self):
        rows = sorted(set(index.row() for index in self.selectedIndexes()), reverse=True)
        if not rows or self.dataframe is None:
            return

        for row in rows:
            # 1. Mise à jour de l'UI
            self.removeRow(row)

            # 2. Mise à jour du DataFrame (en tenant compte de l'offset)
            df_row = row - self.header_row_idx
            if 0 <= df_row < len(self.dataframe):
                self.dataframe = self.dataframe.drop(self.dataframe.index[df_row]).reset_index(drop=True)

        # 3. Notification CRUCIAL : On informe l'agent Excel qu'il doit réécrire le tableau
        self.modification = Modification(ModificationType.NEW_TABLE, [])
        self.modification.df = self.dataframe
        self.notify()
        self.save_checkpoint()

    @withoutconnect
    def remove_selected_columns(self):
        cols = sorted(set(index.column() for index in self.selectedIndexes()), reverse=True)
        if not cols or self.dataframe is None:
            return

        for col in cols:
            col_name = self.get_column_name(col)
            # 1. Mise à jour de l'UI
            self.removeColumn(col)

            # 2. Mise à jour du DataFrame
            if col_name in self.dataframe.columns:
                self.dataframe = self.dataframe.drop(columns=[col_name])

        # 3. Notification
        self.modification = Modification(ModificationType.NEW_TABLE, [])
        self.modification.df = self.dataframe
        self.notify()
        self.save_checkpoint()

    def clear_selection(self):
        self.clearSelection()
        # self.selected_indexes = set()

    def selection_changed(self, selected, deselected):
        selected_indexes = selected.indexes()

        if selected_indexes:
            selected_cells = set((index.row(), index.column()) for index in selected_indexes)
            self.selected_indexes = self.selected_indexes | selected_cells

    def cell_clicked(self, row, column):
        table_widget = self.sender()

        table_widget.clearSelection()
        self.selected_indexes = {(row, column)}

        table_widget.setCurrentCell(row, column)

    @withoutconnect
    def config_font_style(self):
        font, ok = QFontDialog.getFont()
        if ok:
            item_infos = []
            for index in self.selected_indexes:
                row, column = index
                old_item_info = self.snapshot.get(index)
                item_infos.append(old_item_info)
                item = self.item(row, column)
                item.setFont(font)
            self.modification = Modification(ModificationType.UPDATE_INPLACE, item_infos)
            self.clearSelection()
            self.save_checkpoint()

    @withoutconnect
    def config_font_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            item_infos = []
            for index in self.selected_indexes:
                row, column = index
                old_item_info = self.snapshot.get(index)
                item_infos.append(old_item_info)
                item = self.item(row, column)
                item.setData(Qt.TextColorRole, color)
            self.modification = Modification(ModificationType.UPDATE_INPLACE, item_infos)
            self.save_checkpoint()
            self.clearSelection()

    @withoutconnect
    def config_bg_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            item_infos = []
            for index in self.selected_indexes:
                row, column = index
                old_item_info = self.snapshot.get(index)
                item_infos.append(old_item_info)
                item = self.item(row, column)
                item.setBackground(color)
            self.modification = Modification(ModificationType.UPDATE_INPLACE, item_infos)
            self.save_checkpoint()
            self.clearSelection()

    @withoutconnect
    def init_table(self):
        self.setRowCount(20)
        self.setColumnCount(20)
        for row in range(20):
            for col in range(20):
                self.setItem(row, col, QTableWidgetItem(''))

    @withoutconnect
    def open_excel(self):
        filename, filetype = QFileDialog.getOpenFileName(self, "Sélectionner un fichier Excel", "", "*.xlsx;;*.xls;;All Files(*)")
        if not filename:
            return
        ok = self.load_excel(filename)
        if not ok:
            return ok
        self.loaded = True
        return ok

    def load_excel(self, excel_file):
        ok = self.excel_agent.load(excel_file)
        if not ok:
            return ok

        # --- DÉTECTION INTELLIGENTE DE L'EN-TÊTE (pour calculs seulement) ---
        ws = self.excel_agent.ws
        header_row_idx = 1
        max_non_empty = 0
        for i in range(1, min(11, ws.max_row + 1)):
            non_empty_count = sum(1 for cell in ws[i] if cell.value is not None)
            if non_empty_count > max_non_empty:
                max_non_empty = non_empty_count
                header_row_idx = i

        self.header_row_idx = header_row_idx

        # ON AFFICHE TOUT (dès la ligne 1)
        self.setRowCount(self.excel_agent.num_rows)
        self.setColumnCount(self.excel_agent.num_cols)

        # Les en-têtes de colonnes Qt seront simplement A, B, C...
        # car les vrais en-têtes sont dans le tableau
        self.setHorizontalHeaderLabels([self.excel_agent.index_to_excel_index(0, j).split('1')[0] for j in range(self.excel_agent.num_cols)])

        # --- GESTION DES CELLULES FUSIONNÉES (UI) ---
        for merged_range in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            # setSpan(row, column, rowSpan, columnSpan)
            self.setSpan(min_row - 1, min_col - 1, max_row - min_row + 1, max_col - min_col + 1)

        # Chargement de TOUTES les lignes
        header_names = []
        for row in ws.iter_rows(min_row=1):
            for cell in row:
                cell_value = cell.value if cell.value is not None else ""
                i = cell.row - 1
                j = cell.column - 1
                item = QTableWidgetItem(str(cell_value))
                item.custom_dtype = type(cell.value) if cell.value is not None else str
                item = self.format_with_cell(cell, item)

                # Si c'est la ligne d'en-tête détectée, on la met visuellement en évidence
                if cell.row == header_row_idx:
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                    header_names.append(str(cell_value) if cell_value != "" else f"Col{cell.column}")

                self.setItem(i, j, item)
                # On ne snapshot pour l'IA que ce qui est au niveau ou en dessous de l'en-tête
                if cell.row >= header_row_idx:
                    col_name = header_names[j] if j < len(header_names) else f"Col{j+1}"
                    self.snapshot.set(item, (i, j), col_name)

        # Synchronisation de Pandas : on lui dit de sauter les lignes avant l'en-tête
        self.df_agent.load(excel_file, header=header_row_idx - 1)
        self.dataframe = pd.read_excel(excel_file, header=header_row_idx - 1)

        self.resizeColumnsToContents()
        return ok

    def format_with_cell(self, cell, item):
        cell_font = cell.font
        font_name = cell_font.name
        font_size = cell_font.size
        item_font = QFont(font_name, int(font_size))
        item.setFont(item_font)

        font_color = cell_font.color.rgb if cell_font.color else None
        if isinstance(font_color, str) and len(font_color) >= 6:
            try:
                # Gérer les formats ARGB (8 hex) ou RGB (6 hex)
                hex_val = font_color[-6:]
                color = QColor(*hex_to_rgb(hex_val))
                item.setData(Qt.TextColorRole, color)
            except Exception as e:
                logger.debug(f"Erreur couleur texte: {e}")

        # Réactivation et fiabilisation de la couleur de fond
        if cell.fill and hasattr(cell.fill, 'fgColor') and cell.fill.fgColor:
            bg_color = cell.fill.fgColor.rgb
            if isinstance(bg_color, str) and len(bg_color) >= 6:
                try:
                    hex_val = bg_color[-6:]
                    # Éviter le blanc pur ou les couleurs invalides par défaut d'Excel
                    if hex_val.upper() != "000000" or bg_color.startswith("FF"):
                         color = QColor(*hex_to_rgb(hex_val))
                         item.setBackground(color)
                except Exception as e:
                    logger.debug(f"Erreur couleur fond: {e}")
        return item

    def save(self):
        return self.modification

    def _update_inplace(self, item_infos):
        item_info = item_infos[0]
        text = item_info.text()
        index = item_info.index
        font = item_info.font
        font_color = item_info.font_color
        bg_color = item_info.bg_color
        row, column = index
        item = self.item(row, column)
        item.setText(text)
        item.setFont(font)
        item.setData(Qt.TextColorRole, font_color)
        item.setBackground(bg_color)

    def _insert_scalar(self, item_infos):
        item_info = item_infos[0]
        value = item_info.text
        row, col = item_info.index
        self.item(row, col).setText(value)

    def _delete_scalar(self, item_infos):
        item_info = item_infos[0]
        value = item_info.text
        row, col = item_info.index
        self.item(row, col).setText('')

    def restore(self, modifications):
        if modifications is None:
            return
        if len(modifications.item_infos) == 0:
            return
        mtype = modifications.mtype
        if mtype != ModificationType.UPDATE_INPLACE:
            return
        self._update_inplace(modifications.item_infos)
        self.modification = modifications
        self.notify()
        self.modification = None

    @withoutconnect
    def insert_result(self, res):
        stdout, stderror = res
        if stdout is None:
            return
        if isinstance(stdout, pd.DataFrame):
            # L'offset de l'en-tête (1-based Excel -> 0-based index)
            header_row_idx = self.header_row_idx - 1
            data_start_row = self.header_row_idx

            # 1. Mise à jour des noms de colonnes (UNIQUEMENT si l'IA a donné des noms explicites et non génériques)
            for j in range(min(stdout.shape[1], self.columnCount())):
                new_col_name = str(stdout.columns[j])
                # On ne touche pas à l'en-tête si c'est un nom générique Pandas ou si la cellule est fusionnée (item est None)
                if "Unnamed" not in new_col_name and not self.isColumnHidden(j):
                    item = self.item(header_row_idx, j)
                    if item:
                        item.setText(new_col_name)

            # 2. Préparation de la taille (on ne réduit jamais la taille existante)
            new_data_rows = stdout.shape[0]
            total_required_rows = data_start_row + new_data_rows
            if total_required_rows > self.rowCount():
                self.setRowCount(total_required_rows)

            # 3. Insertion chirurgicale des données (préserve styles et fusions)
            new_item_infos = []
            for i in range(stdout.shape[0]):
                for j in range(stdout.shape[1]):
                    value = stdout.iloc[i, j]
                    display_value = "" if pd.isna(value) else str(value)
                    target_row = i + data_start_row

                    item = self.item(target_row, j)
                    if not item:
                        item = QTableWidgetItem(display_value)
                        # Copie du style de la ligne d'en-tête pour les nouvelles cellules
                        sample = self.item(header_row_idx, j)
                        if sample:
                            item.setFont(sample.font())
                            item.setBackground(sample.background())
                        self.setItem(target_row, j, item)
                    else:
                        item.setText(display_value)

                    index = (target_row, j)
                    column_header_item = self.item(header_row_idx, j)
                    current_col_name = column_header_item.text() if column_header_item else f"Col{j+1}"

                    item.custom_dtype = type(value)
                    new_item_infos.append(build_item_info(item, index, current_col_name))

            # On ne fait plus de self.setRowCount(total_rows) pour ne pas supprimer la suite du fichier
            self.dataframe = stdout
            self.modification = Modification(ModificationType.NEW_TABLE, new_item_infos)
            self.modification.df = stdout

            self.resizeColumnsToContents()

            # Défilement automatique vers la zone mise à jour
            self.scrollToItem(self.item(data_start_row, 0))

            self.notify()
            return True
        else:
            # Pour les résultats simples (chiffres, textes), on laisse l'onglet IA ou le chat l'afficher
            return True

    @withoutconnect
    def undo_modification(self):
        self.recoder.undo()

    def save_checkpoint(self):
        self.recoder.backup()

    def get_column_name(self, column):
        # On va chercher le nom dans la ligne d'en-tête détectée (header_row_idx est 1-based)
        header_item = self.item(self.header_row_idx - 1, column)
        if header_item and header_item.text():
            return header_item.text()
        return f"Col{column + 1}"

    def handle_item_changed(self, item):
        if self.dataframe is None:
            return
        item_info = self.snapshot.get(item)
        row = item.row()
        column = item.column()
        index = (row, column)
        column_name = self.get_column_name(column)
        if item_info is None:
            mtype = ModificationType.INSERT_SCALAR
            item_info = build_item_info(item, index, column_name)
        else:
            mtype = ModificationType.UPDATE_INPLACE
        modification = Modification(mtype, item_info)
        self.modification = modification
        self.save_checkpoint()
        self.snapshot.set(item, index, column_name)

    def register_shortcut(self):
        self.shortcut = QShortcut(QKeySequence(Qt.CTRL + Qt.Key_Z), self)
        self.shortcut.activated.connect(self.undo_modification)

    def file_save(self):
        file_filter = "*.xlsx;;*.xls;;All Files(*)"
        path, _ = QFileDialog.getSaveFileName(self, "Enregistrer le fichier Excel", "", file_filter)

        if not path:
            return
        self.excel_agent.save(path, self.fig_dir)

    def attach(self, observer):
        self._observers.append(observer)

    def detach(self, observer):
        self._observers.remove(observer)

    def notify(self):
        for observer in self._observers:
            observer._update(self)
