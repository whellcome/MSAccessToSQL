import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox
from tkinter import ttk

import pandas as pd
import win32com.client


class WidgetsRender:
    def __init__(self, render_params=None, *args, **options):
        """
        Initialization of the Frame, description of the main elements
        :param render_params: General parameters for the arrangement of elements can be set externally
        :param args:
        :param options:
        """
        super().__init__(*args, **options)
        if render_params is None:
            render_params = dict(sticky="ew", padx=5, pady=2)
        self.__render_params = render_params

    def render(self, obj=None, render_params=None):
        """
        Perform element creation and rendering in one command. Without creating a variable unnecessarily.
        Combines general parameters for the arrangement of elements and parameters for a specific element.
        :param obj: Element to rendering
        :param render_params: Dictionary with element parameters
        :return: Rendered element
        """
        if obj:
            render_params = render_params if render_params else {}
            united_pack_params = self.__render_params.copy()
            united_pack_params.update(render_params)
            obj.grid(united_pack_params)
        return obj


class TreeviewDataFrame(WidgetsRender, ttk.Treeview):

    def __init__(self, parent, render_params=None, *args, **kwargs):
        super().__init__(render_params, parent, *args, **kwargs)
        self.df = pd.DataFrame()
        self.filtered_df = pd.DataFrame()
        self.bind("<Button-1>", self.toggle_cell)
        self.svars = {"flag_symbol": {
            "check": "✔",
            "uncheck": " "
        }}
        self.svars["flag_values"] = {
            self.svars["flag_symbol"]["uncheck"]: self.svars["flag_symbol"]["check"],
            self.svars["flag_symbol"]["check"]: self.svars["flag_symbol"]["uncheck"]
        }


    def column(self, column, option=None, **kw):
        """
            Override column method with DataFrame.
        """
        result = super().column(column, option=option, **kw)
        if column not in self.df.columns:
            self.df[column] = ''
        return result

    def insert(self, parent, index, iid=None, **kw):
        """
               Inserts a new row into the Treeview and synchronizes it with the DataFrame.

               :param parent: Parent node for Treeview (usually "" for root-level items).
               :param index: Position to insert the item.
               :param iid: Unique identifier for the row. If None, Treeview generates one.
               :param kw: Additional arguments for Treeview insert (e.g., values).
               """
        # Use the provided iid or let Treeview generate one
        if iid is None:
            iid = super().insert(parent, index, **kw)  # Automatically generate iid
        else:
            super().insert(parent, index, iid=iid, **kw)

        # Ensure values are provided
        values = kw.get("values", [])

        # Convert values to a DataFrame-compatible dictionary
        new_row = {col: val for col, val in zip(self.cget("columns"), values)}

        # Add the new row to the DataFrame, using iid as the index
        self.df.loc[iid] = new_row
        return iid

    def set(self, item, column=None, value=None):
        """
            Enhanced set method for synchronization with a DataFrame.

            :param item: The item ID (iid) in the Treeview.
            :param column: The column name to retrieve or update.
            :param value: The value to set; if None, retrieves the current value.
            :return: The value as returned by the original Treeview method.
        """
        result = super().set(item, column, value)
        if item not in self.df.index:
            raise KeyError(f"Row with index '{item}' not found in DataFrame.")
        is_filtered = True if item in self.filtered_df.index else False
        if value is None:
            if column is None:
                self.df.loc[item] = self.df.loc[item].replace(result)
            else:
                self.df.loc[item, column] = result
        else:
            self.df.loc[item, column] = value
            if is_filtered:
                self.filtered_df.loc[item, column] = value
        return result

    def item(self, item, option=None, **kw):
        """
        Override item method with DataFrame.
        """
        values = kw.get("values", [])
        result = super().item(item, option, **kw)
        is_filtered = True if item in self.filtered_df.index else False
        if option is None and len(values):
            updates = pd.Series(values, index=self.cget("columns"))
            self.df.loc[item] = updates
            if is_filtered:
                self.filtered_df.loc[item] = updates
        return result

    def delete(self, *items, inplace=False):
        """
        Override delete method with DataFrame..
        """
        if inplace:
            for item in items:
                values = self.item(item, "values")
                self.df = self.df[~(self.df[list(self.df.columns)] == values).all(axis=1)]
        super().delete(*items)

    def flag_inverse(self, value: str) -> str:
        flag_values = self.svars["flag_values"]
        return flag_values[value]

    def toggle_cell(self, event):
        """Handles cell clicks to change flags."""
        if self.identify_region(event.x, event.y) != "cell":
            return
        col_num = int(self.identify_column(event.x).replace("#", "")) - 1
        if not col_num:
            return
        col_name = self.cget("columns")[col_num]
        item = self.identify_row(event.y)
        current_value = self.set(item, col_name)
        self.set(item, col_name, self.flag_inverse(current_value))
        self.event_generate("<<TreeToggleCell>>")

    def rebuild_tree(self, dataframe=None):
        if dataframe is None:
            dataframe = self.df
        self.delete(*self.get_children())
        for index, row in dataframe.iterrows():
            self.insert("", "end", iid=index, values=row.to_list())

    def filter_by_name(self, keyword=""):
        """Filter rows based on a keyword and update Treeview."""
        self.filtered_df = self.df[self.df[self.df.columns[0]].str.contains(keyword, case=False)].copy()
        self.rebuild_tree(self.filtered_df)

    def filter_event_evoke(self):
        """Filter updated event."""
        self.event_generate("<<TreeFilterUpdated>>")

    def all_checked(self, column):
        df = self.filtered_df if len(self.filtered_df) else self.df
        return not len(df[df.iloc[:, column] == " "])

    def filter_widget(self, parent):
        widget_frame = ttk.Frame(parent, width=150, borderwidth=1, relief="solid", padding=(2, 2))
        self.render(tk.Label(widget_frame, text="Filter by table name:", font=("Helvetica", 9, "bold")),
                    dict(row=0, column=0, pady=5))
        filter_entry = tk.Entry(widget_frame)
        self.render(filter_entry, dict(row=0, column=1, padx=5, pady=5, sticky="ew"))

        def apply_filter():
            self.filter_by_name(filter_entry.get())
            if len(self.filtered_df) == len(self.df):
                self.filtered_df = pd.DataFrame()
            self.filter_event_evoke()

        def clear_filter():
            self.rebuild_tree()
            filter_entry.delete(0, tk.END)
            self.filtered_df = pd.DataFrame()
            self.filter_event_evoke()

        self.render(ttk.Button(widget_frame, text="Filter", command=apply_filter),
                    dict(row=0, column=2, padx=5, pady=5))
        self.render(ttk.Button(widget_frame, text="Restore", command=clear_filter),
                    dict(row=0, column=3, padx=5, pady=5))
        return widget_frame

    def checkbox_widget(self, parent):

        def toggle_all(ind):
            checked = self.svars['check_all'][ind].get()
            if not checked:
                self.svars['check_all'][ind].set(False)
            for item in self.get_children():
                values = list(self.item(item, "values"))
                if checked:
                    values[ind] = self.svars['flag_symbol']['check']
                else:
                    values[ind] = self.svars['flag_symbol']['uncheck']
                self.item(item, values=values)

        widget_frame = ttk.Frame(parent, padding=(2, 2))
        self.svars["check_all"] = {}
        for i, col in  enumerate(self.cget("columns")[1:]):
            ind = i+1
            self.svars["check_all"][ind] = tk.IntVar(value=0)
            box_text = f"Check all {self.heading(col)['text'] if self.heading(col)['text'] else col}"
            render_params = dict(row=0, column=ind, padx=20)
            self.render(ttk.Checkbutton(widget_frame, text=box_text, variable=self.svars["check_all"][ind],
                                        command=lambda k=ind: toggle_all(k)), render_params)
        return widget_frame


class GetWidgetsFrame(WidgetsRender, ttk.Frame):
    """
    The main class of the program is responsible for constructing the form and interaction of elements
    """

    def __init__(self, render_params=None, *args, **options):
        """
        Initialization of the Frame, description of the main elements
        :param render_params: General parameters for the arrangement of elements can be set externally
        :param args:
        :param options:
        """
        super().__init__(*args, **options)
        self.db = None
        self.svars = {
            'dao_types': {
                1: "Boolean",
                3: "Integer",
                4: "Long",
                5: "Currency",
                7: "Single",
                8: "Double",
                9: "Date",
                10: "Text",
                11: "Binary",
                12: "Text"
            }}
        self.db_path = tk.StringVar(self, "")
        self.label1 = ttk.Label(self, text="", font=("Helvetica", 12))
        self.frame0 = ttk.Frame(self, width=240, borderwidth=1, relief="solid", padding=(2, 2))
        self.frame1 = ttk.Frame(self, width=100, borderwidth=1, relief="solid", padding=(2, 2))
        self.tree = TreeviewDataFrame(self.frame1, columns=("table", "export", "data"), show="headings")
        self.tree.bind("<<TreeFilterUpdated>>", self.on_filter_updated)
        self.tree.bind("<<TreeToggleCell>>", self.on_toggle_cell)
        self.scrollbar = ttk.Scrollbar(self.frame1, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        self.create_widgets()

    def create_widgets(self):
        """Building the main widgets at the beginning of program execution"""
        self.render(self)
        self.render(tk.Label(self, text="MS Access to SQL Export Tool", font=("Helvetica", 14)),
                    dict(row=0, column=0, columnspan=3, pady=5))
        self.render(self.label1, dict(row=1, column=0, columnspan=3))
        self.render(tk.Button(self, text="MS Access File Open", command=self.btn_openf),
                    dict(row=2, column=0, columnspan=2))
        self.render(tk.Button(self, text=" Exit ", command=self.btn_exit), dict(row=2, column=2, columnspan=2))
        self.render(self.frame1, dict(row=4, column=0, columnspan=3))
        self.render(self.tree, dict(row=0, column=0, pady=5))
        self.render(self.scrollbar, dict(row=0, column=3, sticky="ns"))
        self.render(tk.Button(self, text=" Run! ", command=self.btn_run, font=("Helvetica", 12)),
                    dict(row=5, column=0, columnspan=3, ))

    def recreate_widgets(self):
        self.render(self.tree, dict(row=0, column=0, pady=5))
        self.render(self.scrollbar, dict(row=0, column=3, sticky="ns"))
        self.render(self.frame0, dict(row=3, column=0, columnspan=3, sticky="e"))
        self.render(self.tree.filter_widget(self.frame0),
                    dict(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="ew"))
        self.render(self.tree.checkbox_widget(self.frame0),
                    dict(row=4, column=0, columnspan=3, padx=5, pady=5, sticky="e"))

        # self.render(tk.Label(self.frame0, text=" ", width=43),
        #             dict(row=1, column=0, columnspan=2, pady=5))
        # self.svars['check_all'] = tk.IntVar(value=0)
        # self.render(ttk.Checkbutton(self.frame0, text="Check all to Export", variable=self.svars['check_all'],
        #                             command=self.toggle_all_export), dict(row=1, column=2, padx=20))
        # self.svars['check_all_upload'] = tk.IntVar(value=0)
        # self.render(ttk.Checkbutton(self.frame0, text="Check all to Upload", variable=self.svars['check_all_upload'],
        #                             command=self.toggle_all_upload, ), dict(row=1, column=3, padx=20))

    def make_tree(self):
        self.tree.heading("table", text="Table")
        self.tree.heading("export", text="Export Table")
        self.tree.heading("data", text="Upload Data")
        self.tree.column("table", width=150, anchor="w")
        self.tree.column("export", width=50, anchor="center")
        self.tree.column("data", width=50, anchor="center")
        for table in self.db.TableDefs:
            if not table.Name.startswith("MSys"):
                self.tree.insert("", "end", values=(table.Name, " ", " "))
        self.tree.grid(row=3, column=0, columnspan=3, pady=5)
        style = ttk.Style()
        style.map("Treeview",
                  background=[("disabled", "#c0c0c0"), ("selected", "#d9f2d9")],
                  foreground=[("selected", "#000000")]
                  )
        style.configure("Treeview", rowheight=25)
        self.tree.tag_configure("normal")
        self.tree.tag_configure("export", background="#fff0f0")

    def update_data_column(self):
        """..."""
        for item_id in self.tree.get_children():
            if self.tree.set(item_id, "export") == "✔":
                self.tree.item(item_id, tags=("export",))
            else:
                self.tree.item(item_id, tags=("normal",))

    def on_filter_updated(self, event):
        self.svars['check_all'].set(self.tree.all_checked(1))
        self.svars['check_all_upload'].set(self.tree.all_checked(2))

    def on_toggle_cell(self, event):
        """Handles cell clicks to change flags."""
        self.update_data_column()
        # self.svars['check_all'].set(self.tree.all_checked(1))
        # self.svars['check_all_upload'].set(self.tree.all_checked(2))

    def toggle_all_export(self):
        checked = self.svars['check_all'].get()
        if not checked:
            self.svars['check_all_upload'].set(False)
        for item in self.tree.get_children():
            values = list(self.tree.item(item, "values"))
            if checked:
                values[1] = "✔"
            else:
                values[1] = " "
                values[2] = " "
            self.tree.item(item, values=values)

    def toggle_all_upload(self):
        checked = self.svars['check_all_upload'].get()
        for item in self.tree.get_children():
            values = list(self.tree.item(item, "values"))
            if checked:
                values[2] = "✔" if values[1] == "✔" else " "
            else:
                values[2] = " "
            self.tree.item(item, values=values)
        if not self.svars['check_all'].get():
            self.svars['check_all_upload'].set(False)

    def btn_run(self):
        """
        Implementation of the "Run" button
        """
        self.export()

    def btn_exit(self):
        self.destroy()
        root.destroy()

    def btn_openf(self):
        """
        Implementation of the "File Open" button
        After selecting a file, the data can be loaded.
        """
        db_path = filedialog.askopenfilename(filetypes=[("MS Access files", "*.mdb, *.accdb")])
        self.db_path.set(db_path)
        self.label1['text'] = f"MS Access database for export: \"{self.db_path.get().split('/')[-1]}\""
        self.label1.update()
        self.db_connect()
        if self.check_permissions():
            self.make_tree()
            self.recreate_widgets()

    def show_permission_warning(self):
        def open_link(event):
            warning_window.destroy()
            webbrowser.open_new(
                "https://github.com/whellcome/MSAccessToSQL?tab=readme-ov-file#important-note-access-permissions")

        warning_window = tk.Toplevel()
        warning_window.title("Access Permission Error")
        warning_window.geometry("345x185")
        warning_window.resizable(False, False)
        spad = 7
        self.render(ttk.Label(warning_window, text="Access Permission Error", font=("Helvetica", 14)),
                    dict(row=0, column=0, pady=spad, columnspan=3, sticky="ns"))
        message = (
            "The MS Access Export Tool requires access to system tables "
            "MSysObjects and MSysRelationships. Please refer to the "
            "documentation for steps to grant the necessary permissions."
        )
        self.render(ttk.Label(warning_window, text=message, wraplength=350, justify="center"),
                    dict(row=1, column=0, columnspan=3, pady=spad))
        link = ttk.Label(
            warning_window, text="Click here for documentation", foreground="blue", cursor="hand2"
        )
        self.render(link, dict(row=2, column=0, columnspan=3, pady=spad, sticky="ns"))
        link.bind("<Button-1>", open_link)
        self.render(tk.Button(warning_window, text=" Close ", command=warning_window.destroy),
                    dict(row=3, column=1, pady=spad))
        warning_window.transient()
        warning_window.grab_set()
        warning_window.mainloop()

    def db_connect(self):
        db_path = self.db_path.get()
        if not db_path:
            return None
        engine = win32com.client.Dispatch("DAO.DBEngine.120")
        self.db = engine.OpenDatabase(db_path)

    def check_permissions(self):
        try:
            recordset = self.db.OpenRecordset("SELECT TOP 1 * FROM MSysObjects")
            recordset.Close()
            recordset = self.db.OpenRecordset("SELECT TOP 1 * FROM MSysRelationships")
            recordset.Close()
            return True
        except:
            self.show_permission_warning()
            return False

    def get_referenced_tables(self, table_name):
        referenced_tables = []
        query = """
            SELECT DISTINCT szReferencedObject AS ReferencedTable
            FROM MSysRelationships
            WHERE szObject = ?
            """
        query_def = self.db.CreateQueryDef("", query)
        query_def.Parameters(0).Value = table_name
        try:
            results = query_def.OpenRecordset()
            while not results.EOF:
                referenced_tables.append(results.Fields("ReferencedTable").Value)
                results.MoveNext()
            results.Close()
        except Exception as e:
            print(f"Error retrieving relationships for {table_name}: {e}")

        return referenced_tables

    def resolve_dependencies(self, export_list):
        export_set = set(export_list)
        added_tables = set()
        while True:
            new_tables = set()
            for table in list(export_set):
                referenced = self.get_referenced_tables(table)
                for ref_table in referenced:
                    if ref_table not in export_set:
                        new_tables.add(ref_table)
            if not new_tables:
                break
            export_set.update(new_tables)
            added_tables.update(new_tables)

        return list(export_set), list(added_tables)

    def export_prepare(self):
        df = self.tree.df
        export_list = df[df.iloc[:, 1] == "✔"]["table"].to_list()
        upload_list = df[df.iloc[:, 2] == "✔"]["table"].to_list()
        final_list, added_tables = self.resolve_dependencies(export_list)

        if added_tables:
            added_tables_str = "\n".join(added_tables)
            message = (
                "The following tables were added to ensure database integrity:\n\n"
                f"{added_tables_str}\n\n"
                "Do you want to continue the export?"
            )
            if not messagebox.askyesno("Integrity Check", message):
                return False
        expath = self.db_path.get().split('/')
        fname = expath[-1]
        catalog = "/".join(expath[:-1])
        output_sql_path = f"{catalog}/{'_'.join(fname.split('.')[:-1])}.sql"

        return final_list, upload_list, output_sql_path

    def export(self):
        export_lists = self.export_prepare()
        if not export_lists:
            return

        with (open(export_lists[2], "w", encoding="utf-8") as sql_file):
            for tab_name in export_lists[0]:
                table = self.db.TableDefs(tab_name)
                sql_file.write(f"-- Table: {table.Name}\n")
                sql_file.write(f"CREATE TABLE '{table.Name}' (\n")
                column_definitions = []
                for field in table.Fields:
                    cNull = 'NOT NULL' if field.Required else ''
                    fSize = f"({field.Size})" if field.Size else ''
                    column_definitions.append(
                        f" '{field.Name}'"
                        f" {self.svars['dao_types'].get(field.Type, 'Unknown')}{fSize}"
                        f" {cNull}"
                    )
                column_primkeys = []
                for index in table.Indexes:
                    if index.Primary:
                        column_primkeys.append(index.Fields[0].Name)
                if len(column_primkeys):
                    keysStr = ",".join(column_primkeys)
                    column_definitions.append(f" PRIMARY KEY ({keysStr} AUTOINCREMENT)")
                relationships_query = """
                                            SELECT szObject AS FK_Table,
                                                   szColumn AS FK_Column,
                                                   szReferencedObject AS PK_Table,
                                                   szReferencedColumn AS PK_Column
                                            FROM MSysRelationships
                                            WHERE szObject = ?
                                            """
                query_def = self.db.CreateQueryDef("", relationships_query)
                query_def.Parameters(0).Value = table.Name
                results = query_def.OpenRecordset()
                while not results.EOF:
                    fk_column = results.Fields("FK_Column").Value
                    pk_table = results.Fields("PK_Table").Value
                    pk_column = results.Fields("PK_Column").Value
                    column_definitions.append(f" FOREIGN KEY ({fk_column})"
                                              f" REFERENCES {pk_table}({pk_column})"
                                              )
                    results.MoveNext()
                results.Close()
                sql_file.write(",\n".join(column_definitions))
                sql_file.write("\n);\n\n")

                if table.Name in export_lists[1]:
                    ref_columns = [field.Name for field in table.Fields]
                    sql_file.write(f"-- Filling data for {table.Name}\n")
                    sql_file.write(f"INSERT INTO '{table.Name}' ({', '.join(ref_columns)}) VALUES\n")
                    recordset = self.db.OpenRecordset(f"SELECT * FROM [{table.Name}]")
                    while not recordset.EOF:
                        values = []
                        for column in ref_columns:
                            value = recordset.Fields(column).Value
                            if value is None:
                                values.append("NULL")
                            elif isinstance(value, str):
                                values.append(f"'{str(value)}'".replace("'", "''"))
                            elif isinstance(value, (int, float)):
                                values.append(str(value))
                            else:
                                values.append(f"'{str(value)}'")
                        insert_query = f" ({', '.join(values)});\n"
                        sql_file.write(insert_query)
                        recordset.MoveNext()
                    recordset.Close()
                    sql_file.write("\n);\n\n")

            messagebox.showinfo("SQL export completed!", f"File saved as {export_lists[2]}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("MS Access Export")
    app = GetWidgetsFrame(master=root, padding=(2, 2))
    app.mainloop()
