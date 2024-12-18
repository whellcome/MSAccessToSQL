import webbrowser
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import win32com.client
import pandas as pd


class WidgetsRender():
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
               :param kwargs: Additional arguments for Treeview insert (e.g., values).
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

        if value is None:
            if column is None:
                self.df.loc[item] = self.df.loc[item].replace(result)
            else:
                self.df.loc[item, column] = result
        else:
            self.df.loc[item, column] = value
        return result

    def item(self, item, option=None, **kw):
        """
        Override item method with DataFrame.
        """
        values = kw.get("values", [])
        result = super().item(item, option, **kw)
        if option is None and len(values):
            updates = pd.Series(values, index=self.cget("columns"))
            self.df.loc[item] = updates
        return result

    def delete(self, *items, inplace = False):
        """
        Override delete method with DataFrame..
        """
        if inplace:
            for item in items:
                values = self.item(item, "values")
                self.df = self.df[~(self.df[list(self.df.columns)] == values).all(axis=1)]
        super().delete(*items)

    def rebuild_tree(self, dataframe=None):
        if dataframe is None:
            dataframe = self.df
        self.delete(*self.get_children())
        for index, row in dataframe.iterrows():
            self.insert("", "end", iid=index, values=row.to_list())

    def filter_by_name(self, keyword):
        """Filter rows based on a keyword in first column and update Treeview."""
        filtered_df = self.df[self.df[self.df.columns[0]].str.contains(keyword, case=False)]
        self.rebuild_tree(filtered_df)


    def filter_widget(self, parent):
        widget_frame = ttk.Frame(parent, width=150, borderwidth=1, relief="solid", padding=(2, 2))
        self.render(tk.Label(widget_frame, text="Filter by table name:", font=("Helvetica", 9,"bold")),
                    dict(row=0, column=0, pady=5))
        filter_entry = tk.Entry(widget_frame)
        self.render(filter_entry, dict(row=0, column=1, padx=5, pady=5, sticky="ew"))

        def apply_filter():
            self.filter_by_name(filter_entry.get())

        def clear_filter():
            self.rebuild_tree()
            filter_entry.delete(0, tk.END)

        self.render(ttk.Button(widget_frame, text="Filter", command=apply_filter),
                    dict(row=0, column=2, padx=5, pady=5))
        self.render(ttk.Button(widget_frame, text="Restore", command=clear_filter),
                    dict(row=0, column=3, padx=5, pady=5))
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
        self.filter_entry = ttk.Entry(self.frame0)
        self.frame1 = ttk.Frame(self, width=100, borderwidth=1, relief="solid", padding=(2, 2))
        self.tree = TreeviewDataFrame(self.frame1, columns=("table", "export", "data"), show="headings")
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
        self.render(self.tree.filter_widget(self.frame0),dict(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="ew"))
        self.render(tk.Label(self.frame0, text=" ", width=43),
                    dict(row=1, column=0, columnspan=2, pady=5))
        self.svars['check_all'] = tk.IntVar(value=0)
        self.render(ttk.Checkbutton(self.frame0, text="Check all to Export", variable=self.svars['check_all'],
                                    command=self.toggle_all_export), dict(row=1, column=2, padx=20))
        self.svars['check_all_upload'] = tk.IntVar(value=0)
        self.render(ttk.Checkbutton(self.frame0, text="Check all to Upload", variable=self.svars['check_all_upload'],
                                    command=self.toggle_all_upload, ), dict(row=1, column=3, padx=20))

    def make_tree(self):
        self.tree.heading("table", text="Table")
        self.tree.heading("export", text="Export")
        self.tree.heading("data", text="Upload")
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
        self.tree.bind("<Button-1>", self.toggle_cell)
        self.tree.tag_configure("normal")
        self.tree.tag_configure("export", background="#fff0f0")

    def update_data_column(self, event):
        """..."""
        for item_id in self.tree.get_children():
            is_red = self.tree.set(item_id, "export")
            if is_red == "✔":
                self.tree.item(item_id, tags=("export",))
            else:
                self.tree.item(item_id, tags=("normal",))

    def toggle_cell(self, event):
        """Handles cell clicks to change flags."""
        tree = self.tree
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = tree.identify_column(event.x)
        item = tree.identify_row(event.y)
        if col == "#2":
            current_value = tree.set(item, "export")
            tree.set(item, "export", " " if current_value == "✔" else "✔")
            if current_value == "✔":
                tree.set(item, "data", " ")
                if self.svars['check_all'].get() == 1:
                    self.svars['check_all'].set(0)
                    self.svars['check_all_upload'].set(0)

        elif col == "#3":
            current_value = tree.set(item, "data")
            if tree.set(item, "export") == "✔":
                tree.set(item, "data", " " if current_value == "✔" else "✔")
            if current_value == "✔":
                if self.svars['check_all_upload'].get() == 1:
                    self.svars['check_all_upload'].set(0)

        self.update_data_column(None)

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
        except Exception as e:
            self.show_permission_warning()
            return False

    def export(self):
        expath = self.db_path.get().split('/')
        fname = expath[-1]
        catalog = "/".join(expath[:-1])
        output_sql_path = f"{catalog}/{'_'.join(fname.split('.')[:-1])}.sql"
        with (open(output_sql_path, "w", encoding="utf-8") as sql_file):
            for table in self.db.TableDefs:
                if not table.Name.startswith("MSys"):
                    # ********** TableDefs("name") get object by name*************
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
                    column_primkeys = []
                    for index in table.Indexes:
                        if index.Primary:
                            column_primkeys.append(index.Fields[0].Name)
                    if len(column_primkeys):
                        keysStr = ",".join(column_primkeys)
                        column_definitions.append(f" PRIMARY KEY ({keysStr} AUTOINCREMENT)")
                    sql_file.write(",\n".join(column_definitions))
                    sql_file.write("\n);\n\n")
                    if table.Name.startswith("Ref_"):
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
        print(f"SQL export completed. File saved as {output_sql_path}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("MS Access Export")
    app = GetWidgetsFrame(master=root, padding=(2, 2))
    app.mainloop()
