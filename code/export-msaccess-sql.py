import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser
import win32com.client
from tkextras import *
import json
import argparse
from datetime import datetime

parser = argparse.ArgumentParser(description="MS Access to SQL Export Tool")

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
        self.sql_path = tk.StringVar(self, "")
        self.log_path = tk.StringVar(self, "")
        self.config_path = tk.StringVar(self, "")
        self.lable_frame = ttk.Frame(self, padding=(2, 2))
        self.label_db_path = ttk.Label(self.lable_frame, text="MS Access database", font=("Helvetica", 12),
                                       wraplength=650)
        self.label_sql_path = ttk.Label(self.lable_frame, text="Exported SQL script:", font=("Helvetica", 12),
                                        wraplength=650)
        self.log_frame = ttk.Frame(self, borderwidth=1, relief="solid", padding=(2, 2))
        self.label_log_path = ttk.Label(self.log_frame, text="...", font=("Helvetica", 12), wraplength=650)
        self.frame0 = ttk.Frame(self, width=240, borderwidth=1, relief="solid", padding=(2, 2))
        self.frame1 = ttk.Frame(self, width=100, borderwidth=1, relief="solid", padding=(2, 2))
        self.tree = TreeviewDataFrame(self.frame1, columns=("table", "export", "data"), show="headings")
        self.tree.bind("<<TreeFilterUpdated>>", self.on_filter_updated)
        self.tree.bind("<<TreeCheckAllUpdated>>", self.on_check_all_updated)
        self.tree.bind("<<TreeToggleCell>>", self.on_toggle_cell)
        self.scrollbar = ttk.Scrollbar(self.frame1, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        self.create_widgets()
        self.load_config('config.json', True)

    def create_widgets(self):
        """
        Building the main widgets at the beginning of program execution
        Returns:

        """
        grid = self.rgrid
        grid(self)
        grid(tk.Label(self, text="MS Access to SQL Export Tool", font=("Helvetica", 14)),
             dict(row=0, column=0, columnspan=3, pady=5))
        grid(self.lable_frame, dict(row=1, column=0, columnspan=3))
        grid(self.label_db_path, dict(row=0, column=0, columnspan=3, pady=5))
        grid(self.label_sql_path, dict(row=1, column=0, columnspan=3, pady=5))
        grid(tk.Button(self, text="MS Access File Open", command=self.btn_openf, font=("Helvetica", 11)),
             dict(row=2, column=0))
        grid(tk.Button(self, text="Save SQL script as...", command=self.btn_sql_path, font=("Helvetica", 11)),
             dict(row=2, column=1))
        grid(tk.Button(self, text=" Exit ", command=self.btn_exit, font=("Helvetica", 11)),
             dict(row=2, column=2, columnspan=2))
        grid(self.frame1, dict(row=4, column=0, columnspan=3))
        grid(self.tree, dict(row=0, column=0, pady=5))
        grid(self.scrollbar, dict(row=0, column=3, sticky="ns"))
        grid(tk.Button(self, text=" Save default config ", command=self.save_config, font=("Helvetica", 11)),
             dict(row=5, column=0, ))
        grid(tk.Button(self, text=" Save config as... ", command=self.save_config_as, font=("Helvetica", 11)),
             dict(row=5, column=1, ))
        grid(tk.Button(self, text=" Load config ", command=self.load_config, font=("Helvetica", 11)),
             dict(row=5, column=2, ))
        grid(self.log_frame, dict(row=6, column=0, columnspan=3, pady=5))
        self.log_frame.grid_columnconfigure(1, weight=1)
        grid(tk.Button(self.log_frame, text=" Logging in file: ", command=self.btn_log, font=("Helvetica", 11)),
             dict(row=0, column=0, pady=5))
        grid(self.label_log_path, dict(row=0, column=1, pady=5))
        grid(tk.Button(self.log_frame, text="X", command=self.btn_log_delete, font=("Helvetica", 11)),
             dict(row=0, column=2, padx=5, pady=5, sticky="e"))
        grid(tk.Button(self, text=" Run! ", command=self.btn_run, font=("Helvetica", 12, "bold")),
             dict(row=7, column=0, columnspan=3, pady=5))

    def recreate_widgets(self):
        grid = self.rgrid
        grid(self.tree, dict(row=0, column=0, pady=5))
        grid(self.scrollbar, dict(row=0, column=3, sticky="ns"))
        grid(self.frame0, dict(row=3, column=0, columnspan=3, sticky="e"))
        grid(self.tree.filter_widget(self.frame0), dict(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="ew"))
        grid(self.tree.checkbox_widget(self.frame0), dict(row=4, column=0, columnspan=3, padx=5, pady=5, sticky="e"))

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
        self.tree.tag_configure("selected", background="#fff0f0")

    def update_column_style(self):
        """..."""
        for item_id in self.tree.get_children():
            if "✔" in self.tree.item(item_id, "values"):
                self.tree.item(item_id, tags=("selected",))
            else:
                self.tree.item(item_id, tags=("normal",))

    def on_filter_updated(self, event):
        """

        Args:
            event:

        Returns:

        """
        pass

    def on_check_all_updated(self, event):
        self.update_column_style()

    def on_toggle_cell(self, event):
        """Handles cell clicks to change flags."""
        self.update_column_style()

    def btn_log(self):
        log_path = filedialog.asksaveasfilename(
            filetypes=[("Log files", "*.log")],
            initialfile="export-msaccess-sql.log"
        )
        self.log_path.set(log_path)
        self.update_label_log()

    def btn_log_delete(self):
        self.log_path.set("")
        self.update_label_log()

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
        self.update_label_db()

    def btn_sql_path(self):
        sql_path = filedialog.asksaveasfilename(
            filetypes=[("SQL script files", "*.sql")],
            initialfile=self.get_output_sql_name()
        )
        self.sql_path.set(sql_path)
        self.update_label_sql()

    def save_config(self, file_path='config.json'):
        config = {
            "info": "MS Access to SQL Export configuration file",
            "db_path": self.db_path.get(),
            "sql_path": self.sql_path.get(),
            "log_path": self.log_path.get(),
            "tree": self.tree.df.to_dict()
        }
        with open(file_path, 'w') as f:
            json.dump(config, f, indent=4)

    def save_config_as(self):
        file_path = filedialog.asksaveasfilename(
            title="Save configuration as(",
            defaultextension=".json",
            initialfile="config.json",
            filetypes=[("JSON files", "*.json")]
        )
        if file_path:
            self.save_config(file_path)

    def load_config(self, fpath='config.json', loadbyinit=False):

        try:
            with open(fpath, 'r') as f:
                config = json.load(f)

            if ('info' in config) and (config['info'] == "MS Access to SQL Export configuration file"):
                self.db_path.set(config["db_path"])
                self.sql_path.set(config["sql_path"])
                if "log_path" in config:
                    self.log_path.set(config["log_path"])
                self.update_widgets()
                self.tree.df = self.tree.df.from_dict(config["tree"])
                self.tree.rebuild_tree()
                self.tree.all_checked_update()
            else:
                raise
        except:
            if loadbyinit:
                return
            fpath = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
            if fpath:
                self.load_config(fpath)

    def update_label_db(self):
        self.db_connect()
        if self.check_permissions():
            self.label_db_path['text'] = f"MS Access database: \"{self.db_path.get().split('/')[-1]}\""
            self.make_tree()
            self.recreate_widgets()

    def update_label_sql(self):
        self.label_sql_path['text'] = f"Exported SQL script: {self.sql_path.get()}"

    def update_label_log(self):
        self.label_log_path['text'] = f"{self.log_path.get()}"

    def update_widgets(self):
        self.update_label_db()
        self.update_label_sql()
        self.update_label_log()

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
        grid = self.rgrid
        grid(ttk.Label(warning_window, text="Access Permission Error", font=("Helvetica", 14)),
             dict(row=0, column=0, pady=spad, columnspan=3, sticky="ns"))
        message = (
            "The MS Access Export Tool requires access to system tables "
            "MSysObjects and MSysRelationships. Please refer to the "
            "documentation for steps to grant the necessary permissions."
        )
        grid(ttk.Label(warning_window, text=message, wraplength=350, justify="center"),
             dict(row=1, column=0, columnspan=3, pady=spad))
        link = ttk.Label(
            warning_window, text="Click here for documentation", foreground="blue", cursor="hand2"
        )
        grid(link, dict(row=2, column=0, columnspan=3, pady=spad, sticky="ns"))
        link.bind("<Button-1>", open_link)
        grid(tk.Button(warning_window, text=" Close ", command=warning_window.destroy),
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

    def get_output_sql_name(self):
        db_path = self.db_path.get()
        if db_path:
            expath = db_path.split('/')
            fname = expath[-1]
            catalog = "/".join(expath[:-1])
            return f"{catalog}/{'_'.join(fname.split('.')[:-1])}.sql"
        else:
            return "AccessExport.sql"

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

    def export_prepare(self, output_sql_path="", mode=""):
        df = self.tree.df
        export_list = df[df.iloc[:, 1] == "✔"]["table"].to_list()
        upload_list = df[df.iloc[:, 2] == "✔"]["table"].to_list()
        final_list, added_tables = self.resolve_dependencies(export_list)

        if added_tables and mode != "cmd":
            added_tables_str = "\n".join(added_tables)
            message = (
                "The following tables were added to ensure database integrity:\n\n"
                f"{added_tables_str}\n\n"
                "Do you want to continue the export?"
            )
            if not messagebox.askyesno("Integrity Check", message):
                return False
        if not output_sql_path:
            output_sql_path = self.get_output_sql_name()

        return final_list, upload_list, output_sql_path

    def export(self, mode=""):
        export_lists = self.export_prepare(self.sql_path.get(), mode=mode)
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
            if mode == "cmd":
                print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - SQL export completed!", f"File saved as {export_lists[2]}", sep="\n")
            else:
                messagebox.showinfo("SQL export completed!", f"File saved as {export_lists[2]}")

def main():
    parser.add_argument("-c","--config", type=str, help="Path to config file")
    args = parser.parse_args()

    if args.config:
        log_list = []
        conf = args.config
        log_list.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Configuration`s accepted: {conf}")
        root.withdraw()
        app = GetWidgetsFrame(master=root)
        log_list.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - App started: {app.master.title()}")
        app.load_config(conf)
        log_list.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Configuration`s uploaded:")
        log_list.append(f"\t db path: {app.db_path.get()}")
        log_list.append(f"\t sql path: {app.sql_path.get()}")
        log_list.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Export execution...")
        print("\n".join(log_list))
        app.export(mode="cmd")
        app.btn_exit()

    else:
        app = GetWidgetsFrame(master=root, padding=(2, 2))
        app.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    root.title("MS Access Export")
    main()
