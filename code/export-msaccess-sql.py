
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import win32com.client

dao_types = {
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
}


class GetWigetsFrame(tk.Frame):
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
        if render_params is None:
            render_params = dict(sticky="ew", padx=5, pady=2)

        self.__render_params = render_params
        self.db_path = tk.StringVar(self, "")
        self.label1 = tk.Label(self, text="", font=("Helvetica", 12))
        self.frame1 = ttk.Frame(self, width=100, borderwidth=1, relief="solid", padding=(2, 2))
        self.create_widgets()

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

    def create_widgets(self):
        """Building the main widgets at the beginning of program execution"""
        self.render(self)
        self.render(tk.Label(self, text="MS Access to SQL Export Tool", font=("Helvetica", 14)),
                    dict(row=0, column=0, columnspan=2, pady=5))
        self.render(self.label1, dict(row=1, column=0, columnspan=3))
        self.render(tk.Button(self, text="MS Access File Open", command=self.btn_openf), dict(row=2, column=0, columnspan=2))
        self.render(tk.Button(self, text=" Exit ", command=self.quit), dict(row=2, column=2, columnspan=2))
        self.render(tk.Label(self, text=" --------------------------------------------------------- "),
                    dict(row=3, column=0, columnspan=3, pady=5))
        self.render(tk.Button(self, text=" Run! ", command=self.btn_run, font=("Helvetica", 12)),
                    dict(row=4, column=0, columnspan=3, ))

    def recreate_widgets(self):
        pass

    def btn_run(self):
        """
        Implementation of the "Show" button click event
        Based on the form data, a dataframe is generated for plotting and sent for drawing.
        In the case when the comparison mode is selected, a list with two dataframes is sent
        """
        self.export()

    def btn_openf(self):
        """
        Implementation of the "File Open" button click event
        After selecting a file, the data is loaded and cleared.
        If successful, the program continues and loads elements for selecting plotting options
        If unsuccessful, prompts to select another file or exit the program
        """
        db_path = filedialog.askopenfilename(filetypes=[("MS Access files", "*.mdb, *.accdb")])
        self.db_path.set(db_path)
        self.label1['text'] = f"MS Access database for export: \"{self.db_path.get().split('/')[-1]}\""
        self.label1.update()
        self.recreate_widgets()

    def export(self):
        db_path = self.db_path.get()
        if not db_path:
            return None
        engine = win32com.client.Dispatch("DAO.DBEngine.120")
        db = engine.OpenDatabase(db_path)

        output_sql_path = "export_msaccess.sql"

        with (open(output_sql_path, "w", encoding="utf-8") as sql_file):
            for table in db.TableDefs:
                if not table.Name.startswith("MSys"):

                    sql_file.write(f"-- Table: {table.Name}\n")
                    sql_file.write(f"CREATE TABLE '{table.Name}' (\n")

                    column_definitions = []
                    for field in table.Fields:
                        cNull = 'NOT NULL' if field.Required else ''
                        fSize = f"({field.Size})" if field.Size else ''
                        column_definitions.append(
                            f" '{field.Name}'"
                            f" {dao_types.get(field.Type, 'Unknown')}{fSize}"
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
                    query_def = db.CreateQueryDef("", relationships_query)
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
                        recordset = db.OpenRecordset(f"SELECT * FROM [{table.Name}]")

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
    app = GetWigetsFrame(master=root, padx=20, pady=10)
    app.mainloop()
