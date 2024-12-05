
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
# Подключение к базе данных Access
db_path = filedialog.askopenfilename(filetypes=[("MS Access files", "*.mdb, *.accdb")])
if db_path:

    # Инициализация DAO
    engine = win32com.client.Dispatch("DAO.DBEngine.120")
    db = engine.OpenDatabase(db_path)

    # Путь для сохранения SQL-скрипта
    output_sql_path = "export_msaccess.sql"

    # Открытие файла для записи
    with (open(output_sql_path, "w", encoding="utf-8") as sql_file):
        for table in db.TableDefs:
            if not table.Name.startswith("MSys"):  # Пропуск системных таблиц

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
                query_def.Parameters(0).Value = table.Name  # Указываем имя таблицы

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
