# Abrir archivos xlsx
import xlwings as xw
# Manejar Data Frames
import pandas as pd
# Finalizar procesos
import psutil

for process in psutil.process_iter():
    if process.name() == "EXCEL.EXE":
        process.kill()
        
path_working_directory = "H:/Mi unidad/gc/AWS/automation/create_db_from_xlsx"
path_file_xlsx = path_working_directory + '/' + 'DataDefinition.xlsx'

app = xw.App(visible = False)
app.properties(display_alerts = False)

wb_db = app.books.open(path_file_xlsx)
i = 0
text_file = open(path_working_directory+'/'+'Output.txt', "w")
text_file.close()
script = ''
for sheet in wb_db.sheets:
    i = i + 1
    # print(sheet.name)
    db_name = sheet.name
    script_create_DB = 'CREATE Table [' + db_name + ']('
    j = 0
    # Write txt file in python?
    text_file = open(path_working_directory+'/'+'Output.txt', "a")
    text_file.write(script_create_DB)
    pk=''
    for value in sheet.used_range.value:
        
        default_values = ''
        if j>0 :
            column_name = value[0]
            if j==1:
                pk = column_name
            column_position = value[1]
            column_type = value[2].strip()
            column_size = int(value[3])
            column_decimals = value[4]
            column_key = value[5]
            if value[6] == 'X':
                column_not_null = 'NOT NULL'
            else:
                column_not_null = ''    
            column_default = value[7]
            # print(value[0])
            if column_type == 'INTEGER':
                column_type_text = '[int]'
            else:
                column_type_text = f'[{column_type}]({column_size})'

            each_column = f'[{column_name}] {column_type_text} {column_not_null}, \n'
            text_file.write(each_column)
            default_values += f"""  ALTER TABLE [{db_name}] ADD  CONSTRAINT [DF_{db_name}_{column_name}]  DEFAULT ('{column_default}') FOR [{column_name}]\n
                                    GO \n"""
        j = j + 1
    text_file.write(f'CONSTRAINT [PK_{db_name}] PRIMARY KEY CLUSTERED ([{pk}] ASC)\n')
    text_file.write(""" WITH    (PAD_INDEX = OFF, 
                                STATISTICS_NORECOMPUTE = OFF, 
                                IGNORE_DUP_KEY = OFF, 
                                ALLOW_ROW_LOCKS = ON, 
                                ALLOW_PAGE_LOCKS = ON, 
                                OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) 
                            ON [PRIMARY] \n
                        )ON [PRIMARY]\n
                        GO\n""")

    text_file.close()