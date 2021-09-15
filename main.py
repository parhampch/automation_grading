import xlrd
import gspread
import yaml


def find_student(excel_worksheet, student_id_col_index, first_student_row_index, number_of_students, student_id):
    for i in range(number_of_students):
        if excel_worksheet.cell(first_student_row_index + number_of_students, student_id_col_index) == student_id:
            return first_student_row_index + number_of_students
    return -1


# load configs
with open("config.yml", "r") as yml_file:
    cfg = yaml.safe_load(yml_file)

# open worksheet
gc = gspread.service_account(filename=cfg['GENERAL']['credentials_dir'])
google_sheet = gc.open(cfg['SHEET']['sheet_name'])
google_worksheet = google_sheet.worksheet(cfg['SHEET']['worksheet_name'])

# open excel file
excel_sheet = xlrd.open_workbook(cfg['EXCEL']['file_dir'])
excel_worksheet = excel_sheet.sheet_by_name(cfg['EXCEL']['worksheet_name'])

if cfg['GENERAL']['use_grade_col_list_sheet']:
    print('test')
else:
    for i in range(cfg['GENERAL']['number_of_questions']):
        student_id = google_worksheet.cell(cfg['SHEET']['first_student_row_index'] + i,
                                           cfg['SHEET']['student_IDs_col_index']).value
        student_row_index = find_student(excel_worksheet, cfg['EXCEL']['student_IDs_col_index'],
                                         cfg['EXCEL']['first_student_row_index'],
                                         cfg['GENERAL']['number_of_questions'], student_id)


#######################
# val = wsh.cell(2, 1).value
# wsh.update_cell(4, 1, 'parham')
# print(val)
'''
configs = {"GENERAL": {"number_of_questions": 10, "number_of_students": 200, "set_score": True, "set_delay": True,
                       "use_col_list_excel": False, "use_col_list_sheet": False},
           "EXCEL": {"file_dir": "", "worksheet_name": "", "student_IDs_col_index": 3, "first_student_row_index": 3, "first_delay_col_index": 5,
                     "question_col_length": 5, "grade_cols": []},
           "SHEET": {"sheet_name": "", "worksheet_name": "", "student_IDs_col_index": 3, "first_student_row_index": 3,
                     "first_delay_col_index": 5, "question_col_length": 5, "grade_cols": []}}
'''


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
