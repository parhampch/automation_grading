import xlrd
import gspread
import yaml
import time


def find_student(excel_worksheet, student_id_col_index, first_student_row_index, number_of_students, student_id):
    for i in range(number_of_students):
        if excel_worksheet.cell(first_student_row_index + i, student_id_col_index).value == student_id:
            return first_student_row_index + i
    return -1


def load_config(file_dir):
    with open(file_dir, "r") as yml_file:
        cfg = yaml.safe_load(yml_file)
    return cfg


def open_google_worksheet(cfg):
    gc = gspread.service_account(filename=cfg['GENERAL']['credentials_dir'])
    google_sheet = gc.open(cfg['SHEET']['sheet_name'])
    google_worksheet = google_sheet.worksheet(cfg['SHEET']['worksheet_name'])
    return google_worksheet


def open_excel_worksheet(cfg):
    excel_sheet = xlrd.open_workbook(cfg['EXCEL']['file_dir'])
    excel_worksheet = excel_sheet.sheet_by_name(cfg['EXCEL']['worksheet_name'])
    return excel_worksheet


def set_grade_of_a_student(cfg, google_worksheet, excel_worksheet, student_google_sheet_row_index,
                           student_excel_row_index):
    for j in range(cfg['GENERAL']['number_of_questions']):
        if cfg['GENERAL']['use_grade_col_list_excel']:
            excel_question_col_index = cfg['EXCEL']['grade_cols'][j]
            google_sheet_question_col_index = cfg['SHEET']['grade_cols'][j]
        else:
            excel_question_col_index = cfg['EXCEL']['first_question_col_index'] + \
                                       j * cfg['EXCEL']['question_col_length']
            google_sheet_question_col_index = cfg['SHEET']['first_question_col_index'] + \
                                              j * cfg['SHEET']['question_col_length']
        if excel_worksheet.cell_type(student_excel_row_index, excel_question_col_index) == 0:
            grade = 0
        else:
            grade = excel_worksheet.cell(student_excel_row_index, excel_question_col_index).value
        google_worksheet.update_cell(student_google_sheet_row_index, google_sheet_question_col_index, grade)


def set_delay_of_a_student(cfg, google_worksheet, excel_worksheet, student_google_sheet_row_index,
                           student_excel_row_index):
    delay = 0
    for j in range(cfg['GENERAL']['number_of_questions']):
        if cfg['GENERAL']['use_delay_col_list_excel']:
            excel_delay_col_index = cfg['EXCEL']['grade_cols'][j]
        else:
            excel_delay_col_index = cfg['EXCEL']['first_delay_col_index'] + \
                                       j * cfg['EXCEL']['question_col_length']
        if excel_worksheet.cell_type(student_excel_row_index, excel_delay_col_index) == 0:
            temp = 100
        else:
            temp = int(excel_worksheet.cell(student_excel_row_index, excel_delay_col_index).value)
        delay = max(delay, (100 - temp) // 2)
    google_worksheet.update_cell(student_google_sheet_row_index, cfg['SHEET']['delay_col_index'], delay)


def main():
    # load configs
    cfg = load_config('config.yml')

    # open worksheet
    google_worksheet = open_google_worksheet(cfg)

    # open excel file
    excel_worksheet = open_excel_worksheet(cfg)

    for i in range(cfg['GENERAL']['number_of_students']):
        if i != 0 and i % 5 == 0:
            time.sleep(10)
        student_id = google_worksheet.cell(cfg['SHEET']['first_student_row_index'] + i,
                                           cfg['SHEET']['student_IDs_col_index']).value
        student_excel_row_index = find_student(excel_worksheet, cfg['EXCEL']['student_IDs_col_index'],
                                         cfg['EXCEL']['first_student_row_index'],
                                         cfg['GENERAL']['number_of_students'], student_id)
        print(student_id)
        # set grades
        if cfg['GENERAL']['set_score']:
            set_grade_of_a_student(cfg, google_worksheet, excel_worksheet,
                                   cfg['SHEET']['first_student_row_index'] + i, student_excel_row_index)
        # set delays
        if cfg['GENERAL']['set_delay']:
            set_delay_of_a_student(cfg, google_worksheet, excel_worksheet,
                                    cfg['SHEET']['first_student_row_index'] + i, student_excel_row_index)


if __name__ == '__main__':
    main()
