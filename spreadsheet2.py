import gspread
import statistics
import math
from oauth2client.service_account import ServiceAccountCredentials

from numpy import array, matmul, linalg

'''
data is given as a list of dictionaries with the following keys:
- Student name (string)
- English (int)
- Math (int)
- Selective Score (int)
Reads the list and extracts the required marks into a new list
'''
def get_subject_marks(data, subject):
    marks = []
    for student in data:
        marks.append(student[subject])
    return marks

# sum from 0 to n of (x_i - mean(x))(y_i - mean(y))
def sample_covariance(data_x, data_y):
    assert len(data_x) == len(data_y)
    covariance = 0
    for i in range(len(data_x)):
        expression = (data_x[i] - statistics.mean(data_x)) * (data_y[i] - statistics.mean(data_y))
        covariance += expression
    return covariance

# sum from 0 to n of (x_i - mean(x))^2
def sample_std(data_list):
    std = 0
    for i in range(len(data_list)):
        expression = (data_list[i] - statistics.mean(data_list)) ** 2
        std += expression

    return math.sqrt(std)

# cov(X,Y) / [ std(x) * std(y) ]
def sample_correlation(data_x, data_y):
    cov = sample_covariance(data_x, data_y)
    sigma_x = sample_std(data_x)
    sigma_y = sample_std(data_y)

    corr = cov / (sigma_x * sigma_y)
    return corr

# returns a list of the correlation of english, maths and ga to selective marks
def get_all_correlation(data):
    english_marks = get_subject_marks(data, "English")
    math_marks = get_subject_marks(data, "Math")
    ga_marks = get_subject_marks(data, "GA")
    selective_score = get_subject_marks(data, "Selective Score")

    english_corr = sample_correlation(english_marks, selective_score)
    math_corr = sample_correlation(math_marks, selective_score)
    ga_corr = sample_correlation(ga_marks, selective_score)

    return [english_corr, math_corr, ga_corr]

def update_sheet_correlation(data):
    list_corr = get_all_correlation(data)

    # updates the cell ranges 3 spots below the Corr (python) cell found in spreadsheet
    cell = worksheet.find("Corr (Python)")
    for i in range(1, 4):
        worksheet.update_cell(cell.row + i, cell.col, list_corr[i-1])
    return

# assuming order of correlation is english, maths and ga according to sheet
# returns a list of the weights of each subject, based on their correlation values
def get_correlation_weights(data):
    list_corr = get_all_correlation(data)
    total = sum(list_corr)
    list_weight = []
    for i in range(3):
        weight = list_corr[i] / total
        list_weight.append(weight)

    return list_weight

def update_sheet_weight(data):
    cell = worksheet.find("Weight (Python)")
    list_weight = get_correlation_weights(data)
    # updates the cell ranges 3 spots below the Weight (python) cell found in spreadsheet
    for i in range(1, 4):
        worksheet.update_cell(cell.row + i, cell.col, list_weight[i-1])
    return

# calculates the averages, weighted based on the correlations determined previously
def weighted_corr_mark(data, english_mark, math_mark, ga_mark):
    list_weight = get_correlation_weights(data)
    mark = english_mark * list_weight[0] + math_mark * list_weight[1] + ga_mark * list_weight[2]

    return mark

def update_weighted_corr_mark(data):
    english_marks = get_subject_marks(data, "English")
    math_marks = get_subject_marks(data, "Math")
    ga_marks = get_subject_marks(data, "GA")

    cell = worksheet.find("WAM (Python)")
    for i in range(len(english_marks)):
        mark = weighted_corr_mark(data, english_marks[i], math_marks[i], ga_marks[i])
        worksheet.update_cell(cell.row + i + 1, cell.col, mark)
    return

# reads the spreadsheet for given subject input marks
def read_input_marks():
    english = worksheet.find("Input English")
    math = worksheet.find("Input Math")
    ga = worksheet.find("Input GA")

    english_mark = worksheet.cell(english.row, english.col + 1)
    math_mark = worksheet.cell(math.row, math.col + 1)
    ga_mark = worksheet.cell(ga.row, ga.col + 1)

    return [english_mark, math_mark, ga_mark]

def get_num_students(data):
    english_marks = get_subject_marks(data, "English")
    return len(english_marks)

'''
Generates line of best fit using linear equations / matrices
Does so by minimising ||ax - b||, where:
- a = [ < 1 ... 1 > < x_1 ... x_n > ] ^ T
- b = [ < y_1 ... y_n > ] ^ T
Then solving (a^T * a) x = a^T * b
'''
def equation_best_fit(data):
    num_students = get_num_students(data)
    mark_cell = worksheet.find("WAM (Python)")
    sel_cell = worksheet.find("Selective Score")

    # construct matrix a
    marks_list = []
    selective_score_list = []
    for i in range(num_students):
        mark = worksheet.cell(mark_cell.row + i, mark_cell.col).value
        marks_list.append([1, mark])
        selective_mark = worksheet.cell(sel_cell.row + i, sel_cell.col).value
        selective_score_list.append(selective_mark)

    matrix_a = array(marks_list)
    matrix_b = array(selective_score_list)
    print(matrix_a)
    print(matrix_b)

    lhs = matmul(matrix_a.transpose(), matrix_a)        # a^T * a
    rhs = matmul(matrix_a.transpose(), matrix_b)        # a^T * b

    # Solving linear equation to get a pair of solutions: (intercept, gradient)
    result = linalg.solve(lhs, rhs)
    intercept = float(result[0, 0])
    gradient =  float(result[1, 0])

    return intercept, gradient

def update_expected_mark(data):
    # calculate the weighted average mark to interpolate from line of best fit
    input_marks = read_input_marks(data)
    weighted_mark = weighted_corr_mark(data, input_marks[0], input_marks[1], input_marks[1])

    interecept, gradient = equation_best_fit(data)
    expected_selective_mark = intercept + gradient * weighted_mark
    print(expected_selective_mark)

    cell = worksheet.find("Output Mark")
    worksheet.update_cell(cell.row, cell.col + 1, expected_selective_mark)
    return


if __name__ == '__main__':
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('test-ad974d72b0b7.json', scope)
    gc = gspread.authorize(credentials)

    worksheet = gc.open('Test').worksheet("Sheet2")

    # get all input values in each cell
    data = array(worksheet.get_all_records())
    print(data)

    update_sheet_correlation(data)
    update_sheet_weight(data)
    update_weighted_corr_mark(data)
    equation_best_fit(data)
    update_expected_mark(data)
