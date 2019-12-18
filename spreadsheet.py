from numpy import array, matmul, linalg

import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('test-ad974d72b0b7.json', scope)

gc = gspread.authorize(credentials)

sheet = gc.open('Test').sheet1

# get all input values in each cell
data = array(sheet.get_all_records())

print(data)

# minimising ||ax - b|| to get line of best fit

marks_list = []
selective_result = []
for dict in data:
    marks_list.append([1, dict['Mark']])        # matrix a
    selective_result.append([dict['Selective Score']])      # matrix b

matrix_a = array(marks_list)
matrix_b = array(selective_result)

print(matrix_a)
print(matrix_b)

# need to solve (a^T*a)x = a^T*b

lhs = matmul(matrix_a.transpose(), matrix_a)        # a^T * a
rhs = matmul(matrix_a.transpose(), matrix_b)        # a^T * b

result = linalg.solve(lhs, rhs)
print(f"result is {result}")

intercept = float(result[0, 0])
gradient =  float(result[1, 0])

# write to google sheets
sheet.update_acell('F2', intercept)
sheet.update_acell('F3', gradient)

input_mark = int(sheet.acell('F5').value)

print(f"intercept is {intercept}")
print(f"gradient is {gradient}")
print(f"line of best fit is y = {intercept} + {gradient}x")

print(f"input_mark is {input_mark}")

# using the generated line of best fit to estimate the selective mark
expected_selective_mark = intercept + gradient * input_mark
print(f"expected mark is {expected_selective_mark}")

sheet.update_acell('F6', expected_selective_mark)
