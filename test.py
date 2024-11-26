# a = {
#     'worksheets_count': 1,
#     1: {
#         '{{CONTRACT_NUMBER}}': 'str(currentSheetvalue)',
#         '{{CONTRACT_YEAR}}': '(currentSheet.value),  # Короткий год - 24',
#         '{{WORK_NUMBER}}': '(currentSheet[value)',

#     },
#     2: {
#         '{{BILL_NUMBER}}': '(currentShee.value),  # Короткий год - 24',
#         '{{CONTR_DATE}}': '(currentShevalue)',
#     }
# }
# print(a[1]['{{CONTRACT_NUMBER}}'])

variables = {}

a = 'CONTRACT_NUMBER'
variables['{{' + a + '}}'] = '123'
variables['{{CONTRACT_YEAR}}'] = '456'
print(variables)