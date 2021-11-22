#import from Find_Debit_Credit_cells python program
from Find_Debit_Credit_cells import worksheet, debit_row, debit_column, credit_column, credit_row


#Loop and collect values from the selected columns and sum
def loopcolumn(fromA, toB, toprint):

    number = 0.0
    for i in range(debit_row, worksheet.max_row):
        for col in worksheet.iter_cols(fromA,toB):
            if str(col[i].value) != 'None':
                number += float(col[i].value)
    print(toprint + str(number))          
    return number

loopcolumn(debit_column, debit_column, " Total Debit: ")
loopcolumn(credit_column, credit_column, "Total Credit: ")
