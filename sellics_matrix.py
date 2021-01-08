# importing modules
import pandas as pd
import numpy as np

# enabling displaying all columns
pd.set_option('display.max_columns', None)

# importing data
perf_by_asin = pd.read_excel('Emma and Noah Data Sets - September 2020.xlsx',
                             sheet_name='perf_by_asin',
                             usecols=['ad_group_id', 'asin'])
mat = perf_by_asin
ad_group_id = mat.ad_group_id
asin = mat.asin

# get unique values to build matrix
mat_col = list(map(str, mat.ad_group_id.unique()))
mat_row = mat.asin.unique()
mat_row = mat_row[~pd.isnull(mat_row)]
mat_row = list(mat_row)

# building and reframing maxtrix
matrix = pd.DataFrame(np.zeros((len(mat_col)+1, len(mat_row))))
matrix.columns = mat_row
matrix = matrix.drop([0])

matrix = matrix.T
matrix.columns = mat_col

# populating in matrix
for i in range(0, len(asin)-1):
    for j in range(0, len(matrix.index)-1):
        for n in range(0, len(matrix.columns)-1):
            if asin[i] == matrix.index[j]:
                if ad_group_id[i] == int(matrix.columns[n]):
                    matrix.iloc[j][n] = 1

#exporting matrix
matrix.to_excel('sellisc_matrix.xlsx', index = True)
