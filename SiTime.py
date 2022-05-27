import pandas as pd
import os
from openpyxl import load_workbook

writer = pd.ExcelWriter("summary.xlsx")

directory = './Data/'

summary_columns = ['File_Name', 'Column_Name', 'Inferred_Type',
                   'Populated', 'Unique_Values', 'PctUnique',
                   'Nulls', 'PctMissing',
                   'Top_Values', 'mean', 'median', 'max', 'min']

file_count = 0
sheet_count = 0
for file_name in os.listdir(directory):

    if file_name.lower().endswith(".csv") \
            or file_name.lower().endswith(".xlsx") \
            or file_name.lower().endswith(".xls"):
        print(file_name)
        file_count += 1
        if file_name.lower().endswith(".csv"):
            sheet_list = ['.csv']
        elif file_name.lower().endswith(".xlsx"):
            xls = load_workbook('{}{}'.format(directory, file_name), read_only=True, keep_links=False)
            sheet_list = xls.sheetnames
        elif file_name.lower().endswith(".xls"):
            sheet_list = ['.xls']


        summary = []

        for sheet in sheet_list:
            sheet_count += 1
            print(sheet)
            if file_name.lower().endswith(".csv"):
                df = pd.read_csv('{}{}'.format(directory, file_name), low_memory=False, quotechar='"')
            elif file_name.lower().endswith(".xlsx"):
                df = pd.read_excel('{}{}'.format(directory, file_name), sheet_name=sheet)
            elif file_name.lower().endswith(".xls"):
                df = pd.read_excel('{}{}'.format(directory, file_name))

            top_freq = {}

            for col in df:
                col_name = df[col].name
                col_name_lower = col_name.lower()
                col_counts = df[col].value_counts().to_dict()

                stats = {}
                if df[col].dtypes in ["float64", "int64", "datetime64[ns]"]:
                    stats["mean"] = (df[col].mean())
                    stats["min"] = (df[col].min())
                    stats["max"] = (df[col].max())
                    stats["median"] = (df[col].median())

                unique_values = len(col_counts)

                missing = df[col].isnull()
                missing_count = 0
                for v in missing:
                    if v:
                        missing_count += 1

                {k: v for k, v in sorted(col_counts.items(), key=lambda item: item[1])}
                top_values = list(col_counts)[:unique_values if unique_values < 10 else 5]

                entry = [file_name+' | '+sheet if sheet != '.csv' else file_name,
                        col_name_lower, df[col].dtypes,
                        len(df) - missing_count,
                        unique_values, unique_values / len(df) if len(df) > 0 else 0,
                        missing_count, missing_count / len(df) if len(df) > 0 else 0,
                        top_values,
                        (stats["mean"] if "mean" in stats else ''),
                        (stats["median"] if "median" in stats else ''),
                        (stats["min"] if "min" in stats else ''),
                        (stats["max"] if "max" in stats else '')]
                summary.append(entry)

        summary_df = pd.DataFrame(summary, columns=summary_columns)
        summary_df.to_excel(writer, header=True, sheet_name=file_name[0:30], index=False)
print('Processed {} Files with a total of {} Sheets'.format(file_count, sheet_count))
writer.close()