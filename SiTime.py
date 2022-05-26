import pandas as pd
import os

def report():
    print("Summary of {}".format(filename))
    print("Potential Keys:  {}".format(', '.join([str(elem) for elem in possible_keys])))
    print("Top 5 Values by Column")
    for k, v in top_freq.items():
        if k in possible_keys:
            k += '    *****Possible Key'
        print("\nColumn:  {}".format(k))
        for k2, v2, in v.items():
            print('{0:<50}{1:<10}'.format(k2, v2))


id_list = ["id", "key", "pk", "fk"]
not_id = ["middle"]

writer = pd.ExcelWriter("summary.xlsx")

directory = './Data/'

summary_columns = ['File_Name', 'Column_Name',
                   'Populated', 'Uniuqe_Values', 'PctUnique',
                   'Nulls', 'PctMissing',
                   'Top_Values', 'mean', 'median', 'max', 'min']

for file_name in os.listdir(directory):

    if file_name.endswith(".csv") or file_name.endswith(".xlsx"):
        print(file_name)
        possible_keys = []
        if file_name.endswith(".csv"):
            df = pd.read_csv('{}{}'.format(directory, file_name), low_memory=False, quotechar='"')
        elif file_name.endswith(".xls") or file_name.endswith(".xlsx"):
            df = pd.read_excel('{}{}'.format(directory, file_name))

        summary = []

        top_freq = {}

        for col in df:
            col_name = df[col].name
            col_name_lower = col_name.lower()
            col_counts = df[col].value_counts().to_dict()

            stats = {}
            if df[col].dtypes == "float64":
                stats["mean"] = (df[col].mean())
                stats["min"] = (df[col].min())
                stats["max"] = (df[col].max())
                stats["median"] = (df[col].median())
            #                print(stats)
            #            else:
            #                col_counts_nocase = df[col].str.lower().value_counts().to_dict()
            #                unique_values_nocase = len(col_counts_nocase)

            unique_values = len(col_counts)

            missing = df[col].isnull()
            missing_count = 0
            for v in missing:
                if v:
                    missing_count += 1

            {k: v for k, v in sorted(col_counts.items(), key=lambda item: item[1])}
            top_values = list(col_counts)[:unique_values if unique_values < 10 else 5]

            entry = [file_name, col_name_lower, len(df) - missing_count,
                     unique_values, unique_values / len(df),
                     missing_count, missing_count / len(df),
                     top_values,
                     (stats["mean"] if "mean" in stats else 'na'),
                    (stats["median"] if "median" in stats else 'na'),
                     (stats["min"] if "min" in stats else 'na'),
                     (stats["max"] if "max" in stats else 'na')]
            summary.append(entry)

        # print(summary)
        summary_df = pd.DataFrame(summary, columns=summary_columns)
        # print(summary_df)
        summary_df.to_excel(writer, header=True, sheet_name=file_name[0:30], index=False)
writer.close()