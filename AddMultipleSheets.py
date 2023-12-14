import os
import pandas as pd

print('What service is this for: ')
service = input()
Path = ("" + service + "/")
# path needs to be defined
output_file = (service + 'UserAudit.xlsx')

if os.path.exists(output_file):
    xlsx = pd.ExcelFile(output_file)
    sheet_dict = {sheet_name: xlsx.parse(sheet_name)
                  for sheet_name in xlsx.sheet_names}
else:
    sheet_dict = {}

for file in os.listdir(Path):
    if file.endswith(".csv"):
        print(file)
        protocol = file[0:6]
        data = pd.read_csv(Path + str(file))
        df = pd.DataFrame(
            data, columns=["Email", "Study", "First Name", "Last Name", "Access Needed (Yes/No)"])
        df = df.dropna(subset=['Email'])

        condition_Notmdsol = ~df['Email'].str.contains('mdsol', case=False)

        crit_emails = [
            'singhj2@mskcc.org', 'zarskia@mskcc.org', 'nallyb@mskcc.org', 'carond@mskcc.org',
            'chod@mskcc.org', 'pachecoh@mskcc.org', 'truongh@mskcc.org', 'lengfelj@mskcc.org',
            'panzarem@mskcc.org', 'dagostr1@mskcc.org', 'osheas@mskcc.org'
        ]
        condition_crit_emails = ~df['Email'].isin(crit_emails)

        result_df = df[condition_Notmdsol & condition_crit_emails]

        print(result_df)
        sheet_dict[protocol] = result_df

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name, df_sheet in sheet_dict.items():
        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
