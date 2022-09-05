import pandas as pd

#This function will manipulate records in County Column
def Matching(data):
    Argentina = ['ARGPAE', 'ARGTSE', 'ARGVista', 'ARGYPF']
    Colombia = ['Colombia TSE', 'Colombia Ecopetrol', 'Colombia Frontera', 'CGX Guyana', 'Parex']
    LEA = 'UK LEA'

    #Sroting Argentina
    for i in Argentina:
        result = i in str(data)
        #print(str(data), "&", i)
        if (result):
            return "Argentina"

    # Sroting Colombia
    for i in Colombia:
        result = i in str(data)
        #print("result: ", result)
        # print(str(data), "&", i)
        if (result):
         return "Colombia"

    result = LEA in str(data)
    if (result):
        return 'UK'

    return data

print("Processing...")
spreadsheet = pd.ExcelFile('Global LEA Virtual Inventory.xlsx')
req_worksheets = [ 'Saudi LEA','Colombia LEA', 'Argentina LEA', 'Egypt LEA', 'Saudi Azure Stack', 'Ecuador LEA', 'UK LEA']
worksheets = spreadsheet.sheet_names
appended_data = []
for sheet_name in req_worksheets:
    df = pd.read_excel(spreadsheet,sheet_name)
    df.rename(columns={'Country ': 'Country-Tenent'}, inplace=True)
    df.insert(0, 'Country ', '')
    df.insert(16,'Agent Exception', '')

    PS_result = df['Power State'].values
    #Filling blank cells as Running
    df['Power State'].fillna("Running", inplace=True)

    # 2 Lines : copy pasting data from Patching Exception to Agent Exception
    PE_result = df['Patching Exception']
    df['Agent Exception'] = PE_result


    # 2 Lines : Manipulating data in Country columns based on data in Country-Tenent
    result = Matching(df['Country-Tenent'])
    df['Country '] = result

    #Checking and Deleting additional Subscription column
    if "Subscription" in df.columns:
        df.pop("Subscription")
    appended_data.append(df)

# sorting by first name
#appended_data.sort_values("NetBIOS(Host name)", inplace=True)


appended_data = pd.concat(appended_data)
# dropping ALL duplicate values
appended_data.drop_duplicates(subset="NetBIOS(Host name)",
                              keep='first', inplace=True)
#Removing Extra Unamed Columns
appended_data = appended_data.loc[:, ~appended_data.columns.str.contains('^Unnamed')]

appended_data.to_excel('LEA Consolidated.xlsx', index=False)
print("Completed..!")






