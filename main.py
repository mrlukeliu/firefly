import pandas as pd 
import xlsxwriter
import xlrd

'''
Multiple sheet/file functionality and working Locate Price function
'''

store_supplier_names = []
store_supplier_names_2 = []

# Put most recent file in origin_file_1
origin_file_1 = "Total_inventory-2017-2021.xlsx"
origin_file_2 = "test_total_inventory_2.xlsx"

# File to be created (change name every time or delete previous file)
destination_file = "Updated_prices_2020-2.xlsx"

xlsx_1 = pd.ExcelFile(origin_file_1)
xlsx_2 = pd.ExcelFile(origin_file_2)

writer = pd.ExcelWriter(destination_file, engine="xlsxwriter") # pylint: disable=abstract-class-instantiated

def main():

    print('working')

    for sheet in xlsx_1.sheet_names:
        store_supplier_names.append(sheet)


    for sheet in xlsx_2.sheet_names:
        store_supplier_names_2.append(sheet)
    
    x = 0
    y = 0

    for supplier in store_supplier_names:
        df = pd.read_excel(origin_file_1, sheet_name=supplier, skiprows=3)
        price_list = {}
        for index, row in df.iterrows():
            if df.iloc[index, 2] not in price_list:
                price_list[df.iloc[index, 2]] = locate_price(df, index)

        if supplier in store_supplier_names_2:
            for supplier in store_supplier_names_2:
                df2 = pd.read_excel(origin_file_2, sheet_name=supplier, skiprows=3)

                for index, row in df2.iterrows():
                    if df2.iloc[index, 2] not in price_list:
                        price_list[df2.iloc[index, 2]] = locate_price(df, index)
        write_to_excel(price_list, store_supplier_names[x])

        x += 1

# if the suppliers are not in the new file (which is unlikely)
    for supplier in store_supplier_names_2:
        price_list = {}
        if supplier not in store_supplier_names:
            df2 = pd.read_excel(origin_file_2, sheet_name=supplier, skiprows=3)

            for index, row in df2.iterrows():
                if df2.iloc[index, 2] not in price_list:
                    price_list[df2.iloc[index, 2]] = locate_price(df2, index)
            write_to_excel(price_list, store_supplier_names_2[y])

            y += 1
    writer.save()


def locate_price(df, index):

    i = -1
    while i > -len(df.columns) + 1:
        if 'final cost' in df.columns[i] and not pd.isnull(df.iloc[index, i]):
            return df.iloc[index, i]
        i -= 1


def write_to_excel(price_list, supplier):
    pricing = pd.DataFrame.from_dict(price_list, orient='index')
    pricing.to_excel(writer, sheet_name=supplier)
    writer.sheets[supplier].set_column('A:A', 75)
    writer.sheets[supplier].set_column('B:B', 10)
    
    
    
if __name__ == "__main__":
    main()
    print("Successful! Finished.")
    
        
        
