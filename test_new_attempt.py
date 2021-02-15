import pandas as pd
from information import Information
from pandas import ExcelWriter


if __name__ == "__main__":

    """
    Send each sheet individually and store it there
    """
    
    data = "/Users/liulu/OneDrive/Desktop/Projects/Account/firefly/test_and_output/total_inventory_2017_2021.xlsx"
    data_2 = "/Users/liulu/OneDrive/Desktop/Projects/Account/firefly/test_and_output/test_total_inventory_2.xlsx"

    data_list = [data, data_2]

    companies = {}

    

    for d in data_list:
        xlsx = pd.ExcelFile(d)

        sheets = xlsx.sheet_names
        


        for sheet in sheets:
            if sheet != "other":
                df = pd.read_excel(xlsx, sheet_name=sheet, skiprows=2)

                process = Information(companies, df, sheet)

                companies[sheet] = process.process_excel()

    output_file = "tax_return_price.xlsx"

    with ExcelWriter(output_file) as writer: # pylint: disable=abstract-class-instantiated
        for key in companies:
            to_write = pd.DataFrame.from_dict(companies[key], orient="index")
            to_write.to_excel(writer, sheet_name=key)

    print(f"Done! File created: {output_file}")
    # print(companies)
    
        
    

