import numpy as np 
import pandas as pd 
from multiprocessing import Pool
import xlsxwriter

class Information(object):
    """
    Use multiprocessing when gathering information

    Read old excel, add information
    Read new excel, replace information

    OR

    Read new excel, add information
    Read old excel, if information not there then add information

    """
    def __init__(self, products, data, sheet_name):
        try:
            self.products = products[sheet_name]
        except:
            self.products = {}
        self.data = data
        self.sheet_name = sheet_name

    def process_excel(self):
        row_loc = 0
        reference_row = self.data.iloc[row_loc].to_list()
        # print(self.data.iloc[3,2])
        
        for i, row in self.data.iterrows():
            x = row.size-1
            checker = False

            while not checker:
                try:
                    # print("i'm starting!")
                    if not pd.isnull(row.iloc[x]) and row.iloc[x] > 0:
                        # print("It's not null!")
                        # print(self.data.iloc[row_loc, x])
                        if self.data.iloc[row_loc, x] == "final cost":
                            # print("I make it here!")
                            product = row.iloc[2]
                            # print((product, row.iloc[x]))
                            if product not in self.products:
                                # print("added!")
                                self.products[product] = row.iloc[x]
                            checker = True
                    x -= 1
                    # print("It is null!")
                except Exception as e:
                    checker = True
        return self.products
                
        # print(self.products)
    
        
