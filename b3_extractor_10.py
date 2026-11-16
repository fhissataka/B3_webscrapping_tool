import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from datetime import datetime
import requests
from bs4 import BeautifulSoup

def fetch_PRE_from_url(dates_column, writer, col):
    
    for date in dates_column:
        data_to_save = []
        date_obj = pd.to_datetime(date, errors='coerce')

        if not pd.isnull(date_obj):
            
            formatted_date = date_obj.strftime("%d/%m/%Y")
            print("Processing the curve PRE for day " + formatted_date)

            url = f"https://www2.bmf.com.br/pages/portal/bmfbovespa/boletim1/TxRef1.asp?Data={formatted_date}&Data1={date_obj.strftime('%Y%m%d')}&slcTaxa=PRE"

            response = requests.get(url)
            if response.status_code == 200:
                            
                soup = BeautifulSoup(response.content, 'html.parser')
                tabela = soup.find_all("table")[1].getText()
                tabela = tabela.replace('\r', '')
                data_lines = [line.strip() for line in tabela.split('\n') if line.strip()]

                # data_to_save.append([formatted_date, " "])
                for i in range(5, len(data_lines), 3):
                    if i + 1 < len(data_lines):
                        data_to_save.append(data_lines[i:i+2])

            else:
                content = f"Failed to fetch the page. Status code: {response.status_code}"

        df = pd.DataFrame(data_to_save, columns=['day', 'index'])
        
        # Convert data to appropriate types (int and float)
        df['day'] = df['day'].astype(int)
        df['index'] = df['index'].str.replace(',', '.').astype(float)   

        # Rename the columns and write to excel; increase the column count
        df.rename(columns={'day': formatted_date, 'index': ' '}, inplace=True)
        df.to_excel(writer, sheet_name='Sheet1', startcol=col, index=False)
        col += 2

    return data_to_save



def fetch_data_from_url(curve, dates_column, writer, col):
    
    for date in dates_column:
        data_to_save = []
        date_obj = pd.to_datetime(date, errors='coerce')

        if not pd.isnull(date_obj):
            
            formatted_date = date_obj.strftime("%d/%m/%Y")
            print("Processing the curve " + curve + " for day " + formatted_date)

            url = f"https://www2.bmf.com.br/pages/portal/bmfbovespa/boletim1/TxRef1.asp?Data={formatted_date}&Data1={date_obj.strftime('%Y%m%d')}&slcTaxa={curve}"

            response = requests.get(url)
            if response.status_code == 200:
                            
                soup = BeautifulSoup(response.content, 'html.parser')
                tabela = soup.find_all("table")[1].getText()
                tabela = tabela.replace('\r', '')
                data_lines = [line.strip() for line in tabela.split('\n') if line.strip()]

                # data_to_save.append([formatted_date, " "])
                for i in range(4, len(data_lines), 2):
                    if i + 1 < len(data_lines):
                        data_to_save.append(data_lines[i:i+2])
                        #print (data_lines[i:i+2])

            else:
                content = f"Failed to fetch the page. Status code: {response.status_code}"
 
        df = pd.DataFrame(data_to_save, columns=['day', 'index'])
        
        # Convert data to appropriate types (int and float)
        df['day'] = df['day'].astype(int)
        df['index'] = df['index'].str.replace(',', '.').astype(float)   

        # Rename the columns and write to excel; increase the column count
        df.rename(columns={'day': formatted_date, 'index': ' '}, inplace=True)
        df.to_excel(writer, sheet_name='Sheet1', startcol=col, index=False)
        col += 2

    return data_to_save


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    out_directory = 'Output'
    if not os.path.exists(out_directory):
    # Create the directory if it doesn't exist
        os.makedirs(out_directory)

    if file_path:
        df = pd.read_excel(file_path, header=None)
        dates_column = df.iloc[0:, 0]

        # Convert the dates_column to pandas datetime format (if not already)
        dates_column = pd.to_datetime(dates_column)

        # Sort the dates_column in ascending order (oldest dates first)
        sorted_dates_column = dates_column.sort_values()

        #list_of_curves = [ "DOC","DOL", "DIC", "PTX", "SLP", "LIB", "PRE"]
        list_of_curves = ["LIB"]
        for curve in list_of_curves:
            col = 0
            output_file = "./Output/" + curve + ".xlsx"

            # DI x Pré routine.. it must ignore the 3rd column
            if curve == "PRE":
                with pd.ExcelWriter(output_file, mode='w') as writer:
                    fetch_PRE_from_url(sorted_dates_column, writer, col)
            else:
                with pd.ExcelWriter(output_file, mode='w') as writer:
                    fetch_data_from_url(curve, sorted_dates_column, writer, col)
