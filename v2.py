import re, pymupdf, os, pandas as pd, tkinter as tk
from pdfminer.high_level import extract_text


#Function to extract values
def extract_values(filename):
    pdfdata = extract_text(filename)
    field_data_arr=pdfdata.split('\n')
    res={}
    for index, word in enumerate(fields): #For iterating over Column Names
        for phrase in field_data_arr: #for going through each line of pdf
            res['FileName'] = filename
            if word in phrase:
                if chr(8211).lower() in phrase.lower(): 
                    res[keys[index]] = re.sub(r'.*–', '', phrase.strip())
                elif ":" in phrase:
                    res[keys[index]] = re.sub(r'.*:', '', phrase.strip())
                else: 
                    res[keys[index]] = re.sub(r'.*-', '', phrase.strip())
    def extract_first_number(df):
        for col in df.columns: # for iterating over columns
            for cell in df[col]: # for checking each cell data
                if cell is not None:
                    match = re.search(r'\b(?!2024\b)(\d{1,3}(,\d{3})*|\d+)\b', str(cell)) # for checking amount
                    if match:
                        return match.group(0)
        return None
    doc = pymupdf.open(filename)
    page = doc[0]
    tabs = page.find_tables()
    try: 
        df = tabs[0].to_pandas() 
        res['Amount'] = extract_first_number(df)
    except Exception as e:
        pass
    return res

#Function to extract data
def extract():
    try: 
        
        global final_res, keys, fields, directory
        final_res=[]
        
        #Customizable as per needs
        directory = './PDFs'
        keys=['Invoice_Number', 'Invoice_Date', 'Full_Name', 'Pan', 'Phone_Number', 'Amount']
        fields=['Invoice Number', 'Invoice Date', 'Full Name', 'PAN CARD Number', 'Phone Number']
        
        #Regex Definition
        datepattern = re.compile(r'(0[1-9]|[12][0-9]|3[01])(\/|-)(0[1-9]|1[1,2])(\/|-)(19|20)\d{2}')
        panpattern = re.compile(r'[A-Z]{5}[0-9]{4}[A-Z]{1}')
        
        #Iteration of Files
        for filename in os.listdir(directory):
            if not filename.endswith(".pdf"):
                continue
            else: 
                f = os.path.join(directory, filename)
                if os.path.isfile(f):
                    final_res.append(extract_values(f))
        # Error Handling
        for index, i in enumerate(final_res): 
            final_res[index]['Error'] = ''
            if('Invoice_Number' not in final_res[index]):
                final_res[index]['Error'] = 'Invoice_Number'
            if('Invoice_Date' not in final_res[index]):
                a = final_res[index]['Error']
                a += '|Invoice_Date'
                final_res[index]['Error'] = a
            elif(not datepattern.search(final_res[index]['Invoice_Date'])):
                final_res[index]['Invoice_Date'] = ''
                a = final_res[index]['Error']
                a += '|Invoice_Date'
                final_res[index]['Error'] = a
            if('Full_Name' not in final_res[index]):
                a = final_res[index]['Error']
                a += '|Full_Name'
                final_res[index]['Error'] = a
            if('Pan' not in final_res[index]):
                a = final_res[index]['Error']
                a += '|Pan'
                final_res[index]['Error'] = a
            elif(not panpattern.search(final_res[index]['Pan'])):
                final_res[index]['Pan'] = ''
                a = final_res[index]['Error']
                a += '|Pan'
                final_res[index]['Error'] = a
            if('Phone_Number' not in final_res[index]):
                a = final_res[index]['Error']
                a += '|Phone_Number'
                final_res[index]['Error'] = a
            if('Amount' not in final_res[index] or final_res[index]['Amount']==None):
                a = final_res[index]['Error']
                a += '|Amount'
                final_res[index]['Error'] = a

        df = pd.DataFrame(final_res)
        df.to_excel('output.xlsx', index=False)
        result_label.config(text="Invoices parsed Successfully!", fg="#5cb85c", font=('Helvetica', 16), padx=20, pady=20)
    except Exception as e:
        print(e)
        result_label.config(text=f"Error: {e}.\n Contact Admin", fg="#d9534f", font=('Helvetica', 16), padx=20, pady=20)
        
# extract()

window = tk.Tk()
window.title("Invoice Extractor")
window.config(bg="#292b2c")
window.minsize(250, 250)
window.eval('tk::PlaceWindow . center')
label = tk.Label(window, text="Welcome to Invoice Extractor",fg="#5bc0de",  bg="#292b2c", padx=50, pady=50, font=('Helvetica', 16))
label.pack()
button = tk.Button(window, text="Click Me", command=extract)
button.pack()
result_label = tk.Label(window, bg="#292b2c", text="")
result_label.pack()
window.mainloop()
