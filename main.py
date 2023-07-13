print('in work for tom main')

#imports
import pandas as pd
import os
import xlsxwriter

print(pd.__version__)
print(xlsxwriter.__version__)

#for good messure
sheet_dict = {
    'downlode': 'הורדות',
    'RBT_downlode': 'RBT - הורדות',
    'RBT_DMS': 'RBT - דמ"ש',
    'SUM': 'RBT - דמ"ש'
    
}

#step 1 - get files
def get_file_names():
    file_list = []
    file_names = os.listdir('before')
    for name in file_names:
        if name.split('.')[1] == 'xlsx':
            file_list.append(name.split('.')[0])
    return file_list


def get_file(file_name, sheet_name):  
    return pd.read_excel(f'before/{file_name}.xlsx', sheet_name=sheet_name, engine='openpyxl')

     

def get_files(file_dict, file_name):
    for key, sheet in sheet_dict.items():
        df = get_file(file_name, sheet)
        file_dict[key] = df
    return file_dict

def get_collomn(file_dict, sheet_name, collumn_name):
    #has to be after cleaning empty rows
    df = file_dict[sheet_name]
    new_collumn_name = None
    for col in df.columns:
        if df.loc[0, col] == collumn_name:
            new_collumn_name = col
            break
    if new_collumn_name:
        return new_collumn_name
    else:
        print(f'no collomn named {collumn_name}')

#step 2 - delete 'isrc' collom in הורדיות folder

def delete_collom(file_dict, sheet_name ,collumn_name):
    collumn_name = get_collomn(file_dict, sheet_name ,collumn_name)
    if collumn_name in file_dict[sheet_name].columns:
        file_dict[sheet_name] = file_dict[sheet_name].drop(collumn_name, axis=1)

#step 3 - in folder הורדות RBT  make sure the "other" and "content name" celles has up to 25 chars, if their are more simply cut the rest

def check_char_limit(file_dict, sheet_name ,collumn_name, limit):
    collumn_name = get_collomn(file_dict, sheet_name ,collumn_name)
    for index, row in file_dict[sheet_name].iterrows():
        cell_value = row[collumn_name]
        if len(str(cell_value)) > limit:
             file_dict[sheet_name].at[index, collumn_name] = cell_value[:limit]

#step 4 - step 3 on folder "דמש RBT"

#step 5 - in folder "דמש RBT" add collom "charge to the provider" (like הורדות RBT) 
#the value is active_item * מחיר לפריט לספק 
def create_collom(file_dict, sheet_name):
    collumn_name_Author = get_collomn(file_dict, sheet_name ,'Author')
    index = file_dict[sheet_name][file_dict[sheet_name][collumn_name_Author] == 'מחיר לפריט לספק (₪)'].index.item()
    collumn_name_Content_Name = get_collomn(file_dict, sheet_name ,'Content Name')
    price_per_item_for_provider = file_dict[sheet_name].loc[index, collumn_name_Content_Name]
    
    collumn_name_Active_Items = get_collomn(file_dict, sheet_name ,'Active Items')
    
    file_dict[sheet_name]['Charge to the provider'] = None
    file_dict[sheet_name].loc[0, 'Charge to the provider'] = 'Charge to the provider'
    print(type(price_per_item_for_provider))
    print(price_per_item_for_provider)
    
   

    rounded_value = int(round(float(str(price_per_item_for_provider))*100, 0))
    print(rounded_value)
    #rounded_value *= 100
    print(rounded_value)
    
    file_dict[sheet_name]['Charge to the provider'][1:] = '=' + (file_dict[sheet_name][collumn_name_Active_Items] * rounded_value)/100   
    
    
    

#step 6 - clear empty lines

def drop_empty(file_dict):
    for df in file_dict.values():
        df.dropna(how='all', inplace=True)
        df.reset_index(drop=True, inplace=True)

#step 7 - make sure last row starts with "Total" (with capital T)

def check_for_T(file_dict, sheet_name):
    last_row_index = file_dict[sheet_name].index[-1]
    file_dict[sheet_name].loc[last_row_index, file_dict[sheet_name].columns[0]] = 'Total'

#step 8 - save with same name as a xls file (insted of exls)

def save_as_xlsx(file_name, file_dict):
    file_path = f'after/{file_name}.xlsx'
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for sheet_name, df in file_dict.items():
            df.to_excel(writer, sheet_name=sheet_name)




#connect all in func main

def main():
    file_dict = {}
    print('in main')
    file_names = get_file_names()
    for file_name in file_names:
        
        print(file_name)
        
        #step 1
        file_dict = get_files(file_dict, file_name)
        
        #print(file_dict)
        
         #step 6
        drop_empty(file_dict)
        
        #step 2
        delete_collom(file_dict,'downlode', 'isrc')
        
        #step 3
        check_char_limit(file_dict, 'RBT_downlode', 'Author', 25)
        check_char_limit(file_dict, 'RBT_downlode', 'Content Name', 25)
        
        #step 4
        check_char_limit(file_dict, 'RBT_DMS', 'Author', 25)
        check_char_limit(file_dict, 'RBT_DMS', 'Content Name', 25)

        #step 5
        #create_collom(file_dict, 'RBT_DMS')
        
        
        #step 7
        check_for_T(file_dict, 'RBT_downlode')
        check_for_T(file_dict, 'RBT_DMS')
        
        #step 8
        #save_as_xlsx(file_name, file_dict)
        print(file_dict)
        
        break


main()



