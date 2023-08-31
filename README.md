# Find-duplicates-base-on-groups
import pandas as pd
import numpy as np
# Because the name is too long, I rename several names of column
#A=S A1a Street Number From
#B=S A1b Street Number To
#C=S A1c Street Number By
#D=S A1g Street Type Prefix
#E=S A1i Street Name
#F=S A1k Street Type
#G=CSD Code (S A1m CSD Name) (Municipality)

#set variables.
def mark_rows(group):
    marked_indices_i = []
    marked_indices_j = []
    marked_indices_i2 = []
    marked_indices_j2 = []

    for index_i, row_i in group.iterrows():
        if row_i['B'] > 0 and row_i['A'] == 0:#(first situation:)
            for index_j, row_j in group.iterrows():
                if index_j != index_i and row_j['A'] > 0 and row_i['B'] >= row_j['A'] and row_i['B'] <= row_j['B']:
                    if row_j['C'] > 1 and (row_i['B']-row_j['A'])%row_j['C']==0:
                        marked_indices_i.extend([index_i])
                        marked_indices_j.extend([index_j])
                    elif row_j['C'] == 1:
                        marked_indices_i.extend([index_i])
                        marked_indices_j.extend([index_j])                    
        elif row_i['B'] > row_i['A'] > 0:#(second situation:)
            for index_j, row_j in group.iterrows():
                if index_j!= index_i and row_j['B'] > row_j['A'] > 0 and row_i['A']>=row_j['A'] and row_i['A']<=row_j['B']:
                    if row_j['C']>1 and row_i['C']>1 and (row_i['A']-row_j['A'])%min(row_j['C'],row_i['C'])==0 or (row_j['B']-row_i['A'])%min(row_j['C'],row_i['C'])==0:
                        marked_indices_i2.extend([index_i])
                        marked_indices_j2.extend([index_j])
                    elif min(row_i['C'],row_j['C']) == 1:
                        marked_indices_i2.extend([index_i])
                        marked_indices_j2.extend([index_j])
                elif index_j!= index_i and row_j['B'] > row_j['A'] > 0 and row_i['B']>=row_j['A'] and row_i['B']<=row_j['B']:
                    if row_j['C']>1 and row_i['C']>1 and (row_i['B']-row_j['A'])%min(row_j['C'],row_i['C'])==0 or (row_j['B']-row_i['B'])%min(row_j['C'],row_i['C'])==0:
                        marked_indices_i2.extend([index_i])
                        marked_indices_j2.extend([index_j])
                    elif min(row_i['C'],row_j['C']) == 1:
                        marked_indices_i2.extend([index_i])
                        marked_indices_j2.extend([index_j])
            

    group['Marked'] = group.index.isin(marked_indices_i)
    group['Range'] = group.index.isin(marked_indices_j)
    group['Marked2'] = group.index.isin(marked_indices_i2)
    group['Range2'] = group.index.isin(marked_indices_j2)
    return group
#Read excel file.
#When you analyze different files, remember to change the name of the file. STC_MB_CHECK.xlsx is the name of the file
File_check=pd.read_excel('STC_MB_CHECK.xlsx')
#Change the name of columns for the program I created before
New_column_names={
    'S A1a Street Number From':'A',
    'S A1b Street Number To':'B',
    'S A1c Street Number By':'C',
    'S A1g Street Type Prefix':'D',
    'CSD Code (S A1m CSD Name) (Municipality)':'G'
}
File_check=File_check.rename(columns=New_column_names)

#Change the type of columns before rewriting them
#Because there are various types of cells, the program will not work.For example: 4(numeric),4(text) will not be regarded as the same.
File_check['S A1i Street Name']=File_check['S A1i Street Name'].astype(str)
File_check['S A1k Street Type']=File_check['S A1k Street Type'].astype(str)

#Formate column E,F.Because system is very sensitive,for example: Road, road,ROAD,ROad will not be regarded as the same.
File_check['E']=File_check['S A1i Street Name'].str.upper()
File_check['F']=File_check['S A1k Street Type'].str.capitalize()
File_check['D']=File_check['D'].str.capitalize()

#Add the modified column(E) after the original 'S A1i Street Name' column
E_index=File_check.columns.get_loc('S A1i Street Name')
File_check.insert(E_index+1,'E',File_check.pop('E'))
#Add the modified column(F) after the original 'S A1i Street Type' column
F_index=File_check.columns.get_loc('S A1k Street Type')
File_check.insert(F_index+1,'F',File_check.pop('F'))

#Replace French letters(ç, é, â, ê, î, ô, û, à, è, ù, ë, ï, ü) with English letters.(If you find more common cases, I can also add them in this formula)
#Remove whitespace and '
replace_dict={'ç':'c', 'é':'e', 'â':'a', 'ê':'e', 'î':'i', 'ô':'o', 'û':'u', 'à':'a', 'è':'e', 'ù':'u', 'ë':'e', 'ï':'i', 'ü':'u',' ':'',"'":''}
File_check['E'] = File_check['E'].astype(str).replace(replace_dict, regex=True)

#Remove special characters(if you find more common cases, I can also add them in this formula)
File_check['E'] = File_check['E'].astype(str).replace(r'[#-,/+*_() :;<=>?&]', '', regex=True)

#Fill all empty cells in ?. My or our systm is too old to deal with the rows that contain empty cells. 
#I choose ?, because I am sure that ? is not an information, then I will not change, delete, or add any information
File_Na=File_check.fillna('?')

#We need to group data by D,E,F,G,because we find duplicates base on groups
grouped_File=File_Na.groupby(['D', 'E', 'F', 'G'])

#Using the function I have already created to get the result we want 
result_File= grouped_File.apply(mark_rows)

# Find Another type of Duplicates(The easist situation), add 'is_duplicate' column to show the results
result_File.reset_index(drop=True, inplace=True)
result_File['is_duplicate'] = File_check.duplicated(subset=['B','D', 'E', 'F', 'G'], keep=False)

#Change the name of columns back
Back_column_names={
    'A':'S A1a Street Number From',
    'B':'S A1b Street Number To',
    'C':'S A1c Street Number By',
    'D':'S A1g Street Type Prefix',
    'E':'S A1i Street Name_Modified',
    'F':'S A1k Street Type_Modified',
    'G':'CSD Code (S A1m CSD Name) (Municipality)'
}
result_File=result_File.rename(columns=Back_column_names)
#Reset index and save to a new Excel file and Export Excel file
result_File.replace('?', np.NaN).to_excel("Duplicate_MB.xlsx",index=False)
