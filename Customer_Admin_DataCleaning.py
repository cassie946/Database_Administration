import pandas as pd
import re
import numpy as np
df = pd.read_excel(r'C:\Users\Lisa Liu\Documents\softchoice_caseStudy\Database_AdministrationData.xlsx')
#EDA
print(df.head())
print(df.columns)
print(df.shape)
print(df.dtypes)
#Checking missing values in columns
print(df.isnull().sum())
print(df.isnull())

#Checking missing values in rows
null_data = df[df.isnull().any(axis=1)]
print(f"null_data: {null_data}")

#check if 'CONTACTID' follows naming protocol
ContactID_pattern = r'^CID\d{6}AAQ$'
CONTACTID_list= []
for row in range(len(df)):
    if re.match(ContactID_pattern, df['CONTACTID'].loc[row]) :
        print(f"{row}: Matches the pattern")
    else:
        print(f"{row}: Does not match the pattern")
        CONTACTID_list.append(df['CONTACTID'].loc[row])
print(CONTACTID_list)

#check if 'TRANSACTIONID' follows naming protocol
TRANSACTIONID_pattern =r'^ID\d{6}CDG$'
TRANSACTIONID_list = []
for row in range(len(df)):
    if re.match(TRANSACTIONID_pattern, df['TRANSACTIONID'].loc[row]) :
        print(f"{row}: Matches the pattern")
    else:
        print(f"{row}: Does not match the pattern")
        TRANSACTIONID_list.append(df['TRANSACTIONID'].loc[row])
print(TRANSACTIONID_list)

#SWAP CELL BETWEEN CONTACTID AND TRANSACTIONIC
pattern_CON = r'^CID\d{6}AAQ$'
pattern_TRA = r'^ID\d{6}CDG$'
for row in range(len(df)):
    if re.match(pattern_CON, df['CONTACTID'].loc[row]) :
        print(f"{row}: Matches the pattern")
    else:
        print(f"{row}: Does not match the pattern")
        ID_TRA = df.iloc[row]['CONTACTID']
        ID_CON = df.iloc[row]['TRANSACTIONID']
        if re.match(pattern_TRA, ID_TRA) :
            df.iloc[row]['TRANSACTIONID'] = ID_TRA
            print(f"{row} TRANSACTIONID value switch TO CONTACTID  : {df.iloc[row]['TRANSACTIONID']}")
        if re.match(pattern_CON, ID_CON) :
            df.iloc[row]['CONTACTID'] = ID_CON
            print(f"{row} CONTACTID value switch TO TRANSACTIONID : {df.iloc[row]['CONTACTID']}")
print(df.iloc[48][['CONTACTID','TRANSACTIONID']])

# check COUNTRY is either CA or US
COUNTRY_List =['US', 'CA']
CountryError_List=[]
index_list=[]
for row in range(len(df)):
    if df['COUNTRY'].loc[row] in COUNTRY_List:
        print (f"{row}: True")
    else:
        print(f"{row}: False")
        index_list.append(row)
        CountryError_List.append(df['COUNTRY'].loc[row])
print(index_list,COUNTRY_list)

#Check REGION follow the naming convention used in Organization Structure and Country.
Region_List =['CA EAST','US WEST','US EAST','CA WEST']
index_list=[]
RegionError_List=[]

for row in range(len(df)):
    if df['REGION'].loc[row] in Region_List:
        print (f"{row}: True")
    else:
        print(f"{row}: False")
        index_list.append(row)
        RegionError_List.append(df['REGION'].loc[row])
print(index_list,REGION_List,)

# replace the null or error value in REGION based on column 'SALES_LEADER'
ca_East =['Adem Osborne','Iram Luna','Shahzaib Mahoney']
ca_West =['Esa Walton','Zaid Nunez']
us_East =['Nicole Stamp','Hattie Beattie']
us_West =['Matas Walmsley','Atlanta Avila','Taliah Pope']
for row in range(len(df)):
    if df['REGION'].loc[row] in Region_List:
        print (f"{row}: True")
    else:
        print(f"{row}: False")
        if df['SALES_LEADER'].loc[row] in ca_East:
            df['REGION'].loc[row]='CA EAST'
            print(df['REGION'].loc[row])
        if df['SALES_LEADER'].loc[row] in ca_West:
            df['REGION'].loc[row]='CA WEST'
            print(df['REGION'].loc[row])
        if df['SALES_LEADER'].loc[row] in us_East:
            df['REGION'].loc[row]='US EAST'
            print(df['REGION'].loc[row])
        if df['SALES_LEADER'].loc[row] in us_West:
            df['REGION'].loc[row]='US WEST'
            print(df['REGION'].loc[row])
print(df['REGION'].isnull().sum())

# check if SALES_LEADER follow organization hierarchy
SL_list =['Iram Luna','Shahzaib Mahoney' , 'Zaid Nunez','Nicole Stamp','Atlanta Avila','Taliah Pope']
LeaderError_List=[]
for row in range(len(df)):
    if df['SALES_LEADER'].loc[row] in SL_list:
        print (f"{row}: True")
    else:
        print( f"{df['SALES_LEADER'].loc[row]}: sales leader row {row} does not follow organization hierarchy")
        index_list.append(row)
        LeaderError_List.append(df['SALES_LEADER'].loc[row])
print(index_list,LeaderError_List)

#check if SALES_LEADER contains leading space
sl_String=[]
for i in range(len(df)):
    if df['SALES_LEADER'].iloc[i].startswith(' '):
        sl_string.append(df['SALES_LEADER'].iloc[i])
        df['SALES_LEADER'].iloc[i]=df['SALES_LEADER'].iloc[i].strip()
        print(df['SALES_LEADER'].iloc[i])
    else:
        print(f"SALES_LEADER {i}: Does not have leading space")
print(sl_string)


# check if SALES_VICE_PRESIDENT follow organization hierachy
# swap name in SALES_LEAD and SALES_VICE_PRESIDENT
vp_List = ['Adem Osborne','Matas Walmsley','Hattie Beattie','Esa Walton']
sl_List =['Iram Luna','Shahzaib Mahoney' , 'Zaid Nunez','Nicole Stamp','Atlanta Avila','Taliah Pope']
for row in range(len(df)):
    if df['SALES_VICE_PRESIDENT'].loc[row] in vp_List:
        print(f"{row}: True")
    else:
        name1=df.iloc[row]['SALES_VICE_PRESIDENT']
        name2=df.iloc[row]['SALES_LEADER']
        if name1 in sl_List:
            df.iloc[row]['SALES_LEADER']=name1
            print(f"{row} value switch : {df.iloc[row]['SALES_LEADER']}")
        else:
            print(f"error found : {row}, SALES_VICE_PRESIDENT")
        if name2 in vp_List:
            df.iloc[row]['SALES_VICE_PRESIDENT']=name2
        else:
            print(f"error found : {row}, SALES_LEADER")

# check if SALES_VICE_PRESIDENT contains leading space
vp_String=[]
for i in range(len(df)):
    if df['SALES_VICE_PRESIDENT'].iloc[i].startswith(' '):
        vp_String.append(df['SALES_VICE_PRESIDENT'].iloc[i])
        df['SALES_VICE_PRESIDENT'].iloc[i] = df['SALES_VICE_PRESIDENT'].iloc[i].strip()
        print(df['SALES_VICE_PRESIDENT'].iloc[i])
    else:
        print(f"SALES_VICE_PRESIDENT {i}: Does not have leading space")
print(row_List,vp_String)
###there is no leading space in SALES_VICE_PRESIDENT

# drop missing value in FirstName after fulfill all the missing values
df = df.dropna(subset=['FirstName'])


#check if FirstName contains spaces

i_List=[]
FirstName_List=[]
for i in range(len(df)):
    if " " in df['FirstName'].iloc[i]:
        i_List.append(i)
        FirstName_List.append(df['FirstName'].iloc[i])
print(f'FirstName:{i_List},{FirstName_List}')


#check if LastName contains spaces
j_List=[]
LastName_List=[]
for j in range(len(df)):
    if " " in df['LastName'].iloc[j]:
        j_List.append(j)
        LastName_List.append(df['LastName'].iloc[j])
print(f'LastName:{j_List},{LastName_List}')

#check if Email contains spaces
k_List=[]
email_List = []
for k in range(len(df)):
    if " " in df['Email'].iloc[k]:
        k_List.append(k)
        email_List.append(df['Email'].iloc[k])
print(f'email:{k_List},{email_List}')
#[7, 51] ['Yassin .Jackson@fakeemailaddress.com', ' Sydney.Kramer@fakeemailaddress.com']

# Remove All Spaces in Email Using the replace() Method
for k in range(len(df)):
    if " " in df['Email'].iloc[k]:
        k_List.append(k)
        email_List.append(df['Email'].iloc[k])
        df['Email'].iloc[k]=df['Email'].iloc[k].replace(" ", "")
        print(df['Email'].iloc[k])

#save cleaned CUstome_Admin Data to excel file
df.to_excel("ADMINoutput.xlsx")