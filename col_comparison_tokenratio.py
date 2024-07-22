from fuzzywuzzy import process, fuzz
import pandas as pd

dict={}
df =pd.read_excel("C:/Users/Gaurav/Desktop/spell.xlsx",sheet_name=["Sheet1","Sheet2"])
sheet1_df=df["Sheet1"]
sheet2_df=df["Sheet2"]

for index , row in sheet2_df.iterrows():
    dict[row["Company"]]=row["id"]

for index , row in sheet1_df.iterrows():
    
    max_match=0
    for sheet2_comapny in dict.keys():
        
        match=fuzz.token_sort_ratio(row["Company"],sheet2_comapny)
        if max_match<match:
            max_match=fuzz.token_sort_ratio(row["Company"],sheet2_comapny)
            match_comp=sheet2_comapny
            
    sheet1_df.at[index,"id"]=dict[match_comp]
    
with pd.ExcelWriter("C:/Users/Gaurav/Desktop/spell.xlsx", engine='openpyxl') as writer:
    sheet1_df.to_excel(writer, sheet_name="Sheet1", index=False)
    sheet2_df.to_excel(writer, sheet_name="Sheet2", index=False)