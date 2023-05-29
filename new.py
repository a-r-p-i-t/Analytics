import pandas as pd
df=pd.read_excel("C:\\Users\\Arpit Mohanty\\analytics\\output4.xlsx",sheet_name="Trends-Bandra-Mumbai")
def loop_through(date,sheet_name,brand):
    df.Time=pd.to_datetime(df.Time)
    fullness_sum,day_count=0,0


    for ind in df.index:
        if df["Time"][ind]==date and df["Brand Name"][ind]==brand:
            day_count+=1
            fullness_sum+=df["Fullness"][ind]
        else:
            fullness_sum=0
            day_count=1
    return fullness_sum,day_count

stores=["Trends-Bandra-Mumbai","Trends-Koram Mall-Mumbai","Trends-Imperial Mall,PCMC-Pune","Trends-Amanora-Pune","Trends-Market City-Pune","Trends-RIL Mall,Elpro-Pune","Trends-Seawoods-NaviMumbai","Trends-DVictoria-Mumbai"]

for i in stores:
    for j in df["Time"].unique():
        print(loop_through(j,i,"Flormar"))
