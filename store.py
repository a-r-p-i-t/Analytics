import pandas as pd
import cv2
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import datetime as dt
import calendar
import re
import os
from os import listdir
import glob
import xlwings as xw


wb=xw.Book()
sht=wb.sheets[0]
sht.name="Analysis of stores"
folder_dir="C:\\Users\\Arpit Mohanty\\analytics"



def file_extract(sheet_name):
    return pd.read_excel("C:\\Users\\Arpit Mohanty\\analytics\\output4.xlsx",sheet_name=sheet_name)



def avg_fullness(date,sheet_name):
    df=file_extract(sheet_name)
    df.Time=pd.to_datetime(df.Time)
    fullness_sum,day_count=0,0
    for ind in df.index:
        if df["Time"][ind]==date:
            day_count+=1
            fullness_sum+=df["Fullness"][ind]
    
    avg_fullness_=fullness_sum/day_count
    return avg_fullness_




def critical_pts(sheet_name):
    df=file_extract(sheet_name)
    global strt_date,end_date,sht

    df.Time=pd.to_datetime(df.Time)
    dt_ls=[]
    dt_lss=[]
    for i in df.Time:
        dt_ls.append(i.date())
    dt_set=set(dt_ls)
    for i in dt_ls:
        if i in dt_set:
            dt_lss.append(i)
            dt_set.remove(i)




    fullness_list=[]
    for dates in dt_lss:
        fullness_list.append(avg_fullness(dates,sheet_name))
    
    all_obj=zip(dt_lss,fullness_list)
    all_ls=list(all_obj)


    
    strt_date=pd.to_datetime(strt_date)
    end_date=pd.to_datetime(end_date)
    strt_datef=strt_date.date()
    strt_datefr=strt_date.date()
    end_datef=end_date.date()
    delta = dt.timedelta(days=1)
    cust_dates=[]
    m_cust_dates=[]
    while (strt_datef <= end_datef):
        cust_dates.append(strt_datef)
        strt_datef += delta

    cust_full=[]
    for i in range(len(all_ls)):
        if all_ls[i][0] in cust_dates:
            cust_full.append(all_ls[i][1])
            m_cust_dates.append(all_ls[i][0])
        


    s=0
    for i in fullness_list:
        s+=i
    avg_fullnesss=s/len(fullness_list)
    avg_fullnesssr=round(avg_fullnesss,2)

    


    critical_ls=[cust_full[0]]
    for i in range(len(cust_full)-1):
        if abs(cust_full[i+1]-cust_full[i]) >5 :
            critical_ls.append(cust_full[i+1])





    critical_rp=[critical_ls[0]]
    critical_cns=[critical_ls[0]]
    critical_ls_copy=critical_ls.copy()
    markerfacecolor=["r"]
    full_dev=["Ref"]
    rep_rate_ls=[]
    cns_rate_ls=[]

    
    diff_rp=[]
    i=0
    diff_cns=[]
    while i<(len(cust_full)-1):
        if cust_full[i+1]-cust_full[i]>5:
            if cust_full[i+1] in critical_ls and abs(cust_full[i+1]-cust_full[i]) >5 :
                diff=cust_full[i+1]-cust_full[i]
                diff_round=round(diff,2)
                rep_rate_ls.append(diff_round)
                diff_rp.append(str(m_cust_dates[i+1]-m_cust_dates[i])) 
                # date_rp
                full_dev.append(str(diff_round)+"% gain")  
                markerfacecolor.append("g")
                critical_rp.append(cust_full[i+1])
        
            

        if cust_full[i+1]-cust_full[i]<5:
            if cust_full[i+1] in critical_ls and abs(cust_full[i+1]-cust_full[i]) >5:
                diff=cust_full[i]-cust_full[i+1]
                diff_round=round(diff,2)
                cns_rate_ls.append(diff_round)
                diff_cns.append(str(m_cust_dates[i+1]-m_cust_dates[i]))
                full_dev.append(str(diff_round)+"% drop")
                markerfacecolor.append("r")
                critical_cns.append(cust_full[i+1])
            
        i+=1


    sum_rep=0
    sum_cns=0
    for i in rep_rate_ls:
        sum_rep+=i
    for j in cns_rate_ls:
        sum_cns+=j
    
    

    obj=zip(cust_full,m_cust_dates)
    comp_ls=(list(obj))

    date_rp=[]
    
    sub_unique_ls=[]
    obj_1=zip(m_cust_dates,cust_full)
    neutral_ls=list(obj_1)
    for y in range(len(neutral_ls)):
        if neutral_ls[y][1] in critical_ls_copy :
            sub_unique_ls.append(neutral_ls[y][0])
            critical_ls_copy.remove(neutral_ls[y][1])


    
        
    for i in range(len(comp_ls)):
        if comp_ls[i][0] in critical_rp:
            date_rp.append(comp_ls[i][1])
    
    tuned_diff_rp=[]
    diff_rp_1=[]
    for i in range(len(date_rp)-1):
        diff_rp_1.append(str(date_rp[i+1]-date_rp[i]))

    for m in diff_rp_1:
        z=m.split(",")
        tuned_diff_rp.append(z)
    tuned_1_diff_rp=[]
    for i in range(len(tuned_diff_rp)):
        tuned_1_diff_rp.append(tuned_diff_rp[i][0])
    tuned_2_diff_rp=[]
    for i in tuned_1_diff_rp:
        q=re.sub(r'[^0-9]', '', i)
        tuned_2_diff_rp.append(q)
    sum=0
    for i in tuned_2_diff_rp:
        sum+=int(i)
    if len(critical_rp)!=1:
        mttr=sum/(len(rep_rate_ls))
        mttrr=round(mttr,2)
    else:
        mttrr="No replenishment"
    if len(rep_rate_ls)!=0:
        rep_rate=sum_rep/len(rep_rate_ls)
    if len(cns_rate_ls)!=0:
        cns_rate=sum_cns/len(cns_rate_ls)
    if len(rep_rate_ls)!=0:
        rep_rater=round(rep_rate,2)
    else:
        rep_rater="No replenishment"


    if len(cns_rate_ls)!=0:
        cns_rater=round(cns_rate,2)
    else:
        cns_rater="No Consumption"
    
       



    date_cns=[]
    diff_cns_1=[]
    
        
    for i in range(len(comp_ls)):
        if comp_ls[i][0] in critical_cns:
            date_cns.append(comp_ls[i][1])
    for i in range(len(date_cns)-1):
        diff_cns_1.append(str(date_cns[i+1]-date_cns[i]))
    
    tuned_diff_cns=[]

    for m in diff_cns_1:
        z1=m.split(",")
        tuned_diff_cns.append(z1)
    tuned_1_diff_cns=[]
    for i in range(len(tuned_diff_cns)):
        tuned_1_diff_cns.append(tuned_diff_cns[i][0])
    tuned_2_diff_cns=[]
    for i in tuned_1_diff_cns:
        q1=re.sub(r'[^0-9]', '', i)
        tuned_2_diff_cns.append(q1)
    sum1=0
    for i in tuned_2_diff_cns:
        sum1+=int(i)
    if len(critical_cns)!=1:
        mttc=sum1/(len(cns_rate_ls))
        mttcr=round(mttc,2)
    else:
        mttcr="No Consumption"

    day_rp=[]
    for i in date_rp:
        marker=i.weekday()
        day_rp.append(calendar.day_name[marker])
    ob_2=zip(day_rp,rep_rate_ls)
    obj_ls=list(ob_2)
    day_dic={}
    count,count1,count2,count3,count4,count5,count6=0,0,0,0,0,0,0
    for i in range(len(obj_ls)):
        if obj_ls[i][0]=="Sunday":
            count+=obj_ls[i][1]
        if obj_ls[i][0]=="Monday":
            count1+=obj_ls[i][1]
        if obj_ls[i][0]=="Tuesday":
            count2+=obj_ls[i][1]
        if obj_ls[i][0]=="Wednesday":
            count3+=obj_ls[i][1]
        if obj_ls[i][0]=="Thursday":
            count4+=obj_ls[i][1]
        if obj_ls[i][0]=="Friday":
            count5+=obj_ls[i][1]
        if obj_ls[i][0]=="Saturday":
            count6+=obj_ls[i][1]
    

    day_dic["Sunday"]=count
    day_dic["Monday"]=count1
    day_dic["Tuesday"]=count2
    day_dic["Wednesday"]=count3
    day_dic["Thursday"]=count4
    day_dic["Friday"]=count5
    day_dic["Saturday"]=count6

    labels = [i for i in day_dic.keys() if day_dic[i]!=0]
    sizes = [i for i in day_dic.values() if i!=0]
    rep_ls=[]
    sizesr=[]
    for i in sizes:
        sizesr.append(round(i,2))
    for i,j in zip(sizesr,labels):
        rep_ls.append(j+"-"+str(i))

    insert_handing(sht.range("A2"),f"Analysis Charts from {strt_datefr} to {end_datef}")
   
    

    
    fig=plt.figure()
    plt.pie(sizes, labels=rep_ls)
    plt.title(f"Analysis of Rep Rates of {sheet_name}")
    plt.axis('equal')
    plt.savefig(f"Rep Rates of {sheet_name}")
    # plt.show()

    sht.pictures.add(
        fig,name=f"Matplotlib_{sheet_name}",
        update=False,
        left=sht.range("A"+str(k)).left,
        top=sht.range("A"+str(k)).top,
        height=200,
        width=250,
    )



    day_cns=[]
    for i in date_cns:
        marker1=i.weekday()
        day_cns.append(calendar.day_name[marker1])
    ob_3=zip(day_cns,cns_rate_ls)
    obj_ls_1=list(ob_3)
    day_dic1={}
    count_0,count_1,count_2,count_3,count_4,count_5,count_6=0,0,0,0,0,0,0
    for i in range(len(obj_ls_1)):
        if obj_ls_1[i][0]=="Sunday":
            count_0+=obj_ls_1[i][1]
        if obj_ls_1[i][0]=="Monday":
            count_1+=obj_ls_1[i][1]
        if obj_ls_1[i][0]=="Tuesday":
            count_2+=obj_ls_1[i][1]
        if obj_ls_1[i][0]=="Wednesday":
            count_3+=obj_ls_1[i][1]
        if obj_ls_1[i][0]=="Thursday":
            count_4+=obj_ls_1[i][1]
        if obj_ls_1[i][0]=="Friday":
            count_5+=obj_ls_1[i][1]
        if obj_ls_1[i][0]=="Saturday":
            count_6+=obj_ls_1[i][1]
    

    day_dic1["Sunday"]=count_0
    day_dic1["Monday"]=count_1
    day_dic1["Tuesday"]=count_2
    day_dic1["Wednesday"]=count_3
    day_dic1["Thursday"]=count_4
    day_dic1["Friday"]=count_5
    day_dic1["Saturday"]=count_6

    sizes1 = [i for i in day_dic1.values() if i!=0]
    labels1 = [i for i in day_dic1.keys() if day_dic1[i]!=0]
    cns_ls=[]
    sizes1r=[]
    for i in sizes1:
        sizes1r.append(round(i,2))
    for i,j in zip(sizes1r,labels1):
        cns_ls.append(j+"-"+str(i))

    
    fig=plt.figure()
    plt.pie(sizes1, labels=cns_ls)
    plt.title(f"Analysis of Cns Rates of {sheet_name}")
    plt.axis('equal')
    plt.savefig(f"Cns Rates of {sheet_name}")
    # plt.show()



    sht.pictures.add(
        fig,name=f"pie_cns_{sheet_name}",
        update=False,
        left=sht.range("G"+str(k)).left,
        top=sht.range("D"+str(k)).top,
        height=200,
        width=250,
    )





    rep_count=[day_dic["Sunday"],day_dic["Monday"],day_dic["Tuesday"],day_dic["Wednesday"],day_dic["Thursday"],day_dic["Friday"],day_dic["Saturday"]]
    cns_count=[day_dic1["Sunday"],day_dic1["Monday"],day_dic1["Tuesday"],day_dic1["Wednesday"],day_dic1["Thursday"],day_dic1["Friday"],day_dic1["Saturday"]]
    days=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
    fig=plt.figure()
    x_axis=np.arange(len(days))
    plt.bar(x_axis - 0.2, rep_count, 0.4, label = 'Avg Replenishment Rate')
    plt.bar(x_axis + 0.2, cns_count, 0.4, label = 'Avg Consumption Rate')
    
    plt.xticks(x_axis, days)
    plt.xlabel("Days")
    plt.ylabel("Vol. Rates")
    plt.title(f"Analysis of Vol. Rates of {sheet_name}")
    plt.legend()
    plt.savefig(f"Bar representation of {sheet_name}")
    # plt.show()



    sht.pictures.add(
        fig,name=f"bar_representation_{sheet_name}",
        update=False,
        left=sht.range("M"+str(k)).left,
        top=sht.range("M"+str(k)).top,
        height=200,
        width=250,
    )

    
    fig=plt.figure()
    plt.scatter(sub_unique_ls,critical_ls,c=markerfacecolor)
    plt.plot(m_cust_dates,cust_full,label=sheet_name,color="y")
    plt.plot(sub_unique_ls,critical_ls)
    plt.xticks(m_cust_dates,m_cust_dates)
    for i, txt in enumerate(full_dev):
        plt.annotate(txt, (sub_unique_ls[i], critical_ls[i]),rotation=-10)
    
    plt.xlabel("Dates")
    plt.xticks(rotation = 90)
    plt.ylabel("Fullness")
    plt.legend(loc="upper left")
    plt.title("Analysis of fullness")
    
    plt.savefig(f"Criticals of {sheet_name}")
    # plt.show()



    sht.pictures.add(
        fig,name=f"Graphical Rep_{sheet_name}",
        update=False,
        left=sht.range("S"+str(k)).left,
        top=sht.range("S"+str(k)).top,
        height=200,
        width=250,
    )
    return mttrr,mttcr,rep_rater,cns_rater,avg_fullnesssr
    






def insert_handing(rng,text):
    rng.value=text
    rng.font.bold=True
    rng.font.size=24
    rng.font.color=(0,0,139)






sheets=input("Enter the stores name:").split(";")
strt_date=input("Enter the start date:")
end_date=input("Enter the end date:")
df_new=pd.DataFrame(columns=["Trends Beauty Stores","MTTR(Days)","MTTC(Days)","Replenishment Rate(Vol)","Consumption Rate(Vol)","Avg.Shelf-Fullness"])
k=4
for i in range(len(sheets)):
    try:    
        a,b,c,d,e=critical_pts(sheets[i])
        df_new.loc[i]=[sheets[i],a,b,c,d,str(e)+"%"]
        k+=20
        
    except:
        print(f"Data is not available for the selectes dates in store{i+1} or selected dates are not feasible.")

    
sht["J159"].options(pd.DataFrame, header=1, index=True, expand='table').value = df_new
sht["J159"].expand("down").api.Font.Bold = True
sht["J159"].expand("right").api.Font.Bold = True
sht["J159"].expand("right").api.Borders.Weight = 2
sht["J159"].expand("down").api.Borders.Weight = 2

wb.save("chart.xlsx")
print(df_new)




    




    

    

    






    









