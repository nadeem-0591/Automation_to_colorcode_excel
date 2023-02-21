import pandas as pd
import re,os,operator,time
import numpy as np
from styleframe import utils,StyleFrame,Styler


df=pd.read_excel(r"C:\Users\DELL\Downloads\Df.xlsx")
sample=df.copy()
compare_val=list(df.loc[0])

df=df.loc[1:]       ## Remove the operator rows

df['A_4G_DlThp_Data_3UK_L1400']=df['A_4G_DlThp_Data_3UK_L1400']/1000 ### Convert to Mbps


### Determine the operator
def determineoperator(op):
    #print(op)
    if('>' in op):
        if('=' in op):
            value=operator.ge
        else:
            value=operator.gt
    elif '<' in op:
        if('=' in op):
            value=operator.le
        else:
            value=operator.lt
    else:
        value=""
    return value

### compare the operation
def compare(num,cvalue,opr):
    num=float(num)
    val=0
    #print(type(num),num)
    status=opr(cvalue,num)
    if (status):
        val=1
    #(num,type(num),cvalue,opr,val)
    return val
    
### Apply colors and modify cells

def formatexcel(df1,temp):
    df1.iloc[1:,-1:]=temp.iloc[1:,-1:]
    sf = StyleFrame(df1)
    for col_name in df1.columns[:-2]:
        #### Apply Red colour to invalid cells 
        sf.apply_style_by_indexes(
            sf[temp[col_name]==1],cols_to_style=col_name,
            styler_obj=Styler(bg_color='red',font_size=8,
            number_format=utils.number_formats.general_float,
            wrap_text=False))

        #### Apply Green colour to valid cells
        sf.apply_style_by_indexes(
            sf[temp[col_name]==0],cols_to_style=col_name,
            styler_obj=Styler(bg_color="#228B22",font_size=8,
            number_format=utils.number_formats.general_float,
            wrap_text=False))
        #### Apply Yellow colour to Operator cells
        sf.apply_style_by_indexes(
            sf[temp[col_name]==0.0505],cols_to_style=col_name,
            styler_obj=Styler(bg_color="yellow",font_size=8,
            wrap_text=False))
        
    header_style = Styler(bg_color="#c4c4ff",text_rotation=90,
    	font_size=8,wrap_text=False)
    sf.apply_headers_style(styler_obj=header_style)
    sf.set_column_width(columns=sf.columns, width=10)
    sf.to_excel("Output "+str(int(time.time()))+".xlsx").save()
    return True



final=pd.DataFrame()
col_pos=0
for opr,col in zip(compare_val,df.columns):
    #print(col)
    digit=""
    opr=str(opr)
    optr=determineoperator(opr)
    if(optr):
        try:
            digit=int(" ".join(re.findall('([\d.]+)',opr)))
            if('-'in opr):
            	digit=operator.sub(0,digit)
            data1=df[col].apply(compare,cvalue=digit,opr=optr)
            a=[0.0505]
            a.extend(list(data1))
            #print("A list",a,data1)
            final.insert(col_pos,col,a)
            #print(final)
            #display(pd.DataFrame({col:list(data1)}))
            col_pos+=1
            #print("Data",optr,digit)
            #print(data1)
        except Exception as e:
        	a=[0.0505]
        	a.extend(np.zeros(df.shape[1]-1))
        	final.insert(col_pos,col,a)
        	col_pos+=1
        	print("Exception Occured",e)
        	continue
    else:
        print("Escaping",col)
        a=[0.0505]
        a.extend(np.zeros(df.shape[1]-1))
        final.insert(col_pos,col,a)
        col_pos+=1
    #print(digit,opr,col,len(df[col]))

#print(final)

kpi_count=[]
for index,rowdata in final.iterrows():
    kpi_count.append(int(rowdata.values.sum()))

print("KPI_COUNT",kpi_count,col)
final[col]=kpi_count
final.to_excel('final.xlsx',index=False)
#print(final)

formatexcel(sample,final)



### https://github.com/DeepSpace2/StyleFrame