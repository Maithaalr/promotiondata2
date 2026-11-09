import re
from datetime import date, datetime
from io import BytesIO
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='تحليل الترقيات الموحد', page_icon='📈', layout='wide')
st.title('📈 لوحة تحليل الترقيات الموحدة | ERP + Cloud')
st.caption('دمج سجلات الترقيات لجهة واحدة، إزالة التكرار، تحليل آخر ترقية وأنواع الترقيات، وتصدير Excel.')

HIDDEN = re.compile(r'[\u200e\u200f\u202a-\u202e\u2066-\u2069\ufeff]')
def clean(v):
    if pd.isna(v): return ''
    return re.sub(r'\s+', ' ', HIDDEN.sub('', str(v)).replace('\xa0',' ')).strip()

def norm_id(v):
    s=clean(v)
    return re.sub(r'\.0$', '', s) if s.lower() not in ('nan','none','nat') else ''

def parse_date(v):
    if pd.isna(v) or clean(v)=='': return pd.NaT
    if isinstance(v,(pd.Timestamp,datetime,date)): return pd.Timestamp(v).normalize()
    if isinstance(v,(int,float,np.integer,np.floating)) and 20000 <= float(v) <= 80000:
        return (pd.Timestamp('1899-12-30')+pd.to_timedelta(float(v),unit='D')).normalize()
    s=clean(v).translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹','01234567890123456789'))
    s=re.sub(r'\s+\d{1,2}:\d{2}(:\d{2})?.*$', '', s)
    for fmt in ('%d/%m/%Y','%d-%m-%Y','%d.%m.%Y','%d/%m/%y','%d-%m-%y','%d.%m.%y','%Y-%m-%d','%Y/%m/%d'):
        try: return pd.Timestamp(datetime.strptime(s,fmt)).normalize()
        except ValueError: pass
    return pd.NaT

def read_file(f):
    if f.name.lower().endswith('.csv'):
        raw=f.getvalue()
        for enc in ('utf-8-sig','cp1256','latin1'):
            try: return pd.read_csv(BytesIO(raw),encoding=enc,dtype=str)
            except UnicodeDecodeError: pass
        raise ValueError('تعذر قراءة CSV')
    return pd.read_excel(f, dtype=object)

def options(df): return list(df.columns)
def guess(cols, names):
    for n in names:
        for c in cols:
            if clean(c).replace(' ','')==n.replace(' ',''): return c
    return cols[0]

files=st.columns(2)
with files[0]: erp=st.file_uploader('📂 ملف ERP',type=['xlsx','xls','csv'],key='erp')
with files[1]: cloud=st.file_uploader('☁️ ملف Cloud',type=['xlsx','xls','csv'],key='cloud')
if not erp or not cloud:
    st.info('ارفعي الملفين لبدء التحليل. يتم الدمج على الرقم الوظيفي + تاريخ الترقية بعد تحويله إلى تاريخ فعلي.')
    st.stop()
try:
    a,b=read_file(erp),read_file(cloud)
    for d in (a,b): d.columns=[clean(c) for c in d.columns]
except Exception as e:
    st.error(f'تعذر قراءة الملفين: {e}'); st.stop()

st.subheader('1. تحديد الأعمدة')
with st.expander('اختيار أعمدة كل ملف (تلقائي ويمكن تعديلها)',expanded=True):
    c1,c2=st.columns(2)
    selections=[]
    for label,df,container in [('ERP',a,c1),('Cloud',b,c2)]:
        with container:
            st.markdown(f'**{label}**')
            cols=options(df)
            idcol=st.selectbox('الرقم الوظيفي',cols,index=cols.index(guess(cols,['الرقم الوظيفي','رقم الموظف','Employee Number','Employee ID'])),key=f'{label}id')
            dtcol=st.selectbox('تاريخ الترقية',cols,index=cols.index(guess(cols,['تاريخ تطبيق الترقية','تاريخ الترقية','Promotion Date'])),key=f'{label}date')
            typecol=st.selectbox('نوع الترقية',cols,index=cols.index(guess(cols,['نوع الترقية','Promotion Type'])),key=f'{label}type')
            namecol=st.selectbox('اسم الموظف (اختياري)',['—']+cols,index=(1+cols.index(guess(cols,['اسم الموظف','Employee Name']))) if any(clean(x) in ['اسم الموظف','Employee Name'] for x in cols) else 0,key=f'{label}name')
            selections.append((idcol,dtcol,typecol,namecol))

@st.cache_data(show_spinner=False)
def prep(df,sel,source):
    idcol,dtcol,typecol,namecol=sel
    out=pd.DataFrame({'الرقم الوظيفي':df[idcol].map(norm_id),'التاريخ الأصلي':df[dtcol].map(clean),'تاريخ الترقية':df[dtcol].map(parse_date),'نوع الترقية':df[typecol].map(clean),'اسم الموظف':df[namecol].map(clean) if namecol!='—' else '', 'المصدر':source})
    out['رقم الصف الأصلي']=np.arange(2,len(out)+2)
    out['سبب الاستبعاد']=''
    out.loc[out['الرقم الوظيفي']=='','سبب الاستبعاد']='رقم وظيفي مفقود'
    out.loc[out['تاريخ الترقية'].isna(),'سبب الاستبعاد']=out['سبب الاستبعاد'].apply(lambda x: x+'؛ ' if x else '')+'تاريخ غير صالح أو مفقود'
    return out

all_rows=pd.concat([prep(a,selections[0],'ERP'),prep(b,selections[1],'Cloud')],ignore_index=True)
invalid=all_rows[all_rows['سبب الاستبعاد']!=''].copy()
valid=all_rows[all_rows['سبب الاستبعاد']==''].copy()
valid['نوع الترقية']=valid['نوع الترقية'].replace('', 'غير محدد')
# Merge duplicates on employee + parsed date, not type. Preserve provenance and flag type disagreements.
def combine(g):
    types=sorted(set(g['نوع الترقية'])-{'غير محدد'})
    names=[x for x in g['اسم الموظف'] if x]
    src=sorted(set(g['المصدر']))
    return pd.Series({'اسم الموظف':names[0] if names else '', 'نوع الترقية':types[0] if len(types)==1 else ('غير محدد' if not types else 'يحتاج مراجعة'), 'أنواع الترقية الأصلية':' | '.join(types) if types else 'غير محدد', 'حالة النوع':'تعارض نوع الترقية' if len(types)>1 else ('نوع غير محدد' if not types else 'سليم'), 'المصدر':' + '.join(src),'عدد السجلات المدمجة':len(g)})
merged=valid.groupby(['الرقم الوظيفي','تاريخ الترقية'],dropna=False,sort=False).apply(combine,include_groups=False).reset_index()
