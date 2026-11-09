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
merged=merged.sort_values(['الرقم الوظيفي','تاريخ الترقية']).reset_index(drop=True)
merged['ترتيب الترقية']=merged.groupby('الرقم الوظيفي').cumcount()+1
merged['الترقية السابقة']=merged.groupby('الرقم الوظيفي')['تاريخ الترقية'].shift()
merged['المدة بين الترقيات (يوم)']=(merged['تاريخ الترقية']-merged['الترقية السابقة']).dt.days
latest=merged.sort_values('تاريخ الترقية').drop_duplicates('الرقم الوظيفي',keep='last').copy()
latest['عدد الترقيات']=latest['الرقم الوظيفي'].map(merged.groupby('الرقم الوظيفي').size())
ref=st.sidebar.date_input('تاريخ التحليل المرجعي',value=date.today())
ref=pd.Timestamp(ref)
latest['الأيام منذ آخر ترقية']=(ref-latest['تاريخ الترقية']).dt.days
latest['السنوات منذ آخر ترقية']=(latest['الأيام منذ آخر ترقية']/365.25).round(2)
latest['فئة المدة منذ آخر ترقية']=pd.cut(latest['الأيام منذ آخر ترقية'],bins=[-np.inf,-1,365,1095,1826,np.inf],labels=['تاريخ مستقبلي','أقل من سنة','1–3 سنوات','3–5 سنوات','أكثر من 5 سنوات'])

st.subheader('2. ملخص البيانات الموحدة')
metrics=[('سجلات المصدرين',len(all_rows)),('الترقيات بعد إزالة التكرار',len(merged)),('التكرارات المحذوفة',len(valid)-len(merged)),('الموظفون',merged['الرقم الوظيفي'].nunique()),('سجلات مستبعدة',len(invalid)),('تعارضات نوع الترقية',(merged['حالة النوع']=='تعارض نوع الترقية').sum())]
for col,(lab,val) in zip(st.columns(6),metrics): col.metric(lab,f'{val:,}')
st.caption('التكرار: نفس الرقم الوظيفي ونفس تاريخ الترقية، حتى لو اختلف نوع الترقية. التعارض في النوع لا يُحسم تلقائياً.')

st.sidebar.header('🔎 الفلاتر')
types=sorted(merged['نوع الترقية'].unique().tolist())
selected_types=st.sidebar.multiselect('نوع الترقية',types,default=types)
years=sorted(merged['تاريخ الترقية'].dt.year.unique().tolist())
selected_years=st.sidebar.multiselect('سنة الترقية',years,default=years)
search=st.sidebar.text_input('بحث بالرقم الوظيفي')
filt=merged[merged['نوع الترقية'].isin(selected_types)&merged['تاريخ الترقية'].dt.year.isin(selected_years)].copy()
if search: filt=filt[filt['الرقم الوظيفي'].str.contains(re.escape(search),case=False,na=False)]
fl=latest[latest['الرقم الوظيفي'].isin(filt['الرقم الوظيفي'])].copy()

one,two,three,four,five=st.tabs(['📊 نظرة عامة','🏷️ أنواع الترقيات','👤 آخر ترقية','🧹 جودة البيانات','📥 تصدير Excel'])
with one:
    x,y=st.columns(2)
    with x:
        annual=filt.assign(السنة=filt['تاريخ الترقية'].dt.year).groupby('السنة').size().reset_index(name='عدد الترقيات')
        if not annual.empty: st.plotly_chart(px.bar(annual,x='السنة',y='عدد الترقيات',title='الترقيات حسب السنة'),use_container_width=True)
    with y:
        count=fl['فئة المدة منذ آخر ترقية'].value_counts().reset_index()
        count.columns=['فئة المدة','عدد الموظفين']
        if not count.empty: st.plotly_chart(px.bar(count,x='فئة المدة',y='عدد الموظفين',title='المدة منذ آخر ترقية'),use_container_width=True)
    st.info('Insight: الموظفون الذين مرّ على آخر ترقية لهم أكثر من 5 سنوات ليسوا بالضرورة مستحقين للترقية؛ يتطلب الاستحقاق قواعد الموارد البشرية.')
    if len(fl):
        st.write(f"**أكثر من 5 سنوات منذ آخر ترقية:** {(fl['الأيام منذ آخر ترقية']>1826).sum():,} موظف من أصل {len(fl):,} ضمن نتائج البحث.")
    st.dataframe(filt,use_container_width=True,hide_index=True)
with two:
    type_counts=filt.groupby('نوع الترقية').agg(**{'عدد الترقيات':('الرقم الوظيفي','size'),'عدد الموظفين':('الرقم الوظيفي','nunique')}).reset_index().sort_values('عدد الترقيات',ascending=False)
    st.dataframe(type_counts,use_container_width=True,hide_index=True)
    if not type_counts.empty: st.plotly_chart(px.bar(type_counts,x='نوع الترقية',y='عدد الترقيات',text_auto=True,title='أعداد الترقيات حسب النوع'),use_container_width=True)
    annual_type=filt.assign(السنة=filt['تاريخ الترقية'].dt.year).groupby(['السنة','نوع الترقية']).size().reset_index(name='عدد الترقيات')
    if not annual_type.empty: st.plotly_chart(px.bar(annual_type,x='السنة',y='عدد الترقيات',color='نوع الترقية',barmode='group',title='أنواع الترقيات عبر السنوات'),use_container_width=True)
    latest_type=fl.groupby('نوع الترقية').size().reset_index(name='عدد الموظفين آخر ترقيتهم من هذا النوع')
    st.markdown('**نوع آخر ترقية لكل موظف**')
    st.dataframe(latest_type,use_container_width=True,hide_index=True)
with three:
    st.caption('سجل واحد لكل موظف حسب أحدث تاريخ ترقية في البيانات الموحدة. فلتر السنة والنوع يحدد الموظفين المعروضين، لكن آخر ترقية تُحسب من السجل الكامل.')
    st.dataframe(fl.sort_values('تاريخ الترقية',ascending=False),use_container_width=True,hide_index=True)
    if not merged.empty:
        vals=merged.groupby('الرقم الوظيفي').size().value_counts().sort_index().reset_index()
        vals.columns=['عدد الترقيات للموظف','عدد الموظفين']
        st.plotly_chart(px.bar(vals,x='عدد الترقيات للموظف',y='عدد الموظفين',title='توزيع عدد الترقيات لكل موظف'),use_container_width=True)
with four:
    st.write('**سجلات مستبعدة بسبب رقم وظيفي أو تاريخ غير صالح**')
    st.dataframe(invalid,use_container_width=True,hide_index=True)
    conflicts=merged[merged['حالة النوع']!='سليم']
    st.write('**ترقيات تحتاج مراجعة نوع الترقية**')
    st.dataframe(conflicts,use_container_width=True,hide_index=True)
    duplicate_rows=valid[valid.duplicated(['الرقم الوظيفي','تاريخ الترقية'],keep=False)].sort_values(['الرقم الوظيفي','تاريخ الترقية'])
    st.write('**السجلات الأصلية المتكررة (للتدقيق)**')
    st.dataframe(duplicate_rows,use_container_width=True,hide_index=True)

@st.cache_data(show_spinner=False)
def excel_export(merged,latest,types,invalid,duplicates,conflicts,annual):
    buf=BytesIO()
    sheets={'البيانات الموحدة':merged,'آخر ترقية لكل موظف':latest,'أنواع الترقيات':types,'الاتجاه السنوي':annual,'سجلات مستبعدة':invalid,'سجلات متكررة':duplicates,'مراجعة نوع الترقية':conflicts}
    with pd.ExcelWriter(buf,engine='openpyxl',datetime_format='DD/MM/YYYY',date_format='DD/MM/YYYY') as writer:
        for name,df in sheets.items():
            df.to_excel(writer,sheet_name=name,index=False)
            ws=writer.sheets[name]
            ws.freeze_panes='A2'
            ws.auto_filter.ref=ws.dimensions
            ws.sheet_view.rightToLeft=True
            for cells in ws.columns:
                letter=cells[0].column_letter
                width=min(42,max(14,max(len(str(c.value)) if c.value is not None else 0 for c in cells[:min(len(cells),200)])+2))
                ws.column_dimensions[letter].width=width
                for cell in cells[1:]:
                    if isinstance(cell.value,(datetime,pd.Timestamp)): cell.number_format='DD/MM/YYYY'
    buf.seek(0)
    return buf.getvalue()
with five:
    st.write('يشمل التصدير كل البيانات الموحدة، آخر ترقية، أنواع الترقيات، الاتجاه السنوي، السجلات المستبعدة، والتكرارات وتعارضات النوع.')
    annual_all=merged.assign(السنة=merged['تاريخ الترقية'].dt.year).groupby('السنة').size().reset_index(name='عدد الترقيات')
    all_types=merged.groupby('نوع الترقية').agg(**{'عدد الترقيات':('الرقم الوظيفي','size'),'عدد الموظفين':('الرقم الوظيفي','nunique')}).reset_index()
    duplicate_rows=valid[valid.duplicated(['الرقم الوظيفي','تاريخ الترقية'],keep=False)]
    content=excel_export(merged,latest,all_types,invalid,duplicate_rows,merged[merged['حالة النوع']!='سليم'],annual_all)
    st.download_button('📥 تحميل التقرير الشامل Excel',content,file_name='unified_promotion_analysis.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
