import io, re, os
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='Ngọc Mỹ • TikTok Ads',page_icon='📊',layout='wide',initial_sidebar_state='expanded')
st.markdown('''<style>.block-container{padding-top:1.3rem;max-width:1500px}.stMetric{background:#f5f7fb;border:1px solid #e5eaf1;padding:14px;border-radius:12px}div[data-testid="stSidebar"]{background:#f7f9fc}h1{font-size:2rem!important}h2{font-size:1.4rem!important}</style>''',unsafe_allow_html=True)
BASE=Path(__file__).parent
FILES={'overview':'Campaign overview data 20261014 - 20261019.xlsx','campaign':'Product campaign data 2026-10-14 - 2026-10-19.xlsx','product':'Product data 2026-10-14 - 2026-10-19.xlsx','video':'Creative data 2026-10-14 - 2026-10-19 - Product 0.xlsx'}
NAMES={'overview':'📊 Tổng quan','campaign':'📣 Chiến dịch','product':'📦 Sản phẩm','video':'🎬 Video quảng cáo'}
COLS={'date':['Theo ngày'],'name_campaign':['Tên chiến dịch'],'name_product':['Tên sản phẩm'],'name_video':['Mẫu quảng cáo'],'id_product':['ID sản phẩm'],'spend':['Chi phí'],'revenue':['Doanh thu gộp (Cửa hàng hiện tại)','Doanh thu gộp'],'orders':['Số lượng đơn hàng SKU (Cửa hàng hiện tại)','Số lượng đơn hàng SKU'],'impressions':['Số lượt hiển thị quảng cáo sản phẩm'],'clicks':['Số lượt nhấp vào quảng cáo sản phẩm'],'ctr':['Tỷ lệ nhấp vào quảng cáo sản phẩm']}
def read(f):
 if f is None:return None
 if isinstance(f,(str,Path)):return pd.read_excel(f)
 return pd.read_csv(f) if f.name.lower().endswith('.csv') else pd.read_excel(f)
def standard(raw):
 if raw is None:return None
 d=pd.DataFrame(index=raw.index)
 for k,choices in COLS.items():
  for c in choices:
   if c in raw.columns:d[k]=raw[c];break
 for c in ['spend','revenue','orders','impressions','clicks','ctr']:
  if c in d:d[c]=pd.to_numeric(d[c].astype(str).str.replace(',','',regex=False).str.replace('%','',regex=False),errors='coerce')
 if 'date' in d:d['date']=pd.to_datetime(d.date,errors='coerce')
 for c in ['spend','revenue','orders','impressions','clicks']:
  if c not in d:d[c]=np.nan
 return d

def agg(d,key=None):
 metrics=['spend','revenue','orders','impressions','clicks']
 if key: a=d.groupby(key,dropna=False)[metrics].sum(min_count=1).reset_index()
 else:a=pd.DataFrame([d[metrics].sum(min_count=1)])
 a['ROI (ROAS)']=np.where(a.spend>0,a.revenue/a.spend,np.nan)
 a['CPA']=np.where(a.orders>0,a.spend/a.orders,np.nan)
 a['CTR (%)']=np.where(a.impressions>0,a.clicks/a.impressions*100,np.nan)
 a['CVR (%)']=np.where(a.clicks>0,a.orders/a.clicks*100,np.nan)
 return a

def money(x):return '—' if pd.isna(x) else f'{x:,.0f} đ'
def number(x):return '—' if pd.isna(x) else f'{x:,.0f}'
def roi(x):return '—' if pd.isna(x) else f'{x:.2f}x'
def paint(row):
 s=row.get('Đánh giá','')
 color='#ffe2e3' if s.startswith('🔴') else '#def5e7' if s.startswith('🟢') else '#fff0c5' if s.startswith('🟡') else '#edf1f6'
 return [f'background-color:{color};color:#17212f' for _ in row]
def show_table(a,sort='spend'):
 cols={'spend':'Chi phí (đ)','revenue':'Doanh thu (đ)','orders':'Đơn hàng','clicks':'Lượt nhấp','impressions':'Hiển thị'}
 b=a.sort_values(sort,ascending=False).rename(columns=cols)
 formats={'Chi phí (đ)':'{:,.0f}','Doanh thu (đ)':'{:,.0f}','Đơn hàng':'{:,.0f}','Lượt nhấp':'{:,.0f}','Hiển thị':'{:,.0f}','ROI (ROAS)':'{:.2f}','CPA':'{:,.0f}','CTR (%)':'{:.2f}','CVR (%)':'{:.2f}','ROI mục tiêu':'{:.2f}'}
 st.dataframe(b.style.apply(paint,axis=1).format({k:v for k,v in formats.items() if k in b},na_rep='—'),use_container_width=True,hide_index=True,height=min(660,95+len(b)*35))
 st.download_button('⬇️ Xuất kết quả CSV',a.to_csv(index=False).encode('utf-8-sig'),'ngocmy_danh_gia.csv','text/csv')

with st.sidebar:
 st.title('NGỌC MỸ')
 st.caption('Quản trị hiệu quả TikTok Shop Ads')
 page=st.radio('CHỌN BÁO CÁO',list(NAMES),format_func=lambda x:NAMES[x])
 st.divider()
 with st.expander('📂 Tải dữ liệu kỳ mới',expanded=False):
  st.caption('Mỗi mục tương ứng một file; không tải lên sẽ dùng dữ liệu mẫu 14–19/10/2026 nếu file có sẵn.')
  uploads={k:st.file_uploader(NAMES[k],type=['xlsx','csv'],key='up_'+k) for k in FILES}
 with st.expander('⚙️ Quy tắc đánh giá',expanded=False):
  minspend=st.number_input('Chi phí tối thiểu để đánh giá (đ)',0,100000,10000)
  maxcpa=st.number_input('CPA tham khảo tối đa (đ/đơn)',0,80000,10000)
  goodctr=st.number_input('CTR tốt từ (%)',0.0,2.0,0.25)
  minclicks=st.number_input('Lượt nhấp tối thiểu để xem xét video',1,20,5)
  st.caption('CPA là cảnh báo phụ, không thay thế ROI mục tiêu riêng.')

raws={}
for k,path in FILES.items():
 try:raws[k]=read(uploads[k] if uploads[k] is not None else (BASE/path if (BASE/path).exists() else None))
 except Exception as e:raws[k]=None;st.sidebar.error(f'{NAMES[k]}: {e}')
data={k:standard(v) for k,v in raws.items()}
# Date range is based on overview daily dates when available; otherwise extracted from uploaded filename.
def period_from_name(name):
 nums=re.findall(r'20\d{6}|20\d{2}-\d{2}-\d{2}',name)
 try:
  if len(nums)>=2:return pd.to_datetime(nums[0]).date(),pd.to_datetime(nums[1]).date()
 except Exception:pass
 return None
ref=data['overview']
period=(ref.date.min().date(),ref.date.max().date()) if ref is not None and 'date' in ref and ref.date.notna().any() else None
if uploads[page] is not None:
 period=period_from_name(uploads[page].name) or period
if period is None:period=(pd.Timestamp('2026-10-14').date(),pd.Timestamp('2026-10-19').date())
st.title(NAMES[page]+' | TikTok Shop Ngọc Mỹ')
st.caption('ROI (ROAS) = Doanh thu quảng cáo ÷ Chi phí quảng cáo. Không sử dụng lợi nhuận hoặc ROI hòa vốn.')
if page=='overview' and ref is not None and 'date' in ref:
 dates=st.date_input('Khoảng ngày',value=period,min_value=ref.date.min().date(),max_value=ref.date.max().date())
 if isinstance(dates,tuple) and len(dates)==2:period=dates
st.info(f'🗓️ **Kỳ báo cáo: {period[0]:%d/%m/%Y} – {period[1]:%d/%m/%Y}**  •  Nguồn: {"file vừa tải lên" if uploads[page] is not None else "file mẫu đính kèm"}')
d=data[page]
if d is None:
 st.warning('Chưa có file cho trang này. Mở **Tải dữ liệu kỳ mới** ở thanh bên để tải báo cáo tương ứng.');st.stop()
if page=='overview' and 'date' in d:d=d[d.date.dt.date.between(*period)]
if d.empty:st.warning('Không có dòng dữ liệu trong kỳ đã chọn.');st.stop()
required=['spend','revenue','orders']
missing=[x for x in required if d[x].notna().sum()==0]
if missing:st.error('**Thiếu dữ liệu:** File không có giá trị hợp lệ cho: '+', '.join(missing)+'. Không thể đánh giá đầy đủ.')
s=agg(d).iloc[0]
metrics=[('Chi phí',money(s.spend)),('Doanh thu',money(s.revenue)),('Đơn hàng',number(s.orders)),('ROI (ROAS)',roi(s['ROI (ROAS)'])),('CPA',money(s.CPA))]
for col,(label,val) in zip(st.columns(5),metrics):col.metric(label,val)
st.divider()
if page=='overview':
 st.subheader('Diễn biến theo ngày')
 if 'date' in d and d.date.notna().any():
  daily=d.groupby('date',as_index=False)[['spend','revenue','orders']].sum()
  left,right=st.columns([1.5,1])
  with left:st.plotly_chart(px.line(daily,x='date',y=['spend','revenue'],markers=True,labels={'value':'VNĐ','date':'Ngày','variable':'Chỉ số'},title='Chi phí và doanh thu quảng cáo'),use_container_width=True)
  with right:st.plotly_chart(px.bar(daily,x='date',y='orders',title='Đơn hàng theo ngày',labels={'date':'Ngày','orders':'Số đơn'}),use_container_width=True)
  daily['ROI (ROAS)']=np.where(daily.spend>0,daily.revenue/daily.spend,np.nan)
  st.subheader('Bảng chi tiết theo ngày')
  st.dataframe(daily.rename(columns={'date':'Ngày','spend':'Chi phí','revenue':'Doanh thu','orders':'Đơn hàng'}).style.format({'Chi phí':'{:,.0f}','Doanh thu':'{:,.0f}','Đơn hàng':'{:,.0f}','ROI (ROAS)':'{:.2f}'},na_rep='—'),hide_index=True,use_container_width=True)
 else:st.warning('Thiếu cột ngày nên chưa thể vẽ biểu đồ theo ngày.')
 st.caption('Các cấp báo cáo có thể dùng phạm vi ghi nhận khác nhau; không cộng doanh thu/đơn giữa các file.')
 st.stop()

key={'campaign':'name_campaign','product':'name_product','video':'name_video'}[page]
if key not in d:
 st.error(f'Thiếu cột định danh {key}. Hãy kiểm tra báo cáo tải lên.');st.stop()
a=agg(d,key)
# ROI thresholds: product-specific, keyed by product ID when available. Stored for the current Streamlit session.
if 'targets' not in st.session_state:st.session_state.targets={}
product_raw=raws['product']
if product_raw is not None and 'Tên sản phẩm' in product_raw:
 p=product_raw[['Tên sản phẩm']+(['ID sản phẩm'] if 'ID sản phẩm' in product_raw else [])].drop_duplicates().copy()
 p['ROI mục tiêu']=p['Tên sản phẩm'].map(st.session_state.targets)
 if page in ('product','campaign','video'):
  with st.expander('🎯 Cấu hình ROI mục tiêu RIÊNG cho từng sản phẩm',expanded=page=='product'):
   st.caption('Điền ROI mục tiêu do người quản trị quyết định. Ô trống = chưa thiết lập; không tự áp ROI = 3 cho tất cả sản phẩm.')
   edited=st.data_editor(p,hide_index=True,use_container_width=True,disabled=[c for c in p if c!='ROI mục tiêu'],column_config={'ROI mục tiêu':st.column_config.NumberColumn('ROI mục tiêu (x)',min_value=0.01,step=0.1,format='%.2f')},key='targets_editor')
   st.session_state.targets={str(r['Tên sản phẩm']):float(r['ROI mục tiêu']) for _,r in edited.iterrows() if pd.notna(r['ROI mục tiêu']) and float(r['ROI mục tiêu'])>0}
   st.download_button('⬇️ Xuất bảng ngưỡng ROI',edited.to_csv(index=False).encode('utf-8-sig'),'roi_muc_tieu_san_pham.csv','text/csv')
else:st.warning('Chưa có danh mục sản phẩm để thiết lập ROI mục tiêu riêng.')

if page=='product':
 a['ROI mục tiêu']=a[key].map(st.session_state.targets)
elif page=='campaign':
 # Campaign with multiple products: compare product-by-product, not against a made-up shared threshold.
 if data['product'] is not None and 'name_campaign' in data['product'] and 'name_product' in data['product']:
  pd0=agg(data['product'],['name_campaign','name_product']) if False else data['product'].copy()
  pd0['target']=pd0.name_product.map(st.session_state.targets)
  def target_for_campaign(name):
   sub=pd0[pd0.name_campaign==name]
   if sub.empty or sub.target.isna().any() or sub.spend.isna().any() or sub.spend.sum()<=0:return np.nan
   # Weighted threshold is an operational comparison, not break-even ROI.
   return np.average(sub.target,weights=sub.spend) if sub.spend.sum()>0 else np.nan
  a['ROI mục tiêu']=a[key].map(target_for_campaign)
 else:a['ROI mục tiêu']=np.nan
else:
 if 'name_product' in d:
  # Aggregate per creative and require a single product to determine its target.
  by=d.groupby(key)['name_product'].agg(lambda x: list(pd.unique(x.dropna())))
  a['ROI mục tiêu']=a[key].map(lambda n:st.session_state.targets.get(str(by[n][0])) if n in by and len(by[n])==1 else np.nan)
 else:a['ROI mục tiêu']=np.nan

def assess(r):
 if pd.isna(r.spend) or pd.isna(r.revenue) or pd.isna(r.orders):return '⚪ Thiếu dữ liệu','Thiếu chi phí, doanh thu hoặc số đơn trong báo cáo; cần bổ sung file/cột.'
 if r.spend<minspend:return '⚪ Chưa đủ dữ liệu',f'Đã có số liệu nhưng mới chi {money(r.spend)}, dưới ngưỡng đánh giá {money(minspend)}.'
 if pd.isna(r['ROI mục tiêu']):return '⚪ Chưa thiết lập mục tiêu','Chưa có ROI mục tiêu riêng cho sản phẩm (hoặc chiến dịch/video gồm nhiều sản phẩm chưa đủ ngưỡng).'
 t=r['ROI mục tiêu'];actual=r['ROI (ROAS)']
 if page=='video' and r.orders==0 and pd.notna(r['CTR (%)']) and r['CTR (%)']>=goodctr and r.clicks>=minclicks:
  return '🟡 Thu hút tốt, chưa ra đơn',f'CTR {r["CTR (%)"]:.2f}% ≥ {goodctr:.2f}%, có {number(r.clicks)} nhấp nhưng chưa ghi nhận đơn. Kiểm tra trang sản phẩm, giá và ưu đãi.'
 if r.orders==0:return '🔴 Cần xem xét tắt',f'Đã chi {money(r.spend)} nhưng chưa có đơn. Kiểm tra độ trễ ghi nhận trước khi tạm dừng.'
 if actual>=t and (pd.isna(r.CPA) or r.CPA<=maxcpa or maxcpa==0):return '🟢 Hiệu quả tốt',f'ROI {roi(actual)} đạt mục tiêu {roi(t)}; có {number(r.orders)} đơn.'
 if actual<t and r.spend>=2*minspend:return '🔴 Cần tối ưu / xem xét tắt',f'ROI {roi(actual)} thấp hơn mục tiêu {roi(t)}, đã chi {money(r.spend)}. Rà soát phân phối và nội dung.'
 if r.CPA>maxcpa and maxcpa>0:return '🟡 Cần tối ưu CPA',f'CPA {money(r.CPA)} vượt mức tham khảo {money(maxcpa)}; ROI {roi(actual)} so với mục tiêu {roi(t)}.'
 return '🟡 Theo dõi thêm',f'ROI {roi(actual)} so với mục tiêu {roi(t)}; chưa đủ điều kiện kết luận nên tắt.'
res=a.apply(assess,axis=1)
a['Đánh giá']=[r[0] for r in res];a['Lý do / hành động']=[r[1] for r in res]
counts=a['Đánh giá'].str.slice(0,1).value_counts()
for c,(label,prefix) in zip(st.columns(4),[('Tốt','🟢'),('Cần chú ý','🟡'),('Cảnh báo','🔴'),('Chưa thể kết luận','⚪')]):c.metric(label,int(counts.get(prefix,0)))
st.subheader('Đánh giá và lý do cụ thể')
st.caption('Màu đỏ: cảnh báo cần xem xét • Xanh: đạt tiêu chí • Vàng: có tiềm năng hoặc cần theo dõi • Trắng: chưa đủ dữ liệu hoặc chưa cấu hình mục tiêu.')
show_table(a)
left,right=st.columns(2)
with left:st.plotly_chart(px.bar(a.nlargest(15,'spend'),x=key,y='spend',color='Đánh giá',title='Chi phí quảng cáo cao nhất'),use_container_width=True)
with right:st.plotly_chart(px.scatter(a,x='spend',y='revenue',color='Đánh giá',hover_name=key,title='Chi phí so với doanh thu'),use_container_width=True)
with st.expander('ℹ️ Phân biệt “Thiếu dữ liệu” và “Chưa đủ dữ liệu”'):
 st.markdown('**Thiếu dữ liệu:** không có trường cần thiết hoặc giá trị trống nên không tính được chỉ số. **Chưa đủ dữ liệu:** đã có số liệu nhưng lượng chi tiêu/nhấp còn thấp so với ngưỡng kiểm chứng, nên chưa đủ cơ sở kết luận. **Chưa thiết lập mục tiêu:** dữ liệu đầy đủ nhưng chưa nhập ROI mục tiêu riêng; đây là trạng thái khác cả hai trường hợp trên.')
st.caption('Các gợi ý chỉ hỗ trợ quyết định, không tự động tắt chiến dịch. Chỉ so sánh cùng phạm vi phân bổ và kỳ báo cáo.')
