import io
import re
import unicodedata
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='Ngọc Mỹ | TikTok Ads', page_icon='📊', layout='wide')
st.title('📊 NGỌC MỸ | TIKTOK SHOP ADS')
st.caption('Phân tích 4 cấp báo cáo • ROI (ROAS) = Doanh thu quảng cáo / Chi phí quảng cáo')

PAGES = ['1. Tổng quan', '2. Chiến dịch', '3. Sản phẩm', '4. Mẫu quảng cáo/video']
SYN = {
 'date':['date','day','ngay','theo ngay','thoi gian','stat time','reporting date','time'],
 'campaign':['campaign name','campaign','ten chien dich','campaign id'],
 'product':['product name','product','ten san pham','item name','sku name','product id'],
 'video':['video name','video','creative name','ad name','ten mau quang cao','mau quang cao','id bai dang','ad id','creative id','video id'],
 'spend':['cost','spend','total cost','ad spend','chi phi','tieu dung','amount spent'],
 'revenue':['gross revenue','revenue','gmv','gross merchandise value','doanh thu','doanh thu gop','doanh thu gop cua hang hien tai','sales','attributed revenue','total revenue'],
 'orders':['orders','order','purchases','so don','don hang','sku orders','total orders','so luong don hang sku','so luong don hang sku cua hang hien tai'],
 'impressions':['impressions','impression','hien thi','luot hien thi','so luot hien thi quang cao san pham'],
 'clicks':['clicks','click','luot nhap','product clicks','destination clicks','so luot nhap vao quang cao san pham'],
 'ctr':['ctr','click through rate','ty le nhap','ty le nhap vao quang cao san pham'],
 'roi':['roi','roas','return on investment','return on ad spend'],
 'cpc':['cpc','cost per click'],
 'cpm':['cpm','cost per mille'],
 'start':['start date','start time','ngay bat dau'],
 'end':['end date','end time','ngay ket thuc'],
}
def norm(s):
 s=unicodedata.normalize('NFKD',str(s)); s=''.join(c for c in s if not unicodedata.combining(c))
 return re.sub(r'[^a-z0-9]+',' ',s.lower()).strip()
def num(s):
 if pd.api.types.is_numeric_dtype(s): return pd.to_numeric(s,errors='coerce')
 def parse(x):
  t=str(x).strip().replace('₫','').replace('đ','').replace('%','').replace(' ','')
  if not t or t.lower() in ('nan','none','-'):return np.nan
  if ',' in t and '.' in t:
   t=t.replace('.','').replace(',','.') if t.rfind(',')>t.rfind('.') else t.replace(',','')
  elif ',' in t:
   bits=t.split(',');t=''.join(bits) if len(bits[-1])==3 else t.replace(',','.')
  elif '.' in t and len(t.split('.')[-1])==3:t=t.replace('.','')
  try:return float(t)
  except:return np.nan
 return s.map(parse)
def load_file(f):
 raw=f.getvalue()
 if f.name.lower().endswith('.csv'):
  for encoding in ('utf-8-sig','utf-16','cp1258','latin1'):
   try:return pd.read_csv(io.BytesIO(raw),encoding=encoding,sep=None,engine='python')
   except Exception:pass
  raise ValueError('Không đọc được CSV')
 x=pd.ExcelFile(io.BytesIO(raw)); sheet=st.selectbox('Sheet trong '+f.name,x.sheet_names,key='sheet_'+f.name)
 return pd.read_excel(io.BytesIO(raw),sheet_name=sheet)
def map_cols(df,kind):
 st.caption('Ánh xạ cột: hệ thống tự nhận diện tên cột tiếng Việt trong 4 file TikTok; có thể chỉnh lại nếu cần.')
 used={};cols=['— Không có —']+list(df.columns)
 fields=['date','start','end','spend','revenue','orders','impressions','clicks','ctr','roi']
 if kind=='campaign':fields=['campaign']+fields
 if kind=='product':fields=['product']+fields
 if kind=='video':fields=['video','campaign','product']+fields
 for k in fields:
  candidates=[c for c in df.columns if norm(c) in [norm(a) for a in SYN[k]]]
  default=candidates[0] if candidates else '— Không có —'
  pick=st.selectbox(k,cols,index=cols.index(default),key=f'map_{kind}_{k}')
  if pick!='— Không có —':used[k]=pick
 return used
def prepare(df,m):
 out=pd.DataFrame(index=df.index)
 for k,c in m.items():out[k]=df[c]
 for k in ['spend','revenue','orders','impressions','clicks','ctr','roi']:
  if k in out:out[k]=num(out[k]).fillna(0)
 for k in ['date','start','end']:
  if k in out:out[k]=pd.to_datetime(out[k],errors='coerce',dayfirst=True)
 return out

def aggregate(df,key=None):
 sums=[k for k in ['spend','revenue','orders','impressions','clicks'] if k in df]
 if key and key in df:
  result=df.groupby(key,dropna=False)[sums].sum().reset_index() if sums else df[[key]].drop_duplicates()
 else:result=pd.DataFrame([{k:df[k].sum() for k in sums}])
 for k in ['spend','revenue','orders','impressions','clicks']:
  if k not in result:result[k]=0.0
 result['ROI (ROAS)']=np.where(result.spend>0,result.revenue/result.spend,np.nan)
 result['CPA']=np.where(result.orders>0,result.spend/result.orders,np.nan)
 result['CTR (%)']=np.where(result.impressions>0,100*result.clicks/result.impressions,np.nan)
 result['CVR (%)']=np.where(result.clicks>0,100*result.orders/result.clicks,np.nan)
 result['CPC']=np.where(result.clicks>0,result.spend/result.clicks,np.nan)
 return result

with st.sidebar:
 st.header('Dữ liệu & cấu hình')
 st.info('Tải 4 file báo cáo riêng biệt. Dữ liệu chỉ lưu trong phiên ứng dụng hiện tại.')
 uploaded={}
 for kind,title in [('overview','Tổng quan'),('campaign','Chiến dịch'),('product','Sản phẩm'),('video','Mẫu quảng cáo/video')]:
  uploaded[kind]=st.file_uploader(title,type=['xlsx','xls','csv'],key='file_'+kind)
 st.divider()
 target=st.number_input('ROI (ROAS) mục tiêu (doanh thu / chi phí)',min_value=0.1,value=3.0,step=0.25)
 maxcpa=st.number_input('CPA tối đa (VND/đơn)',min_value=0,value=80000,step=10000)
 minspend=st.number_input('Chi phí tối thiểu để đánh giá (VND)',min_value=0,value=100000,step=50000)
 goodctr=st.number_input('CTR tốt tối thiểu (%)',min_value=0.1,value=2.0,step=0.5)
 page=st.radio('Chọn trang báo cáo',PAGES)

kind=['overview','campaign','product','video'][PAGES.index(page)]
f=uploaded[kind]
if f is None:
 st.info('Vui lòng tải file **'+page+'** ở thanh bên trái để bắt đầu.')
 st.stop()
try:raw=load_file(f)
except Exception as e:st.error(f'Lỗi đọc file: {e}');st.stop()
if raw.empty:st.warning('File không có dữ liệu.');st.stop()
with st.expander('Xem cột dữ liệu và thiết lập ánh xạ',expanded=True):
 st.dataframe(raw.head(4),use_container_width=True)
 mapping=map_cols(raw,kind)
if 'spend' not in mapping:st.error('Cần ánh xạ cột chi phí để tính hiệu quả.');st.stop()
df=prepare(raw,mapping)
if 'date' in df:
 valid=df['date'].dropna()
 if not valid.empty:
  low,high=valid.min().date(),valid.max().date()
  dates=st.date_input('Lọc khoảng thời gian',value=(low,high),min_value=low,max_value=high)
  if isinstance(dates,tuple) and len(dates)==2:df=df[df.date.dt.date.between(dates[0],dates[1])]
  st.write(f'**Kỳ dữ liệu:** {low:%d/%m/%Y} – {high:%d/%m/%Y}')
 else:st.warning('Không đọc được ngày trong cột đã chọn.')
elif 'start' in df or 'end' in df:
 bounds=[df[k].dropna() for k in ['start','end'] if k in df]
 bounds=[b for b in bounds if len(b)]
 if bounds:st.write('**Khoảng thời gian ghi trên báo cáo:**',min(b.min() for b in bounds).strftime('%d/%m/%Y'),'–',max(b.max() for b in bounds).strftime('%d/%m/%Y'))
else:
 st.info('Kỳ báo cáo theo tên file: 14/10/2026 – 19/10/2026. Nếu upload kỳ khác, hãy nhập thời gian ở bên dưới.')
 period=st.date_input('Chọn kỳ báo cáo cho file không có cột ngày',value=(pd.to_datetime('2026-10-14').date(),pd.to_datetime('2026-10-19').date()),key='period_'+kind)
 if isinstance(period,tuple) and len(period)==2:st.write(f'**Kỳ đang phân tích:** {period[0]:%d/%m/%Y} – {period[1]:%d/%m/%Y}')
if df.empty:st.warning('Không có dữ liệu trong khoảng đã chọn.');st.stop()
if 'revenue' not in df:st.warning('Thiếu doanh thu: ROAS không thể xác định chính xác.')
if 'orders' not in df:st.warning('Thiếu đơn hàng: không thể đánh giá chuyển đổi/CPA.')

summary=aggregate(df).iloc[0]
cols=st.columns(5)
for col,(label,value,fmt) in zip(cols,[('Chi phí',summary.spend,',.0f'),('Doanh thu',summary.revenue,',.0f'),('Đơn hàng',summary.orders,',.0f'),('ROI (ROAS)',summary['ROI (ROAS)'],'.2f'),('CPA',summary.CPA,',.0f')]):
 col.metric(label,format(value,fmt) if pd.notna(value) else 'N/A')
st.caption('ROI (ROAS) = doanh thu / chi phí; CPA = chi phí / số đơn; CTR = lượt nhấp / hiển thị; CVR = đơn / lượt nhấp. Các tỷ lệ được tính lại từ tổng, không cộng/trung bình tỷ lệ từng dòng.')

if kind=='overview':
 st.subheader('Tổng quan hiệu quả')
 if 'date' in df and df.date.notna().any():
  d=df.groupby('date',as_index=False)[[c for c in ['spend','revenue','orders'] if c in df]].sum()
  fig=px.line(d,x='date',y=[c for c in ['spend','revenue'] if c in d],markers=True,title='Chi phí và doanh thu theo ngày')
  st.plotly_chart(fig,use_container_width=True)
  if 'orders' in d:st.plotly_chart(px.bar(d,x='date',y='orders',title='Đơn hàng theo ngày'),use_container_width=True)
 else:st.info('Thêm cột ngày để vẽ xu hướng theo ngày.')
 st.dataframe(aggregate(df).round(2),use_container_width=True)
else:
 key={'campaign':'campaign','product':'product','video':'video'}[kind]
 if key not in df:st.error(f'Cần ánh xạ cột {key}.');st.stop()
 a=aggregate(df,key)
 def judge(r):
  if r.spend<minspend:return '⚪ Chưa đủ dữ liệu','Chi phí chưa đạt ngưỡng đánh giá.'
  if 'orders' not in df or 'revenue' not in df:return '⚪ Thiếu dữ liệu','Cần doanh thu và đơn hàng để kết luận.'
  if r.orders>0 and r['ROI (ROAS)']>=target and r.CPA<=maxcpa:return '🟢 Tốt','Có đơn; ROI (ROAS) đạt mục tiêu và CPA trong giới hạn.'
  if r.orders==0 and kind=='video' and pd.notna(r['CTR (%)']) and r['CTR (%)']>=goodctr and r.clicks>=20:return '🟡 Có tiềm năng','CTR tốt, có ít nhất 20 nhấp nhưng chưa có đơn; kiểm tra giá, trang sản phẩm, ưu đãi và hành trình mua.'
  if r.orders==0:return '🔴 Cần xem xét tắt','Đã tiêu ngân sách tối thiểu nhưng chưa ghi nhận đơn.'
  if pd.notna(r['ROI (ROAS)']) and r['ROI (ROAS)']<target and r.spend>=2*minspend:return '🔴 Cần xem xét tắt','Chi phí đã vượt gấp đôi ngưỡng đánh giá nhưng ROI (ROAS) vẫn dưới mục tiêu.'
  if pd.notna(r.CPA) and r.CPA>maxcpa:return '🔴 Cần tối ưu','CPA vượt giới hạn cho phép.'
  if r['ROI (ROAS)']>=target:return '🟡 Theo dõi','ROI (ROAS) đạt mục tiêu nhưng CPA vượt giới hạn hoặc chưa đủ điều kiện xanh.'
  return '🟡 Theo dõi','Có đơn nhưng ROI (ROAS) chưa đạt mục tiêu; cần thử tối ưu trước khi tăng ngân sách.'
 decisions=a.apply(judge,axis=1)
 a['Đánh giá']=[x[0] for x in decisions];a['Lý do / hành động']=[x[1] for x in decisions]
 st.subheader({'campaign':'Phân loại chiến dịch','product':'Cảnh báo hiệu quả sản phẩm','video':'Đánh giá mẫu video quảng cáo'}[kind])
 def paint(row):
  status=row['Đánh giá']
  color='#ffe1e1' if '🔴' in status else '#e0f4e7' if '🟢' in status else '#fff1c9' if '🟡' in status else '#edf0f4'
  return [f'background-color: {color}; color: #222222' for _ in row]
 st.dataframe(a.style.apply(paint,axis=1).format({c:'{:,.2f}' for c in ['ROI (ROAS)','CPA','CTR (%)','CVR (%)','CPC']},na_rep='—'),use_container_width=True,hide_index=True)
 st.plotly_chart(px.bar(a.sort_values('spend',ascending=False).head(20),x=key,y='spend',color='Đánh giá',title='Top 20 theo chi phí'),use_container_width=True)
 if 'revenue' in df:
  st.plotly_chart(px.scatter(a,x='spend',y='revenue',size='clicks' if 'clicks' in df else 'orders',color='Đánh giá',hover_name=key,title='Chi phí so với doanh thu'),use_container_width=True)
 st.caption('Các ngưỡng là tiêu chí quản trị có thể điều chỉnh, không phải ngưỡng hòa vốn. Không tự động tắt quảng cáo; lưu ý độ trễ ghi nhận đơn và mức độ đủ dữ liệu.')
 csv=a.to_csv(index=False).encode('utf-8-sig')
 st.download_button('⬇️ Xuất bảng đánh giá CSV',csv,file_name=f'ngocmy_{kind}_assessment.csv',mime='text/csv')
