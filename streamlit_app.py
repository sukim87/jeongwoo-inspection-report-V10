from datetime import date
from io import BytesIO
from pathlib import Path
import hashlib,streamlit as st
from PIL import Image,ImageOps
from modules.application_reader import read_application
from modules.report_text import make_draft
from modules.excel_writer import build_report
BASE=Path(__file__).resolve().parent
st.set_page_config(page_title='Jeongwoo AI 검사결과서',page_icon='📋',layout='wide',initial_sidebar_state='expanded')
CODES=[('서류검토','DR'),('열처리','HT'),('재료검사','MI'),('조립검사','AS'),('마킹검사','MK'),('압력시험','PR'),('취부검사','FU'),('성능시험','PF'),('치수검사','DI'),('비파괴검사','NDE'),('외관검사','VI'),('도장/도금','PT'),('가공검사','MA'),('포장/출하','PK'),('성형검사','FM'),('용접검사','WD')]
st.markdown("""<style>html,body,[class*='css']{font-family:'Malgun Gothic','맑은 고딕',sans-serif!important}.block-container{max-width:1220px;padding-top:2.5rem!important}.hero{background:linear-gradient(110deg,#97001c,#d51c3f);color:white;border-radius:12px;padding:26px}.hero h1{font-size:3rem;line-height:1.18;margin:0 0 10px;font-weight:900}.hero p{font-size:1rem;margin:0;line-height:1.6}.head{background:#dedac8;border:1px solid #222;text-align:center;font-weight:900;padding:8px;margin-top:13px}.box{border:1px solid #222;border-top:0;padding:13px;background:white}.side{background:white;border:1px solid #ddd;border-top:4px solid #c8102e;border-radius:8px;padding:13px;line-height:1.65;font-size:.84rem}.manual{background:#fff8dc;border:1px solid #e7ca6b;border-radius:8px;padding:11px 13px;margin:10px 0}label,.stButton button,.stDownloadButton button{font-weight:800!important}.note{font-size:.8rem;color:#626b75;background:#f5f7f9;border:1px solid #d9dde2;border-radius:7px;padding:9px;margin-top:8px}@media(max-width:760px){.hero h1{font-size:2.1rem}}</style>""",unsafe_allow_html=True)
with st.sidebar:
 logo=BASE/'assets/jeongwoo_logo.jpg'
 if logo.exists():st.image(str(logo),use_container_width=True)
 st.markdown("<div class='side'><b>☎ 시스템 문의 및 지원</b><br><br><b>부서:</b> PS품질팀<br><b>담당자:</b> 김선웅<br><b>이메일:</b> sukim1@jimc.co.kr<br><br><span style='color:#687078'>프로그램 오류, 기능 개선, 검사양식 및 입력항목 변경 요청</span></div>",unsafe_allow_html=True)
st.markdown("<div class='hero'><h1>AI로 더 빠르게 만드는 품질검사 결과서</h1><p>사진과 간략하게 검사내용을 입력하고 결과만 확인하세요. 입력 오타를 다듬어 검사 결과서를 작성합니다.</p></div>",unsafe_allow_html=True)
with st.expander('🖼 완성 결과서 예시 보기'):
 p=BASE/'assets/sample_result.png'
 if p.exists():st.image(str(p),use_container_width=True)
st.markdown("<div class='manual'><b>처음 사용하시는 경우, 작성 전에 사용 매뉴얼을 확인해 주세요.</b><br>입력 순서와 사진 배치, AI 초안 확인방법을 미리 살펴보면 중복 작성과 누락을 줄일 수 있습니다.</div>",unsafe_allow_html=True)
manual=BASE/'assets/Jeongwoo_AI_검사결과서_사용매뉴얼.pdf'
if manual.exists():st.download_button('📘 사용 매뉴얼 PDF 다운로드',manual.read_bytes(),file_name='Jeongwoo_AI_검사결과서_사용매뉴얼.pdf',mime='application/pdf',use_container_width=True)
app=st.file_uploader('검사신청서 PDF 또는 Excel',type=['pdf','xlsx'])
if app:
 sig=hashlib.sha256(app.getvalue()).hexdigest()
 if st.session_state.get('appsig')!=sig:
  info,preview=read_application(app)
  for k in ('customer','part_name','order_no','po_step_no','inspection_location'):st.session_state[k]=info.get(k,'')
  st.session_state.appsig=sig;st.session_state.preview=preview
 st.success('확실한 기본정보를 자동 입력했습니다. 필요한 항목은 수정하세요.')
 if st.session_state.get('preview') is not None:
  with st.expander('방향 보정된 검사신청서'):st.image(st.session_state.preview,use_container_width=True)
st.markdown("<div class='head'>기본정보</div><div class='box'>",unsafe_allow_html=True)
a,b=st.columns(2);customer=a.text_input('고객사',key='customer');part=b.text_input('부품명 *',key='part_name')
a,b=st.columns(2);order=a.text_input('수주번호 *',key='order_no');po=b.text_input('발주/공정번호 *',key='po_step_no')
a,b=st.columns(2);location=a.text_input('검사장소 *',key='inspection_location');inspector=b.text_input('검사자',value='김선웅')
a,b=st.columns(2);seq=a.number_input('문서 순번',1,999,1);doc=f"IRR-{order or 'ORDER'}-{po or 'PO-STEP'}-{int(seq):03d}";b.text_input('문서번호',doc,disabled=True)
st.markdown('</div>',unsafe_allow_html=True)
st.markdown("<div class='head'>검사종류</div><div class='box'>",unsafe_allow_html=True)
cols=st.columns(4);selected=[]
for i,(n,c) in enumerate(CODES):
 label=f'{n} ({c})'
 if cols[i%4].checkbox(label,key=c):selected.append(label)
st.markdown('</div>',unsafe_allow_html=True)
st.markdown("<div class='head'>검사 기준 및 결과</div><div class='box'>",unsafe_allow_html=True)
refs=[];rc=st.columns(3)
for col,x in zip(rc,('도면','구매사양서','구매발주서')):
 if col.checkbox(x,key='ref_'+x):refs.append(x)
manual_spec=st.text_input('추가 검사기준 직접 입력',key='manual_spec',placeholder='예: 승인도면 DWG-PS-101 Rev.2 / ASME B31.1 / DFT Min. 320 μm')
spec_parts=refs+([manual_spec.strip()] if manual_spec.strip() else [])
spec=' / '.join(spec_parts)
st.text_area('검사기준 *',value=spec,height=70,disabled=True)
content=st.text_area('검사내용 *',key='inspection_content',height=90,placeholder='예: 치수, 외관 스펙 만족')
judgment=st.selectbox('최종판정',['합격','불합격','조건부 합격','N/A'])
st.markdown('</div>',unsafe_allow_html=True)
st.markdown("<div class='head'>검사사진</div><div class='box'>",unsafe_allow_html=True)
photos=st.file_uploader('검사사진 최대 8장',type=['jpg','jpeg','png'],accept_multiple_files=True);rots=[];fit_modes=[]
if photos:
 pcs=st.columns(2)
 for i,p in enumerate(photos[:8]):
  with pcs[i%2]:
   im=ImageOps.exif_transpose(Image.open(BytesIO(p.getvalue())))
   r=st.selectbox(f'사진 {i+1} 방향',['자동','오른쪽 90°','180°','왼쪽 90°'],key=f'r{i}')
   deg={'자동':0,'오른쪽 90°':90,'180°':180,'왼쪽 90°':270}[r]
   rots.append(deg)
   fit_mode=st.selectbox(f'사진 {i+1} 맞춤',['전체 사진 표시','사진 칸 채우기'],key=f'fit{i}',help='전체 사진 표시는 원본 비율을 유지해 잘리지 않습니다. 사진 칸 채우기는 일부가 잘릴 수 있습니다.')
   fit_modes.append(fit_mode)
   if deg:im=im.rotate(-deg,expand=True)
   if fit_mode=='전체 사진 표시':
    preview=ImageOps.contain(im,(1200,850),Image.Resampling.LANCZOS)
    canvas=Image.new('RGB',(1200,850),'white')
    canvas.paste(preview,((1200-preview.width)//2,(850-preview.height)//2))
    im=canvas
   else:
    im=ImageOps.fit(im,(1200,850),Image.Resampling.LANCZOS)
   st.image(im,use_container_width=True)
st.markdown("<div class='note'>기본값인 전체 사진 표시는 원본 비율을 유지하여 계측기, 눈금, 제품이 잘리지 않습니다. 갑지에는 사진 1~2, 사진 3~8은 을지에 배치됩니다. 8장을 초과하면 검사결과서(을지) 양식을 복사·편집하여 이어서 사용해 주세요.</div></div>",unsafe_allow_html=True)
sig=hashlib.sha256(('|'.join(selected)+'|'+spec+'|'+content+'|'+judgment+'|'+str(len(photos))).encode()).hexdigest()
if st.session_state.get('draft_sig') not in (None,sig):st.session_state.draft_ready=False
if st.button('✨ 오타 교정 후 AI 초안 작성',type='primary',use_container_width=True):
 if not selected or not spec or not content.strip():st.error('검사종류, 검사기준, 검사내용을 먼저 입력해 주세요.')
 else:
  r,a,c,cs,cc=make_draft(selected,spec,content,judgment,len(photos));st.session_state.ai_result=r;st.session_state.ai_action=a;st.session_state.ai_caption=c;st.session_state.corrected_spec=cs;st.session_state.corrected_content=cc;st.session_state.draft_ready=True;st.session_state.draft_sig=sig;st.rerun()
if st.session_state.get('draft_ready'):
 st.markdown("<div class='head'>AI 작성 결과 확인</div><div class='box'>",unsafe_allow_html=True)
 c1,c2=st.columns(2);c1.text_area('교정된 검사기준',st.session_state.corrected_spec,disabled=True);c2.text_area('교정된 검사내용',st.session_state.corrected_content,disabled=True)
 result=st.text_area('검사결과 문장',key='ai_result');action=st.text_area('조치요구',key='ai_action');caption=st.text_area('사진 설명',key='ai_caption')
 st.caption('검사장소는 결과서 요약 칸에만 표시되며 검사결과 문장에는 반복하지 않습니다.')
 st.markdown('</div>',unsafe_allow_html=True)
 missing=[n for n,v in [('부품명',part),('수주번호',order),('발주/공정번호',po),('검사장소',location)] if not str(v).strip()]
 if missing:st.warning('필수 입력: '+', '.join(missing))
 if st.button('✅ 검사결과서 생성',type='primary',use_container_width=True,disabled=bool(missing)):
  data={'customer':customer,'part_name':part,'order_no':order,'po_step_no':po,'inspection_location':location,'inspector':inspector,'date':date.today(),'doc':doc,'types':' / '.join(selected),'spec':st.session_state.corrected_spec,'inspection_content':st.session_state.corrected_content,'judgment':judgment,'result':result,'action':action,'caption':caption}
  try:st.session_state.out,st.session_state.name=build_report(data,photos[:8],rots,fit_modes);st.success('검사결과서가 완성되었습니다.')
  except Exception as e:st.error(f'검사결과서 생성 중 오류: {e}')
 if st.session_state.get('out'):st.download_button('📥 검사결과서 다운로드',st.session_state.out,file_name=st.session_state.name,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',use_container_width=True)
