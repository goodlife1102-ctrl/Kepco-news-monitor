# -*- coding: utf-8 -*-
import streamlit as st
import requests
import pandas as pd
import plotly.graph_objects as go
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime, timedelta
import re, io
from collections import Counter
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import streamlit.components.v1 as components

try:
    import yfinance as yf
    YF_OK = True
except:
    YF_OK = False

CLIENT_ID     = "mQj4FzmR2tmebJUYk4uq"
CLIENT_SECRET = "zWsKNP7xrB"

MEDIA_GRADE = {
    "조선일보":{"rank":1,"rate":3.73,"grade":"S"},
    "중앙일보":{"rank":2,"rate":2.45,"grade":"A"},
    "동아일보":{"rank":3,"rate":1.95,"grade":"A"},
    "매일경제":{"rank":4,"rate":0.97,"grade":"A"},
    "한겨레":  {"rank":5,"rate":0.62,"grade":"B"},
    "한국경제":{"rank":6,"rate":0.43,"grade":"B"},
    "경향신문":{"rank":7,"rate":0.41,"grade":"B"},
    "한국일보":{"rank":8,"rate":0.31,"grade":"B"},
    "국민일보":{"rank":9,"rate":0.19,"grade":"B"},
    "문화일보":{"rank":10,"rate":0.12,"grade":"C"},
    "서울신문":{"rank":11,"rate":0.10,"grade":"C"},
    "서울경제":{"rank":12,"rate":0.03,"grade":"C"},
    "세계일보":{"rank":13,"rate":0.02,"grade":"D"},
}
GRADE_COLOR = {"S":"#B71C1C","A":"#E64A19","B":"#1565C0","C":"#2E7D32","D":"#616161"}
POSITIVE_WORDS = ["성과","달성","개선","혁신","성장","우수","협력","기여","선정","수상","기대","친환경","안정","흑자","증가","추진","완료","승인","투자","확대","증대","국익","선도","도약","강화","지원","활성화","성공","최초","호평","기록","돌파","수주","계약","협약","양해각서","출범","개통","준공","신기록"]
NEGATIVE_WORDS = ["사고","부패","비리","손실","적자","위기","파업","갈등","논란","문제","우려","하락","감소","실패","사망","중단","취소","반발","지연","비판","폭탄","부담","폐쇄","위반","처벌","고발","감사원","적발","의혹","과태료","경고","해임","부실","낭비","특혜","불법","소송","압수수색","조사","수사","민원","항의","경찰","패소","기소","고소","피해","피의자","벌금","구속","재판","징계","결함","은폐","허위","과장","오염","누출","폭발","붕괴","침수","정전"]
IRRELEVANT_PATTERNS = [r'배구',r'축구',r'야구',r'농구',r'골프',r'올림픽',r'월드컵',r'선수단',r'드래프트',r'챔피언십']
MEDIA_MAP = {"chosun":"조선일보","joongang":"중앙일보","donga":"동아일보","hani":"한겨레","khan":"경향신문","yna":"연합뉴스","ytn":"YTN","imnews":"MBC","kbs":"KBS","sbs":"SBS","mt.co":"머니투데이","edaily":"이데일리","heraldcorp":"헤럴드경제","newsis":"뉴시스","newspim":"뉴스핌","etnews":"전자신문","energy-news":"에너지신문","electimes":"일렉트릭타임스","hankyung":"한국경제","mk.co":"매일경제","sedaily":"서울경제","ajunews":"아주경제","businesspost":"비즈니스포스트","fnnews":"파이낸셜뉴스","inews24":"아이뉴스24","dt.co":"디지털타임스","hankookilbo":"한국일보","munhwa":"문화일보","ohmynews":"오마이뉴스","pressian":"프레시안","energydaily":"에너지데일리","kpinews":"KPI뉴스","naeil":"내일신문","seoul":"서울신문","ekn":"에너지경제","kukminilbo":"국민일보","segyetimes":"세계일보","e2news":"이투뉴스"}
TOPIC_GROUPS = {"전기요금":["전기요금","요금","전력요금","인상","누진제","전기세"],"원전·수출":["원전","수출","원자력","UAE","체코","APR","해외수주"],"재무·경영":["흑자","적자","부채","재무","비상경영","원가","실적","손실"],"전력망·설비":["송전","배전","전력망","변전","선로","전력설비","정전","계통"],"탄소중립·에너지전환":["탄소중립","RE100","온실가스","수소","재생에너지","넷제로","태양광","풍력"],"노사관계":["노사","노조","파업","임금","단체협약","쟁의"],"안전·사고":["안전","사고","재해","산재","폭발","화재","부상","사망"],"AI·디지털혁신":["AI","인공지능","디지털","스마트","자동화","AX","빅데이터"],"공기업·거버넌스":["공기업","감사","이사회","투명","거버넌스","윤리","비리"],"고객·서비스":["서비스","고객","민원","복지","국민","전기복지"],"정책·규제":["정책","규제","법안","제도","정부","국회","의원","경찰","조사","소송"]}
DISAMBIG_MAP = {"김동철":["한전","사장","한국전력","KEPCO"],"이창양":["장관","산업부"]}
FONT_KR = "Malgun Gothic, Apple SD Gothic Neo, sans-serif"
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# ── 유틸 ──────────────────────────────────────────────
def clean(t): return re.sub(r'<[^>]+>','',str(t)).strip()
def get_media(o,l):
    url=o if o else l
    for d,n in MEDIA_MAP.items():
        if d in url: return n
    try: return url.split("//")[-1].split("/")[0].replace("www.","").split(".")[0]
    except: return "기타"
def is_relevant(t): return not any(re.search(p,t) for p in IRRELEVANT_PATTERNS)
def get_sentiment(t):
    p=sum(1 for w in POSITIVE_WORDS if w in t); n=sum(1 for w in NEGATIVE_WORDS if w in t)
    return "긍정" if p>n else "부정" if n>p else "중립"
def summarize(t,n=50):
    t=re.sub(r'\s+',' ',t).strip(); return t[:n]+"..." if len(t)>n else t
def parse_kw(raw):
    raw=raw.replace("(","").replace(")",""); result=[]
    for p in [x.strip() for x in raw.split(",") if x.strip()]:
        if "+" in p:
            sub=[k.strip() for k in p.split("+") if k.strip()]
            result.append({"type":"AND","keywords":sub,"label":" + ".join(sub)})
        else: result.append({"type":"SINGLE","keywords":[p],"label":p})
    return result
def matches_and(t,g): return all(k in t for k in g["keywords"])
def apply_disambig(arts,label):
    for base,req in DISAMBIG_MAP.items():
        if base in label:
            return [a for a in arts if any(r in a["헤드라인"]+" "+a.get("요약","") for r in req)]
    return arts
def get_news(q,mx=1000):
    url="https://openapi.naver.com/v1/search/news.json"
    hdr={"X-Naver-Client-Id":CLIENT_ID,"X-Naver-Client-Secret":CLIENT_SECRET}
    items,s=[],1
    while s<=mx:
        try:
            r=requests.get(url,headers=hdr,params={"query":q,"display":100,"start":s,"sort":"date"},timeout=10)
            batch=r.json().get("items",[])
            if not batch: break
            items.extend(batch)
            if len(batch)<100: break
            s+=100
        except: break
    return items
def auto_cat(arts):
    for a in arts:
        t=a["헤드라인"]+" "+a.get("요약","")
        sc={c:sum(1 for w in ws if w in t) for c,ws in TOPIC_GROUPS.items()}
        sc={k:v for k,v in sc.items() if v>0}
        a["카테고리"]=max(sc,key=sc.get) if sc else "기타"
    return arts
def extract_kws(arts,sent,n=5):
    ft=[a for a in arts if a["감성"]==sent]
    txt=" ".join([a["헤드라인"]+" "+a.get("요약","") for a in ft])
    pool=(NEGATIVE_WORDS if sent=="부정" else POSITIVE_WORDS if sent=="긍정" else
          ["발표","계획","추진","검토","협의","논의","회의","방문","출범","개최","진행","예정"])
    cnt={w:txt.count(w) for w in pool if txt.count(w)>0}
    return sorted(cnt.items(),key=lambda x:x[1],reverse=True)[:n]
def get_key_issues(arts):
    txt=" ".join([a["헤드라인"] for a in arts])
    kws=["전기요금","원전수출","탄소중립","파업","부채","감사원","비상경영","AI","안전","노조","규제","재생에너지","전력망","경찰"]
    found=[(k,txt.count(k)) for k in kws if txt.count(k)>=2]; found.sort(key=lambda x:x[1],reverse=True)
    return [f[0] for f in found[:5]]
def gen_criticisms(arts,kw):
    neg=[a for a in arts if a["감성"]=="부정"]; cat_c=Counter([a["카테고리"] for a in neg])
    DB={"전기요금":{"title":"전기요금 인상 부담","points":["요금 현실화 국민 공감대 부족","저소득층 에너지 부담 대안 미흡"]},"재무·경영":{"title":"재무구조 악화 우려","points":["부채 증가·재무건전성 의문","비상경영 조치 실효성 지적"]},"노사관계":{"title":"노사갈등·파업 리스크","points":["단체협약 갈등으로 경영 불안","공공서비스 안정성 우려"]},"공기업·거버넌스":{"title":"공기업 투명성 지적","points":["경영 비효율 반복 지적","낙하산 인사·지배구조 문제"]},"안전·사고":{"title":"현장 안전사고 우려","points":["안전관리 체계 실효성 의문","협력사 안전망 확대 요구"]},"전력망·설비":{"title":"전력망 노후화 문제","points":["노후 설비 정전·사고 리스크","현대화 투자 속도 부족"]},"탄소중립·에너지전환":{"title":"탄소중립 이행 실효성","points":["이행 속도 저조 지적","전환 비용 현실성 논란"]},"정책·규제":{"title":"정책 투명성·법적 리스크","points":["일방적 정책 추진 지적","경찰 조사·소송 리스크"]},"원전·수출":{"title":"원전 수출 신뢰성","points":["추진 일정 지연·불확실성","안전 기준 논란"]},"AI·디지털혁신":{"title":"디지털 전환 실효성","points":["투자 대비 성과 부족","보안·개인정보 리스크"]},"고객·서비스":{"title":"고객 서비스 대응 미흡","points":["민원 처리 속도 불만","소외계층 접근성 개선 요구"]}}
    result=[]
    for cat,cnt2 in cat_c.most_common(8):
        if cat=="기타": continue
        item=DB.get(cat,{"title":f"{cat} 비판 보도","points":["모니터링 강화 필요","맞춤 대응 메시지 개발"]}).copy()
        item["dots"]=min(5,max(2,cnt2//max(1,len(neg)//10)+2)); result.append(item)
        if len(result)==5: break
    defs=[{"title":"커뮤니케이션 체계 미흡","points":["위기 시 신속 대응 부족","공식 채널 속도 개선 필요"],"dots":3},{"title":"사회적 책임 이행 부족","points":["CSR 기대치 미충족","이해관계자 소통 강화"],"dots":2}]
    while len(result)<5: result.append(defs.pop(0))
    return result[:5]
def gen_actions(arts,kw):
    neg=[a for a in arts if a["감성"]=="부정"]
    nm=Counter([a["매체"] for a in neg]).most_common(3)
    t2=", ".join([m for m,_ in nm[:2]]) if nm else "주요 매체"
    return [{"title":"핵심 성과 홍보","points":["성과 중심 보도자료 배포","긍정 매체 관계 강화"],"dots":5},{"title":"부정 이슈 팩트시트","points":["키워드별 Q&A 준비","공식 해명 지속 업데이트"],"dots":5},{"title":"부정 매체 개별 대응","points":[f"{t2} 전담 대응 운영","현장 취재 지원"],"dots":4},{"title":"SNS·디지털 대응","points":["포털 댓글 모니터링","공식 SNS 긍정 콘텐츠"],"dots":4},{"title":"전문가 활용","points":["전문가 코멘트·기고 추진","정책 성과 평가 자료"],"dots":3}]
def calc_pr_risk(neg_n,total,neg_kws,crisis_found,top_neg_media):
    s=0; neg_r=neg_n/total*100 if total>0 else 0
    s+=min(40,neg_r*0.8)
    if crisis_found: s+=20
    s+=min(20,len(neg_kws)*4)
    sa=[m for m in top_neg_media if MEDIA_GRADE.get(m,{}).get('grade','') in ['S','A']]
    s+=min(20,len(sa)*7)
    s=min(100,round(s,1))
    if s>=70: return s,"HIGH","#C62828"
    elif s>=40: return s,"MEDIUM","#E65100"
    return s,"LOW","#2E7D32"

# ── 시장 데이터 ───────────────────────────────────────
@st.cache_data(ttl=1800)
def get_market_data():
    d={"kospi":"—","kospi_c":"","kospi_p":"","kospi_up":True,"kosdaq":"—","kosdaq_c":"","kosdaq_p":"","kosdaq_up":True,"kepco_k":"—","kepco_kc":"","kepco_k_up":True,"kepco_u":"—","kepco_uc":"","kepco_u_up":True,"usd_krw":"—","usd_c":"","usd_up":True,"oil":"—","oil_c":"","oil_up":True,"smp_avg":"—","smp_h":"—","smp_l":"—","updated":datetime.now().strftime("%Y.%m.%d %H:%M")}
    if YF_OK:
        for sym,key in {"^KS11":"kospi","^KQ11":"kosdaq","015760.KS":"kepco_k","KEP":"kepco_u","USDKRW=X":"usd","BZ=F":"oil"}.items():
            try:
                h=yf.Ticker(sym).history(period="2d")
                if h.empty: continue
                cur=float(h["Close"].iloc[-1]); prev=float(h["Close"].iloc[-2]) if len(h)>=2 else cur
                chg=cur-prev; pct=chg/prev*100 if prev else 0
                arr="▲" if chg>=0 else "▼"; up=(chg>=0)
                if key=="kospi": d.update({"kospi":f"{cur:,.2f}","kospi_c":f"{arr}{abs(chg):,.2f}","kospi_p":f"{pct:+.2f}%","kospi_up":up})
                elif key=="kosdaq": d.update({"kosdaq":f"{cur:,.2f}","kosdaq_c":f"{arr}{abs(chg):,.2f}","kosdaq_p":f"{pct:+.2f}%","kosdaq_up":up})
                elif key=="kepco_k": d.update({"kepco_k":f"{cur:,}원","kepco_kc":f"{arr}{abs(chg):,.0f}","kepco_k_up":up})
                elif key=="kepco_u": d.update({"kepco_u":f"{cur:.2f}USD","kepco_uc":f"{arr}{abs(chg):.2f}","kepco_u_up":up})
                elif key=="usd": d.update({"usd_krw":f"{cur:,.2f}","usd_c":f"{arr}{abs(chg):,.2f}","usd_up":up})
                elif key=="oil": d.update({"oil":f"{cur:.2f}","oil_c":f"{arr}{abs(chg):.2f}","oil_up":up})
            except: pass
    try:
        r=requests.get("https://new.kpx.or.kr/powerSource/getSmpCurrentDay.do",headers={"User-Agent":"Mozilla/5.0","Referer":"https://new.kpx.or.kr/"},params={"area":"1","yyyymmdd":datetime.now().strftime("%Y%m%d")},timeout=5)
        if r.status_code==200:
            vals=[float(x.get("smp",0)) for x in (r.json().get("list",[]) or r.json().get("data",[])) if x.get("smp")]
            if vals: d.update({"smp_avg":f"{sum(vals)/len(vals):.2f}","smp_h":f"{max(vals):.2f}","smp_l":f"{min(vals):.2f}"})
    except: pass
    return d

def mhdr(d):
    def cs(v,up): c="#C62828" if up else "#1565C0"; return f"<span style='color:{c};font-size:10px;font-weight:600;'>{v}</span>"
    smp="" if d["smp_avg"]=="—" else f"<div style='border-left:1px solid #ddd;padding-left:10px;margin-left:8px;'><div style='font-size:8px;color:#888;font-weight:700;'>SMP육지</div><div style='font-size:12px;font-weight:700;color:#003366;'>{d['smp_avg']}</div><div style='font-size:8px;color:#777;'>고{d['smp_h']}/저{d['smp_l']}</div></div>"
    return f"""<div style='background:white;border:1px solid #ddd;border-radius:5px;padding:7px 14px;margin-bottom:8px;display:flex;align-items:center;flex-wrap:wrap;gap:3px;font-family:sans-serif;'>
<div style='margin-right:10px;'><div style='font-size:8px;color:#888;font-weight:700;'>코스피</div><div style='font-size:13px;font-weight:700;color:#003366;'>{d['kospi']}</div><div>{cs(d['kospi_c']+" "+d['kospi_p'],d['kospi_up'])}</div></div>
<div style='margin-right:12px;border-left:1px solid #eee;padding-left:10px;'><div style='font-size:8px;color:#888;font-weight:700;'>코스닥</div><div style='font-size:13px;font-weight:700;color:#003366;'>{d['kosdaq']}</div><div>{cs(d['kosdaq_c']+" "+d['kosdaq_p'],d['kosdaq_up'])}</div></div>
<div style='border-left:2px solid #003366;height:30px;margin:0 10px;'></div>
<div style='margin-right:4px;font-size:8px;color:#003366;font-weight:700;'>⚡ KEPCO</div>
<div style='margin-right:10px;'><div style='font-size:8px;color:#888;'>KOSPI</div><div style='font-size:12px;font-weight:700;color:#003366;'>{d['kepco_k']}</div><div>{cs(d['kepco_kc'],d['kepco_k_up'])}</div></div>
<div style='margin-right:12px;border-left:1px solid #eee;padding-left:10px;'><div style='font-size:8px;color:#888;'>NYSE</div><div style='font-size:12px;font-weight:700;color:#003366;'>{d['kepco_u']}</div><div>{cs(d['kepco_uc'],d['kepco_u_up'])}</div></div>
<div style='border-left:2px solid #ddd;height:30px;margin:0 10px;'></div>
<div style='margin-right:10px;'><div style='font-size:8px;color:#888;font-weight:700;'>USD/KRW</div><div style='font-size:12px;font-weight:700;color:#333;'>{d['usd_krw']}</div><div>{cs(d['usd_c'],d['usd_up'])}</div></div>
<div style='border-left:1px solid #eee;padding-left:10px;margin-right:10px;'><div style='font-size:8px;color:#888;font-weight:700;'>두바이유($/bbl)</div><div style='font-size:12px;font-weight:700;color:#333;'>{d['oil']}</div><div>{cs(d['oil_c'],d['oil_up'])}</div></div>
{smp}<div style='margin-left:auto;font-size:8px;color:#ccc;'>{d['updated']}</div></div>"""

# ── 차트 ──────────────────────────────────────────────
def mini_chart_config(): return {'displayModeBar':False,'staticPlot':False}

def plot_buzz(df):
    daily=df.groupby('일자').size().reset_index(name='건수')
    daily['dt']=pd.to_datetime(daily['일자'])
    by_sent=df.groupby(['일자','감성']).size().unstack(fill_value=0)
    fig=go.Figure()
    for sent,color in [('부정','#FFCDD2'),('중립','#E0E0E0'),('긍정','#BBDEFB')]:
        if sent in by_sent.columns:
            y=by_sent[sent].reindex(daily['일자'],fill_value=0).values
            fig.add_trace(go.Bar(x=daily['dt'],y=y,name=sent,marker_color=color,hovertemplate=f'{sent}: %{{y}}건<extra></extra>'))
    fig.add_trace(go.Scatter(x=daily['dt'],y=daily['건수'],mode='lines+markers',name='전체',line=dict(color='#003366',width=2),marker=dict(size=5,color='white',line=dict(width=2,color='#003366')),hovertemplate='%{x|%Y-%m-%d}<br>전체: <b>%{y}건</b><extra></extra>'))
    fig.update_layout(barmode='stack',plot_bgcolor='white',paper_bgcolor='white',font=dict(family=FONT_KR,size=10),margin=dict(l=40,r=10,t=25,b=35),height=190,hovermode='x unified',showlegend=True,legend=dict(orientation='h',y=1.12,x=1,xanchor='right',font=dict(size=10)),xaxis=dict(tickformat='%m/%d',showgrid=False,tickangle=-30),yaxis=dict(showgrid=True,gridcolor='#f5f5f5',rangemode='tozero'))
    return fig

def plot_kw_trend(df,kw,mode='daily'):
    mask=df['헤드라인'].str.contains(kw,na=False,regex=False)|df['요약'].str.contains(kw,na=False,regex=False)
    kdf=df[mask].copy()
    if kdf.empty: return None,[]
    sel_days=[]
    color_map={'부정':'#C62828','긍정':'#1565C0','중립':'#555555'}
    if mode=='daily':
        all_dates=sorted(df['일자'].unique())
        grouped=kdf.groupby(['일자','감성']).size().unstack(fill_value=0).reindex(all_dates,fill_value=0)
        x=[pd.to_datetime(d) for d in grouped.index]; xfmt='%m/%d'
    elif mode=='monthly':
        grouped=kdf.groupby(['월','감성']).size().unstack(fill_value=0)
        x=grouped.index.tolist(); xfmt=None
    else:
        cutoff=(datetime.now()-timedelta(days=7)).strftime('%Y-%m-%d')
        avail=sorted(kdf['일자'].unique())
        sel_days=[d for d in avail if d>=cutoff] or avail[-7:]
        kdf2=kdf[kdf['일자'].isin(sel_days)].copy()
        kdf2['시간_int']=pd.to_numeric(kdf2['시간'],errors='coerce').fillna(0).astype(int)
        grouped=kdf2.groupby(['시간_int','감성']).size().unstack(fill_value=0).reindex(range(24),fill_value=0)
        x=list(range(24)); xfmt=None
    fig=go.Figure()
    for sent,color in color_map.items():
        y=grouped[sent].tolist() if sent in grouped.columns else [0]*len(x)
        fig.add_trace(go.Scatter(x=x,y=y,mode='lines+markers',name=sent,line=dict(color=color,width=2),marker=dict(size=4),hovertemplate=f'<b>{sent}</b> %{{y}}건<extra></extra>'))
    mode_label={'daily':'일별','monthly':'월별','hourly':'시간대별(최근1주)'}
    fig.update_layout(title=dict(text=f'<b>「{kw}」 {mode_label.get(mode,"")}  추이</b>',font=dict(size=12,color='#003366',family=FONT_KR)),plot_bgcolor='white',paper_bgcolor='white',font=dict(family=FONT_KR,size=10),margin=dict(l=40,r=10,t=35,b=35),height=220,hovermode='x unified',legend=dict(orientation='h',y=1.15,x=1,xanchor='right',font=dict(size=10)),xaxis=dict(showgrid=True,gridcolor='#f5f5f5',tickangle=-30,tickformat=xfmt,tickmode='linear' if mode=='hourly' else 'auto',dtick=2 if mode=='hourly' else None),yaxis=dict(showgrid=True,gridcolor='#f5f5f5',rangemode='tozero'))
    return fig,sel_days

def plot_crisis(df,ck):
    mask=df['헤드라인'].str.contains(ck,na=False,regex=False)|df['요약'].str.contains(ck,na=False,regex=False)
    ckdf=df[mask].copy()
    if ckdf.empty: return None,0,0
    total=len(ckdf); neg=int(ckdf['감성'].value_counts().get('부정',0))
    all_dates=sorted(df['일자'].unique())
    dt=ckdf.groupby('일자').size().reindex(all_dates,fill_value=0)
    dn=ckdf[ckdf['감성']=='부정'].groupby('일자').size().reindex(all_dates,fill_value=0)
    x=[pd.to_datetime(d) for d in all_dates]
    fig=go.Figure()
    fig.add_trace(go.Bar(x=x,y=dt.values,name='전체',marker_color='#E3F2FD',hovertemplate='전체 %{y}건<extra></extra>'))
    fig.add_trace(go.Scatter(x=x,y=dn.values,mode='lines+markers',name='부정',line=dict(color='#C62828',width=2),marker=dict(size=4),hovertemplate='부정 %{y}건<extra></extra>'))
    fig.update_layout(title=dict(text=f'<b>「{ck}」</b> <span style="color:#C62828;font-size:11px;">총 {total}건 | 부정 {neg}건</span>',font=dict(size=12,color='#003366',family=FONT_KR)),barmode='overlay',plot_bgcolor='white',paper_bgcolor='white',font=dict(family=FONT_KR,size=10),margin=dict(l=40,r=10,t=35,b=35),height=200,hovermode='x unified',legend=dict(orientation='h',y=1.15,x=1,xanchor='right',font=dict(size=10)),xaxis=dict(tickformat='%m/%d',showgrid=False,tickangle=-30),yaxis=dict(showgrid=True,gridcolor='#f5f5f5',rangemode='tozero'))
    return fig,total,neg

def plot_heatmap(df):
    top_m=df["매체"].value_counts().head(8).index.tolist()
    top_c=[c for c in TOPIC_GROUPS if c in df["카테고리"].values][:7]
    if not top_m or not top_c: return None
    z=np.array([[round(len(df[(df["매체"]==m)&(df["카테고리"]==cat)&(df["감성"]=="부정")])/max(1,len(df[(df["매체"]==m)&(df["카테고리"]==cat)]))*100,0) for cat in top_c] for m in top_m])
    fig=go.Figure(data=go.Heatmap(z=z,x=top_c,y=top_m,colorscale=[[0,'#F1F8E9'],[0.5,'#FFF9C4'],[1,'#B71C1C']],zmin=0,zmax=100,text=[[f"{v:.0f}%" for v in row] for row in z],texttemplate="%{text}",textfont={"size":10},hovertemplate='%{y} × %{x}<br>부정률: <b>%{z:.0f}%</b><extra></extra>'))
    fig.update_layout(xaxis=dict(tickangle=-30,side='bottom'),yaxis=dict(autorange='reversed'),plot_bgcolor='white',paper_bgcolor='white',font=dict(family=FONT_KR,size=10),margin=dict(l=90,r=10,t=15,b=60),height=290)
    return fig

def plot_radar(items,color,title):
    cats=[c["title"][:7] for c in items]; vals=[c["dots"] for c in items]
    N=len(cats); a=[n/float(N)*2*np.pi for n in range(N)]; a+=a[:1]; v=vals+vals[:1]
    fig,ax=plt.subplots(figsize=(3.2,3.2),subplot_kw=dict(polar=True)); fig.patch.set_facecolor('white')
    ax.plot(a,v,'o-',color=color,lw=1.5); ax.fill(a,v,alpha=0.15,color=color)
    ax.set_xticks(a[:-1]); ax.set_xticklabels(cats,fontsize=7); ax.set_ylim(0,5)
    ax.tick_params(pad=3); ax.set_facecolor('white')
    plt.tight_layout(pad=0.5)
    return fig

# ── Word 전체 보고서 ──────────────────────────────────
def make_full_word(cd):
    doc=Document(); label=cd['label']; period_str=cd['period_str']; df=cd['df']; total=cd['total']
    pos_n=cd['pos_n']; neg_n=cd['neg_n']; neu_n=cd['neu_n']; neg_rate=neg_n/total*100
    pr_s=cd.get('pr_score',0); pr_l=cd.get('pr_lvl','—')
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("한국전력 언론보도 유형분석 보고서"); r.bold=True; r.font.size=Pt(18); r.font.color.rgb=RGBColor(0,51,102)
    p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"{label}  |  {period_str}  |  {datetime.now().strftime('%Y년 %m월 %d일')}")
    doc.add_paragraph()
    def hd(txt,lv=1):
        h=doc.add_heading(txt,level=lv); h.runs[0].font.color.rgb=RGBColor(0,51,102); return h
    hd("00. 종합 결론 및 제언")
    doc.add_paragraph(cd['insights_text'])
    doc.add_paragraph(f"PR 리스크 스코어: {pr_s}점 / {pr_l}")
    doc.add_paragraph()
    hd("01. Executive Summary")
    tone=("부정 우세" if neg_n>pos_n*1.5 else "긍정 우세" if pos_n>neg_n*1.5 else "균형")
    for ln in [f"기간: {period_str} | 총 {total}건 | PR리스크 {pr_s}점 ({pr_l})",f"논조: {tone} — 부정 {neg_n}건({neg_rate:.1f}%) | 긍정 {pos_n}건({pos_n/total*100:.1f}%) | 중립 {neu_n}건({neu_n/total*100:.1f}%)",f"핵심 비판: {cd['top_neg_cat']} | 부정1위 키워드: {cd['top_neg_kw'] or '—'}",f"주요 매체: {cd['top3_media']}"]:
        doc.add_paragraph(ln)
    doc.add_paragraph()
    hd("02. 논조별 키워드 TOP5")
    for label_,kws in [("부정",cd['neg_kws']),("긍정",cd['pos_kws']),("중립",cd['neu_kws'])]:
        p3=doc.add_paragraph(); p3.add_run(f"[{label_}] ").bold=True; p3.add_run(", ".join([f"{k}({v}회)" for k,v in kws]))
    doc.add_paragraph()
    hd("03. 주요 비판 포인트")
    for i,c in enumerate(cd['criticisms'],1):
        p4=doc.add_paragraph(); p4.add_run(f"{i}. {c['title']}  ({'●'*c['dots']}{'○'*(5-c['dots'])})").bold=True
        for pt in c["points"]: doc.add_paragraph(f"   · {pt}",style='List Bullet')
    doc.add_paragraph()
    hd("04. 언론 대응 우선순위")
    for i,a in enumerate(cd['actions'],1):
        p5=doc.add_paragraph(); p5.add_run(f"{i}. {a['title']}  ({'●'*a['dots']}{'○'*(5-a['dots'])})").bold=True
        for pt in a["points"]: doc.add_paragraph(f"   · {pt}",style='List Bullet')
    doc.add_paragraph()
    hd("05. 매체별 논조 분포")
    tbl=doc.add_table(rows=1,cols=5); tbl.style='Table Grid'
    for i,h_txt in enumerate(["매체","등급","열독률","부정","긍정"]): tbl.rows[0].cells[i].text=h_txt
    for mname in df["매체"].value_counts().head(12).index:
        gi=MEDIA_GRADE.get(mname,{}); cells=tbl.add_row().cells
        cells[0].text=mname; cells[1].text=gi.get("grade","—"); cells[2].text=f"{gi.get('rate','')}%"
        cells[3].text=str(int(df[(df["매체"]==mname)&(df["감성"]=="부정")].shape[0]))
        cells[4].text=str(int(df[(df["매체"]==mname)&(df["감성"]=="긍정")].shape[0]))
    doc.add_paragraph()
    hd("06. 주간 PR 브리핑")
    top5=df[df['감성']=='부정'].sort_values('일자',ascending=False).head(5)
    lines="\n".join([f"  · [{r2['일자']}] {r2['매체']} — {r2['헤드라인']}" for _,r2 in top5.iterrows()])
    doc.add_paragraph(f"[한국전력 PR 브리핑] {period_str}\n총 {total}건 | 부정 {neg_n}건({neg_rate:.0f}%) | PR리스크 {pr_s}점/{pr_l}\n핵심: {cd['top_neg_cat']} | 키워드: {', '.join([k for k,_ in cd['neg_kws'][:3]])}\n\n주요 부정 기사:\n{lines}")
    doc.add_paragraph()
    hd("07. 기사 전체 목록")
    tbl2=doc.add_table(rows=1,cols=5); tbl2.style='Table Grid'
    for i,h_txt in enumerate(["번호","일자","매체","헤드라인","논조"]): tbl2.rows[0].cells[i].text=h_txt
    for idx,row in enumerate(df.sort_values("일자",ascending=False).to_dict("records"),1):
        cells2=tbl2.add_row().cells; cells2[0].text=str(idx); cells2[1].text=str(row["일자"]); cells2[2].text=str(row["매체"]); cells2[3].text=str(row["헤드라인"]); cells2[4].text=str(row["감성"])
    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf

def copy_link_btn():
    components.html("""<button id="cpbtn" onclick="(function(){var u=window.parent.location.href;if(navigator.clipboard){navigator.clipboard.writeText(u).then(function(){document.getElementById('cpbtn').innerText='✅ 복사됨!';document.getElementById('cpbtn').style.background='#2E7D32';setTimeout(function(){document.getElementById('cpbtn').innerText='🔗 링크 복사';document.getElementById('cpbtn').style.background='#003366';},2000);})}else{var el=document.createElement('input');el.value=u;document.body.appendChild(el);el.select();document.execCommand('copy');document.body.removeChild(el);document.getElementById('cpbtn').innerText='✅ 복사됨!';setTimeout(function(){document.getElementById('cpbtn').innerText='🔗 링크 복사';},2000);}})();" style="background:#003366;color:white;border:none;padding:8px 16px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;font-family:sans-serif;width:100%;">🔗 링크 복사</button>""",height=40)

def divider(n): st.markdown(f"<div style='font-size:11px;font-weight:700;color:#003366;letter-spacing:.6px;border-bottom:1.5px solid #003366;padding-bottom:3px;margin:14px 0 7px;'>{n}</div>",unsafe_allow_html=True)

# ══ 보고서 렌더링 ══════════════════════════════════════
def render_report(cd):
    label=cd['label']; period_str=cd['period_str']; df=cd['df']
    total=cd['total']; pos_n=cd['pos_n']; neg_n=cd['neg_n']; neu_n=cd['neu_n']
    neg_kws=cd['neg_kws']; neu_kws=cd['neu_kws']; pos_kws=cd['pos_kws']
    top_neg_kw=cd['top_neg_kw']; criticisms=cd['criticisms']; actions=cd['actions']
    insights_text=cd['insights_text']; top_neg_cat=cd['top_neg_cat']; top_pos_cat=cd['top_pos_cat']
    top3_media=cd['top3_media']; trend_txt=cd['trend_txt']; crisis_kws=cd['crisis_kws']
    pr_s=cd.get('pr_score',0); pr_l=cd.get('pr_lvl','—'); pr_c=cd.get('pr_color','#888')
    neg_rate=neg_n/total*100; pos_rate=pos_n/total*100
    tone_sym=("🔴" if neg_n>pos_n*1.5 else "🟢" if pos_n>neg_n*1.5 else "🟡")
    tone_txt=("부정 우세" if neg_n>pos_n*1.5 else "긍정 우세" if pos_n>neg_n*1.5 else "균형")
    neg_kw_str=", ".join([f'{k}({v}회)' for k,v in neg_kws[:3]]) if neg_kws else "없음"
    neg_media_top=df[df['감성']=='부정']['매체'].value_counts().head(3)
    top_neg_m=", ".join([f"{m}({n}건)" for m,n in neg_media_top.items()]) if not neg_media_top.empty else "해당없음"

    # ─ 헤더 바 ─
    hc1,hc2=st.columns([5,1])
    with hc1:
        st.markdown(f"""<div style='background:#003366;color:white;padding:10px 16px;border-radius:5px;'>
        <span style='font-size:15px;font-weight:700;'>{label}</span>
        <span style='font-size:11px;opacity:.75;margin-left:10px;'>{period_str} | 총 {total}건</span>
        <span style='float:right;font-size:20px;font-weight:900;color:{pr_c};'>{pr_s}점</span>
        <span style='float:right;font-size:10px;color:rgba(255,255,255,.7);margin-right:6px;margin-top:4px;'>PR리스크</span>
        </div>""",unsafe_allow_html=True)
    with hc2:
        wb=make_full_word(cd)
        st.download_button("📄 보고서 워드",data=wb,file_name=f"KEPCO_{label}_{datetime.now().strftime('%Y%m%d')}.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True,key=f"wd_{label}")

    # ═══════════════════════════════════════════════════
    # BLOCK A: 결론 + KPI (한 줄 압축)
    # ═══════════════════════════════════════════════════
    divider("00 · 종합 결론 및 제언")

    # KPI 가로 6칸
    k1,k2,k3,k4,k5,k6=st.columns(6)
    kpi_items=[
        (k1,str(total),"총 기사","#003366",trend_txt[:18]),
        (k2,f"{neg_n}건","부정","#C62828",f"{neg_rate:.0f}%  {top_neg_cat[:6]}"),
        (k3,f"{pos_n}건","긍정","#1565C0",f"{pos_rate:.0f}%  {top_pos_cat[:6]}"),
        (k4,f"{neu_n}건","중립","#555",f"{neu_n/total*100:.0f}%"),
        (k5,tone_sym,"논조","#333",tone_txt),
        (k6,f"{pr_s}","PR리스크",pr_c,f"{pr_l}  /100점"),
    ]
    for col,val,lbl,color,sub in kpi_items:
        col.markdown(f"""<div style='background:white;border:1px solid #e8e8e8;border-top:3px solid {color};border-radius:4px;padding:8px 6px;text-align:center;'>
        <div style='font-size:19px;font-weight:700;color:{color};line-height:1.1;'>{val}</div>
        <div style='font-size:9px;font-weight:700;color:#999;letter-spacing:.4px;margin-top:2px;'>{lbl}</div>
        <div style='font-size:9px;color:#bbb;margin-top:1px;white-space:nowrap;overflow:hidden;'>{sub}</div></div>""",unsafe_allow_html=True)

    # PR 리스크 게이지 + 결론 텍스트
    g1,g2=st.columns([1,3])
    with g1:
        # 게이지 차트
        fig_g=go.Figure(go.Indicator(
            mode="gauge+number",value=pr_s,
            number={"suffix":"점","font":{"size":22,"color":pr_c,"family":FONT_KR}},
            gauge={"axis":{"range":[0,100],"tickwidth":1,"tickcolor":"#ccc","tickfont":{"size":9}},"bar":{"color":pr_c,"thickness":0.25},"bgcolor":"#f5f5f5","borderwidth":0,"steps":[{"range":[0,40],"color":"#E8F5E9"},{"range":[40,70],"color":"#FFF8E1"},{"range":[70,100],"color":"#FFEBEE"}],"threshold":{"line":{"color":pr_c,"width":3},"thickness":0.75,"value":pr_s}}))
        fig_g.update_layout(height=140,margin=dict(l=20,r=20,t=10,b=10),paper_bgcolor='white',font=dict(family=FONT_KR))
        st.plotly_chart(fig_g,use_container_width=True,config=mini_chart_config())
    with g2:
        neg_kw_top3=[k for k,_ in neg_kws[:3]]
        act_top2=[a['title'] for a in actions[:2]]
        st.markdown(f"""<div style='background:#F8F9FA;border-left:3px solid #003366;border-radius:0 4px 4px 0;padding:10px 14px;font-size:12px;line-height:1.8;'>
        <span style='color:#003366;font-weight:700;'>핵심 비판:</span> {top_neg_cat} &nbsp;|&nbsp; <span style='color:#003366;font-weight:700;'>부정 키워드:</span> {neg_kw_str}<br>
        <span style='color:#003366;font-weight:700;'>부정 집중 매체:</span> {top_neg_m}<br>
        <span style='color:#003366;font-weight:700;'>결론:</span> {insights_text[:120]}…<br>
        <span style='color:#C62828;font-weight:700;'>즉각 대응:</span> {act_top2[0]} &nbsp;/&nbsp; {act_top2[1]}
        </div>""",unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════
    # BLOCK B: 논조 파이 + 버즈량 + 매체 테이블 (3분할)
    # ═══════════════════════════════════════════════════
    divider("01 · 논조 분포  ·  02 · 버즈량 추이  ·  03 · 매체별 논조")
    b1,b2,b3=st.columns([1,2,2])
    with b1:
        fig_pie,ax=plt.subplots(figsize=(3,3)); fig_pie.patch.set_facecolor('white')
        pd2=[(pos_n,"긍정","#1565C0"),(neu_n,"중립","#9E9E9E"),(neg_n,"부정","#C62828")]
        pd2=[x for x in pd2 if x[0]>0]
        _,_,auts=ax.pie([x[0] for x in pd2],labels=[f"{x[1]}\n{x[0]}건" for x in pd2],colors=[x[2] for x in pd2],autopct='%1.0f%%',startangle=90,wedgeprops=dict(width=0.55,edgecolor='white',linewidth=1.5),pctdistance=0.75,labeldistance=1.15)
        for at in auts: at.set_fontsize(8); at.set_color('white'); at.set_fontweight('bold')
        plt.tight_layout(pad=0.3); st.pyplot(fig_pie,use_container_width=True); plt.close()
    with b2:
        st.plotly_chart(plot_buzz(df),use_container_width=True,config=mini_chart_config())
    with b3:
        rows_m=""
        for mname in df["매체"].value_counts().head(8).index:
            gi=MEDIA_GRADE.get(mname,{}); grade=gi.get("grade","—"); rate=gi.get("rate",""); gc=GRADE_COLOR.get(grade,"#ccc")
            n_neg2=int(df[(df["매체"]==mname)&(df["감성"]=="부정")].shape[0])
            n_pos2=int(df[(df["매체"]==mname)&(df["감성"]=="긍정")].shape[0])
            n_tot2=int(df[df["매체"]==mname].shape[0])
            bar=int(n_neg2/n_tot2*12) if n_tot2>0 else 0
            rows_m+=f"<tr><td style='font-size:11px;'>{mname} <span style='background:{gc};color:white;padding:0px 3px;border-radius:2px;font-size:8px;font-weight:700;'>{grade}</span></td><td style='color:#888;font-size:9px;text-align:right;'>{rate}%</td><td style='color:#C62828;font-size:11px;font-weight:600;text-align:center;'>{n_neg2}</td><td style='color:#1565C0;font-size:11px;text-align:center;'>{n_pos2}</td><td style='font-size:10px;'><span style='color:#C62828;'>{'■'*bar}</span><span style='color:#eee;'>{'■'*(12-bar)}</span></td></tr>"
        st.markdown(f"""<table style='width:100%;border-collapse:collapse;margin-top:4px;'>
        <tr style='background:#003366;color:white;font-size:10px;'><th style='padding:4px 6px;text-align:left;'>매체</th><th>열독률</th><th>부정</th><th>긍정</th><th>부정비중</th></tr>{rows_m}</table>""",unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════
    # BLOCK C: 키워드 TOP5 + 추세 탭
    # ═══════════════════════════════════════════════════
    divider("04 · 논조별 키워드 노출량 — 키워드 선택 시 추세 분석")
    kc1,kc2,kc3=st.columns(3)
    def kw_block(kws,accent,title):
        rows="".join([f"<div style='display:flex;justify-content:space-between;align-items:center;padding:3px 0;border-bottom:1px solid #f5f5f5;'><span style='font-size:11px;'>{kw}</span><span style='background:{accent};color:white;padding:1px 8px;border-radius:10px;font-size:10px;font-weight:700;'>{cnt}회</span></div>" for kw,cnt in kws] or "<div style='font-size:10px;color:#ccc;'>해당없음</div>")
        return f"<div style='background:white;border:1px solid #e8e8e8;border-top:2px solid {accent};border-radius:4px;padding:8px 10px;'><div style='font-size:10px;font-weight:700;color:{accent};margin-bottom:5px;'>{title}</div>{rows}</div>"
    with kc1: st.markdown(kw_block(neg_kws,"#C62828","🔴 부정 키워드 TOP5"),unsafe_allow_html=True)
    with kc2: st.markdown(kw_block(neu_kws,"#333","⬛ 중립 키워드 TOP5"),unsafe_allow_html=True)
    with kc3: st.markdown(kw_block(pos_kws,"#1565C0","🔵 긍정 키워드 TOP5"),unsafe_allow_html=True)

    all_kws=[(f"🔴 {k}",k) for k,_ in neg_kws]+[(f"🔵 {k}",k) for k,_ in pos_kws]+[(f"⬛ {k}",k) for k,_ in neu_kws]
    if all_kws:
        sel_disp=st.selectbox("📌 키워드 선택 → 추세 분석",[d for d,_ in all_kws],key=f"kwsel_{label}",label_visibility="collapsed")
        sel_kw=dict(all_kws)[sel_disp]
        t1,t2,t3=st.tabs(["📅 월별","📆 일자별 (전체 기간)","🕐 시간대별 (최근 1주일)"])
        with t1:
            f,_=plot_kw_trend(df,sel_kw,'monthly')
            if f: st.plotly_chart(f,use_container_width=True,config=mini_chart_config())
            else: st.caption("데이터 없음")
        with t2:
            f,_=plot_kw_trend(df,sel_kw,'daily')
            if f: st.plotly_chart(f,use_container_width=True,config=mini_chart_config())
            else: st.caption("데이터 없음")
        with t3:
            f,sd=plot_kw_trend(df,sel_kw,'hourly')
            if f:
                if sd: st.caption(f"최근 7일({', '.join(sd[:3])}{'...' if len(sd)>3 else ''}) 기준")
                st.plotly_chart(f,use_container_width=True,config=mini_chart_config())
            else: st.caption("데이터 없음")

    # ═══════════════════════════════════════════════════
    # BLOCK D: 위기 키워드 + 히트맵 (2분할)
    # ═══════════════════════════════════════════════════
    divider("05 · 위기관리 키워드  ·  06 · 매체×이슈 부정 보도율")
    d1,d2=st.columns([3,2])
    with d1:
        found_c=False
        for ck in crisis_kws:
            fg,ct,cn=plot_crisis(df,ck)
            if fg:
                found_c=True
                st.plotly_chart(fg,use_container_width=True,config=mini_chart_config())
        if not found_c: st.caption("위기 키워드가 수집 기사에 없습니다.")
    with d2:
        fg_hm=plot_heatmap(df)
        if fg_hm:
            st.caption("셀 값 = 해당 매체의 카테고리별 부정 보도율(%)")
            st.plotly_chart(fg_hm,use_container_width=True,config=mini_chart_config())
        else: st.caption("히트맵: 데이터 부족")

    # ═══════════════════════════════════════════════════
    # BLOCK E: 비판 포인트 + 대응 우선순위 (레이더 + 텍스트)
    # ═══════════════════════════════════════════════════
    divider("07 · 주요 비판 포인트  ·  08 · 언론 대응 우선순위")
    e1,e2,e3,e4=st.columns([1,2,1,2])
    with e1:
        fig_r=plot_radar(criticisms,"#C62828","비판 강도")
        st.pyplot(fig_r,use_container_width=True); plt.close()
    with e2:
        html_c=""
        for i,c in enumerate(criticisms,1):
            dots="●"*c["dots"]+"○"*(5-c["dots"])
            pts=" / ".join(c["points"])
            html_c+=f"<div style='border-left:2px solid #C62828;padding:4px 8px;margin-bottom:4px;background:white;'><div style='display:flex;justify-content:space-between;'><span style='font-size:11px;font-weight:700;color:#003366;'>{i}. {c['title']}</span><span style='color:#C62828;font-size:10px;letter-spacing:1px;'>{dots}</span></div><div style='font-size:10px;color:#666;'>{pts}</div></div>"
        st.markdown(html_c,unsafe_allow_html=True)
    with e3:
        fig_r2=plot_radar(actions,"#003366","대응 우선순위")
        st.pyplot(fig_r2,use_container_width=True); plt.close()
    with e4:
        html_a=""
        for i,a in enumerate(actions,1):
            dots="●"*a["dots"]+"○"*(5-a["dots"])
            pts=" / ".join(a["points"])
            html_a+=f"<div style='border-left:2px solid #003366;padding:4px 8px;margin-bottom:4px;background:white;'><div style='display:flex;justify-content:space-between;'><span style='font-size:11px;font-weight:700;color:#003366;'>{i}. {a['title']}</span><span style='color:#003366;font-size:10px;letter-spacing:1px;'>{dots}</span></div><div style='font-size:10px;color:#666;'>{pts}</div></div>"
        st.markdown(html_a,unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════
    # BLOCK F: 주간 브리핑 원페이저
    # ═══════════════════════════════════════════════════
    divider("09 · 주간 PR 브리핑 원페이저")
    top5=df[df['감성']=='부정'].sort_values('일자',ascending=False).head(5)
    lines="\n".join([f"  · [{r2['일자']}] {r2['매체']} — {r2['헤드라인']}" for _,r2 in top5.iterrows()])
    briefing=f"""[한국전력 언론 주간 브리핑]  {period_str}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
총 {total}건 | 부정 {neg_n}건({neg_rate:.0f}%) | 긍정 {pos_n}건({pos_rate:.0f}%) | PR리스크 {pr_s}점/{pr_l}
핵심 이슈: {top_neg_cat} | 부정 키워드: {', '.join([k for k,_ in neg_kws[:3]])} | 집중 매체: {top_neg_m}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🔴 주요 부정 기사 TOP5:
{lines}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ 대응: {actions[0]['title']} / {actions[1]['title']}"""
    st.text_area("",briefing,height=200,key=f"brief_{label}",label_visibility="collapsed")

    # ═══════════════════════════════════════════════════
    # BLOCK G: Raw Data
    # ═══════════════════════════════════════════════════
    divider("10 · 기사 목록")
    gc1,gc2=st.columns(2)
    with gc1: df_f=st.selectbox("일자",["전체"]+sorted(df["일자"].unique().tolist(),reverse=True),key=f"dff_{label}",label_visibility="collapsed")
    with gc2: cf_f=st.selectbox("카테고리",["전체"]+sorted(df["카테고리"].unique().tolist()),key=f"cff_{label}",label_visibility="collapsed")
    fdf=df.copy()
    if df_f!="전체": fdf=fdf[fdf["일자"]==df_f]
    if cf_f!="전체": fdf=fdf[fdf["카테고리"]==cf_f]
    fdf=fdf.sort_values("일자",ascending=False).reset_index(drop=True)
    sk=f"s_{label}"
    if sk not in st.session_state: st.session_state[sk]=30
    ddf=fdf.iloc[:st.session_state[sk]]
    rh=""
    for i,row in enumerate(ddf.to_dict("records"),1):
        bg={"긍정":"<span style='background:#E3F2FD;color:#1565C0;padding:0px 6px;border-radius:8px;font-size:10px;font-weight:600;'>긍정</span>","부정":"<span style='background:#FFEBEE;color:#C62828;padding:0px 6px;border-radius:8px;font-size:10px;font-weight:600;'>부정</span>","중립":"<span style='background:#F5F5F5;color:#555;padding:0px 6px;border-radius:8px;font-size:10px;font-weight:600;'>중립</span>"}.get(row["감성"],"")
        gi2=MEDIA_GRADE.get(row["매체"],{}); grade=gi2.get("grade",""); gc_=GRADE_COLOR.get(grade,"#ccc")
        gs=f"<span style='background:{gc_};color:white;padding:0px 3px;border-radius:2px;font-size:8px;font-weight:700;'>{grade}</span>" if grade else ""
        rh+=f"<tr><td style='text-align:center;color:#ccc;font-size:10px;'>{i}</td><td style='font-size:10px;'>{row['일자']}</td><td style='font-size:10px;'>{row['매체']} {gs}</td><td><a href='{row['링크']}' target='_blank' style='color:#003366;text-decoration:none;font-size:11px;'>{row['헤드라인']}</a></td><td style='color:#666;font-size:10px;'>{row['요약']}</td><td>{bg}</td><td style='color:#999;font-size:9px;'>{row.get('카테고리','—')}</td></tr>"
    st.markdown(f"""<div style='overflow-x:auto;'><table style='width:100%;border-collapse:collapse;'><thead><tr style='background:#003366;color:white;'><th style='padding:5px 6px;font-size:10px;'>#</th><th>일자</th><th>언론사</th><th>헤드라인</th><th>요약</th><th>논조</th><th>카테고리</th></tr></thead><tbody>{rh}</tbody></table></div>""",unsafe_allow_html=True)
    st.caption(f"전체 {len(fdf)}건 중 {min(st.session_state[sk],len(fdf))}건")
    if st.session_state[sk]<len(fdf):
        if st.button("▼ 더보기",key=f"more_{label}"): st.session_state[sk]+=30; st.rerun()

    # 다운로드 바
    dl1,dl2,dl3=st.columns(3)
    with dl1:
        out=io.BytesIO()
        with pd.ExcelWriter(out,engine='openpyxl') as w: df.to_excel(w,index=False,sheet_name="데이터")
        out.seek(0)
        st.download_button("📥 엑셀",data=out,file_name=f"한전뉴스_{label}_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key=f"xl_{label}")
    with dl2:
        wb2=make_full_word(cd)
        st.download_button("📄 전체 보고서 워드",data=wb2,file_name=f"KEPCO_{label}_{datetime.now().strftime('%Y%m%d')}.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True,key=f"wd2_{label}")
    with dl3: copy_link_btn()

    st.markdown(f"<div style='background:#003366;color:white;text-align:center;padding:7px;border-radius:4px;margin-top:10px;font-size:10px;opacity:.8;'>⚡ 한국전력 뉴스 유형분석 자동화 시스템 | {datetime.now().strftime('%Y.%m.%d')} | 열독률: 언론진흥재단('23) | 두바이유 기준</div>",unsafe_allow_html=True)
    st.markdown("---")

# ══ APP ═══════════════════════════════════════════════
st.set_page_config(page_title="한국전력 뉴스 유형분석 자동화 시스템",layout="wide",page_icon="⚡",initial_sidebar_state="expanded")
st.markdown("""<style>
.main .block-container{padding-top:.5rem;padding-bottom:.5rem;max-width:1400px;}
[data-testid="stSidebar"]{background:#F4F6F9;}
.stTabs [data-baseweb="tab"]{font-size:11px;padding:4px 12px;}
div[data-testid="stVerticalBlock"]>div{gap:0.3rem;}
</style>""",unsafe_allow_html=True)

for k,v in [("history",[]),("analysis_cache",{}),("active_key",None)]:
    if k not in st.session_state: st.session_state[k]=v

if not YF_OK: st.warning("📦 주가: cmd에서 pip install yfinance 실행 필요",icon="⚠️")
md=get_market_data()
st.markdown(f"""<div style='background:#003366;color:white;padding:8px 16px;border-radius:5px;display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;'><span style='font-size:15px;font-weight:700;'>⚡ 한국전력 뉴스 유형분석 자동화 시스템</span><span style='font-size:8px;opacity:.65;'>{datetime.now().strftime('%Y.%m.%d')} | 열독률 등급 기반 | 네이버 뉴스 API</span></div>""",unsafe_allow_html=True)
st.markdown(mhdr(md),unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 분석 설정")
    with st.form("mf",clear_on_submit=False):
        kc1,kc2=st.columns([5,1])
        with kc1: keywords_input=st.text_input("🔍 키워드 (Enter=분석)","한국전력",placeholder="키워드 입력 후 Enter")
        with kc2:
            st.markdown("<div style='padding-top:24px;'>",unsafe_allow_html=True)
            rb=st.form_submit_button("🚀",use_container_width=True)
            st.markdown("</div>",unsafe_allow_html=True)
        st.caption("쉼표(,)=개별  |  플러스(+)=동시포함")
        cs1,cs2=st.columns(2)
        with cs1: start_date=st.date_input("시작일",datetime.now()-timedelta(days=7))
        with cs2: end_date=st.date_input("종료일",datetime.now())
        max_articles=st.select_slider("수집 기사 수",[500,1000,2000,3000,5000],value=1000)
        crisis_input=st.text_input("🚨 위기 키워드","전기요금 폭탄,정전,파업,감사원")
        run=st.form_submit_button("🚀 분석 시작",use_container_width=True)
        run=run or rb
    st.markdown("---")
    if st.session_state.history:
        st.markdown("**📋 분석 이력**",unsafe_allow_html=True)
        for i,h in enumerate(st.session_state.history[:10],1):
            nr=h['neg']/h['total']*100 if h['total']>0 else 0
            active=(st.session_state.active_key==h['cache_key'])
            if st.button(f"{'▶ ' if active else ''}#{i} {h['keyword']}\n{h['period']} | 부정{nr:.0f}%",key=f"hb_{i}",use_container_width=True):
                st.session_state.active_key=h['cache_key']; st.rerun()
    else: st.caption("분석 후 이력이 쌓입니다")

if run:
    st.session_state.active_key=None
    kw_groups=parse_kw(keywords_input)
    crisis_kws=[k.strip() for k in crisis_input.split(",") if k.strip()]
    all_res=[]
    for g in kw_groups:
        lbl=g["label"]
        with st.spinner(f"'{lbl}' 수집 중... (최대 {max_articles}건)"):
            raw=get_news(" ".join(g["keywords"]),max_articles)
            for a in raw:
                pub=a.get("pubDate","")
                try:
                    ad=datetime.strptime(pub[:16],"%a, %d %b %Y").date()
                    if not (start_date<=ad<=end_date): continue
                    ds=ad.strftime("%Y-%m-%d"); hs=pub[17:19] if len(pub)>18 else "00"
                except: ds=pub[:10]; hs="00"
                title=clean(a.get("title","")); desc=clean(a.get("description",""))
                text=title+" "+desc; orig=a.get("originallink",""); link=a.get("link","")
                if not is_relevant(text): continue
                if g["type"]=="AND" and not matches_and(text,g): continue
                media=get_media(orig,link); gi=MEDIA_GRADE.get(media,{})
                all_res.append({"키워드그룹":lbl,"일자":ds,"월":ds[:7],"시간":hs,"매체":media,"등급":gi.get("grade","—"),"열독률":gi.get("rate",0.05),"헤드라인":title,"요약":summarize(desc,50),"감성":get_sentiment(text),"카테고리":"","링크":orig if orig else link})
    if not all_res: st.error("수집된 기사가 없습니다."); st.stop()
    all_res=auto_cat(all_res)
    all_res=[a for item in [apply_disambig([a],a["키워드그룹"]) for a in all_res] for a in item]
    df_all=pd.DataFrame(all_res)
    for g in kw_groups:
        lbl=g["label"]; df=df_all[df_all["키워드그룹"]==lbl].copy()
        if df.empty: st.warning(f"'{lbl}' 기사 없음"); continue
        arts=df.to_dict("records"); total=len(df); cv=df["감성"].value_counts()
        pos_n=int(cv.get("긍정",0)); neg_n=int(cv.get("부정",0)); neu_n=int(cv.get("중립",0))
        period_str=f"{start_date.strftime('%Y.%m.%d')} ~ {end_date.strftime('%m.%d')}"
        top3m=", ".join(list(df["매체"].value_counts().index[:3]))
        ki=get_key_issues(arts); ti=ki[0] if ki else "—"
        tnc=df[df["감성"]=="부정"]["카테고리"].value_counts().index[0] if neg_n>0 else "없음"
        tpc=df[df["감성"]=="긍정"]["카테고리"].value_counts().index[0] if pos_n>0 else "없음"
        nk=extract_kws(arts,"부정"); uk=extract_kws(arts,"중립"); pk=extract_kws(arts,"긍정")
        tnkw=nk[0][0] if nk else None
        daily=df.groupby("일자").size()
        if len(daily)>=2:
            fh=daily.iloc[:len(daily)//2].mean(); sh=daily.iloc[len(daily)//2:].mean()
            tt=(f"증가 추세({fh:.1f}→{sh:.1f}건/일)" if sh>fh*1.3 else f"감소 추세({fh:.1f}→{sh:.1f}건/일)" if sh<fh*0.7 else f"안정적(일평균 {daily.mean():.1f}건)")
        else: tt=f"총 {total}건"
        it=(f"'{lbl}' 관련 {period_str} 분석 결과, '{tnc}' 이슈 중심으로 부정 언론 환경이 형성되어 선제적 대응이 필요합니다. '{tpc}' 관련 성과는 수치 중심으로 적극 홍보해야 합니다. {tt}.")
        crs=gen_criticisms(arts,lbl); acts=gen_actions(arts,lbl)
        neg_med=[m for m,_ in df[df["감성"]=="부정"]["매체"].value_counts().head(5).items()]
        crisis_found=any(any(ck in a["헤드라인"] or ck in a.get("요약","") for a in arts) for ck in crisis_kws)
        pr_s,pr_l,pr_c=calc_pr_risk(neg_n,total,nk,crisis_found,neg_med)
        ck=f"{lbl}_{period_str}"
        cd={"label":lbl,"period_str":period_str,"df":df,"articles":arts,"total":total,"pos_n":pos_n,"neg_n":neg_n,"neu_n":neu_n,"neg_kws":nk,"neu_kws":uk,"pos_kws":pk,"top_neg_kw":tnkw,"key_issues":ki,"criticisms":crs,"actions":acts,"insights_text":it,"top_neg_cat":tnc,"top_pos_cat":tpc,"top3_media":top3m,"trend_txt":tt,"crisis_kws":crisis_kws,"pr_score":pr_s,"pr_lvl":pr_l,"pr_color":pr_c}
        st.session_state.analysis_cache[ck]=cd; st.session_state.active_key=ck
        st.session_state.history=[h for h in st.session_state.history if not (h["keyword"]==lbl and h["period"]==period_str)]
        st.session_state.history.insert(0,{"keyword":lbl,"period":period_str,"total":total,"pos":pos_n,"neg":neg_n,"neu":neu_n,"top_issue":ti,"cache_key":ck})
        st.session_state.history=st.session_state.history[:10]
        render_report(cd)
elif st.session_state.active_key:
    cd2=st.session_state.analysis_cache.get(st.session_state.active_key)
    if cd2: render_report(cd2)
    else: st.warning("다시 분석해 주세요.")
else:
    if st.session_state.history:
        st.markdown("<div style='font-size:12px;font-weight:700;color:#003366;border-bottom:1.5px solid #003366;padding-bottom:3px;margin:6px 0 10px;'>📋 분석 이력</div>",unsafe_allow_html=True)
        for i,h in enumerate(st.session_state.history[:10],1):
            nr=h['neg']/h['total']*100 if h['total']>0 else 0; pr=h['pos']/h['total']*100 if h['total']>0 else 0
            ca,cb=st.columns([5,1])
            with ca:
                st.markdown(f"""<div style='background:white;border:1px solid #e8e8e8;border-radius:4px;padding:8px 12px;margin-bottom:4px;'><span style='font-size:13px;font-weight:700;color:#003366;'>#{i} {h['keyword']}</span><span style='color:#aaa;font-size:10px;margin-left:8px;'>{h['period']}</span><br><span style='font-size:11px;'>총 {h['total']}건 | <span style='color:#C62828;'>부정 {h['neg']}건({nr:.0f}%)</span> | <span style='color:#1565C0;'>긍정 {h['pos']}건({pr:.0f}%)</span> | {h.get('top_issue','—')}</span></div>""",unsafe_allow_html=True)
            with cb:
                if st.button("열람",key=f"v_{i}",use_container_width=True):
                    st.session_state.active_key=h['cache_key']; st.rerun()
    else:
        st.markdown("""<div style='text-align:center;padding:50px;color:#aaa;'><div style='font-size:32px;'>⚡</div><div style='font-size:15px;font-weight:600;color:#003366;margin-top:8px;'>한국전력 뉴스 유형분석 자동화 시스템</div><div style='font-size:12px;margin-top:6px;'>좌측 키워드 입력 후 Enter 또는 🚀</div></div>""",unsafe_allow_html=True)