# -*- coding: utf-8 -*-
import streamlit as st
import requests
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from datetime import datetime, timedelta
import re, io, random
from collections import Counter
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import streamlit.components.v1 as components

try:
    import yfinance as yf
    YF_OK = True
except:
    YF_OK = False

CLIENT_ID     = "mQj4FzmR2tmebJUYk4uq"
CLIENT_SECRET = "zWsKNP7xrB"
FONT_KR = "'Noto Sans KR', 'Apple SD Gothic Neo', 'Malgun Gothic', Arial, sans-serif"

MEDIA_GRADE = {
    "조선일보":{"rank":1,"rate":3.73,"grade":"S"},"중앙일보":{"rank":2,"rate":2.45,"grade":"A"},
    "동아일보":{"rank":3,"rate":1.95,"grade":"A"},"매일경제":{"rank":4,"rate":0.97,"grade":"A"},
    "한겨레":{"rank":5,"rate":0.62,"grade":"B"},"한국경제":{"rank":6,"rate":0.43,"grade":"B"},
    "경향신문":{"rank":7,"rate":0.41,"grade":"B"},"연합뉴스":{"rank":8,"rate":0.38,"grade":"B"},
    "YTN":{"rank":9,"rate":0.35,"grade":"B"},"KBS":{"rank":10,"rate":0.30,"grade":"B"},
    "MBC":{"rank":11,"rate":0.28,"grade":"B"},"SBS":{"rank":12,"rate":0.25,"grade":"B"},
    "한국일보":{"rank":13,"rate":0.31,"grade":"B"},"국민일보":{"rank":14,"rate":0.19,"grade":"C"},
    "문화일보":{"rank":15,"rate":0.12,"grade":"C"},"서울신문":{"rank":16,"rate":0.10,"grade":"C"},
    "서울경제":{"rank":17,"rate":0.08,"grade":"C"},"세계일보":{"rank":18,"rate":0.02,"grade":"D"},
    "머니투데이":{"rank":19,"rate":0.05,"grade":"D"},"이데일리":{"rank":20,"rate":0.04,"grade":"D"},
    "헤럴드경제":{"rank":21,"rate":0.03,"grade":"D"},"뉴시스":{"rank":22,"rate":0.02,"grade":"D"},
    "뉴스핌":{"rank":23,"rate":0.02,"grade":"D"},
}
GRADE_COLOR = {"S":"#B71C1C","A":"#E64A19","B":"#1565C0","C":"#2E7D32","D":"#616161"}

POSITIVE_WORDS = ["성과","달성","개선","혁신","성장","우수","협력","기여","선정","수상","기대","친환경","안정","흑자","증가","추진","완료","승인","투자","확대","증대","국익","선도","도약","강화","지원","활성화","성공","최초","호평","기록","돌파","수주","계약","협약","양해각서","출범","개통","준공","신기록"]
NEGATIVE_WORDS = ["사고","부패","비리","손실","적자","위기","파업","갈등","논란","문제","우려","하락","감소","실패","사망","중단","취소","반발","지연","비판","폭탄","부담","폐쇄","위반","처벌","고발","감사원","적발","의혹","과태료","경고","해임","부실","낭비","특혜","불법","소송","압수수색","조사","수사","민원","항의","경찰","패소","기소","고소","피해","피의자","벌금","구속","재판","징계","결함","은폐","허위","과장","오염","누출","폭발","붕괴","침수","정전"]

IRRELEVANT_PATTERNS = [
    r'배구',r'축구',r'야구',r'농구',r'골프',r'올림픽',r'월드컵',r'선수단',r'드래프트',r'챔피언십',
    r'세트스코어',r'\d+세트',r'배구단',r'득점왕',r'홈런',r'퇴장',r'심판',r'리그',r'경기장',
    r'선수(?:\s|가|은|는|이|을|의|도)',r'감독(?:\s|이|은|의)',r'코치(?:\s|이|의)',
]

MEDIA_MAP = {
    "chosun":"조선일보","joongang":"중앙일보","donga":"동아일보","hani":"한겨레","khan":"경향신문",
    "yna":"연합뉴스","ytn":"YTN","imnews":"MBC","kbs":"KBS","sbs":"SBS","mt.co":"머니투데이",
    "edaily":"이데일리","heraldcorp":"헤럴드경제","newsis":"뉴시스","newspim":"뉴스핌",
    "etnews":"전자신문","energy-news":"에너지신문","electimes":"일렉트릭타임스",
    "hankyung":"한국경제","mk.co":"매일경제","sedaily":"서울경제","ajunews":"아주경제",
    "businesspost":"비즈니스포스트","fnnews":"파이낸셜뉴스","inews24":"아이뉴스24",
    "dt.co":"디지털타임스","hankookilbo":"한국일보","munhwa":"문화일보","ohmynews":"오마이뉴스",
    "pressian":"프레시안","energydaily":"에너지데일리","naeil":"내일신문","seoul":"서울신문",
    "ekn":"에너지경제","kukminilbo":"국민일보","segyetimes":"세계일보","e2news":"이투뉴스",
}

TOPIC_GROUPS = {
    "전기요금":["전기요금","요금","전력요금","인상","누진제","전기세"],
    "원전·수출":["원전","수출","원자력","UAE","체코","APR","해외수주"],
    "재무·경영":["흑자","적자","부채","재무","비상경영","원가","실적","손실"],
    "전력망·설비":["송전","배전","전력망","변전","선로","전력설비","정전","계통"],
    "탄소중립·에너지전환":["탄소중립","RE100","온실가스","수소","재생에너지","넷제로","태양광","풍력"],
    "노사관계":["노사","노조","파업","임금","단체협약","쟁의"],
    "안전·사고":["안전","사고","재해","산재","폭발","화재","부상","사망"],
    "AI·디지털혁신":["AI","인공지능","디지털","스마트","자동화","AX","빅데이터"],
    "공기업·거버넌스":["공기업","감사","이사회","투명","거버넌스","윤리","비리"],
    "고객·서비스":["서비스","고객","민원","복지","국민","전기복지"],
    "정책·규제":["정책","규제","법안","제도","정부","국회","의원","경찰","조사","소송"],
}
DISAMBIG_MAP = {"김동철":["한전","사장","한국전력","KEPCO"],"김성환":["장관","산업부"]}

INSIGHT_DB = {
    "전기요금":{"bg":"요금 인상 이슈는 소비자 민감도 최상위. 언론 집중 공세 대상.","action":"원가회수율·에너지복지 데이터 팩트시트 배포 및 취약계층 지원 성과 수치화","steps":["원가 회수율·지원 가구 수 팩트시트 즉시 배포","취약계층 전기요금 지원 성과 수치화","핵심 언론 1:1 설명회 개최"],"msg":"감정이 아닌 숫자(원가회수율, 지원가구수)로 설득하라."},
    "재무·경영":{"bg":"부채·적자 보도는 공기업 신뢰도 직격. 투자자·여론 이중 영향.","action":"흑자전환·부채 감축 실적 정량 데이터 선제 공개","steps":["흑자 전환·부채 감축 실적 정량 보도자료","경영 효율화 구체 사례 스토리","경제지 전담 관계 강화"],"msg":"'얼마나 줄었나'를 먼저 말하라. 과거 대비 개선폭이 핵심."},
    "노사관계":{"bg":"파업·갈등 이슈는 공공서비스 안정성 우려로 확산되는 전형적 패턴.","action":"협상 진행 현황 정기 브리핑으로 루머 선제 차단","steps":["협상 진행 상황 정기 브리핑","공공서비스 안정 유지 메시지 선점","노사 공동 성명 화해 시그널 제공"],"msg":"침묵은 최악. 협상 중이라도 '소통하고 있다'는 사실을 보여라."},
    "공기업·거버넌스":{"bg":"감사·비리 보도는 기관 전체 신뢰도 훼손. 확산 속도 빠름.","action":"감사 결과 자진 공개 및 윤리경영 조치사항 적극 홍보","steps":["감사 결과 자진 공개로 선제 대응","윤리경영 구체 조치 언론 제공","외부 감사 제3자 검증 활용"],"msg":"숨기면 더 크게 터진다. 먼저 공개하는 것이 신뢰 전략."},
    "안전·사고":{"bg":"안전사고는 사회적 충격이 크고 언론이 가장 오래 추적하는 이슈.","action":"사고 원인·재발방지책 48시간 내 공식 발표","steps":["사고 원인·재발방지책 48시간 내 공식 발표","현장 안전 투자 금액·건수 데이터 제공","협력업체 안전망 확대 조치 동시 발표"],"msg":"'얼마나 빨리, 얼마나 구체적으로' 대응하느냐가 2차 피해를 막는다."},
    "전력망·설비":{"bg":"노후 설비·정전 보도는 국민 생활 직결. 반복 이슈화 위험.","action":"설비 현대화 로드맵 및 투자 계획 수치 중심 공개","steps":["투자 계획·진행 현황 수치 중심 보도자료","정전 원인 신속 공개 및 복구 타임라인 제시","스마트그리드·디지털 전환 성과 홍보"],"msg":"'문제 있다'는 언론보다 '우리가 먼저 알고 고치고 있다'가 우선."},
    "탄소중립·에너지전환":{"bg":"이행 속도 논란은 국제 기준·정부 정책과 연계돼 압박 수위 상승.","action":"온실가스 감축 실적 및 재생에너지 투자 정량 공개","steps":["온실가스 감축 실적 정량 공개","재생에너지 투자·발전량 구체 수치 제시","국제 협약 대비 성과 비교 자료 제공"],"msg":"추상적 목표보다 '작년 대비 몇% 줄었다'는 구체 숫자가 신뢰를 만든다."},
    "정책·규제":{"bg":"경찰·소송 관련 보도는 기관 이미지에 즉각적 타격. 법적 리스크 동반.","action":"법무팀 공식 입장 즉시 발표 및 조사 협조 의지 메시지 선점","steps":["법무팀 공식 입장 즉시 발표","사실 관계 오보 정정 요청 적극 집행","조사 협조 의지·투명성 강조 메시지 선점"],"msg":"'법적 대응'보다 '적극 협조·투명하게 소명' 메시지가 여론에 유리."},
    "원전·수출":{"bg":"수출 지연·안전 논란은 국가 신인도와 연계돼 국제적 관심 이슈.","action":"원전 수출 계약·협상 진행 상황 정기 업데이트 공개","steps":["계약·협상 진행 상황 정기 업데이트","안전 기준·국제 인증 현황 구체 자료","UAE 바라카 성공 사례 집중 레퍼런스"],"msg":"불확실성이 비판을 낳는다. 알려줄 수 있는 정보는 먼저 알려라."},
    "AI·디지털혁신":{"bg":"투자 대비 성과 불명확 시 예산 낭비 비판으로 이어짐.","action":"AI 도입 전후 효율 지표 수치 비교 보도자료 배포","steps":["AI 도입 전후 효율 지표 수치 비교","구체적 서비스 개선 사례(응답시간, 오류율) 제시","보안·개인정보 보호 조치 별도 홍보"],"msg":"'AI 도입'이 아니라 '덕분에 이렇게 달라졌다'는 Before-After 스토리가 효과적."},
    "고객·서비스":{"bg":"민원·서비스 불만은 SNS 확산이 빠르고 감정적 반응 유발.","action":"민원 처리 속도·만족도 지표 공개 및 취약계층 서비스 사례 강조","steps":["민원 처리 속도·만족도 지표 공개","취약계층 특화 서비스 사례 강조","24시간 대응 체계 구축 사실 홍보"],"msg":"민원 통계보다 '실제 해결된 사람의 이야기'가 언론에 더 잘 먹힌다."},
}
DEFAULT_INSIGHT = {"bg":"커뮤니케이션 공백이 부정 보도의 가장 큰 원인.","action":"해당 이슈 공식 입장 48시간 내 발표 및 담당 부서 창구 일원화","steps":["해당 이슈 공식 입장 48시간 내 발표","담당 부서 창구 일원화","미디어 대응 매뉴얼 사전 준비"],"msg":"말하지 않으면 언론이 대신 말한다. 먼저, 빠르게, 구체적으로."}

def gen_paired_insights(criticisms):
    result = []
    for c in criticisms:
        cat = c.get("category", c["title"])
        db = INSIGHT_DB.get(cat, DEFAULT_INSIGHT)
        result.append({"criticism": c, "db": db})
    return result

def gen_criticisms(arts, kw):
    neg=[a for a in arts if a["감성"]=="부정"]; cat_c=Counter([a["카테고리"] for a in neg])
    DB={
        "전기요금":{"title":"전기요금 인상 부담","points":["요금 현실화 국민 공감대 부족","저소득층 에너지 부담 대안 미흡"],"category":"전기요금"},
        "재무·경영":{"title":"재무구조 악화 우려","points":["부채 증가·재무건전성 의문","비상경영 조치 실효성 지적"],"category":"재무·경영"},
        "노사관계":{"title":"노사갈등·파업 리스크","points":["단체협약 갈등으로 경영 불안","공공서비스 안정성 우려"],"category":"노사관계"},
        "공기업·거버넌스":{"title":"공기업 투명성 지적","points":["경영 비효율 반복 지적","낙하산 인사·지배구조 문제"],"category":"공기업·거버넌스"},
        "안전·사고":{"title":"현장 안전사고 우려","points":["안전관리 체계 실효성 의문","협력사 안전망 확대 요구"],"category":"안전·사고"},
        "전력망·설비":{"title":"전력망 노후화 문제","points":["노후 설비 정전·사고 리스크","현대화 투자 속도 부족"],"category":"전력망·설비"},
        "탄소중립·에너지전환":{"title":"탄소중립 이행 실효성","points":["이행 속도 저조 지적","전환 비용 현실성 논란"],"category":"탄소중립·에너지전환"},
        "정책·규제":{"title":"정책 투명성·법적 리스크","points":["일방적 정책 추진 지적","경찰 조사·소송 리스크"],"category":"정책·규제"},
        "원전·수출":{"title":"원전 수출 신뢰성","points":["추진 일정 지연·불확실성","안전 기준 논란"],"category":"원전·수출"},
        "AI·디지털혁신":{"title":"디지털 전환 실효성","points":["투자 대비 성과 부족","보안·개인정보 리스크"],"category":"AI·디지털혁신"},
        "고객·서비스":{"title":"고객 서비스 대응 미흡","points":["민원 처리 속도 불만","소외계층 접근성 개선 요구"],"category":"고객·서비스"},
    }
    result=[]
    for cat,cnt2 in cat_c.most_common(8):
        if cat=="기타": continue
        item=DB.get(cat,{"title":f"{cat} 비판 보도","points":["모니터링 강화 필요","맞춤 대응 메시지 개발"],"category":cat}).copy()
        item["dots"]=min(5,max(2,cnt2//max(1,len(neg)//10)+2)); result.append(item)
        if len(result)==3: break
    defs=[{"title":"커뮤니케이션 체계 미흡","points":["위기 시 신속 대응 부족","공식 채널 속도 개선 필요"],"dots":3,"category":"기타"},
          {"title":"사회적 책임 이행 부족","points":["CSR 기대치 미충족","이해관계자 소통 강화"],"dots":2,"category":"기타"},
          {"title":"미디어 관계 강화 필요","points":["전담 기자 관계 구축","정기 브리핑 채널 부재"],"dots":2,"category":"기타"}]
    while len(result)<3: result.append(defs.pop(0))
    return result[:3]

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
def summarize(t,n=30):
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
    pool=(NEGATIVE_WORDS if sent=="부정" else POSITIVE_WORDS)
    cnt={w:txt.count(w) for w in pool if txt.count(w)>0}
    return sorted(cnt.items(),key=lambda x:x[1],reverse=True)[:n]
def get_media_rank(media): return MEDIA_GRADE.get(media,{}).get("rank",999)
def sentiment_light(s): return {"긍정":"🟢","부정":"🔴","중립":"🟡"}.get(s,"⚪")
def calc_pr_risk(neg_n,total,neg_kws,crisis_found,top_neg_media):
    s=0; neg_r=neg_n/total*100 if total>0 else 0
    s+=min(40,neg_r*0.8)
    if crisis_found: s+=20
    s+=min(20,len(neg_kws)*4)
    sa=[m for m in top_neg_media if MEDIA_GRADE.get(m,{}).get("grade","") in ["S","A"]]
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
    return f"""<div style='background:white;border:1px solid #ddd;border-radius:5px;padding:7px 14px;margin-bottom:8px;display:flex;align-items:center;flex-wrap:wrap;gap:3px;font-family:{FONT_KR};'>
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

# ── 차트 함수 ──────────────────────────────────────────
def cfg(): return {'displayModeBar':False}

def plot_wordcloud(df):
    random.seed(42)
    word_data = {}
    for sent, words in [('부정', NEGATIVE_WORDS), ('긍정', POSITIVE_WORDS)]:
        sub = df[df['감성']==sent]
        txt = " ".join(sub['헤드라인'].tolist())
        for w in words:
            cnt = txt.count(w)
            if cnt >= 1:
                if w not in word_data or word_data[w][1] < cnt:
                    word_data[w] = (sent, cnt)
    items = sorted(word_data.items(), key=lambda x: -x[1][1])[:28]
    max_cnt = items[0][1][1] if items else 1

    xs, ys, texts, sizes, cols, hover = [0], [0], ['한국전력'], [48], ['#003366'], [f'한국전력 | 총 {len(df)}건']
    angle_step = 2.399; r_step = 0.15; base_r = 0.45
    for i, (word, (sent, cnt)) in enumerate(items):
        angle = i * angle_step
        r = base_r + r_step * (i // 6)
        x = r * np.cos(angle) * 2.2 + random.uniform(-0.1, 0.1)
        y = r * np.sin(angle) + random.uniform(-0.07, 0.07)
        size = max(13, min(34, int(13 + (cnt / max_cnt) * 22)))
        color = '#C62828' if sent == '부정' else '#1565C0'
        xs.append(x); ys.append(y); texts.append(word)
        sizes.append(size); cols.append(color)
        hover.append(f'{word} | {sent} | {cnt}회')

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=xs, y=ys, mode='text',
        text=texts,
        textfont=dict(size=sizes, color=cols, family=FONT_KR),
        hovertext=hover, hoverinfo='text',
        customdata=hover,
    ))
    fig.add_annotation(x=0, y=0, text='', showarrow=False)
    fig.update_layout(
        xaxis=dict(showgrid=False, showticklabels=False, zeroline=False, range=[-3.5, 3.5]),
        yaxis=dict(showgrid=False, showticklabels=False, zeroline=False, range=[-1.6, 1.6]),
        paper_bgcolor='white', plot_bgcolor='#FAFBFC',
        margin=dict(l=5, r=5, t=5, b=5), height=280,
        font=dict(family=FONT_KR),
    )
    return fig

def plot_donut(pos_n, neg_n, neu_n, total):
    fig = go.Figure(data=[go.Pie(
        labels=['긍정','중립','부정'], values=[pos_n, neu_n, neg_n],
        hole=0.55, marker=dict(colors=['#1565C0','#9E9E9E','#C62828'], line=dict(color='white', width=2)),
        textinfo='percent+label', textfont=dict(size=11, family=FONT_KR),
        hovertemplate='%{label}: %{value}건 (%{percent})<extra></extra>',
        direction='clockwise', sort=False, rotation=90,
    )])
    fig.update_layout(
        showlegend=False, margin=dict(l=5, r=5, t=5, b=5), height=230,
        paper_bgcolor='white', font=dict(family=FONT_KR),
        annotations=[dict(text=f"<b>{total}</b><br>건", x=0.5, y=0.5, font_size=16, showarrow=False, font=dict(family=FONT_KR, color='#003366'))]
    )
    return fig

def plot_buzz(df):
    daily = df.groupby('일자').size().reset_index(name='건수')
    daily['dt'] = pd.to_datetime(daily['일자'])
    by_sent = df.groupby(['일자','감성']).size().unstack(fill_value=0)
    fig = go.Figure()
    for sent, color in [('부정','#FFCDD2'),('중립','#E0E0E0'),('긍정','#BBDEFB')]:
        if sent in by_sent.columns:
            y = by_sent[sent].reindex(daily['일자'], fill_value=0).values
            fig.add_trace(go.Bar(x=daily['dt'], y=y, name=sent, marker_color=color, hovertemplate=f'{sent}: %{{y}}건<extra></extra>'))
    fig.add_trace(go.Scatter(x=daily['dt'], y=daily['건수'], mode='lines+markers', name='전체', line=dict(color='#003366', width=2), marker=dict(size=5, color='white', line=dict(width=2, color='#003366')), hovertemplate='%{x|%Y-%m-%d}<br>전체: <b>%{y}건</b><extra></extra>'))
    fig.update_layout(barmode='stack', plot_bgcolor='white', paper_bgcolor='white', font=dict(family=FONT_KR, size=11), margin=dict(l=40, r=10, t=10, b=35), height=230, hovermode='x unified', showlegend=True, legend=dict(orientation='h', y=1.08, x=1, xanchor='right', font=dict(size=10)), xaxis=dict(tickformat='%m/%d', showgrid=False, tickangle=-30), yaxis=dict(showgrid=True, gridcolor='#f5f5f5', rangemode='tozero'))
    return fig

def plot_kw_trend(df, kw, mode='daily', date_from=None, date_to=None):
    mask = df['헤드라인'].str.contains(kw, na=False, regex=False) | df['요약'].str.contains(kw, na=False, regex=False)
    kdf = df[mask].copy()
    if kdf.empty: return None
    if date_from: kdf = kdf[kdf['일자'] >= str(date_from)]
    if date_to:   kdf = kdf[kdf['일자'] <= str(date_to)]
    if kdf.empty: return None
    color_map = {'부정':'#C62828','긍정':'#1565C0','중립':'#777777'}
    if mode == 'daily':
        try:
            d_from = pd.to_datetime(str(date_from)) if date_from else pd.to_datetime(kdf['일자'].min())
            d_to   = pd.to_datetime(str(date_to))   if date_to   else pd.to_datetime(kdf['일자'].max())
            all_dates_full = pd.date_range(d_from, d_to, freq='D').strftime('%Y-%m-%d').tolist()
        except:
            all_dates_full = sorted(kdf['일자'].unique())
        grouped = kdf.groupby(['일자','감성']).size().unstack(fill_value=0).reindex(all_dates_full, fill_value=0)
        x = [pd.to_datetime(d) for d in grouped.index]; tick_fmt = '%m/%d'
        n_days = len(all_dates_full)
        dtick_ms = max(1, n_days // 20) * 86400000
    elif mode == 'monthly':
        kdf = kdf.copy(); kdf['월'] = kdf['일자'].str[:7]
        all_months = sorted(kdf['월'].unique())
        grouped = kdf.groupby(['월','감성']).size().unstack(fill_value=0).reindex(all_months, fill_value=0)
        x = grouped.index.tolist(); tick_fmt = None; dtick_ms = None
    else:
        kdf2 = kdf.copy()
        kdf2['시간_int'] = pd.to_numeric(kdf2['시간'], errors='coerce').fillna(0).astype(int)
        grouped = kdf2.groupby(['시간_int','감성']).size().unstack(fill_value=0).reindex(range(24), fill_value=0)
        x = list(range(24)); tick_fmt = None; dtick_ms = 2
    fig = go.Figure()
    for sent, color in color_map.items():
        y = grouped[sent].tolist() if sent in grouped.columns else [0]*len(x)
        m_mode = 'lines' if (mode == 'daily' and len(x) > 14) else 'lines+markers'
        fig.add_trace(go.Scatter(x=x, y=y, mode=m_mode, name=sent,
            line=dict(color=color, width=1.5), marker=dict(size=3),
            hovertemplate=f'<b>{sent}</b> %{{y}}건<extra></extra>'))
    mode_lbl = {'daily':'일자별','monthly':'월별','hourly':'시간대별'}
    xaxis_cfg = dict(showgrid=True, gridcolor='#f0f0f0', tickangle=-45, tickformat=tick_fmt)
    if mode == 'daily' and dtick_ms: xaxis_cfg.update({'dtick': dtick_ms, 'tickmode': 'linear'})
    elif mode == 'hourly': xaxis_cfg.update({'dtick': 2, 'tickmode': 'linear'})
    fig.update_layout(
        title=dict(text=f"<b>「{kw}」 {mode_lbl.get(mode,'')} 추이</b>", font=dict(size=13, color='#003366', family=FONT_KR)),
        plot_bgcolor='white', paper_bgcolor='white', font=dict(family=FONT_KR, size=11),
        margin=dict(l=40, r=10, t=40, b=55), height=270, hovermode='x unified',
        legend=dict(orientation='h', y=1.15, x=1, xanchor='right', font=dict(size=10)),
        xaxis=xaxis_cfg, yaxis=dict(showgrid=True, gridcolor='#f5f5f5', rangemode='tozero')
    )
    return fig

def plot_heatmap_with_hover(df):
    media_counts = df["매체"].value_counts().head(10)
    top_m = sorted(media_counts.index.tolist(), key=get_media_rank)[:8]
    top_c = [c for c in TOPIC_GROUPS if c in df["카테고리"].values][:8]
    if not top_m or not top_c: return None

    z = np.array([[
        round(len(df[(df["매체"]==m)&(df["카테고리"]==cat)&(df["감성"]=="부정")])/
              max(1, len(df[(df["매체"]==m)&(df["카테고리"]==cat)]))*100, 0)
        for cat in top_c] for m in top_m])

    # Build hover text with article headlines per cell
    hover_text = []
    for i, m in enumerate(top_m):
        row_hover = []
        for j, cat in enumerate(top_c):
            arts_cell = df[(df["매체"]==m)&(df["카테고리"]==cat)&(df["감성"]=="부정")]
            arts_cell = arts_cell.sort_values("일자", ascending=False).head(3)
            rate = z[i][j]
            if len(arts_cell) > 0:
                lines = [f"<b>{m} × {cat} ({rate:.0f}%)</b><br>──────────────"]
                for _, r2 in arts_cell.iterrows():
                    lines.append(f"· {r2['일자']}  {r2['헤드라인'][:22]}")
                cell_str = "<br>".join(lines)
            else:
                cell_str = f"<b>{m} × {cat}</b><br>부정 기사 없음"
            row_hover.append(cell_str)
        hover_text.append(row_hover)

    # Top3 annotations
    flat = [(z[i][j], i, j) for i in range(len(top_m)) for j in range(len(top_c))]
    flat.sort(reverse=True)
    top3_cells = [(r,c) for _,r,c in flat[:3]]

    annotations = []
    for i in range(len(top_m)):
        for j in range(len(top_c)):
            rank = next((k+1 for k,(r,c) in enumerate(top3_cells) if r==i and c==j), None)
            txt = f"#{rank} {z[i][j]:.0f}%" if rank else f"{z[i][j]:.0f}%"
            annotations.append(dict(x=j, y=i, text=txt, showarrow=False,
                font=dict(size=10 if not rank else 11, color='white' if z[i][j]>55 else '#333',
                          family=FONT_KR, weight='bold' if rank else 'normal')))

    fig = go.Figure(data=go.Heatmap(
        z=z, x=top_c, y=top_m,
        colorscale=[[0,'#F1F8E9'],[0.4,'#FFF9C4'],[0.7,'#FFB74D'],[1,'#B71C1C']],
        zmin=0, zmax=100,
        text=hover_text,
        hovertemplate='%{text}<extra></extra>',
        showscale=True,
    ))
    fig.update_layout(
        xaxis=dict(tickangle=-30, side='bottom', tickfont=dict(family=FONT_KR, size=10)),
        yaxis=dict(autorange='reversed', tickfont=dict(family=FONT_KR, size=10)),
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(family=FONT_KR, size=10),
        margin=dict(l=90, r=10, t=10, b=70), height=310,
        annotations=annotations
    )
    return fig

def plot_pr_gauge(pr_s, pr_c):
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=pr_s,
        number={"suffix":"점","font":{"size":22,"color":pr_c,"family":FONT_KR}},
        gauge={"axis":{"range":[0,100],"tickwidth":1,"tickcolor":"#ccc","tickfont":{"size":9}},
               "bar":{"color":pr_c,"thickness":0.25},"bgcolor":"#f5f5f5","borderwidth":0,
               "steps":[{"range":[0,40],"color":"#E8F5E9"},{"range":[40,70],"color":"#FFF8E1"},{"range":[70,100],"color":"#FFEBEE"}],
               "threshold":{"line":{"color":pr_c,"width":3},"thickness":0.75,"value":pr_s}}))
    fig.update_layout(height=140, margin=dict(l=20,r=20,t=10,b=10), paper_bgcolor='white', font=dict(family=FONT_KR))
    return fig

# ── Word 보고서 ────────────────────────────────────────
def set_table_header(table, headers, bg="003366", fg="FFFFFF"):
    row = table.rows[0]
    for i, h in enumerate(headers):
        cell = row.cells[i]; cell.text = h
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0] if p.runs else p.add_run(h)
        run.bold = True; run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(int(fg[:2],16), int(fg[2:4],16), int(fg[4:],16))
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),bg)
        tcPr.append(shd)

def make_full_word(cd):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(2); sec.bottom_margin = Cm(2)
        sec.left_margin = Cm(2.5); sec.right_margin = Cm(2.5)
    label=cd['label']; period_str=cd['period_str']; df=cd['df']; total=cd['total']
    pos_n=cd['pos_n']; neg_n=cd['neg_n']; neu_n=cd['neu_n']
    neg_rate=neg_n/total*100; pos_rate=pos_n/total*100
    pr_s=cd.get('pr_score',0); pr_l=cd.get('pr_lvl','—')
    neg_kws=cd['neg_kws']; pos_kws=cd['pos_kws']
    criticisms=cd['criticisms']; top_neg_cat=cd['top_neg_cat']; top_pos_cat=cd['top_pos_cat']

    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("한국전력 언론보도 유형분석 보고서"); r.bold=True; r.font.size=Pt(18); r.font.color.rgb=RGBColor(0,51,102)
    p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"{label}  |  {period_str}  |  {datetime.now().strftime('%Y년 %m월 %d일')}")
    doc.add_paragraph()
    def hd(txt, lv=1):
        h=doc.add_heading(txt,level=lv); h.runs[0].font.color.rgb=RGBColor(0,51,102); return h

    hd("00. 종합 결론 및 PR 리스크")
    tone="부정 우세" if neg_n>pos_n*1.5 else "긍정 우세" if pos_n>neg_n*1.5 else "균형"
    doc.add_paragraph(cd['insights_text'])
    doc.add_paragraph(f"PR 리스크 스코어: {pr_s}점 ({pr_l}) | 논조: {tone}")
    doc.add_paragraph()

    hd("01. 논조 분석")
    tbl0=doc.add_table(rows=2,cols=4); tbl0.style='Table Grid'
    set_table_header(tbl0, ["구분","부정","긍정","중립"])
    row=tbl0.rows[1].cells
    row[0].text="건수 (비율)"; row[1].text=f"{neg_n}건 ({neg_rate:.1f}%)"
    row[2].text=f"{pos_n}건 ({pos_rate:.1f}%)"; row[3].text=f"{neu_n}건 ({neu_n/total*100:.1f}%)"
    doc.add_paragraph()

    hd("02. 매체별 논조 분석")
    media_list=df["매체"].value_counts().head(15).index.tolist()
    media_sorted=sorted(media_list,key=lambda m:(-df[(df["매체"]==m)&(df["감성"]=="부정")].shape[0]/max(1,df[df["매체"]==m].shape[0]), get_media_rank(m)))
    tbl1=doc.add_table(rows=1+len(media_sorted),cols=6); tbl1.style='Table Grid'
    set_table_header(tbl1, ["매체","등급","열독률(%)","총기사","부정","긍정"])
    for i,mname in enumerate(media_sorted,1):
        gi=MEDIA_GRADE.get(mname,{}); cells=tbl1.rows[i].cells
        cells[0].text=mname; cells[1].text=gi.get("grade","—"); cells[2].text=f"{gi.get('rate','')}"
        cells[3].text=str(df[df["매체"]==mname].shape[0])
        cells[4].text=str(df[(df["매체"]==mname)&(df["감성"]=="부정")].shape[0])
        cells[5].text=str(df[(df["매체"]==mname)&(df["감성"]=="긍정")].shape[0])
    doc.add_paragraph()

    hd("03. 논조별 키워드 TOP5")
    tbl2=doc.add_table(rows=1,cols=2); tbl2.style='Table Grid'
    set_table_header(tbl2, ["부정 키워드","긍정 키워드"])
    mx=max(len(neg_kws),len(pos_kws),1)
    for i in range(mx):
        r=tbl2.add_row().cells
        r[0].text=f"{neg_kws[i][0]}({neg_kws[i][1]}회)" if i<len(neg_kws) else ""
        r[1].text=f"{pos_kws[i][0]}({pos_kws[i][1]}회)" if i<len(pos_kws) else ""
    doc.add_paragraph()

    hd("04. 매체×이슈 부정 보도율")
    top_m_l=sorted(df["매체"].value_counts().head(8).index.tolist(),key=get_media_rank)
    top_c_l=[c for c in TOPIC_GROUPS if c in df["카테고리"].values][:7]
    if top_m_l and top_c_l:
        tbl3=doc.add_table(rows=1+len(top_m_l),cols=1+len(top_c_l)); tbl3.style='Table Grid'
        set_table_header(tbl3, ["매체"]+top_c_l)
        for i,mname in enumerate(top_m_l,1):
            cells=tbl3.rows[i].cells; cells[0].text=mname
            for j,cat in enumerate(top_c_l):
                nm=len(df[(df["매체"]==mname)&(df["카테고리"]==cat)])
                nn=len(df[(df["매체"]==mname)&(df["카테고리"]==cat)&(df["감성"]=="부정")])
                cells[j+1].text=f"{round(nn/max(1,nm)*100)}%" if nm>0 else "—"
    doc.add_paragraph()

    hd("05. 비판 포인트 & 대응 전략")
    paired=gen_paired_insights(criticisms)
    tbl4=doc.add_table(rows=1+len(paired),cols=4); tbl4.style='Table Grid'
    set_table_header(tbl4, ["비판 이슈","심각도","대응 전략","핵심 메시지"])
    for i,item in enumerate(paired,1):
        c=item["criticism"]; db=item["db"]; cells=tbl4.rows[i].cells
        cells[0].text=c["title"]; cells[1].text="●"*c["dots"]+"○"*(5-c["dots"])
        cells[2].text=db["action"]; cells[3].text=db["msg"]
    doc.add_paragraph()

    hd("06. 기사 전체 목록")
    df_s=df.copy(); df_s['_r']=df_s['매체'].apply(get_media_rank)
    df_s=df_s.sort_values(['일자','_r'],ascending=[False,True]).reset_index(drop=True)
    tbl5=doc.add_table(rows=1,cols=5); tbl5.style='Table Grid'
    set_table_header(tbl5, ["No.","일자","매체","헤드라인","논조"])
    for idx,row in enumerate(df_s.to_dict("records"),1):
        cells=tbl5.add_row().cells
        cells[0].text=str(idx); cells[1].text=str(row["일자"])
        cells[2].text=str(row["매체"]); cells[3].text=str(row["헤드라인"])
        cells[4].text=str(row["감성"])
    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf

# ── 섹션 헤더 ──────────────────────────────────────────
def divider(n, count_html=""):
    st.markdown(f"<div style='font-size:15px;font-weight:800;color:#003366;letter-spacing:.5px;border-bottom:2px solid #003366;padding-bottom:6px;margin:20px 0 10px;font-family:{FONT_KR};'>{n}{count_html}</div>", unsafe_allow_html=True)

# ── 위기관리 권고 박스 ──────────────────────────────────
def show_crisis_recommendation(pr_s, pr_l, label):
    if pr_s >= 80:
        grade_txt = "A등급(심각)"; bg = "#FFEBEE"; border = "#C62828"; icon = "🚨"
        action = "외부 전문가 공조 체제 즉시 구축. 그룹 전체 위기관리 프로토콜 가동 필요."
        criteria = "중앙부처 문의·조사 개시 / 중앙지·방송 집중 보도 / 전국 규모 단체 항의 가능성"
    elif pr_s >= 70:
        grade_txt = "B등급(경계)"; bg = "#FFF3E0"; border = "#E65100"; icon = "⚠️"
        action = "사업소 단위 위기관리 가동. 지방 언론 및 지역 단체 대응 강화 즉시 실행."
        criteria = "시·도 관공서 문의 / 지방 신문·지역방송 보도 / 지역 단체 항의 가능성"
    else:
        return
    st.markdown(f"""<div style='background:{bg};border:2px solid {border};border-radius:6px;padding:12px 16px;margin-bottom:12px;font-family:{FONT_KR};'>
<div style='font-size:14px;font-weight:800;color:{border};margin-bottom:6px;'>{icon} PR 리스크 {pr_s}점 — {grade_txt} 위기관리 절차 즉시 시행 권고</div>
<div style='font-size:12px;color:#333;margin-bottom:4px;'><b>판단 기준:</b> {criteria}</div>
<div style='font-size:12px;color:{border};font-weight:700;'><b>조치사항:</b> {action}</div>
<div style='font-size:11px;color:#888;margin-top:5px;'>※ 위기관리 메뉴얼 참조: 확산 범위 재확인 → 등급별 대응팀 소집 → 48시간 내 공식 입장 발표</div>
</div>""", unsafe_allow_html=True)

# ══ 보고서 렌더링 ══════════════════════════════════════
def render_report(cd):
    label=cd['label']; period_str=cd['period_str']; df=cd['df']
    total=cd['total']; pos_n=cd['pos_n']; neg_n=cd['neg_n']; neu_n=cd['neu_n']
    neg_kws=cd['neg_kws']; pos_kws=cd['pos_kws']
    top_neg_kw=cd['top_neg_kw']; criticisms=cd['criticisms']
    insights_text=cd['insights_text']; top_neg_cat=cd['top_neg_cat']; top_pos_cat=cd['top_pos_cat']
    top3_media=cd['top3_media']; trend_txt=cd['trend_txt']
    pr_s=cd.get('pr_score',0); pr_l=cd.get('pr_lvl','—'); pr_c=cd.get('pr_color','#888')
    neg_rate=neg_n/total*100; pos_rate=pos_n/total*100
    tone_sym="🔴" if neg_n>pos_n*1.5 else "🟢" if pos_n>neg_n*1.5 else "🟡"
    tone_txt="부정 우세" if neg_n>pos_n*1.5 else "긍정 우세" if pos_n>neg_n*1.5 else "균형"
    neg_kw_str=", ".join([f'{k}({v}회)' for k,v in neg_kws[:3]]) if neg_kws else "없음"
    neg_media_top=df[df['감성']=='부정']['매체'].value_counts().head(3)
    top_neg_m=", ".join([f"{m}({n}건)" for m,n in neg_media_top.items()]) if not neg_media_top.empty else "해당없음"

    # ─ 헤더 ─
    hc1, hc2 = st.columns([5,1])
    with hc1:
        st.markdown(f"""<div style='background:#003366;color:white;padding:10px 16px;border-radius:5px;font-family:{FONT_KR};'>
        <span style='font-size:15px;font-weight:700;'>{label}</span>
        <span style='font-size:11px;opacity:.75;margin-left:10px;'>{period_str} | 총 {total}건</span>
        <span style='float:right;font-size:20px;font-weight:900;color:{pr_c};'>{pr_s}점</span>
        <span style='float:right;font-size:10px;color:rgba(255,255,255,.7);margin-right:6px;margin-top:4px;'>PR리스크</span>
        </div>""", unsafe_allow_html=True)
    with hc2:
        wb=make_full_word(cd)
        st.download_button("📄 보고서 워드", data=wb, file_name=f"KEPCO_{label}_{datetime.now().strftime('%Y%m%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key=f"wd_{label}")

    # 위기관리 권고 (조건부)
    show_crisis_recommendation(pr_s, pr_l, label)

    # ═══ 조사 개요 블록 ═══
    neg_media_list = df[df['감성']=='부정']['매체'].value_counts().head(4)
    pos_media_list = df[df['감성']=='긍정']['매체'].value_counts().head(3)
    neu_media_list = df[df['감성']=='중립']['매체'].value_counts().head(3)

    def media_row(mv, color):
        items = [f"<span style='display:inline-block;background:{color}15;border:1px solid {color}44;border-radius:12px;padding:2px 8px;font-size:10px;color:{color};margin:2px 2px;font-weight:600;'>{m}({n}건)</span>" for m,n in mv.items()]
        return "".join(items) if items else "<span style='font-size:10px;color:#aaa;'>해당없음</span>"

    neg_media_html = media_row(neg_media_list, "#C62828")
    pos_media_html = media_row(pos_media_list, "#1565C0")
    neu_media_html = media_row(neu_media_list, "#555555")

    # 부정 주요 이슈 TOP3
    neg_cats = df[df['감성']=='부정']['카테고리'].value_counts().head(3)
    issue_html = "".join([f"<li style='font-size:11px;color:#333;margin-bottom:3px;'><b>{cat}</b> ({n}건)</li>" for cat,n in neg_cats.items()])
    pos_cats = df[df['감성']=='긍정']['카테고리'].value_counts().head(3)
    pos_issue_html = "".join([f"<li style='font-size:11px;color:#333;margin-bottom:3px;'><b>{cat}</b> ({n}건)</li>" for cat,n in pos_cats.items()])

    tone_overall = "부정 우세" if neg_n > pos_n*1.5 else "긍정 우세" if pos_n > neg_n*1.5 else "균형"
    tone_color   = "#C62828" if neg_n > pos_n*1.5 else "#1565C0" if pos_n > neg_n*1.5 else "#555"

    st.markdown(f"""<div style='background:white;border:1px solid #ddd;border-radius:8px;padding:16px 20px;margin-bottom:12px;font-family:{FONT_KR};'>
  <div style='font-size:16px;font-weight:800;color:#003366;margin-bottom:12px;border-bottom:2px solid #003366;padding-bottom:6px;'>
    📋 {label} 언론보도 분석 — 조사 개요
  </div>
  <div style='display:flex;gap:12px;margin-bottom:14px;flex-wrap:wrap;'>
    <div style='flex:1;min-width:140px;background:#F4F6F9;border-radius:6px;padding:10px 14px;'>
      <div style='font-size:9px;font-weight:700;color:#888;letter-spacing:.8px;margin-bottom:4px;'>조사 기간</div>
      <div style='font-size:12px;font-weight:700;color:#003366;'>{period_str}</div>
    </div>
    <div style='flex:1;min-width:140px;background:#F4F6F9;border-radius:6px;padding:10px 14px;'>
      <div style='font-size:9px;font-weight:700;color:#888;letter-spacing:.8px;margin-bottom:4px;'>분석 방법</div>
      <div style='font-size:12px;font-weight:700;color:#003366;'>네이버 뉴스 API<br>키워드·감성 분석</div>
    </div>
    <div style='flex:1;min-width:140px;background:#F4F6F9;border-radius:6px;padding:10px 14px;'>
      <div style='font-size:9px;font-weight:700;color:#888;letter-spacing:.8px;margin-bottom:4px;'>분석 범위</div>
      <div style='font-size:12px;font-weight:700;color:#003366;'>논조 구분, 매체별 비중<br>키워드·이슈 우선순위</div>
    </div>
    <div style='flex:1;min-width:140px;background:#F4F6F9;border-radius:6px;padding:10px 14px;'>
      <div style='font-size:9px;font-weight:700;color:#888;letter-spacing:.8px;margin-bottom:4px;'>총 수집 기사</div>
      <div style='font-size:14px;font-weight:800;color:#003366;'>{total}건</div>
      <div style='font-size:10px;color:{tone_color};font-weight:700;'>{tone_overall} 기조</div>
    </div>
  </div>
  <div style='display:flex;gap:12px;flex-wrap:wrap;'>
    <div style='flex:2;min-width:200px;'>
      <div style='font-size:11px;font-weight:800;color:#C62828;margin-bottom:5px;'>🔴 부정 논조 주요 매체 ({neg_n}건 · {neg_rate:.0f}%)</div>
      <div style='margin-bottom:8px;'>{neg_media_html}</div>
      <div style='font-size:11px;font-weight:800;color:#1565C0;margin-bottom:5px;'>🔵 긍정 논조 주요 매체 ({pos_n}건 · {pos_rate:.0f}%)</div>
      <div style='margin-bottom:8px;'>{pos_media_html}</div>
      <div style='font-size:11px;font-weight:800;color:#555;margin-bottom:5px;'>⚫ 중립 논조 주요 매체 ({neu_n}건 · {neu_n/total*100:.0f}%)</div>
      <div>{neu_media_html}</div>
    </div>
    <div style='flex:1;min-width:160px;'>
      <div style='font-size:11px;font-weight:800;color:#C62828;margin-bottom:5px;'>주요 비판 이슈</div>
      <ul style='margin:0;padding-left:14px;'>{issue_html}</ul>
    </div>
    <div style='flex:1;min-width:160px;'>
      <div style='font-size:11px;font-weight:800;color:#1565C0;margin-bottom:5px;'>주요 긍정 이슈</div>
      <ul style='margin:0;padding-left:14px;'>{pos_issue_html}</ul>
    </div>
  </div>
</div>""", unsafe_allow_html=True)

    # ═══ 00. KPI + 결론 ═══
    divider("00 · 종합 결론 및 제언")
    k1,k2,k3,k4,k5,k6=st.columns(6)
    for col,val,lbl,color,sub in [
        (k1,str(total),"총 기사","#003366",trend_txt[:18]),
        (k2,f"{neg_n}건","부정","#C62828",f"{neg_rate:.0f}%  {top_neg_cat[:6]}"),
        (k3,f"{pos_n}건","긍정","#1565C0",f"{pos_rate:.0f}%  {top_pos_cat[:6]}"),
        (k4,f"{neu_n}건","중립","#555",f"{neu_n/total*100:.0f}%"),
        (k5,tone_sym,"논조","#333",tone_txt),
        (k6,f"{pr_s}","PR리스크",pr_c,f"{pr_l}  /100점"),
    ]:
        col.markdown(f"""<div style='background:white;border:1px solid #e8e8e8;border-top:3px solid {color};border-radius:4px;padding:8px 6px;text-align:center;font-family:{FONT_KR};'>
        <div style='font-size:19px;font-weight:700;color:{color};line-height:1.1;'>{val}</div>
        <div style='font-size:9px;font-weight:700;color:#999;letter-spacing:.4px;margin-top:2px;'>{lbl}</div>
        <div style='font-size:9px;color:#bbb;margin-top:1px;white-space:nowrap;overflow:hidden;'>{sub}</div></div>""", unsafe_allow_html=True)

    g1,g2=st.columns([1,3])
    with g1:
        st.plotly_chart(plot_pr_gauge(pr_s, pr_c), use_container_width=True, config=cfg())
    with g2:
        paired0 = gen_paired_insights(criticisms)
        a0 = paired0[0]['db']['action'] if paired0 else "—"
        a1 = paired0[1]['db']['action'] if len(paired0)>1 else "—"
        st.markdown(f"""<div style='background:#F8F9FA;border-left:3px solid #003366;border-radius:0 4px 4px 0;padding:10px 14px;font-size:12px;line-height:1.8;font-family:{FONT_KR};'>
        <span style='color:#003366;font-weight:700;'>핵심 비판:</span> {top_neg_cat} &nbsp;|&nbsp; <span style='color:#003366;font-weight:700;'>부정 키워드:</span> {neg_kw_str}<br>
        <span style='color:#003366;font-weight:700;'>부정 집중 매체:</span> {top_neg_m}<br>
        <span style='color:#003366;font-weight:700;'>결론:</span> {insights_text[:120]}…<br>
        <span style='color:#C62828;font-weight:700;'>즉각 대응:</span> {a0}
        </div>""", unsafe_allow_html=True)

    # ═══ 01. 워드클라우드 ═══
    divider("01 · 키워드 워드클라우드")
    wc1, wc2 = st.columns([3,1])
    with wc1:
        fig_wc = plot_wordcloud(df)
        st.plotly_chart(fig_wc, use_container_width=True, config=cfg())
    with wc2:
        st.markdown(f"""<div style='background:#F8F9FA;border-radius:6px;padding:10px;font-family:{FONT_KR};font-size:11px;'>
        <div style='font-weight:700;color:#003366;margin-bottom:6px;'>범례</div>
        <div style='margin-bottom:4px;'><span style='color:#C62828;font-weight:700;'>■</span> 부정 키워드</div>
        <div style='margin-bottom:4px;'><span style='color:#1565C0;font-weight:700;'>■</span> 긍정 키워드</div>
        <div style='margin-bottom:8px;'><span style='color:#888;font-weight:700;'>■</span> 중립 키워드</div>
        <div style='font-size:10px;color:#aaa;'>글자 크기 = 언급 빈도<br>커서를 단어에 올리면<br>상세 정보 표시</div>
        </div>""", unsafe_allow_html=True)

    # ═══ 02. 언론노출 추이 및 논조 분석 ═══
    divider("02 · 언론노출 추이 및 논조 분석")
    b1, b2 = st.columns([1, 2])
    with b1:
        st.plotly_chart(plot_donut(pos_n, neg_n, neu_n, total), use_container_width=True, config=cfg())
    with b2:
        st.plotly_chart(plot_buzz(df), use_container_width=True, config=cfg())

    # ═══ 03. 매체별 논조 ═══
    divider("03 · 매체별 논조")
    media_all = df["매체"].value_counts().head(12).index.tolist()
    media_sorted = sorted(media_all, key=lambda m: (
        -df[(df["매체"]==m)&(df["감성"]=="부정")].shape[0]/max(1,df[df["매체"]==m].shape[0]),
        get_media_rank(m)
    ))
    rows_m = ""
    for mname in media_sorted:
        gi=MEDIA_GRADE.get(mname,{}); grade=gi.get("grade","—"); rate=gi.get("rate",""); gc=GRADE_COLOR.get(grade,"#ccc")
        n_neg=int(df[(df["매체"]==mname)&(df["감성"]=="부정")].shape[0])
        n_pos=int(df[(df["매체"]==mname)&(df["감성"]=="긍정")].shape[0])
        n_tot=int(df[df["매체"]==mname].shape[0])
        bar=int(n_neg/n_tot*14) if n_tot>0 else 0
        rows_m+=f"<tr><td style='padding:4px 8px;font-size:11px;'>{mname} <span style='background:{gc};color:white;padding:0 3px;border-radius:2px;font-size:8px;font-weight:700;'>{grade}</span></td><td style='text-align:right;color:#888;font-size:9px;padding:4px 8px;'>{rate}%</td><td style='text-align:center;color:#C62828;font-size:11px;font-weight:600;padding:4px 8px;'>{n_neg}</td><td style='text-align:center;color:#1565C0;font-size:11px;padding:4px 8px;'>{n_pos}</td><td style='padding:4px 8px;font-size:10px;'><span style='color:#C62828;'>{'■'*bar}</span><span style='color:#eee;'>{'■'*(14-bar)}</span></td></tr>"
    st.markdown(f"""<table style='width:100%;border-collapse:collapse;font-family:{FONT_KR};'>
    <tr style='background:#003366;color:white;font-size:10px;'><th style='padding:5px 8px;text-align:left;'>매체</th><th style='padding:5px 8px;'>열독률</th><th style='padding:5px 8px;'>부정</th><th style='padding:5px 8px;'>긍정</th><th style='padding:5px 8px;text-align:left;'>부정비중</th></tr>{rows_m}</table>""", unsafe_allow_html=True)

    # ═══ 04. 논조별 키워드 + 클릭 추세 ═══
    divider("04 · 논조별 키워드 노출량 — 키워드 클릭 시 추세 분석")

    sel_kw_key = f"sel_kw_{label}"
    if sel_kw_key not in st.session_state:
        st.session_state[sel_kw_key] = neg_kws[0][0] if neg_kws else (pos_kws[0][0] if pos_kws else "")

    kc1, kc2 = st.columns(2)
    with kc1:
        st.markdown(f"<div style='background:#FFEBEE;border:2px solid #C62828;border-radius:8px 8px 0 0;padding:6px 12px;font-size:12px;font-weight:800;color:#C62828;font-family:{FONT_KR};'>🔴 부정 키워드 TOP5 — 클릭하면 추세 분석</div>", unsafe_allow_html=True)
        with st.container():
            btn_cols = st.columns(min(5, len(neg_kws)) if neg_kws else 1)
            for i, (kw, cnt) in enumerate(neg_kws):
                with btn_cols[i % len(btn_cols)]:
                    is_sel = (st.session_state[sel_kw_key] == kw)
                    label_txt = f"{'▶ ' if is_sel else ''}{kw}\n{cnt}회"
                    if st.button(label_txt, key=f"neg_kw_{label}_{kw}", use_container_width=True):
                        st.session_state[sel_kw_key] = kw; st.rerun()
    with kc2:
        st.markdown(f"<div style='background:#E3F2FD;border:2px solid #1565C0;border-radius:8px 8px 0 0;padding:6px 12px;font-size:12px;font-weight:800;color:#1565C0;font-family:{FONT_KR};'>🔵 긍정 키워드 TOP5 — 클릭하면 추세 분석</div>", unsafe_allow_html=True)
        with st.container():
            btn_cols2 = st.columns(min(5, len(pos_kws)) if pos_kws else 1)
            for i, (kw, cnt) in enumerate(pos_kws):
                with btn_cols2[i % len(btn_cols2)]:
                    is_sel = (st.session_state[sel_kw_key] == kw)
                    label_txt = f"{'▶ ' if is_sel else ''}{kw}\n{cnt}회"
                    if st.button(label_txt, key=f"pos_kw_{label}_{kw}", use_container_width=True):
                        st.session_state[sel_kw_key] = kw; st.rerun()

    sel_kw = st.session_state[sel_kw_key]
    if sel_kw:
        st.markdown(f"<div style='font-size:12px;font-weight:700;color:#003366;margin:8px 0 4px;font-family:{FONT_KR};'>📊 선택 키워드: 「{sel_kw}」 추세 분석</div>", unsafe_allow_html=True)
        try:
            df_min_d = datetime.strptime(df['일자'].min(), '%Y-%m-%d').date()
            df_max_d = datetime.strptime(df['일자'].max(), '%Y-%m-%d').date()
        except:
            df_min_d = datetime.now().date()-timedelta(days=30)
            df_max_d = datetime.now().date()

        trend_key = f"trend_{label}"
        if trend_key not in st.session_state:
            st.session_state[trend_key] = {"from": df_min_d, "to": df_max_d, "applied": False}

        dc1, dc2, dc3 = st.columns([2, 2, 1])
        with dc1:
            d_from = st.date_input("시작일", st.session_state[trend_key]["from"], key=f"tf_{label}_{sel_kw}")
        with dc2:
            d_to = st.date_input("종료일", st.session_state[trend_key]["to"], key=f"tt_{label}_{sel_kw}")
        with dc3:
            st.markdown("<div style='padding-top:24px;'>", unsafe_allow_html=True)
            if st.button("📊 언론노출 추세 확인", key=f"tbtn_{label}_{sel_kw}", use_container_width=True):
                st.session_state[trend_key] = {"from": d_from, "to": d_to, "applied": True}
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        applied_from = st.session_state[trend_key]["from"]
        applied_to = st.session_state[trend_key]["to"]

        tt1, tt2 = st.tabs(["📅 월별","📆 일자별"])
        with tt1:
            f = plot_kw_trend(df, sel_kw, 'monthly', applied_from, applied_to)
            if f: st.plotly_chart(f, use_container_width=True, config=cfg())
            else: st.caption("해당 기간 데이터 없음")
        with tt2:
            f = plot_kw_trend(df, sel_kw, 'daily', applied_from, applied_to)
            if f: st.plotly_chart(f, use_container_width=True, config=cfg())
            else: st.caption("해당 기간 데이터 없음")

    # ═══ 05. 매체×이슈 히트맵 ═══
    divider("05 · 매체×이슈 부정 보도율 — 커서를 셀에 올리면 기사 확인")
    fig_hm = plot_heatmap_with_hover(df)
    if fig_hm:
        st.plotly_chart(fig_hm, use_container_width=True, config=cfg())
    else:
        st.caption("데이터 부족으로 히트맵 생성 불가")

    # ═══ 06. 비판 포인트 레이더 + As-Is/To-Be ═══
    divider("06 · 비판 포인트 & 대응 전략 — 문제(As-Is) → 개선(To-Be)")
    paired = gen_paired_insights(criticisms)

    # 레이더차트: 카테고리별 부정 건수 기반 6각형
    cat_labels_all = list(TOPIC_GROUPS.keys())
    cat_neg_counts = {cat: int(df[(df['카테고리']==cat)&(df['감성']=='부정')].shape[0]) for cat in cat_labels_all}
    cat_neg_sorted = sorted(cat_neg_counts.items(), key=lambda x: -x[1])
    radar_cats = [c for c,v in cat_neg_sorted if v > 0][:6]
    if len(radar_cats) < 3:
        radar_cats = [c for c,v in cat_neg_sorted][:6]
    radar_vals = [cat_neg_counts.get(c, 0) for c in radar_cats]
    max_v = max(radar_vals) if radar_vals else 1
    radar_norm = [round(v/max_v*5, 1) for v in radar_vals]

    # 레이더 Plotly
    radar_labels_short = [c[:7] for c in radar_cats]
    fig_radar = go.Figure()
    fig_radar.add_trace(go.Scatterpolar(
        r=radar_norm + [radar_norm[0]],
        theta=radar_labels_short + [radar_labels_short[0]],
        fill='toself',
        fillcolor='rgba(198,40,40,0.15)',
        line=dict(color='#C62828', width=2),
        marker=dict(size=7, color='#C62828'),
        name='비판 강도',
        hovertemplate='<b>%{theta}</b><br>강도: %{r:.1f}/5<br>' +
            '<br>'.join([f"{radar_cats[i]}: {radar_vals[i]}건" for i in range(len(radar_cats))]) + '<extra></extra>',
    ))
    # 평균 라인
    avg_v = sum(radar_norm)/len(radar_norm) if radar_norm else 2.5
    fig_radar.add_trace(go.Scatterpolar(
        r=[avg_v]*len(radar_labels_short) + [avg_v],
        theta=radar_labels_short + [radar_labels_short[0]],
        mode='lines',
        line=dict(color='#003366', width=1, dash='dot'),
        name='평균',
        hoverinfo='skip',
    ))
    fig_radar.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 5], tickfont=dict(size=9), showticklabels=True, gridcolor='#eee'),
            angularaxis=dict(tickfont=dict(size=11, family=FONT_KR, color='#333'), gridcolor='#eee'),
            bgcolor='white',
        ),
        showlegend=True,
        legend=dict(orientation='h', y=-0.08, x=0.5, xanchor='center', font=dict(size=10, family=FONT_KR)),
        paper_bgcolor='white',
        font=dict(family=FONT_KR),
        margin=dict(l=50, r=50, t=20, b=40),
        height=320,
        title=dict(text=f"<b>주요 비판 포인트 우선순위</b><br><sub style='font-size:9px;color:#888;'>비판 강도(5점 만점)</sub>",
                   font=dict(size=12, color='#003366', family=FONT_KR), x=0.5, xanchor='center'),
    )

    r_col, detail_col = st.columns([1, 2])
    with r_col:
        st.plotly_chart(fig_radar, use_container_width=True, config=cfg())

    with detail_col:
        # 카테고리별 순위 + 점수
        st.markdown(f"<div style='font-size:12px;font-weight:800;color:#003366;margin-bottom:8px;font-family:{FONT_KR};'>카테고리별 비판 강도 순위</div>", unsafe_allow_html=True)
        for rank_i, (cat, val) in enumerate(cat_neg_sorted[:6], 1):
            if val == 0: continue
            score = round(val/max_v*5, 1)
            bar_w = int(score/5*100)
            num_circle = ["①","②","③","④","⑤","⑥"][rank_i-1]
            st.markdown(f"""<div style='display:flex;align-items:center;gap:10px;margin-bottom:6px;font-family:{FONT_KR};'>
  <span style='font-size:15px;font-weight:800;color:#C62828;min-width:24px;'>{num_circle}</span>
  <div style='flex:1;'>
    <div style='display:flex;justify-content:space-between;margin-bottom:2px;'>
      <span style='font-size:11px;font-weight:700;color:#333;'>{cat}</span>
      <span style='font-size:11px;font-weight:700;color:#C62828;'>{score}/5점 ({val}건)</span>
    </div>
    <div style='background:#f5f5f5;border-radius:4px;height:6px;'><div style='background:#C62828;width:{bar_w}%;height:6px;border-radius:4px;'></div></div>
  </div>
</div>""", unsafe_allow_html=True)

    # As-Is / To-Be
    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
    col_asis, col_tobe = st.columns(2)
    with col_asis:
        st.markdown(f"<div style='background:#C62828;color:white;padding:7px 14px;border-radius:6px 6px 0 0;font-size:13px;font-weight:800;font-family:{FONT_KR};'>🔴 현재 문제점 (As-Is)</div>", unsafe_allow_html=True)
        for i, item in enumerate(paired, 1):
            c = item["criticism"]
            dots_str = "●"*c["dots"]+"○"*(5-c["dots"])
            pts_html = "  /  ".join(c["points"])
            st.markdown(f"""<div style='border:1px solid #FFCDD2;border-top:none;background:#FFF8F8;padding:10px 14px;margin-bottom:3px;font-family:{FONT_KR};'>
  <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;'>
    <span style='font-size:12px;font-weight:700;color:#C62828;'>이슈 {i}. {c["title"]}</span>
    <span style='font-size:10px;color:#C62828;letter-spacing:1px;'>{dots_str}</span>
  </div>
  <div style='font-size:11px;color:#666;'>{pts_html}</div>
</div>""", unsafe_allow_html=True)
    with col_tobe:
        st.markdown(f"<div style='background:#1565C0;color:white;padding:7px 14px;border-radius:6px 6px 0 0;font-size:13px;font-weight:800;font-family:{FONT_KR};'>✅ 개선 방향 (To-Be)</div>", unsafe_allow_html=True)
        for i, item in enumerate(paired, 1):
            db = item["db"]
            steps_html = "  /  ".join(db["steps"][:2])
            st.markdown(f"""<div style='border:1px solid #BBDEFB;border-top:none;background:#F0F8FF;padding:10px 14px;margin-bottom:3px;font-family:{FONT_KR};'>
  <div style='font-size:12px;font-weight:700;color:#1565C0;margin-bottom:4px;'>전략 {i}. {db["action"]}</div>
  <div style='font-size:11px;color:#555;margin-bottom:5px;'>{steps_html}</div>
  <div style='font-size:11px;background:white;border-left:3px solid #003366;padding:4px 8px;color:#003366;font-weight:700;'>📌 {db["msg"]}</div>
</div>""", unsafe_allow_html=True)

    # ═══ 07. 기사 목록 ═══
    neg_cnt = int(df['감성'].value_counts().get('부정',0))
    neu_cnt = int(df['감성'].value_counts().get('중립',0))
    pos_cnt = int(df['감성'].value_counts().get('긍정',0))
    count_html = f" <span style='font-size:12px;font-weight:400;color:#888;'>🔴 부정 {neg_cnt} · 🟡 중립 {neu_cnt} · 🟢 긍정 {pos_cnt} · 총 {total}건</span>"
    divider("07 · 기사 목록", count_html)

    fdf = df.copy()
    fdf['_rank'] = fdf['매체'].apply(get_media_rank)
    fdf = fdf.sort_values(['일자','_rank'], ascending=[False, True]).reset_index(drop=True)

    # 컬럼별 필터 (엑셀 스타일 — 헤더 바로 위)
    fl1,fl2,fl3,fl4 = st.columns(4)
    all_dates = ["전체"]+sorted(fdf["일자"].unique().tolist(), reverse=True)
    all_media = ["전체"]+sorted(fdf["매체"].unique().tolist(), key=get_media_rank)
    all_sent = ["전체","부정","중립","긍정"]
    all_cat = ["전체"]+sorted(fdf["카테고리"].unique().tolist())
    with fl1: f_date = st.selectbox("📅 일자별 필터", all_dates, key=f"fd_{label}", label_visibility="visible")
    with fl2: f_media = st.selectbox("📰 언론사별 필터", all_media, key=f"fm_{label}", label_visibility="visible")
    with fl3: f_sent = st.selectbox("🎨 논조별 필터", all_sent, key=f"fs_{label}", label_visibility="visible")
    with fl4: f_cat = st.selectbox("🏷️ 카테고리별 필터", all_cat, key=f"fc_{label}", label_visibility="visible")

    if f_date!="전체": fdf=fdf[fdf["일자"]==f_date]
    if f_media!="전체": fdf=fdf[fdf["매체"]==f_media]
    if f_sent!="전체": fdf=fdf[fdf["감성"]==f_sent]
    if f_cat!="전체": fdf=fdf[fdf["카테고리"]==f_cat]
    fdf = fdf.reset_index(drop=True)

    sk = f"s_{label}"
    if sk not in st.session_state: st.session_state[sk] = 30
    ddf = fdf.iloc[:st.session_state[sk]]

    rh = ""
    for i, row in enumerate(ddf.to_dict("records"), 1):
        light = sentiment_light(row["감성"])
        gi2 = MEDIA_GRADE.get(row["매체"],{}); grade=gi2.get("grade",""); gc_=GRADE_COLOR.get(grade,"#ccc")
        gs = f"<span style='background:{gc_};color:white;padding:0 3px;border-radius:2px;font-size:8px;font-weight:700;'>{grade}</span>" if grade else ""
        summ = str(row.get('요약',''))[:30]
        rh += f"<tr><td style='text-align:center;color:#aaa;font-size:10px;padding:4px 6px;'>{i}</td><td style='font-size:10px;padding:4px 6px;'>{row['일자']}</td><td style='font-size:10px;padding:4px 6px;'>{row['매체']} {gs}</td><td style='padding:4px 6px;'><a href='{row['링크']}' target='_blank' style='color:#003366;text-decoration:none;font-size:11px;'>{row['헤드라인']}</a></td><td style='color:#666;font-size:10px;padding:4px 6px;'>{summ}</td><td style='text-align:center;font-size:16px;padding:4px 6px;'>{light}</td><td style='color:#999;font-size:9px;padding:4px 6px;'>{row.get('카테고리','—')}</td></tr>"

    st.markdown(f"""<div style='overflow-x:auto;margin-top:6px;'><table style='width:100%;border-collapse:collapse;font-family:{FONT_KR};'>
    <thead><tr style='background:#003366;color:white;font-size:11px;'>
      <th style='padding:6px 8px;'>No.</th><th style='padding:6px 8px;'>일자</th><th style='padding:6px 8px;'>언론사</th><th style='padding:6px 8px;'>헤드라인</th><th style='padding:6px 8px;'>30자 요약</th><th style='padding:6px 8px;'>논조</th><th style='padding:6px 8px;'>카테고리</th>
    </tr></thead>
    <tbody>{rh}</tbody></table></div>""", unsafe_allow_html=True)

    if st.session_state[sk] < len(fdf):
        if st.button("▼ 더보기", key=f"more_{label}"): st.session_state[sk]+=30; st.rerun()

    # 다운로드 바
    dl1,dl2,dl3 = st.columns(3)
    with dl1:
        out=io.BytesIO()
        with pd.ExcelWriter(out,engine='openpyxl') as w: df.to_excel(w,index=False,sheet_name="데이터")
        out.seek(0)
        st.download_button("📥 엑셀", data=out, file_name=f"한전뉴스_{label}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=f"xl_{label}")
    with dl2:
        wb2=make_full_word(cd)
        st.download_button("📄 전체 보고서 워드", data=wb2, file_name=f"KEPCO_{label}_{datetime.now().strftime('%Y%m%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key=f"wd2_{label}")
    with dl3:
        components.html("""<button id="cpbtn" onclick="(function(){var u=window.parent.location.href;navigator.clipboard.writeText(u).then(function(){document.getElementById('cpbtn').innerText='✅ 복사됨!';document.getElementById('cpbtn').style.background='#2E7D32';setTimeout(function(){document.getElementById('cpbtn').innerText='🔗 링크 복사';document.getElementById('cpbtn').style.background='#003366';},2000);});})();" style="background:#003366;color:white;border:none;padding:8px 16px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;width:100%;">🔗 링크 복사</button>""", height=40)

    st.markdown(f"<div style='background:#003366;color:white;text-align:center;padding:7px;border-radius:4px;margin-top:10px;font-size:10px;opacity:.8;font-family:{FONT_KR};'>⚡ 한국전력 뉴스 유형분석 자동화 시스템 | {datetime.now().strftime('%Y.%m.%d')} | 열독률: 언론진흥재단('23)</div>", unsafe_allow_html=True)
    st.markdown("---")

# ══ APP ═══════════════════════════════════════════════
st.set_page_config(page_title="한국전력 뉴스분석 자동화 시스템", layout="wide", page_icon="⚡", initial_sidebar_state="expanded")
st.markdown(f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;800&display=swap');
.main .block-container{{padding-top:.5rem;padding-bottom:.5rem;max-width:1400px;}}
[data-testid="stSidebar"]{{background:#F4F6F9;}}
.stTabs [data-baseweb="tab"]{{font-size:12px;padding:5px 14px;font-family:{FONT_KR};}}
div[data-testid="stVerticalBlock"]>div{{gap:0.3rem;}}
.main p, .main div, .main span, .main td, .main th, .main label {{font-family:{FONT_KR};}}
</style>""", unsafe_allow_html=True)

for k,v in [("history",[]),("analysis_cache",{}),("active_key",None)]:
    if k not in st.session_state: st.session_state[k]=v

if not YF_OK: st.warning("📦 주가: pip install yfinance 실행 필요", icon="⚠️")
md = get_market_data()
st.markdown(f"""<div style='background:#003366;color:white;padding:8px 16px;border-radius:5px;display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;font-family:{FONT_KR};'><span style='font-size:15px;font-weight:700;'>⚡ 한국전력 뉴스 유형분석 자동화 시스템</span><span style='font-size:8px;opacity:.65;'>{datetime.now().strftime('%Y.%m.%d')} | 열독률 등급 기반 | 네이버 뉴스 API</span></div>""", unsafe_allow_html=True)
st.markdown(mhdr(md), unsafe_allow_html=True)

with st.sidebar:
    st.markdown(f"<h3 style='font-family:{FONT_KR};'>분석 설정</h3>", unsafe_allow_html=True)
    with st.form("mf", clear_on_submit=False):
        kc1s,kc2s = st.columns([5,1])
        with kc1s: keywords_input = st.text_input("🔍 키워드 (Enter=분석)", "한국전력", placeholder="키워드 입력 후 Enter")
        with kc2s:
            st.markdown("<div style='padding-top:24px;'>", unsafe_allow_html=True)
            run = st.form_submit_button("🚀", use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)
        st.caption("쉼표(,)=개별  |  플러스(+)=동시포함")
        cs1,cs2 = st.columns(2)
        with cs1: start_date = st.date_input("시작일", datetime.now()-timedelta(days=7))
        with cs2: end_date = st.date_input("종료일", datetime.now())
        max_articles = st.select_slider("수집 기사 수", [500,1000,2000,3000,5000], value=1000)
        crisis_input = st.text_input("🚨 위기 키워드", "전기요금 폭탄,정전,파업,감사원")
    st.markdown("---")
    if st.session_state.history:
        st.markdown("**📋 분석 이력**")
        for i,h in enumerate(st.session_state.history[:10],1):
            nr=h['neg']/h['total']*100 if h['total']>0 else 0
            active=(st.session_state.active_key==h['cache_key'])
            if st.button(f"{'▶ ' if active else ''}#{i} {h['keyword']}\n{h['period']} | 부정{nr:.0f}%", key=f"hb_{i}", use_container_width=True):
                st.session_state.active_key=h['cache_key']; st.rerun()
    else: st.caption("분석 후 이력이 쌓입니다")

if run:
    st.session_state.active_key = None
    kw_groups = parse_kw(keywords_input)
    crisis_kws = [k.strip() for k in crisis_input.split(",") if k.strip()]
    all_res = []
    for g in kw_groups:
        lbl = g["label"]
        with st.spinner(f"'{lbl}' 수집 중... (최대 {max_articles}건)"):
            raw = get_news(" ".join(g["keywords"]), max_articles)
            for a in raw:
                pub = a.get("pubDate","")
                try:
                    ad = datetime.strptime(pub[:16], "%a, %d %b %Y").date()
                    if not (start_date<=ad<=end_date): continue
                    ds = ad.strftime("%Y-%m-%d"); hs = pub[17:19] if len(pub)>18 else "00"
                except: ds=pub[:10]; hs="00"
                title=clean(a.get("title","")); desc=clean(a.get("description",""))
                text=title+" "+desc; orig=a.get("originallink",""); link=a.get("link","")
                if not is_relevant(text): continue
                if g["type"]=="AND" and not matches_and(text,g): continue
                media=get_media(orig,link); gi=MEDIA_GRADE.get(media,{})
                all_res.append({"키워드그룹":lbl,"일자":ds,"월":ds[:7],"시간":hs,"매체":media,"등급":gi.get("grade","—"),"열독률":gi.get("rate",0.05),"헤드라인":title,"요약":summarize(desc,30),"감성":get_sentiment(text),"카테고리":"","기자":"—","링크":orig if orig else link})
    if not all_res: st.error("수집된 기사가 없습니다."); st.stop()
    all_res = auto_cat(all_res)
    all_res = [a for item in [apply_disambig([a],a["키워드그룹"]) for a in all_res] for a in item]
    df_all = pd.DataFrame(all_res)
    for g in kw_groups:
        lbl=g["label"]; df=df_all[df_all["키워드그룹"]==lbl].copy()
        if df.empty: st.warning(f"'{lbl}' 기사 없음"); continue
        arts=df.to_dict("records"); total=len(df); cv=df["감성"].value_counts()
        pos_n=int(cv.get("긍정",0)); neg_n=int(cv.get("부정",0)); neu_n=int(cv.get("중립",0))
        period_str=f"{start_date.strftime('%Y.%m.%d')} ~ {end_date.strftime('%m.%d')}"
        top3m=", ".join(list(df["매체"].value_counts().index[:3]))
        tnc=df[df["감성"]=="부정"]["카테고리"].value_counts().index[0] if neg_n>0 else "없음"
        tpc=df[df["감성"]=="긍정"]["카테고리"].value_counts().index[0] if pos_n>0 else "없음"
        nk=extract_kws(arts,"부정"); uk=[]; pk=extract_kws(arts,"긍정")
        tnkw=nk[0][0] if nk else None
        daily=df.groupby("일자").size()
        if len(daily)>=2:
            fh=daily.iloc[:len(daily)//2].mean(); sh=daily.iloc[len(daily)//2:].mean()
            tt=(f"증가 추세({fh:.1f}→{sh:.1f}건/일)" if sh>fh*1.3 else f"감소 추세({fh:.1f}→{sh:.1f}건/일)" if sh<fh*0.7 else f"안정적(일평균 {daily.mean():.1f}건)")
        else: tt=f"총 {total}건"
        it=f"'{lbl}' 관련 {period_str} 분석 결과, '{tnc}' 이슈 중심으로 부정 언론 환경이 형성되어 선제적 대응이 필요합니다. '{tpc}' 관련 성과는 수치 중심으로 적극 홍보해야 합니다. {tt}."
        crs=gen_criticisms(arts,lbl)
        neg_med=[m for m,_ in df[df["감성"]=="부정"]["매체"].value_counts().head(5).items()]
        crisis_found=any(any(ck in a["헤드라인"] or ck in a.get("요약","") for a in arts) for ck in crisis_kws)
        pr_s,pr_l,pr_c=calc_pr_risk(neg_n,total,nk,crisis_found,neg_med)
        ck=f"{lbl}_{period_str}"
        cd={"label":lbl,"period_str":period_str,"df":df,"articles":arts,"total":total,"pos_n":pos_n,"neg_n":neg_n,"neu_n":neu_n,"neg_kws":nk,"neu_kws":uk,"pos_kws":pk,"top_neg_kw":tnkw,"criticisms":crs,"insights_text":it,"top_neg_cat":tnc,"top_pos_cat":tpc,"top3_media":top3m,"trend_txt":tt,"crisis_kws":crisis_kws,"pr_score":pr_s,"pr_lvl":pr_l,"pr_color":pr_c}
        st.session_state.analysis_cache[ck]=cd; st.session_state.active_key=ck
        st.session_state.history=[h for h in st.session_state.history if not (h["keyword"]==lbl and h["period"]==period_str)]
        st.session_state.history.insert(0,{"keyword":lbl,"period":period_str,"total":total,"pos":pos_n,"neg":neg_n,"neu":neu_n,"top_issue":tnc,"cache_key":ck})
        st.session_state.history=st.session_state.history[:10]
        render_report(cd)
elif st.session_state.active_key:
    cd2=st.session_state.analysis_cache.get(st.session_state.active_key)
    if cd2: render_report(cd2)
    else: st.warning("다시 분석해 주세요.")
else:
    if st.session_state.history:
        st.markdown(f"<div style='font-size:14px;font-weight:800;color:#003366;border-bottom:2px solid #003366;padding-bottom:5px;margin:10px 0;font-family:{FONT_KR};'>📋 분석 이력</div>", unsafe_allow_html=True)
        for i,h in enumerate(st.session_state.history[:10],1):
            nr=h['neg']/h['total']*100 if h['total']>0 else 0; pr=h['pos']/h['total']*100 if h['total']>0 else 0
            ca,cb=st.columns([5,1])
            with ca:
                st.markdown(f"""<div style='background:white;border:1px solid #e8e8e8;border-radius:4px;padding:8px 12px;margin-bottom:4px;font-family:{FONT_KR};'><span style='font-size:13px;font-weight:700;color:#003366;'>#{i} {h['keyword']}</span><span style='color:#aaa;font-size:10px;margin-left:8px;'>{h['period']}</span><br><span style='font-size:11px;'>총 {h['total']}건 | <span style='color:#C62828;'>부정 {h['neg']}건({nr:.0f}%)</span> | <span style='color:#1565C0;'>긍정 {h['pos']}건({pr:.0f}%)</span></span></div>""", unsafe_allow_html=True)
            with cb:
                if st.button("열람", key=f"v_{i}", use_container_width=True):
                    st.session_state.active_key=h['cache_key']; st.rerun()
    else:
        st.markdown(f"""<div style='text-align:center;padding:50px;color:#aaa;font-family:{FONT_KR};'><div style='font-size:32px;'>⚡</div><div style='font-size:15px;font-weight:600;color:#003366;margin-top:8px;'>한국전력 뉴스 유형분석 자동화 시스템</div><div style='font-size:12px;margin-top:6px;'>좌측 키워드 입력 후 🚀 클릭</div></div>""", unsafe_allow_html=True)
