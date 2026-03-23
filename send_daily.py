# -*- coding: utf-8 -*-
"""
GitHub Actions 에서 매일 정시 실행되는 뉴스 알리미 발송 스크립트
"""
import os, json, smtplib, requests, re, random
import pandas as pd
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from collections import Counter

# ── 환경변수에서 시크릿 로드 ──
NAVER_CLIENT_ID     = os.environ["NAVER_CLIENT_ID"]
NAVER_CLIENT_SECRET = os.environ["NAVER_CLIENT_SECRET"]
GMAIL_SENDER        = os.environ["GMAIL_SENDER"].strip().replace("\xa0","").replace('"','').replace("'",'')
_raw_pw             = os.environ["GMAIL_APP_PW"]
GMAIL_APP_PW        = ''.join(c for c in _raw_pw if c.isalnum())
print(f"발신 계정: {GMAIL_SENDER}")
print(f"앱 비밀번호 길이: {len(GMAIL_APP_PW)}자리")
# 구독자 목록: subscribers.json 파일 우선, 없으면 환경변수 SUBSCRIBERS
SUBSCRIBERS = []
if os.path.exists("subscribers.json"):
    try:
        with open("subscribers.json", "r", encoding="utf-8") as f:
            SUBSCRIBERS = json.load(f)
        print(f"subscribers.json 에서 {len(SUBSCRIBERS)}명 로드")
    except: pass

if not SUBSCRIBERS:
    try:
        SUBSCRIBERS = json.loads(os.environ.get("SUBSCRIBERS", "[]"))
        print(f"환경변수에서 {len(SUBSCRIBERS)}명 로드")
    except: pass

if not SUBSCRIBERS:
    print("구독자 없음. 종료.")
    exit(0)

print(f"구독자 {len(SUBSCRIBERS)}명 대상 발송 시작")

# ── 상수 ──
APP_URL = "https://kepco-news-monitor-gbff2xm5nzatkkmvjsd9tm.streamlit.app"

POSITIVE_WORDS = ["성과","달성","개선","혁신","성장","우수","협력","기여","선정","수상","기대","친환경","안정","흑자","증가","추진","완료","승인","투자","확대","증대","국익","선도","도약","강화","지원","활성화","성공","최초","호평","기록","돌파","수주","계약","협약","양해각서","출범","개통","준공","신기록"]
NEGATIVE_WORDS = ["사고","부패","비리","손실","적자","위기","파업","갈등","논란","문제","우려","하락","감소","실패","사망","중단","취소","반발","지연","비판","폭탄","부담","폐쇄","위반","처벌","고발","감사원","적발","의혹","과태료","경고","해임","부실","낭비","특혜","불법","소송","압수수색","조사","수사","민원","항의","경찰","패소","기소","고소","피해","피의자","벌금","구속","재판","징계","결함","은폐","허위","과장","오염","누출","폭발","붕괴","침수","정전"]
IRRELEVANT_PATTERNS = [r'배구',r'축구',r'야구',r'농구',r'골프',r'올림픽',r'월드컵',r'선수단',r'선수(?:\s|가|은|는|이|을|의|도)',r'감독(?:\s|이|은|의)']

MEDIA_GRADE = {
    "조선일보":{"rank":1,"rate":3.73,"grade":"S"},"중앙일보":{"rank":2,"rate":2.45,"grade":"A"},
    "동아일보":{"rank":3,"rate":1.95,"grade":"A"},"매일경제":{"rank":4,"rate":0.97,"grade":"A"},
    "한겨레":{"rank":5,"rate":0.62,"grade":"B"},"한국경제":{"rank":6,"rate":0.43,"grade":"B"},
    "경향신문":{"rank":7,"rate":0.41,"grade":"B"},"연합뉴스":{"rank":8,"rate":0.38,"grade":"B"},
}
GRADE_COLOR = {"S":"#B71C1C","A":"#E64A19","B":"#1565C0","C":"#2E7D32","D":"#616161"}

def clean(t): return re.sub(r'<[^>]+>','',str(t)).strip()
def is_relevant(t): return not any(re.search(p,t) for p in IRRELEVANT_PATTERNS)
def get_sentiment(t):
    p=sum(1 for w in POSITIVE_WORDS if w in t)
    n=sum(1 for w in NEGATIVE_WORDS if w in t)
    return "긍정" if p>n else "부정" if n>p else "중립"
def summarize(t,n=30):
    t=re.sub(r'\s+',' ',t).strip(); return t[:n]+"..." if len(t)>n else t
def get_media(o,l):
    url=o if o else l
    MEDIA_MAP = {
        "chosun":"조선일보","joongang":"중앙일보","donga":"동아일보","hani":"한겨레",
        "khan":"경향신문","yna":"연합뉴스","ytn":"YTN","hankyung":"한국경제",
        "mk.co":"매일경제","edaily":"이데일리","heraldcorp":"헤럴드경제",
        "newsis":"뉴시스","kbs":"KBS","mbc":"MBC","sbs":"SBS",
    }
    for d,n in MEDIA_MAP.items():
        if d in url: return n
    return "기타"

def get_news(q, mx=1000):
    url = "https://openapi.naver.com/v1/search/news.json"
    hdr = {"X-Naver-Client-Id": NAVER_CLIENT_ID, "X-Naver-Client-Secret": NAVER_CLIENT_SECRET}
    items, s = [], 1
    while s <= mx:
        try:
            r = requests.get(url, headers=hdr,
                params={"query":q,"display":100,"start":s,"sort":"date"}, timeout=10)
            batch = r.json().get("items", [])
            if not batch: break
            items.extend(batch)
            if len(batch) < 100: break
            s += 100
        except: break
    return items

def collect_and_build_html(label, days):
    end_dt   = datetime.now().date()
    start_dt = end_dt - timedelta(days=max(1, int(days)))
    print(f"    수집 기간: {start_dt} ~ {end_dt}")
    raw = get_news(label, 1000)
    print(f"    네이버 API 수집: {len(raw)}건")
    arts = []
    for a in raw:
        pub = a.get("pubDate","")
        try:
            ad = datetime.strptime(pub[:16], "%a, %d %b %Y").date()
            ds = ad.strftime("%Y-%m-%d")
        except:
            ds = pub[:10]
        title = clean(a.get("title","")); desc = clean(a.get("description",""))
        text  = title + " " + desc
        orig  = a.get("originallink",""); link = a.get("link","")
        if not is_relevant(text): continue
        media = get_media(orig, link)
        arts.append({"일자":ds,"매체":media,"헤드라인":title,
                     "요약":summarize(desc,30),"감성":get_sentiment(text),
                     "링크":orig if orig else link})
    print(f"    필터 후 기사: {len(arts)}건")
    if not arts:
        return None, None

    df = pd.DataFrame(arts)
    period_str = f"{start_dt.strftime('%Y.%m.%d')} ~ {end_dt.strftime('%m.%d')}"
    total = len(df)
    cv    = df["감성"].value_counts()
    neg_n = int(cv.get("부정",0)); pos_n = int(cv.get("긍정",0)); neu_n = int(cv.get("중립",0))
    neg_rate = neg_n/total*100; pos_rate = pos_n/total*100
    tone_txt = "부정 우세" if neg_n>pos_n*1.5 else "긍정 우세" if pos_n>neg_n*1.5 else "균형"
    tone_color = "#C62828" if tone_txt=="부정 우세" else "#1565C0" if tone_txt=="긍정 우세" else "#E65100"
    now_str = (datetime.utcnow()+timedelta(hours=9)).strftime("%Y년 %m월 %d일 %H:%M")

    # 키워드
    neg_txt = " ".join(df[df["감성"]=="부정"]["헤드라인"].tolist())
    pos_txt = " ".join(df[df["감성"]=="긍정"]["헤드라인"].tolist())
    nk = sorted({w:neg_txt.count(w) for w in NEGATIVE_WORDS if neg_txt.count(w)>0}.items(), key=lambda x:-x[1])[:5]
    pk = sorted({w:pos_txt.count(w) for w in POSITIVE_WORDS if pos_txt.count(w)>0}.items(), key=lambda x:-x[1])[:5]

    neg_kw_html = " ".join([f"<span style='background:#FFEBEE;color:#C62828;padding:3px 8px;border-radius:12px;font-size:11px;font-weight:700;margin:2px;display:inline-block;'>{k}({v})</span>" for k,v in nk])
    pos_kw_html = " ".join([f"<span style='background:#E3F2FD;color:#1565C0;padding:3px 8px;border-radius:12px;font-size:11px;font-weight:700;margin:2px;display:inline-block;'>{k}({v})</span>" for k,v in pk])

    # 주요 기사
    neg_top = df[df["감성"]=="부정"].sort_values("일자",ascending=False).head(5)
    pos_top = df[df["감성"]=="긍정"].sort_values("일자",ascending=False).head(3)

    neg_rows = "".join([
        f"<tr><td style='padding:6px 8px;font-size:11px;color:#777;border-bottom:1px solid #f0f0f0;'>{r['매체']}</td>"
        f"<td style='padding:6px 8px;font-size:12px;border-bottom:1px solid #f0f0f0;'>"
        f"<a href='{r['링크']}' style='color:#003366;text-decoration:none;'>{r['헤드라인']}</a></td>"
        f"<td style='padding:6px 8px;font-size:11px;color:#999;border-bottom:1px solid #f0f0f0;white-space:nowrap;'>{r['일자']}</td></tr>"
        for _,r in neg_top.iterrows()
    ])
    pos_rows = "".join([
        f"<tr><td style='padding:6px 8px;font-size:11px;color:#777;border-bottom:1px solid #f0f0f0;'>{r['매체']}</td>"
        f"<td style='padding:6px 8px;font-size:12px;border-bottom:1px solid #f0f0f0;'>"
        f"<a href='{r['링크']}' style='color:#1565C0;text-decoration:none;'>{r['헤드라인']}</a></td>"
        f"<td style='padding:6px 8px;font-size:11px;color:#999;border-bottom:1px solid #f0f0f0;white-space:nowrap;'>{r['일자']}</td></tr>"
        for _,r in pos_top.iterrows()
    ])

    html = f"""<!DOCTYPE html><html><head><meta charset='utf-8'></head><body style='font-family:Malgun Gothic,Apple SD Gothic Neo,Arial,sans-serif;margin:0;padding:0;background:#f0f2f5;'>
<div style='max-width:700px;margin:20px auto;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.12);'>
  <div style='background:#003366;color:white;padding:20px 24px;'>
    <div style='font-size:20px;font-weight:800;'>⚡ {label} 뉴스 모니터링 리포트</div>
    <div style='font-size:11px;opacity:.75;margin-top:5px;'>{period_str} | {now_str} 자동 발송</div>
  </div>
  <div style='padding:18px 24px;border-bottom:1px solid #eee;'>
    <table style='width:100%;border-collapse:collapse;table-layout:fixed;'>
      <tr>
        <td style='text-align:center;padding:10px;background:#F4F6F9;border-radius:6px;border-top:3px solid #003366;'>
          <div style='font-size:22px;font-weight:800;color:#003366;'>{total}</div>
          <div style='font-size:10px;color:#888;'>총 기사</div>
        </td><td style='width:6px;'></td>
        <td style='text-align:center;padding:10px;background:#FFF8F8;border-radius:6px;border-top:3px solid #C62828;'>
          <div style='font-size:22px;font-weight:800;color:#C62828;'>{neg_n}</div>
          <div style='font-size:10px;color:#888;'>부정 ({neg_rate:.0f}%)</div>
        </td><td style='width:6px;'></td>
        <td style='text-align:center;padding:10px;background:#F0F8FF;border-radius:6px;border-top:3px solid #1565C0;'>
          <div style='font-size:22px;font-weight:800;color:#1565C0;'>{pos_n}</div>
          <div style='font-size:10px;color:#888;'>긍정 ({pos_rate:.0f}%)</div>
        </td><td style='width:6px;'></td>
        <td style='text-align:center;padding:10px;background:#F5F5F5;border-radius:6px;border-top:3px solid #888;'>
          <div style='font-size:22px;font-weight:800;color:#555;'>{neu_n}</div>
          <div style='font-size:10px;color:#888;'>중립</div>
        </td><td style='width:6px;'></td>
        <td style='text-align:center;padding:10px;background:#F4F6F9;border-radius:6px;border-top:3px solid {tone_color};'>
          <div style='font-size:20px;font-weight:800;color:{tone_color};'>{'🔴' if tone_txt=='부정 우세' else '🔵' if tone_txt=='긍정 우세' else '🟡'}</div>
          <div style='font-size:10px;color:#888;'>{tone_txt}</div>
        </td>
      </tr>
    </table>
  </div>
  <div style='padding:16px 24px;border-bottom:1px solid #eee;'>
    <div style='font-size:13px;font-weight:800;color:#003366;border-left:4px solid #003366;padding-left:10px;margin-bottom:10px;'>🔑 주요 키워드</div>
    <div style='margin-bottom:6px;'><b style='font-size:11px;color:#C62828;'>부정 키워드</b><br>{neg_kw_html if neg_kw_html else '<span style="color:#aaa;font-size:11px;">없음</span>'}</div>
    <div><b style='font-size:11px;color:#1565C0;'>긍정 키워드</b><br>{pos_kw_html if pos_kw_html else '<span style="color:#aaa;font-size:11px;">없음</span>'}</div>
  </div>
  <div style='padding:16px 24px;border-bottom:1px solid #eee;'>
    <div style='font-size:13px;font-weight:800;color:#003366;border-left:4px solid #C62828;padding-left:10px;margin-bottom:10px;'>🔴 주요 부정 기사 TOP5</div>
    {'<table style="width:100%;border-collapse:collapse;">' + neg_rows + '</table>' if neg_rows else '<p style="color:#aaa;font-size:12px;">부정 기사 없음</p>'}
  </div>
  <div style='padding:16px 24px;border-bottom:1px solid #eee;'>
    <div style='font-size:13px;font-weight:800;color:#003366;border-left:4px solid #1565C0;padding-left:10px;margin-bottom:10px;'>🔵 주요 긍정 기사 TOP3</div>
    {'<table style="width:100%;border-collapse:collapse;">' + pos_rows + '</table>' if pos_rows else '<p style="color:#aaa;font-size:12px;">긍정 기사 없음</p>'}
  </div>
  <div style='padding:14px 24px;text-align:center;border-bottom:1px solid #eee;'>
    <a href='{APP_URL}' target='_blank' style='display:inline-block;background:#003366;color:white;padding:10px 28px;border-radius:20px;font-size:13px;font-weight:700;text-decoration:none;'>⚡ 전체 분석 보고서 앱에서 보기 →</a>
  </div>
  <div style='background:#f8f8f8;padding:12px 24px;font-size:10px;color:#aaa;text-align:center;'>
    ⚡ 홍보실에 꼭 필요한 뉴스 분석시스템 by 글쓰는 여행자 | 네이버 뉴스 API 기반<br>
    본 메일은 자동 발송되었습니다. 수신을 원하지 않으면 구독을 해제해 주세요.
  </div>
</div></body></html>"""
    return html, period_str

# ── 발송 실행 ──
today_str = (datetime.utcnow()+timedelta(hours=9)).strftime('%Y.%m.%d')
fail_list = []

with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.ehlo()
    server.starttls()
    server.login(GMAIL_SENDER, GMAIL_APP_PW)
    for sub in SUBSCRIBERS:
        label = sub.get("keyword", "뉴스")
        days  = int(sub.get("days", 1))
        addr  = sub.get("email","")
        if not addr:
            continue
        print(f"  [{label}] → {addr} 발송 중...")
        html_body, period_str = collect_and_build_html(label, days)
        if html_body is None:
            print(f"    ⚠️ 기사 없음 — 건너뜀")
            fail_list.append(f"{addr}(기사없음)")
            continue
        subject = f"[{today_str}] {label} 뉴스 모니터링 리포트"
        msg = MIMEMultipart("alternative")
        msg["Subject"] = Header(subject, "utf-8")
        msg["From"]    = GMAIL_SENDER
        msg["To"]      = addr
        msg.attach(MIMEText(html_body, "html", "utf-8"))
        server.send_message(msg)
        print(f"    ✅ 발송 완료")

if fail_list:
    print(f"\n⚠️ 실패: {', '.join(fail_list)}")
else:
    print(f"\n✅ 전체 {len(SUBSCRIBERS)}명 발송 완료")
