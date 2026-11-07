# -*- coding: utf-8 -*-
"""주문 자동 정리: 받은주문.txt + 입금내역.csv + 카드결제.csv -> 주문정리_날짜.xlsx
사용: python order_auto.py
"""
import csv, re, sys, datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

D = Path(__file__).parent
TODAY = datetime.date.today().strftime("%Y%m%d")
KNUM = {"한": 1, "하나": 1, "두": 2, "둘": 2, "세": 3, "셋": 3, "네": 4, "넷": 4, "다섯": 5}


def read_csv(name):
    p = D / name
    if not p.exists():
        return []
    for enc in ("utf-8-sig", "cp949"):
        try:
            with open(p, encoding=enc) as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            pass
    return []


def won(s):
    return int(re.sub(r"[^\d]", "", str(s)) or 0)


# ---------- 가격표 ----------
prices, ship_fee = [], 0
for r in read_csv("가격표.csv"):
    if r["상품"] == "배송비":
        ship_fee = won(r["가격"])
    else:
        prices.append(r)
product_names = sorted({r["상품"] for r in prices}, key=len, reverse=True)


# ---------- 주문 글 해석 ----------
def parse_block(text):
    o = {"원문": text.strip(), "경로": "", "이름": "", "전화": "", "주소": "",
         "품목": [], "결제방식": "", "확인필요": []}
    m = re.match(r"\s*\[(문자|카톡)\]", text)
    if m:
        o["경로"] = m.group(1)
    m = re.search(r"01[016789][-\s.]?\d{3,4}[-\s.]?\d{4}", text)
    if m:
        d = re.sub(r"\D", "", m.group())
        o["전화"] = f"{d[:3]}-{d[3:-4]}-{d[-4:]}"
    m = (re.search(r"(?:이름|성함|받는\s*분)\s*[:：]?\s*([가-힣]{2,4})", text)
         or re.search(r"([가-힣]{2,4})(?:입니다|이에요|예요|이요)", text))
    if m and not re.search(r"(박스|상자|포|단|망|포기|개|키로|kg)$", m.group(1)) \
            and m.group(1) not in product_names and m.group(1)[0] not in KNUM:
        o["이름"] = m.group(1)
    for line in re.split(r"[\n/]", text):
        m = re.search(r"((?:서울|경기|인천|부산|대구|광주|대전|울산|세종|강원|충북|충남|전북|전남|경북|경남|제주|[가-힣]+시\s)[^\n]*)", line)
        if m:
            addr = re.sub(r"\s*(카드|계좌|입금|결제).*$", "", m.group(1)).strip()
            if re.search(r"(로|길|동|읍|면|리)\s*\d", addr) or re.search(r"\d+(번길|호|동)", addr):
                o["주소"] = addr
                break
    if "카드" in text:
        o["결제방식"] = "카드"
    elif re.search(r"입금|계좌|송금", text):
        o["결제방식"] = "입금"

    # 품목 찾기
    pat = "|".join(map(re.escape, product_names))
    spots = [(m.start(), m.group()) for m in re.finditer(pat, text)]
    for i, (pos, name) in enumerate(spots):
        end = spots[i + 1][0] if i + 1 < len(spots) else len(text)
        seg = text[pos + len(name):end]
        seg = re.split(r"[\n,]|01[016789]|주소|서울|경기|[가-힣]+[시구]\s", seg)[0]
        opts = [r for r in prices if r["상품"] == name]
        row = opts[0]
        for r in opts:
            if r["옵션"] and re.search(re.escape(r["옵션"]).replace("kg", r"\s*(kg|키로|킬로)"), seg, re.I):
                row = r
                seg = re.sub(re.escape(r["옵션"]).replace("kg", r"\s*(kg|키로|킬로)"), "", seg, count=1, flags=re.I)
                break
        if len(opts) > 1 and not any(r["옵션"] and r is row for r in opts):
            o["확인필요"].append(f"{name} 옵션(무게) 없음")
        m = re.search(r"(\d+)\s*(박스|상자|포|단|kg|키로|망|포기|개)?", seg)
        km = re.search(r"(" + "|".join(KNUM) + r")\s*(박스|상자|포|단|망|포기|개|키로)", seg)
        if m and (not km or m.start() < km.start()):
            qty = int(m.group(1))
        elif km:
            qty = KNUM[km.group(1)]
            o["확인필요"].append(f"{name} 수량 글자('{km.group()}')→{qty}로 읽음")
        else:
            qty = 1
            o["확인필요"].append(f"{name} 수량 없음→1로 봄")
        label = name + (f" {row['옵션']}" if row["옵션"] else "")
        o["품목"].append({"품목": label, "수량": qty, "단위": row["단위"],
                         "단가": won(row["가격"]), "금액": qty * won(row["가격"])})
    for k in ("이름", "전화", "주소"):
        if not o[k]:
            o["확인필요"].append(f"{k} 없음")
    if not o["품목"]:
        o["확인필요"].append("품목 못 찾음")
    o["상품금액"] = sum(i["금액"] for i in o["품목"])
    o["합계"] = o["상품금액"] + (ship_fee if o["품목"] else 0)
    return o


raw = (D / "받은주문.txt").read_text(encoding="utf-8")
orders = [parse_block(b) for b in re.split(r"\n\s*\n", raw) if b.strip()]

# ---------- 결제 맞추기 ----------
bank = [dict(r, 사용=False) for r in read_csv("입금내역.csv")]
card = [dict(r, 사용=False) for r in read_csv("카드결제.csv")]

for o in orders:
    o["결제상태"], o["결제근거"] = "미확인", ""
    # 1) 입금: 이름+금액
    for b in bank:
        if not b["사용"] and o["이름"] and o["이름"] in b["입금자명"] and won(b["입금액"]) == o["합계"]:
            b["사용"] = True
            o["결제상태"], o["결제근거"] = "입금확인", f"{b['거래일시']} {b['입금자명']} {won(b['입금액']):,}"
            break
    if o["결제상태"] != "미확인":
        continue
    # 2) 카드: 금액(카드내역엔 이름이 없어 금액으로 맞춤)
    for c in card:
        if not c["사용"] and won(c["금액"]) == o["합계"]:
            c["사용"] = True
            o["결제상태"], o["결제근거"] = "카드확인", f"{c['승인일시']} {c['카드사']} {c['승인번호']} {won(c['금액']):,}"
            if o["결제방식"] != "카드":
                o["확인필요"].append("카드는 금액만 맞음 - 본인 결제인지 확인")
            break
    if o["결제상태"] != "미확인":
        continue
    # 3) 이름은 맞는데 금액 다름
    for b in bank:
        if not b["사용"] and o["이름"] and o["이름"] in b["입금자명"]:
            b["사용"] = True
            diff = won(b["입금액"]) - o["합계"]
            o["결제상태"] = "금액불일치"
            o["결제근거"] = f"{b['입금자명']} {won(b['입금액']):,} 입금 ({'+' if diff > 0 else ''}{diff:,})"
            break

# ---------- 엑셀 ----------
wb = Workbook()
RED = PatternFill("solid", fgColor="FFC7CE")
GRN = PatternFill("solid", fgColor="C6EFCE")
YEL = PatternFill("solid", fgColor="FFEB9C")
B = Font(bold=True)


def sheet(ws, head, rows, widths):
    ws.append(head)
    for c in ws[1]:
        c.font = B
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "A2"


ws = wb.active
ws.title = "주문목록"
rows = []
for n, o in enumerate(orders, 1):
    items = ", ".join(f"{i['품목']} {i['수량']}{i['단위']}" for i in o["품목"])
    rows.append([n, o["경로"], o["이름"], o["전화"], o["주소"], items, o["상품금액"], o["합계"],
                 o["결제상태"], o["결제근거"], " / ".join(o["확인필요"]), o["원문"]])
sheet(ws, ["번호", "경로", "이름", "전화", "주소", "품목", "상품금액", "합계(배송비포함)",
           "결제상태", "결제근거", "확인필요", "원문"], rows, [5, 6, 8, 14, 40, 30, 10, 12, 10, 34, 34, 50])
for row in ws.iter_rows(min_row=2):
    st = row[8].value
    row[8].fill = GRN if st in ("입금확인", "카드확인") else RED if st == "금액불일치" else YEL
    if row[10].value:
        row[10].fill = YEL
    for c in (row[6], row[7]):
        c.number_format = "#,##0"

# 발송목록: 결제 확인 + 확인필요 없는 것만 (택배 대량등록용)
ws2 = wb.create_sheet("발송목록")
ship = [o for o in orders if o["결제상태"] in ("입금확인", "카드확인") and not o["확인필요"]]
sheet(ws2, ["받는분", "전화번호", "주소", "상품명", "수량", "메모"],
      [[o["이름"], o["전화"], o["주소"], i["품목"], i["수량"], ""] for o in ship for i in o["품목"]],
      [8, 14, 45, 14, 6, 20])

ws3 = wb.create_sheet("확인필요")
hold = [o for o in orders if o not in ship]
sheet(ws3, ["이름", "전화", "결제상태", "결제근거", "확인할 것", "원문"],
      [[o["이름"], o["전화"], o["결제상태"], o["결제근거"], " / ".join(o["확인필요"]), o["원문"]] for o in hold],
      [8, 14, 10, 34, 40, 60])

ws4 = wb.create_sheet("짝없는결제")
sheet(ws4, ["구분", "일시", "이름/카드사", "금액"],
      [["입금", b["거래일시"], b["입금자명"], won(b["입금액"])] for b in bank if not b["사용"]]
      + [["카드", c["승인일시"], c["카드사"], won(c["금액"])] for c in card if not c["사용"]],
      [6, 18, 12, 10])

# 품목별 합계(포장·출고 준비용)
ws5 = wb.create_sheet("품목합계")
tot = {}
for o in ship:
    for i in o["품목"]:
        tot[(i["품목"], i["단위"])] = tot.get((i["품목"], i["단위"]), 0) + i["수량"]
sheet(ws5, ["품목", "수량", "단위"], [[k[0], v, k[1]] for k, v in tot.items()], [14, 8, 6])

out = D / f"주문정리_{TODAY}.xlsx"
wb.save(out)
print(f"주문 {len(orders)}건 | 발송가능 {len(ship)}건 | 확인필요 {len(hold)}건")
for o in hold:
    print(f"  - {o['이름'] or '(이름없음)'}: {o['결제상태']} {o['결제근거']} {' / '.join(o['확인필요'])}")
print(f"저장: {out}")
