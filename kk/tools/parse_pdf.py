# -*- coding: utf-8 -*-
import pdfplumber, re, csv, json, os, collections

def colof(x0):
    if x0<105: return "code"
    if x0<250: return "name"
    if x0<372: return "addr_op"
    if x0<455: return "phone_kin"
    if x0<540: return "opener"     # 開設者=法人名+役職
    if x0<620: return "manager"    # 管理者氏名
    if x0<690: return "desig"
    if x0<779: return "bed"
    return "kind_dep"

DG=re.compile(r"[^0-9]")
TEL=re.compile(r"^0\d{1,4}[-－]\d")
BEDNUM=re.compile(r"(一般（感染）|一般|療養|精神|結核|感染|全床)\s*([0-9,]+)")

def parse_pdf(path, pref, mt):
    recs={}
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            words=pg.extract_words(keep_blank_chars=False)
            if not words: continue
            rowmap=collections.defaultdict(list)
            for w in words: rowmap[round(w['top'])].append(w)
            tops=sorted(rowmap.keys())
            cur=None; cur_top=None
            def flush():
                nonlocal cur
                if cur and cur.get("med_code"):
                    txt=cur.pop("_bedtxt","")
                    b=0
                    for lab,nu in BEDNUM.findall(txt):
                        try:b+=int(DG.sub("",nu) or 0)
                        except:pass
                    cur["beds"]=str(b) if b>0 else ""
                    dep=cur.pop("_deptxt","")
                    STOP=set(["病院","診療所","現存","休止","一般","療養","精神","結核","感染","（感染）","(医","勤","常","非常勤:","勤:"])
                    toks=[t for t in re.split(r"[\s　]+",dep) if t and not re.search(r"\d",t) and t not in STOP and "－" not in t]
                    cur["depts"]="　".join(toks)[:500]
                    cur.pop("_kind",None)
                    recs[cur["med_code"]]=cur
                cur=None
            for tp in tops:
                cells=collections.defaultdict(list)
                for w in sorted(rowmap[tp],key=lambda x:x['x0']):
                    cells[colof(w['x0'])].append(w['text'])
                codetxt="".join(cells.get("code",[]))
                nm="".join(cells.get("name",[]))
                alld=DG.sub("",codetxt)
                if len(alld)>=7 and nm:
                    flush()
                    ph="".join(cells.get("phone_kin",[]))
                    kd="".join(cells.get("kind_dep",[]))
                    cur=dict(med_type=mt,pref_code=pref,med_code=alld[-7:],name=nm[:120],
                             address="".join(cells.get("addr_op",[]))[:200],
                             phone=ph if TEL.match(ph) else "",
                             opener="".join(cells.get("opener",[]))[:150],
                             manager="".join(cells.get("manager",[]))[:120],  # 管理者は初行のみ
                             depts="",
                             _bedtxt=" ".join(cells.get("bed",[])),
                             _deptxt=" ".join(cells.get("bed",[]))+" "+kd,
                             designated_date="".join(cells.get("desig",[]))[:20],
                             source=os.path.basename(path))
                    if "病院" in kd: cur["_kind"]="病院"
                    elif "診療所" in kd: cur["_kind"]="診療所"
                elif cur:
                    a="".join(cells.get("addr_op",[]))
                    opn="".join(cells.get("opener",[]))
                    ph="".join(cells.get("phone_kin",[]))
                    bd=" ".join(cells.get("bed",[]))
                    kd="".join(cells.get("kind_dep",[]))
                    if a: cur["address"]=(cur["address"]+a)[:200]
                    if opn: cur["opener"]=(cur["opener"]+opn)[:150]  # 開設者は複数行連結
                    # 管理者(540-620)継続は初行で確定済のため無視
                    if ph and not cur["phone"] and TEL.match(ph): cur["phone"]=ph
                    if bd: cur["_bedtxt"]=cur.get("_bedtxt","")+" "+bd; cur["_deptxt"]=cur.get("_deptxt","")+" "+bd
                    if kd: cur["_deptxt"]=cur.get("_deptxt","")+" "+kd
            flush()
    for r in recs.values():
        if mt=="医科":
            b=int(r["beds"]) if r["beds"].isdigit() else 0
            r["med_type"]=r.get("_kind") or ("病院" if b>=20 else "診療所")
        r.pop("_kind",None)
    return list(recs.values())

def merge_addr_opener(rec):
    # 開設者法人名は所在地列末尾に含まれる場合がある→addressから法人名部分をopenerへ移送
    a=rec.get("address",""); op=rec.get("opener","")
    keys=["医療法人","公益財団","公益社団","一般財団","一般社団","社会福祉","社会医療","株式会社","有限会社","独立行政","地方独立","国立","日本"]
    pos=-1
    for k in keys:
        i=a.find(k)
        if i>0 and (pos<0 or i<pos): pos=i
    if pos<0:
        m=re.search(r"[^\s]{1,6}(市長|県知事|区長|町長|村長)",a)
        if m: pos=m.start()
    if pos>0:
        rec["opener"]=(a[pos:].strip()+op)[:200]
        rec["address"]=a[:pos].strip()[:200]
    return rec

if __name__=="__main__":
    man=json.load(open("/home/claude/kk/pdf_manifest.json"))
    out=collections.defaultdict(dict); fails=[]
    for m in man:
        fp="/home/claude/kk/pdf/"+os.path.basename(m["path"])
        if not os.path.exists(fp): fails.append((m["path"],"missing")); continue
        try:
            for r in parse_pdf(fp, m["pref"], m["mt"]):
                merge_addr_opener(r)
                out[r["pref_code"]][(r["med_type"],r["med_code"])]=r
        except Exception as e:
            fails.append((os.path.basename(fp),str(e)[:80]))
    cols=["med_type","pref_code","med_code","name","address","phone","opener","manager","depts","beds","designated_date","source"]
    tot=collections.Counter()
    os.makedirs("/home/claude/kk/pdf_out",exist_ok=True)
    for pc,d in out.items():
        with open(f"/home/claude/kk/pdf_out/pref_{pc}.csv","w",newline="",encoding="utf-8") as f:
            w=csv.DictWriter(f,fieldnames=cols); w.writeheader()
            for r in d.values():
                w.writerow({k:r.get(k,"") for k in cols}); tot[r["med_type"]]+=1
    print("by type:",dict(tot),"| prefs:",sorted(out.keys()),"| fails:",len(fails),fails[:5])
