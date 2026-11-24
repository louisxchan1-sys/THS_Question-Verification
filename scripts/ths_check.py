#!/usr/bin/env python3
"""THS 試卷硬性指標檢查（出題自查與驗證共用）。

用法：
  python3 scripts/ths_check.py <檔案> [<檔案> ...]   # .docx 或 .md（問題紙／評卷參考／試卷草稿）
  python3 scripts/ths_check.py --selftest             # 用 2026 公開試原題核對程式及參數

只檢查可機械判斷的指標（題型比例、NOT 題數、答案分布、DBQ 分值及字數、LQ 背景字數及分值、
參考DSE題目年份、外校名、標註外洩、評卷參考頁碼）。難度、知識準確性、誘答質素等仍須按
驗證清單人手／AI 判斷。所有數字讀自 考評標準/出題總則.md「機讀參數」。
結果：✓ 合格／⚠ 偏離目標（需說明或修改）／✗ 違反硬性規則。有 ✗ 時退出碼為 1。
"""
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RULES_FILE = ROOT / "考評標準" / "出題總則.md"


def load_params():
    t = RULES_FILE.read_text(encoding="utf-8")
    m = re.search(r"<!-- 機讀參數開始 -->\s*```json\s*(\{.*?\})\s*```\s*<!-- 機讀參數結束 -->", t, re.S)
    if not m:
        sys.exit(f"找不到機讀參數：{RULES_FILE}")
    return json.loads(m.group(1))


P = load_params()

# ---------- 讀檔 ----------

NOISE = re.compile(r"^\s*(##\s*PDF Page|<!--|# Raw OCR|\d{4}-DSE-THS|.*Text by HL)")


def read_docx(path):
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    d = docx.Document(str(path))
    lines = []
    for el in d.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            lines.append(Paragraph(el, d).text)
        elif tag == "tbl":
            for r in Table(el, d).rows:
                cells, seen = [], set()
                for c in r.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    cells.append(c.text.replace("\n", " "))
                lines.append("|" + "|".join(cells) + "|")
    meta = {"title": d.core_properties.title or ""}
    with zipfile.ZipFile(path) as z:
        footers = [z.read(n).decode("utf8", "ignore") for n in z.namelist() if re.match(r"word/footer\d*\.xml", n)]
    meta["footer_page"] = any("PAGE" in f.replace("NUMPAGES", "") for f in footers)
    meta["footer_numpages"] = any("NUMPAGES" in f for f in footers)
    return lines, meta


def read_md(path):
    t = path.read_text(encoding="utf-8")
    t = re.sub(r"^---\n.*?\n---\n", "", t, count=1, flags=re.S)
    t = re.sub(r"\n---\n.*?\n---\n", "\n", t, flags=re.S)  # 內嵌 frontmatter（公開試轉檔）
    lines = [l for l in t.split("\n") if not NOISE.match(l)]
    return lines, {}


def split_draft(lines):
    """試卷草稿：按「問題紙全文」「評卷參考全文」標題切開。"""
    qi = ai = end = None
    for i, l in enumerate(lines):
        if re.match(r"^#+\s", l):
            if "問題紙" in l and qi is None:
                qi = i
            elif "評卷參考" in l and ai is None:
                ai = i
            elif ai is not None and end is None and i > ai and re.match(r"^##\s", l):
                end = i
    if qi is None or ai is None:
        return None
    return lines[qi + 1:ai], lines[ai + 1:end]


# ---------- 工具 ----------

def norm(s):
    s = s.replace("（", "(").replace("）", ")").replace("　", " ")
    return re.sub(r"\s+", "", s)


CJK = re.compile(r"[一-鿿]")
MARK = re.compile(r"[（(]\s*(\d+)\s*分\s*[）)]")
NOT_WORDS = ["不是", "不正確", "不恰當", "不適合", "並非", "不屬於", "最不", "不包括"]


class Report:
    def __init__(self, name):
        self.name, self.rows = name, []

    def add(self, status, item, detail=""):
        self.rows.append((status, item, detail))

    def ok(self, item, detail=""):
        self.add("✓", item, detail)

    def warn(self, item, detail=""):
        self.add("⚠", item, detail)

    def fail(self, item, detail=""):
        self.add("✗", item, detail)

    def show(self):
        print(f"\n=== {self.name} ===")
        for s, i, d in self.rows:
            print(f"{s} {i}" + (f"：{d}" if d else ""))

    @property
    def failed(self):
        return any(s == "✗" for s, _, _ in self.rows)


# ---------- 分部 ----------

def sections(lines):
    """回傳 (MC 行, DBQ 行, LQ 行)。"""
    jia = next((i for i, l in enumerate(lines) if re.match(r"^\s*甲部", l)), None)
    yi = next((i for i, l in enumerate(lines) if re.match(r"^\s*乙部", l)), None)
    has_dbq = any("參考所提供的資料" in l for l in lines)
    mc = dbq = lq = []
    if jia is not None:
        mc = lines[jia:yi] if yi is not None and yi > jia else lines[jia:]
        if yi is not None and has_dbq:
            dbq = lines[yi:]
    elif has_dbq:
        dbq = lines
    elif any(re.match(r"^\s*\(a\)", l) for l in lines):
        lq = lines
    elif any(re.match(r"^\s*[A-D][.．、\s\t　]", l) for l in lines):
        mc = lines
    return mc, dbq, lq


# ---------- MC ----------

def parse_mc(lines):
    qs, cur, n = [], None, 1
    for l in lines:
        m = re.match(rf"^\s*{n}(?!\d)[.．、\s\t　]*(\S.*)$", l)
        if m and not re.match(r"^\s*\d+\s*分", l[len(str(n)):]):
            cur = {"no": n, "lines": [m.group(1)]}
            qs.append(cur)
            n += 1
        elif cur is not None:
            cur["lines"].append(l)
    for q in qs:
        stem, stmts, opts, table = [], set(), {}, False
        for l in q["lines"]:
            if l.startswith("|"):
                table = True
                continue
            ms = re.match(r"^\s*[（(]\s*(\d)\s*[)）]", l)
            mo = re.match(r"^\s*([A-D])(?:[.．、\t　]|\s|$)(.*)$", l)
            if mo and (opts or len(mo.group(2).strip()) > 0):
                txt = mo.group(2)
                opts[mo.group(1)] = txt
                if re.search(r"\S\s{2,}\S|\t", txt.strip()):
                    table = True
            elif ms and not opts:
                stmts.add(int(ms.group(1)))
            elif not opts and not stmts:
                stem.append(l)
        hdr = [s.strip() for s in stem if s.strip()]
        if hdr and len(hdr) > 1 and re.search(r"\S\s+\S", hdr[-1]) and not re.search(r"[？?。，,：:／/、]", hdr[-1]):
            table = True  # pdftotext 轉出的表格：選項前一行是欄名
        q["stem"] = "".join(stem)
        q["opts"] = opts
        if 4 in stmts:
            q["type"] = "combo4"
        elif {1, 2, 3} <= stmts:
            no = [norm(o) for o in opts.values()]
            single = sum(bool(re.fullmatch(r"只有\(\d\)[。.]?", o)) for o in no)
            double = sum(bool(re.fullmatch(r"只有\(\d\)及\(\d\)[。.]?", o)) for o in no)
            full = any("(1)" in o and "(2)" in o and "(3)" in o and "只有" not in o for o in no)
            q["type"] = "combo3-new" if single == 2 and double == 2 and not full else "combo3-trad"
            q["full_opt"] = next((k for k, o in opts.items() if "(1)" in norm(o) and "(3)" in norm(o) and "只有" not in o), None)
        elif table:
            q["type"] = "table"
        else:
            q["type"] = "direct"
        q["not"] = [w for w in NOT_WORDS if w in q["stem"]]
    return qs


def check_mc_questions(qs, r):
    n = len(qs)
    if not n:
        return
    r.ok("MC 題數", str(n))
    cnt = Counter(q["type"] for q in qs)
    share = P["MC"]["type_share"]
    parts = []
    for t, s in share.items():
        target = round(s * n)
        got = cnt.get(t, 0)
        parts.append(f"{t} {got}（目標約 {target}）")
        if abs(got - target) > max(1, round(n * 0.07)):
            r.warn(f"MC 題型比例 {t}", f"{got} 題，目標約 {target} 題（MC 規格 §二）")
    r.ok("MC 題型分布", "；".join(parts))
    bad = [f"Q{q['no']}({q['type']})" for q in qs if q["type"] in P["MC"]["banned_types"]]
    if bad:
        r.fail("MC 非 2026 格式（4 陳述／含全選或混合選項）", "、".join(bad))
    nots = [f"Q{q['no']}「{'/'.join(q['not'])}」" for q in qs if q["not"]]
    limit = P["MC"]["not_max"].get(str(n), P["MC"]["not_max"]["other"])
    (r.fail if len(nots) > limit else r.ok)(f"NOT 題（上限 {limit}）", f"{len(nots)} 題 " + "、".join(nots))


def check_mc_answers(ans, r, n_q=None):
    if not ans:
        return
    nums = sorted(ans)
    seq = [ans[i] for i in nums]
    n = len(seq)
    if n_q and n != n_q:
        r.fail("MC 答案數與題數不符", f"答案 {n}、題目 {n_q}")
    c = Counter(seq)
    dist = " ".join(f"{k}{c.get(k, 0)}" for k in "ABCD")
    mx = P["MC"]["letter_max_share"]
    over = [k for k in "ABCD" if c.get(k, 0) > mx * n]
    if over:
        r.fail(f"MC 答案分布（任一字母 ≤{int(mx * 100)}%）", f"{dist}；超標：{''.join(over)}")
    else:
        lo, hi = (7, 8) if n == 30 else (round(n / 4) - 1, round(n / 4) + 1)
        uneven = [k for k in "ABCD" if not lo <= c.get(k, 0) <= hi]
        (r.warn if uneven else r.ok)("MC 答案分布", dist + (f"（目標各 {lo}–{hi}）" if uneven else ""))
    run, runs = 1, []
    for i in range(1, n):
        run = run + 1 if seq[i] == seq[i - 1] else 1
        if run == P["MC"]["max_same_run"] + 1:
            runs.append(f"Q{nums[i - run + 1]}–Q{nums[i]}={seq[i]}")
    (r.fail if runs else r.ok)("MC 連續同答案（不可 3 題或以上）", "、".join(runs))


def parse_mc_answers(lines):
    ans = {}
    rows = [[c.strip() for c in l.strip().strip("|").split("|")] for l in lines if l.strip().startswith("|")]
    for a, b in zip(rows, rows[1:]):
        if a and b and a[0].startswith("題號") and b[0].startswith("答案"):
            for k, v in zip(a[1:], b[1:]):
                if k.isdigit() and re.fullmatch(r"[A-D]", v):
                    ans[int(k)] = v
    for row in rows:
        if row and row[0].startswith("題號"):
            continue
        for k, v in zip(row[0::2], row[1::2]):  # 「題號｜答案」成對欄（含「A (88%)」格式）
            mv = re.match(r"^([A-D])(?:\s*\(\d+%\))?$", v)
            if k.rstrip(".").isdigit() and mv:
                ans.setdefault(int(k.rstrip(".")), mv.group(1))
            else:
                break
    return ans


# ---------- DBQ ----------

def check_dbq(lines, r):
    marks_idx = [i for i, l in enumerate(lines) if "參考所提供的資料" in l]
    if not marks_idx:
        return
    cfg = P["DBQ"]
    start = 0
    for qn, mi in enumerate(marks_idx, 1):
        nxt = marks_idx[qn] if qn < len(marks_idx) else len(lines)
        region = lines[mi + 1:nxt]
        last = max((i for i, l in enumerate(region) if MARK.search(l)), default=-1)
        if qn < len(marks_idx):
            sub = region[:last + 1]
            nxt_start = mi + 1 + last + 1
        else:
            sub, nxt_start = region, len(lines)
        data = [l for l in lines[start:mi] if not re.match(r"^\s*(乙部|本部分|本部共|.*選答)", l)]
        chars = sum(len(CJK.findall(l)) for l in data)
        marks = [int(x) for l in sub for x in MARK.findall(l)]
        lo, hi = cfg["data_chars"]
        st = r.fail if chars > cfg["data_max"] else (r.warn if not lo <= chars <= hi else r.ok)
        st(f"DBQ 第{qn}題 資料字數（中文字）", f"{chars}（目標 {lo}–{hi}，上限 {cfg['data_max']}；轉檔文字或有誤差）")
        tot = sum(marks)
        (r.fail if tot != cfg["total"] else r.ok)(f"DBQ 第{qn}題 總分", f"{tot}（{'+'.join(map(str, marks))}）")
        if sorted(marks, reverse=True) != sorted(cfg["marks"], reverse=True):
            r.warn(f"DBQ 第{qn}題 分值配置", f"{sorted(marks, reverse=True)}，2026 為 {cfg['marks']}（{cfg['subparts']} 個計分分題）")
        start = nxt_start


# ---------- LQ ----------

def check_lq(lines, r):
    cfg = P["LQ"]
    qs, cur = [], None
    for l in lines:
        m = re.match(r"^\s*(\d)(?!\d)[.．、\s\t　]+(\S.*)$", l)
        if m and not re.match(r"^\s*\(", l):
            cur = {"no": m.group(1), "bg": [m.group(2)], "sub": [], "in_sub": False}
            qs.append(cur)
        elif cur is not None:
            if re.match(r"^\s*\([a-b]\)", l):
                cur["in_sub"] = True
            (cur["sub"] if cur["in_sub"] else cur["bg"]).append(l)
    for q in qs:
        bg = "".join(q["bg"])
        chars = len(re.sub(r"\s", "", bg))
        sents = len(re.findall(r"[。？！]", bg))
        marks = [int(x) for l in q["sub"] for x in MARK.findall(l)]
        lo, hi = cfg["bg_chars"]
        st = r.fail if chars > cfg["bg_max"] else (r.warn if not lo <= chars <= hi else r.ok)
        st(f"LQ 第{q['no']}題 背景字數（連標點）", f"{chars}（目標 {lo}–{hi}，上限 {cfg['bg_max']}；{sents} 句）")
        if sents != cfg["sentences"]:
            r.warn(f"LQ 第{q['no']}題 背景句數", f"{sents} 句（2026 為 {cfg['sentences']} 句）")
        ok = sum(marks) == cfg["content_total"] and len(marks) == 2 and all(m in cfg["part_marks"] for m in marks)
        (r.ok if ok else r.fail)(f"LQ 第{q['no']}題 分值", "+".join(map(str, marks)) + f"（須 8+10 或 10+8＝{cfg['content_total']}）")


# ---------- 評卷參考 ----------

def check_annotations(lines, r):
    text = "\n".join(lines)
    segs = re.findall(r"參考DSE題目[：:]([^\n｜|]*(?:[｜|][^\n]*)?)", text)
    if not segs:
        r.warn("參考DSE題目標註", "找不到任何標註")
        return
    y0, y1 = P["anchor_years"]
    bad = []
    for s in segs:
        years = [int(y) for y in re.findall(r"(?<!\d)((?:19|20)\d\d)(?!\d)", s)]
        if "—" in s or "未見" in s or not years:
            bad.append(s.strip()[:30] or "（空白）")
        elif any(not y0 <= y <= y1 for y in years):
            bad.append(s.strip()[:30])
    (r.fail if bad else r.ok)(f"參考DSE題目（只可 {y0}–{y1}，不可「—」）", f"{len(segs)} 處" + ("；問題：" + "；".join(bad[:8]) if bad else ""))
    extra = [w for w in ["目標誤解", "誤解來源", "錯點", "資料立足點", "預估答對率", "推理步數"] if w in text]
    if extra:
        r.fail("評卷參考含不應出現的標註項", "、".join(extra))


# ---------- 通用 ----------

def check_common(lines, meta, r, role):
    text = "\n".join(lines) + "\n" + meta.get("title", "")
    others = {m for m in re.findall(r"香海正覺蓮社佛教\S{2,10}?中學", text) if m != P["school"]}
    label = "評卷參考" if role == "answer" else "問題紙" if role == "question" else role
    (r.fail if others else r.ok)(f"外校名（{label}，含文件屬性標題）", "、".join(others))
    if role == "question":
        leak = [w for w in ["難度：", "參考DSE題目", "目標誤解", "評分說明"] if w in text]
        (r.fail if leak else r.ok)("問題紙無標註外洩", "、".join(leak))
    if role == "answer" and "footer_numpages" in meta:
        ok = meta["footer_page"] and meta["footer_numpages"]
        (r.ok if ok else r.fail)("評卷參考頁腳頁碼（PAGE／NUMPAGES）", "" if ok else "缺頁碼欄位")


def role_of(name, lines):
    if "草稿" in name:
        return "draft"
    if "評卷參考" in name or any(l.strip() == "評卷參考" for l in lines[:12]):
        return "answer"
    if "答題簿" in name:
        return "booklet"
    return "question"


def check_file(path):
    path = Path(path)
    lines, meta = read_docx(path) if path.suffix == ".docx" else read_md(path)
    r = Report(path.name)
    role = role_of(path.name, lines)
    if role == "draft":
        parts = split_draft(lines)
        if not parts:
            r.fail("草稿結構", "找不到「問題紙全文」「評卷參考全文」標題（見 出題流程/草稿範本.md）")
            return r
        qlines, alines = parts
    elif role == "answer":
        qlines, alines = [], lines
    elif role == "booklet":
        check_common(lines, meta, r, role)
        r.ok("答題簿", "只檢查校名")
        return r
    else:
        qlines, alines = lines, []
    if qlines:
        check_common(qlines, meta, r, "question")
        mc, dbq, lq = sections(qlines)
        qs = parse_mc(mc)
        check_mc_questions(qs, r)
        check_dbq(dbq, r)
        check_lq(lq, r)
    else:
        qs = []
    if alines:
        check_common(alines, meta, r, "answer")
        ans = parse_mc_answers(alines)
        check_mc_answers(ans, r, len(qs) or None)
        for q in qs:
            if q.get("full_opt") and ans.get(q["no"]) == q["full_opt"]:
                r.fail(f"MC Q{q['no']} 答案為全選項", "三陳述全對＝送分（經驗庫 #8）")
        check_annotations(alines, r)
    return r


# ---------- 自我測試 ----------

def selftest():
    base = ROOT / "參考資料" / "公開試"
    fails = 0

    def expect(cond, msg):
        nonlocal fails
        print(("✓ " if cond else "✗ ") + msg)
        fails += not cond

    mc_lines, _ = read_md(next((base / "01_選擇題").glob("2026_*.md")))
    qs = parse_mc(sections(mc_lines)[0])
    cnt = Counter(q["type"] for q in qs)
    expect(len(qs) == 30, f"2026 MC 題數 30（得 {len(qs)}）")
    expect(abs(cnt["combo3-new"] - 19) <= 1 and abs(cnt["table"] - 4) <= 1,
           f"2026 MC 題型約 19/7/4（得 新格式{cnt['combo3-new']} 直接{cnt['direct']} 表格{cnt['table']} 其他{len(qs) - sum(cnt[k] for k in ('combo3-new', 'direct', 'table'))}）")
    nots = [q["no"] for q in qs if q["not"]]
    expect(21 in nots, f"2026 MC NOT 題含 Q21（得 {nots}）")
    ans, _ = read_md(next((base / "04_選擇題答案").glob("2025_*.md")))
    a = parse_mc_answers(ans)
    expect(len(a) == 30, f"2025 MC 答案解析 30 題（得 {len(a)}）")

    r = Report("2026 DBQ")
    dbq_lines, _ = read_md(next((base / "02_資料題").glob("2026_*.md")))
    check_dbq(sections(dbq_lines)[1], r)
    totals = [d for s, i, d in r.rows if "總分" in i]
    expect(len(totals) == 3 and all(d.startswith("10") for d in totals), f"2026 DBQ 三題各 10 分（得 {totals}）")
    chars = [int(d.split("（")[0]) for s, i, d in r.rows if "資料字數" in i]
    expect(len(chars) == 3 and all(abs(c - t) <= 25 for c, t in zip(chars, [246, 279, 259])),
           f"2026 DBQ 資料字數約 246/279/259（得 {chars}）")

    r = Report("2026 LQ")
    lq_lines, _ = read_md(next((base / "03_論述題").glob("2026_*.md")))
    check_lq(sections(lq_lines)[2], r)
    bgs = [int(d.split("（")[0]) for s, i, d in r.rows if "背景字數" in i]
    expect(len(bgs) == 5 and all(abs(b - t) <= 4 for b, t in zip(bgs, [57, 64, 60, 71, 66])),
           f"2026 LQ 背景字數約 57/64/60/71/66（得 {bgs}）")
    marks_ok = all(s == "✓" for s, i, d in r.rows if "分值" in i)
    expect(marks_ok, "2026 LQ 五題分值均為 8+10／10+8")
    print("\n自我測試" + ("全部通過" if not fails else f"有 {fails} 項不符"))
    return fails


def main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    if argv[0] == "--selftest":
        return 1 if selftest() else 0
    failed = False
    for f in argv:
        r = check_file(f)
        r.show()
        failed |= r.failed
    print("\n結果：" + ("有 ✗ 項，須修改" if failed else "沒有 ✗ 項（⚠ 項須在草稿說明或修改）"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
