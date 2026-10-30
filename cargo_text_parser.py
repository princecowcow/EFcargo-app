"""
信杰對單系統 — 客戶標準格式「貼上即填格」原型 (v2)
==================================================
痛點：客戶經微信落單，單據太多，同事人手入資料入唔切。
方案：客戶按統一表格格式填（騰訊文檔）→ 同事直接 copy 貼入，
      或填 file_id + tab 直接讀騰訊文檔 → 系統自動拆字段填格。

真實統一格式（客戶填的表格，列頭 + 每行一條記錄）：
    日期 | 公司(客戶自己公司名) | 流水号(序號) | 供應商(進)/出貨客戶(對手方名)
    | 型号 | 數量(出貨打負號 e.g. -500) | 備註(DC/LOT/貨來源) | 入倉号/AWB

運行：
    pip install streamlit pandas
    streamlit run cargo_text_parser.py
"""

import re
import os
import json
import subprocess
from datetime import date
import pandas as pd
import streamlit as st
from typing import Dict, List, Optional

# --------------------------------------------------------------------------
# ① 發給客戶的標準填寫格式（騰訊文檔表格，統一格式）
# --------------------------------------------------------------------------
TEMPLATE_TABLE = (
    "【客戶填單格式】日期留空＝當天；公司打「公司-流水号」（例 HT-1，唔使再打 E.流水号）\n"
    "\n"
    "【一個型號】\n"
    "A. 日期 :\n"
    "D. 公司 :\n"
    "F. 進出 :\n"
    "G. 供應商(進)/出貨客戶 :\n"
    "H. 型号 :      I. 數量 :\n"
    "J. 備註 :\n"
    "K. 入倉号/AWB :\n"
    "\n"
    "【兩個型號以上】H.型号+I.數量 / J.備註 / K.入倉号 為一組，幾多個型號就 copy 幾多組：\n"
    "A. 日期 : 20261030\n"
    "D. 公司 : HT-1\n"
    "F. 進出 :\n"
    "G. 供應商(進)/出貨客戶 :\n"
    "H. 型号 :      I. 數量 :\n"
    "J. 備註 :\n"
    "K. 入倉号/AWB :\n"
    "H. 型号 :      I. 數量 :\n"
    "J. 備註 :\n"
    "K. 入倉号/AWB :\n"
    "\n"
    "【示例】\n"
    "A. 日期 : 20261030\n"
    "D. 公司 : HT-1\n"
    "F. 進出 : 提\n"
    "G. 供應商(進)/出貨客戶 : 赛特恩\n"
    "H. 型号 : MZILG15THBLA-00A07  I. 數量 : 76\n"
    "J. 備註 : 库存\n"
    "K. 入倉号/AWB : DC2503\n"
    "H. 型号 : MZILG15THBLA-00A08  I. 數量 : 100\n"
    "J. 備註 :\n"
    "K. 入倉号/AWB : DC2549\n"
    "\n"
    "（舊格式 E. 流水号 都照認；B.完成 / C.簽收 係內部欄，客戶唔使填）"
)

# 填好的示例（演示 / 測試用；B完成 / C簽收 係內部欄，客戶唔填，解析自動忽略）
EXAMPLE_TABLE = (
    "日期\t公司\t流水号\t進出\t供應商(進)/出貨客戶\t型号\t數量\t備註\t入倉号/AWB\n"
    "2026-11-01\t信杰電子\t1\t進\tACEPAK\tTHGBMTG5D1LBAIL\t5120\tDC2345/LOT A1B2\tWH-001\n"
    "2026-11-01\t信杰電子\t2\t出\t出貨客戶B\tK4A8G085WB\t-500\tDC2401/LOT C3D4\tWH-002\n"
    "2026-11-02\t信杰電子\t3\t進\t供應商C\tMTFC8GAMALBH\t2560\tDC2500\tWH-003"
)

# --------------------------------------------------------------------------
# ② 字段識別：列名 → 系統內部欄位（用「包含關鍵字 + 優先級」避免複雜列名誤判）
#    順序很重要：counterparty 必須排在 customer 之前，否則「出貨客戶」會被歸錯。
# --------------------------------------------------------------------------
HEADER_RULES = [
    ("date",         ["日期", "date"]),
    ("seq",          ["流水", "序號", "序号", "seq", "no"]),
    ("counterparty", ["供應商", "供应商", "出貨", "出货", "對手", "对手", "客戶名"]),
    ("customer",     ["公司", "客戶", "客户"]),
    ("mpn",          ["型号", "型號", "mpn", "model"]),
    ("qty",          ["數量", "数量", "qty", "pcs"]),
    ("remark",       ["備註", "备注", "remark", "note", "dc", "lot"]),
    ("awb",          ["入倉", "入仓", "awb", "入倉號"]),
]


def _normalize(col: str) -> str:
    """只保留字母、數字、中文，轉小寫。把「供應商(進)/出貨客戶」→「供應商進出貨客戶」。"""
    return re.sub(r"[^\w\u4e00-\u9fff]", "", col).lower()


def _classify_header(col: str) -> Optional[str]:
    """把一個列名映射到內部欄位；認唔到回 None。
    「進出」是獨立一列（F欄），精確匹配歸 inout，避免同 G欄「供應商(進)/出貨客戶」撞。"""
    c = _normalize(col)
    if c in ("進出", "进出"):
        return "inout"
    for field, keys in HEADER_RULES:
        if any(k in c for k in keys):
            return field
    return None


# --------------------------------------------------------------------------
# ③ 解析：表格（列頭 + 數據行）為主；舊 key:value 格式作回退
# --------------------------------------------------------------------------
def _split_row(line: str) -> List[str]:
    """一行拆成單元格：優先 tab（騰訊文檔複製），其次逗號（CSV），再次雙空格。"""
    line = line.rstrip("\n")
    if "\t" in line:
        return [c.strip() for c in line.split("\t")]
    if "," in line:
        return [c.strip() for c in line.split(",")]
    return [c.strip() for c in re.split(r"\s{2,}", line.strip())]


def _coerce(rec: Dict, warnings: List[str]) -> None:
    """數值欄位轉型：數量支援負號（出貨）。失敗留原字並告警。"""
    for num_field, label in (("qty", "數量"), ("seq", "流水号")):
        if rec.get(num_field):
            try:
                rec[num_field] = int(re.sub(r"[^\d-]", "", str(rec[num_field])))
            except ValueError:
                warnings.append(f"「{rec.get('mpn', rec.get('counterparty', '?'))}」的{label}不是數字：{rec[num_field]}")


def _parse_table(rows: List[List[str]], warnings: List[str]) -> Optional[List[Dict]]:
    """識別表頭 → 建 col 索引映射 → 逐行成記錄。唔似表格回 None。"""
    if not rows:
        return None
    header = rows[0]
    mapping: Dict[int, str] = {}
    for i, cell in enumerate(header):
        f = _classify_header(cell)
        if f:
            mapping[i] = f
    if len(mapping) < 2:          # 表頭都認唔到 → 唔係表格
        return None
    items: List[Dict] = []
    for r in rows[1:]:
        if not any(c.strip() for c in r):
            continue
        rec: Dict = {}
        for i, field in mapping.items():
            if i < len(r):
                rec[field] = r[i].strip()
        if rec.get("mpn") or rec.get("qty") or rec.get("counterparty"):
            _coerce(rec, warnings)
            items.append(rec)
        else:
            warnings.append(f"跳過無法辨識的資料行：{r}")
    return items


# 舊格式回退用的 key:value 別名（兼容客戶若用手打 key:value）
_KV_ALIASES = {
    "date": ["日期", "date"],
    "customer": ["公司", "客戶", "客户", "customer"],
    "seq": ["流水号", "流水號", "序號", "序号", "seq"],
    "inout": ["進出", "进出"],
    "counterparty": ["供應商", "供应商", "出貨客戶", "出货客户", "對手方", "对手方"],
    "mpn": ["型号", "型號", "mpn", "model"],
    "qty": ["數量", "数量", "qty"],
    "remark": ["備註", "备注", "remark"],
    "awb": ["入倉号", "入倉號", "入仓号", "awb"],
}


_NUM_LINE_RE = re.compile(r"^\s*(\d+)\s*[.、)．]\s*(.+)$")


def _parse_numbered_lines(raw: str, warnings: List[str]):
    """微信簡易格式：'1. 供應商A, 型號X, 5120, DC1234, WH-001' 每條一行（流水號開頭）。
    '公司：X' / '日期：X' 行作單頭。回 (items, header)；一條都認唔到回 (None, {})。"""
    items: List[Dict] = []
    header: Dict = {}
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith(("【", "#")):
            continue
        hm = re.match(r"^(公司|客戶|客户|日期)\s*[:：]\s*(.+)$", line)
        if hm:
            header["customer" if hm.group(1) != "日期" else "date"] = hm.group(2).strip()
            continue
        nm = _NUM_LINE_RE.match(line)
        if not nm:
            continue
        parts = [p.strip() for p in re.split(r"[,，]", nm.group(2))]
        if len(parts) < 3:
            continue
        rec: Dict = {"seq": int(nm.group(1))}
        for i, f in enumerate(("counterparty", "mpn", "qty", "remark", "awb")):
            if i < len(parts):
                rec[f] = parts[i]
        _coerce(rec, warnings)
        items.append(rec)
    return (items, header) if items else (None, {})


_LABEL_RE = re.compile(r"^\s*(?:[A-Za-z]\s*\.?\s*)?([\u4e00-\u9fffA-Za-z/（）()\-]+?)\s*[:：=]\s*(.*)$")


# 單頭 / 訂單級（同一單多型號共用）/ 細項級 欄位分類
_ORDER_LEVEL = {"seq", "inout", "counterparty"}
_ITEM_LEVEL = {"mpn", "qty", "remark"}


def _split_company_seq(v: str):
    """'HT-4' → (公司='HT', 流水号='4')；冇 -數字就原樣返。"""
    m = re.match(r"^(.+)-(\d+)$", v.strip())
    if m and m.group(1).strip():
        return m.group(1).strip(), m.group(2)
    return v.strip(), ""


_EMBED_LABEL_RE = re.compile(r"([A-Za-z])\s*\.\s*([\u4e00-\u9fffA-Za-z/（）()\-]+?)\s*[:：=]")


def _iter_line_labels(label: str, value: str):
    """一行可能有多個標籤（客戶會打 'H. 型号 :xxx  I. 數量 :76'）→ 拆做 [(label, value), ...]。"""
    ms = list(_EMBED_LABEL_RE.finditer(value))
    if not ms:
        return [(label, value)]
    out = [(label, value[:ms[0].start()].strip())]
    for i, mm in enumerate(ms):
        v_end = ms[i + 1].start() if i + 1 < len(ms) else len(value)
        out.append((mm.group(2), value[mm.end():v_end].strip()))
    return out


def _parse_labeled_blocks(raw: str, warnings: List[str]):
    """垂直標籤格式：'A. 日期 : 2026-11-01' / 'D. 公司 : HT-4' 每行一欄（一行多標籤都認）。
    D. 公司 末尾 -數字 會拆做流水号（例：HT-4 → 公司 HT、流水号 4，唔使再打 E）。
    同一單出多型號：重複嘅 H/I/J 自動拆成多行（共用 日期/公司/流水号/進出/供應商）。
    date / company 當單頭；備註等多行內容會併入上一個欄位。回 (items, header)。"""
    items: List[Dict] = []
    header: Dict = {}
    cur: Dict = {}
    last_field = ""  # 最後一個配到標籤嘅欄位（連續行歸入同一格）

    def _assign(field: str, value: str):
        nonlocal cur, last_field
        if field == "customer":
            comp, seqnum = _split_company_seq(value)
            header["customer"] = comp
            if seqnum and not cur.get("seq"):
                cur["seq"] = seqnum
        elif field == "date":
            header[field] = value
        elif field == "seq":
            # 新流水號 + 上一件已有細項 → 新單（全部重置）
            if cur.get("mpn") or cur.get("qty") or cur.get("remark") or cur.get("counterparty"):
                _coerce(cur, warnings)
                items.append(cur)
                cur = {}
            cur["seq"] = value
        else:
            # 型号(H) 重複出現 = 同一單出第二個型號 → 新細項（保留 訂單級 欄位）
            if field == "mpn" and cur.get("mpn"):
                _coerce(cur, warnings)
                items.append(cur)
                cur = {k: v for k, v in cur.items() if k in _ORDER_LEVEL}
            cur[field] = value
        last_field = field

    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith(("【", "#")):
            continue
        m = _LABEL_RE.match(line)
        if m:
            pairs = _iter_line_labels(m.group(1), m.group(2).strip())
            if all(not _classify_header(lb) for lb, _ in pairs) and re.match(r"^[A-Za-z]\s*\.", line):
                pass  # 成行都係認唔到嘅標籤（例：'B. 完成 : …'）→ 略過
            else:
                for lb, val in pairs:
                    fld = _classify_header(lb)
                    if fld:
                        _assign(fld, val)
                    elif not re.match(r"^[A-Za-z]\s*\.", lb):
                        # 認唔到又唔似內部標籤（冇 'X.' 前綴）→ 當備註一部分
                        if last_field and last_field not in ("date", "customer"):
                            cur[last_field] = (cur.get(last_field, "") + "\n" + (lb + "：" + val).strip()).strip()
            continue
        # 無冒號嘅連續行（客戶備註好多時多行）→ 併入上一個欄位
        if last_field and last_field not in ("date", "customer"):
            cur[last_field] = (cur.get(last_field, "") + "\n" + line).strip()
    if cur:
        _coerce(cur, warnings)
        items.append(cur)
    return (items, header) if items else (None, {})


def _parse_kv_blocks(raw: str, warnings: List[str]) -> List[Dict]:
    """回退：每行 '型号: xxx, 數量: 5120' 這類 key:value，或 --- 分單頭/明細。"""
    lookup = {_normalize(a): f for f, als in _KV_ALIASES.items() for a in als}
    items: List[Dict] = []
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("---"):
            continue
        rec: Dict = {}
        for part in re.split(r"\s*[\|\u3001,;]\s*", line):
            m = re.match(r"^(.*?)[:：=]\s*(.*)$", part.strip())
            if m:
                f = lookup.get(_normalize(m.group(1)))
                if f:
                    rec[f] = m.group(2).strip()
        if rec.get("mpn") or rec.get("qty") or rec.get("counterparty"):
            _coerce(rec, warnings)
            items.append(rec)
    return items


def parse_text(raw: str) -> Dict:
    """核心：文字 → 結構化 {header, items, warnings}。"""
    result: Dict = {"header": {}, "items": [], "warnings": []}
    if not raw or not raw.strip():
        result["warnings"].append("貼上的內容是空的。")
        return result

    rows = [_split_row(l) for l in raw.splitlines() if l.strip()]
    items = None
    extra_header: Dict = {}
    # 1) 垂直標籤格式 A. 日期 : 值
    items, extra_header = _parse_labeled_blocks(raw, result["warnings"])
    # 2) 表格（表頭 + 數據行）
    if items is None:
        items = _parse_table(rows, result["warnings"])
    # 3) 簡易逐行 1. 供應商, 型號...
    if items is None:
        items, h2 = _parse_numbered_lines(raw, result["warnings"])
        extra_header.update(h2)
    # 4) 舊 key:value
    if items is None:
        items = _parse_kv_blocks(raw, result["warnings"])

    # 單頭（日期/公司）：先取簡易格式的單頭行，再從明細首個非空值補
    header: Dict = dict(extra_header)
    for f in ("date", "customer"):
        if header.get(f):
            continue
        for it in items:
            if it.get(f):
                header[f] = it[f]
                break
    result["items"] = items
    # 日期可選：客戶不打就預設填當天（都是當天的單）
    if not header.get("date"):
        header["date"] = date.today().isoformat()
    result["header"] = header
    if not items:
        result["warnings"].append("沒有解析到任何貨品記錄，請檢查格式（要有表頭 + 數據行）。")
    return result


# --------------------------------------------------------------------------
# ④ 直連騰訊文檔 MCP（OpenAPI）：讀 + 写回，复用你 sync.js 同款 token
#    —— 唔再依賴 mcporter，只要有 TENCENT_DOCS_TOKEN 就掂
# --------------------------------------------------------------------------
import time
import urllib.request
import urllib.error

_TDOCS_EP = os.environ.get("TDOCS_ENDPOINT", "https://docs.qq.com/openapi/mcp")


def _load_tdoc_token() -> str:
    """優先 env，其次 Streamlit secrets，再本機 .env（含 Desktop/mpc-sync/.env）。"""
    return _read_secret("TENCENT_DOCS_TOKEN")


# --------------------------------------------------------------------------
# 密碼登入門（出公網時擋住閒人；密碼來源：環境變數 → Streamlit secrets → 本機 .env）
# --------------------------------------------------------------------------
def _read_secret(key: str) -> str:
    """優先 env，其次 Streamlit secrets，再本機 .env。"""
    v = os.environ.get(key, "").strip()
    if v:
        return v
    try:
        v = str(st.secrets.get(key, "")).strip()
        if v:
            return v
    except Exception:
        pass
    for pp in (os.path.join(os.path.dirname(__file__), ".env"),
               r"C:\Users\ExpertFair12D2\Desktop\mpc-sync\.env"):
        if os.path.exists(pp):
            for line in open(pp, encoding="utf-8"):
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, val = line.split("=", 1)
                if k.strip() == key:
                    return val.strip()
    return ""


def _app_password() -> str:
    return _read_secret("APP_PASSWORD")


def _login_guard():
    """未登入就只顯示密碼框；登入後先 render 主程式。"""
    if st.session_state.get("authed"):
        return
    st.markdown("## 🔒 登入")
    st.caption("此工具可寫入客戶騰訊文檔，請輸入密碼先使用。")
    pw = st.text_input("密碼", type="password", key="login_pw")
    if st.button("登入", type="primary"):
        correct = _app_password()
        if not correct:
            st.error("⚠️ 尚未設定 APP_PASSWORD：請於環境變數 / Streamlit secrets / 本機 .env 加一行 APP_PASSWORD=你的密碼，先可以使用。")
        elif pw == correct:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("密碼錯誤")
    st.stop()

class TencentDocsMCP:
    """极简 MCP client：initialize + tools/call。对应 sync.js 嘅 rpc/mcp。"""

    def __init__(self, token: str, ep: str = _TDOCS_EP):
        self.token = token
        self.ep = ep
        self.auth = "Bearer " + token if token else ""
        self.session = None
        self.tools = None

    def _post(self, method, params=None, notify=False):
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
            "Authorization": self.auth,
        }
        if self.session:
            headers["Mcp-Session-Id"] = self.session
        body = json.dumps({
            "jsonrpc": "2.0", "id": int(time.time() * 1000),
            "method": method, "params": params or {},
        }).encode("utf-8")
        req = urllib.request.Request(self.ep, data=body, headers=headers, method="POST")
        try:
            resp = urllib.request.urlopen(req, timeout=60)
        except urllib.error.HTTPError as e:
            raise RuntimeError(f"HTTP {e.code}：{e.read().decode('utf-8', 'replace')[:300]}")
        sid = resp.headers.get("mcp-session-id")
        if sid:
            self.session = sid
        text = resp.read().decode("utf-8", "replace")
        if notify:
            return None
        obj = self._parse(text)
        if obj and obj.get("error"):
            msg = obj["error"].get("message", str(obj["error"]))
            if "file not exist" in msg or "10016" in msg:
                msg += "（→ file_id 唔啱，或者個文檔未 share 俾 TENCENT_DOCS_TOKEN 個帳號）"
            elif "sheet_id" in msg and "missing" in msg:
                msg += "（→ 呢個客戶未揀子表，去「⚙️ 管理客戶」撳「列出子表」揀頁）"
            raise RuntimeError(msg)
        return obj.get("result") if obj else None

    @staticmethod
    def _parse(text):
        text = text.strip()
        if text.startswith("{"):
            try:
                return json.loads(text)
            except Exception:
                pass
        lines = [l for l in text.splitlines() if l.startswith("data:")]
        if lines:
            try:
                return json.loads(lines[-1][5:].strip())
            except Exception:
                return None
        return None

    def _unwrap(self, r):
        if isinstance(r, dict):
            if "structuredContent" in r:
                return r["structuredContent"]
            c = r.get("content")
            if c and isinstance(c[0], dict) and isinstance(c[0].get("text"), str):
                try:
                    return json.loads(c[0]["text"])
                except Exception:
                    return c[0]["text"]
        return r

    def ensure_init(self):
        if self.tools is not None:
            return
        try:
            r = self._post("tools/list", {})
        except Exception as e:
            if "initialize" in str(e).lower():
                self._post("initialize", {"protocolVersion": "2025-03-26",
                                          "capabilities": {}, "clientInfo": {"name": "cargo-parser", "version": "1"}})
                try:
                    self._post("notifications/initialized", {}, notify=True)
                except Exception:
                    pass
                r = self._post("tools/list", {})
            else:
                raise
        self.tools = [t["name"] for t in (r or {}).get("tools", [])]

    def _find(self, *cands):
        names = self.tools or []
        for c in cands:
            hit = next((n for n in names if n == c), None)
            if not hit:
                hit = next((n for n in names if n.endswith(c)), None)
            if not hit:
                hit = next((n for n in names if c in n), None)
            if hit:
                return hit
        return None

    def _call(self, cands, args):
        name = self._find(*cands)
        if not name:
            raise RuntimeError("搵唔到工具：" + " / ".join(cands))
        return self._unwrap(self._post("tools/call", {"name": name, "arguments": args}))

    # ---- 讀 ----
    def _sheet_row_count(self, file_id, sheet_id) -> int:
        info = self._call(["get_sheet_info"], {"file_id": file_id})
        sheets = (info.get("sheets") or info.get("sheet_list") or []) if isinstance(info, dict) else []
        for s in sheets:
            if (s.get("sheet_id") or s.get("id")) == sheet_id:
                return int(s.get("row_count") or 1000)
        return 1000

    @staticmethod
    def _parse_csv(text) -> List[List[str]]:
        """同 sync.js parseCsv 一致：處理引號包住嘅欄。"""
        rows: List[List[str]] = []
        row: List[str] = []
        field = ""
        in_q = False
        i = 0
        while i < len(text):
            ch = text[i]
            if in_q:
                if ch == '"':
                    if text[i + 1: i + 2] == '"':
                        field += '"'
                        i += 1
                    else:
                        in_q = False
                else:
                    field += ch
            elif ch == '"':
                in_q = True
            elif ch == ",":
                row.append(field)
                field = ""
            elif ch == "\n":
                row.append(field)
                rows.append(row)
                row, field = [], ""
            elif ch != "\r":
                field += ch
            i += 1
        if field != "" or row:
            row.append(field)
            rows.append(row)
        return rows

    def read_grid(self, file_id, sheet_id) -> List[List[str]]:
        """讀晒所有有資料嘅行（同 sync.js readAllRows 一樣逐級擴大範圍）。"""
        row_count = self._sheet_row_count(file_id, sheet_id)
        end = min(row_count, 200)
        rows: List[List[str]] = []
        for _ in range(20):
            d = self._call(["get_cell_data"], {
                "file_id": file_id, "sheet_id": sheet_id,
                "start_row": 0, "start_col": 0,
                "end_row": end - 1, "end_col": 12, "return_csv": True,
            })
            csv_text = d.get("csv_data") if isinstance(d, dict) else None
            if csv_text is None:
                raise RuntimeError("讀唔到文檔內容（冇 csv_data，確認 tab_id 正確）")
            rows = self._parse_csv(csv_text)
            last = -1
            for i, r in enumerate(rows):
                if any(str(c).strip() for c in r):
                    last = i
            if last < 0:
                return []
            if last >= len(rows) - 1 and end < row_count:
                end = min(row_count, end * 2)
                continue
            return rows[:last + 1]
        return rows

    def read_csv_text(self, file_id, sheet_id) -> str:
        grid = self.read_grid(file_id, sheet_id)
        return "\n".join(
            ",".join(f'"{str(v).replace(chr(34), chr(34) * 2)}"' for v in r) for r in grid
        )

    def last_used_row(self, file_id, sheet_id) -> int:
        """最後一行有資料嘅 row index（0-based）。空表回 -1。"""
        grid = self.read_grid(file_id, sheet_id)
        return len(grid) - 1 if grid else -1

    # ---- 写（append）----
    def append_rows(self, file_id, sheet_id, rows):
        start = self.last_used_row(file_id, sheet_id) + 1  # 加喺最尾
        csv = "\n".join(
            ",".join(f'"{str(v).replace(chr(34), chr(34) * 2)}"' for v in r) for r in rows
        )
        csv_tool = self._find("set_range_value_by_csv")
        if csv_tool:
            self._call([csv_tool], {"file_id": file_id, "sheet_id": sheet_id,
                                    "start_row": start, "start_col": 0, "csv_data": csv})
            return start
        # fallback：逐格 set_range_value
        vals = []
        for ri, r in enumerate(rows):
            for ci, v in enumerate(r):
                s = "" if v is None else str(v)
                num = s.replace(",", "")
                if s != "" and num.lstrip("-").isdigit():
                    vals.append({"row": start + ri, "col": ci, "value_type": "NUMBER", "number_value": int(num)})
                else:
                    vals.append({"row": start + ri, "col": ci, "value_type": "STRING", "string_value": s})
        self._call(["set_range_value"], {"file_id": file_id, "sheet_id": sheet_id, "values": vals})
        return start


def read_tencent_sheet(file_id: str, sheet_id: str) -> str:
    """直連騰訊文檔 MCP 拉取子表 CSV 文字（免 mcporter）。"""
    token = _load_tdoc_token()
    if not token:
        raise RuntimeError("未搵到 TENCENT_DOCS_TOKEN：請設 env 或放 .env（參考你 Desktop/mpc-sync/.env）")
    sheet_id = resolve_sheet_id(file_id, sheet_id)
    mcp = TencentDocsMCP(token)
    mcp.ensure_init()
    return mcp.read_csv_text(file_id, sheet_id)


# 写回用嘅目標欄位順序：預設跟 HT 文檔月份頁實際 13 欄：
# 日期 | 完成 | 簽收 | 公司 | 序号 | 進出 | 供應商(進)/出貨客戶 | 型号 | 數量 | 備註 | 电话地址 | 入倉号/AWB | (圖片/OCR)
# 空字串 = 該欄留空。其他客戶文檔欄位唔同嘅話，喺 clients.csv 加一欄 cols（逗號分隔 field key）。
DEFAULT_DOC_COLS = "date,,,customer,seq,inout,counterparty,mpn,qty,remark,,awb,"


def _cols_from_spec(spec: str) -> List[str]:
    keys = [c.strip() for c in (spec or "").split(",")]
    return keys if any(keys) else DEFAULT_DOC_COLS.split(",")


def _fmt_doc_date(v) -> str:
    """各種客戶日期寫法 → 10月29日（跟文檔統一格式）；認唔出就照出。"""
    s = str(v or "").strip()
    m = re.match(r"^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?$", s)  # 2026-10-29 / 2026/10/29 / 2026年10月29日
    if not m:
        m = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)  # 20261029（無分隔符）
    if m:
        return f"{int(m.group(2))}月{int(m.group(3))}日"
    return s


_DOC_FIELDS = {
    "date":        lambda h, it: _fmt_doc_date(h.get("date", "")),
    "customer":    lambda h, it: h.get("customer", ""),
    "seq":         lambda h, it: it.get("seq", ""),
    "inout":       lambda h, it: it.get("inout", ""),
    "counterparty": lambda h, it: it.get("counterparty", ""),
    "mpn":         lambda h, it: it.get("mpn", ""),
    "qty":         lambda h, it: it.get("qty", ""),
    "remark":      lambda h, it: it.get("remark", ""),
    "awb":         lambda h, it: it.get("awb", ""),
}


def build_doc_rows(header: Dict, items: List[Dict], cols_spec: str = "") -> List[List[str]]:
    """寫落文檔嘅行。同一單（同序号）後續行嘅 公司/序号/進出 留空（跟你哋平時手寫習慣，
    唔會成頁重複晒單頭）。"""
    cols = _cols_from_spec(cols_spec)
    out = []
    prev_seq = None
    for it in items:
        seq = str(it.get("seq", "") or "").strip()
        cont = bool(seq) and seq == prev_seq  # 同一單延續行
        out.append([
            "" if (cont and c in ("customer", "seq", "inout"))
            else _DOC_FIELDS.get(c, lambda h, i: "")(header, it)
            for c in cols
        ])
        prev_seq = seq
    return out


def _inout_cell_style(v) -> str:
    """進出格 → 跟文檔標籤顏色：送/寄 類粉紅，提/收 類綠。"""
    s = re.sub(r"\s+", "", str(v or "")).translate(_S2T)
    if any(k in s for k in ("送貨", "寄貨", "寄付", "轉倉(出)")) or "寄" in s:
        return "background-color:#fde3e3;color:#c0392b"
    if any(k in s for k in ("提貨", "自提", "收貨", "收票", "轉倉(收)")):
        return "background-color:#dcf3e3;color:#1e8449"
    return ""


def _neg_qty_style(v) -> str:
    """數量格 → 負數紅字粗體（跟文檔）。"""
    try:
        return "color:#c0392b;font-weight:bold" if float(str(v).strip()) < 0 else ""
    except (ValueError, TypeError):
        return ""


def _style_prev_df(df: pd.DataFrame):
    """⑦ 預覽套用文檔風格。"""
    sty = df.style
    for c in df.columns:
        cs = str(c)
        if "進出" in cs:
            sty = sty.map(_inout_cell_style, subset=[c]) if hasattr(sty, "map") else sty.applymap(_inout_cell_style, subset=[c])
    for c in df.columns:
        if "數量" in str(c):
            sty = sty.map(_neg_qty_style, subset=[c]) if hasattr(sty, "map") else sty.applymap(_neg_qty_style, subset=[c])
    return sty


def write_items_to_doc(file_id: str, sheet_id: str, header: Dict, items: List[Dict],
                       cols_spec: str = "") -> int:
    """解析結果 append 落客戶專用文檔，回傳寫入起始 row（0-based）。
    跟上一行唔同日子 → 先自動加一行紅色分隔（寫日期 + 塗紅底），資料寫喺後面，
    咁樣分隔行有值，append 定位唔會撞正佢。"""
    token = _load_tdoc_token()
    if not token:
        raise RuntimeError("未搵到 TENCENT_DOCS_TOKEN：請設 env 或放 .env（參考你 Desktop/mpc-sync/.env）")
    sheet_id = resolve_sheet_id(file_id, sheet_id)
    rows = build_doc_rows(header, items, cols_spec)
    if not rows:
        raise RuntimeError("冇貨品可寫入")
    mcp = TencentDocsMCP(token)
    mcp.ensure_init()

    # 紅色日期分隔行：上一行日期同今次唔同（新一日）先加
    today = _fmt_doc_date(header.get("date", ""))
    if today:
        try:
            grid = mcp.read_grid(file_id, sheet_id)
        except Exception:
            grid = []
        if grid:
            last = grid[-1] if grid else []
            last_date = _fmt_doc_date(last[0]) if last else ""
            if last_date and last_date != today:
                ncol = len(rows[0])
                divider = [today] + [""] * (ncol - 1)
                dstart = mcp.append_rows(file_id, sheet_id, [divider])
                try:  # 塗紅底（成行 13 欄）；塗唔到都唔阻住寫資料
                    mcp._call(["set_cell_style"], {
                        "file_id": file_id, "sheet_id": sheet_id,
                        "start_row": dstart, "end_row": dstart,
                        "start_col": 0, "end_col": ncol - 1,
                        "bg_color": "FFFF0000",
                    })
                except Exception:
                    pass

    return mcp.append_rows(file_id, sheet_id, rows)


# --------------------------------------------------------------------------
# ⑤ LLM 兜底接口（保留）：客戶亂填、規則解析不到時交 OpenRouter 文字模型
# --------------------------------------------------------------------------
def parse_text_with_llm(raw: str, llm_fn=None) -> Dict:
    res = parse_text(raw)
    if not res["items"] and llm_fn:
        try:
            llm_json = llm_fn(raw)
            if isinstance(llm_json, dict) and llm_json.get("items"):
                llm_json.setdefault("warnings", ["（此筆由 LLM 智能解析）"])
                return llm_json
        except Exception as e:  # noqa
            res["warnings"].append(f"LLM 解析失敗，沿用規則結果：{e}")
    return res


# --------------------------------------------------------------------------
# ⑤b 客戶映射：客戶名 ↔ 騰訊文檔 file_id / tab_id
# --------------------------------------------------------------------------
import os

CLIENTS_CSV = os.path.join(os.path.dirname(__file__), "clients.csv")


def load_clients() -> pd.DataFrame:
    """載入 clients.csv（客戶名 / file_id / tab_id）；不存在則建示例。"""
    if not os.path.exists(CLIENTS_CSV):
        pd.DataFrame([{"client_name": "信杰電子", "file_id": "DTmdNSW1nY2RvRUdM", "tab_id": "i6h4jj"}]).to_csv(CLIENTS_CSV, index=False)
    return pd.read_csv(CLIENTS_CSV, dtype=str).fillna("")


def client_doc_url(file_id: str, tab_id: str) -> str:
    url = f"https://docs.qq.com/sheet/{file_id}"
    if tab_id:
        url += f"?tab={tab_id}"
    return url


def list_sheets(file_id: str) -> List[Dict]:
    """列出文檔入面所有子表：[{sheet_id, sheet_name}, ...]（用 get_sheet_info）。"""
    token = _load_tdoc_token()
    if not token:
        raise RuntimeError("未搵到 TENCENT_DOCS_TOKEN：請設 env 或放 .env")
    mcp = TencentDocsMCP(token)
    mcp.ensure_init()
    info = mcp._call(["get_sheet_info"], {"file_id": file_id})
    sheets = (info.get("sheets") or info.get("sheet_list") or []) if isinstance(info, dict) else []
    out = []
    for s in sheets:
        sid = s.get("sheet_id") or s.get("id")
        nm = s.get("sheet_name") or s.get("name") or sid
        if sid:
            out.append({"sheet_id": sid, "sheet_name": nm})
    return out


def resolve_sheet_id(file_id: str, sheet_id: str) -> str:
    """tab_id 冇填：得一個子表就自動用佢；多過一個就拒絕（防寫錯頁）。"""
    if sheet_id:
        return sheet_id
    sheets = list_sheets(file_id)
    if not sheets:
        raise RuntimeError("呢個文檔搵唔到任何子表（確認 file_id 正確）")
    if len(sheets) == 1:
        return sheets[0]["sheet_id"]
    raise RuntimeError(
        f"呢個文檔有 {len(sheets)} 個子表，唔知寫邊頁。去「⚙️ 管理客戶」撳「📋 列出子表」揀定一頁再保存。"
    )


def save_client(name: str, file_id: str, tab_id: str, cols: str = "") -> None:
    """新增或按客戶名修改（同名就覆寫 file_id/tab_id），寫回 clients.csv。"""
    df = load_clients()
    df = df[df["client_name"] != name]
    df = pd.concat(
        [df, pd.DataFrame([{"client_name": name, "file_id": file_id, "tab_id": tab_id, "cols": cols}])],
        ignore_index=True,
    )
    df.to_csv(CLIENTS_CSV, index=False)


def delete_client(name: str) -> None:
    df = load_clients()
    df = df[df["client_name"] != name]
    df.to_csv(CLIENTS_CSV, index=False)


@st.cache_data(ttl=300, show_spinner=False)
def _list_sheets_cached(file_id: str) -> List[tuple]:
    """[(sheet_id, sheet_name), ...]；失敗回空。"""
    try:
        return [(s["sheet_id"], s["sheet_name"]) for s in list_sheets(file_id)]
    except Exception:
        return []


@st.cache_data(ttl=60, show_spinner=False)
def _target_header_cached(file_id: str, sheet_id: str) -> List[str]:
    """目標子表嘅表頭行（⑦預覽欄名用，一眼睇到有冇錯位）。"""
    try:
        mcp_ = TencentDocsMCP(_load_tdoc_token())
        mcp_.ensure_init()
        grid = mcp_.read_grid(file_id, sheet_id)
        return [str(x) for x in grid[0]] if grid else []
    except Exception:
        return []


# --------------------------------------------------------------------------
# ⑥ Streamlit UI
# --------------------------------------------------------------------------
ITEM_COLS = [("seq", "流水号"), ("inout", "進出"), ("counterparty", "供應商(進)/出貨客戶"),
             ("mpn", "型号"), ("qty", "數量(出貨-)"), ("remark", "備註"), ("awb", "入倉号/AWB")]

# 進出下拉選項（跟客戶專用文檔入面嘅下拉一致）
INOUT_OPTIONS = ["收貨", "送貨", "自提", "提貨", "轉倉(收)", "轉倉(出)",
                 "寄貨DHL", "寄貨FEDEX", "寄貨UPS", "順豐寄付", "順豐到付",
                 "收貨 FEDEX", "收貨 DHL", "收貨 UPS", "收貨 順豐", "收票"]

_INOUT_KW = [  # 客戶亂打時嘅關鍵字對照（特定 courier 排先，通用排後）
    ("轉倉出", "轉倉(出)"), ("轉倉收", "轉倉(收)"), ("轉倉", "轉倉(收)"),
    ("寄dhl", "寄貨DHL"), ("寄fedex", "寄貨FEDEX"), ("寄ups", "寄貨UPS"),
    ("收fedex", "收貨 FEDEX"), ("收dhl", "收貨 DHL"), ("收ups", "收貨 UPS"), ("收順豐", "收貨 順豐"),
    ("寄付", "順豐寄付"), ("到付", "順豐到付"), ("收票", "收票"),
    ("收貨", "收貨"), ("送貨", "送貨"), ("自提", "自提"), ("提貨", "提貨"), ("提", "提貨"),
]


_S2T = str.maketrans("货仓丰转递", "貨倉豐轉遞")  # 簡體字 → 繁體，方便比對


def _match_inout(v) -> str:
    """客戶打嘅「進出」→ 盡量對返文檔下拉嘅標準選項；對唔到就保留原字（同事可下拉改）。"""
    s = re.sub(r"\s+", "", str(v or "")).translate(_S2T)
    if not s:
        return ""
    for o in INOUT_OPTIONS:  # 直接對（忽略空格／大小寫／簡繁）
        if s == re.sub(r"\s+", "", o).translate(_S2T) or s.lower() == re.sub(r"\s+", "", o).lower().translate(_S2T):
            return o
    sl = s.lower()
    for k, o in _INOUT_KW:  # 關鍵字
        if k in sl or k in s:
            return o
    for brand, recv, ship in (("fedex", "收貨 FEDEX", "寄貨FEDEX"),
                              ("dhl", "收貨 DHL", "寄貨DHL"),
                              ("ups", "收貨 UPS", "寄貨UPS")):  # 倒裝 e.g.「FEDEX收」
        if brand in sl:
            return recv if "收" in s else ship
    return str(v).strip()


def main():
    st.set_page_config(page_title="信杰對單 · 客戶格式自動填格", layout="wide")
    st.title("📋 客戶單據「貼上即填格」")
    _login_guard()  # 未登入只顯示密碼框
    if st.sidebar.button("🚪 登出", help="測試密碼門用；登出後要再打密碼"):
        st.session_state["authed"] = False
        st.rerun()
    st.caption("客戶按統一表格格式填（騰訊文檔）→ 同事貼入 / 直接讀文檔 → 自動拆字段填格 → 路由到客戶專用文檔")

    # ① 客戶選擇（路由到該客戶專用文檔）
    st.markdown("### ① 揀客戶名（路由到該客戶專用文檔）")
    clients = load_clients()
    cfid = ctid = ccols = ""
    if clients.empty:
        st.warning("clients.csv 而家冇任何客戶——去下面「⚙️ 管理客戶」加返一個（客戶名 + file_id + 揀子表），加完刷新再揀。")
    else:
        sel = st.selectbox("客戶", clients["client_name"].tolist(), key="client_sel")
        _selrow = clients[clients["client_name"] == sel]
        if _selrow.empty:
            # 下拉仲記住已刪除/已改名嘅舊客戶 → 暫用第一個客戶，唔好爆錯
            st.caption(f"⚠️ 之前揀嘅「{sel}」已經唔存在（刪咗或改咗名），暫用第一個客戶，請重新揀。")
            _selrow = clients.iloc[[0]]
        crow = _selrow.iloc[0]
        cfid, ctid = str(crow.get("file_id", "")), str(crow.get("tab_id", ""))
        if cfid and not ctid:
            _sheets = _list_sheets_cached(cfid)
            if len(_sheets) == 1:
                ctid = _sheets[0][0]
            elif len(_sheets) > 1:
                _opts = ["（未揀——寫回前必須揀）"] + [f"{nm}　(id: {sid})" for sid, nm in _sheets]
                _pick = st.selectbox("⚠️ 呢個客戶未設定子表，即場揀要寫邊頁：", _opts, key="tmp_sheet_pick")
                if not _pick.startswith("（未揀"):
                    _sid, _nm = _sheets[_opts.index(_pick) - 1]
                    ctid = _sid
                    st.caption(f"今次會寫入「{_nm}」。記得去「⚙️ 管理客戶」→「📋 列出子表」揀定佢並保存。")
        ccols = str(crow.get("cols", "") or "")
    c1, c2 = st.columns([3, 1])
    c1.caption(f"file_id: `{cfid}`　tab: `{ctid}`")
    c2.markdown(f"[🔗 打開客戶專用文檔]({client_doc_url(cfid, ctid)})")

    # ①b 管理客戶（同事唔使改 CSV）
    with st.expander("⚙️ 管理客戶（新增 / 修改 / 刪除）", expanded=False):
        st.info("同事唔使改 CSV：填客戶名 + file_id，撳「列出子表」揀頁就自動填 tab_id，撳保存即寫入 clients.csv。")
        nc_name = st.text_input("客戶名", key="nc_name")
        nc_fid = st.text_input("文檔 file_id", key="nc_fid")
        if st.button("📋 列出子表（先填 file_id）", key="btn_listsheets"):
            if not nc_fid.strip():
                st.warning("先填 file_id 先撳")
            else:
                try:
                    sheets = list_sheets(nc_fid.strip())
                    if sheets:
                        st.session_state["sheet_opts"] = sheets
                        st.success(f"搵到 {len(sheets)} 個子表，落去揀。")
                    else:
                        st.warning("呢個文檔搵唔到子表（確認 file_id 正確）")
                except Exception as e:
                    st.error(str(e))
        sheets = st.session_state.get("sheet_opts", [])
        labels = [f"{s['sheet_name']}  (id: {s['sheet_id']})" for s in sheets]

        def _pick_sheet():
            p = st.session_state.get("sheet_pick")
            if p in labels:
                st.session_state["nc_tid"] = sheets[labels.index(p)]["sheet_id"]

        if sheets:
            st.selectbox("揀要填既子表（揀完自動填落 tab_id）", labels, key="sheet_pick", on_change=_pick_sheet)
            if st.session_state.get("nc_tid"):
                st.caption(f"✅ tab_id 已自動填：`{st.session_state['nc_tid']}`")
        nc_tid = st.text_input("子表 tab_id（揀上面會自動填，也可手填）", value=st.session_state.get("nc_tid", ""), key="nc_tid")
        nc_cols = st.text_input(
            "欄位順序（進階，留空＝預設 13 欄）",
            value=st.session_state.get("nc_cols", ""),
            key="nc_cols",
            help="逗號分隔：date,完成,簽收,customer,seq,inout,counterparty,mpn,qty,remark,電話,awb,圖片。留空位打唔落 field 名就得。",
        )
        cs, cd = st.columns(2)
        with cs:
            if st.button("💾 保存客戶", key="btn_saveclient", use_container_width=True):
                if not nc_name.strip() or not nc_fid.strip():
                    st.warning("客戶名 同 file_id 必填")
                else:
                    tid = nc_tid.strip()
                    if not tid and sheets:  # 冇手填 tab_id 就用上面揀既子表
                        p = st.session_state.get("sheet_pick", "")
                        if p in labels:
                            tid = sheets[labels.index(p)]["sheet_id"]
                    save_client(nc_name.strip(), nc_fid.strip(), tid, nc_cols.strip())
                    st.success(f"已保存「{nc_name.strip()}」（tab_id: {tid or '（空——寫回前要揀子表）'}）。上方客戶列表已更新，解析前先重選佢。")
                    st.rerun()
        with cd:
            del_name = st.selectbox("刪除客戶", clients["client_name"].tolist(), key="del_name")
            if st.button("🗑 刪除", key="btn_delclient", use_container_width=True):
                delete_client(del_name)
                st.rerun()

    with st.expander("② 發給客戶的統一格式（微信/文檔通用）", expanded=False):
        st.code(TEMPLATE_TABLE, language="text")
        st.info("三種格式都認：① 上面「A.日期 / D.公司 …」垂直標籤格式（微信直接打，最推介）；② 騰訊文檔表格（表頭：日期/公司/流水号/供應商(進)/出貨客戶/型号/數量/備註/入倉号/AWB）；③ 簡易一行式「1. 供應商, 型號, 數量…」。")

    # ③ 直接讀騰訊文檔（免複製）
    with st.expander("③ 或直接從騰訊文檔讀取（免複製貼上）"):
        c1, c2, c3 = st.columns([2, 2, 1])
        fid = c1.text_input("文檔 file_id", value=cfid or "DTmdNSW1nY2RvRUdM")
        sid = c2.text_input("子表 tab_id", value=ctid or "i6h4jj")
        if c3.button("📥 讀取文檔"):
            try:
                csv_text = read_tencent_sheet(fid.strip(), sid.strip())
                st.session_state["raw"] = csv_text
                st.rerun()
            except Exception as e:
                st.error(str(e))

    # ④ 貼上區（預設空白，等同事直接貼客戶填嘅內容）
    st.markdown("### ④ 同事把客戶填好的表格貼到這裡（或直接讀文檔）")
    if "raw" not in st.session_state:
        st.session_state["raw"] = ""
    raw = st.text_area("貼上客戶填好的單據（tab / 逗號 / 雙空格分隔皆可）",
                       height=220, value=st.session_state["raw"], key="raw",
                       placeholder="貼上客戶填好的內容……（格式例子喺上面 ② 度睇）")

    col_parse, col_llm = st.columns([1, 1])
    with col_parse:
        run = st.button("🚀 自動解析填格", type="primary", use_container_width=True)
    with col_llm:
        use_llm = st.checkbox("客戶亂填時用 LLM 兜底（需接 OpenRouter）", value=False)

    if run and raw.strip():
        res = parse_text_with_llm(raw, llm_fn=None) if use_llm else parse_text(raw)
        # 新一次解析 → 清走舊確認/舊編輯，避免殘留
        for k in ("h_date", "h_cust", "confirm_write", "items_editor"):
            st.session_state.pop(k, None)
        st.session_state["parsed"] = res
    elif run:
        st.warning("請先貼上客戶填好的文字，或先「讀取文檔」。")

    # 結果區：存咗入 session 就一直顯示，等同事可以慢慢改表格 / tick 確認（唔會一撳嘢就消失）
    res = st.session_state.get("parsed")
    if res:
        h = dict(res["header"])

        st.markdown("### ⑤ 自動填好的單頭")
        d1, d2 = st.columns(2)
        d1v = d1.text_input("日期", value=h.get("date", ""), key="h_date")
        d2v = d2.text_input("公司（客戶自己公司名）", value=h.get("customer", ""), key="h_cust")
        if d1v:
            h["date"] = d1v
        if d2v:
            h["customer"] = d2v

        st.markdown("### ⑥ 自動填好的貨品明細（每行一格，可即場改）")
        if res["items"]:
            # 進出：保留客戶原字（唔自動轉），下拉選項包含原字 + 文檔標準值
            _orig_inout = list(dict.fromkeys(
                it.get("inout", "") for it in res["items"] if it.get("inout")))
            _inout_opts = INOUT_OPTIONS + [v for v in _orig_inout if v not in INOUT_OPTIONS]
            df = pd.DataFrame(res["items"])
            for f, _ in ITEM_COLS:
                if f not in df.columns:
                    df[f] = ""
            df = df[[f for f, _ in ITEM_COLS]]
            df.columns = [label for _, label in ITEM_COLS]
            df = df.fillna("")  # 空格顯示空白，唔好出 None

            # 進出：格內照顯示客戶原字（點一下格就有下拉）；上面再俾個永遠見到既下拉快速套用
            _QUICK_NONE = "（唔改——保留客戶原字）"

            def _quick_inout_changed():
                v = st.session_state.get("quick_inout_pick")
                if v and v != _QUICK_NONE:
                    st.session_state["force_inout_once"] = v
                    st.session_state["quick_inout_pick"] = _QUICK_NONE  # 用完彈返返預設，等下次再揀

            st.selectbox("⚡ 快速改「進出」——揀一個即套用落全部行（格內照顯示客戶原字，要改單行就點嗰格揀）",
                         [_QUICK_NONE] + INOUT_OPTIONS, key="quick_inout_pick", on_change=_quick_inout_changed)
            _fv = st.session_state.pop("force_inout_once", None)
            if _fv:
                df["進出"] = _fv
            edited = st.data_editor(
                df, use_container_width=True, num_rows="dynamic", key="items_editor", hide_index=True,
                column_config={
                    "流水号": st.column_config.NumberColumn("流水号", width="small", format="%d"),
                    "進出": st.column_config.SelectboxColumn(
                        "進出", options=_inout_opts, width="small",
                        help="顯示客戶原字；同事可下拉改做文檔標準值"),
                    "供應商(進)/出貨客戶": st.column_config.TextColumn("供應商(進)/出貨客戶", width="medium"),
                    "型号": st.column_config.TextColumn("型号", width="large"),
                    "數量(出貨-)": st.column_config.NumberColumn("數量(出貨-)", width="small", format="%d"),
                    "備註": st.column_config.TextColumn("備註", width="large"),
                    "入倉号/AWB": st.column_config.TextColumn("入倉号/AWB", width="medium"),
                },
            )

            # 同事可能喺表格改咗嘢 → 用改後嘅版本做落文檔 / 匯出（空格轉 ""，唔會出 nan）
            lbl2f = {lbl: f for f, lbl in ITEM_COLS}
            items_for_write = [
                {f: (row[lbl] if pd.notna(row[lbl]) else "") for lbl, f in lbl2f.items()}
                for _, row in edited.iterrows()
            ]

            # 複製結果貼到客戶專用文檔
            tab_txt = "\t".join(lbl for _, lbl in ITEM_COLS) + "\n"
            tab_txt += "\n".join("\t".join(str(rec.get(f, "")) for f, _ in ITEM_COLS) for rec in items_for_write)
            st.text_area("📋 複製下面貼到客戶專用文檔", tab_txt, height=140)

            with st.expander("匯出給主系統（JSON）"):
                payload = {"header": h, "items": items_for_write}
                st.json(payload)
                st.code(json.dumps(payload, ensure_ascii=False, indent=2), language="json")

            # ⑦ 自動寫回客户專用文檔（append 到最尾，唔覆蓋舊資料）
            st.markdown("### ⑦ 自動寫回客户專用文檔")
            st.caption(f"目標：{client_doc_url(cfid, ctid)}　（加係最尾，唔會覆蓋舊資料）")
            prev_rows = build_doc_rows(h, items_for_write, ccols)
            hdr = _target_header_cached(cfid, ctid) if ctid else []
            if hdr and len(hdr) == len(prev_rows[0]):
                prev_cols = [x if x.strip() else f"欄{i + 1}" for i, x in enumerate(hdr)]
            else:
                prev_cols = [f"欄{i + 1}" for i in range(len(prev_rows[0]))]
            st.dataframe(_style_prev_df(pd.DataFrame(prev_rows, columns=prev_cols)),
                         use_container_width=True, hide_index=True)
            if not hdr:
                st.caption("（攞唔到目標表頭，上面欄名暫以 欄1/欄2… 顯示）")
            confirm = st.checkbox("✅ 確認上面既行寫落客户文檔（加喺最尾）", key="confirm_write")
            if st.button("📤 寫回文檔", type="primary", disabled=not confirm, use_container_width=True):
                try:
                    start = write_items_to_doc(cfid, ctid, h, items_for_write, ccols)
                    st.success(f"✅ 已寫落 {len(items_for_write)} 行（由第 {start + 1} 行开始）。去文檔核對下。")
                    st.session_state.pop("confirm_write", None)
                except Exception as e:
                    st.error(str(e))
        else:
            st.warning("冇解析到貨品，請檢查格式或開 LLM 兜底。")

        if res["warnings"]:
            with st.expander("⚠️ 解析提示 / 異常", expanded=bool(res["warnings"])):
                for w in res["warnings"]:
                    st.write("•", w)


if __name__ == "__main__":
    main()
