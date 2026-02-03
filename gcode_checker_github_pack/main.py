import os
import re
import bisect
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from typing import Any, Dict, List, Optional, Tuple

# ============================
# （可選）拖拉開檔支援：tkinterdnd2
# 若你有安裝 tkinterdnd2，會自動啟用拖拉到視窗開檔
# ============================
DND_AVAILABLE = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False
    DND_FILES = None
    TkinterDnD = None

# ============================
# 狀態（中文）
# ============================
STATUS_OK = "通過"
STATUS_MISSING = "缺少"
STATUS_ERROR = "錯誤"

# ============================
# 路徑刪除參數（程式內部用，不顯示在介面）
# ============================
TRIM_START_Z_MAX = 0.1
TRIM_END_Z_MIN = 10.0
EPS = 1e-12


# ============================
# 讀檔（支援常見編碼 + 無副檔名）
# ============================
def read_text_with_fallback(path: str) -> Tuple[str, str]:
    encodings = ["utf-8-sig", "utf-8", "cp950", "big5", "gb18030"]
    last_err = None
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read(), enc
        except Exception as e:
            last_err = e
    try:
        with open(path, "rb") as f:
            raw = f.read()
        text = raw.decode("utf-8", errors="replace")
        return text, "utf-8（錯誤字元已替換）"
    except Exception as e:
        raise RuntimeError(f"讀檔失敗：{path}") from (last_err or e)


# ============================
# 行/列索引
# ============================
def build_line_starts(text: str) -> List[int]:
    starts = [0]
    for i, ch in enumerate(text):
        if ch == "\n":
            starts.append(i + 1)
    return starts


def index_to_line_col(idx: int, line_starts: List[int]) -> Tuple[int, int]:
    line = bisect.bisect_right(line_starts, idx) - 1
    return line + 1, idx - line_starts[line] + 1


# ============================
# 註解遮罩（保留長度與索引）
# 支援：(...)、以及 ; 到行尾
# ============================
def mask_gcode_comments_keep_length(text: str) -> str:
    chars = list(text)
    in_paren = 0
    i = 0
    n = len(chars)

    while i < n:
        ch = chars[i]

        if ch == "(":
            in_paren += 1
            chars[i] = " "
            i += 1
            continue

        if ch == ")" and in_paren > 0:
            chars[i] = " "
            in_paren -= 1
            i += 1
            continue

        if in_paren > 0:
            if ch != "\n":
                chars[i] = " "
            i += 1
            continue

        if ch == ";":
            while i < n and chars[i] != "\n":
                chars[i] = " "
                i += 1
            continue

        i += 1

    return "".join(chars)


# ============================
# Token helpers（允許連寫：G90Z.1、T01M06）
# - 避免 G54 誤判到 G543：後面不是數字才算 token 結束
# ============================
def token_pattern(token: str) -> str:
    return rf"(^|[^0-9A-Z])({re.escape(token)})(?=[^0-9]|$)"


def line_has_token(line: str, token: str) -> bool:
    return re.search(token_pattern(token), line, flags=re.IGNORECASE) is not None


def find_gcode_token(masked_text: str, token: str) -> Optional[re.Match]:
    return re.search(token_pattern(token), masked_text, flags=re.IGNORECASE | re.MULTILINE)


def extract_z_value(line: str) -> Optional[float]:
    m = re.search(
        r"(^|[^0-9A-Z])Z([+-]?(?:\d+\.\d*|\d*\.\d+|\d+))(?=[^0-9]|$)",
        line,
        flags=re.IGNORECASE
    )
    if not m:
        return None
    try:
        return float(m.group(2))
    except ValueError:
        return None


def collect_addr_numbers(masked_text: str, addr_letter: str) -> Tuple[List[str], Optional[int]]:
    addr_letter = addr_letter.upper()
    pattern = rf"(^|[^0-9A-Z])({addr_letter})([0-9]+)(?=[^0-9]|$)"

    first_idx: Optional[int] = None
    seen: Dict[int, str] = {}

    for m in re.finditer(pattern, masked_text, flags=re.IGNORECASE | re.MULTILINE):
        raw_num = m.group(3)
        try:
            num_int = int(raw_num)
        except ValueError:
            continue

        if num_int not in seen:
            seen[num_int] = raw_num

        idx = m.start(2)
        if first_idx is None or idx < first_idx:
            first_idx = idx

    nums_sorted = [seen[k] for k in sorted(seen.keys())]
    return nums_sorted, first_idx


# ============================
# 原始內容節錄（行範圍）
# ============================
def clamp(n: int, lo: int, hi: int) -> int:
    return max(lo, min(hi, n))


def make_range_excerpt(lines: List[str], center_line: int, radius: int = 2) -> Tuple[int, int, str, str]:
    """
    回傳：
      start_line, end_line, range_text, excerpt_text(含行號，多行可捲動看)
    """
    if not lines or center_line <= 0:
        return 0, 0, "", ""

    total = len(lines)
    s = clamp(center_line - radius, 1, total)
    e = clamp(center_line + radius, 1, total)

    excerpt_lines = []
    for ln in range(s, e + 1):
        excerpt_lines.append(f"{ln:6d}: {lines[ln - 1]}")
    excerpt = "\n".join(excerpt_lines)

    range_text = f"第{s}-{e}行" if s != e else f"第{s}行"
    return s, e, range_text, excerpt


# ============================
# 檢查項目（都回傳 pos_idx / line / col，方便抓原始內容範圍）
# ============================
def mk_result(
    rid: str,
    label: str,
    typ: str,
    status: str,
    evidence: str = "",
    pos_idx: Optional[int] = None,
    line: Optional[int] = None,
    col: Optional[int] = None,
) -> Dict[str, Any]:
    return {
        "id": rid,
        "label": label,
        "type": typ,
        "status": status,
        "evidence": evidence,
        "pos_idx": pos_idx,
        "line": line,
        "col": col,
    }


def eval_declaration_line(original_text: str, masked_text: str, line_starts: List[int]) -> Dict[str, Any]:
    required = ["G17", "G40", "G49", "G80", "G90"]
    missing = [t for t in required if find_gcode_token(masked_text, t) is None]
    if missing:
        return mk_result("DECL", "宣告碼檢查（G17 G40 G49 G80 G90）", "宣告碼", STATUS_MISSING, "缺少：" + ", ".join(missing))

    masked_lines = masked_text.splitlines(keepends=True)
    original_lines = original_text.splitlines(keepends=True)

    offset = 0
    for ml, ol in zip(masked_lines, original_lines):
        if all(line_has_token(ml, t) for t in required):
            line_no, col_no = index_to_line_col(offset, line_starts)
            evidence = f"同一行宣告：第{line_no}行 | {ol.strip()}"
            return mk_result("DECL", "宣告碼檢查（G17 G40 G49 G80 G90）", "宣告碼", STATUS_OK, evidence, pos_idx=offset, line=line_no, col=col_no)
        offset += len(ml)

    # 分散多行：用第一個命中位置當 pos
    first_pos = None
    first_token = None
    for t in required:
        m = find_gcode_token(masked_text, t)
        if m:
            idx = m.start(2)
            if first_pos is None or idx < first_pos:
                first_pos = idx
                first_token = t

    if first_pos is not None:
        line_no, col_no = index_to_line_col(first_pos, line_starts)
        evidence = "分散於多行（仍通過）。首次出現：" + f"{first_token}@第{line_no}行"
        return mk_result("DECL", "宣告碼檢查（G17 G40 G49 G80 G90）", "宣告碼", STATUS_OK, evidence, pos_idx=first_pos, line=line_no, col=col_no)

    return mk_result("DECL", "宣告碼檢查（G17 G40 G49 G80 G90）", "宣告碼", STATUS_OK, "已找到所有宣告碼（未能定位）")


def eval_home_return_g91_g28_z0(original_text: str, masked_text: str, line_starts: List[int], lookahead_lines: int = 3) -> Dict[str, Any]:
    orig_lines = original_text.splitlines(keepends=True)
    mask_lines = masked_text.splitlines(keepends=True)

    # 同行命中
    offset = 0
    for ol, ml in zip(orig_lines, mask_lines):
        z = extract_z_value(ml)
        if line_has_token(ml, "G91") and line_has_token(ml, "G28") and (z is not None) and abs(z - 0.0) <= 1e-6:
            line_no, col_no = index_to_line_col(offset, line_starts)
            evidence = f"同行命中：第{line_no}行 | {ol.strip()}"
            return mk_result("HOME", "回零檢查（G91 G28 Z0.0）", "安全回零", STATUS_OK, evidence, pos_idx=offset, line=line_no, col=col_no)
        offset += len(ml)

    # 跨行：以 G91/G28 那行當定位
    for i, ml in enumerate(mask_lines):
        if line_has_token(ml, "G91") and line_has_token(ml, "G28"):
            for j in range(i, min(i + 1 + lookahead_lines, len(mask_lines))):
                z2 = extract_z_value(mask_lines[j])
                if (z2 is not None) and abs(z2 - 0.0) <= 1e-6:
                    off_i = sum(len(x) for x in mask_lines[:i])
                    li, ci = index_to_line_col(off_i, line_starts)
                    evidence = f"跨行命中：第{li}行（G91/G28） + 第{j+1}行（Z0.0）"
                    return mk_result("HOME", "回零檢查（G91 G28 Z0.0）", "安全回零", STATUS_OK, evidence, pos_idx=off_i, line=li, col=ci)

    return mk_result("HOME", "回零檢查（G91 G28 Z0.0）", "安全回零", STATUS_MISSING, "未找到同一行或相鄰行組合（G91 + G28 + Z0.0）")


def eval_spindle_m03(original_text: str, masked_text: str, line_starts: List[int]) -> Dict[str, Any]:
    m = find_gcode_token(masked_text, "M03")
    token_used = "M03"
    if not m:
        m = find_gcode_token(masked_text, "M3")
        token_used = "M3"
    if not m:
        return mk_result("M03", "主軸正轉檢查（M03）", "主軸", STATUS_MISSING, "未找到 M03（或 M3）")

    idx = m.start(2)
    line_no, col_no = index_to_line_col(idx, line_starts)
    evidence = f"命中：{token_used}｜第{line_no}行，第{col_no}列"
    return mk_result("M03", "主軸正轉檢查（M03）", "主軸", STATUS_OK, evidence, pos_idx=idx, line=line_no, col=col_no)


def eval_gcode_quick(original_text: str, masked_text: str, line_starts: List[int], tokens: List[str], need_t: bool, need_h: bool) -> List[Dict[str, Any]]:
    results: List[Dict[str, Any]] = []

    for token in tokens:
        m = find_gcode_token(masked_text, token)
        if not m:
            results.append(mk_result(token, token, "G碼", STATUS_MISSING))
        else:
            idx = m.start(2)
            line_no, col_no = index_to_line_col(idx, line_starts)
            evidence = f"第{line_no}行，第{col_no}列"
            results.append(mk_result(token, token, "G碼", STATUS_OK, evidence, pos_idx=idx, line=line_no, col=col_no))

    if need_t:
        nums, first_idx = collect_addr_numbers(masked_text, "T")
        if not nums:
            results.append(mk_result("T", "T 號", "地址碼", STATUS_MISSING))
        else:
            ev = f"找到：{', '.join('T'+x for x in nums)}（共{len(nums)}個）"
            if first_idx is not None:
                line_no, col_no = index_to_line_col(first_idx, line_starts)
                ev += f"｜首次：第{line_no}行，第{col_no}列"
                results.append(mk_result("T", "T 號", "地址碼", STATUS_OK, ev, pos_idx=first_idx, line=line_no, col=col_no))
            else:
                results.append(mk_result("T", "T 號", "地址碼", STATUS_OK, ev))

    if need_h:
        nums, first_idx = collect_addr_numbers(masked_text, "H")
        if not nums:
            results.append(mk_result("H", "H 號", "地址碼", STATUS_MISSING))
        else:
            ev = f"找到：{', '.join('H'+x for x in nums)}（共{len(nums)}個）"
            if first_idx is not None:
                line_no, col_no = index_to_line_col(first_idx, line_starts)
                ev += f"｜首次：第{line_no}行，第{col_no}列"
                results.append(mk_result("H", "H 號", "地址碼", STATUS_OK, ev, pos_idx=first_idx, line=line_no, col=col_no))
            else:
                results.append(mk_result("H", "H 號", "地址碼", STATUS_OK, ev))

    return results


# ============================
# 移除 TOOL LIST 區塊（另存修正版用）
# ============================
TOOL_LIST_TITLE_RE = re.compile(r"\(\s*T\s*O\s*O\s*L\s*L\s*I\s*S\s*T\s*\)", re.IGNORECASE)

def _is_comment_line(s: str) -> bool:
    st = s.strip()
    return st.startswith("(") and st.endswith(")")

def _is_blank(s: str) -> bool:
    return s.strip() == ""

def _should_remove_following_spindle_line(line: str) -> bool:
    ml = mask_gcode_comments_keep_length(line)
    if not line_has_token(ml, "G00"):
        return False
    has_spindle = (re.search(r"(^|[^0-9A-Z])S[0-9]+(?=[^0-9]|$)", ml, re.IGNORECASE) is not None) or line_has_token(ml, "M03") or line_has_token(ml, "M3")
    has_xy = (re.search(r"(^|[^0-9A-Z])X[+-]?(?:\d+\.\d*|\d*\.\d+|\d+)(?=[^0-9]|$)", ml, re.IGNORECASE) is not None) or \
             (re.search(r"(^|[^0-9A-Z])Y[+-]?(?:\d+\.\d*|\d*\.\d+|\d+)(?=[^0-9]|$)", ml, re.IGNORECASE) is not None)
    return has_spindle and has_xy

def remove_tool_list_blocks(text: str) -> Tuple[str, int, int]:
    lines = text.splitlines(keepends=True)
    n = len(lines)
    remove = [False] * n
    blocks = 0
    removed_lines = 0

    i = 0
    while i < n:
        if TOOL_LIST_TITLE_RE.search(lines[i]):
            start = i
            j = i - 1
            while j >= 0 and (_is_comment_line(lines[j]) or _is_blank(lines[j])):
                start = j
                j -= 1

            end = i + 1
            k = i + 1
            while k < n and (_is_comment_line(lines[k]) or _is_blank(lines[k])):
                end = k + 1
                k += 1

            for idx in range(start, end):
                if not remove[idx]:
                    remove[idx] = True
                    removed_lines += 1

            blocks += 1

            if end < n and _should_remove_following_spindle_line(lines[end]):
                if not remove[end]:
                    remove[end] = True
                    removed_lines += 1
                end += 1

            i = end
            continue

        i += 1

    out = [lines[idx] for idx in range(n) if not remove[idx]]
    return "".join(out), blocks, removed_lines


# ============================
# 路徑刪除（另存修正版用）
# ============================
IMPORTANT_KEEP_RE = re.compile(
    r"(^|[^0-9A-Z])("
    r"G43|G54|G55|G56|G57|G58|G59|"
    r"T[0-9]+|H[0-9]+|"
    r"M0?6"
    r")(?=[^0-9]|$)",
    re.IGNORECASE
)

def trim_toolpath_by_markers(original_text: str, masked_text: str) -> Tuple[str, str]:
    orig_lines = original_text.splitlines(keepends=True)
    mask_lines = masked_text.splitlines(keepends=True)

    pos_mode = None
    motion_mode = None

    in_trim = False
    saw_below_safe = False
    start_line = None
    end_line = None
    removed_lines = 0

    out: List[str] = []

    for i, (ol, ml) in enumerate(zip(orig_lines, mask_lines)):
        if line_has_token(ml, "G90"):
            pos_mode = "G90"
        elif line_has_token(ml, "G91"):
            pos_mode = "G91"

        if line_has_token(ml, "G00"):
            motion_mode = "G00"
        elif line_has_token(ml, "G01"):
            motion_mode = "G01"
        elif line_has_token(ml, "G02"):
            motion_mode = "G02"
        elif line_has_token(ml, "G03"):
            motion_mode = "G03"

        z = extract_z_value(ml)

        if (not in_trim) and (pos_mode == "G90") and (z is not None) and (z <= TRIM_START_Z_MAX + EPS):
            if not (line_has_token(ml, "G28") or line_has_token(ml, "G53")):
                in_trim = True
                start_line = i + 1
                removed_lines += 1
                continue

        if (not in_trim) and (start_line is None) and (pos_mode == "G90") and (z is not None) and (z < TRIM_END_Z_MIN - EPS):
            if not (line_has_token(ml, "G28") or line_has_token(ml, "G53")):
                in_trim = True
                start_line = i + 1
                saw_below_safe = True
                removed_lines += 1
                continue

        if in_trim:
            if (z is not None) and (z < TRIM_END_Z_MIN - EPS):
                saw_below_safe = True

            if saw_below_safe and (motion_mode == "G00") and (z is not None) and (z >= TRIM_END_Z_MIN - EPS):
                in_trim = False
                end_line = i + 1
                out.append(ol)
                continue

            if IMPORTANT_KEEP_RE.search(ml):
                out.append(ol)
                continue

            removed_lines += 1
            continue

        out.append(ol)

    new_text = "".join(out)

    if start_line is None:
        return original_text, "未找到刪除起點"
    if end_line is None:
        return new_text, f"已從第{start_line}行開始刪除，但未找到安全回升終點。刪除行數={removed_lines}"
    return new_text, f"刪除完成：第{start_line}行起刪除，至第{end_line}行停止（保留該行）。刪除行數={removed_lines}"


# ============================
# GUI App（拖拉：若有 tkinterdnd2 則啟用）
# ============================
BaseTk = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk

class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("程式內容檢查工具（G碼 / T / H）")
        self.geometry("1200x950")
        self.minsize(1100, 860)

        self.file_path_var = tk.StringVar()

        # Quick 勾選
        self.chk_g54 = tk.BooleanVar(value=True)
        self.chk_g55 = tk.BooleanVar(value=False)
        self.chk_g56 = tk.BooleanVar(value=False)
        self.chk_g57 = tk.BooleanVar(value=False)
        self.chk_g58 = tk.BooleanVar(value=False)
        self.chk_g59 = tk.BooleanVar(value=False)
        self.chk_g43 = tk.BooleanVar(value=True)
        self.chk_tnum = tk.BooleanVar(value=True)
        self.chk_hnum = tk.BooleanVar(value=True)

        # 路徑刪除開關
        self.chk_trim_path = tk.BooleanVar(value=False)

        # 上方摘要
        self.t_summary_var = tk.StringVar(value="T: -")
        self.h_summary_var = tk.StringVar(value="H: -")
        self.file_info_var = tk.StringVar(value="尚未載入檔案")
        self.status_var = tk.StringVar(value="就緒")

        # 暫存
        self.last_loaded_text: Optional[str] = None
        self.last_encoding: Optional[str] = None
        self.last_lines: List[str] = []
        self.last_results: List[Dict[str, Any]] = []

        self._apply_style()
        self._build_ui()

        # 預設結果區更大
        self.after(100, self._set_default_sash)

        # 拖拉註冊（若可用）
        if DND_AVAILABLE:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self.on_drop_file)
            except Exception:
                pass

    def _apply_style(self):
        style = ttk.Style(self)
        for theme in ("clam", "vista", "xpnative"):
            try:
                style.theme_use(theme)
                break
            except tk.TclError:
                continue

        base_font = ("Microsoft JhengHei UI", 9)
        bold_font = ("Microsoft JhengHei UI", 9, "bold")
        self.option_add("*Font", base_font)

        style.configure("TButton", padding=(10, 5))
        style.configure("TCheckbutton", padding=(4, 1))
        style.configure("TLabelframe.Label", font=bold_font)

        # 結果表格行高更高
        style.configure("Treeview", rowheight=42)
        style.configure("Treeview.Heading", font=bold_font)

        style.configure("Pill.TLabel", padding=(10, 4), relief="solid")

    def _build_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(self, padding=(12, 10))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        ttk.Label(header, text="程式內容檢查工具（G碼 / T / H）", font=("Microsoft JhengHei UI", 12, "bold")).grid(row=0, column=0, sticky="w")

        tool = ttk.Frame(header)
        tool.grid(row=0, column=2, sticky="e")
        ttk.Button(tool, text="開始檢查", command=self.run_check).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(tool, text="另存修正版", command=self.save_processed_program).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(tool, text="清除結果", command=self.clear_results).grid(row=0, column=2)

        main = ttk.Frame(self, padding=(12, 0, 12, 12))
        main.grid(row=1, column=0, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        file_group = ttk.Labelframe(main, text="檔案", padding=10)
        file_group.grid(row=0, column=0, sticky="ew")
        file_group.grid_columnconfigure(1, weight=1)

        ttk.Label(file_group, text="程式檔案：").grid(row=0, column=0, sticky="w")
        self.file_entry = ttk.Entry(file_group, textvariable=self.file_path_var)
        self.file_entry.grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(file_group, text="瀏覽...", command=self.browse_text_file).grid(row=0, column=2, padx=(0, 6))
        ttk.Button(file_group, text="載入", command=self.load_text_file).grid(row=0, column=3)

        hint = "可直接在路徑欄貼上/輸入路徑後按「載入」"
        if DND_AVAILABLE:
            hint += "，或把檔案拖到視窗/路徑欄直接開啟"
        ttk.Label(file_group, text=hint).grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))

        ttk.Label(file_group, textvariable=self.file_info_var).grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))

        # 若可拖拉，也把 Entry 設成 drop target
        if DND_AVAILABLE:
            try:
                self.file_entry.drop_target_register(DND_FILES)
                self.file_entry.dnd_bind("<<Drop>>", self.on_drop_file)
            except Exception:
                pass

        self.paned = ttk.Panedwindow(main, orient="vertical")
        self.paned.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

        top_area = ttk.Frame(self.paned)
        bottom_area = ttk.Frame(self.paned)
        self.paned.add(top_area, weight=1)
        self.paned.add(bottom_area, weight=3)

        # ---------- top: quick options ----------
        top_area.grid_columnconfigure(0, weight=1)

        grp_offset = ttk.Labelframe(top_area, text="工件座標系（G54 ~ G59）", padding=10)
        grp_offset.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        for c in range(6):
            grp_offset.grid_columnconfigure(c, weight=1)

        ttk.Checkbutton(grp_offset, text="G54", variable=self.chk_g54).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(grp_offset, text="G55", variable=self.chk_g55).grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(grp_offset, text="G56", variable=self.chk_g56).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(grp_offset, text="G57", variable=self.chk_g57).grid(row=0, column=3, sticky="w")
        ttk.Checkbutton(grp_offset, text="G58", variable=self.chk_g58).grid(row=0, column=4, sticky="w")
        ttk.Checkbutton(grp_offset, text="G59", variable=self.chk_g59).grid(row=0, column=5, sticky="w")

        ttk.Button(grp_offset, text="全選", command=self.select_all_offsets).grid(row=1, column=4, sticky="e", pady=(8, 0))
        ttk.Button(grp_offset, text="全不選", command=self.clear_all_offsets).grid(row=1, column=5, sticky="w", pady=(8, 0))

        grp_tool = ttk.Labelframe(top_area, text="刀長補正 / 刀具資訊", padding=10)
        grp_tool.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        grp_tool.grid_columnconfigure(3, weight=1)

        ttk.Checkbutton(grp_tool, text="檢查 G43", variable=self.chk_g43).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(grp_tool, text="顯示 T 號", variable=self.chk_tnum).grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Checkbutton(grp_tool, text="顯示 H 號", variable=self.chk_hnum).grid(row=0, column=2, sticky="w", padx=(12, 0))
        ttk.Label(grp_tool, text="（確認刀長補正與工件座標）").grid(row=0, column=3, sticky="e")

        grp_trim = ttk.Labelframe(top_area, text="路徑刪除（另存修正版）", padding=10)
        grp_trim.grid(row=2, column=0, sticky="ew")
        grp_trim.grid_columnconfigure(1, weight=1)

        ttk.Checkbutton(grp_trim, text="啟用路徑刪除開關", variable=self.chk_trim_path).grid(row=0, column=0, sticky="w")
        ttk.Button(grp_trim, text="另存修正版程式...", command=self.save_processed_program).grid(row=0, column=1, sticky="e")

        # ---------- bottom: results + detail ----------
        bottom_area.grid_rowconfigure(3, weight=1)
        bottom_area.grid_columnconfigure(0, weight=1)

        res_group = ttk.Labelframe(bottom_area, text="檢查結果（可拖拉高度）", padding=10)
        res_group.grid(row=0, column=0, sticky="nsew")
        res_group.grid_rowconfigure(3, weight=1)
        res_group.grid_columnconfigure(0, weight=1)

        ttk.Label(res_group, textvariable=self.status_var).grid(row=0, column=0, sticky="w", pady=(0, 8))

        pill_bar = ttk.Frame(res_group)
        pill_bar.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        pill_bar.grid_columnconfigure(1, weight=1)

        ttk.Label(pill_bar, text="T 號：").grid(row=0, column=0, sticky="w")
        ttk.Label(pill_bar, textvariable=self.t_summary_var, style="Pill.TLabel").grid(row=0, column=1, sticky="w", padx=(6, 18))
        ttk.Label(pill_bar, text="H 號：").grid(row=0, column=2, sticky="w")
        ttk.Label(pill_bar, textvariable=self.h_summary_var, style="Pill.TLabel").grid(row=0, column=3, sticky="w", padx=(6, 0))

        # ✅ 新增一欄「範圍」，並新增一欄「原始內容（節錄）」
        columns = ("label", "type", "status", "range", "evidence", "raw")
        self.tree = ttk.Treeview(res_group, columns=columns, show="headings")
        self.tree.grid(row=2, column=0, sticky="nsew")

        self.tree.heading("label", text="項目 / 規則")
        self.tree.heading("type", text="類型")
        self.tree.heading("status", text="狀態")
        self.tree.heading("range", text="範圍")
        self.tree.heading("evidence", text="位置 / 證據")
        self.tree.heading("raw", text="原始內容（節錄）")

        self.tree.column("label", width=210, anchor="w")
        self.tree.column("type", width=90, anchor="w")
        self.tree.column("status", width=80, anchor="w")
        self.tree.column("range", width=110, anchor="w")
        self.tree.column("evidence", width=220, anchor="w")
        self.tree.column("raw", width=520, anchor="w")

        vsb = ttk.Scrollbar(res_group, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=2, column=1, sticky="ns", padx=(8, 0))

        hsb = ttk.Scrollbar(res_group, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=3, column=0, sticky="ew", pady=(6, 0))

        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("odd", background="#f6f7fb")
        self.tree.tag_configure("ok", foreground="#1b5e20")
        self.tree.tag_configure("miss", foreground="#b71c1c")
        self.tree.tag_configure("err", foreground="#e65100")

        detail_frame = ttk.Labelframe(bottom_area, text="詳細內容（可捲動看原始內容）", padding=10)
        detail_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        detail_frame.grid_rowconfigure(0, weight=1)
        detail_frame.grid_columnconfigure(0, weight=1)

        self.detail_box = ScrolledText(detail_frame, height=10, wrap="none")
        self.detail_box.grid(row=0, column=0, sticky="nsew")
        self._set_detail_text("點選上方結果列，可在這裡捲動查看『原始內容範圍』。")

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

    def _set_default_sash(self):
        try:
            total_h = self.paned.winfo_height()
            if total_h <= 0:
                return
            self.paned.sashpos(0, int(total_h * 0.33))
        except Exception:
            pass

    # ----------------------------
    # 拖拉開檔
    # ----------------------------
    def on_drop_file(self, event):
        try:
            # event.data 可能是多個檔案；取第一個
            paths = self.tk.splitlist(event.data)
            if not paths:
                return
            path = paths[0]
            if path and os.path.exists(path):
                self.file_path_var.set(path)
                self.load_text_file()
        except Exception:
            pass

    # ----------------------------
    # UI helpers
    # ----------------------------
    def select_all_offsets(self):
        self.chk_g54.set(True)
        self.chk_g55.set(True)
        self.chk_g56.set(True)
        self.chk_g57.set(True)
        self.chk_g58.set(True)
        self.chk_g59.set(True)

    def clear_all_offsets(self):
        self.chk_g54.set(False)
        self.chk_g55.set(False)
        self.chk_g56.set(False)
        self.chk_g57.set(False)
        self.chk_g58.set(False)
        self.chk_g59.set(False)

    def browse_text_file(self):
        path = filedialog.askopenfilename(
            title="選擇程式檔案",
            filetypes=[("所有檔案", "*")],
        )
        if path:
            self.file_path_var.set(path)

    def load_text_file(self):
        path = self.file_path_var.get().strip().strip('"')
        if not path:
            messagebox.showerror("錯誤", "請輸入或選擇檔案路徑。")
            return
        if not os.path.exists(path):
            messagebox.showerror("錯誤", "檔案不存在。")
            return
        try:
            text, enc = read_text_with_fallback(path)
            self.last_loaded_text = text
            self.last_encoding = enc
            self.last_lines = text.splitlines()

            line_count = len(self.last_lines)
            size_kb = os.path.getsize(path) / 1024.0
            self.file_info_var.set(
                f"已載入：{os.path.basename(path)}｜行數：{line_count}｜大小：{size_kb:.1f} KB｜編碼：{enc}"
            )
            self.status_var.set("已載入檔案，請按「開始檢查」。")
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def clear_results(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.last_results = []
        self.status_var.set("已清除結果。")
        self.t_summary_var.set("T: -")
        self.h_summary_var.set("H: -")
        self._set_detail_text("")

    def _set_detail_text(self, text: str):
        self.detail_box.configure(state="normal")
        self.detail_box.delete("1.0", "end")
        self.detail_box.insert("1.0", text or "")
        self.detail_box.configure(state="disabled")

    def on_tree_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        vals = self.tree.item(item_id, "values")
        # (label,type,status,range,evidence,raw)
        label = vals[0] if len(vals) > 0 else ""
        typ = vals[1] if len(vals) > 1 else ""
        status = vals[2] if len(vals) > 2 else ""
        rng = vals[3] if len(vals) > 3 else ""
        evidence = vals[4] if len(vals) > 4 else ""
        raw = vals[5] if len(vals) > 5 else ""

        # 從 last_results 找對應的 full excerpt
        full_excerpt = ""
        for r in self.last_results:
            if r.get("_row_key") == item_id:
                full_excerpt = r.get("full_excerpt", "") or ""
                break

        self._set_detail_text(
            f"項目：{label}\n"
            f"類型：{typ}\n"
            f"狀態：{status}\n"
            f"範圍：{rng}\n"
            f"位置/證據：{evidence}\n\n"
            f"原始內容（節錄）：\n{raw}\n\n"
            f"原始內容（範圍內完整）：\n{full_excerpt}"
        )

    def _get_selected_gcode_tokens(self) -> Tuple[List[str], bool, bool]:
        tokens: List[str] = []
        if self.chk_g54.get(): tokens.append("G54")
        if self.chk_g55.get(): tokens.append("G55")
        if self.chk_g56.get(): tokens.append("G56")
        if self.chk_g57.get(): tokens.append("G57")
        if self.chk_g58.get(): tokens.append("G58")
        if self.chk_g59.get(): tokens.append("G59")
        if self.chk_g43.get(): tokens.append("G43")
        return tokens, self.chk_tnum.get(), self.chk_hnum.get()

    # ----------------------------
    # 另存：路徑刪除修正版（+ 移除 TOOL LIST）
    # ----------------------------
    def save_processed_program(self):
        if self.last_loaded_text is None:
            self.load_text_file()
            if self.last_loaded_text is None:
                return

        if not self.chk_trim_path.get():
            messagebox.showinfo("提示", "請先勾選「啟用路徑刪除開關」再另存。")
            return

        text = self.last_loaded_text
        masked = mask_gcode_comments_keep_length(text)

        trimmed_text, trim_msg = trim_toolpath_by_markers(text, masked)
        cleaned_text, _, _ = remove_tool_list_blocks(trimmed_text)

        default_name = "processed.nc"
        in_path = self.file_path_var.get().strip()
        if in_path:
            base = os.path.basename(in_path)
            root, ext = os.path.splitext(base)
            if not ext:
                ext = ".nc"
            default_name = f"{root}_processed{ext}"

        save_path = filedialog.asksaveasfilename(
            title="另存修正版程式",
            initialfile=default_name,
            defaultextension=".nc",
            filetypes=[("所有檔案", "*")],
        )
        if not save_path:
            return

        try:
            with open(save_path, "w", encoding="utf-8", newline="") as f:
                f.write(cleaned_text)
            messagebox.showinfo("完成", f"已另存修正版：\n{save_path}\n\n{trim_msg}")
        except Exception as e:
            messagebox.showerror("錯誤", f"另存失敗：{e}")

    # ----------------------------
    # 核心：開始檢查
    # ----------------------------
    def run_check(self):
        if self.last_loaded_text is None:
            self.load_text_file()
            if self.last_loaded_text is None:
                return

        text = self.last_loaded_text
        masked_all = mask_gcode_comments_keep_length(text)
        line_starts = build_line_starts(text)

        decl = eval_declaration_line(text, masked_all, line_starts)
        home = eval_home_return_g91_g28_z0(text, masked_all, line_starts)
        m03 = eval_spindle_m03(text, masked_all, line_starts)

        tokens, need_t, need_h = self._get_selected_gcode_tokens()

        # 上方 T/H 摘要
        if need_t:
            t_nums, _ = collect_addr_numbers(masked_all, "T")
            self.t_summary_var.set("T: " + (", ".join("T" + n for n in t_nums) if t_nums else "(缺少)"))
        else:
            self.t_summary_var.set("T:（未勾選）")

        if need_h:
            h_nums, _ = collect_addr_numbers(masked_all, "H")
            self.h_summary_var.set("H: " + (", ".join("H" + n for n in h_nums) if h_nums else "(缺少)"))
        else:
            self.h_summary_var.set("H:（未勾選）")

        quick = eval_gcode_quick(text, masked_all, line_starts, tokens, need_t, need_h)
        results = [decl, home, m03] + quick

        # 結果寫入 Treeview（新增範圍、原始內容節錄）
        for i in self.tree.get_children():
            self.tree.delete(i)

        self.last_results = []
        missing = sum(1 for r in results if r["status"] == STATUS_MISSING)
        errors = sum(1 for r in results if r["status"] == STATUS_ERROR)
        self.status_var.set(f"完成：缺少={missing}｜錯誤={errors}｜總計={len(results)}")

        for idx, r in enumerate(results):
            tags = ["even" if idx % 2 == 0 else "odd"]
            if r["status"] == STATUS_OK:
                tags.append("ok")
            elif r["status"] == STATUS_MISSING:
                tags.append("miss")
            elif r["status"] == STATUS_ERROR:
                tags.append("err")

            line_no = r.get("line") or 0
            if line_no > 0 and self.last_lines:
                _, _, range_text, full_excerpt = make_range_excerpt(self.last_lines, line_no, radius=2)
                raw_line = self.last_lines[line_no - 1] if 1 <= line_no <= len(self.last_lines) else ""
            else:
                range_text, full_excerpt, raw_line = "", "", ""

            raw_short = raw_line.strip()
            if len(raw_short) > 140:
                raw_short = raw_short[:140] + "…"

            item = self.tree.insert(
                "",
                "end",
                values=(
                    r.get("label", ""),
                    r.get("type", ""),
                    r.get("status", ""),
                    range_text,
                    r.get("evidence", ""),
                    raw_short,
                ),
                tags=tags,
            )

            # 將完整 excerpt 存起來，點選時下方可捲動看
            r["_row_key"] = item
            r["full_excerpt"] = full_excerpt
            self.last_results.append(r)

    # ----------------------------
    # 結束
    # ----------------------------


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
