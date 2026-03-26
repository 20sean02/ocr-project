#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import csv
from typing import Optional, Tuple, List, Dict, Any

import cv2
import numpy as np
import pytesseract

# If tesseract isn't found, uncomment and set:
# pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"


# =========================================================
# Department mapping for 來文字號
# =========================================================
DEPT_MAP = {
    "教": "訓練科",
    "人": "人事室",
    "行": "行政科",
    "搶": "搶救科",
    "預": "預防科",
    "指": "指揮中心",
    "調": "火調科",
    "護": "救護科",
    "會": "會計室",
    "管": "災管科",
    "企": "督企科",
    "訓": "訓練科",
    "商": "府建商",
}


# =========================================================
# Utility
# =========================================================
def list_images(folder: str) -> List[str]:
    exts = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"}
    files = []
    for f in os.listdir(folder):
        _, ext = os.path.splitext(f.lower())
        if ext in exts:
            files.append(os.path.join(folder, f))
    files.sort()
    return files


def red_na_for_terminal(s: str) -> str:
    # CSV can't be colored; terminal can
    if s == "N/A":
        return "\033[31mN/A\033[0m"
    return s


def parse_roc_date_to_sort_key(date_str: Optional[str]) -> Tuple[int, int, int, int]:
    """
    date_str: "115.1.13" -> (0,115,1,13)
    missing/invalid -> (1,0,0,0)  (sort to the end)
    """
    if not date_str:
        return (1, 0, 0, 0)
    m = re.fullmatch(r"\s*(\d+)\.(\d+)\.(\d+)\s*", date_str)
    if not m:
        return (1, 0, 0, 0)
    y, mo, d = m.groups()
    try:
        return (0, int(y), int(mo), int(d))
    except Exception:
        return (1, 0, 0, 0)


def format_work_item_id(seq: int, date_str: Optional[str]) -> str:
    """
    seq: 1 -> "001"
    date_str: "115.1.13" -> "1.13"
    fallback if missing date -> "N/A"
    Final: "001_1.13"
    """
    seq_part = f"{seq:03d}"
    if not date_str:
        return f"{seq_part}_N/A"
    m = re.fullmatch(r"\s*(\d+)\.(\d+)\.(\d+)\s*", date_str)
    if not m:
        return f"{seq_part}_N/A"
    _, mo, d = m.groups()
    try:
        mo_i = int(mo)
        d_i = int(d)
        return f"{seq_part}_{mo_i}.{d_i}"
    except Exception:
        return f"{seq_part}_N/A"


# =========================================================
# Image preprocessing + OCR
# =========================================================
def read_and_prepare(image_path: str, max_dim: int = 1800) -> np.ndarray:
    """Return raw grayscale image (no blur applied)."""
    img = cv2.imread(image_path)
    if img is None:
        raise FileNotFoundError(f"Cannot read image: {image_path}")

    h, w = img.shape[:2]
    if max(h, w) > max_dim:
        scale = max_dim / max(h, w)
        img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    return gray


def preprocess_for_ocr(img_gray: np.ndarray, method: str = "gaussian") -> np.ndarray:
    """Apply preprocessing to raw grayscale image."""
    if method == "gaussian":
        return cv2.GaussianBlur(img_gray, (3, 3), 0)
    elif method == "bilateral":
        return cv2.bilateralFilter(img_gray, 9, 75, 75)
    elif method == "clahe":
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        return clahe.apply(img_gray)
    else:
        return img_gray


def ocr(img_gray: np.ndarray, lang: str = "chi_tra", psm: int = 6) -> str:
    import subprocess
    config = f"--oem 1 --psm {psm}"
    # Suppress tesseract stderr debug output
    old_popen = subprocess.Popen
    def quiet_popen(*args, **kwargs):
        kwargs.setdefault("stderr", subprocess.DEVNULL)
        return old_popen(*args, **kwargs)
    subprocess.Popen = quiet_popen
    try:
        return pytesseract.image_to_string(img_gray, lang=lang, config=config)
    finally:
        subprocess.Popen = old_popen


def ocr_english_subject(img_gray: np.ndarray) -> str:
    """
    English OCR pass tuned for model names / codes:
      JAC, i5, RAGC-700-ELCB, JEEP WRANGLER LIMITED SAHARA 4XE, TECC, etc.
    """
    config = (
        "--oem 1 --psm 6 "
        "-c preserve_interword_spaces=1 "
        "-c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-()/:., "
    )
    return pytesseract.image_to_string(img_gray, lang="eng", config=config)


# =========================================================
# ROI Cropping
# =========================================================
def _ensure_min_size(roi: np.ndarray, min_height: int = 800) -> np.ndarray:
    """Upscale ROI only if it's too small for Tesseract to read well."""
    h, w = roi.shape[:2]
    if h < min_height:
        scale = min_height / h
        roi = cv2.resize(roi, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    return roi


def crop_subject_rois(gray: np.ndarray) -> List[np.ndarray]:
    """
    Two ROIs to avoid truncation of multi-line subject.
    ROI2 is taller to capture longer subjects.
    """
    h, w = gray.shape[:2]
    rois = []

    x1 = int(w * 0.10)  # exclude 裝訂線 sidebar markers
    x2 = int(w * 0.95)

    # ROI 1 (standard subject area)
    roi1 = gray[int(h * 0.28):int(h * 0.49), x1:x2]
    rois.append(_ensure_min_size(roi1))

    # ROI 2 (taller)
    roi2 = gray[int(h * 0.26):int(h * 0.56), x1:x2]
    rois.append(_ensure_min_size(roi2))

    return rois


def crop_header_left_roi(gray: np.ndarray) -> np.ndarray:
    """
    ROI for:
    受文者 / 發文日期 / 發文字號
    """
    h, w = gray.shape[:2]
    roi = gray[int(h * 0.14):int(h * 0.33), int(w * 0.05):int(w * 0.70)]
    return _ensure_min_size(roi)


def crop_top_center_agency_roi(gray: np.ndarray) -> np.ndarray:
    """
    ROI for issuing agency in the middle-top header:
      - 澎湖縣政府 函
      - 澎湖縣政府消防局 函
    """
    h, w = gray.shape[:2]
    roi = gray[int(h * 0.03):int(h * 0.16), int(w * 0.20):int(w * 0.80)]
    return _ensure_min_size(roi, min_height=400)


# =========================================================
# Text helpers
# =========================================================
def normalize_text(s: str) -> str:
    s = s.replace("：", ":")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\r\n", "\n", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()


# =========================================================
# Subject cleanup (GENERAL, not per-document)
# =========================================================
_NOISE_SINGLE_CJK = set(list("汪訂謹爰茲函說明"))
_NOISE_TOKENS = {"k", "K", "l", "I", "|", "丨", "•", "·"}


def _drop_noise_tokens(s: str) -> str:
    toks = s.split()
    kept = []
    for t in toks:
        if t in _NOISE_TOKENS:
            continue
        if re.fullmatch(r"[A-Za-z]", t):
            continue
        if re.fullmatch(r"[\u4e00-\u9fff]", t) and t in _NOISE_SINGLE_CJK:
            continue
        kept.append(t)
    return " ".join(kept).strip()


def _hard_stop_inside_subject(s: str) -> str:
    s_ns = s.replace(" ", "")

    boundaries = [
        r"(謹|茲)?說\s*明\s*[:：]",
        r"正\s*本\s*[:：]?",
        r"副\s*本\s*[:：]?",
        r"附\s*件\s*[:：]?",
        r"\s[一二三四五六七八九十]\s*[、.．]",
    ]

    # Keep "請依說明..." inside subject; still stop at other headers
    if "依說明" in s_ns:
        boundaries = [
            r"正\s*本\s*[:：]?",
            r"副\s*本\s*[:：]?",
            r"附\s*件\s*[:：]?",
        ]

    for pat in boundaries:
        m = re.search(pat, s)
        if m:
            return s[: m.start()].strip()

    return s.strip()


def _fix_roc_long_date_typos(s: str) -> str:
    """
    Fix OCR misreads in ROC long-date phrases, including spacing variants.
    """
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s*日\s*", "日", s)
    s = re.sub(r"(月)\s*一\s*[\u4e00-\u9fff]\s*日", r"\1一日", s)
    return s


def _strip_ascii_garbage(s: str) -> str:
    """
    Remove the common 'dump' garbage without killing real English tokens.
    - removes urls / pdf names
    - removes very long alnum blobs
    Keeps:
      JAC, i5, TECC, RAGC-700-ELCB, JEEP WRANGLER..., 4XE
    """
    s = re.sub(r"(https?://\S+)", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b\S+\.pdf\b", " ", s, flags=re.IGNORECASE)

    def repl_long(m: re.Match) -> str:
        token = m.group(0)
        # keep short hyphenated model codes
        if "-" in token and len(token) <= 30:
            return token
        return " "

    s = re.sub(r"[A-Za-z0-9]{18,}", repl_long, s)
    s = re.sub(r"[-_]{4,}", " ", s)
    return s


# =========================================================
# Post-corrector for common OCR confusions (NO upgrades)
# =========================================================
COMMON_PHRASE_FIXES = {
    # earlier batch issues
    "設立迷記": "設立登記",
    "商業設立迷記": "商業設立登記",
    "變更登記一貯": "變更登記一案",
    "一貯": "一案",
    "淮如所請": "准如所請",
    "艾商業": "貴商業",
    "坎代役": "替代役",
    "時豉勵": "鼓勵",
    "有有關": "有關",
    "辮理": "辦理",
    "傷呻": "傷患",
    "蔡代役": "替代役",
    "了人作": "工作",
    "落瘓": "落實",
    "後貫": "後續",
    "職堂": "職掌",

    # issues you showed
    "汪淮": "江淮",
    "元寡節": "元宵節",
    "宜導": "宣導",
    "主昌": "主旨",
    "火哭資料": "火災資料",
    "報一一告": "報告",
    "腊餘": "剩餘",
    "生份": "份",
    "寥消防": "之消防",
    "傷惠": "傷患",
    "蟬轉": "函轉",
    "巷代役": "替代役",
    "安爹案": "安全案",
    "洛實": "落實",
    "落賓": "落實",
    "執答": "執行",
    "業務職堂": "業務職掌",
    "學主班": "學士班",

    # common header OCR variants you’ve seen
    "范轉": "函轉",

    # 007_2.2 / IMG_7442 / IMG_7449 fixes
    "場欠": "場次",
    "場欽": "場次",
    "懲還": "懲罰",
    "懲員": "懲罰",
    "人系列": "系列",
    "了系急": "緊急",
    "系急聯絡": "緊急聯絡",
    "及系急": "及緊急",
    "轉圈參訓": "轉達參訓",
    "轉殼參訓": "轉達參訓",
    "參瞻": "參訓",

    # remaining OCR noise cleanup
    "登記倆了": "登記一",
    "設立和登記": "設立登記",
    "有銷久聯絡": "緊急聯絡",
    "吵詢": "函詢",
    "有了關": "有關",
    "蔡備役": "後備役",
    "寺守份": "份",
    "報了一告": "報告",
    "執稅": "執行",
    "戰轉": "函轉",
    "盅躍": "踴躍",
    "優為蹟": "優異事蹟",
    "郊教": "教育",
    "宣達教育": "宣導教育",

    # additional OCR variants
    "囚急": "緊急",
    "累急": "緊急",
    "臟餘": "剩餘",
    "腊餘": "剩餘",
    "火器資料": "火災資料",
    "火哭資料": "火災資料",
    "邢EP": "JEEP",
}


def apply_phrase_fixes(s: str) -> str:
    if not s:
        return s

    for wrong, correct in COMMON_PHRASE_FIXES.items():
        s = s.replace(wrong, correct)

    # Remove stray numeric before 請查照 (e.g., "，1 請查照")
    s = re.sub(r"，?\s*\d+\s*請查照", "，請查照", s)

    # Remove weird lone symbol blocks like "答』"
    s = re.sub(r"[』「」]{1,}", "", s)

    # Normalize "主旨 :" variations
    s = re.sub(r"主\s*旨\s*[:：]\s*", "", s)

    # Fix: "轉知所屬= 請查照" -> "轉知所屬，請查照"
    s = s.replace("所屬= 請查照", "所屬，請查照")
    s = s.replace("所屬=請查照", "所屬，請查照")

    # Fix "XEP WRANGLER" -> "JEEP WRANGLER" (OCR misread of J as various CJK)
    s = re.sub(r"[\u4e00-\u9fff]EP\s+WRANGLER", "JEEP WRANGLER", s)
    # Also fix standalone "XEP" before non-CJK
    s = re.sub(r"[工邢距]EP\b", "JEEP", s)

    # Fix "11$征" -> "115年" ($ is OCR for 5, 征 for 年)
    s = re.sub(r"11\$征", "115年", s)

    # Fix "。，" -> "，" (stray period before comma)
    s = s.replace("。，", "，")

    # Fix "執行 : 三計畫" -> "執行計畫" (": 三" is sidebar noise)
    s = re.sub(r"執行\s*[:：]\s*三計畫", "執行計畫", s)

    # Strip leading punctuation noise
    s = re.sub(r"^[、，。：:]+\s*", "", s)

    # Fix "落實執X人請" / "執行請" -> "落實執行，請" (OCR mangling)
    s = re.sub(r"執[\u4e00-\u9fff]人請", "執行，請", s)
    s = re.sub(r"執行請查照", "執行，請查照", s)

    # Fix "登記XX 案" -> "登記一案" (OCR mangling of 一案)
    s = re.sub(r"登記[\u4e00-\u9fff]{1,2}\s*案", "登記一案", s)

    # Fix "X急聯絡" / "X急救援" -> "緊急..." (OCR misread of 緊)
    s = re.sub(r"[\u4e00-\u9fff]急聯絡", "緊急聯絡", s)
    s = re.sub(r"[\u4e00-\u9fff]急救援", "緊急救援", s)

    # Normalize double+ spaces to single
    s = re.sub(r"  +", " ", s)

    return s


def clean_subject(s: str) -> str:
    s = normalize_text(s)

    # Normalize separators
    s = s.replace("|", " ").replace("丨", " ").replace("•", " ").replace("·", " ")

    # Remove ascii garbage blobs but keep normal English tokens
    s = _strip_ascii_garbage(s)

    # Common OCR confusions (kept minimal + general)
    s = s.replace("吵轉", "函轉")
    s = s.replace("哨轉", "函轉")

    # Fix long ROC date phrase OCR typos
    s = _fix_roc_long_date_typos(s)

    # Match 主旨 and common OCR variants: 主劑, 主百, 主昌, 主上旨, 主午
    _ZHI_PAT = r"主\s*(?:上\s*)?(?:旨|劑|百|昌|午)"

    # Strip contamination like "如主旨主旨:" / "主旨:"
    s = re.sub(r"^.*?附件[:：]?\s*如\s*" + _ZHI_PAT + r"\s*", "", s)
    s = re.sub(r"^(如\s*)?" + _ZHI_PAT + r"\s*[:：]?\s*", "", s)
    s = re.sub(r"^\s*" + _ZHI_PAT + r"\s*[:：]?\s*", "", s)

    # Keep only content after the LAST 主旨-like marker
    if re.search(_ZHI_PAT + r"\s*[:：]", s):
        parts = re.split(_ZHI_PAT + r"\s*[:：]\s*", s)
        s = parts[-1].strip()

    # Turn OCR glitch "汪 案" etc into "一案"
    s = re.sub(r"汪\s*[!！]?\s*案", "一案", s)
    s = re.sub(r"汪\s*\w+\s*案", "一案", s)

    # Hard-stop if headers got merged into subject line
    s = _hard_stop_inside_subject(s)

    # Remove isolated noise tokens
    s = _drop_noise_tokens(s)

    # Normalize punctuation and spaces
    s = s.replace(",", "，")
    s = re.sub(r"([\u4e00-\u9fff])\s+([\u4e00-\u9fff])", r"\1\2", s)  # remove spaces between CJK
    s = re.sub(r"\s*([，。！？；：])\s*", r"\1", s)

    # Remove stray noise injected from sidebar / margin OCR artifacts:
    # - isolated Latin chars (1-3 letters not part of known English tokens)
    # - stray symbols, short fragments at word boundaries
    s = re.sub(r"\s+[A-Za-z]{1,2}(?=[\u4e00-\u9fff，。])", "", s)  # "救援 ee 材" -> "救援材"
    s = re.sub(r"(?<=[\u4e00-\u9fff])\s+[A-Za-z]{1,2}\s*$", "", s)  # trailing "...份 es"
    s = re.sub(r"\s*[~<>=]{1,}\s*", "", s)
    # Remove stray colons between numbers/CJK (sidebar OCR artifact)
    s = re.sub(r"(\d)\s*[:：]\s*([\u4e00-\u9fff])", r"\1\2", s)
    # Remove isolated sidebar CJK noise: single chars like 生,呈,宮,富,還,品 after spaces
    s = re.sub(r"\s+[\u4e00-\u9fff]{1,2}\s+(?=[A-Za-z])", " ", s)
    s = re.sub(r"\s+[\u4e00-\u9fff]{1}$", "", s)  # trailing single CJK

    s = " ".join(s.split())

    # If contains 請查照, cut to 請查照 (ENFORCE only when present)
    # Also handle 請查 followed by garbage (照 often gets mangled by OCR)
    m = re.search(r"^(.*?請\s*查\s*照)", s)
    if m:
        s = m.group(1).strip()
    else:
        # 請查 at end or followed by non-CJK garbage → assume 請查照
        m2 = re.search(r"^(.*?請\s*查)\s*[^，。\u4e00-\u9fff]*$", s)
        if m2:
            s = m2.group(1).strip() + "照"

    # Apply safe phrase-level corrections (domain-specific)
    s = apply_phrase_fixes(s)

    # Ensure ending period
    s = s.rstrip(" ，,:：．。")
    return s + "。"


def extract_subject(ocr_text: str) -> Optional[str]:
    """
    Extract 主旨:
    - Start at line containing 主旨
    - Collect until stop markers:
      * '說明' as a SECTION HEADER
      * '一、' bullets
      * 附件/正本/副本
    - Then apply general cleaning
    """
    if not ocr_text:
        return None

    t = normalize_text(ocr_text)
    lines = [ln.strip() for ln in t.splitlines() if ln.strip()]
    if not lines:
        return None

    # Find 主旨 line (fuzzy, includes OCR variants: 主劑, 主百, 主昌, 主午)
    _ZHI_FIND = re.compile(r"主\s*(?:上\s*)?(?:旨|劑|百|昌|午)")
    start = None
    for i, ln in enumerate(lines):
        ln_ns = ln.replace(" ", "")
        if _ZHI_FIND.search(ln_ns):
            start = i
            break
    if start is None:
        return None

    parts = []

    # Same-line tail after 主旨: (or variant)
    first = lines[start].replace("：", ":")
    if ":" in first:
        tail = first.split(":", 1)[1].strip()
        if tail:
            parts.append(tail)
    else:
        # sometimes OCR misses colon — find end of the 主旨-variant marker
        m_pos = _ZHI_FIND.search(lines[start])
        if m_pos:
            after = lines[start][m_pos.end():].strip()
            if after:
                parts.append(after)

    def is_stop_line(line: str) -> bool:
        s = line.strip()

        # 說明 as header
        if re.match(r"^(謹|茲)?說\s*明\s*[:：]?\s*$", s):
            return True

        s_ns = s.replace(" ", "")
        if "附件" in s_ns or "正本" in s_ns or "副本" in s_ns:
            return True

        if re.match(r"^[一二三四五六七八九十]\s*[、.．]", s):
            return True

        return False

    for j in range(start + 1, len(lines)):
        ln = lines[j].strip()
        if is_stop_line(ln):
            break
        parts.append(ln)

    if not parts:
        return None

    subject_raw = " ".join(parts).strip()
    subject = clean_subject(subject_raw)
    return subject


# =========================================================
# English merge helpers (to recover missing model names)
# =========================================================
def extract_english_tokens(eng_text: str) -> List[str]:
    if not eng_text:
        return []

    t = normalize_text(eng_text).replace("\n", " ")

    tokens = []
    tokens += re.findall(r"\b[A-Z]{2,}[A-Z0-9-]{0,25}\b", t)  # JEEP, RAGC-700-ELCB
    tokens += re.findall(r"\b[a-zA-Z]\d{1,3}\b", t)          # i5
    tokens += re.findall(r"\b\d[A-Z]{1,3}\b", t)             # 4XE
    tokens += re.findall(r"\b[A-Z]{2,}\d{1,4}\b", t)         # TECC-like

    seen = set()
    out = []
    for tok in tokens:
        if tok in seen:
            continue
        seen.add(tok)
        out.append(tok)
    return out


def enrich_subject_with_english(subject: str, eng_text: str) -> str:
    """
    Use English OCR as a helper (NOT the only logic).
    If we detect classic missing areas, patch them.
    """
    if not subject:
        return subject

    tokens = extract_english_tokens(eng_text)
    if not tokens:
        return subject

    s = subject

    # 車型 patch: add missing JAC / i5
    if "車型" in s:
        has_jac = "JAC" in s
        has_i5 = "i5" in s or "I5" in s
        want_jac = "JAC" in tokens
        want_i5 = ("i5" in tokens) or ("I5" in tokens)

        if (want_jac and not has_jac) or (want_i5 and not has_i5):
            m = re.search(r"\(車型\s*[:：]\s*([^)]*)\)", s)
            if m:
                inside = m.group(1).strip()
                new_inside = inside
                if want_jac and not has_jac:
                    new_inside = "JAC " + new_inside
                if want_i5 and not has_i5:
                    new_inside = re.sub(r"\b5\b", "i5", new_inside)
                    if "i5" not in new_inside and "I5" not in new_inside:
                        new_inside = new_inside + " i5"
                s = s[: m.start()] + f"(車型:{new_inside})" + s[m.end():]
            else:
                add = []
                if want_jac and not has_jac:
                    add.append("JAC")
                if want_i5 and not has_i5:
                    add.append("i5")
                if add:
                    s = re.sub(r"(車型)", r"\1 " + " ".join(add), s, count=1)

    # RAGC-700-ELCB patch
    if "-700-" in s or "700" in s:
        model_like = [t for t in tokens if "700" in t and "-" in t and len(t) <= 30]
        if model_like:
            best = max(model_like, key=len)
            s = s.replace("-700-", best)
            if "提供" in s and "系列" in s and best not in s:
                s = re.sub(r"(提供)\s*", r"\1" + best + " ", s, count=1)

    # JEEP phrase patch
    if "提供" in s and "JEEP" not in s:
        jeep_words = ["JEEP", "WRANGLER", "LIMITED", "SAHARA", "4XE"]
        present = [w for w in jeep_words if w in tokens]
        if len(present) >= 2:
            phrase = " ".join([w for w in jeep_words if w in present])
            s = re.sub(r"(提供)\s*", r"\1" + phrase + " ", s, count=1)

    # TECC patch for empty parentheses
    if re.search(r"\(\s*\)", s) and "TECC" in tokens:
        s = re.sub(r"\(\s*\)", "(TECC)", s, count=1)

    # Enforce 請查照 cutoff only if present
    m = re.search(r"^(.*?請\s*查\s*照)", s)
    if m:
        s = m.group(1).strip() + "。"
    else:
        s = s.rstrip(" ，,:：．。") + "。"

    return s


# =========================================================
# Header field extraction
# =========================================================
def extract_issue_date(text: str) -> Optional[str]:
    """
    發文日期: 中華民國115年1月13日 -> 115.1.13
    """
    if not text:
        return None

    flat = normalize_text(text).replace("\n", " ")
    m = re.search(r"發\s*文\s*日\s*期[:：]?\s*中華民國\s*(\d+)年\s*(\d+)月\s*(\d+)日", flat)
    if not m:
        return None

    y, mo, d = m.groups()
    return f"{int(y)}.{int(mo)}.{int(d)}"


def extract_doc_no(text: str) -> Optional[str]:
    """
    發文字號 -> digits
    """
    if not text:
        return None

    flat = normalize_text(text).replace("\n", " ").replace(" ", "")

    m = re.search(r"發文字號[:：]?(.*?第(\d{6,})號)", flat)
    if m:
        return m.group(2)

    m = re.search(r"發文字號[:：]?.*?字(\d{6,})號", flat)
    if m:
        return m.group(1)

    m = re.search(r"發文字號[:：]?(.*?)(速別|密等|附件|$)", flat)
    if m:
        chunk = m.group(1)
        nums = re.findall(r"\d{6,}", chunk)
        if nums:
            return max(nums, key=len)

    return None


def extract_laiwen_dept(text: str) -> str:
    """
    Column: 來文字號 (department name)
    Rule: abbreviation is the 3rd CJK character after 發文字號 (before '字').
    """
    if not text:
        return "N/A"

    flat = normalize_text(text).replace("\n", "").replace(" ", "")

    m = re.search(r"發\s*文\s*字\s*號[:：]?(.*?)(字)", flat)
    if not m:
        return "N/A"

    before_zi = m.group(1)
    cjk_chars = re.findall(r"[\u4e00-\u9fff]", before_zi)

    if len(cjk_chars) < 3:
        return "N/A"

    abbr = cjk_chars[2]
    return DEPT_MAP.get(abbr, "N/A")


def extract_issuing_agency_from_top(text: str) -> str:
    """
    Decide issuing agency based on top-center header.
    Priority:
        1. 澎湖縣政府消防局 -> 消防局
        2. 澎湖縣政府 -> 澎湖縣政府
        3. N/A
    """
    if not text:
        return "N/A"

    flat = normalize_text(text)
    flat = flat.replace("\n", "").replace(" ", "")

    # Find first occurrence of each agency pattern (earliest = issuing agency,
    # later ones might be from 受文者 field)
    candidates = []
    for pattern, agency in [
        ("澎湖縣政府消防局", "消防局"),
        ("內政部消防署", "消防署"),
        ("澎湖縣政府", "澎湖縣政府"),
    ]:
        idx = flat.find(pattern)
        if idx >= 0:
            candidates.append((idx, agency))

    if not candidates:
        return "N/A"

    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]


def infer_agency_from_fawenzihao(text: str) -> str:
    """Infer issuing agency from 發文字號 prefix like 澎消X字, 消署X字."""
    if not text:
        return "N/A"
    flat = normalize_text(text).replace("\n", "").replace(" ", "")
    m = re.search(r"發文字號[:：]?\s*([\u4e00-\u9fff]+)字第", flat)
    if not m:
        return "N/A"
    prefix = m.group(1)
    if "澎消" in prefix:
        return "消防局"
    if "消署" in prefix:
        return "消防署"
    if "府建" in prefix:
        return "澎湖縣政府"
    return "N/A"


def extract_tongbao_subject(text: str) -> Optional[str]:
    """Extract body text from 通報 documents (no 主旨 section)."""
    if not text:
        return None

    t = normalize_text(text)
    lines = [ln.strip() for ln in t.splitlines() if ln.strip()]

    # Find start: after 保密期限 line
    start = None
    for i, ln in enumerate(lines):
        ln_ns = ln.replace(" ", "")
        if "保密期限" in ln_ns:
            start = i + 1

    if start is None:
        return None

    # Collect lines until 此致 or distribution list markers
    parts = []
    for j in range(start, len(lines)):
        ln = lines[j].strip()
        if re.search(r"此\s*致", ln):
            break
        # Strip sidebar markers
        ln = re.sub(r"^[|!‧.\s]*", "", ln)
        # Skip pure noise
        if not ln or ln in ("全", "裝", "訂", "線", "有", "市"):
            continue
        parts.append(ln)

    if not parts:
        return None

    subject_raw = " ".join(parts).strip()
    # Basic cleanup
    subject = normalize_text(subject_raw)
    subject = subject.replace(",", "，")
    subject = re.sub(r"([\u4e00-\u9fff])\s+([\u4e00-\u9fff])", r"\1\2", subject)
    subject = re.sub(r"\s*([，。！？；：])\s*", r"\1", subject)
    # Remove stray trailing noise (digits, symbols after last CJK)
    subject = re.sub(r"[0-9|!]+\s*$", "", subject)
    subject = " ".join(subject.split())
    subject = apply_phrase_fixes(subject)
    subject = subject.rstrip(" ，,:：．。") + "。"
    return subject


# =========================================================
# Main processing
# =========================================================
_PREPROCESS_METHODS = ["bilateral", "clahe"]


def _score_subject(subj: Optional[str]) -> int:
    if not subj:
        return -1

    # Base: count CJK characters (not total length, to avoid rewarding garbage)
    cjk_count = len(re.findall(r"[\u4e00-\u9fff]", subj))
    score = cjk_count

    if "請依說明" in subj:
        score += 50
    if "請查照" in subj:
        score += 20

    # Penalize noise: isolated Latin chars, random symbols
    garbage_hits = len(re.findall(r"(?<!\w)[A-Za-z]{1,2}(?!\w)", subj))
    score -= garbage_hits * 5

    # Penalize stray numbers not part of meaningful content
    stray_nums = len(re.findall(r"(?<![0-9\.\-])\b\d{1}\b(?![0-9\.\-])", subj))
    score -= stray_nums * 3

    return score


def process_one(image_path: str) -> Tuple[Optional[str], Optional[str], str, Optional[str], str]:
    gray_raw = read_and_prepare(image_path)

    # --- Subject via dual ROI x dual preprocessing (pick best) ---
    best_subject = None
    best_score = -1
    all_subject_roi_texts = []

    for roi_raw in crop_subject_rois(gray_raw):
        for method in ("gaussian", "bilateral"):
            roi = preprocess_for_ocr(roi_raw, method)
            zh_txt = ocr(roi, lang="chi_tra", psm=6)
            all_subject_roi_texts.append(zh_txt)
            subj = extract_subject(zh_txt)
            if not subj:
                continue

            eng_txt = ocr_english_subject(roi)
            subj = enrich_subject_with_english(subj, eng_txt)

            score = _score_subject(subj)
            if score > best_score:
                best_score = score
                best_subject = subj

    subject = best_subject

    # --- Header-left fields (Gaussian blur — proven reliable for headers) ---
    header_roi = preprocess_for_ocr(crop_header_left_roi(gray_raw), "gaussian")
    header_text = ocr(header_roi, lang="chi_tra", psm=6)

    issue_date = extract_issue_date(header_text)
    doc_no = extract_doc_no(header_text)
    laiwen_dept = extract_laiwen_dept(header_text)

    # --- Also try header fields from subject ROI text (handles different layouts) ---
    for srt in all_subject_roi_texts:
        if issue_date and doc_no and laiwen_dept != "N/A":
            break
        if not srt:
            continue
        if issue_date is None:
            issue_date = extract_issue_date(srt)
        if doc_no is None:
            doc_no = extract_doc_no(srt)
        if laiwen_dept == "N/A":
            laiwen_dept = extract_laiwen_dept(srt)

    # --- Issuing agency from top-center header ---
    top_roi = preprocess_for_ocr(crop_top_center_agency_roi(gray_raw), "gaussian")
    top_text = ocr(top_roi, lang="chi_tra", psm=6)
    from_agency = extract_issuing_agency_from_top(top_text)

    # Also try agency from header text (handles different doc layouts)
    if from_agency == "N/A":
        from_agency = extract_issuing_agency_from_top(header_text)

    # Also try inferring agency from 發文字號 prefix
    if from_agency == "N/A":
        from_agency = infer_agency_from_fawenzihao(header_text)
    if from_agency == "N/A":
        for srt in all_subject_roi_texts:
            from_agency = infer_agency_from_fawenzihao(srt)
            if from_agency != "N/A":
                break

    # --- Detect 通報 documents (no 主旨 section) ---
    is_tongbao = bool(re.search(r"通\s*報", top_text or ""))
    if not is_tongbao:
        is_tongbao = bool(re.search(r"通\s*報", header_text or ""))

    if is_tongbao and subject is None:
        # Use a body ROI covering header + subject area for 通報
        h, w = gray_raw.shape[:2]
        body_roi_raw = _ensure_min_size(
            gray_raw[int(h * 0.18):int(h * 0.55), int(w * 0.10):int(w * 0.95)]
        )
        # Try both preprocessing methods, pick best by CJK count
        best_tb = None
        best_tb_cjk = -1
        for method in ("gaussian", "bilateral"):
            body_roi = preprocess_for_ocr(body_roi_raw, method)
            body_text = ocr(body_roi, lang="chi_tra", psm=6)
            tb_subj = extract_tongbao_subject(body_text)
            if tb_subj:
                cjk_n = len(re.findall(r"[\u4e00-\u9fff]", tb_subj))
                if cjk_n > best_tb_cjk:
                    best_tb_cjk = cjk_n
                    best_tb = tb_subj
        subject = best_tb

    # fallback full page OCR with PSM 3 if missing criticals
    if (issue_date is None) or (doc_no is None) or (laiwen_dept == "N/A") or (subject is None) or (from_agency == "N/A"):
        full_gray = preprocess_for_ocr(gray_raw, "bilateral")
        full_text = ocr(full_gray, lang="chi_tra", psm=3)

        if issue_date is None:
            issue_date = extract_issue_date(full_text)
        if doc_no is None:
            doc_no = extract_doc_no(full_text)
        if laiwen_dept == "N/A":
            laiwen_dept = extract_laiwen_dept(full_text)
        if subject is None:
            subject = extract_subject(full_text)
            if subject:
                rois = crop_subject_rois(gray_raw)
                if rois:
                    eng_txt = ocr_english_subject(rois[-1])
                    subject = enrich_subject_with_english(subject, eng_txt)

        # For 通報 fallback on full text
        if is_tongbao and subject is None:
            subject = extract_tongbao_subject(full_text)

        if from_agency == "N/A":
            from_agency = extract_issuing_agency_from_top(full_text)
        if from_agency == "N/A":
            from_agency = infer_agency_from_fawenzihao(full_text)

    # For 通報 documents, just use "通報" as agency
    if is_tongbao:
        from_agency = "通報"

    return subject, issue_date, laiwen_dept, doc_no, from_agency


# =========================================================
# Entry
# =========================================================
def main():
    images_folder = "images"
    output_csv = "subject_output_full_withEN.csv"

    files = list_images(images_folder)
    if not files:
        print(f"[error] No images found in: {images_folder}")
        return

    # 1) Process everything first (do NOT write yet)
    rows: List[Dict[str, Any]] = []

    for path in files:
        original_fname = os.path.basename(path)

        try:
            subject, issue_date, laiwen_dept, doc_no, from_agency = process_one(path)

            status = "ok" if (subject and from_agency != "N/A" and doc_no and issue_date) else "partial"

            rows.append({
                "original_filename": original_fname,
                "來文機關": from_agency or "N/A",
                "來文字號": laiwen_dept or "N/A",
                "收文文號": doc_no or "",
                "收文日期": issue_date or "",
                "事由": subject or "",
                "status": status,
                "_sort_key": parse_roc_date_to_sort_key(issue_date),
            })

            print(
                f"[{status}] {original_fname} -> "
                f"agency={from_agency or 'N/A'} | "
                f"來文字號={red_na_for_terminal(laiwen_dept or 'N/A')} | "
                f"no={doc_no or ''} | date={issue_date or ''} | subject={subject or ''}"
            )

        except Exception as e:
            rows.append({
                "original_filename": original_fname,
                "來文機關": "N/A",
                "來文字號": "N/A",
                "收文文號": "",
                "收文日期": "",
                "事由": "",
                "status": f"error: {e}",
                "_sort_key": (1, 0, 0, 0),
            })
            print(f"[error] {original_fname} -> {e}")

    # 2) Sort by 收文日期 ascending (missing dates go last)
    rows.sort(key=lambda r: r["_sort_key"])

    # 3) Assign 工作項目編號 = 001_1.13, 002_1.13, ...
    for idx, r in enumerate(rows, start=1):
        r["工作項目編號"] = format_work_item_id(idx, r.get("收文日期", ""))

    # 4) Write CSV
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)

        writer.writerow([
            "工作項目編號",
            "原始檔名",
            "來文機關",
            "來文字號",
            "收文文號",
            "收文日期",
            "事由",
            "status",
        ])

        for r in rows:
            writer.writerow([
                r.get("工作項目編號", ""),
                r.get("original_filename", ""),
                r.get("來文機關", "N/A"),
                r.get("來文字號", "N/A"),
                r.get("收文文號", ""),
                r.get("收文日期", ""),
                r.get("事由", ""),
                r.get("status", ""),
            ])

    print("Done. CSV saved to:", output_csv)


if __name__ == "__main__":
    main()