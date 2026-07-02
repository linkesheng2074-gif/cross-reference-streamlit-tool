# -*- coding: utf-8 -*-
"""
XTX NOR Flash Cross Reference Extractor

功能：
1) 读取 Cross Reference Excel 模板，保留原模板格式/合并单元格/列宽/颜色；
2) 从友商规格书 PDF 中抽取 ACDC 参数和 Command Opcode；
3) 按模板 Part Number 自动匹配 PDF，并输出相同格式的 Cross Reference；
4) 新增 Review_ACDC / Review_Command / PDF_Index 工作表，用于“模板已填值 vs PDF抽取值”的二次确认；
5) 支持 Streamlit 页面部署，也支持命令行批处理。

推荐运行：
    streamlit run cross_reference_streamlit_app.py

命令行运行：
    python cross_reference_streamlit_app.py --cli --template "XTX_NOR_Flash_High_Density_Cross_Reference_模板.xlsx" --specs "datasheets.zip" --out "CrossRef_Output.xlsx"

依赖：
    pip install streamlit openpyxl pymupdf pdfplumber
"""

from __future__ import annotations

import argparse
import copy
import difflib
import hashlib
import io
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import fitz  # PyMuPDF
except Exception as exc:  # pragma: no cover
    raise RuntimeError("缺少 PyMuPDF，请先执行：pip install pymupdf") from exc

try:
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except Exception as exc:  # pragma: no cover
    raise RuntimeError("缺少 openpyxl，请先执行：pip install openpyxl") from exc

try:
    import pdfplumber  # optional, improves table extraction
except Exception:  # pragma: no cover
    pdfplumber = None


# -----------------------------
# 1. Default extraction rules
# -----------------------------

DEFAULT_ACDC_RULES: Dict[str, Dict[str, object]] = {
    # DC current
    "Input Leakage Current": {"aliases": ["Input Leakage Current", "ILI", "I LI", "Input Leakage", "Leakage Current", "IOZ"], "units": ["uA", "µA", "μA"]},
    "Output Leakage Current": {"aliases": ["Output Leakage Current", "ILO", "I LO", "Output Leakage", "IOZ"], "units": ["uA", "µA", "μA"]},
    "stanby_current": {"aliases": ["Standby Current", "Stand-by Current", "standby current", "ICC1", "ICC2", "ICC standby", "Standby"], "units": ["uA", "µA", "μA", "mA"]},
    "deep_power_down_current": {"aliases": ["Deep Power-Down Current", "Deep Power Down Current", "DPD Current", "ICC DPD", "Ultra-Deep Power-Down Current"], "units": ["uA", "µA", "μA"]},
    "read_current(Single)": {"aliases": ["Single Read Current", "Read Current", "Active Read Current", "ICC Read", "ICC3", "Serial Read"], "units": ["mA", "uA", "µA", "μA"]},
    "read_current(Duad)": {"aliases": ["Dual Read Current", "Dual Output", "Dual I/O", "Dual IO", "ICC Read"], "units": ["mA", "uA", "µA", "μA"]},
    "read_current(Quad)": {"aliases": ["Quad Read Current", "Quad Output", "Quad I/O", "Quad IO", "Quad Read", "ICC Read"], "units": ["mA", "uA", "µA", "μA"]},
    "read_current(DTR Quad)": {"aliases": ["DTR Quad", "DDR Quad", "DTR Read Current", "DDR Read Current"], "units": ["mA", "uA", "µA", "μA"]},
    "program_current\nWRSR_current": {"aliases": ["Program Current", "Page Program Current", "Write Status Register Current", "WRSR Current", "ICC Program"], "units": ["mA", "uA", "µA", "μA"]},
    "erase_current": {"aliases": ["Erase Current", "Sector Erase Current", "Block Erase Current", "Chip Erase Current", "ICC Erase"], "units": ["mA", "uA", "µA", "μA"]},
    # DC voltage
    "Input Low Voltage": {"aliases": ["Input Low Voltage", "VIL", "V IL"], "units": ["V"]},
    "Input High Voltage": {"aliases": ["Input High Voltage", "VIH", "V IH"], "units": ["V"]},
    "Output Low Voltage": {"aliases": ["Output Low Voltage", "VOL", "V OL"], "units": ["V"]},
    "Output High Voltage": {"aliases": ["Output High Voltage", "VOH", "V OH"], "units": ["V"]},
    # Frequency / AC
    "fC1(Serial Clock Frequency For: all commands except Read (03H))": {"aliases": ["Clock Frequency", "Serial Clock Frequency", "fC", "f C", "SCLK frequency", "Frequency"], "units": ["MHz"]},
    "fC2(Serial Clock Frequency For: DTR Read)": {"aliases": ["DTR Clock Frequency", "DDR Clock Frequency", "DTR", "fC2"], "units": ["MHz"]},
    "fR(Serial Clock Frequency For: Read (03H))": {"aliases": ["Read Clock Frequency", "Read Frequency", "Read (03h)", "fR"], "units": ["MHz"]},
    "tCLH(Serial Clock High Time)": {"aliases": ["Serial Clock High Time", "Clock High Time", "tCLH", "tCH"], "units": ["ns"]},
    "tCLL(Serial Clock Low Time)": {"aliases": ["Serial Clock Low Time", "Clock Low Time", "tCLL", "tCL"], "units": ["ns"]},
    "tCLCH(Serial Clock Rise/Fall Time (Slew Rate))": {"aliases": ["Clock Rise Time", "Rise Time", "tCLCH", "tR"], "units": ["ns", "V/ns"]},
    "tCHCL(Serial Clock Rise/Fall Time (Slew Rate))": {"aliases": ["Clock Fall Time", "Fall Time", "tCHCL", "tF"], "units": ["ns", "V/ns"]},
    "tSLCH(CS# Active Setup Time)": {"aliases": ["CS# Active Setup Time", "CS# Setup Time", "Chip Select Setup", "tSLCH", "tCSS"], "units": ["ns"]},
    "tCHSH(CS# Active Hold Time)": {"aliases": ["CS# Active Hold Time", "CS# Hold Time", "tCHSH", "tCSH"], "units": ["ns"]},
    "tCLSH(CS# Active Hold Time)": {"aliases": ["CS# Active Hold Time", "tCLSH"], "units": ["ns"]},
    "tSHCH(CS# Not Active Setup Time)": {"aliases": ["CS# Not Active Setup Time", "CS# High Setup", "tSHCH"], "units": ["ns"]},
    "tCHSL(CS# Not Active Hold Time)": {"aliases": ["CS# Not Active Hold Time", "tCHSL"], "units": ["ns"]},
    "tSHSL(CS# High Time)": {"aliases": ["CS# High Time", "CS High Time", "tSHSL", "tCS"], "units": ["ns"]},
    "tSHQZ(Output Disable Time)": {"aliases": ["Output Disable Time", "tSHQZ", "tDIS"], "units": ["ns"]},
    "tCLQX(Output Hold Time)": {"aliases": ["Output Hold Time", "tCLQX", "tHO"], "units": ["ns"]},
    "tCLQV(Clock Transient To Output Valid)": {"aliases": ["Clock to Output Valid", "Output Valid", "tCLQV", "tV"], "units": ["ns"]},
    "tDVCH(Data In Setup Time)": {"aliases": ["Data In Setup Time", "Data Setup Time", "tDVCH", "tDS"], "units": ["ns"]},
    "tCHDX(Data In Hold Time)": {"aliases": ["Data In Hold Time", "Data Hold Time", "tCHDX", "tDH"], "units": ["ns"]},
    # Program / erase timing
    "tW(write status register Cycle Time)": {"aliases": ["Write Status Register Cycle Time", "Write Status Register", "tW"], "units": ["ms", "us", "µs", "μs"]},
    "byte program(First)": {"aliases": ["Byte Program", "First Byte Program", "tBP"], "units": ["us", "µs", "μs", "ms"]},
    "byte program": {"aliases": ["Byte Program", "tBP"], "units": ["us", "µs", "μs", "ms"]},
    "page program(256byte)": {"aliases": ["Page Program", "256 byte", "256-byte", "tPP"], "units": ["ms", "us", "µs", "μs"]},
    "page program(512byte)": {"aliases": ["Page Program", "512 byte", "512-byte", "tPP"], "units": ["ms", "us", "µs", "μs"]},
    "4k_erase": {"aliases": ["4KB Sector Erase", "4 Kbyte", "4KB Erase", "Subsector Erase", "tSE"], "units": ["ms", "s"]},
    "32k_erase": {"aliases": ["32KB Block Erase", "32KB Erase", "tBE32"], "units": ["ms", "s"]},
    "64k_erase": {"aliases": ["64KB Block Erase", "64KB Erase", "Sector Erase", "tSE", "tBE"], "units": ["ms", "s"]},
    "256K_erase": {"aliases": ["256KB Block Erase", "256KB Erase", "tBE256"], "units": ["ms", "s"]},
    "512Mb bulk erase time": {"aliases": ["Bulk Erase", "Die Erase", "512Mb Bulk", "tBE"], "units": ["s"]},
    "chip_erase": {"aliases": ["Chip Erase", "Bulk Erase", "tCE"], "units": ["s"]},
    # Power-on / reset / suspend
    "VWI": {"aliases": ["Write Inhibit Voltage", "VWI", "V WI"], "units": ["V"]},
    "VCC Power On ramp rate": {"aliases": ["VCC Power On ramp rate", "Power On ramp rate", "VCC Ramp"], "units": ["V/us", "V/µs", "V/μs", "V/ms"]},
    "VCC Power Down ramp rate": {"aliases": ["VCC Power Down ramp rate", "Power Down ramp rate"], "units": ["V/us", "V/µs", "V/μs", "V/ms"]},
    "tVSL": {"aliases": ["VCC(min) to CS# Low", "tVSL", "VSL"], "units": ["us", "µs", "μs", "ms"]},
    "tRLRH(Reset Pulse Width)": {"aliases": ["Reset Pulse Width", "RESET# Pulse Width", "tRLRH", "tRP"], "units": ["ns", "us", "µs", "μs"]},
    "tRHSL(Reset Hold time before next Operation)": {"aliases": ["Reset Hold time", "Reset Hold", "tRHSL"], "units": ["ns", "us", "µs", "μs"]},
    "tRB(Reset Recovery Time)": {"aliases": ["Reset Recovery Time", "tRB", "tRST"], "units": ["us", "µs", "μs", "ms"]},
    "tSUS(CS# High To Next Command After Suspend)": {"aliases": ["Suspend Latency", "Suspend", "tSUS"], "units": ["us", "µs", "μs", "ms"]},
    "tRS(Latency Between Resume And Next Suspend)": {"aliases": ["Latency Between Resume", "Resume latency", "tRS"], "units": ["us", "µs", "μs", "ms"]},
    "Suspend latency": {"aliases": ["Suspend latency", "Suspend Latency", "tSUS"], "units": ["us", "µs", "μs", "ms"]},
}

# 命令抽取：建议通过 rules/crossref_rules.json 持续补充别名。
DEFAULT_COMMAND_ALIASES: Dict[str, List[str]] = {
    "enabale_reset": ["Enable Reset", "Reset Enable", "Software Reset Enable", "RSTEN"],
    "reset": ["Reset", "Software Reset", "Reset Memory", "RST"],
    "Mode Bit Reset": ["Mode Bit Reset", "MBR"],
    "NOP": ["No Operation", "NOP"],
    "Legacy Software Reset": ["Legacy Software Reset", "Software Reset"],
    "read_manufacture_id": ["Read Manufacturer", "Read Manufacture", "Manufacturer ID", "Read ID"],
    "read_manufacture_id_d_io)": ["Dual I/O Read ID", "Dual IO Read ID", "Read Manufacturer/Device ID Dual"],
    "read_manufacture_id_q_io)": ["Quad I/O Read ID", "Quad IO Read ID", "Read Manufacturer/Device ID Quad"],
    "read_jedec_id": ["Read JEDEC ID", "JEDEC ID", "Read Identification"],
    "read_unique_id": ["Read Unique ID", "Unique ID", "Read Unique Identifier"],
    "read_sfpd": ["Read SFDP", "SFDP", "Serial Flash Discoverable Parameters"],
    "Read Electronic Signature": ["Read Electronic Signature", "RES"],
    "READ STATUS REGISTER-1": ["Read Status Register-1", "Read Status Register", "RDSR", "Status Register-1"],
    "READ STATUS REGISTER-2": ["Read Status Register-2", "Status Register-2"],
    "READ STATUS REGISTER-3": ["Read Status Register-3", "Status Register-3"],
    "read_config_register": ["Read Configuration Register", "Configuration Register", "RDCR"],
    "READ FLAG STATUS REGISTER": ["Read Flag Status Register", "Flag Status Register"],
    "READ NONVOLATILE CONFIGURATION REGISTER": ["Read Nonvolatile Configuration Register", "Nonvolatile Configuration Register"],
    "READ VOLATILE CONFIGURATION REGISTER": ["Read Volatile Configuration Register", "Volatile Configuration Register"],
    "READ ENHANCED VOLATILE CONFIGURATION REGISTER": ["Read Enhanced Volatile Configuration Register", "Enhanced Volatile Configuration Register"],
    "READ EXTENDED ADDRESS REGISTER": ["Read Extended Address Register", "Extended Address Register"],
    "READ ANY REGISTER": ["Read Any Register"],
    "WRITE STATUS REGISTER-1": ["Write Status Register-1", "Write Status Register", "WRSR"],
    "WRITE STATUS REGISTER-2": ["Write Status Register-2"],
    "WRITE STATUS REGISTER-3": ["Write Status Register-3"],
    "write_enable": ["Write Enable", "WREN"],
    "write_disable": ["Write Disable", "WRDI"],
    "write_enable_volatile": ["Write Enable for Volatile", "Volatile Write Enable"],
    "CLEAR FLAG STATUS REGISTER": ["Clear Flag Status Register"],
    "CLEAR ECC STATUS REGISTER": ["Clear ECC Status Register"],
    "ENTER 4-BYTE ADDRESS MODE": ["Enter 4-Byte Address Mode", "Enter 4 Byte Address Mode", "EN4B"],
    "EXIT 4-BYTE ADDRESS MODE": ["Exit 4-Byte Address Mode", "Exit 4 Byte Address Mode", "EX4B"],
    "enable_qpi": ["Enter QPI", "Enable QPI", "QPI Enable"],
    "disable_qpi": ["Exit QPI", "Disable QPI", "QPI Disable"],
    "Set Read Parameter": ["Set Read Parameters", "Set Read Parameter"],
    "Set Burst with Wrap": ["Set Burst with Wrap", "Burst with Wrap"],
    "normal_read": ["Read Data", "Normal Read", "Read", "Read Memory"],
    "fast_read": ["Fast Read", "Fast Read Data"],
    "dual_output_fast_read": ["Dual Output Fast Read", "Dual Output Read"],
    "dual_io_fast_read": ["Dual I/O Fast Read", "Dual IO Fast Read", "Dual Input/Output Fast Read"],
    "quad_output_fast_read": ["Quad Output Fast Read", "Quad Output Read"],
    "quad_io_fast_read": ["Quad I/O Fast Read", "Quad IO Fast Read", "Quad Input/Output Fast Read"],
    "quad_Io_word_read": ["Quad I/O Word Read", "Quad IO Word Read"],
    "Burst Read with Wrap": ["Burst Read with Wrap", "Wrap Read"],
    "DTR FAST READ": ["DTR Fast Read", "DDR Fast Read"],
    "DTR DUAL OUTPUT FAST READ": ["DTR Dual Output Fast Read", "DDR Dual Output Fast Read"],
    "DTR DUAL INPUT/OUTPUT FAST\nREAD": ["DTR Dual I/O Fast Read", "DTR Dual Input/Output Fast Read"],
    "DTR QUAD OUTPUT FAST READ": ["DTR Quad Output Fast Read", "DDR Quad Output Fast Read"],
    "DTR QUAD INPUT/OUTPUT FAST READ": ["DTR Quad I/O Fast Read", "DTR Quad Input/Output Fast Read", "DDR Quad I/O Fast Read"],
    "4-BYTE READ": ["4-Byte Read", "4 Byte Read"],
    "4-BYTE FAST READ": ["4-Byte Fast Read", "4 Byte Fast Read"],
    "4-BYTE DUAL OUTPUT FAST READ": ["4-Byte Dual Output Fast Read"],
    "4-BYTE DUAL INPUT/OUTPUT FAST READ": ["4-Byte Dual I/O Fast Read", "4-Byte Dual Input/Output Fast Read"],
    "4-BYTE QUAD OUTPUT FAST READ": ["4-Byte Quad Output Fast Read"],
    "4-BYTE QUAD INPUT/OUTPUT FAST READ": ["4-Byte Quad I/O Fast Read", "4-Byte Quad Input/Output Fast Read"],
    "PAGE PROGRAM": ["Page Program"],
    "QUAD INPUT FAST PROGRAM": ["Quad Input Fast Program", "Quad Input Page Program"],
    "DUAL INPUT FAST PROGRAM": ["Dual Input Fast Program", "Dual Input Page Program"],
    "4-BYTE PAGE PROGRAM": ["4-Byte Page Program", "4 Byte Page Program"],
    "4-BYTE QUAD INPUT FAST PROGRAM": ["4-Byte Quad Input Fast Program"],
    "sector_erase": ["Sector Erase", "4KB Sector Erase", "Subsector Erase"],
    "32k_block_erase": ["32KB Block Erase", "32 Kbyte Block Erase"],
    "64k_block_erase": ["64KB Block Erase", "64 Kbyte Block Erase", "Block Erase"],
    "256K_block_erase": ["256KB Block Erase", "256 Kbyte Block Erase"],
    "chip_erase": ["Chip Erase", "Bulk Erase"],
    "program_erase_suspend": ["Program/Erase Suspend", "Suspend"],
    "program_erase_resume": ["Program/Erase Resume", "Resume"],
    "erase_suspend": ["Erase Suspend"],
    "erase_resume": ["Erase Resume"],
    "program_suspend": ["Program Suspend"],
    "program_resume": ["Program Resume"],
    "erase_security_register": ["Erase Security Register", "Erase OTP"],
    "program_security_register": ["Program Security Register", "Program OTP"],
    "read_security_register": ["Read Security Register", "Read OTP"],
    "deep_powerdown": ["Deep Power-Down", "Deep Power Down"],
    "release_from_deep_powerdown": ["Release from Deep Power-Down", "Release from Power-Down", "Release Power-Down"],
    "CYCLIC REDUNDANCY CHECK": ["Cyclic Redundancy Check", "CRC"],
}

SECTION_PREFIXES = (
    "一、", "二、", "三、", "四、", "五、", "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10."
)


@dataclass
class PdfInfo:
    source_file: str
    vendor_folder: str
    detected_part: str
    title_part_candidates: str
    pages: int
    sha1: str
    matched_columns: str = ""
    match_score: float = 0.0


@dataclass
class ExtractedValue:
    value: str
    confidence: float
    evidence: str
    source_file: str
    source_page: Optional[int] = None


@dataclass
class ReviewRow:
    sheet: str
    vendor: str
    part_number: str
    item: str
    template_value: str
    extracted_value: str
    match_status: str
    similarity: float
    confidence: float
    source_file: str
    evidence: str
    row: int
    column: int


# -----------------------------
# 2. Utilities
# -----------------------------

def safe_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:31]


def clean_text(s: object) -> str:
    if s is None:
        return ""
    text = str(s)
    text = text.replace("_x000D_", "\n")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\u00a0", " ")
    text = text.replace("µ", "u").replace("μ", "u")
    text = text.replace("℃", "C")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def norm_key(s: object) -> str:
    s = clean_text(s).lower()
    s = s.replace("duad", "dual")
    return re.sub(r"[^a-z0-9]+", "", s)


def norm_for_compare(s: object) -> str:
    s = clean_text(s).lower()
    s = s.replace("maximum", "max").replace("minimum", "min").replace("typical", "typ")
    s = s.replace("µ", "u").replace("μ", "u").replace("℃", "c")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^a-z0-9+\-./@%<>:=]", "", s)
    return s


def normalize_part(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", clean_text(s).upper())


def short_filename(path: str) -> str:
    return Path(path).name


def calculate_sha1(path: Path) -> str:
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()[:12]


def is_section_item(item: object) -> bool:
    s = clean_text(item)
    if not s:
        return True
    if any(s.startswith(prefix) for prefix in SECTION_PREFIXES):
        return True
    # Chinese section header or broad header rows
    if "Characteristics" in s and not re.search(r"[a-zA-Z]\w+\(", s):
        return True
    if s in {"Read register operation", "Write register operation", "Write operation", "Read memory operation", "DTR Read memory operation", "Program memory operation", "Erase memory operation"}:
        return True
    if s.endswith("operation") or s.endswith("Operation") or s.endswith("Operations"):
        # Some rows are just group headers. Leave specific commands to alias rules.
        return True
    return False


def unpack_specs(input_path: Path, work_dir: Path) -> Path:
    """Return directory that contains PDFs. Supports zip or folder."""
    input_path = Path(input_path)
    if input_path.is_dir():
        return input_path
    if input_path.suffix.lower() != ".zip":
        raise ValueError("规格书输入必须是 PDF 文件夹或 .zip 压缩包。")
    target = work_dir / "unzipped_specs"
    target.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(input_path, "r") as z:
        # Defend against zip-slip
        for member in z.infolist():
            member_path = target / member.filename
            if not str(member_path.resolve()).startswith(str(target.resolve())):
                raise ValueError(f"压缩包中存在不安全路径：{member.filename}")
        z.extractall(target)
    return target


def list_pdf_files(spec_dir: Path) -> List[Path]:
    pdfs = []
    for p in spec_dir.rglob("*"):
        if p.is_file() and p.suffix.lower() == ".pdf":
            pdfs.append(p)
    return sorted(pdfs)


def locate_template_xlsx(spec_dir: Path) -> Optional[Path]:
    candidates = []
    for p in spec_dir.rglob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        candidates.append(p)
    if not candidates:
        return None
    # Prefer Cross Reference file name
    candidates.sort(key=lambda p: ("cross" not in p.name.lower(), len(str(p))))
    return candidates[0]


# -----------------------------
# 3. PDF text / table extraction
# -----------------------------

def extract_pdf_pages(pdf_path: Path, max_pages: Optional[int] = None) -> List[Tuple[int, str]]:
    pages: List[Tuple[int, str]] = []
    doc = fitz.open(str(pdf_path))
    page_count = doc.page_count if max_pages is None else min(doc.page_count, max_pages)
    for i in range(page_count):
        text = doc.load_page(i).get_text("text") or ""
        text = clean_text(text)
        pages.append((i + 1, text))
    return pages


def extract_pdf_text(pdf_path: Path, max_pages: Optional[int] = None) -> str:
    return "\n".join(text for _, text in extract_pdf_pages(pdf_path, max_pages=max_pages))


def extract_pdf_tables_as_lines(pdf_path: Path, max_pages: Optional[int] = None) -> List[Tuple[int, str]]:
    lines: List[Tuple[int, str]] = []
    if pdfplumber is None:
        return lines
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            pages = pdf.pages if max_pages is None else pdf.pages[:max_pages]
            for idx, page in enumerate(pages, start=1):
                try:
                    tables = page.extract_tables() or []
                except Exception:
                    tables = []
                for table in tables:
                    for row in table or []:
                        cells = [clean_text(c) for c in (row or []) if clean_text(c)]
                        if len(cells) >= 2:
                            lines.append((idx, " | ".join(cells)))
    except Exception:
        return lines
    return lines


def detect_part_from_text_and_filename(pdf_path: Path, first_pages_text: str) -> Tuple[str, List[str]]:
    name = pdf_path.stem.upper()
    text = first_pages_text[:5000].upper()
    candidates = set()

    patterns = [
        r"\b(?:GD|W25|MX25|MX66|MT25|S25|S70|XM25|EN25|IS25|DS25|GT25|XT25|XT55)[A-Z0-9\-_/\.]{4,28}\b",
        r"\bW25Q[0-9A-Z]{4,20}\b",
        r"\bMX(?:25|66)[0-9A-Z]{4,24}\b",
        r"\bMT25Q[UL][0-9A-Z]{4,24}\b",
        r"\bGD(?:25|55)[A-Z0-9]{4,24}\b",
        r"\bS(?:25|70)[A-Z0-9]{4,24}\b",
    ]
    for blob in [name, text]:
        for pat in patterns:
            for m in re.finditer(pat, blob):
                token = m.group(0).strip("_-.,;:()[]{}")
                if len(token) >= 7 and not token.endswith("PDF"):
                    candidates.add(token)

    # Filename often includes revision/noise; take first high quality part candidate.
    sorted_candidates = sorted(candidates, key=lambda x: (-len(normalize_part(x)), x))
    if not sorted_candidates:
        return "", []
    # Prefer tokens that appear in filename.
    filename_norm = normalize_part(name)
    sorted_candidates.sort(key=lambda x: (normalize_part(x) not in filename_norm, -len(normalize_part(x))))
    return sorted_candidates[0], sorted_candidates[:10]


def build_pdf_info(pdf_path: Path, base_dir: Path) -> PdfInfo:
    rel = pdf_path.relative_to(base_dir)
    # Use the immediate parent folder as vendor folder, e.g. Cypress/GD/MXIC/Winbond.
    vendor_folder = rel.parts[-2] if len(rel.parts) > 1 else pdf_path.parent.name
    first_text = extract_pdf_text(pdf_path, max_pages=3)
    try:
        pages = fitz.open(str(pdf_path)).page_count
    except Exception:
        pages = 0
    part, candidates = detect_part_from_text_and_filename(pdf_path, first_text)
    return PdfInfo(
        source_file=str(pdf_path),
        vendor_folder=vendor_folder,
        detected_part=part,
        title_part_candidates=", ".join(candidates),
        pages=pages,
        sha1=calculate_sha1(pdf_path),
    )



def find_snippets_with_alias_text(pages: List[Tuple[int, str]], aliases: Sequence[str], window_chars: int = 900, max_hits: int = 5) -> List[Tuple[int, str]]:
    """Fast literal search on page text. This avoids scanning every PDF line repeatedly."""
    clean_aliases = [clean_text(a).lower() for a in aliases if clean_text(a)]
    # Prefer longer, more specific aliases first.
    clean_aliases.sort(key=len, reverse=True)
    hits: List[Tuple[int, str]] = []
    for page_no, page_text in pages:
        lower = page_text.lower()
        for alias in clean_aliases:
            pos = lower.find(alias)
            if pos >= 0:
                start = max(0, pos - window_chars // 3)
                end = min(len(page_text), pos + window_chars)
                snippet = clean_text(page_text[start:end])
                hits.append((page_no, snippet))
                break
        if len(hits) >= max_hits:
            break
    return hits


def find_lines_with_alias(pages: List[Tuple[int, str]], aliases: Sequence[str], window: int = 4) -> List[Tuple[int, str]]:
    alias_keys = [norm_key(a) for a in aliases if a]
    hits: List[Tuple[int, str]] = []
    for page_no, page_text in pages:
        raw_lines = [clean_text(x) for x in page_text.splitlines() if clean_text(x)]
        line_keys = [norm_key(x) for x in raw_lines]
        for i, key in enumerate(line_keys):
            if any(a and a in key for a in alias_keys):
                start = max(0, i - 1)
                end = min(len(raw_lines), i + window)
                snippet = " | ".join(raw_lines[start:end])
                hits.append((page_no, snippet))
                if len(hits) >= 5:
                    return hits
    return hits


def find_table_line_with_alias(table_lines: List[Tuple[int, str]], aliases: Sequence[str]) -> Optional[Tuple[int, str]]:
    alias_keys = [norm_key(a) for a in aliases if a]
    best = None
    best_score = 0
    for page_no, line in table_lines:
        key = norm_key(line)
        score = max((len(a) for a in alias_keys if a and a in key), default=0)
        if score > best_score:
            best = (page_no, line)
            best_score = score
    return best


def value_patterns_for_units(units: Sequence[str]) -> List[re.Pattern]:
    # Normalize micro units to u before applying.
    unit_group = "|".join(sorted({re.escape(u.replace("µ", "u").replace("μ", "u")) for u in units}, key=len, reverse=True))
    if not unit_group:
        unit_group = r"uA|mA|V|MHz|ns|us|ms|s"
    # Capture common min/typ/max triples and conditions such as @85C, @133MHz.
    return [
        re.compile(rf"(?i)(?:min\.?|typ\.?|max\.?|maximum|minimum|typical)?\s*[<>≤≥=:\-]*\s*[+-]?\d+(?:\.\d+)?\s*(?:/{0,1}\s*[+-]?\d+(?:\.\d+)?\s*)?(?:{unit_group})(?:\s*(?:max|min|typ))?(?:\s*@\s*[\w./+\-°]+)?"),
        re.compile(rf"(?i)[+-]?\d+(?:\.\d+)?\s*/\s*[+-]?\d+(?:\.\d+)?\s*(?:{unit_group})(?:\s*@\s*[\w./+\-°]+)?"),
        re.compile(rf"(?i)(?:[0-9]+(?:\.[0-9]+)?\s*(?:{unit_group}))(?:\s*@\s*[\w./+\-°]+)?"),
    ]


def extract_value_from_snippet(snippet: str, units: Sequence[str]) -> str:
    normalized = clean_text(snippet)
    normalized = normalized.replace("µ", "u").replace("μ", "u")
    patterns = value_patterns_for_units(units)
    values: List[str] = []
    for pat in patterns:
        for m in pat.finditer(normalized):
            token = clean_text(m.group(0))
            # Filter out page numbers or unrelated tiny tokens.
            if re.search(r"(?i)(uA|mA|V|MHz|ns|us|ms|s|V/us|V/ms)", token):
                values.append(token)
        if values:
            break
    # Remove obvious duplicate tokens while preserving order.
    seen = set()
    uniq = []
    for v in values:
        key = norm_for_compare(v)
        if key not in seen:
            seen.add(key)
            uniq.append(v)
    if uniq:
        return "\n".join(uniq[:8])
    return ""


def extract_acdc_value(pdf_path: Path, pages: List[Tuple[int, str]], table_lines: List[Tuple[int, str]], rule: Dict[str, object]) -> Optional[ExtractedValue]:
    aliases = rule.get("aliases", []) if isinstance(rule, dict) else []
    units = rule.get("units", []) if isinstance(rule, dict) else []
    if not aliases:
        return None

    # Prefer table rows because spec values are usually tabulated.
    table_hit = find_table_line_with_alias(table_lines, aliases)
    if table_hit:
        page, line = table_hit
        value = extract_value_from_snippet(line, units)
        evidence = line[:700]
        if value:
            return ExtractedValue(value=value, confidence=0.72, evidence=evidence, source_file=str(pdf_path), source_page=page)
        return ExtractedValue(value=evidence[:300], confidence=0.35, evidence=evidence, source_file=str(pdf_path), source_page=page)

    hits = find_snippets_with_alias_text(pages, aliases, window_chars=1000)
    if not hits:
        return None
    for page, snippet in hits:
        value = extract_value_from_snippet(snippet, units)
        if value:
            return ExtractedValue(value=value, confidence=0.55, evidence=snippet[:700], source_file=str(pdf_path), source_page=page)
    page, snippet = hits[0]
    return ExtractedValue(value="", confidence=0.20, evidence=snippet[:700], source_file=str(pdf_path), source_page=page)


def opcode_regex() -> re.Pattern:
    return re.compile(r"(?i)(?:opcode|code|instruction|command)?\s*[:=\-]?\s*\b([0-9A-F]{2})\s*H\b|\b([0-9A-F]{2})h\b|\b0x([0-9A-F]{2})\b")


def extract_opcode_from_snippet(snippet: str) -> str:
    # Avoid IDs like 85C/105C; regex only H/0x.
    values: List[str] = []
    for m in opcode_regex().finditer(snippet):
        code = next((g for g in m.groups() if g), "")
        if code:
            values.append(code.upper() + "H")
    seen = set()
    uniq = []
    for v in values:
        if v not in seen:
            seen.add(v)
            uniq.append(v)
    return "/".join(uniq[:3])


def extract_command_value(pdf_path: Path, pages: List[Tuple[int, str]], table_lines: List[Tuple[int, str]], item: str, aliases: Sequence[str]) -> Optional[ExtractedValue]:
    if not aliases:
        aliases = [item]

    # Table first.
    hit = find_table_line_with_alias(table_lines, aliases)
    if hit:
        page, line = hit
        opcode = extract_opcode_from_snippet(line)
        if opcode:
            suffix = "(QPI)" if re.search(r"\bQPI\b", line, flags=re.I) else ""
            return ExtractedValue(value=opcode + suffix, confidence=0.78, evidence=line[:700], source_file=str(pdf_path), source_page=page)

    hits = find_snippets_with_alias_text(pages, aliases, window_chars=900)
    for page, snippet in hits:
        opcode = extract_opcode_from_snippet(snippet)
        if opcode:
            suffix = "(QPI)" if re.search(r"\bQPI\b", snippet, flags=re.I) else ""
            return ExtractedValue(value=opcode + suffix, confidence=0.55, evidence=snippet[:700], source_file=str(pdf_path), source_page=page)
    if hits:
        page, snippet = hits[0]
        return ExtractedValue(value="", confidence=0.20, evidence=snippet[:700], source_file=str(pdf_path), source_page=page)
    return None


# -----------------------------
# 4. Workbook mapping / writing
# -----------------------------

def get_template_items(ws) -> Dict[int, str]:
    items = {}
    for row in range(1, ws.max_row + 1):
        value = clean_text(ws.cell(row, 2).value)
        if value:
            items[row] = value
    return items


def get_header_columns(ws) -> Dict[int, Tuple[str, str]]:
    cols = {}
    for col in range(3, ws.max_column + 1):
        vendor = clean_text(ws.cell(2, col).value)
        part = clean_text(ws.cell(3, col).value)
        if vendor or part:
            cols[col] = (vendor, part)
    return cols


def vendor_compatible(template_vendor: str, folder_vendor: str) -> bool:
    t = norm_key(template_vendor)
    f = norm_key(folder_vendor)
    if not t or not f:
        return True
    aliases = {
        "spansion": {"cypress", "spansion", "infineon"},
        "cypress": {"cypress", "spansion", "infineon"},
        "winbond": {"winbond"},
        "gd": {"gd", "giga", "gigadevice"},
        "mxic": {"mxic", "macronix"},
        "macronix": {"mxic", "macronix"},
        "micron": {"micron"},
    }
    for k, vals in aliases.items():
        if k in t:
            return any(v in f for v in vals)
    return t in f or f in t


def match_score_for_pdf_to_column(pdf_info: PdfInfo, template_vendor: str, template_part: str) -> float:
    if not template_part:
        return 0.0
    part_norm = normalize_part(template_part)
    if "目标" in template_part or "XT" in part_norm[:3]:
        # Usually target XTX columns should not be filled from competitor PDFs.
        return 0.0
    vendor_bonus = 0.08 if vendor_compatible(template_vendor, pdf_info.vendor_folder) else -0.15
    filename_norm = normalize_part(Path(pdf_info.source_file).stem)
    detected_norm = normalize_part(pdf_info.detected_part)
    candidates_norm = [normalize_part(x) for x in pdf_info.title_part_candidates.split(",") if x.strip()]

    scores = []
    for cand in [detected_norm, filename_norm] + candidates_norm:
        if not cand:
            continue
        if part_norm and part_norm in cand:
            scores.append(0.92 + vendor_bonus)
        elif cand and cand in part_norm:
            scores.append(0.88 + vendor_bonus)
        else:
            ratio = difflib.SequenceMatcher(None, part_norm, cand).ratio()
            # Avoid false matches such as GD25LB512ME -> GD25LT512ME.
            # Non-contained fuzzy matches must be very close to pass.
            # For cross reference, false-positive column filling is more harmful than missing a match.
            # Therefore non-contained fuzzy matches are capped below the default threshold.
            scores.append(min(ratio + vendor_bonus, 0.65))
    return max(scores or [0.0])


def map_pdfs_to_columns(wb, pdf_infos: List[PdfInfo], threshold: float = 0.70) -> Dict[Tuple[str, int], PdfInfo]:
    mapping: Dict[Tuple[str, int], PdfInfo] = {}
    for ws in wb.worksheets:
        if not ws.title.startswith("Cross Reference"):
            continue
        cols = get_header_columns(ws)
        for col, (vendor, part) in cols.items():
            best_pdf = None
            best_score = 0.0
            for pdf_info in pdf_infos:
                score = match_score_for_pdf_to_column(pdf_info, vendor, part)
                if score > best_score:
                    best_pdf, best_score = pdf_info, score
            if best_pdf and best_score >= threshold:
                mapping[(ws.title, col)] = best_pdf
    # Update pdf_info matched_columns summary.
    for pdf_info in pdf_infos:
        matched = []
        max_score = 0.0
        for (sheet, col), pinfo in mapping.items():
            if pinfo.source_file == pdf_info.source_file:
                matched.append(f"{sheet}!{get_column_letter(col)}")
                max_score = max(max_score, match_score_for_pdf_to_column(pdf_info, "", ""))
        pdf_info.matched_columns = ", ".join(matched)
    return mapping


def build_merged_anchor_map(ws) -> Dict[Tuple[int, int], Tuple[int, int]]:
    anchors: Dict[Tuple[int, int], Tuple[int, int]] = {}
    for rng in ws.merged_cells.ranges:
        for row in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                anchors[(row, col)] = (rng.min_row, rng.min_col)
    return anchors


def set_cell_value_safely(ws, row: int, col: int, value: object, anchor_map: Dict[Tuple[int, int], Tuple[int, int]]) -> None:
    arow, acol = anchor_map.get((row, col), (row, col))
    ws.cell(arow, acol).value = value
    ws.cell(arow, acol).alignment = copy.copy(ws.cell(arow, acol).alignment)
    ws.cell(arow, acol).alignment = Alignment(
        horizontal=ws.cell(arow, acol).alignment.horizontal,
        vertical=ws.cell(arow, acol).alignment.vertical,
        wrap_text=True,
    )


def clear_cell_safely(ws, row: int, col: int, anchor_map: Dict[Tuple[int, int], Tuple[int, int]]) -> None:
    arow, acol = anchor_map.get((row, col), (row, col))
    if (arow, acol) == (row, col):
        ws.cell(row, col).value = None


def similarity_status(template_value: str, extracted_value: str) -> Tuple[str, float]:
    t = norm_for_compare(template_value)
    e = norm_for_compare(extracted_value)
    if not t and not e:
        return "Both Blank", 1.0
    if t and not e:
        return "Missing", 0.0
    if not t and e:
        return "New Extracted", 0.0
    if t == e:
        return "Exact", 1.0
    ratio = difflib.SequenceMatcher(None, t, e).ratio()
    if ratio >= 0.86:
        return "Similar", ratio
    if ratio >= 0.55:
        return "Check", ratio
    return "Different", ratio


def load_or_create_rules(rules_path: Optional[Path]) -> Dict[str, object]:
    default_rules = {"acdc": DEFAULT_ACDC_RULES, "command_aliases": DEFAULT_COMMAND_ALIASES}
    if rules_path is None:
        return default_rules
    rules_path = Path(rules_path)
    if not rules_path.exists():
        rules_path.parent.mkdir(parents=True, exist_ok=True)
        with open(rules_path, "w", encoding="utf-8") as f:
            json.dump(default_rules, f, ensure_ascii=False, indent=2)
        return default_rules
    with open(rules_path, "r", encoding="utf-8") as f:
        user_rules = json.load(f)
    # Merge user rules over default rules.
    merged = default_rules
    if isinstance(user_rules, dict):
        for k in ["acdc", "command_aliases"]:
            if isinstance(user_rules.get(k), dict):
                merged[k].update(user_rules[k])
    return merged


def extract_values_for_pdf(
    pdf_path: Path,
    acdc_items: Iterable[str],
    command_items: Iterable[str],
    rules: Dict[str, object],
    max_pages: Optional[int] = None,
    use_pdfplumber: bool = False,
) -> Tuple[Dict[str, ExtractedValue], Dict[str, ExtractedValue]]:
    pages = extract_pdf_pages(pdf_path, max_pages=max_pages)
    table_lines = extract_pdf_tables_as_lines(pdf_path, max_pages=max_pages) if use_pdfplumber else []
    acdc_rules = rules.get("acdc", {}) if isinstance(rules, dict) else {}
    command_aliases = rules.get("command_aliases", {}) if isinstance(rules, dict) else {}

    acdc_results: Dict[str, ExtractedValue] = {}
    for item in acdc_items:
        if is_section_item(item):
            continue
        rule = acdc_rules.get(item) or acdc_rules.get(clean_text(item))
        if not rule:
            # fallback uses the item itself as alias and broad units
            rule = {"aliases": [item], "units": ["uA", "mA", "V", "MHz", "ns", "us", "ms", "s"]}
        val = extract_acdc_value(pdf_path, pages, table_lines, rule)
        if val and (val.value or val.evidence):
            acdc_results[item] = val

    cmd_results: Dict[str, ExtractedValue] = {}
    for item in command_items:
        if is_section_item(item):
            continue
        aliases = command_aliases.get(item) or command_aliases.get(clean_text(item)) or [item]
        val = extract_command_value(pdf_path, pages, table_lines, item, aliases)
        if val and (val.value or val.evidence):
            cmd_results[item] = val
    return acdc_results, cmd_results


def ensure_review_sheet(wb, sheet_name: str):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    headers = [
        "Sheet", "Vendor", "Part Number", "Item", "Template_Value", "Extracted_Value",
        "Match_Status", "Similarity", "Confidence", "Source_File", "Evidence", "Cell"
    ]
    ws.append(headers)
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L1"
    widths = [22, 14, 24, 34, 34, 34, 14, 12, 12, 42, 70, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def style_review_row(row_cells, status: str):
    colors = {
        "Exact": "C6EFCE",
        "Similar": "D9EAD3",
        "Check": "FFF2CC",
        "Different": "F4CCCC",
        "Missing": "FCE4D6",
        "New Extracted": "DDEBF7",
        "Both Blank": "E7E6E6",
    }
    fill = PatternFill("solid", fgColor=colors.get(status, "FFFFFF"))
    for cell in row_cells:
        cell.fill = fill
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="D9D9D9"))


def create_pdf_index_sheet(wb, pdf_infos: List[PdfInfo], mapping: Dict[Tuple[str, int], PdfInfo]):
    if "PDF_Index" in wb.sheetnames:
        del wb["PDF_Index"]
    ws = wb.create_sheet("PDF_Index", 0)
    headers = ["Vendor_Folder", "Detected_Part", "Candidates", "Pages", "SHA1", "Matched_Columns", "Source_File"]
    ws.append(headers)
    for pinfo in pdf_infos:
        matched = []
        for (sheet, col), mpinfo in mapping.items():
            if mpinfo.source_file == pinfo.source_file:
                matched.append(f"{sheet}!{get_column_letter(col)}")
        ws.append([pinfo.vendor_folder, pinfo.detected_part, pinfo.title_part_candidates, pinfo.pages, pinfo.sha1, ", ".join(matched), short_filename(pinfo.source_file)])
    fill = PatternFill("solid", fgColor="548235")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    widths = [18, 24, 60, 10, 16, 55, 75]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{ws.max_row}"
    return ws


def generate_cross_reference(
    template_xlsx: Path,
    specs_input: Path,
    output_xlsx: Path,
    rules_path: Optional[Path] = None,
    match_threshold: float = 0.70,
    clear_matched_cells: bool = True,
    max_pages: Optional[int] = None,
    use_pdfplumber: bool = False,
    progress_callback=None,
) -> Dict[str, object]:
    start = time.time()
    work_dir = Path(tempfile.mkdtemp(prefix="xtx_crossref_"))
    try:
        specs_dir = unpack_specs(Path(specs_input), work_dir)
        if template_xlsx is None or not Path(template_xlsx).exists():
            found_template = locate_template_xlsx(specs_dir)
            if found_template is None:
                raise FileNotFoundError("没有找到 Cross Reference 模板 xlsx。")
            template_xlsx = found_template

        pdf_paths = list_pdf_files(specs_dir)
        if not pdf_paths:
            raise FileNotFoundError("规格书目录/压缩包中没有找到 PDF 文件。")

        if progress_callback:
            progress_callback(f"发现 {len(pdf_paths)} 个 PDF，开始读取模板。")

        wb = load_workbook(str(template_xlsx))
        if not any(s.startswith("Cross Reference") for s in wb.sheetnames):
            raise ValueError("模板中没有找到 Cross Reference 工作表。")

        # Cache template values before overwriting.
        template_values: Dict[Tuple[str, int, int], str] = {}
        acdc_items: List[str] = []
        command_items: List[str] = []
        for ws in wb.worksheets:
            if ws.title == "Cross Reference--ACDC":
                acdc_items = list(get_template_items(ws).values())
            elif ws.title == "Cross Reference--Command":
                command_items = list(get_template_items(ws).values())
            if ws.title.startswith("Cross Reference"):
                for row in range(1, ws.max_row + 1):
                    for col in range(1, ws.max_column + 1):
                        template_values[(ws.title, row, col)] = clean_text(ws.cell(row, col).value)

        rules = load_or_create_rules(rules_path)

        pdf_infos: List[PdfInfo] = []
        for idx, pdf in enumerate(pdf_paths, start=1):
            if progress_callback:
                progress_callback(f"识别 PDF：{idx}/{len(pdf_paths)} - {pdf.name}")
            try:
                pdf_infos.append(build_pdf_info(pdf, specs_dir))
            except Exception as exc:
                pdf_infos.append(PdfInfo(str(pdf), pdf.parent.name, "", f"ERROR: {exc}", 0, calculate_sha1(pdf)))

        mapping = map_pdfs_to_columns(wb, pdf_infos, threshold=match_threshold)
        if progress_callback:
            progress_callback(f"完成 PDF 与模板列匹配：{len(mapping)} 个列匹配成功。")

        # Extract per unique PDF only once.
        unique_pdf_paths = sorted({pinfo.source_file for pinfo in mapping.values()})
        extracted_cache: Dict[str, Tuple[Dict[str, ExtractedValue], Dict[str, ExtractedValue]]] = {}
        for idx, pdf_str in enumerate(unique_pdf_paths, start=1):
            pdf = Path(pdf_str)
            if progress_callback:
                progress_callback(f"抽取 PDF：{idx}/{len(unique_pdf_paths)} - {pdf.name}")
            extracted_cache[pdf_str] = extract_values_for_pdf(pdf, acdc_items, command_items, rules, max_pages=max_pages, use_pdfplumber=use_pdfplumber)

        review_acdc = ensure_review_sheet(wb, "Review_ACDC")
        review_command = ensure_review_sheet(wb, "Review_Command")

        reviews: List[ReviewRow] = []
        filled_cells = 0
        for ws in list(wb.worksheets):
            if not ws.title.startswith("Cross Reference"):
                continue
            anchor_map = build_merged_anchor_map(ws)
            items = get_template_items(ws)
            headers = get_header_columns(ws)
            for col, (vendor, part) in headers.items():
                key = (ws.title, col)
                pinfo = mapping.get(key)
                if not pinfo:
                    continue
                acdc_result, cmd_result = extracted_cache.get(pinfo.source_file, ({}, {}))
                result_map = acdc_result if ws.title == "Cross Reference--ACDC" else cmd_result
                if clear_matched_cells:
                    for row, item in items.items():
                        if row <= 3 or is_section_item(item):
                            continue
                        clear_cell_safely(ws, row, col, anchor_map)
                for row, item in items.items():
                    if row <= 3 or is_section_item(item):
                        continue
                    extracted = result_map.get(item)
                    template_value = template_values.get((ws.title, row, col), "")
                    extracted_value = clean_text(extracted.value if extracted else "")
                    evidence = clean_text(extracted.evidence if extracted else "")
                    conf = float(extracted.confidence if extracted else 0.0)
                    source_file = short_filename(pinfo.source_file)
                    status, sim = similarity_status(template_value, extracted_value)
                    # Write only extracted values to target format sheet. Evidence goes to Review sheet.
                    if extracted_value:
                        set_cell_value_safely(ws, row, col, extracted_value, anchor_map)
                        filled_cells += 1
                    rr = ReviewRow(
                        sheet=ws.title,
                        vendor=vendor,
                        part_number=part,
                        item=item,
                        template_value=template_value,
                        extracted_value=extracted_value,
                        match_status=status,
                        similarity=round(sim, 3),
                        confidence=round(conf, 2),
                        source_file=source_file,
                        evidence=evidence,
                        row=row,
                        column=col,
                    )
                    reviews.append(rr)
                    target_review = review_acdc if ws.title == "Cross Reference--ACDC" else review_command
                    target_review.append([
                        rr.sheet, rr.vendor, rr.part_number, rr.item, rr.template_value, rr.extracted_value,
                        rr.match_status, rr.similarity, rr.confidence, rr.source_file, rr.evidence,
                        f"{get_column_letter(col)}{row}",
                    ])
                    style_review_row(target_review[target_review.max_row], rr.match_status)

        create_pdf_index_sheet(wb, pdf_infos, mapping)

        # Basic layout fix for original crossref sheets.
        for ws in wb.worksheets:
            if ws.title.startswith("Cross Reference"):
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is not None:
                            cell.alignment = Alignment(
                                horizontal=cell.alignment.horizontal,
                                vertical=cell.alignment.vertical or "center",
                                wrap_text=True,
                            )

        output_xlsx = Path(output_xlsx)
        output_xlsx.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(output_xlsx))
        elapsed = round(time.time() - start, 1)
        return {
            "output_xlsx": str(output_xlsx),
            "template_xlsx": str(template_xlsx),
            "pdf_count": len(pdf_paths),
            "matched_columns": len(mapping),
            "filled_cells": filled_cells,
            "review_rows": len(reviews),
            "elapsed_sec": elapsed,
        }
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


# -----------------------------
# 5. Streamlit UI
# -----------------------------

def run_streamlit_app():
    try:
        import streamlit as st
    except Exception as exc:  # pragma: no cover
        print("当前环境没有安装 Streamlit。请先执行：pip install streamlit，然后运行：streamlit run cross_reference_streamlit_app.py")
        raise exc

    st.set_page_config(page_title="XTX Cross Reference Extractor", layout="wide")
    st.title("XTX NOR Flash Cross Reference 自动提取工具")
    st.caption("上传 Cross Reference 模板和友商规格书 ZIP，自动生成同格式 Cross Reference，并附带 Review 表进行二次确认。")

    with st.sidebar:
        st.header("输入文件")
        template_file = st.file_uploader("Cross Reference 模板 Excel（.xlsx）", type=["xlsx"])
        specs_zip = st.file_uploader("友商规格书压缩包（.zip，内含 PDF）", type=["zip"])
        st.divider()
        st.header("提取设置")
        match_threshold = st.slider("PDF 与模板列匹配阈值", min_value=0.50, max_value=0.95, value=0.70, step=0.01)
        clear_matched_cells = st.checkbox("输出表中先清空匹配列旧值，再写入抽取值", value=True)
        max_pages_mode = st.selectbox("PDF 扫描范围", ["仅前 80 页", "仅前 40 页", "全文扫描"], index=0)
        max_pages = None if max_pages_mode == "全文扫描" else int(re.search(r"\d+", max_pages_mode).group(0))
        use_pdfplumber = st.checkbox("启用 pdfplumber 表格增强抽取（更准但更慢）", value=False)
        output_name = st.text_input("输出文件名", value="Cross_Reference_Extracted_Review.xlsx")
        st.divider()
        st.markdown("**规则优化建议**：第一次结果不用追求 100%，先看 Review 表中 Missing/Different，再把别名补充到 `crossref_rules.json`。")

    st.subheader("处理逻辑")
    st.markdown(
        "1. 复制模板，保留原有格式；  "
        "2. 根据 PDF 文件名/首页 Part Number 自动匹配模板列；  "
        "3. 从 PDF 表格/文本抽取 ACDC 与 Command；  "
        "4. 生成原格式 Cross Reference + Review_ACDC + Review_Command + PDF_Index。"
    )

    run_btn = st.button("开始生成", type="primary", disabled=(template_file is None or specs_zip is None))

    if run_btn:
        temp_root = Path(tempfile.mkdtemp(prefix="xtx_crossref_ui_"))
        try:
            template_path = temp_root / "template.xlsx"
            specs_path = temp_root / "specs.zip"
            out_path = temp_root / output_name
            rules_path = temp_root / "crossref_rules.json"
            template_path.write_bytes(template_file.getvalue())
            specs_path.write_bytes(specs_zip.getvalue())

            status_box = st.empty()
            progress_bar = st.progress(0)
            messages = []

            def progress(msg: str):
                messages.append(msg)
                # not exact but useful.
                progress_bar.progress(min(95, len(messages) * 3))
                status_box.info("\n".join(messages[-8:]))

            result = generate_cross_reference(
                template_xlsx=template_path,
                specs_input=specs_path,
                output_xlsx=out_path,
                rules_path=rules_path,
                match_threshold=match_threshold,
                clear_matched_cells=clear_matched_cells,
                max_pages=max_pages,
                use_pdfplumber=use_pdfplumber,
                progress_callback=progress,
            )
            progress_bar.progress(100)
            st.success("生成完成")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("PDF 数量", result["pdf_count"])
            c2.metric("匹配列数", result["matched_columns"])
            c3.metric("写入单元格", result["filled_cells"])
            c4.metric("Review 行数", result["review_rows"])
            st.download_button(
                "下载结果 Excel",
                data=out_path.read_bytes(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.download_button(
                "下载默认规则 JSON（用于后续优化）",
                data=rules_path.read_bytes(),
                file_name="crossref_rules.json",
                mime="application/json",
            )
        except Exception as exc:
            st.error(f"生成失败：{exc}")
            st.exception(exc)
        finally:
            shutil.rmtree(temp_root, ignore_errors=True)


# -----------------------------
# 6. CLI entrypoint
# -----------------------------

def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="XTX NOR Flash Cross Reference Extractor")
    parser.add_argument("--cli", action="store_true", help="使用命令行模式运行。")
    parser.add_argument("--template", type=str, help="Cross Reference 模板 xlsx 路径。")
    parser.add_argument("--specs", type=str, help="规格书 zip 或文件夹路径。")
    parser.add_argument("--out", type=str, default="Cross_Reference_Extracted_Review.xlsx", help="输出 xlsx 路径。")
    parser.add_argument("--rules", type=str, default="crossref_rules.json", help="规则 JSON 路径，不存在会自动生成。")
    parser.add_argument("--threshold", type=float, default=0.70, help="PDF 与模板列匹配阈值。")
    parser.add_argument("--keep-old", action="store_true", help="不清空旧值，只写入抽取到的新值。")
    parser.add_argument("--max-pages", type=int, default=0, help="最多扫描 PDF 页数；0 表示全文。")
    parser.add_argument("--use-pdfplumber", action="store_true", help="启用 pdfplumber 表格增强抽取，速度较慢。")
    args = parser.parse_args(argv)

    if not args.cli:
        run_streamlit_app()
        return 0

    if not args.template or not args.specs:
        parser.error("命令行模式必须提供 --template 和 --specs")
    result = generate_cross_reference(
        template_xlsx=Path(args.template),
        specs_input=Path(args.specs),
        output_xlsx=Path(args.out),
        rules_path=Path(args.rules),
        match_threshold=args.threshold,
        clear_matched_cells=not args.keep_old,
        max_pages=None if args.max_pages == 0 else args.max_pages,
        use_pdfplumber=args.use_pdfplumber,
        progress_callback=lambda msg: print(msg),
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
