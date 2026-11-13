# -*- coding: utf-8 -*-
"""Flux/Qwen生成地标图与Original原图双图对照筛选，Qwen3.5-9B。
安装: py -m pip install Pillow
测试: py 地标生成图_对照原图筛选_Qwen3.5-9B.py --original "G:\Original" --flux "G:\Flux" --qwen "G:\Qwen" --limit 20
全量: 去掉--limit 20。只传--original则会交互询问Flux、Qwen目录。
结果: PASS/FAIL/REVIEW按来源复制；筛选总表.xlsx/.csv；断点缓存.jsonl。
原图匹配不到时直接REVIEW，不乱选原图，也不删除原始文件。
"""
from __future__ import annotations

import argparse
import base64
import csv
import io
import json
import os
import re
import shutil
import sys
import time
import urllib.request
import zipfile
from collections import defaultdict, deque
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

try:
    from PIL import Image, ImageOps
except ImportError:
    raise SystemExit("请安装 Pillow：py -m pip install Pillow")

API_URL = "http://10.10.200.59:5002/v1/chat/completions"
MODEL_NAME = "koboldcpp"
PROMPT_VERSION = "landmark-pair-v1"
EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
SKIP_DIRS = {"mask", "masks", "ps", "contrast", "__pycache__", ".git"}
FIELDS = [
    "模型来源", "生成图文件名", "原图文件名", "结果", "生成图识别地标", "原图地标名",
    "生成图主体可见", "结构是否一致", "判定理由", "匹配方法", "生成图相对路径",
    "生成图完整路径", "原图完整路径", "耗时秒", "处理时间", "错误信息"
]
PROMPT = """你是地标编辑生成图的视觉一致性审核员。当前消息内有两张图片：
【图1】真实地标参考原图；【图2】Flux或Qwen生成图。

目标：审核图2是否仍包含图1中的同一地标主体及关键建筑结构，不是要求两图像素相同。
分别观察两张图，并核对有辨识度的整体轮廓、塔冠/球体/桥塔/拱廊/雕塑/立面结构、各组成部分关系及基本比例。
允许天气、日夜、季节、光照、颜色、背景、植被、人物、拍摄角度、距离、构图发生变化。这些变化本身不构成不合格。
不得只因为生成图看上去像同一城市、或文字写了同一地标名称就判合格。

判定：
PASS：图2明确能识别图1所示目标地标，关键结构特征一致；局部遮挡或视角变化可以接受。
FAIL：图2明确没有目标地标、换成其他建筑、或关键结构已明显改变而不再是同一地标。
REVIEW：地标太小、严重遮挡或模糊、难以对齐；或证据不足不能确定。
不要把仅有局部小瑕疵但主体仍然相同的图误判为FAIL；不确定时选REVIEW。
只输出一个JSON对象，禁止Markdown和额外说明：
{"result":"PASS或FAIL或REVIEW","target_visible":"是或否或不确定","structural_match":"一致或不一致或无法判断","generated_landmark":"图2地标名或无法确认","reason":"简短中文理由，指出具体可见结构或缺失/冲突"}"""


def path_of(s: str) -> Path:
    return Path(s.strip().strip('"').strip("'")).expanduser().resolve()


def inside(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def scan(root: Path, output: Path) -> list[Path]:
    if not root.is_dir():
        raise FileNotFoundError("目录不存在：" + str(root))
    items = []
    for cur, dirs, names in os.walk(root, followlinks=False):
        base = Path(cur)
        dirs[:] = sorted((d for d in dirs if d.casefold() not in SKIP_DIRS
                          and not inside(base / d, output)), key=str.casefold)
        for n in sorted(names, key=str.casefold):
            p = base / n
            if p.suffix.lower() in EXTS and not re.search(r"(?:^|[_-])mask(?:$|[_-])", p.stem, re.I):
                items.append(p)
    return items


TOKENS = {"original", "real", "landmark", "fake", "flux", "qwen", "qwenimage",
          "edit", "edited", "image", "img", "generated", "result", "p2i"}


def filename_key(stem: str, drop_tail: bool = True) -> str:
    parts = [p for p in re.split(r"[_\s-]+", stem.casefold()) if p]
    if drop_tail and len(parts) > 1 and parts[-1].isdigit() and len(parts[-1]) < 5:
        parts.pop()
    return "_".join(p for p in parts if p not in TOKENS)


def serial(p: Path) -> str:
    match = re.match(r"^(\d{5})(?:[_-]|$)", p.stem)
    return match.group(1) if match else ""


def landmark_key(p: Path, root: Path) -> str:
    relative = p.parent.relative_to(root)
    if relative.parts:
        folder = relative.parts[-1]
        if folder.casefold() not in {"original", "flux", "qwen", "qwenimage", "images", "原始", "地标"}:
            return re.sub(r"[_\s-]+", "", folder.casefold())
    parts = [x for x in re.split(r"[_\s-]+", p.stem.casefold()) if x]
    if "landmark" in parts:
        parts = parts[parts.index("landmark") + 1:]
    elif parts and re.fullmatch(r"\d{5}", parts[0]):
        parts = parts[1:]
    parts = [x for x in parts if x not in TOKENS]
    while parts and parts[-1].isdigit():
        parts.pop()
    return "".join(parts)


class OriginalIndex:
    def __init__(self, originals: list[Path], root: Path):
        self.root = root
        self.by_name, self.by_number, self.by_landmark = defaultdict(list), defaultdict(list), defaultdict(list)
        for p in originals:
            self.by_name[filename_key(p.stem, False)].append(p)
            if serial(p):
                self.by_number[serial(p)].append(p)
            if landmark_key(p, root):
                self.by_landmark[landmark_key(p, root)].append(p)

    def match(self, img: Path, root: Path, group_mode: bool) -> tuple[Path | None, str, str]:
        matches = self.by_name.get(filename_key(img.stem), [])
        if len(matches) == 1:
            return matches[0], "规范化文件名", ""
        if len(matches) > 1:
            return None, "重名冲突", f"找到{len(matches)}张同名原图"
        number = serial(img)
        if number:
            matches = self.by_number.get(number, [])
            if len(matches) == 1:
                return matches[0], "五位原图编号", ""
            if len(matches) > 1:
                return None, "编号冲突", f"编号{number}对应多张原图"
        if group_mode:
            group = self.by_landmark.get(landmark_key(img, root), [])
            if group:
                # 同一地标组可用于“主体是否相同”审核，但不是相同拍摄原图。
                return sorted(group, key=lambda p: str(p).casefold())[0], "同地标组参照", f"组内{len(group)}张；非原图一对一对应"
        return None, "未匹配", "找不到唯一对应原图；如文件名不含原图编号可尝试--allow-group"


def load_reference_xlsx(path: Path) -> tuple[dict[str, dict], set[str]]:
    """可选：读取前一脚本的原图分类XLSX，只按唯一文件名关联，不按行号猜测。"""
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as z:
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            doc = ET.fromstring(z.read("xl/sharedStrings.xml"))
            strings = ["".join(n.text or "" for n in si.findall(".//x:t", ns))
                       for si in doc.findall("x:si", ns)]
        doc = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    table = []
    for row in doc.findall(".//x:sheetData/x:row", ns):
        values = {}
        for cell in row.findall("x:c", ns):
            a = re.match(r"[A-Z]+", cell.get("r", ""))
            if not a:
                continue
            ix = 0
            for ch in a.group(0):
                ix = ix * 26 + ord(ch) - 64
            if cell.get("t") == "inlineStr":
                value = "".join(n.text or "" for n in cell.findall(".//x:t", ns))
            else:
                v = cell.find("x:v", ns)
                value = v.text or "" if v is not None else ""
                if cell.get("t") == "s":
                    value = strings[int(value)]
            values[ix - 1] = value
        table.append(values)
    if not table:
        return {}, set()
    headers = {name: index for index, name in table[0].items()}
    if "文件名" not in headers:
        raise ValueError("参考XLSX缺少“文件名”列")
    data, duplicates = {}, set()
    for row in table[1:]:
        name = row.get(headers["文件名"], "").casefold()
        if not name:
            continue
        if name in data:
            duplicates.add(name)
        data[name] = {field: row.get(headers[field], "") for field in
                      ("地标名称", "地标身份特征") if field in headers}
    return data, duplicates


def to_image_uri(p: Path, edge: int, quality: int) -> str:
    with Image.open(p) as src:
        im = ImageOps.exif_transpose(src)
        if im.mode in ("RGBA", "LA") or "transparency" in im.info:
            rgba = im.convert("RGBA")
            bg = Image.new("RGB", rgba.size, "white")
            bg.paste(rgba, mask=rgba.getchannel("A"))
            im = bg
        else:
            im = im.convert("RGB")
        im.thumbnail((edge, edge), Image.Resampling.LANCZOS)
        b = io.BytesIO()
        im.save(b, "JPEG", quality=quality)
    return "data:image/jpeg;base64," + base64.b64encode(b.getvalue()).decode("ascii")


def parse_reply(raw: object) -> dict:
    if isinstance(raw, list):
        raw = "".join(str(x.get("text", "")) for x in raw if isinstance(x, dict) and x.get("type") == "text")
    if not isinstance(raw, str):
        raise ValueError("无可读文本")
    raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.I | re.S)
    decoder, result = json.JSONDecoder(), None
    for m in re.finditer(r"\{", raw):
        try:
            obj, _ = decoder.raw_decode(raw[m.start():])
            if isinstance(obj, dict):
                result = obj
                break
        except json.JSONDecodeError:
            pass
    if not result:
        raise ValueError("模型没有返回有效JSON：" + raw[:140].replace("\n", " "))
    decision = str(result.get("result", "")).upper().strip()
    visible = str(result.get("target_visible", "")).strip()
    structure = str(result.get("structural_match", "")).strip()
    if decision not in {"PASS", "FAIL", "REVIEW"} or visible not in {"是", "否", "不确定"} or structure not in {"一致", "不一致", "无法判断"}:
        raise ValueError("模型分类字段非法：" + str(result)[:160])
    if decision == "PASS" and (visible != "是" or structure != "一致"):
        decision = "REVIEW"
    if decision == "FAIL" and visible == "是" and structure == "一致":
        decision = "REVIEW"
    return {"decision": decision, "visible": visible, "structure": structure,
            "generated_landmark": str(result.get("generated_landmark", "无法确认"))[:100],
            "reason": str(result.get("reason", "未提供说明"))[:700]}


def judge(original: Path, generated: Path, reference: dict, args: argparse.Namespace) -> dict:
    extra = ""
    if reference:
        extra = ("\n原图分类表中的辅助信息（必须仍以两图可见内容为准）："
                 + reference.get("地标名称", "")[:60] + "；稳定身份特征："
                 + reference.get("地标身份特征", "")[:300])
    text = f"图1为原图：{original.name}；图2为生成图：{generated.name}。{extra}\n\n{PROMPT}"
    payload = {"model": args.model, "temperature": 0, "max_tokens": args.max_tokens,
               "messages": [{"role": "user", "content": [
                   {"type": "text", "text": text},
                   {"type": "image_url", "image_url": {"url": to_image_uri(original, args.max_edge, args.quality)}},
                   {"type": "image_url", "image_url": {"url": to_image_uri(generated, args.max_edge, args.quality)}}
               ]}]}
    if args.disable_thinking:
        payload["chat_template_kwargs"] = {"enable_thinking": False}
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    last = None
    for n in range(1, args.retries + 1):
        try:
            req = urllib.request.Request(args.api_url, body,
                                         headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(req, timeout=args.timeout) as res:
                obj = json.loads(res.read().decode("utf-8"))
            return parse_reply(obj["choices"][0]["message"]["content"])
        except Exception as exc:
            last = exc
            print(f"    第{n}/{args.retries}次失败：{str(exc)[:160]}", flush=True)
            if n < args.retries:
                time.sleep(min(2 ** (n - 1), 8))
    raise RuntimeError("Qwen双图对照失败：" + str(last))


def fingerprint(p: Path) -> dict:
    st = p.stat()
    return {"path": str(p.resolve()), "size": st.st_size, "mtime_ns": st.st_mtime_ns}


def key_for(model: str, p: Path) -> str:
    return model + "|" + str(p.resolve()).casefold()


def load_cache(p: Path) -> dict:
    records = {}
    if p.is_file():
        with p.open(encoding="utf-8-sig") as f:
            for line in f:
                try:
                    item = json.loads(line)
                    if isinstance(item, dict) and "key" in item:
                        records[item["key"]] = item
                except json.JSONDecodeError:
                    pass
    return records


def append_cache(p: Path, item: dict) -> None:
    with p.open("a", encoding="utf-8") as f:
        f.write(json.dumps(item, ensure_ascii=False) + "\n")
        f.flush()
        os.fsync(f.fileno())


def reusable(item: dict | None, image: Path, original: Path | None, xlsx: Path | None) -> bool:
    try:
        return bool(item and item.get("version") == PROMPT_VERSION
                    and item.get("generated") == fingerprint(image)
                    and item.get("original") == (fingerprint(original) if original else None)
                    and item.get("reference_xlsx") == (fingerprint(xlsx) if xlsx else None)
                    and item.get("decision") in {"PASS", "FAIL", "REVIEW"})
    except OSError:
        return False


def excel_col(n: int) -> str:
    name = ""
    while n:
        n, remainder = divmod(n - 1, 26)
        name = chr(65 + remainder) + name
    return name


def safe_cell(s: object) -> str:
    value = str("" if s is None else s)
    value = "".join(c for c in value if c in "\t\n\r" or
                    (ord(c) >= 32 and not 0xD800 <= ord(c) <= 0xDFFF and ord(c) not in (0xFFFE, 0xFFFF)))
    if value.startswith(("=", "+", "-", "@")):
        value = "'" + value
    return xml_escape(value)


def export_excel(path: Path, rows: list[dict]) -> None:
    """标准库写XLSX，不需要openpyxl或pandas，支持表头筛选与冻结。"""
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             f'<worksheet xmlns="{ns}"><sheetViews><sheetView workbookViewId="0">'
             '<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>'
             '</sheetView></sheetViews><cols>']
    for i, name in enumerate(FIELDS, 1):
        width = 65 if name in {"判定理由", "生成图完整路径", "原图完整路径"} else 44 if "文件名" in name or "路径" in name else 18
        parts.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
    parts.append("</cols><sheetData>")
    for ri, row in enumerate([dict(zip(FIELDS, FIELDS))] + rows, 1):
        parts.append(f'<row r="{ri}" ht="{28 if ri==1 else 48}" customHeight="1">')
        for ci, k in enumerate(FIELDS, 1):
            parts.append(f'<c r="{excel_col(ci)}{ri}" s="{1 if ri==1 else 2}" t="inlineStr"><is><t xml:space="preserve">{safe_cell(row.get(k, ""))}</t></is></c>')
        parts.append("</row>")
    parts.append("</sheetData>")
    parts.append(f'<autoFilter ref="A1:{excel_col(len(FIELDS))}{len(rows)+1}"/></worksheet>')
    style = (f'<styleSheet xmlns="{ns}"><fonts count="2">'
             '<font><sz val="10"/><name val="Microsoft YaHei"/></font>'
             '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/></font></fonts>'
             '<fills count="3"><fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             '<fill><patternFill patternType="solid"><fgColor rgb="FF264B6C"/></patternFill></fill></fills>'
             '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
             '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
             '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
             '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf>'
             '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>'
             '</cellXfs></styleSheet>')
    ct = ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
    relationships = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb = (f'<workbook xmlns="{ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<sheets><sheet name="地标筛选结果" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wb_rel = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
              '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    temp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as z:
        for filename, data in {"[Content_Types].xml": ct, "_rels/.rels": relationships,
                               "xl/workbook.xml": wb, "xl/_rels/workbook.xml.rels": wb_rel,
                               "xl/styles.xml": style, "xl/worksheets/sheet1.xml": "".join(parts)}.items():
            z.writestr(filename, data if data.startswith("<?xml") else
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + data)
    try:
        os.replace(temp, path)
    except PermissionError:
        print("[提示] XLSX被其他软件占用，最新表格临时文件：" + str(temp))


def reports(output: Path, tasks: list[tuple[str, Path, Path]], state: dict) -> None:
    rows = []
    for model, root, generated in tasks:
        record = state.get(key_for(model, generated), {})
        original = record.get("original") or {}
        rows.append({
            "模型来源": model, "生成图文件名": generated.name,
            "原图文件名": Path(original["path"]).name if original else "",
            "结果": record.get("decision", "待处理"),
            "生成图识别地标": record.get("generated_landmark", ""),
            "原图地标名": record.get("original_landmark", ""),
            "生成图主体可见": record.get("visible", ""),
            "结构是否一致": record.get("structure", ""),
            "判定理由": record.get("reason", ""),
            "匹配方法": record.get("match_method", ""),
            "生成图相对路径": str(generated.relative_to(root)),
            "生成图完整路径": str(generated.resolve()),
            "原图完整路径": original.get("path", ""),
            "耗时秒": record.get("seconds", ""),
            "处理时间": record.get("at", ""),
            "错误信息": record.get("error", "")
        })
    path = output / "地标生成图_筛选总表.csv"
    temp = path.with_suffix(".tmp.csv")
    with temp.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)
    try:
        os.replace(temp, path)
    except PermissionError:
        print("[提示] CSV被占用，最新文件：" + str(temp))
    export_excel(output / "地标生成图_筛选总表.xlsx", rows)


def copy_to_group(image: Path, source_root: Path, model: str, out: Path, decision: str) -> None:
    relative = image.relative_to(source_root)
    dst = out / decision / model / relative
    dst.parent.mkdir(parents=True, exist_ok=True)
    if not dst.exists() or dst.stat().st_size != image.stat().st_size or dst.stat().st_mtime_ns != image.stat().st_mtime_ns:
        shutil.copy2(image, dst)
    for other in ("PASS", "FAIL", "REVIEW"):
        if other != decision:
            old = out / other / model / relative
            if old.is_file():
                old.unlink()


def run() -> None:
    p = argparse.ArgumentParser(description="Qwen3.5-9B双图地标身份筛选：Flux+Qwen vs Original")
    p.add_argument("--original", help="地标原图Original目录")
    p.add_argument("--flux", action="append", default=[], help="Flux生成目录，可重复")
    p.add_argument("--qwen", action="append", default=[], help="Qwen生成目录，可重复")
    p.add_argument("--reference-xlsx", help="可选：前一脚本的1969张原图分类XLSX")
    p.add_argument("--output", help="输出目录")
    p.add_argument("--allow-group", action="store_true", help="找不到同编号原图时允许以同地标组首张原图作参考")
    p.add_argument("--api-url", default=API_URL)
    p.add_argument("--model", default=MODEL_NAME)
    p.add_argument("--max-edge", type=int, default=768, help="双图分别缩放到最长边768；可设1024")
    p.add_argument("--quality", type=int, default=88)
    p.add_argument("--max-tokens", type=int, default=450)
    p.add_argument("--timeout", type=int, default=300)
    p.add_argument("--retries", type=int, default=3)
    p.add_argument("--save-every", type=int, default=25)
    p.add_argument("--limit", type=int, default=0, help="先测试前N张，0为全量")
    p.add_argument("--force", action="store_true", help="忽略旧结果重新判断")
    p.add_argument("--dry-run", action="store_true", help="只预检原图匹配，不调用模型，也不分类复制")
    p.add_argument("--disable-thinking", action="store_true", help="仅后端支持时使用")
    args = p.parse_args()
    if not args.original:
        args.original = input("地标Original原图目录：").strip().strip('"')
    if not args.flux and not args.qwen:
        a = input("Flux生成图目录（可留空）：").strip().strip('"')
        b = input("Qwen生成图目录（可留空）：").strip().strip('"')
        if a:
            args.flux.append(a)
        if b:
            args.qwen.append(b)
    if not args.flux and not args.qwen:
        p.error("至少指定一个--flux或--qwen目录")
    if args.max_edge < 256 or args.retries < 1 or args.save_every < 1:
        p.error("max-edge应>=256，retries/save-every应>=1")
    original_dir = path_of(args.original)
    output = path_of(args.output) if args.output else Path(__file__).resolve().parent / "地标筛选_输出"
    output.mkdir(parents=True, exist_ok=True)
    originals = scan(original_dir, output)
    if not originals:
        raise RuntimeError("Original目录未找到图片")
    index = OriginalIndex(originals, original_dir)
    tasks, seen = [], set()
    for model, folders in (("Flux", args.flux), ("Qwen", args.qwen)):
        for folder in folders:
            root = path_of(folder)
            for img in scan(root, output):
                k = key_for(model, img)
                if k not in seen:
                    tasks.append((model, root, img))
                    seen.add(k)
    tasks.sort(key=lambda t: (t[0], str(t[2]).casefold()))
    if args.limit > 0:
        tasks = tasks[:args.limit]
    if not tasks:
        raise RuntimeError("未发现Flux/Qwen生成图片")
    xlsx = path_of(args.reference_xlsx) if args.reference_xlsx else None
    ref_data, repeated = load_reference_xlsx(xlsx) if xlsx else ({}, set())
    checkpoint = output / "筛选断点缓存.jsonl"
    state = load_cache(checkpoint)
    print(f"原图{len(originals)}张，Flux={sum(x[0]=='Flux' for x in tasks)}张，Qwen={sum(x[0]=='Qwen' for x in tasks)}张")
    print("模型：" + args.api_url + "；输出目录：" + str(output))
    durations = deque(maxlen=40)
    for i, (model, root, img) in enumerate(tasks, 1):
        k = key_for(model, img)
        original, method, issue = index.match(img, root, args.allow_group)
        cached = state.get(k)
        if not args.force and reusable(cached, img, original, xlsx):
            print(f"[{i}/{len(tasks)}] 缓存 {model} {img.name} => {cached['decision']}")
        else:
            start = time.monotonic()
            info = ref_data.get(original.name.casefold(), {}) if original and original.name.casefold() not in repeated else {}
            try:
                if args.dry_run:
                    decision = "REVIEW"
                    answer = {"visible": "", "structure": "", "generated_landmark": "",
                              "reason": "路径匹配预检；未执行视觉判断"}
                elif original is None:
                    decision = "REVIEW"
                    answer = {"visible": "不确定", "structure": "无法判断",
                              "generated_landmark": "无法确认", "reason": issue}
                else:
                    answer = judge(original, img, info, args)
                    decision = answer["decision"]
                elapsed = round(time.monotonic() - start, 2)
                record = {
                    "key": k, "version": PROMPT_VERSION, "generated": fingerprint(img),
                    "original": fingerprint(original) if original else None,
                    "reference_xlsx": fingerprint(xlsx) if xlsx else None,
                    "decision": decision, "visible": answer["visible"], "structure": answer["structure"],
                    "generated_landmark": answer["generated_landmark"],
                    "original_landmark": info.get("地标名称", ""),
                    "reason": answer["reason"], "match_method": method,
                    "seconds": elapsed, "at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "error": ""
                }
                if not args.dry_run:
                    append_cache(checkpoint, record)
                state[k] = record
                durations.append(elapsed)
                print(f"[{i}/{len(tasks)}] {model} {decision} {img.name} | {method} | {elapsed:.1f}秒", flush=True)
            except KeyboardInterrupt:
                print("中断，保存已有结果。")
                break
            except Exception as err:
                elapsed = round(time.monotonic() - start, 2)
                state[k] = {
                    "key": k, "generated": fingerprint(img),
                    "original": fingerprint(original) if original else None,
                    "decision": "ERROR", "match_method": method,
                    "error": str(err), "seconds": elapsed, "at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                durations.append(elapsed)
                print(f"[{i}/{len(tasks)}] ERROR {img.name}：{err}", flush=True)
        record = state.get(k, {})
        if not args.dry_run and record.get("decision") in {"PASS", "FAIL", "REVIEW"}:
            copy_to_group(img, root, model, output, record["decision"])
        if i % args.save_every == 0 or i == len(tasks):
            reports(output, tasks, state)
        if durations:
            average = sum(durations) / len(durations)
            print(f"  最近均速{average:.1f}秒/张，剩余约{(len(tasks)-i)*average/3600:.1f}小时", flush=True)
    reports(output, tasks, state)
    totals = defaultdict(int)
    for model, root, img in tasks:
        totals[state.get(key_for(model, img), {}).get("decision", "待处理")] += 1
    print("分类汇总：" + "，".join(f"{k}={v}" for k, v in sorted(totals.items())))
    print("结果总表：" + str(output / "地标生成图_筛选总表.xlsx"))


if __name__ == "__main__":
    try:
        run()
    except KeyboardInterrupt:
        print("已停止；下次运行可从断点继续。")
    except Exception as exc:
        print("错误：" + str(exc), file=sys.stderr)
        raise
