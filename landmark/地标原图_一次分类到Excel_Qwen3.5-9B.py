# -*- coding: utf-8 -*-
"""
1969张地标原图：Qwen3.5-9B一次视觉分析 -> 单个Excel
兼容旧脚本使用的 KoboldCpp OpenAI 接口。
安装：py -m pip install Pillow
测试：py 地标原图_一次分类到Excel_Qwen3.5-9B.py --input "G:\地标\Original" --limit 20
正式：py 地标原图_一次分类到Excel_Qwen3.5-9B.py --input "G:\地标\Original"
不填 --input 时可直接交互输入文件夹。已有成功结果会自动跳过。
"""
from __future__ import annotations

import argparse
import base64
import io
import json
import os
import re
import sys
import time
import urllib.request
import zipfile
from collections import deque
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

try:
    from PIL import Image, ImageOps
except ImportError:
    raise SystemExit("请先安装 Pillow：py -m pip install Pillow")

# 与此前“人物图片风格多样筛选_无人脸识别.py”的接口一致
API_URL = "http://10.10.200.59:5002/v1/chat/completions"
MODEL_NAME = "koboldcpp"
EXPECTED_COUNT = 1969
EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
EXCLUDE_DIRS = {"mask", "masks", "ps", "contrast", ".git", "__pycache__"}
MODEL_FIELDS = [
    "详细描述", "地标身份特征", "地标名称", "地标类型", "场景", "季节",
    "天气", "时间段", "拍摄角度", "拍摄距离", "光照", "色温", "色调",
    "氛围", "图像风格", "地面状态", "构图", "天空", "植被",
    "周围环境", "地标可见程度",
]
HEADERS = ["文件名"] + MODEL_FIELDS + ["相对路径", "完整路径", "状态", "耗时秒", "错误信息"]
CHOICES = {
    "场景": "城市建筑群/城市街道/城市广场/公园园林/滨水河岸/海滨/山地/乡村/室内/其他/无法判断",
    "季节": "春季/夏季/秋季/冬季/无法判断",
    "天气": "晴朗/少云/多云/阴天/雨天/雪天/雾天/无法判断",
    "时间段": "日出/上午/正午/下午/日落/蓝调时刻/夜间/无法判断",
    "拍摄角度": "正面/左侧/右侧/背面/斜侧/仰拍/俯拍/航拍/无法判断",
    "拍摄距离": "全景/远景/中景/近景/特写/无法判断",
    "光照": "自然直射光/自然漫射光/侧光/逆光/人工灯光/混合光/无法判断",
    "色温": "暖色/中性/冷色/混合/无法判断",
    "色调": "明亮鲜艳/柔和低饱和/暗调/高对比/自然均衡/黑白/无法判断",
    "氛围": "宁静/热闹/庄严/浪漫/神秘/清新/阴郁/现代/历史感/其他/无法判断",
    "图像风格": "真实摄影/建筑摄影/风光摄影/艺术摄影/航拍摄影/插画/三维渲染/无法判断",
    "地面状态": "干燥/湿润/积水/积雪/结冰/未出现地面/无法判断",
    "构图": "主体居中/主体偏左/主体偏右/对称构图/引导线构图/其他/无法判断",
    "地标可见程度": "完整清晰/部分遮挡/严重遮挡/模糊/无法辨认",
}
CHOICE_SETS = {k: set(v.split("/")) for k, v in CHOICES.items()}

PROMPT = """你是地标图像数据集专业视觉标注员。针对当前唯一的一张原图，只依据真实可见画面，输出一个完整合法JSON对象。该结果将用于AI编辑生成后的地标身份一致性筛选和场景属性分析。

要求：
1. “详细描述”写100至180个汉字，客观描述地标可见形状、组成、空间位置、建筑颜色与材质、周边环境（建筑/道路/树木/水面/天空）、构图、天气和光影，不得套话或编造。
2. “地标身份特征”写40至100个汉字，提取2至5个稳定且有区分力的建筑特征，如独特轮廓、屋顶、塔冠、拱门、桥塔、球形结构及结构关系；不要写天气、背景、灯光或拍摄时间。用于以后判定生成图是否仍包含同一个地标。
3. “地标名称”仅在识别可靠时填写，否则写“无法确认”。文件名及上级目录只是可能的名称线索，不能代替对图像的观察。
4. 季节仅根据植被、雪、落叶等可见依据判断；天气依据天空或降水；时间依据真实自然光和天空条件。证据不足必须填“无法判断”，不能根据地理位置和后期色调猜测。
5. 区分地标的稳定结构与可变的季节、天气、时间、光照、背景、视角。图像遮挡或模糊需如实标注。天空或植被不在画面中则写“无”。
6. 下列分类字段每项必须从其候选值中精确选择一个：
""" + "\n".join(f"{k}：{v}" for k, v in CHOICES.items()) + """
7. 其他字段自由用简洁中文表述。“周围环境”写20至70字，“地标类型”可写桥梁/塔楼/雕塑/纪念碑/宫殿/历史建筑等。
8. 返回且只返回下面JSON的全部字段，禁止Markdown、解释性前后缀和多余字段：
""" + json.dumps({k: "" for k in MODEL_FIELDS}, ensure_ascii=False)


def timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def norm_path(value: str) -> Path:
    return Path(value.strip().strip('"').strip("'")).expanduser().resolve()


def scan_images(roots: list[Path], out_dir: Path) -> list[tuple[Path, Path]]:
    found = []
    visited = set()
    for root in roots:
        if not root.is_dir():
            raise FileNotFoundError(f"输入目录不存在：{root}")
        for current, dirs, names in os.walk(root):
            base = Path(current)
            dirs[:] = sorted((d for d in dirs
                              if d.casefold() not in EXCLUDE_DIRS
                              and (base / d).resolve() != out_dir.resolve()),
                             key=str.casefold)
            for name in sorted(names, key=str.casefold):
                p = base / name
                if p.suffix.lower() not in EXTS:
                    continue
                if re.search(r"(?:^|[_-])mask(?:$|[_-])", p.stem, re.I):
                    continue
                key = str(p.resolve()).casefold()
                if key not in visited:
                    visited.add(key)
                    found.append((p, root))
    found.sort(key=lambda pair: str(pair[0]).casefold())
    return found


def image_to_base64(path: Path, max_edge: int, quality: int) -> str:
    with Image.open(path) as src:
        image = ImageOps.exif_transpose(src)
        if image.mode in ("RGBA", "LA") or "transparency" in image.info:
            rgba = image.convert("RGBA")
            bg = Image.new("RGB", rgba.size, "white")
            bg.paste(rgba, mask=rgba.getchannel("A"))
            image = bg
        else:
            image = image.convert("RGB")
        image.thumbnail((max_edge, max_edge), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        image.save(buf, "JPEG", quality=quality)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def parse_model_output(content: object) -> dict[str, str]:
    if isinstance(content, list):
        content = "".join(x.get("text", "") for x in content
                          if isinstance(x, dict) and x.get("type") == "text")
    if not isinstance(content, str):
        raise ValueError("模型没有返回文本")
    content = re.sub(r"<think>.*?</think>", "", content, flags=re.I | re.S)
    decoder = json.JSONDecoder()
    data = None
    for start in re.finditer(r"\{", content):
        try:
            obj, _ = decoder.raw_decode(content[start.start():])
            if isinstance(obj, dict):
                data = obj
                break
        except json.JSONDecodeError:
            pass
    if data is None:
        raise ValueError("无有效JSON：" + content[:160].replace("\n", " "))
    result = {}
    for key in MODEL_FIELDS:
        if key not in data:
            raise ValueError("JSON缺少字段：" + key)
        value = data[key]
        if isinstance(value, list):
            value = "；".join(map(str, value))
        value = str(value if value is not None else "").strip() or "无法判断"
        if key in CHOICE_SETS and value not in CHOICE_SETS[key]:
            raise ValueError(f"非法分类值：{key}={value}")
        result[key] = value
    if len(result["详细描述"]) < 30:
        raise ValueError("详细描述过短")
    return result


def infer(path: Path, args: argparse.Namespace) -> dict[str, str]:
    prompt = ("当前图片文件名：" + path.name + "\n"
              "所在目录名：" + path.parent.name +
              "\n文件名仅供辅助识别，不可编造任何视觉内容。\n\n" + PROMPT)
    payload = {
        "model": args.model,
        "temperature": 0,
        "max_tokens": args.max_tokens,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": image_to_base64(path, args.max_edge, args.quality)}},
            ],
        }],
    }
    if args.disable_thinking:
        payload["chat_template_kwargs"] = {"enable_thinking": False}
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    last_error = None
    for attempt in range(1, args.retries + 1):
        try:
            request = urllib.request.Request(args.api_url, body,
                                             headers={"Content-Type": "application/json"},
                                             method="POST")
            with urllib.request.urlopen(request, timeout=args.timeout) as response:
                answer = json.loads(response.read().decode("utf-8"))
            return parse_model_output(answer["choices"][0]["message"]["content"])
        except Exception as ex:
            last_error = ex
            print(f"    重试 {attempt}/{args.retries}：{str(ex)[:180]}", flush=True)
            if attempt < args.retries:
                time.sleep(min(2 ** (attempt - 1), 8))
    raise RuntimeError(f"视觉分析失败：{last_error}")


def image_info(path: Path) -> dict:
    stat = path.stat()
    return {"key": str(path.resolve()).casefold(), "size": stat.st_size,
            "mtime_ns": stat.st_mtime_ns}


def load_records(path: Path) -> dict[str, dict]:
    result = {}
    if path.is_file():
        with path.open(encoding="utf-8-sig") as f:
            for line in f:
                try:
                    entry = json.loads(line)
                    if isinstance(entry, dict) and "key" in entry:
                        result[entry["key"]] = entry
                except json.JSONDecodeError:
                    continue
    return result


def append_record(path: Path, record: dict) -> None:
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")
        f.flush()
        os.fsync(f.fileno())


def cache_valid(old: dict | None, info: dict) -> bool:
    return bool(isinstance(old, dict)
                and old.get("size") == info["size"]
                and old.get("mtime_ns") == info["mtime_ns"]
                and isinstance(old.get("data"), dict)
                and all(key in old["data"] for key in MODEL_FIELDS))


def col_name(number: int) -> str:
    output = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        output = chr(65 + remainder) + output
    return output


def xml_value(value: object) -> str:
    s = str("" if value is None else value)
    s = "".join(c for c in s if c in "\t\n\r" or
                (ord(c) >= 32 and not 0xD800 <= ord(c) <= 0xDFFF
                 and ord(c) not in (0xFFFE, 0xFFFF)))
    return escape(s)


def export_xlsx(dest: Path, images: list[tuple[Path, Path]],
                successes: dict[str, dict], errors: dict[str, dict]) -> None:
    """仅使用stdlib生成Excel XLSX；可在Excel/WPS打开，带筛选/冻结/换行。"""
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="{ns}">',
        '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews>',
        '<sheetFormatPr defaultRowHeight="19"/>', '<cols>'
    ]
    count = len(HEADERS)
    for c in range(1, count + 1):
        width = 42 if c == 1 else 85 if c == 2 else 66 if c == 3 else 18
        if HEADERS[c-1] in {"周围环境", "天空", "植被"}:
            width = 35
        if HEADERS[c-1] in {"相对路径", "完整路径"}:
            width = 64
        if HEADERS[c-1] == "错误信息":
            width = 58
        parts.append(f'<col min="{c}" max="{c}" width="{width}" customWidth="1"/>')
    parts.append("</cols><sheetData>")

    for row_number in range(1, len(images) + 2):
        if row_number == 1:
            row = HEADERS
        else:
            path, root = images[row_number - 2]
            info = image_info(path)
            old = successes.get(info["key"])
            good = cache_valid(old, info)
            err = errors.get(info["key"])
            fields = old["data"] if good else {}
            row = [path.name] + [fields.get(k, "") for k in MODEL_FIELDS] + [
                str(path.relative_to(root)), str(path.resolve()),
                "成功" if good else "失败" if err else "待处理",
                str(old.get("seconds", "")) if good else "",
                str(err.get("error", "")) if err and not good else "",
            ]
        style = 1 if row_number == 1 else 2
        height = 30 if row_number == 1 else 65
        parts.append(f'<row r="{row_number}" ht="{height}" customHeight="1">')
        for c, value in enumerate(row, 1):
            parts.append(f'<c r="{col_name(c)}{row_number}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{xml_value(value)}</t></is></c>')
        parts.append("</row>")
    parts.append("</sheetData>")
    parts.append(f'<autoFilter ref="A1:{col_name(count)}{len(images)+1}"/>')
    parts.append("</worksheet>")
    sheet_xml = "".join(parts)

    style_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{ns}">'
        '<fonts count="2"><font><sz val="10"/><name val="Microsoft YaHei"/></font>'
        '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font></fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF274564"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" horizontal="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs>'
        '</styleSheet>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="地标原图分类" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )
    entries = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": rels,
        "xl/workbook.xml": workbook,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/styles.xml": style_xml,
        "xl/worksheets/sheet1.xml": sheet_xml,
    }
    temporary = dest.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for filename, text in entries.items():
            z.writestr(filename, text)
    try:
        os.replace(temporary, dest)
    except PermissionError:
        print(f"[提示] 目标xlsx正在被Excel占用，最新版本暂存：{temporary}", flush=True)


def execute(args: argparse.Namespace) -> None:
    if not args.input:
        typed = input("请输入地标Original原图目录路径：").strip()
        if not typed:
            raise SystemExit("未输入路径")
        args.input = [typed]
    roots = [norm_path(s) for s in args.input]
    out_dir = norm_path(args.output) if args.output else Path(__file__).resolve().parent / "地标原图分类_输出"
    out_dir.mkdir(parents=True, exist_ok=True)
    cache_file = out_dir / "分类断点缓存.jsonl"
    error_file = out_dir / "失败日志.jsonl"
    xlsx_file = out_dir / "地标原图_1969张_详细描述与属性分类.xlsx"
    images = scan_images(roots, out_dir)
    if args.limit > 0:
        images = images[:args.limit]
    if not images:
        raise SystemExit("未找到原图；请检查--input是否指向Original目录")
    print(f"共扫描 {len(images)} 张原图，预期 {EXPECTED_COUNT} 张。")
    if args.limit == 0 and len(images) != EXPECTED_COUNT:
        print("[提醒] 数量不是1969，请检查是否选择正确目录。")
    print(f"接口：{args.api_url}，模型：{args.model}")
    print(f"输出：{xlsx_file}")

    done = load_records(cache_file)
    errors = load_records(error_file)
    pending = []
    for path, root in images:
        info = image_info(path)
        if not args.force and cache_valid(done.get(info["key"]), info):
            continue
        pending.append((path, root, info))
    print(f"已成功缓存 {len(images)-len(pending)} 张，本轮待处理 {len(pending)} 张。")
    export_xlsx(xlsx_file, images, done, errors)
    if args.dry_run or not pending:
        print("未调用模型，Excel准备完毕。")
        return

    recent = deque(maxlen=30)
    success_count = fail_count = 0
    for i, (path, root, info) in enumerate(pending, 1):
        t = time.monotonic()
        print(f"[{i}/{len(pending)}] {path.name}", flush=True)
        try:
            result = infer(path, args)
            seconds = round(time.monotonic()-t, 2)
            item = {**info, "data": result, "seconds": seconds, "at": timestamp()}
            append_record(cache_file, item)
            done[info["key"]] = item
            errors.pop(info["key"], None)
            success_count += 1
            print(f"    OK | {seconds:.1f}秒 | {result['地标名称']} | {result['天气']} / {result['季节']}")
        except KeyboardInterrupt:
            print("中止，本次结果将写入Excel；下次运行自动续跑。")
            break
        except Exception as ex:
            seconds = round(time.monotonic()-t, 2)
            item = {**info, "error": str(ex), "at": timestamp()}
            append_record(error_file, item)
            errors[info["key"]] = item
            fail_count += 1
            print(f"    错误：{str(ex)[:220]}")
        recent.append(seconds)
        if i % args.save_every == 0 or i == len(pending):
            export_xlsx(xlsx_file, images, done, errors)
        average = sum(recent)/len(recent)
        eta = (len(pending)-i)*average/3600
        print(f"    本轮成功{success_count} 失败{fail_count} 最近均速{average:.1f}秒/张 预计剩余{eta:.1f}小时", flush=True)

    export_xlsx(xlsx_file, images, done, errors)
    valid = sum(cache_valid(done.get(image_info(p)["key"]), image_info(p)) for p, _ in images)
    print(f"\n完成：共{len(images)}张、成功{valid}张、待处理或失败{len(images)-valid}张。")
    print(f"Excel：{xlsx_file}")
    print("断点缓存保存在输出目录，下次同命令会跳过成功图片并重试失败图片。")


def main() -> None:
    parser = argparse.ArgumentParser(description="地标原图1969张一次性Qwen视觉分析到XLSX")
    parser.add_argument("--input", action="append", default=[], help="原图文件夹，可多次指定")
    parser.add_argument("--output", default="", help="结果目录，默认脚本同目录下")
    parser.add_argument("--api-url", default=API_URL)
    parser.add_argument("--model", default=MODEL_NAME)
    parser.add_argument("--max-edge", type=int, default=1024)
    parser.add_argument("--quality", type=int, default=88)
    parser.add_argument("--max-tokens", type=int, default=1200)
    parser.add_argument("--timeout", type=int, default=300)
    parser.add_argument("--retries", type=int, default=3)
    parser.add_argument("--save-every", type=int, default=25)
    parser.add_argument("--limit", type=int, default=0, help="仅前N张先测试，0为全部")
    parser.add_argument("--force", action="store_true", help="强制重新分析已有缓存")
    parser.add_argument("--dry-run", action="store_true", help="仅扫描图片并生成空表，不请求模型")
    parser.add_argument("--disable-thinking", action="store_true", help="后端支持时关闭思考模式")
    args = parser.parse_args()
    if args.max_edge < 256 or args.save_every < 1 or args.retries < 1:
        parser.error("max-edge需>=256，save-every及retries需>=1")
    execute(args)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("已终止。下次运行可继续。")
    except Exception as e:
        print(f"运行异常：{e}", file=sys.stderr)
        raise
