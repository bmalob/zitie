#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
字帖生成器 —— 把任意中文内容生成 A4 可打印 Word 字帖（米字格 / 楷体）。

快速开始:
  1. 把要练的内容放进 content.txt（换内容只改这个文件）
  2. python3 zitie.py
  生成 字帖.docx，直接打印即可。

常用参数:
  --title "早发白帝城·李白"   左侧竖排标题
  --cell 20                 格子大小 mm（默认 15，数字越大格子越大）
  --rows 7 / --cols 6       直接指定每页行数/每行字数（格子大小自动算）
  --order horizontal        现代横排（默认 vertical 传统竖排，从右往左）
  --blanks 2                每个字后面留 2 个空格（临写练习）
  --repeat 3                每个字连续写 3 遍
  --trace                   浅灰色字，适合描红
  --grid-style tian         tian 田字格 / mizi 米字格(默认) / box 方框
  --font "楷体"            换字体（Windows 用楷体，Mac 默认华文楷体）
完整参数见 python3 zitie.py -h
"""

import argparse
import glob
import math
import os
import re
import shutil
import subprocess
import sys
import uuid
import zipfile

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Mm
from docx.enum.section import WD_ORIENT

PAPERS = {"A4": (210, 297), "A5": (148, 210), "B5": (176, 250)}

MM_TO_TWIP = 1440.0 / 25.4
MM_TO_EMU = 36000.0

HERE = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(HERE, "fonts")
if os.name == "nt":        # Windows：用户字体目录
    USER_FONT_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                                 "Microsoft", "Windows", "Fonts")
elif sys.platform == "darwin":
    USER_FONT_DIR = os.path.expanduser("~/Library/Fonts")
else:                      # Linux
    USER_FONT_DIR = os.path.expanduser("~/.local/share/fonts")


def load_env():
    """读取 .env 配置（可用环境变量 ZITIE_ENV 指定路径；否则找当前目录再找脚本目录）。
    返回 {ZITIE_XXX: 值} 字典，不依赖第三方库。"""
    candidates = []
    if os.environ.get("ZITIE_ENV"):
        candidates.append(os.path.expanduser(os.environ["ZITIE_ENV"]))
    candidates.append(os.path.join(os.getcwd(), ".env"))
    candidates.append(os.path.join(HERE, ".env"))
    cfg = {}
    for path in candidates:
        if os.path.isfile(path):
            with open(path, encoding="utf-8") as fh:
                for line in fh:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, _, val = line.partition("=")
                    cfg[key.strip()] = val.strip().strip('"').strip("'")
            break
    return cfg


def env_bool(cfg, key, default=False):
    v = cfg.get(key)
    return default if v is None else v.lower() in ("1", "true", "yes", "y", "on", "是")


def env_float(cfg, key, default):
    try:
        return float(cfg[key])
    except (KeyError, ValueError):
        return default


def find_soffice():
    """定位 LibreOffice：优先 .env 的 ZITIE_SOFFICE，再 PATH，再常见安装路径。"""
    p = load_env().get("ZITIE_SOFFICE")
    if p:
        p = os.path.expanduser(p)
        if os.path.isfile(p):
            return p
    found = shutil.which("soffice")
    if found:
        return found
    for cand in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
    ):
        if os.path.isfile(cand):
            return cand
    return None

# 内容处理用到的正则 / 标点
CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
SPLIT_RE = re.compile(r"[，。！？；、…—\n\r,.;:!?]")
PUNCT = "，。、；：？！“”‘’「」『』（）《》〈〉…—·～,.;:?!()~"

# 系统自带字体预设（无需文件）
SYSTEM_FONT_PRESETS = {
    "stkaiti": "STKaiti",   # 华文楷体（macOS 自带，默认）
    "kaiti": "楷体",         # Windows / macOS 楷体
}

# 随附开源字体（用其真实字体名，不重命名）
BUNDLED_NATIVE = {
    "wenkai": ("LXGW WenKai", "LXGWWenKai-Regular.ttf"),  # 霞鹜文楷（免费可商用）
}


def font_family_from_file(path):
    """读取 .ttf/.otf/.ttc 字体文件的字体族名。"""
    from fontTools.ttLib import TTFont, TTCollection
    if path.lower().endswith(".ttc"):
        font = TTCollection(path).fonts[0]
    else:
        font = TTFont(path)
    name = font["name"]
    n = (name.getName(16, 3, 1, 0x409) or name.getName(1, 3, 1, 0x409)
         or name.getName(1, 1, 0, 0))
    return n.toUnicode()


def friendly_font_name(filename):
    """从文件名得到友好的中文字体名，如 田英章硬笔楷书简体_mianfeiziti.com.ttf -> 田英章硬笔楷书简体。"""
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.split(r"[_\-]", name)[0]          # 去掉 _mianfeiziti.com、-Regular 等后缀
    name = re.sub(r"^[0-9]+", "", name)          # 去掉开头编号，如 088
    return name.strip()


def list_bundled_fonts():
    """扫描 fonts/ 目录，返回 [{key, name, file, path}]（开源原生字体 + 书家字体）。"""
    out = []
    for key, (fam, fn) in BUNDLED_NATIVE.items():
        p = os.path.join(FONTS_DIR, fn)
        if os.path.isfile(p):
            out.append({"key": key, "name": fam, "file": fn, "path": p, "native": True})
    if os.path.isdir(FONTS_DIR):
        for p in sorted(glob.glob(os.path.join(FONTS_DIR, "*"))):
            if p.lower().endswith((".ttf", ".otf", ".ttc")):
                fn = os.path.basename(p)
                if any(fn == v[1] for v in BUNDLED_NATIVE.values()):
                    continue
                out.append({"key": "", "name": friendly_font_name(fn),
                            "file": fn, "path": p, "native": False})
    return out


def install_font_renamed(src_path, family_name):
    """把字体改名安装为唯一字体族名，返回（安装路径, 字体族名）。
    用 ASCII 族名（中文书名字体在 Word for Mac 上按中文名匹配不稳，
    英文族名可被 Word/WPS/LibreOffice 稳定匹配），字形仍是中文。
    同时避免不同字体内部名相同/怪异（如 Lw42.0、TYZyingbikai）导致选错字体。"""
    import hashlib
    from fontTools.ttLib import TTFont
    os.makedirs(USER_FONT_DIR, exist_ok=True)
    safe = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]", "", family_name)
    family_ascii = "ZT" + hashlib.md5(family_name.encode("utf-8")).hexdigest()[:8].upper()
    dst = os.path.join(USER_FONT_DIR, f"zt_{family_ascii}.ttf")
    if not os.path.isfile(dst):
        if src_path.lower().endswith(".ttc"):
            dst = dst[:-4] + ".ttc"
            shutil.copy(src_path, dst)
            return dst, family_ascii
        font = TTFont(src_path)
        nt = font["name"]
        for plat, enc, lang in ((3, 1, 0x409), (3, 1, 0x804)):
            nt.setName(family_ascii, 1, plat, enc, lang)    # 字体族名（ASCII，稳定匹配）
            nt.setName(family_ascii, 16, plat, enc, lang)   # 排版族名
            nt.setName(f"{family_ascii} Regular", 4, plat, enc, lang)  # 全名
            nt.setName("Regular", 2, plat, enc, lang)       # 子族：常规
            nt.setName("Regular", 17, plat, enc, lang)      # 排版子族
            nt.setName(family_ascii, 6, plat, enc, lang)    # PostScript 名
        try:
            font["OS/2"].fsType = 0   # 工作副本统一设为可嵌入，避免受限(fsType=2)字体被 Word 拒载
        except Exception:
            pass
        font.save(dst)
    return dst, family_ascii


def install_native(src_path):
    """直接复制安装（不改名），返回目标路径。"""
    os.makedirs(USER_FONT_DIR, exist_ok=True)
    dst = os.path.join(USER_FONT_DIR, os.path.basename(src_path))
    if not os.path.isfile(dst):
        shutil.copy(src_path, dst)
    return dst


def resolve_font(font_arg, font_file_arg):
    """解析（字体族名, 字体文件路径）。
    优先级：--font-file > 系统预设(stkaiti/kaiti) > 开源预设(wenkai) >
    fonts/ 目录书家字体（按名字模糊匹配）> 已安装的系统字体族名。"""
    # 1) 显式字体文件
    if font_file_arg:
        path = os.path.expanduser(font_file_arg)
        if not os.path.isfile(path):
            sys.exit(f"找不到字体文件：{path}")
        family = friendly_font_name(path)
        dst, family_ascii = install_font_renamed(path, family)
        print(f"✍️  使用字体文件：{os.path.basename(path)}（{family}）")
        return family_ascii, dst

    # 2) 系统预设
    if font_arg in SYSTEM_FONT_PRESETS:
        family = SYSTEM_FONT_PRESETS[font_arg]
        return family, find_font_file(family)

    # 3) 开源原生预设
    if font_arg in BUNDLED_NATIVE:
        family, fn = BUNDLED_NATIVE[font_arg]
        p = os.path.join(FONTS_DIR, fn)
        if os.path.isfile(p):
            dst = install_native(p)
            return family, dst

    # 4) fonts/ 目录书家字体（精确名 > 模糊包含）
    bundled = list_bundled_fonts()
    target = font_arg.lower().replace(" ", "")
    exact = [b for b in bundled if b["name"].lower().replace(" ", "") == target]
    if exact:
        b = exact[0]
    else:
        fuzzy = [b for b in bundled
                 if target in b["name"].lower().replace(" ", "")
                 or target in b["file"].lower()]
        b = fuzzy[0] if fuzzy else None
    if b:
        if b.get("native"):
            dst = install_native(b["path"])
            family_out = b["name"]
        else:
            dst, family_out = install_font_renamed(b["path"], b["name"])
        print(f"✍️  使用书家字体：{b['name']}")
        return family_out, dst

    # 5) 当作已安装的系统字体族名
    return font_arg, find_font_file(font_arg)


def print_available_fonts():
    print("可用字体：")
    print("  系统自带：")
    for key, fam in SYSTEM_FONT_PRESETS.items():
        print(f"    --font {key:<10} （{fam}）")
    print("  fonts/ 目录（书家字体，直接用名字或关键词选择）：")
    for b in list_bundled_fonts():
        tag = b["key"] or b["name"]
        print(f"    --font {tag}")


def find_font_file(family):
    """按字体族名在常见字体目录里找字体文件（找不到返回 None）。"""
    dirs = [os.path.expanduser("~/Library/Fonts"), "/Library/Fonts",
            "/System/Library/Fonts", "/System/Library/Fonts/Supplemental"]
    target = family.lower().replace(" ", "")
    for d in dirs:
        for ext in ("ttf", "otf", "ttc"):
            for p in glob.glob(os.path.join(d, f"*.{ext}")):
                try:
                    fam = font_family_from_file(p).lower().replace(" ", "")
                    if target in fam or fam in target:
                        return p
                except Exception:
                    continue
    for p in glob.glob("/System/Library/AssetsV2/**/AssetData/Kai*.ttc",
                       recursive=True):
        return p
    return None


def _obfuscate_ttf(data, guid):
    """OOXML 内嵌字体混淆（Word/WPS 标准）。

    密钥 = 把 GUID（去大括号/连字符后的 16 字节）整体倒序，
    再用它对字体文件前 32 字节循环 XOR。已用 Word 自己导出的
    .odttf + w:fontKey 反推校验通过。
    """
    h = guid.replace("-", "").replace("{", "").replace("}", "")
    key = bytes.fromhex(h)[::-1]
    out = bytearray(data)
    for i in range(min(32, len(out))):
        out[i] ^= key[i % 16]
    return bytes(out)


def _xml_root_open_tag(xml, root_local):
    """返回根元素开标签匹配（<w:root ...> 或 <w:root .../>）。"""
    pat = re.compile(r"<[\w:]*" + re.escape(root_local) + r"\b[^>]*?(/?)>",
                     re.DOTALL)
    m = pat.search(xml)
    if not m:
        # 去掉 XML 声明再找
        m = pat.search(xml[xml.find("?>") + 2:] if "?>" in xml else xml)
    return m


def _insert_root_child(xml, root_local, child_xml, position="first"):
    """在根元素内插入子节点，兼容自闭合 <root .../> 与成对 <root></root>。"""
    m = _xml_root_open_tag(xml, root_local)
    if not m:
        raise ValueError(f"找不到根元素 <{root_local}>")
    tag = m.group(0)
    prefix = tag[:tag.rindex(root_local)]  # 如 "<w:"
    close = f"</{prefix[1:]}{root_local}>"
    if tag.endswith("/>"):
        new_tag = tag[:-2].rstrip() + ">" + child_xml + close
        return xml[:m.start()] + new_tag + xml[m.end():]
    if position == "first":
        return xml[:m.end()] + child_xml + xml[m.end():]
    # last：插到闭合标签前
    close_idx = xml.rfind(close)
    return xml[:close_idx] + child_xml + xml[close_idx:]


def _ensure_root_attr(xml, root_local, attr, value):
    """确保根元素开标签上有某属性（如 xmlns:r）。"""
    m = _xml_root_open_tag(xml, root_local)
    tag = m.group(0)
    if attr in tag:
        return xml
    new_tag = tag[:-2] + f' {attr}="{value}"' + ("/>" if tag.endswith("/>") else ">")
    return xml[:m.start()] + new_tag + xml[m.end():]


def _insert_settings_embed(settings_xml):
    """把 <w:embedTrueTypeFonts/> 插到 <w:settings> 内 schema 正确的位置
    （在 writeProtection/view/zoom 等之后，proofState/defaultTabStop 等之前）。"""
    tag = "<w:embedTrueTypeFonts/>"
    # 这些元素在 schema 中都排在 embedTrueTypeFonts 之后，插到最先出现的那个之前
    later = ["proofState", "formsDesign", "attachedTemplate", "linkStyles",
             "stylePaneFormatFilter", "documentType", "mailMerge",
             "revisionView", "trackChanges", "documentProtection",
             "autoFormatOverride", "defaultTabStop", "characterSpacingControl",
             "compat", "rsids"]
    best = None
    for name in later:
        m = re.search(r"<w:" + name + r"\b", settings_xml)
        if m and (best is None or m.start() < best.start()):
            best = m
    if best:
        return settings_xml[:best.start()] + tag + settings_xml[best.start():]
    return _insert_root_child(settings_xml, "settings", tag, position="last")


def _embedded_font_entry_xml(family_name, ttf_path, rid, font_key):
    """构造带完整元数据的 <w:font> 条目。
    Word 需要 panose/charset/sig 等元数据判断内嵌字体覆盖中文，否则会回退。"""
    inner = ""
    try:
        from fontTools.ttLib import TTFont
        f = TTFont(ttf_path)
        os2 = f["OS/2"]
        p = os2.panose
        pan = bytes([p.bFamilyType, p.bSerifStyle, p.bWeight, p.bProportion,
                     p.bContrast, p.bStrokeVariation, p.bArmStyle,
                     p.bLetterForm, p.bMidline, p.bXHeight])
        cp1 = os2.ulCodePageRange1 or 0
        cp2 = os2.ulCodePageRange2 or 0
        if cp1 & 0x40000:        # bit18 GB2312 简体
            charset = "86"
        elif cp1 & 0x80000:      # bit19 Hangul
            charset = "81"
        elif cp1 & 0x20000:      # bit17 ShiftJIS
            charset = "80"
        else:
            charset = "86"       # 本工具字体均按中文处理
        sig = (f"<w:sig w:usb0='{os2.ulUnicodeRange1 or 0:08X}' "
               f"w:usb1='{os2.ulUnicodeRange2 or 0:08X}' "
               f"w:usb2='{os2.ulUnicodeRange3 or 0:08X}' "
               f"w:usb3='{os2.ulUnicodeRange4 or 0:08X}' "
               f"w:csb0='{cp1:08X}' w:csb1='{cp2:08X}'/>")
        inner = (f"<w:panose1 w:val='{pan.hex().upper()}'/>"
                 f"<w:charset w:val='{charset}'/>"
                 f"<w:family w:val='auto'/>"
                 f"<w:pitch w:val='variable'/>{sig}")
    except Exception:
        pass
    return (f'<w:font w:name="{family_name}">{inner}'
            f'<w:embedRegular r:id="{rid}" w:fontKey="{font_key}"/></w:font>')


def embed_font_in_docx(docx_path, ttf_path, family_name):
    """把 TTF 内嵌进 docx（Word/WPS 打开即使用，无需系统安装该字体）。

    关键点：
      - word/settings.xml 里加 <w:embedTrueTypeFonts/>（必须在 <w:settings> 根元素内部）；
      - word/fontTable.xml 里登记 <w:font w:name=...><w:embedRegular r:id=GUID/>；
      - word/_rels/fontTable.xml.rels 里加关系，Id=GUID，Target=fonts/embed1.odttf；
      - GUID 同时作为混淆密钥（Word 标准 XOR 前 32 字节），关系 Id 即该 GUID。
    """
    guid = str(uuid.uuid4()).upper()
    with open(ttf_path, "rb") as f:
        obf = _obfuscate_ttf(f.read(), guid)

    tmp = docx_path + ".embed"
    with zipfile.ZipFile(docx_path, "r") as zin:
        names = set(zin.namelist())
        ct = zin.read("[Content_Types].xml").decode("utf-8")
        settings = zin.read("word/settings.xml").decode("utf-8")
        if "word/fontTable.xml" in names:
            ftable = zin.read("word/fontTable.xml").decode("utf-8")
        else:
            ftable = ('<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')
        rels_name = "word/_rels/fontTable.xml.rels"
        rels = (zin.read(rels_name).decode("utf-8") if rels_name in names else
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')

    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    if "Extension=\"odttf\"" not in ct:
        ct = ct.replace("</Types>",
                        '<Default Extension="odttf" '
                        'ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/></Types>')

    if "embedTrueTypeFonts" not in settings:
        settings = _insert_settings_embed(settings)

    # 关系 Id 用普通 rId（fontTable 的 rels 独立编号，取已有最大号 +1）
    existing = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    rid = "rId" + str(max(existing) + 1 if existing else 1)
    font_key = "{" + guid + "}"   # Word 用 w:fontKey 携带混淆密钥 GUID（带大括号）

    ftable = _ensure_root_attr(ftable, "fonts", "xmlns:r", R_NS)
    if f'w:name="{family_name}"' not in ftable:
        entry = _embedded_font_entry_xml(family_name, ttf_path, rid, font_key)
        ftable = _insert_root_child(ftable, "fonts", entry, position="last")

    rel = (f'<Relationship Id="{rid}" '
           f'Type="{R_NS}/font" '
           'Target="fonts/embed1.odttf"/>')
    if f'Id="{rid}"' not in rels:
        rels = _insert_root_child(rels, "Relationships", rel, position="last")

    with zipfile.ZipFile(docx_path, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = ct.encode("utf-8")
            elif item.filename == "word/settings.xml":
                data = settings.encode("utf-8")
            elif item.filename == "word/fontTable.xml":
                data = ftable.encode("utf-8")
            elif item.filename == rels_name:
                data = rels.encode("utf-8")
            zout.writestr(item, data)
        if "word/fontTable.xml" not in names:
            zout.writestr("word/fontTable.xml", ftable.encode("utf-8"))
        if rels_name not in names:
            zout.writestr(rels_name, rels.encode("utf-8"))
        zout.writestr("word/fonts/embed1.odttf", obf)
    os.replace(tmp, docx_path)


def _export_pdf_word(docx_abs, pdf_abs):
    """用 Microsoft Word 导出 PDF（macOS）。Word 对内嵌字体/横版渲染最准确，
    且横版不会出现 LibreOffice 的个别长线渲染问题。成功返回 True。"""
    if sys.platform != "darwin" or not os.path.isdir("/Applications/Microsoft Word.app"):
        return False
    # 关键：按“文件名”精确定位我们刚打开的文档，绝不抓“活动文档”，
    # 以免误导出/关闭用户在 Word 里已打开的其它文档。
    tname = os.path.basename(docx_abs)
    script = (
        'with timeout of 200 seconds\n'
        '  tell application "Microsoft Word"\n'
        '    activate\n'
        f'    open POSIX file "{docx_abs}"\n'
        f'    set tname to "{tname}"\n'
        '    set d to missing value\n'
        '    repeat 90 times\n'
        '      try\n'
        '        set nm to name of every document\n'
        '        if nm contains tname then\n'
        '          repeat with doc in documents\n'
        '            if name of doc is tname then set d to doc\n'
        '          end repeat\n'
        '        end if\n'
        '      end try\n'
        '      if d is not missing value then exit repeat\n'
        '      delay 1\n'
        '    end repeat\n'
        '    if d is not missing value then\n'
        '      delay 2\n'
        f'      save as d file format format PDF file name "{pdf_abs}"\n'
        '      delay 1\n'
        '      close d saving no\n'
        '    end if\n'
        '  end tell\n'
        'end timeout\n'
    )
    try:
        r = subprocess.run(["osascript", "-e", script], check=False,
                           capture_output=True, timeout=200, text=True)
        if r.returncode != 0:
            print("   （Word 导出提示：" + (r.stderr or r.stdout or "").strip()[:120] + "）")
    except Exception as e:
        print(f"   （Word 导出超时/失败，改用 LibreOffice：{e}）")
        return False
    return os.path.isfile(pdf_abs)


def export_pdf(docx_path, outdir, font_path=None):
    """导出可直接打印的 PDF。优先用 Microsoft Word（内嵌字体/横版渲染最准），
    没有 Word 再退回 LibreOffice（会把所需字体置入其 profile，避免回退）。"""
    docx_abs = os.path.abspath(docx_path)
    pdf_abs = os.path.join(os.path.abspath(outdir),
                           os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if os.path.isfile(pdf_abs):
        os.remove(pdf_abs)

    if _export_pdf_word(docx_abs, pdf_abs):
        print(f"🧾 已用 Microsoft Word 导出可直接打印的 PDF：{pdf_abs}")
        return

    soffice = find_soffice()
    if not soffice:
        print("⚠️  未找到 Word/soffice，跳过 PDF 导出（docx 已生成）。")
        return
    profile = os.path.join(os.path.expanduser("~"), ".cache", "zitie-lo-profile")
    fonts_dir = os.path.join(profile, "user", "fonts")
    os.makedirs(fonts_dir, exist_ok=True)
    for p in [find_font_file("STKaiti"), font_path]:
        if p and os.path.isfile(p):
            dst = os.path.join(fonts_dir, os.path.basename(p))
            if not os.path.isfile(dst):
                shutil.copy(p, dst)
    subprocess.run([soffice, f"-env:UserInstallation=file://{profile}",
                    "--headless", "--convert-to", "pdf", "--outdir", outdir,
                    docx_path],
                   check=False, capture_output=True)
    if os.path.isfile(pdf_abs):
        print(f"🧾 已用 LibreOffice 导出可直接打印的 PDF：{pdf_abs}")
    else:
        print("⚠️  PDF 导出失败，但 docx 已正常生成，可直接用 Word 打开打印。")


def mm_twip(mm):
    return int(round(mm * MM_TO_TWIP))


def mm_pt(mm):
    return mm * MM_TO_TWIP / 20.0


def read_clauses(source, keep_punct):
    """读内容，拆成一条条（竖排时一条=一列，横排时一条=一行）。"""
    if source is None:
        source = "content.txt"
    if os.path.isfile(source):
        with open(source, "r", encoding="utf-8-sig") as f:
            text = f.read()
    else:
        text = source
    split_re = re.compile(r"[\n\r]") if keep_punct else SPLIT_RE
    clauses = []
    for part in split_re.split(text):
        chars = [ch for ch in part
                 if CJK_RE.match(ch) or (keep_punct and ch in PUNCT)]
        if chars:
            clauses.append(chars)
    return clauses


def build_items(clauses, repeat, blanks):
    """展开成单元格流：字符 / None(空格) / 'BREAK'(换列或换行)。"""
    items = []
    for clause in clauses:
        for ch in clause:
            items.extend([ch] * max(1, repeat))
            items.extend([None] * blanks)
        items.append("BREAK")
    if items and items[-1] == "BREAK":
        items.pop()
    return items


def paginate(items, cols, rows, order, top_rows=0):
    if order == "vertical":
        # 竖排：按列从右往左填充
        capacity = cols * rows
        step = rows
        pages = [{}]
        cursor = 0
        for item in items:
            if item == "BREAK":
                if cursor % step:
                    cursor = ((cursor // step) + 1) * step
                continue
            page_idx = cursor // capacity
            while len(pages) <= page_idx:
                pages.append({})
            pages[page_idx][cursor % capacity] = item
            cursor += 1
        return pages

    # 横排：每句占一行（超长自动换行），每一行在格子里水平居中
    eff_rows = rows - top_rows
    clauses, cur = [], []
    for it in items:
        if it == "BREAK":
            clauses.append(cur)
            cur = []
        else:
            cur.append(it)
    if cur:
        clauses.append(cur)

    pages = [{}]
    row = 0
    for cl in clauses:
        last = max((i for i, ch in enumerate(cl) if ch), default=-1)
        seg = cl[:last + 1]            # 去掉句尾空白
        L = len(seg)
        nrows = max(1, math.ceil(L / cols))
        if row + nrows > eff_rows:     # 当前页放不下，换页
            pages.append({})
            row = 0
        for r in range(nrows):
            part = seg[r * cols:(r + 1) * cols]
            plen = len(part)
            start = (cols - plen) // 2 if plen < cols else 0
            for i, ch in enumerate(part):
                if ch:
                    pages[-1][(row + r) * cols + start + i] = ch
        row += nrows
    return pages


def _remove(parent, tag):
    for el in parent.findall(qn(tag)):
        parent.remove(el)


def setup_table(tbl, widths_mm):
    """固定布局、无边框、零边距。"""
    tblPr = tbl._tbl.tblPr
    for tag in ("w:tblW", "w:tblBorders", "w:tblLayout", "w:tblCellMar"):
        _remove(tblPr, tag)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(sum(mm_twip(w) for w in widths_mm)))
    tblW.set(qn("w:type"), "dxa")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    mar = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    anchor_el = tblPr.find(qn("w:tblLook"))
    for el in (tblW, borders, layout, mar):
        if anchor_el is not None:
            anchor_el.addprevious(el)
        else:
            tblPr.append(el)
    grid = tbl._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl._tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths_mm:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(mm_twip(w)))
        grid.append(gc)
    tbl._tbl.insert(list(tbl._tbl).index(tblPr) + 1, grid)


def setup_cell(cell, width_mm, height_mm):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("w:tcW", "w:vAlign"):
        _remove(tcPr, tag)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(mm_twip(width_mm)))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)

    pPr = cell.paragraphs[0]._p.get_or_add_pPr()
    for tag in ("w:spacing", "w:jc"):
        _remove(pPr, tag)
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(mm_twip(height_mm)))
    spacing.set(qn("w:lineRule"), "exact")
    pPr.append(spacing)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    pPr.append(jc)


def write_char(cell, ch, font, font_pt, color):
    run = cell.paragraphs[0].add_run(ch)
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:hint"), "eastAsia")
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), font)
    rPr.append(rFonts)
    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rPr.append(c)
    sz_val = str(int(round(font_pt * 2)))
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), sz_val)
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), sz_val)
    rPr.append(szCs)


def _anchor_open(shape_id, x_mm, y_mm, w_mm, h_mm, behind):
    return f'''<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
 relativeHeight="{100 if behind else 200}" behindDoc="{1 if behind else 0}"
 locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>{int(x_mm * MM_TO_EMU)}</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>{int(y_mm * MM_TO_EMU)}</wp:posOffset></wp:positionV>
<wp:extent cx="{int(max(w_mm, 0.01) * MM_TO_EMU)}" cy="{int(max(h_mm, 0.01) * MM_TO_EMU)}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/>
<wp:docPr id="{shape_id}" name="Shape{shape_id}"/>
<wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:cNvSpPr/><wps:spPr>'''


ANCHOR_CLOSE = '''</wps:spPr><wps:bodyPr/></wps:wsp>
</a:graphicData></a:graphic>
</wp:anchor></w:drawing>'''


def dml_line(shape_id, x1, y1, x2, y2, color, weight_mm, dash=False, rot=False):
    """页面绝对定位直线（放页眉，每页重复，压在文字下方）。"""
    color = color.lstrip("#")
    w = max(abs(x2 - x1), 0.01)
    h = max(abs(y2 - y1), 0.01)
    rot_attr = ' rot="5400000"' if rot else ""
    dash_el = '<a:prstDash val="dash"/>' if dash else ""
    w_emu = int(weight_mm * MM_TO_EMU)
    body = f'''
<a:xfrm{rot_attr}><a:off x="0" y="0"/><a:ext cx="{int(w * MM_TO_EMU)}" cy="{int(h * MM_TO_EMU)}"/></a:xfrm>
<a:prstGeom prst="line"><a:avLst/></a:prstGeom>
<a:ln w="{w_emu}" cap="flat"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{dash_el}</a:ln>'''
    return _anchor_open(shape_id, x1, y1, w, h, behind=True) + body + ANCHOR_CLOSE


def dml_title_chars(shape_id_start, text, x0, y0, band_w, grid_h, font, color="000000", cell=15.0, rows=17):
    """标题：每个字一个水平文本框，绝对定位在标题带内垂直堆叠居中。
    水平文本框的居中在 Word / WPS / LibreOffice 里渲染都一致。"""
    frags = []
    n = len(text)
    # 标题字比格内范字小（约 0.68 格），整组在标题带内垂直居中、每字占一个格高
    pitch = cell
    char_mm = min(band_w, cell) * 0.68
    sz = int(round(mm_pt(char_mm) * 2))
    start_row = (rows - n) / 2.0
    for j, ch in enumerate(text):
        cy = y0 + (start_row + j + 0.5) * cell
        y = cy - pitch / 2
        body = f'''
<a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(band_w * MM_TO_EMU)}" cy="{int(pitch * MM_TO_EMU)}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:noFill/><a:ln><a:noFill/></a:ln>'''
        close = f'''</wps:spPr>
<wps:txbx><w:txbxContent>
<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
<w:r><w:rPr><w:rFonts w:hint="eastAsia" w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>
<w:color w:val="{color}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>
<w:t>{ch}</w:t></w:r>
</w:p>
</w:txbxContent></wps:txbx>
<wps:bodyPr anchor="ctr" anchorCtr="1" lIns="0" tIns="0" rIns="0" bIns="0"/>
</wps:wsp>
</a:graphicData></a:graphic>
</wp:anchor></w:drawing>'''
        frags.append(_anchor_open(shape_id_start + j, x0, y, band_w, pitch,
                                  behind=False) + body + close)
    return frags


def build_grid_shapes(x0, y0, title_w, cell, cols, rows, style,
                      grid_color, guide_color, top_rows=0):
    """米字格/田字格/方框：实线边框 + 虚线辅助线。"""
    frags = []
    sid = [1]

    def L(x1, y1, x2, y2, color, weight, dash=False, rot=False):
        frags.append(dml_line(sid[0], x1, y1, x2, y2, color, weight, dash, rot))
        sid[0] += 1

    tw = title_w
    total_w = tw + cols * cell
    total_h = rows * cell
    border = 0.16
    guide = 0.09

    # 竖线分界 x（含标题带边界）
    xs = [x0]
    if tw > 0:
        xs.append(x0 + tw)
    for k in range(1, cols + 1):
        xs.append(x0 + tw + k * cell)

    # 横线：逐段短线条（横向页面下，整页一根长线在个别渲染器里会被拉到页宽）
    for k in range(rows + 1):
        y = y0 + k * cell
        segs = list(zip(xs, xs[1:]))
        if tw > 0 and 0 < k < rows:
            segs = segs[1:]   # 标题带内部不画横线，保持干净竖条；仅顶/底封边
        for a, b in segs:
            L(a, y, b, y, grid_color, border)
    y_grid0 = y0   # 顶部标题行同样画格子（标题字写在田/米字格里）
    for x in xs:
        L(x, y_grid0, x, y0 + total_h, grid_color, border)

    if style == "box":
        return frags, sid[0]

    # 竖中线：整列贯穿（虚线）
    for i in range(cols):
        mx = x0 + tw + i * cell + cell / 2
        L(mx, y_grid0, mx, y0 + total_h, guide_color, guide, dash=True)
    # 横中线：逐列短虚线
    for j in range(rows):
        my = y0 + (j + 0.5) * cell
        for i in range(cols):
            cx0 = x0 + tw + i * cell
            L(cx0, my, cx0 + cell, my, guide_color, guide, dash=True)

    if style == "mizi":
        for i in range(cols):
            for j in range(rows):
                cx0 = x0 + tw + i * cell
                cy0 = y0 + j * cell
                L(cx0, cy0, cx0 + cell, cy0 + cell, guide_color, guide, dash=True)
                L(cx0, cy0, cx0 + cell, cy0 + cell, guide_color, guide, dash=True, rot=True)
    return frags, sid[0]


def add_header_grid(doc, frags):
    p = doc.sections[0].header.paragraphs[0]
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    pPr.append(spacing)
    for xml in frags:
        r = OxmlElement("w:r")
        r.append(parse_xml(xml))
        p._p.append(r)


def add_section_break(doc, header_part, page_w, page_h, m):
    """分节符（下一页），显式引用同一页眉，确保每页都有网格。"""
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "20")
    spacing.set(qn("w:lineRule"), "exact")
    pPr.append(spacing)
    rPr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "2")
    rPr.append(sz)
    pPr.append(rPr)
    sectPr = OxmlElement("w:sectPr")
    rid = doc.part.relate_to(
        header_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header")
    href = OxmlElement("w:headerReference")
    href.set(qn("w:type"), "default")
    href.set(qn("r:id"), rid)
    sectPr.append(href)
    pgSz = OxmlElement("w:pgSz")
    pgSz.set(qn("w:w"), str(mm_twip(page_w)))
    pgSz.set(qn("w:h"), str(mm_twip(page_h)))
    if page_w > page_h:
        pgSz.set(qn("w:orient"), "landscape")
    sectPr.append(pgSz)
    pgMar = OxmlElement("w:pgMar")
    for attr, v in (("top", m["top"]), ("right", m["right"]),
                    ("bottom", m["bottom"]), ("left", m["left"]),
                    ("header", 5), ("footer", 5), ("gutter", 0)):
        pgMar.set(qn(f"w:{attr}"), str(mm_twip(v)))
    sectPr.append(pgMar)
    pPr.append(sectPr)


def build_page(doc, cells, cols, rows, cell_mm, title_w, font,
               char_pt, color, order, title=None, top_rows=0):
    band = 1 if title_w > 0 else 0
    ncols = cols + band
    tbl = doc.add_table(rows=rows, cols=ncols)
    tbl.autofit = False
    widths = ([title_w] if band else []) + [cell_mm] * cols
    setup_table(tbl, widths)

    for j in range(rows):
        trPr = tbl.rows[j]._tr.get_or_add_trPr()
        trH = OxmlElement("w:trHeight")
        trH.set(qn("w:val"), str(mm_twip(cell_mm)))
        trH.set(qn("w:hRule"), "exact")
        trPr.append(trH)
        for i in range(ncols):
            setup_cell(tbl.cell(j, i), widths[i], cell_mm)

    # 顶部标题行（横排标准字帖）：标题每个字写进田/米字格，整行居中
    if top_rows and title:
        tchars = [ch for ch in title if not ch.isspace()]
        start = max(0, (ncols - len(tchars)) // 2)
        for i, ch in enumerate(tchars):
            col = start + i
            if col < ncols:
                write_char(tbl.cell(0, col), ch, font, char_pt, color)

    for k, ch in cells.items():
        if not ch:
            continue
        if order == "vertical":
            r = k % rows
            c = k // rows
            table_col = ncols - 1 - c
        else:
            r = k // cols + top_rows
            c = k % cols
            table_col = band + c
        write_char(tbl.cell(r, table_col), ch, font, char_pt, color)

    # 竖排时的标题由页眉文本框输出；横排顶部标题已在首行合并单元格中处理


def main():
    ap = argparse.ArgumentParser(
        description="生成 A4 可打印中文楷书字帖（Word 文档）")
    ap.add_argument("source", nargs="?", default="content.txt",
                    help="文本内容或 .txt 文件路径（默认 content.txt）")
    ap.add_argument("-o", "--output", default="字帖.docx")
    ap.add_argument("--title", default=None, help="标题：横排写在第一行（居中），竖排写在左侧")
    ap.add_argument("--paper", default="A4", choices=list(PAPERS.keys()))
    ap.add_argument("--orient", choices=["portrait", "landscape"], default="portrait",
                    help="portrait 纵向（默认）/ landscape 横向")
    ap.add_argument("--cell", type=float, default=15.0, help="格子边长 mm（默认 15）")
    ap.add_argument("--cols", type=int, default=0, help="强制每行字数")
    ap.add_argument("--rows", type=int, default=0, help="强制每页行数")
    ap.add_argument("--margin", type=float, default=14.0, help="页边距 mm（默认 14）")
    ap.add_argument("--title-col", type=float, default=18.0, help="标题列宽 mm（默认 18）")
    ap.add_argument("--font", default="stkaiti",
                    help="字体：stkaiti 华文楷体(默认) / wenkai 霞鹜文楷 / kaiti 楷体，"
                         "或 fonts/ 里的书家字体名，如 田英章楷书；用 --list-fonts 查看全部")
    ap.add_argument("--list-fonts", action="store_true", help="列出所有可用字体后退出")
    ap.add_argument("--font-file", default=None,
                    help="字体文件路径（.ttf/.otf/.ttc），自动识别字体名并安装；优先级高于 --font")
    ap.add_argument("--pdf", action="store_true",
                    help="同时导出 PDF（字体内嵌，适合发打印店/手机）")
    ap.add_argument("--char-scale", type=float, default=0.80, help="字占格比例（默认 0.80）")
    ap.add_argument("--order", choices=["vertical", "horizontal"], default="vertical",
                    help="vertical 竖排从右往左（默认）/ horizontal 横排")
    ap.add_argument("--grid-style", choices=["mizi", "tian", "box"], default="mizi",
                    help="mizi 米字格（默认）/ tian 田字格 / box 方框")
    ap.add_argument("--repeat", type=int, default=1, help="每字重复次数")
    ap.add_argument("--blanks", type=int, default=0, help="每字后留空字数")
    ap.add_argument("--trace", action="store_true", help="描红：浅灰色字")
    ap.add_argument("--pen-color", default=None,
                    help="范字颜色（十六进制，如 808080 灰色让笔迹更细淡；默认黑色）")
    ap.add_argument("--keep-punct", action="store_true", help="标点也占格")
    ap.add_argument("--grid-color", default="#000000")
    ap.add_argument("--guide-color", default="#A6A6A6")
    # .env 配置作为默认值（命令行显式参数仍优先）
    global FONTS_DIR
    env = load_env()
    if env.get("ZITIE_FONTS_DIR"):
        d = os.path.expanduser(env["ZITIE_FONTS_DIR"])
        FONTS_DIR = d if os.path.isabs(d) else os.path.join(os.getcwd(), d)
    ap.set_defaults(
        font=env.get("ZITIE_FONT", "stkaiti"),
        title=(env.get("ZITIE_TITLE") or None),
        order=env.get("ZITIE_ORDER", "vertical"),
        paper=env.get("ZITIE_PAPER", "A4"),
        grid_style=env.get("ZITIE_GRID_STYLE", "mizi"),
        output=env.get("ZITIE_OUTPUT", "字帖.docx"),
        cell=env_float(env, "ZITIE_CELL", 15.0),
        margin=env_float(env, "ZITIE_MARGIN", 14.0),
        char_scale=env_float(env, "ZITIE_CHAR_SCALE", 0.80),
        pdf=env_bool(env, "ZITIE_PDF", False),
    )
    args = ap.parse_args()

    if args.list_fonts:
        print_available_fonts()
        sys.exit(0)

    font_family, font_path = resolve_font(args.font, args.font_file)

    clauses = read_clauses(args.source, args.keep_punct)
    if not clauses:
        sys.exit("没有可写入的汉字，请检查内容。")
    items = build_items(clauses, args.repeat, args.blanks)

    page_w, page_h = PAPERS[args.paper]
    landscape = args.orient == "landscape"
    if landscape:
        page_w, page_h = max(page_w, page_h), min(page_w, page_h)  # 横向：宽>高
    m = {"top": args.margin, "bottom": args.margin,
         "left": args.margin, "right": args.margin}
    avail_w = page_w - m["left"] - m["right"]
    avail_h = page_h - m["top"] - m["bottom"] - 6
    top_title = bool(args.title) and args.order == "horizontal"
    top_rows = 1 if top_title else 0
    if top_title:
        title_w = 0.0                       # 横排：不要左侧标题列
    else:
        title_w = args.title_col if args.title else 0.0
    grid_w = avail_w - title_w

    if args.cols or args.rows:
        cell = min(grid_w / args.cols if args.cols else 1e9,
                   avail_h / args.rows if args.rows else 1e9)
        cell = math.floor(cell * 10) / 10
        cols = args.cols or int(grid_w // cell)
        rows = args.rows or int(avail_h // cell)
    else:
        cell = args.cell
        cols = int(grid_w // cell)
        rows = int(avail_h // cell)
    if cols < 1 or rows < 1:
        sys.exit("格子太大或页边距太大，一页放不下任何格子。")

    char_pt = mm_pt(cell) * args.char_scale
    if args.trace:
        color = "#D9D9D9"
    elif args.pen_color:
        color = "#" + args.pen_color.lstrip("#")
    else:
        color = None

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    sec.page_width = Mm(page_w)
    sec.page_height = Mm(page_h)
    sec.top_margin = Mm(m["top"])
    sec.bottom_margin = Mm(m["bottom"])
    sec.left_margin = Mm(m["left"])
    sec.right_margin = Mm(m["right"])
    sec.header_distance = Mm(5)
    sec.footer_distance = Mm(5)

    frags, next_id = build_grid_shapes(
        m["left"], m["top"], title_w, cell, cols, rows,
        args.grid_style, args.grid_color, args.guide_color, top_rows)
    if title_w > 0 and args.title:
        frags.extend(dml_title_chars(next_id, args.title, m["left"], m["top"],
                                     title_w, rows * cell, font_family, color or "000000",
                                     cell, rows))
    add_header_grid(doc, frags)

    pages = paginate(items, cols, rows, args.order, top_rows)
    for idx, cells in enumerate(pages):
        build_page(doc, cells, cols, rows, cell, title_w,
                  font_family, char_pt, color, args.order,
                  title=args.title if top_title else None, top_rows=top_rows)
        if idx < len(pages) - 1:
            add_section_break(doc, doc.sections[0].header.part, page_w, page_h, m)

    doc.save(args.output)
    total = sum(1 for it in items if it != "BREAK")
    out_abs = os.path.abspath(args.output)
    if font_path and str(font_path).lower().endswith(".ttf") and os.path.isfile(font_path):
        try:
            embed_font_in_docx(out_abs, font_path, font_family)
            print(f"📎 已内嵌字体（{font_family}）：Word/WPS 打开即显示，无需另装字体")
        except Exception as e:
            print(f"⚠️  字体内嵌失败（不影响使用，可装字体或用 PDF）：{e}")
    print(f"✅ 已生成 {args.output}")
    practice_rows = rows - top_rows
    print(f"   {args.paper}｜{cell:.1f}mm 格｜每页 {practice_rows} 行 × {cols} 格 = {practice_rows*cols} 字"
          f"｜{len(pages)} 页｜{total} 字（含空格）｜{args.grid_style}｜{args.order}")
    print(f"   字体：{font_family}")
    if args.pdf:
        export_pdf(out_abs, os.path.dirname(out_abs), font_path)


if __name__ == "__main__":
    main()
