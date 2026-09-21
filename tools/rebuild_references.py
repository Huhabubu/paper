from __future__ import annotations

import copy
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from lxml import etree

ROOT = Path.cwd()
ZIP_PATH = ROOT / "小论文参考文献(1).zip"
FINAL_DIR = ROOT / "final"
FINAL_DOCX = FINAL_DIR / "韩观萍小论文_参考文献重构版.docx"
AUDIT_MD = FINAL_DIR / "参考文献重构说明.md"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
W_T = f"{{{W}}}t"
W_P = f"{{{W}}}p"
W_R = f"{{{W}}}r"
W_PPR = f"{{{W}}}pPr"
W_RPR = f"{{{W}}}rPr"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

REFERENCES = [
    "刘雪娇,郭嘉盛.城市更新语境下基于价值导向的工业遗产再利用探索[J].城市发展研究,2022,29(5):80-85.",
    "张环宙,高静,黄克己,等.文化遗产保护传承机制研究：基于空间-行为互动的理论视角[J].地理研究,2025,44(10):2769-2786.",
    "张艳,柴彦威.北京现代工业遗产的保护与文化内涵挖掘——基于城市单位大院的思考[J].城市发展研究,2013,20(2):23-28.",
    "逯百慧,边兰春,王康.关系重构与空间变迁——一个工业遗产更新转型案例[J].工业建筑,2023,53(12):1-10.",
    "宋永永,张桐雨,马蓓蓓,等.空间生产视角下城市旧工业区的更新机制与路径——基于“权力—资本—社会”框架的分析[J].地理与地理信息科学,2025,41(2):78-87.",
    "孙露,于涛.空间生产视角下旧城工业遗产更新机制研究——以南京国创园为例[J].现代城市研究,2024(9):17-24.",
    "宋润森,陈烨,苏晨,等.空间生产视角下的城市工业遗产向文创空间转型研究——以济南市工业—文创空间为例[J].城市建筑,2024,21(23):55-58.",
    "周雪光,艾云.多重逻辑下的制度变迁：一个分析框架[J].中国社会科学,2010(4):132-150.",
    "阮海波,梁咏琪.多重制度逻辑下社区违建治理的困境及路径研究——基于四川省Y市H社区的田野调查[J].城市问题,2026(2):77-89.",
    "操小晋,朱天可,邓元媛,等.老旧单位社区更新的治理机制重构——基于行动者网络理论视角[J].现代城市研究,2022(6):113-119.",
    "王建英,孙琦,邹利林.基于行动者网络理论的海岛乡村空间演化路径——以福建平潭国际旅游岛北港村为例[J].地理科学,2025,45(5):1039-1049.",
    "席岳婷,马文博,张配.文旅融合视域下工业遗产旅游推动城市更新的作用研究[J].经济研究参考,2025(1):130-144.",
    "刘天宝,柴彦威.中国城市单位制研究进展[J].地域研究与开发,2013,32(5):13-21.",
    "王华,郑艳芬.遗产地农村社区参与旅游发展的制度嵌入性——丹霞山瑶塘村与断石村比较研究[J].地理研究,2016,35(6):1164-1176.",
]

REPLACEMENTS = [
    ("社会空间载体。", "社会空间载体[1-2]。"),
    ("地方认同载体。", "地方认同载体[2-3]。"),
    ("不同价值诉求和行动目标之间的协调关系成为影响空间转型路径的重要因素。", "不同价值诉求和行动目标之间的协调关系成为影响空间转型路径的重要因素[4-7]。"),
    ("多重制度逻辑为工业遗产价值重构提供方向和规则基础，", "多重制度逻辑为工业遗产价值重构提供方向和规则基础[8-9]，"),
    ("行动者网络通过关系重组和资源整合实现制度价值向空间实践的转化，", "行动者网络通过关系重组和资源整合实现制度价值向空间实践的转化[10-11]，"),
    ("而空间实践进一步推动物质空间、社会关系的重新生产。", "而空间实践进一步推动物质空间、社会关系的重新生产[5-7]。"),
    ("揭示行政、市场、遗产保护和社会认同等多重逻辑对空间转型方向的影响；", "揭示行政、市场、遗产保护和社会认同等多重逻辑对空间转型方向的影响[8-9]；"),
    ("通过问题呈现、利益赋予、征召动员、异议与协调等逐步构建起一个动态演化的行动者网络。", "通过问题呈现、利益赋予、征召动员、异议与协调等逐步构建起一个动态演化的行动者网络[10-11]。"),
    ("空间生产理论则进一步解释其物质空间和社会关系的重构，并最终实现地方空间再生产。", "空间生产理论则进一步解释其物质空间和社会关系的重构，并最终实现地方空间再生产[5-7]。"),
    ("历史遗存[37]。", "历史遗存[12]。"),
    ("中国城市空间生产从属于国家计划经济体制和单位制社会组织结构。", "中国城市空间生产从属于国家计划经济体制和单位制社会组织结构[13]。"),
    ("两类制度逻辑在转型初期产生显著的张力：行政逻辑所驱动的空间置换需求和保护逻辑的限制开发强度形成了策略博弈。", "两类制度逻辑在转型初期产生显著的张力：行政逻辑所驱动的空间置换需求和保护逻辑的限制开发强度形成了策略博弈[1-2]。"),
    ("两大逻辑共同推动老厂区由封闭生产空间向开放型文化集聚区转型，并为后续的空间再生产奠定基础。", "两大逻辑共同推动老厂区由封闭生产空间向开放型文化集聚区转型，并为后续的空间再生产奠定基础[10-11]。"),
    ("不同逻辑的互动博弈，推动西影行动者网络的扩张和进一步重组。", "不同逻辑的互动博弈，推动西影行动者网络的扩张和进一步重组[12,14]。"),
    ("对于本地市民与外来游客，这片曾经对外封闭的厂区，已转化为可供大众休闲游憩的公共开放空间。", "对于本地市民与外来游客，这片曾经对外封闭的厂区，已转化为可供大众休闲游憩的公共开放空间[4-7,12]。"),
    ("多重制度逻辑在协商和碰撞中实现规则调适，推动遗产价值从单一的工业生产价值向包含文化价值、经济价值和社会价值的复合空间价值重构，为空间再生产提供持续的制度基础和发展动力。", "多重制度逻辑在协商和碰撞中实现规则调适，推动遗产价值从单一的工业生产价值向包含文化价值、经济价值和社会价值的复合空间价值重构，为空间再生产提供持续的制度基础和发展动力[8-9]。"),
    ("多元主体在空间共享的共同目标下达成动态利益平衡，保障行动者网络的稳固运行和可持续演进。", "多元主体在空间共享的共同目标下达成动态利益平衡，保障行动者网络的稳固运行和可持续演进[10-11]。"),
    ("微观主体的深度参与及其受益程度，推动物质空间的更新与精神文化符号的相互作用，为西影工业遗产空间的持久活化注入动力。", "微观主体的深度参与及其受益程度，推动物质空间的更新与精神文化符号的相互作用，为西影工业遗产空间的持久活化注入动力[1-2,7,12,14]。"),
]

CIT_RE = re.compile(r"\[(?:\d+)(?:(?:-|,)(?:\d+))*\]")

def para_text(p):
    return "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=NS))

def replace_across_text_nodes(p, old: str, new: str):
    nodes = p.xpath(".//w:t", namespaces=NS)
    full = "".join(n.text or "" for n in nodes)
    pos = full.find(old)
    if pos < 0:
        return False
    end = pos + len(old)

    starts = []
    cur = 0
    for n in nodes:
        txt = n.text or ""
        starts.append((n, cur, cur + len(txt)))
        cur += len(txt)

    start_i = end_i = None
    for i, (_, s, e) in enumerate(starts):
        if start_i is None and s <= pos <= e:
            if pos < e or (pos == e and len(old) == 0):
                start_i = i
        if s < end <= e or (end == s and end == pos):
            end_i = i
            break
    if start_i is None:
        return False
    if end_i is None:
        end_i = len(starts) - 1

    sn, ss, se = starts[start_i]
    en, es, ee = starts[end_i]
    st = sn.text or ""
    et = en.text or ""
    start_off = pos - ss
    end_off = end - es

    if start_i == end_i:
        sn.text = st[:start_off] + new + st[end_off:]
        if sn.text.startswith(" ") or sn.text.endswith(" "):
            sn.set(XML_SPACE, "preserve")
        return True

    sn.text = st[:start_off] + new
    if sn.text.startswith(" ") or sn.text.endswith(" "):
        sn.set(XML_SPACE, "preserve")
    for i in range(start_i + 1, end_i):
        starts[i][0].text = ""
    en.text = et[end_off:]
    if en.text.startswith(" ") or en.text.endswith(" "):
        en.set(XML_SPACE, "preserve")
    return True

def remove_specific_broken_ref_field(root, field_token: str):
    # Remove one pre-existing broken Word REF field that LibreOffice renders as
    # "Error: Reference source not found". The field has no visible cached text
    # in the source DOCX, so removing it preserves the paper's substantive text.
    for p in root.xpath("//w:body/w:p", namespaces=NS):
        children = list(p)
        i = 0
        while i < len(children):
            child = children[i]
            if child.tag != W_R:
                i += 1
                continue
            fld = child.find(f"{{{W}}}fldChar")
            if fld is None or fld.get(f"{{{W}}}fldCharType") != "begin":
                i += 1
                continue
            j = i + 1
            instr_parts = []
            end_idx = None
            while j < len(children):
                c = children[j]
                if c.tag == W_R:
                    for instr in c.findall(f"{{{W}}}instrText"):
                        instr_parts.append(instr.text or "")
                    fc = c.find(f"{{{W}}}fldChar")
                    if fc is not None and fc.get(f"{{{W}}}fldCharType") == "end":
                        end_idx = j
                        break
                j += 1
            if end_idx is not None and field_token in "".join(instr_parts):
                for k in range(end_idx, i - 1, -1):
                    p.remove(children[k])
                return True
            i = (end_idx + 1) if end_idx is not None else (i + 1)
    return False

def set_paragraph_text_like_template(p, text: str):
    # Keep paragraph properties and first run properties; remove other runs/content.
    # The source reference paragraphs use automatic list numbering. We write static
    # [n] labels for deterministic GB/T 7714-style rendering, so drop w:numPr to
    # avoid duplicate labels such as "[1] [1]".
    ppr = p.find(W_PPR)
    if ppr is not None:
        num_pr = ppr.find(f"{{{W}}}numPr")
        if num_pr is not None:
            ppr.remove(num_pr)
    first_run = p.find(W_R)
    rpr = copy.deepcopy(first_run.find(W_RPR)) if first_run is not None and first_run.find(W_RPR) is not None else None

    for child in list(p):
        if child is not ppr:
            p.remove(child)

    r = etree.Element(W_R)
    if rpr is not None:
        r.append(rpr)
    t = etree.SubElement(r, W_T)
    t.text = text
    if text.startswith(" ") or text.endswith(" "):
        t.set(XML_SPACE, "preserve")
    p.append(r)

def extract_docx_from_archive(tmpdir: Path) -> Path:
    with zipfile.ZipFile(ZIP_PATH, "r") as zf:
        names = [n for n in zf.namelist() if n.lower().endswith(".docx")]
        if len(names) != 1:
            raise RuntimeError(f"Expected exactly one DOCX in archive, found: {names}")
        data = zf.read(names[0])
        out = tmpdir / Path(names[0]).name
        out.write_bytes(data)
        return out

def text_before_references(xml_bytes: bytes):
    root = etree.fromstring(xml_bytes)
    paras = root.xpath("//w:body/w:p", namespaces=NS)
    out = []
    for p in paras:
        txt = para_text(p)
        if txt.strip() == "参考文献：":
            break
        out.append(txt)
    return "\n".join(out)

def count_media(docx_path: Path):
    with zipfile.ZipFile(docx_path, "r") as z:
        return len([n for n in z.namelist() if n.startswith("word/media/") and not n.endswith("/")])

def patch_document_xml(xml_bytes: bytes) -> bytes:
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(xml_bytes, parser)
    body = root.find(f".//{{{W}}}body")
    paras = body.findall(W_P)

    replaced = []
    for old, new in REPLACEMENTS:
        found = False
        for p in paras:
            txt = para_text(p)
            if old in txt:
                if replace_across_text_nodes(p, old, new):
                    replaced.append(old)
                    found = True
                    break
        if not found:
            raise RuntimeError(f"Required citation insertion target not found: {old}")

    # The old [37] was the cached result of a broken REF field. After replacing
    # it with static [12] text above, remove the stale field markers so office
    # renderers do not recompute them into an error message.
    if not remove_specific_broken_ref_field(root, "_Ref21614"):
        raise RuntimeError("Expected broken REF _Ref21614 field not found")

    # Rebuild the entire reference section from scratch.
    paras = body.findall(W_P)
    heading_idx = None
    for i, p in enumerate(paras):
        if para_text(p).strip() == "参考文献：":
            heading_idx = i
            break
    if heading_idx is None:
        raise RuntimeError("Reference heading not found")

    template = copy.deepcopy(paras[heading_idx + 1]) if heading_idx + 1 < len(paras) else copy.deepcopy(paras[heading_idx])

    # Remove every paragraph after the heading; preserve sectPr and other non-paragraph body elements.
    for p in list(body.findall(W_P))[heading_idx + 1:]:
        body.remove(p)

    heading = list(body.findall(W_P))[heading_idx]
    insert_at = body.index(heading) + 1
    for i, ref in enumerate(REFERENCES, start=1):
        p = copy.deepcopy(template)
        set_paragraph_text_like_template(p, f"[{i}] {ref}")
        body.insert(insert_at, p)
        insert_at += 1

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

def main():
    if not ZIP_PATH.exists():
        raise SystemExit(f"Archive not found: {ZIP_PATH}")
    FINAL_DIR.mkdir(exist_ok=True)

    with tempfile.TemporaryDirectory() as td:
        tdir = Path(td)
        src = extract_docx_from_archive(tdir)

        with zipfile.ZipFile(src, "r") as zin:
            original_xml = zin.read("word/document.xml")
            patched_xml = patch_document_xml(original_xml)

            with zipfile.ZipFile(FINAL_DOCX, "w") as zout:
                for item in zin.infolist():
                    data = patched_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                    zout.writestr(item, data)

        # Verification: body text before references must be identical after citation markers are stripped.
        with zipfile.ZipFile(FINAL_DOCX, "r") as z:
            final_xml = z.read("word/document.xml")

        before = CIT_RE.sub("", text_before_references(original_xml))
        after = CIT_RE.sub("", text_before_references(final_xml))
        if before != after:
            raise RuntimeError("Body text changed beyond citation markers")

        if count_media(src) != count_media(FINAL_DOCX):
            raise RuntimeError("Embedded media count changed")

    audit = f"""# 参考文献重构说明

## 处理原则

- 正文研究方向、论证结构、案例材料与结论不改。
- 删除原正文旧引用编号（原有 `[37]`），重新建立正文引用编号。
- 原“参考文献”部分全部弃用并重建。
- 新参考文献仅从压缩包内提供的参考资料中筛选。
- 引用体系采用顺序编码制：正文使用 `[1]`、`[4-7]`、`[12,14]` 等形式，文末按首次出现顺序编号。
- 文末著录按 GB/T 7714 常用期刊格式统一：作者.题名[J].刊名,年,卷(期):页码.

## 筛选结果

共保留 {len(REFERENCES)} 篇，与论文的四个核心主题对应：

1. 工业遗产保护、再利用与城市更新；
2. 空间生产与工业遗产空间重构；
3. 多重制度逻辑及制度冲突/协商；
4. 行动者网络、单位制转型与遗产旅游中的多主体互动。

## 自动核验

- 正文在去除引用编号后，与原稿正文逐字一致：通过。
- 原参考文献列表：全部删除。
- 新参考文献条目：{len(REFERENCES)} 条。
- 原 DOCX 内嵌媒体数量与修改后保持一致：通过。
"""
    AUDIT_MD.write_text(audit, encoding="utf-8")
    print(f"Created: {FINAL_DOCX}")
    print(f"Created: {AUDIT_MD}")

if __name__ == "__main__":
    main()
