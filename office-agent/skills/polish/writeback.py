"""修订写回：把"原句→改句"以 Word 修订标记（w:del / w:ins）写入 docx，
并为每处改动附批注说明理由。交付物打开即为标准修订视图，可逐条接受/拒绝。

v0 约束（README 有说明）：
- 修订按段落内子串定位；同段多处改动按顺序逐条应用；
- 重建段落时保留段落样式与首个 run 的字符格式（公文场景段内格式通常一致）。
"""
import copy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.run import Run

REV_DATE = "2026-06-12T00:00:00Z"


def _new_run(text: str, rpr=None, text_tag: str = "w:t"):
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement(text_tag)
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _wrap(tag: str, run_el, author: str, wid: int):
    el = OxmlElement(tag)
    el.set(qn("w:id"), str(wid))
    el.set(qn("w:author"), author)
    el.set(qn("w:date"), REV_DATE)
    el.append(run_el)
    return el


class _Counter:
    def __init__(self, start=1000):
        self.v = start

    def next(self):
        self.v += 1
        return self.v


def apply_revisions_to_paragraph(doc, paragraph, revisions, author="AI润色", counter=None):
    """revisions: [{old, new, reason}]，按出现顺序应用于同一段落。
    返回成功应用的条数。"""
    counter = counter or _Counter()
    full = paragraph.text
    # 先在纯文本层面校验并切分
    pieces, cursor, applied = [], 0, []
    for rev in revisions:
        idx = full.find(rev["old"], cursor)
        if idx < 0:
            continue
        if idx > cursor:
            pieces.append(("keep", full[cursor:idx], None))
        pieces.append(("change", rev["old"], rev))
        cursor = idx + len(rev["old"])
        applied.append(rev)
    if not applied:
        return 0
    if cursor < len(full):
        pieces.append(("keep", full[cursor:], None))

    # 保留首个 run 的字符格式
    rpr = None
    if paragraph.runs and paragraph.runs[0]._r.rPr is not None:
        rpr = paragraph.runs[0]._r.rPr
    for r in list(paragraph._p.findall(qn("w:r"))):
        paragraph._p.remove(r)

    # 重建：keep 为普通 run；change 为 w:del(原文) + w:ins(新文)，
    # 紧随其后放一个零宽锚点 run 用于挂批注。
    for kind, text, rev in pieces:
        if kind == "keep":
            paragraph._p.append(_new_run(text, rpr))
        else:
            del_run = _new_run(text, rpr, text_tag="w:delText")
            paragraph._p.append(_wrap("w:del", del_run, author, counter.next()))
            ins_run = _new_run(rev["new"], rpr)
            paragraph._p.append(_wrap("w:ins", ins_run, author, counter.next()))
            anchor_el = _new_run("", rpr)
            paragraph._p.append(anchor_el)
            if rev.get("reason"):
                anchor = Run(anchor_el, paragraph)
                doc.add_comment(anchor, text=rev["reason"], author=author, initials="AI")
    return len(applied)


def apply_revisions(doc, revisions, author="AI润色"):
    """revisions: [{para_index, old, new, reason}]，para_index 基于全文段落序。
    返回成功应用的条数。"""
    by_para: dict[int, list] = {}
    for rev in revisions:
        by_para.setdefault(int(rev["para_index"]), []).append(rev)
    counter = _Counter()
    total = 0
    paras = doc.paragraphs
    for i, revs in sorted(by_para.items()):
        if 0 <= i < len(paras):
            total += apply_revisions_to_paragraph(doc, paras[i], revs, author, counter)
    return total
