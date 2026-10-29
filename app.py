# -*- coding: utf-8 -*-
"""
报关单自动生成器 —— 网页接口
流程：上传深圳SO销售订单PDF → 解析(客户/报价号/商品编码/数量/单价) → 按ERP编码查HS对照表
      → 写入报关单(隐藏商品编码列，HS补入) → 子表(发票/箱单/合同/申报要素)公式自动联动
"""
import os, re, json, sys, subprocess, shutil, zipfile, tempfile, datetime
from urllib.parse import quote
from flask import Flask, request, jsonify, send_file, send_from_directory, render_template_string, Response

BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE, 'base_template.xlsx')
MASTER_XLSX = os.path.join(BASE, 'master.xlsx')          # 用户填好的 HS 对照表(程序用)
MASTER_DOWNLOAD = os.path.join(BASE, 'HS对照表_待填.xlsx') # 给用户下载填写的表
BOM_JSON = os.path.join(BASE, 'master_seed', 'bom_full.json')
UPLOAD = os.path.join(BASE, 'uploads')
OUTPUT = os.path.join(BASE, 'output')
# 可选依赖：workbuddy-pdf skill（PDF → Markdown）。用环境变量覆盖，缺省按工作台安装位置探测。
SKILL = os.environ.get('PDF_SKILL_DIR') or os.path.join(
    os.environ.get('LOCALAPPDATA', ''), 'Programs', 'WorkBuddy', 'resources',
    'app.asar.unpacked', 'resources', 'plugins', 'workbuddy-builtin',
    'skills', 'workbuddy-pdf')

for d in (UPLOAD, OUTPUT):
    os.makedirs(d, exist_ok=True)

# 首次运行（或在别的机器上 clone 下来）时，用脱敏样例初始化两张对照表，保证开箱能跑。
for _real, _demo in ((MASTER_XLSX, 'master.example.xlsx'),
                     (MASTER_DOWNLOAD, 'HS对照表_待填.example.xlsx')):
    _src = os.path.join(BASE, _demo)
    if not os.path.exists(_real) and os.path.exists(_src):
        shutil.copy2(_src, _real)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

# ============ 主数据(HS对照表)读写 ============

def master_path():
    return MASTER_XLSX if os.path.exists(MASTER_XLSX) else MASTER_DOWNLOAD

def load_master():
    """读 HS 对照表 → {ERP编码: {hs,name,brand,model,unit,price,gw,nw}}
    按表头名定位列（用户增删列也不会错位）。"""
    import openpyxl
    p = master_path()
    if not os.path.exists(p):
        return {}
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb.active
    header = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]

    def idx(*names):
        for n in names:
            for i, h in enumerate(header):
                if n and n in h:
                    return i
        return None

    i_erp = idx('ERP商品编码', 'ERP', '商品编码')
    i_hs = idx('HS')
    i_name = idx('品名')
    i_model = idx('型号')
    i_brand = idx('品牌')
    i_unit = idx('单位')
    i_price = idx('单价')
    i_gw = idx('毛重')
    i_nw = idx('净重')
    if i_erp is None:
        wb.close()
        return {}

    def cell(row, i):
        if i is None or i >= len(row) or row[i] is None:
            return ''
        v = row[i]
        return v.strip() if isinstance(v, str) else v

    m = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[i_erp] in (None, ''):
            continue
        erp = str(row[i_erp]).strip()
        m[erp] = {'hs': str(cell(row, i_hs) or '').strip(),
                  'name': str(cell(row, i_name) or '').strip(),
                  'brand': str(cell(row, i_brand) or '').strip(),
                  'model': str(cell(row, i_model) or '').strip(),
                  'unit': str(cell(row, i_unit) or '').strip(),
                  'price': cell(row, i_price),
                  'gw': cell(row, i_gw),
                  'nw': cell(row, i_nw)}
    wb.close()
    return m

def bom_names():
    """BOM 全量名称(用于品名兜底，ERP→中文名)"""
    if not os.path.exists(BOM_JSON):
        return {}
    return {r['erp']: r['name'] for r in json.load(open(BOM_JSON, encoding='utf-8'))}

def bom_meta():
    """BOM 品牌/型号 (ERP→{brand, model})"""
    if not os.path.exists(BOM_JSON):
        return {}
    return {r['erp']: {'brand': r.get('brand', ''), 'model': r.get('model', '')}
            for r in json.load(open(BOM_JSON, encoding='utf-8'))}

# ============ PDF 解析 ============

def _find_sh():
    import shutil as _s
    p = _s.which('sh')
    if p:
        return p
    for c in (r'C:\Users\mt\AppData\Local\Programs\WorkBuddy\resources\vendor\PortableGit\usr\bin\sh.exe',
              r'C:\Users\mt\.workbuddy\binaries\PortableGit\versions\1.3.0\usr\bin\sh.exe',
              r'C:\Program Files\Git\usr\bin\sh.exe',
              r'C:\Program Files\Git\bin\sh.exe'):
        if os.path.exists(c):
            return c
    return 'sh'

def _pdf_skill_dir():
    """定位可选的 PDF → Markdown 转换工具（workbuddy-pdf skill）。

    可用环境变量 PDF_SKILL_DIR 指定；找不到就返回 None，走 pdftotext 兜底，
    这样换台机器 / 别人 clone 也能跑，不依赖某台电脑上的绝对路径。
    """
    cands = [os.environ.get('PDF_SKILL_DIR', ''), SKILL]
    for c in cands:
        if c and os.path.exists(os.path.join(c, 'scripts', 'run.sh')):
            return c
    return None


def extract_md(pdf):
    out = os.path.join(tempfile.mkdtemp(), 'c.md')
    sk = _pdf_skill_dir()
    if sk:
        r = subprocess.run([_find_sh(), os.path.join(sk, 'scripts', 'run.sh'),
                            pdf, out, '-f', 'md', '-q'],
                           capture_output=True, text=True)
        if r.returncode != 0:
            raise RuntimeError('PDF 解析失败: ' + (r.stderr or '')[:200])
        return open(out, encoding='utf-8').read()

    # 兜底：poppler 的 pdftotext（保留版面，表格列不会挤成一团）
    exe = shutil.which('pdftotext')
    if not exe:
        raise RuntimeError(
            '未找到 PDF 解析工具。请二选一：\n'
            '  1) 安装 poppler 并确保 pdftotext 在 PATH 里；\n'
            '  2) 设置环境变量 PDF_SKILL_DIR 指向 workbuddy-pdf skill 目录。')
    r = subprocess.run([exe, '-layout', '-enc', 'UTF-8', pdf, out],
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError('PDF 解析失败: ' + (r.stderr or '')[:200])
    return open(out, encoding='utf-8').read()

def parse_order(md):
    norm = lambda s: re.sub(r'(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])', '', s)
    lines = [norm(l.rstrip()) for l in md.splitlines()]
    text = '\n'.join(lines)
    hdr = {'客户': '', '联系人': '', '报价日期': '', '报价号码': '', '币别': '', '信用条件': ''}
    for key, pat in [('客户', r'客户\s*[：:]\s*([^\n]+)'),
                     ('联系人', r'联系人\s*[：:]\s*([^\n]+)'),
                     ('报价日期', r'报价日期[：:]\s*(\S+)'),
                     ('报价号码', r'报价号码[：:]\s*(\S+)'),
                     ('币别', r'币别[：:]\s*(\S+)'),
                     ('信用条件', r'信用条件[：:]\s*(\S+?)(?=\s*地址|$)')]:
        m = re.search(pat, text)
        if m:
            hdr[key] = m.group(1).strip()
    # 商品行：号序 ERP编码 品名 规格 单位 数量 单价 税率 税额 金额
    items = []
    started = False
    buf = None
    def flush(buf):
        if not buf:
            return
        m = re.match(r'^(\d+)\s+([A-Za-z]{1,4}\d{2,})\s+(.*)$', buf)
        if not m:
            return
        no, erp, rest = m.group(1), m.group(2), m.group(3)
        tm = re.findall(r'((?:套|PCS|个|pcs|PCS|Pcs|件))\s+([\d.]+)\s+([\d,]+\.?\d*)\s+(\d+%)\s+([\d,]+\.?\d*)', rest)
        if not tm:
            return
        unit, qty, price, tax, amt = tm[-1]
        name = rest[:rest.find(unit)].strip()
        items.append({'no': int(no), 'erp': erp, 'name': name, 'unit': unit,
                      'qty': float(qty), 'price': float(price.replace(',', '')),
                      'amount': float(amt.replace(',', ''))})
    for l in lines:
        if '号序' in l and '商品编码' in l:
            started = True
            continue
        if '合计' in l:
            flush(buf)
            break
        if not started:
            continue
        if re.match(r'^\d+\s+[A-Za-z]{1,4}\d{2,}', l):
            flush(buf)
            buf = l
        else:
            buf = (buf + ' ' + l) if buf else l
    return hdr, items

# ============ 报关单写入 ============

def _colnum(letters):
    """列字母 → 列序号（A=1, Z=26, AA=27）。"""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _near_style(xml, col_letter, row):
    """某列要新增单元格时，取同列最近一行的样式（优先上方、没有就取下方）。

    模板每个 SKU 区只把前 1~2 行做了完整样式，第 3 行往后是空白格；
    直接写值会得到「无边框 / 默认对齐」的裸格，所以借用同列已有样式。
    """
    pat = re.compile(r'<c r="%s(\d+)"([^>]*?)(?:/>|>)' % col_letter)
    up, down = None, None
    for m in pat.finditer(xml):
        r = int(m.group(1))
        sm = re.search(r'\bs="(\d+)"', m.group(2))
        if not sm:
            continue
        if r < row and (up is None or r > up[0]):
            up = (r, sm.group(1))
        elif r > row and (down is None or r < down[0]):
            down = (r, sm.group(1))
    pick = up or down
    return pick[1] if pick else None


def put_cell(xml, ref, kind, content, is_str=False, style=None):
    col_letter = re.match(r'([A-Z]+)', ref).group(1)
    row = int(re.match(r'[A-Z]+(\d+)', ref).group(1))
    pat = re.compile(r'<c r="%s"([^>]*?)(?:/>|>.*?</c>)' % re.escape(ref), re.DOTALL)
    m = pat.search(xml)
    if m:
        attrs = m.group(1)
        sm = re.search(r'\bs="([^"]+)"', attrs)
        sattr = (' s="%s"' % sm.group(1)) if sm else ''
        tattr = ' t="str"' if is_str else ''
        inner = ('<f>%s</f>' % content) if kind == 'f' else ('<v>%s</v>' % esc(content))
        return xml[:m.start()] + '<c r="%s"%s%s>%s</c>' % (ref, sattr, tattr, inner) + xml[m.end():]
    if style is None:
        style = _near_style(xml, col_letter, row)
    sattr = (' s="%s"' % style) if style else ''
    tattr = ' t="str"' if is_str else ''
    inner = ('<f>%s</f>' % content) if kind == 'f' else ('<v>%s</v>' % esc(content))
    newcell = '<c r="%s"%s%s>%s</c>' % (ref, sattr, tattr, inner)
    rowpat = re.compile(r'(<row r="%d"[^>]*>)(.*?)(</row>)' % row, re.DOTALL)
    rm = rowpat.search(xml)
    if not rm:
        return xml.replace('</sheetData>', '<row r="%d">%s</row></sheetData>' % (row, newcell), 1)
    # 按列序插入，避免新单元格落到行尾导致 <c> 顺序错乱（Excel / WPS 都要求列升序）
    body = rm.group(2)
    tgt = _colnum(col_letter)
    pos = len(body)
    for cm in re.finditer(r'<c\b[^>]*?\br="([A-Z]+)\d+"', body):
        if _colnum(cm.group(1)) > tgt:
            pos = cm.start()
            break
    return xml[:rm.start()] + rm.group(1) + body[:pos] + newcell + body[pos:] + '</row>' + xml[rm.end():]

def unify_col(x, col, rows, ref_row):
    """把某一列若干行的样式统一成参考行的样式（模板各行样式不齐时用）。"""
    sidx = _cell_style(x, '%s%d' % (col, ref_row))
    if sidx is None:
        return x
    for r in rows:
        if r == ref_row:
            continue
        x = set_cell_style(x, '%s%d' % (col, r), int(sidx))
    return x


def esc(t):
    return str(t).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


# ============ 单元格样式：统一字体 / 货币 / 居中 ============
# 模板每个 SKU 区只有前 1~2 行是完整样式（字体、字号、货币格式都对），
# 第 3 行往后是宋体10 / Arial11 / 常规数字格式 / 右对齐 —— 直接写值就会「一行一个样」。
# 所以按单元格克隆 xf：只改字体 或 只改数字格式+对齐，字体/边框/填充其余部分原样沿用。
MONEY_FMT_XML = '&quot;US$&quot;#,##0.00;\\-&quot;US$&quot;#,##0.00'


class StyleForge:
    """按需克隆 cellXfs，最后一次性写回 styles.xml。"""

    def __init__(self, sx):
        self.sx = sx
        self.numfmts = {int(a): b for a, b in
                        re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', sx)}
        j = sx.find('<cellXfs')
        self.ok = j >= 0
        if not self.ok:
            return
        body = sx[sx.find('>', j) + 1:sx.find('</cellXfs>')]
        self.xfs = re.findall(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', body, re.S)
        self.n_orig = len(self.xfs)
        self.cache = {}

    def _attr(self, sidx, name, default='0'):
        m = re.search(r'\b%s="([^"]*)"' % name, self.xfs[sidx])
        return m.group(1) if m else default

    def font_of(self, sidx):
        return int(self._attr(sidx, 'fontId'))

    def numfmt_of(self, sidx):
        return int(self._attr(sidx, 'numFmtId'))

    def halign_of(self, sidx):
        m = re.search(r'<alignment[^>]*\bhorizontal="([^"]*)"', self.xfs[sidx])
        return m.group(1) if m else None

    def valign_of(self, sidx):
        m = re.search(r'<alignment[^>]*\bvertical="([^"]*)"', self.xfs[sidx])
        return m.group(1) if m else None

    def ensure_numfmt(self, code):
        """找到或登记一个数字格式，返回 numFmtId。"""
        for fid, c in self.numfmts.items():
            if c == code:
                return fid
        nid = max(list(self.numfmts) + [163]) + 1
        self.numfmts[nid] = code
        entry = '<numFmt numFmtId="%d" formatCode="%s"/>' % (nid, code)
        m = re.search(r'<numFmts count="(\d+)">', self.sx)
        if m:
            self.sx = self.sx.replace(m.group(0), '<numFmts count="%d">' % (int(m.group(1)) + 1), 1)
            self.sx = self.sx.replace('</numFmts>', entry + '</numFmts>', 1)
        else:
            self.sx = re.sub(r'(<fonts\b)', '<numFmts count="1">%s</numFmts>' % entry + r'\1',
                             self.sx, count=1)
        return nid

    def variant(self, sidx, font_id=None, numfmt_id=None, halign=None, valign=None):
        """在 sidx 基础上只改指定项，返回新的样式索引（相同请求复用同一条）。"""
        if not self.ok or sidx is None or sidx >= len(self.xfs):
            return sidx
        if ((font_id is None or font_id == self.font_of(sidx))
                and (numfmt_id is None or numfmt_id == self.numfmt_of(sidx))
                and (halign is None or halign == self.halign_of(sidx))
                and (valign is None or valign == self.valign_of(sidx))):
            return sidx          # 已经是目标样式，不必新增
        key = (sidx, font_id, numfmt_id, halign, valign)
        if key in self.cache:
            return self.cache[key]
        xf = self.xfs[sidx]
        if font_id is not None and font_id != self.font_of(sidx):
            xf = re.sub(r'fontId="\d+"', 'fontId="%d"' % font_id, xf, count=1)
            xf = (re.sub(r'applyFont="[^"]*"', 'applyFont="1"', xf) if 'applyFont' in xf
                  else xf.replace('<xf ', '<xf applyFont="1" ', 1))
        if numfmt_id is not None:
            xf = re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % numfmt_id, xf, count=1)
            xf = (re.sub(r'applyNumberFormat="[^"]*"', 'applyNumberFormat="1"', xf)
                  if 'applyNumberFormat' in xf else xf.replace('<xf ', '<xf applyNumberFormat="1" ', 1))
        if halign is not None or valign is not None:
            xf = (re.sub(r'applyAlignment="[^"]*"', 'applyAlignment="1"', xf)
                  if 'applyAlignment' in xf else xf.replace('<xf ', '<xf applyAlignment="1" ', 1))
            if '<alignment' not in xf and xf.endswith('/>'):
                xf = xf[:-2] + '></xf>'      # 自闭合 <xf/> 先展开，好塞 alignment
            for name, val in (('horizontal', halign), ('vertical', valign)):
                if val is None:
                    continue
                if re.search(r'<alignment[^>]*\b%s="[^"]*"' % name, xf):
                    xf = re.sub(r'(<alignment[^>]*?)\b%s="[^"]*"' % name, r'\1%s="%s"' % (name, val), xf)
                elif '<alignment' in xf:
                    xf = xf.replace('<alignment', '<alignment %s="%s"' % (name, val), 1)
                else:
                    xf = xf.replace('</xf>', '<alignment %s="%s"/></xf>' % (name, val), 1)
        new = len(self.xfs)
        self.xfs.append(xf)
        self.cache[key] = new
        return new

    def dump(self):
        if not self.ok or len(self.xfs) == self.n_orig:
            return self.sx
        sx = self.sx
        j = sx.find('<cellXfs')
        return (sx[:j] + '<cellXfs count="%d">' % len(self.xfs)
                + ''.join(self.xfs) + sx[sx.find('</cellXfs>'):])



def _cell_style(x, ref):
    m = re.search(r'<c r="%s"([^>]*?)(?:/>|>)' % re.escape(ref), x)
    if not m:
        return None
    sm = re.search(r'\bs="(\d+)"', m.group(1))
    return sm.group(1) if sm else None


def set_cell_style(x, ref, sidx):
    """把某个单元格的样式索引换成 sidx（不动内容）。"""
    m = re.search(r'<c r="%s"([^>]*?)(?:/>|>)' % re.escape(ref), x)
    if not m or sidx is None:
        return x
    attrs = m.group(1)
    if re.search(r'\bs="\d+"', attrs):
        attrs = re.sub(r'\bs="\d+"', 's="%d"' % sidx, attrs)
    else:
        attrs = ' s="%d"' % sidx + attrs
    return x[:m.start(1)] + attrs + x[m.end(1):]

def _put_hs(x, ref, hs):
    """HS 编码：纯数字就按数字写，避免 Excel/WPS 一路弹「以文本形式存储的数字」绿三角。
    只有以 0 开头（如 0301）才保留文本，否则前导 0 会丢。"""
    hs = (hs or '').strip()
    if hs.isdigit() and not hs.startswith('0') and len(hs) <= 11:
        return put_cell(x, ref, 'v', int(hs))
    return put_cell(x, ref, 'v', hs, is_str=True)


def _put_blank(x, ref, sidx):
    """放一个只保留样式的空单元格。

    合并区域里被吞掉的格子不能删：Excel 是根据区域内**每一格**的边框来画合并框的，
    删掉非锚点格会让合并区的右边框、上下框线延伸段消失。
    """
    pat = re.compile(r'<c r="%s"(?![0-9])[^>]*?(?:/>|>.*?</c>)' % re.escape(ref), re.DOTALL)
    new = '<c r="%s" s="%d"/>' % (ref, sidx)
    if pat.search(x):
        return pat.sub(new, x, count=1)
    col = re.match(r'([A-Z]+)', ref).group(1)
    row = int(re.match(r'[A-Z]+(\d+)', ref).group(1))
    rm = re.search(r'(<row r="%d"[^>]*>)(.*?)(</row>)' % row, x, re.DOTALL)
    if not rm:
        return x
    body = rm.group(2)
    tgt = _colnum(col)
    pos = len(body)
    for cm in re.finditer(r'<c\b[^>]*?\br="([A-Z]+)\d+"', body):
        if _colnum(cm.group(1)) > tgt:
            pos = cm.start()
            break
    return x[:rm.start(1)] + rm.group(1) + body[:pos] + new + body[pos:] + '</row>' + x[rm.end():]


def _amount_merge_gi(x, first, last):
    """合同明细行「金额」列统一成 G:I 合并。

    模板第 1、2 行是 G:I 合并，第 3 行起只合了 H:I —— 金额看着被拆成两截、多一条竖线。
    这里把第 2 行之后改成 G:I；H / I 只清内容、留着并沿用 G 的样式，否则合并区边框会缺一段。
    """
    g_sidx = _cell_style(x, 'G%d' % first)
    g_style = int(g_sidx) if g_sidx not in (None, '') else None
    for r in range(first + 1, last + 1):
        x = x.replace('ref="H%d:I%d"' % (r, r), 'ref="G%d:I%d"' % (r, r))
        if g_style is None:
            continue
        x = set_cell_style(x, 'G%d' % r, g_style)
        for col in ('H', 'I'):
            x = _put_blank(x, '%s%d' % (col, r), g_style)
    return x


def rebuild_master_sheet(master):
    """根据用户填好的 HS 对照表(ERP→HS/品名/单位/单价/重量)重建隐藏的「产品主数据」sheet7.xml。
    该表以 HS 为主键(A列)，报关单项表用 INDEX/MATCH($B行, A列) 联动。
    列结构: A=HS编码 B=ERP商品编码 C=品名(中) D=品牌 E=型号 F=单位 G=单价USD H=单件毛重KG I=单件净重KG"""
    bnames = bom_names()
    bmeta = bom_meta()  # erp -> {brand, model}

    rows = []
    for erp, m in master.items():
        hs = m.get('hs', '')
        if not hs:
            continue  # 未填 HS 的不入主数据（生成时留空由用户手填）
        meta = bmeta.get(erp, {})
        name = m.get('name', '') or bnames.get(erp, '')
        brand = m.get('brand', '') or meta.get('brand', '')
        model = m.get('model', '') or meta.get('model', '')
        unit = m.get('unit', '')
        price = m.get('price', '')
        gw = m.get('gw', '')
        nw = m.get('nw', '')
        rows.append([hs, erp, name, brand, model, unit, price, gw, nw])

    # 按 HS 去重，保持顺序
    seen, dedup = set(), []
    for r in rows:
        if r[0] in seen:
            continue
        seen.add(r[0])
        dedup.append(r)

    if not dedup:
        return None, None   # 无任何 HS，保留原主数据表不动

    cols = ['HS编码', 'ERP商品编码', '品名(中)', '品牌', '型号(规格)', '单位', '单价USD', '单件毛重KG', '单件净重KG']
    letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']

    def cx(ref, val):
        if val is None or str(val).strip() == '':
            return ''
        return '<c r="%s" t="str"><v>%s</v></c>' % (ref, esc(val))

    row_xml = ['<row r="1">%s</row>' % ''.join(
        '<c r="%s1" t="str"><v>%s</v></c>' % (letters[i], cols[i]) for i in range(len(cols)))]
    for i, r in enumerate(dedup):
        rn = 2 + i
        row_xml.append('<row r="%d">%s</row>' % (
            rn, ''.join(cx(letters[c] + str(rn), r[c]) for c in range(len(cols)))))

    LAST = 1 + len(dedup)
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
           '<dimension ref="A1:I%d"/>\n<sheetData>\n%s\n</sheetData>\n</worksheet>'
           % (LAST, '\n'.join(row_xml)))
    return xml, LAST


def fix_packing_list(parts, item_count):
    """箱单 sheet4：第10行是第1项，11-18行是第2-9项。
    把原来有 bug 的 OFFSET 公式改成直接引用报关单项行(20+3j)，并按实际项数清空多余行。"""
    x = parts['xl/worksheets/sheet4.xml'].decode('utf-8')
    for j in range(1, 9):            # 箱单第 10+j 行 = 第 j+1 项
        pl_row = 10 + j
        bg_row = 20 + 3 * j          # 报关单对应项行
        if j < item_count:
            x = put_cell(x, 'A%d' % pl_row, 'v', j + 1)
            x = put_cell(x, 'B%d' % pl_row, 'f', '报关单!D%d' % bg_row, is_str=True)
            x = put_cell(x, 'C%d' % pl_row, 'f', '报关单!V%d' % bg_row)
            x = put_cell(x, 'F%d' % pl_row, 'f', '报关单!G%d' % bg_row)
            x = put_cell(x, 'G%d' % pl_row, 'f', '报关单!H%d' % bg_row, is_str=True)
            x = put_cell(x, 'H%d' % pl_row, 'f', '报关单!U%d' % bg_row)
            x = put_cell(x, 'I%d' % pl_row, 'f', '报关单!T%d' % bg_row)
        else:
            # 清空多余项行（保留边框样式，去掉原 OFFSET 公式）
            for ref in ('A%d', 'B%d', 'C%d', 'F%d', 'G%d', 'H%d', 'I%d'):
                x = clear_cell_keep_style(x, ref % pl_row)
    parts['xl/worksheets/sheet4.xml'] = x.encode('utf-8')


def _style_of(xml, col, row):
    m = re.search(r'<c r="%s%d"([^>]*?)(?:/>|>)' % (col, row), xml)
    if not m:
        return None
    sm = re.search(r'\bs="(\d+)"', m.group(1))
    return sm.group(1) if sm else None


# ---------- 项表区统一样式：12号字 + 全框线 + 自动换行 ----------

_ITEM_CELL_RE = re.compile(r'<c\b[^>]*?(?:/>|>.*?</c>)', re.DOTALL)


def _col_letter(ref):
    return re.match(r'([A-Z]+)', ref).group(1)


def _col_num(c):
    n = 0
    for ch in c:
        n = n * 26 + (ord(ch) - 64)
    return n


def _xmap_build(parts, used):
    """为 used 中的原样式派生新样式：字号统一12号 + 自动换行；
    边框/底色/数字格式完全沿用原样式。返回 {old_s: new_s}。"""
    s = parts['xl/styles.xml'].decode('utf-8')
    fm = re.search(r'<fonts count="(\d+)"[^>]*>(.*?)</fonts>', s, re.DOTALL)
    xm = re.search(r'<cellXfs count="(\d+)"[^>]*>(.*?)</cellXfs>', s, re.DOTALL)
    if not fm or not xm:
        return {}
    fonts = re.findall(r'<font>.*?</font>|<font/>', fm.group(2), re.DOTALL)
    nfont = int(fm.group(1))
    cnt = int(xm.group(1))
    xfs = re.findall(r'<xf [^>]*?(?:/>|>.*?</xf>)', xm.group(2), re.DOTALL)

    font_map, new_fonts = {}, []

    def font12(fid):
        if fid in font_map:
            return font_map[fid]
        if fid >= len(fonts):
            font_map[fid] = fid
            return fid
        f = fonts[fid]
        if re.search(r'<sz val="12(\.0)?"', f):
            font_map[fid] = fid
            return fid
        nf = re.sub(r'<sz val="[\d.]+"/>', '<sz val="12"/>', f)
        if '<sz ' not in nf:
            nf = nf.replace('<font>', '<font><sz val="12"/>', 1)
            nf = re.sub(r'<font/>', '<font><sz val="12"/></font>', nf)
        new_fonts.append(nf)
        nid = nfont + len(new_fonts) - 1
        font_map[fid] = nid
        return nid

    xmap, new_xfs = {}, []
    for sid in sorted(used):
        if sid >= len(xfs):
            continue
        old = xfs[sid]
        mf = re.search(r'fontId="(\d+)"', old)
        fid = int(mf.group(1)) if mf else 0
        nfid = font12(fid)
        x = old
        if nfid != fid:
            x = re.sub(r'fontId="\d+"', 'fontId="%d"' % nfid, x)
            if 'applyFont' not in x:
                x = x.replace('<xf ', '<xf applyFont="1" ', 1)
        if '<alignment' in x:
            x = re.sub(r'\s*shrinkToFit="1"', '', x)
            if 'wrapText' not in x:
                x = x.replace('<alignment ', '<alignment wrapText="1" ', 1)
        else:
            if x.rstrip().endswith('/>'):
                x = x.rstrip()[:-2] + '><alignment vertical="center" wrapText="1"/></xf>'
            else:
                x = x.replace('</xf>', '<alignment vertical="center" wrapText="1"/></xf>')
        if 'applyAlignment' not in x:
            x = x.replace('<xf ', '<xf applyAlignment="1" ', 1)
        new_xfs.append(x)
        xmap[sid] = cnt + len(new_xfs) - 1

    if new_fonts:
        s = s.replace('<fonts count="%d"' % nfont,
                      '<fonts count="%d"' % (nfont + len(new_fonts)), 1)
        s = s.replace('</fonts>', ''.join(new_fonts) + '</fonts>', 1)
    if new_xfs:
        s = s.replace('<cellXfs count="%d"' % cnt,
                      '<cellXfs count="%d"' % (cnt + len(new_xfs)), 1)
        s = s.replace('</cellXfs>', ''.join(new_xfs) + '</cellXfs>', 1)
    parts['xl/styles.xml'] = s.encode('utf-8')
    return xmap


def _row_style_ids(row_xml):
    return set(int(m.group(1)) for m in re.finditer(r'\bs="(\d+)"', row_xml))


def restyle_sheet_rows(sheet_xml, row_nums, xmap):
    """把指定行内已有单元格的样式换成派生样式（不改边框/底色/数值，不新增单元格）。"""
    if not xmap:
        return sheet_xml
    rs = set(row_nums)

    def rep(m):
        txt = m.group(0)
        rn = int(re.search(r'<row r="(\d+)"', txt).group(1))
        if rn not in rs:
            return txt

        def crep(cm):
            c = cm.group(0)
            sm = re.search(r'\bs="(\d+)"', c)
            if not sm:
                return c
            old = int(sm.group(1))
            new = xmap.get(old)
            if new is None or new == old:
                return c
            return re.sub(r'\bs="\d+"', 's="%d"' % new, c, count=1)

        return _ITEM_CELL_RE.sub(crep, txt)

    return re.sub(r'<row r="\d+"[^>]*>.*?</row>', rep, sheet_xml, flags=re.DOTALL)



# ---------- 项表区：一个产品 = 一行，只保留上下行线，下面全无行线 ----------

def _split_elems(block, tag):
    return re.findall(r'<%s\b[^>]*?/>|<%s\b[^>]*?>.*?</%s>' % (tag, tag, tag), block, re.DOTALL)


def _edge_re(e):
    return re.compile(r'<%s\b[^>]*?/>|<%s\b[^>]*?>.*?</%s>' % (e, e, e), re.DOTALL)


def _set_edges(border, top='keep', bottom='keep', left='keep', right='keep'):
    """改写一个 <border> 的四条边。how: keep / none / thin 等边框线型。"""
    b = border.strip()
    if b.endswith('/>'):
        b = '<border><left/><right/><top/><bottom/><diagonal/></border>'
    for e, how in (('left', left), ('right', right), ('top', top), ('bottom', bottom)):
        if how == 'keep':
            continue
        new = '<%s/>' % e if how == 'none' else '<%s style="%s"><color auto="1"/></%s>' % (e, how, e)
        rx = _edge_re(e)
        if rx.search(b):
            b = rx.sub(new, b, 1)
        elif '<diagonal' in b:
            b = b.replace('<diagonal', new + '<diagonal', 1)
        else:
            b = b.replace('</border>', new + '</border>', 1)
    return b


def derive_border_styles(parts, jobs):
    """jobs: {'line': {sid,...}, 'blank': {sid,...}}
    line  → 保留左右竖线 + 底部横线（顶部沿用上一行底边，避免双线）
    blank → 四边全去掉
    返回 {mode: {old_sid: new_sid}}"""
    s = parts['xl/styles.xml'].decode('utf-8')
    bm = re.search(r'<borders count="(\d+)"[^>]*>(.*?)</borders>', s, re.DOTALL)
    xm = re.search(r'<cellXfs count="(\d+)"[^>]*>(.*?)</cellXfs>', s, re.DOTALL)
    if not bm or not xm:
        return {}
    borders = _split_elems(bm.group(2), 'border')
    xfs = _split_elems(xm.group(2), 'xf')
    nb0, nx0 = int(bm.group(1)), int(xm.group(1))
    add_b, add_x, cb, cx, out = [], [], {}, {}, {}

    def border_for(bid, mode):
        key = (bid, mode)
        if key in cb:
            return cb[key]
        src = borders[bid] if bid < len(borders) else '<border><left/><right/><top/><bottom/><diagonal/></border>'
        if mode == 'line2':
            new = _set_edges(src, top='none', bottom='thin')
        elif mode in ('line', 'bot'):
            new = _set_edges(src, top='keep', bottom='thin')
        elif mode == 'spacer':
            new = _set_edges(src, top='none', bottom='none')
        else:
            new = _set_edges(src, 'none', 'none', 'none', 'none')
        if new in borders:
            idx = borders.index(new)
        elif new in add_b:
            idx = nb0 + add_b.index(new)
        else:
            add_b.append(new)
            idx = nb0 + len(add_b) - 1
        cb[key] = idx
        return idx

    for mode, sids in jobs.items():
        out[mode] = {}
        for sid in sorted(set(s for s in sids if s is not None)):
            if sid >= len(xfs):
                continue
            x = xfs[sid]
            m = re.search(r'borderId="(\d+)"', x)
            nbid = border_for(int(m.group(1)) if m else 0, mode)
            key = (sid, nbid)
            if key in cx:
                out[mode][sid] = cx[key]
                continue
            nx = re.sub(r'borderId="\d+"', 'borderId="%d"' % nbid, x)
            if 'borderId=' not in nx:
                nx = nx.replace('<xf', '<xf borderId="%d"' % nbid, 1)
            if 'applyBorder' not in nx:
                nx = nx.replace('<xf ', '<xf applyBorder="1" ', 1)
            if nx in xfs:
                nid = xfs.index(nx)
            elif nx in add_x:
                nid = nx0 + add_x.index(nx)
            else:
                add_x.append(nx)
                nid = nx0 + len(add_x) - 1
            cx[key] = nid
            out[mode][sid] = nid

    if add_b:
        s = s.replace('<borders count="%d"' % nb0, '<borders count="%d"' % (nb0 + len(add_b)), 1)
        s = s.replace('</borders>', ''.join(add_b) + '</borders>', 1)
    if add_x:
        s = s.replace('<cellXfs count="%d"' % nx0, '<cellXfs count="%d"' % (nx0 + len(add_x)), 1)
        s = s.replace('</cellXfs>', ''.join(add_x) + '</cellXfs>', 1)
    parts['xl/styles.xml'] = s.encode('utf-8')
    return out


def apply_style_map(xml, row_map):
    """row_map: {行号: {旧样式: 新样式}}"""
    def rep(m):
        txt = m.group(0)
        rn = int(re.search(r'<row r="(\d+)"', txt).group(1))
        mp = row_map.get(rn)
        if not mp:
            return txt

        def crep(cm):
            c = cm.group(0)
            sm = re.search(r'\bs="(\d+)"', c)
            if not sm:
                return c
            new = mp.get(int(sm.group(1)))
            if new is None:
                return c
            return re.sub(r'\bs="\d+"', 's="%d"' % new, c, count=1)

        return _ITEM_CELL_RE.sub(crep, txt)

    return re.sub(r'<row r="\d+"[^>]*>.*?</row>', rep, xml, flags=re.DOTALL)


def set_row_heights(xml, heights):
    """heights: {行号: 磅值 或 None(恢复自适应)}"""
    def rep(m):
        txt = m.group(0)
        rn = int(re.search(r'<row r="(\d+)"', txt).group(1))
        if rn not in heights:
            return txt
        first = re.match(r'<row r="%d"[^>]*>' % rn, txt).group(0)
        a = re.sub(r'\s(ht|customHeight|hidden)="[^"]*"', '', first[4:-1])
        a = re.sub(r'\s*r="\d+"', '', a)
        h = heights[rn]
        nf = ('<row r="%d" ht="%s" customHeight="1"%s>' % (rn, h, a)) if h is not None else ('<row r="%d"%s>' % (rn, a))
        return nf + txt[len(first):]

    return re.sub(r'<row r="\d+"[^>]*>.*?</row>', rep, xml, flags=re.DOTALL)


ITEM_FIRST_ROW = 20      # 项表区第一块起始行
ITEM_BLOCK_STEP = 3      # 每块 3 行（品名行 + 2 行续行）
ITEM_BLOCK_MAX = 46      # 模板固定 46 块 -> 20 ~ 155 行


def item_blocks(xml):
    """报关单项表区每个产品块的起始行。
    按模板固定结构定位（20 起、每 3 行一块），不依赖单元格内容——
    项号写回或清空后按内容识别会漏块。"""
    present = set(int(m.group(1)) for m in re.finditer(r'<row r="(\d+)"', xml))
    return [ITEM_FIRST_ROW + ITEM_BLOCK_STEP * i
            for i in range(ITEM_BLOCK_MAX)
            if (ITEM_FIRST_ROW + ITEM_BLOCK_STEP * i) in present]


def add_bottom_border(parts, sheet_key, rows, col_from='A', col_to='I'):
    """给商品行补下框线（箱单/合同商品行只有左右竖线，缺横线）。
    表内列补"仅底边"样式（缺格子就补空格子保证线连续），表外多余格子一律去框线。"""
    if sheet_key not in parts or not rows:
        return
    lo, hi = _col_num(col_from), _col_num(col_to)
    x = parts[sheet_key].decode('utf-8')
    rows_xml, sids = {}, set()
    for rn in rows:
        m = re.search(r'<row r="%d"[^>]*>.*?</row>' % rn, x, re.DOTALL)
        if m:
            rows_xml[rn] = m.group(0)
            sids |= set(int(v) for v in re.findall(r'\bs="(\d+)"', m.group(0)))
    if not rows_xml:
        return
    sids.add(0)
    maps = derive_border_styles(parts, {'bot': sids, 'blank': sids})
    bot, blank = maps.get('bot', {}), maps.get('blank', {})

    for rn, txt in rows_xml.items():
        cells = {}
        for c in _ITEM_CELL_RE.findall(txt):
            m = re.search(r'<c r="([A-Z]+)%d"' % rn, c)
            if m:
                cells[_col_num(m.group(1))] = c
        out = []
        for cc in range(lo, hi + 1):
            c = cells.pop(cc, None)
            if c is None:
                letter = ''
                n = cc
                while n:
                    n, r = divmod(n - 1, 26)
                    letter = chr(65 + r) + letter
                out.append('<c r="%s%d" s="%d"/>' % (letter, rn, bot.get(0, 0)))
                continue
            sm = re.search(r'\bs="(\d+)"', c)
            old = int(sm.group(1)) if sm else 0
            out.append(re.sub(r'\bs="\d+"', 's="%d"' % bot.get(old, old), c, count=1) if sm
                       else c.replace('<c r=', '<c s="%d" r=' % bot.get(0, 0), 1))
        # 表外的格子：去掉框线
        for cc, c in sorted(cells.items()):
            sm = re.search(r'\bs="(\d+)"', c)
            if sm:
                c = re.sub(r'\bs="\d+"', 's="%d"' % blank.get(int(sm.group(1)), int(sm.group(1))), c, count=1)
            out.append(c)
        head = re.match(r'<row r="%d"[^>]*>' % rn, txt).group(0)
        x = x.replace(txt, head + ''.join(out) + '</row>', 1)
    parts[sheet_key] = x.encode('utf-8')


def blank_table_tail(parts, sheet_key, start_row, min_cols=3):
    """把表格里从 start_row 起、无内容却还留着框线的空行清干净（去框线 + 高度压到 1 磅），
    直到遇到有内容的行或不再像表格行的行为止。"""
    if sheet_key not in parts or start_row < 1:
        return
    x = parts[sheet_key].decode('utf-8')
    rows = {}
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.DOTALL):
        rows[int(m.group(1))] = m.group(2)
    target = []
    rn = start_row
    while rn in rows:
        body = rows[rn]
        cells = re.findall(r'<c r="[A-Z]+\d+"([^>]*?)(?:/>|>)', body)
        bordered = len(cells)
        has_text = bool(re.search(r'<v>\s*[^<\s]', body)) or '<f>' in body
        if has_text or bordered < min_cols:
            break
        target.append(rn)
        rn += 1
    if not target:
        return

    def sids(rows_):
        got = set()
        for r_ in rows_:
            m = re.search(r'<row r="%d"[^>]*>.*?</row>' % r_, x, re.DOTALL)
            if m:
                got |= set(int(v) for v in re.findall(r'\bs="(\d+)"', m.group(0)))
        return got

    maps = derive_border_styles(parts, {'blank': sids(target)})
    row_map = {}
    for r_ in target:
        m = re.search(r'<row r="%d"[^>]*>.*?</row>' % r_, x, re.DOTALL)
        if m:
            row_map[r_] = {int(v): maps['blank'].get(int(v), int(v))
                           for v in re.findall(r'\bs="(\d+)"', m.group(0))}
    x = apply_style_map(x, row_map)
    x = set_row_heights(x, {r_: 1 for r_ in target})
    parts[sheet_key] = x.encode('utf-8')


def format_item_area(parts, item_count):
    """一个产品只占一行：产品行加底边线，续行/空行高度压到 1 磅且无框线。"""
    x = parts['xl/worksheets/sheet1.xml'].decode('utf-8')
    blocks = item_blocks(x)
    if not blocks:
        return
    n = max(item_count, 0)
    line_rows, spacer_rows, blank_rows = [], [], []
    for i, r in enumerate(blocks):
        if i < n:
            line_rows.append(r)
            spacer_rows.extend([r + 1, r + 2])   # 续行：保留竖线，去掉横线
        else:
            blank_rows.extend([r, r + 1, r + 2])

    def sids(rows):
        got = set()
        for rn in rows:
            m = re.search(r'<row r="%d"[^>]*>.*?</row>' % rn, x, re.DOTALL)
            if m:
                got |= set(int(v) for v in re.findall(r'\bs="(\d+)"', m.group(0)))
        return got

    first = line_rows[:1]
    rest = line_rows[1:]
    jobs = {'line': sids(first), 'line2': sids(rest),
            'spacer': sids(spacer_rows), 'blank': sids(blank_rows)}
    maps = derive_border_styles(parts, jobs)
    row_map = {}
    for mode, rows in (('line', first), ('line2', rest),
                       ('spacer', spacer_rows), ('blank', blank_rows)):
        for rn in rows:
            m = re.search(r'<row r="%d"[^>]*>.*?</row>' % rn, x, re.DOTALL)
            if not m:
                continue
            row_map[rn] = {int(v): maps.get(mode, {}).get(int(v), int(v))
                           for v in re.findall(r'\bs="(\d+)"', m.group(0))}
    x = apply_style_map(x, row_map)

    heights = {}
    for r in line_rows:
        m = re.search(r'<c r="D%d"[^>]*>.*?<v>(.*?)</v>' % r, x, re.DOTALL)
        txt = m.group(1) if m else ''
        units = sum(2 if ord(ch) > 127 else 1 for ch in txt)
        lines = max(1, -(-units // 42))          # 品名列 D:F 约合 42 个半角宽
        heights[r] = round(18 + (lines - 1) * 17.5, 1)
    heights.update({r: 0 for r in spacer_rows})
    heights.update({r: 0 for r in blank_rows})
    x = set_row_heights(x, heights)
    parts['xl/worksheets/sheet1.xml'] = x.encode('utf-8')


def clear_cell_keep_style(xml, ref):
    """清空单元格内容但保留样式（边框/底色），避免出现无框白块。"""
    pat = re.compile(r'<c r="%s"([^>]*?)(?:/>|>.*?</c>)' % re.escape(ref), re.DOTALL)
    def rep(m):
        sm = re.search(r'\bs="(\d+)"', m.group(1))
        return '<c r="%s"%s/>' % (ref, (' s="%s"' % sm.group(1)) if sm else '')
    return pat.sub(rep, xml)



def _shift_rows(xml, start_row, delta):
    """把 row 号 >= start_row 的行/单元格/合并区/维度整体下移 delta。"""
    if delta <= 0:
        return xml
    def shift_ref(ref):
        m = re.match(r'([A-Z]+)(\d+)$', ref)
        if not m:
            return ref
        r = int(m.group(2))
        return '%s%d' % (m.group(1), r + delta if r >= start_row else r)
    xml = re.sub(r'<row r="(\d+)"',
                 lambda m: '<row r="%d"' % (int(m.group(1)) + delta if int(m.group(1)) >= start_row else int(m.group(1))),
                 xml)
    xml = re.sub(r'<c r="([A-Z]+)(\d+)"',
                 lambda m: '<c r="%s%d"' % (m.group(1), int(m.group(2)) + delta if int(m.group(2)) >= start_row else int(m.group(2))),
                 xml)
    xml = re.sub(r'<mergeCell ref="([A-Z]+\d+):([A-Z]+\d+)"',
                 lambda m: '<mergeCell ref="%s:%s"' % (shift_ref(m.group(1)), shift_ref(m.group(2))), xml)
    xml = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+\d+)"',
                 lambda m: '<dimension ref="%s:%s"' % (shift_ref(m.group(1)), shift_ref(m.group(2))), xml)
    return xml


def _insert_item_rows(sheet_xml, template_row, n_items, row_builder,
                      template_merges=None, total_update=None):
    """把单项目行模板复制成 n_items 行：**整行克隆（含所有列的样式与边框）**，
    行高改为自适应，避免新行丢框线。"""
    delta = n_items - 1
    if delta <= 0:
        return sheet_xml
    tm = re.search(r'<row r="%d"[^>]*>.*?</row>' % template_row, sheet_xml, re.DOTALL)
    if not tm:
        return sheet_xml
    tpl = tm.group(0)

    # 模板行：列 -> 样式索引
    tpl_cells = {}
    for m in _ITEM_CELL_RE.finditer(tpl):
        rf = re.search(r'r="([A-Z]+)\d+"', m.group(0))
        if rf:
            sm = re.search(r'\bs="(\d+)"', m.group(0))
            tpl_cells[rf.group(1)] = sm.group(1) if sm else None
    am = re.search(r'<row r="%d"([^>]*)>' % template_row, tpl)
    tpl_attrs = re.sub(r'\s(ht|customHeight)="[^"]*"', '', am.group(1) if am else '')

    # 模板行只去掉固定行高（统一行距），内容/公式原样保留
    tpl_new = re.sub(r'<row r="%d"[^>]*>' % template_row,
                     '<row r="%d"%s>' % (template_row, tpl_attrs), tpl, count=1)
    sheet_xml = sheet_xml.replace(tpl, tpl_new, 1)

    # 下移后续行
    sheet_xml = _shift_rows(sheet_xml, template_row + 1, delta)

    new_rows = []
    for k in range(1, delta + 1):
        r = template_row + k
        overrides = {}
        for (col, kind, content, is_str) in row_builder(k, r):
            overrides[col] = (kind, content, is_str)
        cols = sorted(set(tpl_cells) | set(overrides), key=_col_num)
        cx = []
        for col in cols:
            ref = '%s%d' % (col, r)
            s = tpl_cells.get(col)
            sattr = (' s="%s"' % s) if s else ''
            if col in overrides:
                kind, content, is_str = overrides[col]
                tattr = ' t="str"' if is_str else ''
                inner = ('<f>%s</f>' % content) if kind == 'f' else ('<v>%s</v>' % esc(content))
                cx.append('<c r="%s"%s%s>%s</c>' % (ref, sattr, tattr, inner))
            else:
                cx.append('<c r="%s"%s/>' % (ref, sattr))
        new_rows.append('<row r="%d"%s>%s</row>' % (r, tpl_attrs, ''.join(cx)))

    m2 = re.search(r'<row r="%d"[^>]*>.*?</row>' % template_row, sheet_xml, re.DOTALL)
    sheet_xml = sheet_xml[:m2.end()] + ''.join(new_rows) + sheet_xml[m2.end():]
    if template_merges:
        add = []
        for k in range(1, delta + 1):
            r = template_row + k
            for (c1, c2) in template_merges:
                add.append('<mergeCell ref="%s%d:%s%d"/>' % (c1, r, c2, r))
        mc = re.search(r'<mergeCells count="(\d+)"', sheet_xml)
        if mc:
            cnt = int(mc.group(1))
            sheet_xml = re.sub(r'<mergeCells count="\d+">',
                               '<mergeCells count="%d">' % (cnt + len(add)), sheet_xml, count=1)
            sheet_xml = sheet_xml.replace('</mergeCells>', ''.join(add) + '</mergeCells>', 1)
    if total_update:
        sheet_xml = total_update(sheet_xml, n_items, delta)
    return sheet_xml


def extend_subsheets(parts, items, master):
    """把 发票 / 合同 / 申报要素 从单项目扩充为多项目。"""
    n = len(items)
    bmeta = bom_meta()

    # 申报要素第 1 项：模板行原来靠 INDEX 按 HS 取品牌/型号，
    # 若多个 ERP 共用同一 HS 会取错 → 改为按 ERP 静态写入
    if n >= 1:
        erp0 = items[0].get('erp', '')
        mp0 = master.get(erp0, {})
        meta0 = bmeta.get(erp0, {})
        s2 = parts['xl/worksheets/sheet2.xml'].decode('utf-8')
        s2 = put_cell(s2, 'E10', 'v', mp0.get('brand') or meta0.get('brand', ''), is_str=True)
        s2 = put_cell(s2, 'F10', 'v', mp0.get('model') or meta0.get('model', ''), is_str=True)
        parts['xl/worksheets/sheet2.xml'] = s2.encode('utf-8')

    if n <= 1:
        return

    # ---------- 发票 (sheet3) 模板行 10，合并 B:D / G:H，合计行 11 ----------
    x = parts['xl/worksheets/sheet3.xml'].decode('utf-8')
    def fp_builder(k, r):
        bg = 20 + 3 * k
        return [
            ('A', 'v', k + 1, False),
            ('B', 'f', '报关单!D%d' % bg, True),
            ('E', 'f', '报关单!B%d' % bg, True),
            ('F', 'f', '报关单!G%d' % bg, False),
            ('G', 'f', '报关单!I%d' % bg, False),
            ('I', 'f', 'G%d*F%d' % (r, r), False),
        ]
    def fp_total(xml, n, delta):
        tr = 11 + delta
        xml = re.sub(r'<c r="F%d"[^>]*?(?:/>|>.*?</c>)' % tr,
                     '<c r="F%d"><f>SUM(F10:F%d)</f></c>' % (tr, 10 + delta), xml, flags=re.DOTALL)
        xml = re.sub(r'<c r="I%d"[^>]*?(?:/>|>.*?</c>)' % tr,
                     '<c r="I%d"><f>SUM(I10:I%d)</f></c>' % (tr, 10 + delta), xml, flags=re.DOTALL)
        return xml
    x = _insert_item_rows(x, 10, n, fp_builder, template_merges=[('B', 'D'), ('G', 'H')], total_update=fp_total)
    parts['xl/worksheets/sheet3.xml'] = x.encode('utf-8')

    # ---------- 合同 (sheet5) 模板行 18，合并 B:C / G:I ----------
    x = parts['xl/worksheets/sheet5.xml'].decode('utf-8')
    def ht_builder(k, r):
        bg = 20 + 3 * k
        return [
            ('B', 'f', '报关单!D%d' % bg, True),
            ('D', 'f', '报关单!G%d' % bg, False),
            ('E', 'f', '报关单!H%d' % bg, True),
            ('F', 'f', '报关单!I%d' % bg, False),
            ('G', 'f', 'F%d*D%d' % (r, r), False),
        ]
    x = _insert_item_rows(x, 18, n, ht_builder, template_merges=[('B', 'C'), ('G', 'I')])
    parts['xl/worksheets/sheet5.xml'] = x.encode('utf-8')

    # ---------- 申报要素 (sheet2) 模板行 10，无合并 ----------
    x = parts['xl/worksheets/sheet2.xml'].decode('utf-8')
    def sb_builder(k, r):
        bg = 20 + 3 * k
        erp = items[k].get('erp', '')
        meta = bmeta.get(erp, {})
        mp = master.get(erp, {})
        return [
            ('A', 'v', k + 1, False),
            ('B', 'f', '报关单!D%d' % bg, True),
            ('C', 'f', '报关单!B%d' % bg, True),
            ('E', 'v', mp.get('brand') or meta.get('brand', ''), True),
            ('F', 'v', mp.get('model') or meta.get('model', ''), True),
            ('G', 'v', '境内自主品牌', True),
            ('H', 'v', '不能确定在最终目的国（地区）享受优惠关税', True),
            ('I', 'v', '用于扫描', True),
        ]
    x = _insert_item_rows(x, 10, n, sb_builder)
    parts['xl/worksheets/sheet2.xml'] = x.encode('utf-8')


def apply_item_styles(parts, item_count):
    """统一各单据项表区字体为 12 号并开启自动换行；
    边框、底色、数字格式全部沿用模板原样（不新增网格，避免出现多余横线）。"""
    n = max(item_count, 1)
    targets = {
        'xl/worksheets/sheet1.xml': list(range(20, 158)),            # 报关单 46 个产品块 × 3 行
        'xl/worksheets/sheet2.xml': [9] + list(range(10, 10 + n)),  # 申报要素
        'xl/worksheets/sheet3.xml': [9] + list(range(10, 10 + n + 1)),  # 发票(含合计行)
        'xl/worksheets/sheet4.xml': [9] + list(range(10, 19)),      # 箱单 9 个项行
        'xl/worksheets/sheet5.xml': [17] + list(range(18, 18 + n)),  # 合同
    }
    # 先收集这些行实际用到的样式索引
    used = set()
    for f, rows in targets.items():
        if f not in parts:
            continue
        x = parts[f].decode('utf-8')
        for rn in rows:
            m = re.search(r'<row r="%d"[^>]*>.*?</row>' % rn, x, re.DOTALL)
            if m:
                used |= _row_style_ids(m.group(0))
    if not used:
        return
    xmap = _xmap_build(parts, used)
    if not xmap:
        return
    for f, rows in targets.items():
        if f not in parts:
            continue
        xml = restyle_sheet_rows(parts[f].decode('utf-8'), rows, xmap)
        # 项行去掉固定行高，让换行后的内容自动撑开（表头行保留原高度）
        head = {'xl/worksheets/sheet2.xml': 9, 'xl/worksheets/sheet3.xml': 9,
                'xl/worksheets/sheet4.xml': 9, 'xl/worksheets/sheet5.xml': 17}
        body_rows = [r for r in rows if r != head.get(f)]
        parts[f] = _strip_row_height(xml, body_rows).encode('utf-8')


def _strip_row_height(sheet_xml, rows):
    rs = set(rows)

    def rep(m):
        txt = m.group(0)
        rn = int(re.search(r'<row r="(\d+)"', txt).group(1))
        if rn not in rs:
            return txt
        first = re.match(r'<row r="%d"[^>]*>' % rn, txt).group(0)
        newfirst = re.sub(r'\s(ht|customHeight)="[^"]*"', '', first)
        return txt.replace(first, newfirst, 1)

    return re.sub(r'<row r="\d+"[^>]*>.*?</row>', rep, sheet_xml, flags=re.DOTALL)


TEMPLATE_V2 = os.path.join(BASE, 'base_template_v2.xlsx')

# v2 模板坐标：商品明细一行一个 SKU，模板已预置 8 行（不需插行）
BGD_FIRST, BGD_MAX = 20, 8      # 报关单 20~27
FP_FIRST, FP_MAX = 10, 8        # 发票 10~17，合计固定在第 18 行
FP_TOTAL = 18
SB_FIRST, SB_MAX = 10, 8        # 申报要素 10~17
XD_FIRST, XD_MAX = 10, 9        # 箱单 10~18，合计 19
HT_FIRST, HT_MAX = 18, 12       # 合同 18~29，总值 30
HT_TOTAL = 30

def _cfg(key, default):
    """读本地配置 local_config.json（不入库），让品牌名 / 货源地这类公司标识可配置。"""
    try:
        p = os.path.join(BASE, 'local_config.json')
        if os.path.exists(p):
            v = json.load(open(p, encoding='utf-8')).get(key)
            if v:
                return str(v)
    except Exception:
        pass
    return default


SB_BRAND = _cfg('brand', '示例品牌')                     # 品牌：固定（local_config.json 覆盖）
SB_BRAND_TYPE = '境内自主品牌'                        # 品牌类型：固定
SB_PREFER = '不能确定在最终目的国（地区）享受优惠关税'   # 出口享惠情况：固定
BGD_ORIGIN = _cfg('origin', '中国深圳')                # 原产国 / 境内货源地
BGD_DUTY = '照章征税'                                # 征免


def _blank_row(x, r, cols):
    """只清内容、保留样式（不动字体/行高/边框/合并）。"""
    for c in cols:
        x = clear_cell_keep_style(x, '%s%d' % (c, r))
    return x


def _qty(v):
    """数量取整，避免出现 1.0 这种显示。"""
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return v


def _unit_of(it, master):
    """单位以 HS 对照表(主数据)为准，表里没填才用 PDF 上的（PDF 常是 PCS 这类英文单位）。"""
    mp = master.get(it.get('erp'), {})
    return mp.get('unit') or it.get('unit') or ''


def _fit_row_height(x, r, text, avail, font_sz):
    """按品名长度估算需要的行高，只上调不下调。

    模板里品名格都开了自动换行，但**合并单元格在 Excel/WPS 里不会自动撑行高**，
    长品名会被裁掉半行，所以按可用宽度补一个够用的高度。
    avail = 该品名所占列区的宽度合计（Excel 列宽单位，1 个汉字约等于 2）。
    """
    units = sum(2 if ord(ch) > 127 else 1 for ch in str(text or ''))
    lines = max(1, -(-units // max(int(avail), 1)))
    need = round(lines * font_sz * 1.45 + 3, 1)
    m = re.search(r'<row r="%d"([^>]*)>' % r, x)
    if not m:
        return x
    attr = m.group(1)
    cur = re.search(r'\bht="([\d.]+)"', attr)
    if cur and float(cur.group(1)) >= need:
        return x
    if not cur and lines <= 1:
        return x          # 单行且模板是自动行高：完全不动
    newattr = re.sub(r'\s(?:ht|customHeight)="[^"]*"', '', attr)
    newattr += ' ht="%s" customHeight="1"' % need
    return x[:m.start(1)] + newattr + x[m.end(1):]


def build_docs(hdr, items, master, output_path):
    """v2 模板（8 行 SKU 预置、已取消框线）：只写值和公式，不碰任何样式。

    固定规则：
      · 报关单「最终目的国」绑定「贸易国」→ 写公式 =E10，改贸易国自动跟随；
      · 申报要素品牌固定（见 SB_BRAND）、「品牌类型」固定「境内自主品牌」；
      · 报关单原产国 / 境内货源地固定（见 BGD_ORIGIN），征免固定「照章征税」；
      · 未用到的 SKU 行只清内容，样式 / 行高 / 列宽 / 边框 / 合并 / logo 一律不动。
    """
    z = zipfile.ZipFile(TEMPLATE_V2)
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()

    # 打开即重算，避免公式沿用上一单的缓存值（合计 / 总值 / 发票号）
    w = parts['xl/workbook.xml'].decode('utf-8')
    if 'fullCalcOnLoad' not in w:
        w = w.replace('<calcPr ', '<calcPr fullCalcOnLoad="1" ', 1)
        parts['xl/workbook.xml'] = w.encode('utf-8')

    so = (hdr.get('报价号码') or '').strip()
    cust = (hdr.get('客户') or '').strip()
    date = (hdr.get('报价日期') or '').strip().replace('-', '/')
    bmeta = bom_meta()
    n = min(len(items), BGD_MAX)
    written = []

    # ---------- 报关单 ----------
    x = parts['xl/worksheets/sheet1.xml'].decode('utf-8')
    if so:
        x = put_cell(x, 'A10', 'v', so, is_str=True)     # 合同协议号
    if cust:
        x = put_cell(x, 'A6', 'v', cust, is_str=True)    # 境外收货人
    for i in range(BGD_MAX):
        r = BGD_FIRST + i
        if i >= n:
            x = _blank_row(x, r, ('A', 'B', 'D', 'G', 'H', 'I', 'K', 'M', 'P', 'S'))
            continue
        it = items[i]
        x = put_cell(x, 'A%d' % r, 'v', i + 1)
        x = _put_hs(x, 'B%d' % r, it.get('hs'))
        x = put_cell(x, 'D%d' % r, 'v', it.get('name', ''), is_str=True)
        x = _fit_row_height(x, r, it.get('name'), 42, 12)   # D:F ≈ 44 宽（留 2 单位余量）
        x = put_cell(x, 'G%d' % r, 'v', _qty(it.get('qty', 1)))
        x = put_cell(x, 'H%d' % r, 'v', _unit_of(it, master), is_str=True)
        if it.get('price') not in (None, ''):
            x = put_cell(x, 'I%d' % r, 'v', it['price'])
        x = put_cell(x, 'K%d' % r, 'v', BGD_ORIGIN, is_str=True)   # 原产国
        x = put_cell(x, 'M%d' % r, 'f', 'E10')                     # 最终目的国 = 贸易国
        x = put_cell(x, 'P%d' % r, 'v', BGD_ORIGIN, is_str=True)   # 境内货源地
        x = put_cell(x, 'S%d' % r, 'v', BGD_DUTY, is_str=True)     # 征免
        written.append({'no': i + 1, 'erp': it.get('erp', ''), 'hs': it.get('hs', ''),
                        'name': it.get('name', ''), 'qty': it.get('qty', 1),
                        'price': it.get('price')})
    parts['xl/worksheets/sheet1.xml'] = x.encode('utf-8')

    # ---------- 发票 ----------
    f = parts['xl/worksheets/sheet3.xml'].decode('utf-8')
    if so:
        f = put_cell(f, 'I8', 'v', so, is_str=True)      # Reference
        f = put_cell(f, 'I4', 'v', '', is_str=True)      # Invoice No. 留空手填
    if date:
        f = put_cell(f, 'I5', 'v', date, is_str=True)
    for i in range(FP_MAX):
        r = FP_FIRST + i
        if i >= n:
            f = _blank_row(f, r, ('A', 'B', 'E', 'F', 'G', 'I'))
            continue
        it = items[i]
        f = put_cell(f, 'A%d' % r, 'v', i + 1)
        f = put_cell(f, 'B%d' % r, 'v', it.get('name', ''), is_str=True)
        f = _fit_row_height(f, r, it.get('name'), 28, 9)    # B:D ≈ 30 宽（留余量）
        f = _put_hs(f, 'E%d' % r, it.get('hs'))
        f = put_cell(f, 'F%d' % r, 'v', _qty(it.get('qty', 1)))
        f = put_cell(f, 'G%d' % r, 'v',
                     it.get('price') if it.get('price') not in (None, '') else 0)
        f = put_cell(f, 'I%d' % r, 'f', 'G%d*F%d' % (r, r))
    f = put_cell(f, 'F%d' % FP_TOTAL, 'f', 'SUM(F%d:F%d)' % (FP_FIRST, FP_FIRST + FP_MAX - 1))
    f = put_cell(f, 'I%d' % FP_TOTAL, 'f', 'SUM(I%d:I%d)' % (FP_FIRST, FP_FIRST + FP_MAX - 1))
    parts['xl/worksheets/sheet3.xml'] = f.encode('utf-8')

    # ---------- 申报要素 ----------
    s2 = parts['xl/worksheets/sheet2.xml'].decode('utf-8')
    for i in range(SB_MAX):
        r = SB_FIRST + i
        if i >= n:
            s2 = _blank_row(s2, r, ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'))
            continue
        it = items[i]
        mp = master.get(it.get('erp'), {})
        meta = bmeta.get(it.get('erp'), {})
        s2 = put_cell(s2, 'A%d' % r, 'v', i + 1)
        s2 = put_cell(s2, 'B%d' % r, 'v', it.get('name', ''), is_str=True)
        s2 = _fit_row_height(s2, r, it.get('name'), 30, 12)   # B 列 ≈ 32 宽（留余量）
        s2 = _put_hs(s2, 'C%d' % r, it.get('hs'))
        s2 = put_cell(s2, 'E%d' % r, 'v', SB_BRAND, is_str=True)          # 品牌固定
        s2 = put_cell(s2, 'F%d' % r, 'v',
                      mp.get('model') or meta.get('model') or '', is_str=True)
        s2 = put_cell(s2, 'G%d' % r, 'v', SB_BRAND_TYPE, is_str=True)     # 品牌类型固定
        s2 = put_cell(s2, 'H%d' % r, 'v', SB_PREFER, is_str=True)
    parts['xl/worksheets/sheet2.xml'] = s2.encode('utf-8')

    # ---------- 箱单 ----------
    p = parts['xl/worksheets/sheet4.xml'].decode('utf-8')
    if date:
        p = put_cell(p, 'I3', 'v', date, is_str=True)
    if so:
        p = put_cell(p, 'I5', 'v', so, is_str=True)
    if cust:
        p = put_cell(p, 'B5', 'v', cust, is_str=True)
    for i in range(XD_MAX):
        r = XD_FIRST + i
        if i >= n:
            p = _blank_row(p, r, ('A', 'B', 'C', 'F', 'G', 'H', 'I'))
            continue
        it = items[i]
        mp = master.get(it.get('erp'), {})
        p = put_cell(p, 'A%d' % r, 'v', i + 1)
        p = put_cell(p, 'B%d' % r, 'v', it.get('name', ''), is_str=True)
        p = _fit_row_height(p, r, it.get('name'), 23, 11)   # B 列 ≈ 25 宽（留余量）
        p = _blank_row(p, r, ('C', 'H', 'I'))     # 原有 OFFSET 公式按旧版排版写的，清掉改直填
        p = put_cell(p, 'C%d' % r, 'v', 1)        # 箱数默认 1 箱/项，可手改
        p = put_cell(p, 'F%d' % r, 'v', _qty(it.get('qty', 1)))
        p = put_cell(p, 'G%d' % r, 'v', _unit_of(it, master), is_str=True)
        try:
            gw = float(mp.get('gw') or 0) * float(it.get('qty') or 0)
            nw = float(mp.get('nw') or 0) * float(it.get('qty') or 0)
        except (TypeError, ValueError):
            gw = nw = 0
        if gw > 0:
            p = put_cell(p, 'H%d' % r, 'v', round(gw, 4))
        if nw > 0:
            p = put_cell(p, 'I%d' % r, 'v', round(nw, 4))
    parts['xl/worksheets/sheet4.xml'] = p.encode('utf-8')

    # ---------- 合同 ----------
    c = parts['xl/worksheets/sheet5.xml'].decode('utf-8')
    if so:
        c = put_cell(c, 'G6', 'v', so, is_str=True)
    if date:
        c = put_cell(c, 'G8', 'v', date, is_str=True)
    for i in range(HT_MAX):
        r = HT_FIRST + i
        if i >= n:
            c = _blank_row(c, r, ('B', 'D', 'E', 'F', 'G'))
            continue
        it = items[i]
        c = put_cell(c, 'B%d' % r, 'v', it.get('name', ''), is_str=True)
        c = _fit_row_height(c, r, it.get('name'), 28, 11)   # B:C ≈ 30 宽（留余量）
        c = put_cell(c, 'D%d' % r, 'v', _qty(it.get('qty', 1)))
        c = put_cell(c, 'E%d' % r, 'v', _unit_of(it, master), is_str=True)
        if it.get('price') not in (None, ''):
            c = put_cell(c, 'F%d' % r, 'v', it['price'])
        c = put_cell(c, 'G%d' % r, 'f', 'F%d*D%d' % (r, r))
    # 第 3 行起「金额」列被拆成 G + H:I 两段，统一成 G:I（与第 1、2 行一致）
    c = _amount_merge_gi(c, HT_FIRST, HT_FIRST + HT_MAX - 1)
    parts['xl/worksheets/sheet5.xml'] = c.encode('utf-8')

    # ---------- 统一字体 / 字号 / 货币 / 居中 ----------
    # 明细区列清单：模板第 1 行是对的，第 2 行起字体字号、数字格式各不一样，全部对齐第 1 行
    detail = {
        'xl/worksheets/sheet1.xml': (('A', 'B', 'D', 'G', 'H', 'I', 'K', 'M', 'P', 'S',
                                      'T', 'U', 'V'),
                                     BGD_FIRST, BGD_FIRST + BGD_MAX - 1),
        'xl/worksheets/sheet2.xml': (('A', 'B', 'C', 'E', 'F', 'G', 'H'),
                                     SB_FIRST, SB_FIRST + SB_MAX - 1),
        'xl/worksheets/sheet3.xml': (('A', 'B', 'E', 'F', 'G', 'I'),
                                     FP_FIRST, FP_FIRST + FP_MAX - 1),
        'xl/worksheets/sheet4.xml': (('A', 'B', 'C', 'F', 'G', 'H', 'I'),
                                     XD_FIRST, XD_FIRST + XD_MAX - 1),
        'xl/worksheets/sheet5.xml': (('B', 'D', 'E', 'F', 'G'),
                                     HT_FIRST, HT_FIRST + HT_MAX - 1),
    }
    # 单价 / 金额：报关单 I、发票 G+I+I18、合同 F+G+G30
    money_refs = {
        'xl/worksheets/sheet1.xml': ['I%d' % r for r in range(BGD_FIRST, BGD_FIRST + BGD_MAX)],
        'xl/worksheets/sheet3.xml': (['G%d' % r for r in range(FP_FIRST, FP_FIRST + FP_MAX)]
                                     + ['I%d' % r for r in range(FP_FIRST, FP_FIRST + FP_MAX)]
                                     + ['I%d' % FP_TOTAL]),
        'xl/worksheets/sheet5.xml': (['F%d' % r for r in range(HT_FIRST, HT_FIRST + HT_MAX)]
                                     + ['G%d' % r for r in range(HT_FIRST, HT_FIRST + HT_MAX)]
                                     + ['G%d' % HT_TOTAL]),
    }

    forge = StyleForge(parts['xl/styles.xml'].decode('utf-8'))

    # 1) 明细区：字体 / 字号对齐第 1 行；水平 + 垂直都居中
    for name, (cols, first, last) in detail.items():
        src = parts[name].decode('utf-8')
        for col in cols:
            ref_style = _cell_style(src, '%s%d' % (col, first))
            if ref_style is None:
                continue
            fid = forge.font_of(int(ref_style))
            nf = forge.numfmt_of(int(ref_style))
            for r in range(first, last + 1):
                cur = _cell_style(src, '%s%d' % (col, r))
                if cur is None:
                    continue
                src = set_cell_style(src, '%s%d' % (col, r),
                                     forge.variant(int(cur), font_id=fid, numfmt_id=nf,
                                                   halign='center', valign='center'))
        parts[name] = src.encode('utf-8')

    # 2) 单价 / 金额：货币符号 + 居中
    nid = forge.ensure_numfmt(MONEY_FMT_XML)
    for name, refs in money_refs.items():
        src = parts[name].decode('utf-8')
        for ref in refs:
            cur = _cell_style(src, ref)
            if cur is None:
                continue
            src = set_cell_style(src, ref,
                                 forge.variant(int(cur), numfmt_id=nid,
                                               halign='center', valign='center'))
        parts[name] = src.encode('utf-8')

    # 3) 合计行也居中
    for name, refs in {'xl/worksheets/sheet3.xml': ['A%d' % FP_TOTAL, 'F%d' % FP_TOTAL, 'I%d' % FP_TOTAL],
                       'xl/worksheets/sheet5.xml': ['F%d' % HT_TOTAL, 'G%d' % HT_TOTAL]}.items():
        src = parts[name].decode('utf-8')
        for ref in refs:
            cur = _cell_style(src, ref)
            if cur is None:
                continue
            src = set_cell_style(src, ref, forge.variant(int(cur), halign='center', valign='center'))
        parts[name] = src.encode('utf-8')

    parts['xl/styles.xml'] = forge.dump().encode('utf-8')

    # ---------- 写出（文件被占用时退让命名） ----------
    base, ext = os.path.splitext(output_path)
    for cand in (output_path, '%s(新)%s' % (base, ext)):
        try:
            with zipfile.ZipFile(cand, 'w', zipfile.ZIP_DEFLATED) as zo:
                for name, data in parts.items():
                    zo.writestr(name, data)
            return written, cand
        except (PermissionError, OSError):
            continue
    cand = '%s_%s%s' % (base, datetime.datetime.now().strftime('%H%M%S'), ext)
    with zipfile.ZipFile(cand, 'w', zipfile.ZIP_DEFLATED) as zo:
        for name, data in parts.items():
            zo.writestr(name, data)
    return written, cand


def build_baoguandan(hdr, items, master, output_path):
    """把解析结果写入报关单。items 已带 hs(可能为空)。"""
    z = zipfile.ZipFile(TEMPLATE)
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()

    # 重建隐藏主数据表（HS为主键），使 INDEX/MATCH 联动正确
    master_xml, LAST = rebuild_master_sheet(master)
    x = parts['xl/worksheets/sheet1.xml'].decode('utf-8')
    if master_xml:
        parts['xl/worksheets/sheet7.xml'] = master_xml.encode('utf-8')
        # 修正公式里的范围：$A$2:$A$5 → 动态 $A$2:$A$LAST（报关单 + 申报要素）
        pat = r"产品主数据'!\$([A-I])\$2:\$[A-I]\$\d+"
        repl = lambda m: "产品主数据'!$%s$2:$%s$%d" % (m.group(1), m.group(1), LAST)
        x = re.sub(pat, repl, x)
        s2 = parts['xl/worksheets/sheet2.xml'].decode('utf-8')
        parts['xl/worksheets/sheet2.xml'] = re.sub(pat, repl, s2).encode('utf-8')

    # 表头：合同协议号、境外收货人、币别
    if hdr.get('报价号码'):
        x = put_cell(x, 'A10', 'v', hdr['报价号码'], is_str=True)
    if hdr.get('客户'):
        x = put_cell(x, 'A6', 'v', hdr['客户'], is_str=True)   # 境外收货人
        x = put_cell(x, 'V8', 'v', hdr['客户'], is_str=True)

    item_rows = [20 + 3 * i for i in range(40)]
    written = []
    for i, it in enumerate(items):
        if i >= len(item_rows):
            break
        r = item_rows[i]
        hs = it.get('hs', '')
        mp = master.get(it['erp'], {})
        x = put_cell(x, 'B%d' % r, 'v', hs, is_str=True)           # 商品编号(HS)，空则留待手填
        x = put_cell(x, 'D%d' % r, 'v', it.get('name', ''), is_str=True)  # 品名(静态，保证无HS也有品名)
        qty = it.get('qty', 1)
        x = put_cell(x, 'G%d' % r, 'v', qty)                        # 数量
        x = put_cell(x, 'H%d' % r, 'v', it.get('unit') or mp.get('unit', ''), is_str=True)  # 单位
        price = it.get('price')
        if price is not None and price != '':
            x = put_cell(x, 'I%d' % r, 'v', price)                  # 单价(订单价)
        x = put_cell(x, 'V%d' % r, 'v', qty)                        # 件数=数量
        # 毛/净重：主数据有单件重量则 qty×单件；否则留空
        try:
            gw = float(mp.get('gw') or 0)
            nw = float(mp.get('nw') or 0)
        except (TypeError, ValueError):
            gw = nw = 0
        if gw > 0:
            x = put_cell(x, 'U%d' % r, 'v', round(gw * qty, 4))
        if nw > 0:
            x = put_cell(x, 'T%d' % r, 'v', round(nw * qty, 4))
        written.append({'no': it['no'], 'erp': it['erp'], 'hs': hs, 'qty': qty,
                        'name': it.get('name', ''), 'price': price})

    # 清空未使用的项行（模板共 46 块，全部扫一遍）：清掉预置的
    # 项号 / 中国 / 中国香港 / 中国深圳 / 照章征税 等常量，避免空行显乱
    blocks = item_blocks(x)
    for i, b in enumerate(blocks):
        if i < len(items):
            x = put_cell(x, 'A%d' % b, 'v', i + 1)      # 项号按实际顺序重写
        else:
            for rr in (b, b + 1, b + 2):
                for col in ('A', 'B', 'D', 'G', 'H', 'I', 'K', 'M', 'P', 'S', 'T', 'U', 'V'):
                    x = clear_cell_keep_style(x, '%s%d' % (col, rr))

    parts['xl/worksheets/sheet1.xml'] = x.encode('utf-8')

    # 子表扩展为多项目（发票/合同/申报要素）
    extend_subsheets(parts, items, master)

    # 箱单多行映射（按实际项数）
    fix_packing_list(parts, len(items))

    # 统一项表区样式：12号字 + 全框线 + 自动换行 + 行高自适应
    apply_item_styles(parts, len(items))

    # 项表区收口：一个产品一行、只留上下行线，未用行高度压平且无框线
    format_item_area(parts, len(items))

    # 箱单 / 合同：商品行补下框线（模板原本只有竖线），下面空表行去掉框线并压平
    add_bottom_border(parts, 'xl/worksheets/sheet4.xml', list(range(10, 10 + len(items))), 'A', 'I')
    add_bottom_border(parts, 'xl/worksheets/sheet5.xml', list(range(18, 18 + len(items))), 'B', 'I')
    blank_table_tail(parts, 'xl/worksheets/sheet4.xml', 10 + len(items))
    blank_table_tail(parts, 'xl/worksheets/sheet5.xml', 18 + len(items))

    # 直接覆盖写；文件正被 Excel/WPS 打开时，退到 …(新).xlsx，再不行加时间戳
    base, ext = os.path.splitext(output_path)
    cands = [output_path, '%s(新)%s' % (base, ext)]
    last = None
    for c in cands:
        try:
            with zipfile.ZipFile(c, 'w', zipfile.ZIP_DEFLATED) as zo:
                for n, data in parts.items():
                    zo.writestr(n, data)
            return written, c
        except (PermissionError, OSError) as e:
            last = e
    c = '%s_%s%s' % (base, datetime.datetime.now().strftime('%H%M%S'), ext)
    with zipfile.ZipFile(c, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n, data in parts.items():
            zo.writestr(n, data)
    return written, c
# ============ 路由 ============

PAGE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>报关单自动生成器</title>
<style>
* { box-sizing: border-box; }
body { font-family: -apple-system, "Microsoft YaHei", sans-serif; margin:0; background:#f5f6f8; color:#222; }
.wrap { max-width: 920px; margin: 0 auto; padding: 24px 20px 60px; }
h1 { font-size: 22px; margin: 0 0 4px; }
.sub { color:#888; font-size: 13px; margin-bottom: 20px; }
.card { background:#fff; border-radius:12px; padding:20px; margin-bottom:18px; box-shadow:0 1px 4px rgba(0,0,0,.06); }
.card h2 { font-size:16px; margin:0 0 12px; }
.btn { display:inline-block; padding:9px 18px; border-radius:8px; border:none; cursor:pointer; font-size:14px; }
.btn-primary { background:#2f5597; color:#fff; }
.btn-primary:disabled { background:#a9b8d0; cursor:not-allowed; }
.btn-ghost { background:#eef1f6; color:#333; }
.drop { border:2px dashed #c8cfdb; border-radius:10px; padding:32px; text-align:center; color:#888; cursor:pointer; transition:.15s; }
.drop.hover { border-color:#2f5597; background:#f0f4fb; }
.drop.done { border-style:solid; border-color:#34a853; background:#f2fbf4; color:#333; }
.muted { color:#999; font-size:12px; }
table { width:100%; border-collapse:collapse; font-size:13px; margin-top:8px; }
th,td { border:1px solid #e3e7ee; padding:7px 9px; text-align:left; }
th { background:#f0f3f8; }
.ok { color:#34a853; font-weight:600; }
.miss { color:#e8710a; font-weight:600; }
.hidden { display:none; }
.row { display:flex; gap:12px; align-items:center; flex-wrap:wrap; }
.grow { flex:1; }
.mt8 { margin-top:8px; }
.pill { display:inline-block; background:#eef1f6; border-radius:20px; padding:2px 10px; font-size:12px; color:#555; margin-right:6px; }
.step { color:#2f5597; font-weight:700; }
</style>
</head>
<body>
<div class="wrap">
  <h1>报关单自动生成器</h1>
  <div class="sub">深圳销售订单(SO) PDF → 自动识别商品编码 → 生成报关单（子表自动联动）</div>

  <div class="card">
    <h2><span class="step">第 0 步</span> · 下载/上传 HS 对照表</h2>
    <p class="muted">先下载对照表，填入每个「ERP商品编码」对应的 <b>海关 HS 编码</b>（没填的=特殊品，生成时留空你手动补）。填好后上传回来，程序据此自动补 HS。</p>
    <div class="row">
      <a class="btn btn-ghost" href="/master/export" download>⬇ 下载 HS 对照表</a>
      <label class="btn btn-ghost" style="position:relative;overflow:hidden;">
        上传已填好的对照表
        <input type="file" id="masterFile" accept=".xlsx" style="position:absolute;left:0;top:0;opacity:0;width:100%;height:100%;cursor:pointer;">
      </label>
      <span id="masterStatus" class="muted">未上传（将用默认模板，HS 全空）</span>
    </div>
  </div>

  <div class="card">
    <h2><span class="step">第 1 步</span> · 上传深圳销售订单 PDF</h2>
    <div class="drop" id="drop">点击或拖拽 PDF 到这里<br><span class="muted">支持多张销售订单(SO) PDF</span></div>
    <input type="file" id="fileInput" accept=".pdf" multiple class="hidden">
    <div id="fileList" class="mt8"></div>
  </div>

  <div class="card hidden" id="resultCard">
    <h2><span class="step">第 2 步</span> · 生成结果</h2>
    <div id="resultBody"></div>
  </div>
</div>

<script>
const drop=document.getElementById('drop'), fileInput=document.getElementById('fileInput');
drop.onclick=()=>fileInput.click();
drop.ondragover=e=>{e.preventDefault();drop.classList.add('hover');};
drop.ondragleave=()=>drop.classList.remove('hover');
drop.ondrop=e=>{e.preventDefault();drop.classList.remove('hover');handleFiles(e.dataTransfer.files);};
fileInput.onchange=e=>handleFiles(e.target.files);

let masterReady=false;

document.getElementById('masterFile').onchange=async e=>{
  const f=e.target.files[0]; if(!f)return;
  const fd=new FormData(); fd.append('file',f);
  const r=await fetch('/master/upload',{method:'POST',body:fd}).then(x=>x.json());
  const s=document.getElementById('masterStatus');
  if(r.ok){ masterReady=true; s.innerHTML='<span class="ok">已加载 '+r.count+' 个商品的 HS 映射</span>'; }
  else{ s.innerHTML='<span class="miss">上传失败：'+r.error+'</span>'; }
};

async function handleFiles(files){
  const list=document.getElementById('fileList');
  const filesArr=[...files].filter(f=>f.name.toLowerCase().endsWith('.pdf'));
  if(!filesArr.length){ list.innerHTML='<span class="miss">未选择 PDF 文件</span>'; return; }
  drop.classList.add('done');
  const t0=Date.now();
  const results=[];
  for(let i=0;i<filesArr.length;i++){
    const f=filesArr[i];
    list.innerHTML='<div class="muted">⏳ 正在解析 '+(i+1)+'/'+filesArr.length+'：'+escAttr(f.name)+'…</div>';
    try{
      const fd=new FormData(); fd.append('file',f);
      const r=await fetch('/parse',{method:'POST',body:fd}).then(x=>x.json());
      results.push({name:f.name, data:r});
    }catch(e){
      results.push({name:f.name, data:{ok:false, error:String(e&&e.message||e)}});
    }
  }
  const okN=results.filter(r=>r.data&&r.data.ok).length;
  const badN=results.length-okN;
  const sec=((Date.now()-t0)/1000).toFixed(1);
  let s='';
  if(okN) s+='<span class="ok">✅ 解析完毕 '+okN+' 个文件</span>';
  if(badN) s+=(s?'　|　':'')+'<span class="miss">❌ '+badN+' 个解析失败</span>';
  s+='　|　<span class="muted">耗时 '+sec+' 秒</span>';
  list.innerHTML=s;
  renderResults(results);
}

function renderResults(results){
  const card=document.getElementById('resultCard');
  const body=document.getElementById('resultBody');
  card.classList.remove('hidden');
  let html='';
  for(const r of results){
    const d=r.data;
    if(!d.ok){ html+='<p class="miss">'+r.name+'：解析失败 '+d.error+'</p>'; continue; }
    html+='<div class="card"><div style="font-weight:600;margin-bottom:6px;">📄 '+r.name+'</div>';
    html+='<div class="muted">客户：'+ (d.hdr['客户']||'-') +'　|　报价号：'+ (d.hdr['报价号码']||'-') +'　|　币别：'+ (d.hdr['币别']||'-') +'</div>';
    html+='<table><tr><th>序号</th><th>ERP编码</th><th>品名</th><th>数量</th><th>单价</th><th>HS编码</th></tr>';
    for(const it of d.items){
      const hs=it.hs?('<span class="ok">'+it.hs+'</span>'):('<span class="miss">待填</span>');
      html+='<tr><td>'+it.no+'</td><td>'+it.erp+'</td><td>'+it.name+'</td><td>'+it.qty+'</td><td>'+it.price+'</td><td>'+hs+'</td></tr>';
    }
    html+='</table>';
    html+='<div class="row mt8"><button class="btn btn-primary" data-token="'+d.token+'" data-name="'+escAttr(r.name)+'" onclick="generate(this.dataset.token,this.dataset.name)">生成报关单</button></div>';
    html+='<div id="out_'+d.token+'" class="mt8"></div>';
    html+='</div>';
  }
  body.innerHTML=html;
}

function escAttr(s){
  return String(s==null?'':s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

async function generate(token, name){
  const box=document.getElementById('out_'+token);
  if(box) box.innerHTML='<span class="muted">正在生成…</span>';
  try{
    const fd=new FormData(); fd.append('token',token);
    const r=await fetch('/generate',{method:'POST',body:fd});
    const ct=r.headers.get('Content-Type')||'';
    if(!r.ok || ct.indexOf('json')<0){
      const t=await r.text();
      throw new Error('HTTP '+r.status+' '+t.slice(0,200));
    }
    const d=await r.json();
    if(!box) return;
    // 下载区：主按钮用普通 <a download>，即使 iframe 拦截自动下载也可手动点
    let h='<div style="border:1px solid #cfe3cf;background:#f4faf4;border-radius:8px;padding:12px;">';
    h+='<div style="font-weight:600;color:#1a7f3c;margin-bottom:8px;">✅ 已生成：'+d.file+'</div>';
    if(d.warning){ h+='<div style="background:#fff7e6;border:1px solid #ffd591;color:#a15c00;border-radius:6px;padding:8px 10px;margin-bottom:8px;font-size:13px;">⚠ '+d.warning+'</div>'; }
    h+='<a href="'+d.download+'" download style="display:inline-block;background:#2f6fed;color:#fff;text-decoration:none;padding:9px 18px;border-radius:6px;font-weight:600;margin-right:8px;">⬇ 下载报关单</a>';
    h+='<a href="'+d.download+'?inline=1" target="_blank" style="display:inline-block;background:#eef2ff;color:#2f6fed;text-decoration:none;padding:9px 16px;border-radius:6px;margin-right:8px;">在新标签页打开</a>';
    h+='<button class="btn" onclick="openOutput()">打开输出文件夹</button>';
    h+='<div class="muted" style="margin-top:8px;font-size:12px;">文件位置：'+d.path+'</div>';
    h+='<div class="muted" style="margin-top:4px;font-size:12px;">如果点「下载报关单」没反应，说明当前是内嵌预览窗口（浏览器禁止其中下载）。请点「打开输出文件夹」直接取文件，或用系统浏览器访问 http://127.0.0.1:8765</div>';
    h+='</div>';
    box.innerHTML=h;
    // 仍尝试自动下载（失败不影响上面的按钮）
    try{
      const a=document.createElement('a');
      a.href=d.download; a.download=d.file; a.style.display='none';
      document.body.appendChild(a); a.click();
      setTimeout(function(){ a.remove(); }, 3000);
    }catch(e){}
  }catch(err){
    if(box) box.innerHTML='<span class="miss">生成失败：'+err.message+'</span>';
    else alert('生成失败：'+err.message);
  }
}

async function openOutput(){
  try{
    const r=await fetch('/open_output');
    const d=await r.json();
    if(!d.ok) alert('无法自动打开，请手动前往：'+d.path);
  }catch(e){ alert('无法自动打开输出文件夹'); }
}
</script>
</body>
</html>"""

@app.route('/')
def index():
    return render_template_string(PAGE)

@app.route('/master/export')
def master_export():
    p = MASTER_DOWNLOAD
    if not os.path.exists(p):
        return jsonify({'ok': False, 'error': '对照表文件不存在'}), 404
    return send_file(p, as_attachment=True, download_name='HS对照表_待填.xlsx')

@app.route('/master/upload', methods=['POST'])
def master_upload():
    f = request.files.get('file')
    if not f:
        return jsonify({'ok': False, 'error': '无文件'})
    src = os.path.join(UPLOAD, 'master_upload.xlsx')
    f.save(src)
    try:
        shutil.copy2(src, MASTER_XLSX)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})
    m = load_master()
    filled = sum(1 for v in m.values() if v['hs'])
    return jsonify({'ok': True, 'count': len(m), 'filled': filled})

@app.route('/parse', methods=['POST'])
def parse():
    f = request.files.get('file')
    if not f:
        return jsonify({'ok': False, 'error': '无文件'})
    name = f.filename
    pdf = os.path.join(UPLOAD, 'latest.pdf')
    f.save(pdf)
    try:
        md = extract_md(pdf)
        hdr, items = parse_order(md)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})
    master = load_master()
    bnames = bom_names()
    # 匹配 HS；品名以 PDF 为准（与报关单一致），PDF 取不到时用 BOM 兜底
    for it in items:
        mp = master.get(it['erp'], {})
        it['hs'] = mp.get('hs', '')
        if not it.get('name') and it['erp'] in bnames:
            it['name'] = bnames[it['erp']]
    # 存解析结果供 generate 用（按 token 隔离，支持多文件）
    import uuid
    token = uuid.uuid4().hex[:12]
    jf = os.path.join(UPLOAD, 'parsed_%s.json' % token)
    json.dump({'name': name, 'hdr': hdr, 'items': items}, open(jf, 'w', encoding='utf-8'), ensure_ascii=False)
    return jsonify({'ok': True, 'token': token, 'name': name, 'hdr': hdr, 'items': items})

@app.route('/generate', methods=['POST'])
def generate():
    token = request.form.get('token', '')
    jf = os.path.join(UPLOAD, 'parsed_%s.json' % token) if token else ''
    if not token or not os.path.exists(jf):
        return jsonify({'ok': False, 'error': '请先上传并解析 PDF'}), 400
    d = json.load(open(jf, encoding='utf-8'))
    hdr, items = d['hdr'], d['items']
    master = load_master()
    for it in items:
        mp = master.get(it['erp'], {})
        it['hs'] = mp.get('hs', '')
    # 输出文件名：报关_报价号.xlsx
    so = hdr.get('报价号码', '').replace('/', '-') or '未命名'
    fn = '报关_%s.xlsx' % so
    out = os.path.join(OUTPUT, fn)
    written, out = build_docs(hdr, items, master, out)
    fn = os.path.basename(out)
    warn = ''
    if len(items) > BGD_MAX:
        warn = '共 %d 个 SKU，模板只预留 %d 行，第 %d 项起没有写入，请手工补或先加行' % (
            len(items), BGD_MAX, BGD_MAX + 1)
    # 返回 JSON（前端自己决定怎么下载），同时保留直链
    return jsonify({'ok': True, 'file': fn, 'path': out, 'rows': written,
                    'warning': warn, 'download': '/download/' + quote(fn)})

@app.route('/download/<path:fn>')
def download(fn):
    """供前端或系统浏览器直接下载/打开已生成文件。"""
    p = os.path.join(OUTPUT, fn)
    if not os.path.exists(p):
        return jsonify({'ok': False, 'error': '文件不存在：' + fn}), 404
    as_att = request.args.get('inline') != '1'
    return send_from_directory(OUTPUT, fn, as_attachment=as_att, download_name=fn)

@app.route('/open_output')
def open_output():
    """在资源管理器中打开 output 目录，便于手动取文件。"""
    try:
        if os.name == 'nt':
            os.startfile(OUTPUT)
        else:
            import subprocess as sp
            sp.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', OUTPUT])
        return jsonify({'ok': True, 'path': OUTPUT})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e), 'path': OUTPUT}), 500

if __name__ == '__main__':
    if os.environ.get('OPEN_BROWSER') == '1':
        import threading, webbrowser
        threading.Timer(1.5, lambda: webbrowser.open('http://127.0.0.1:8765')).start()
    app.run(host='127.0.0.1', port=8765, debug=False)
