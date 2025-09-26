# =============================================
# file: danfe_report.py  (versão 2 - corrigido)
# =============================================
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

# fonte
try:
    pdfmetrics.registerFont(TTFont('DejaVu', 'DejaVuSans.ttf'))
    pdfmetrics.registerFont(TTFont('DejaVu-Bold', 'DejaVuSans-Bold.ttf')) # Registrar a fonte em negrito
    BASE_FONT = 'DejaVu'
except Exception:
    BASE_FONT = 'Helvetica'

PAGE_W, PAGE_H = A4
MARGIN = 10 * mm

# --- ESTILOS PARA PARAGRAPH ---
_styles = getSampleStyleSheet()
STYLE_NORMAL = ParagraphStyle(
    "normal",
    parent=_styles["Normal"],
    fontName=BASE_FONT,
    fontSize=8,
    leading=9.5,
)
STYLE_DESC_ITEM = ParagraphStyle(
    "desc_item",
    parent=STYLE_NORMAL,
    fontSize=8,
    leading=9.2,
)

# ---------------- utils ----------------
def _coalesce(*vals, default=""):
    for v in vals:
        if v is None:
            continue
        s = str(v).strip()
        if s:
            return s
    return default

def _label(c, x, y, text, size=8):
    c.setFont(BASE_FONT, size)
    c.drawString(x, y, str(text or ""))

def _box(c, x, y, w, h, stroke=0.8):
    c.setLineWidth(stroke)
    c.rect(x, y, w, h)

def _barcode_code128(c, x, y, w, h, chave):
    try:
        from reportlab.graphics.barcode import code128
        if not chave:
            raise ValueError("chave vazia")
        bc = code128.Code128(str(chave), barHeight=h - 8, barWidth=0.35 * mm)
        dx = x + (w - bc.width) / 2.0
        dy = y + (h - bc.height) / 2.0
        bc.drawOn(c, dx, dy)
    except Exception:
        p = Paragraph(f"<font size=7>Chave (texto)</font><br/>{chave or ''}", STYLE_NORMAL)
        pw, ph = p.wrapOn(c, w - 6, h - 6)
        p.drawOn(c, x + 3, y + h - ph - 3)

def _barcode_qr(c, x, y, w, h, data):
    _box(c, x, y, w, h, stroke=0.8)
    _label(c, x + 3, y + h - 10, "Consulta via QR Code", 7)
    if not data:
        _label(c, x + 3, y + h - 22, "(sem QR Code)", 8)
        return
    try:
        from reportlab.graphics.barcode import qr
        from reportlab.graphics.shapes import Drawing
        from reportlab.graphics import renderPDF
        size = min(w - 10, h - 26)
        qrobj = qr.QrCodeWidget(data)
        bounds = qrobj.getBounds()
        bw = bounds[2] - bounds[0]
        bh = bounds[3] - bounds[1]
        scale = size / max(bw, bh)
        d = Drawing(size, size)
        d.add(qrobj)
        d.scale(scale, scale)
        dx = x + (w - size) / 2.0
        dy = y + (h - size) / 2.0 - 5
        renderPDF.draw(d, c, dx, dy)
    except Exception:
        _label(c, x + 3, y + h - 22, str(data)[:100], 7)

# ------------- formatações -------------
def _format_emissor(doc):
    return _coalesce(doc.get("Empresa"), "")

def _format_endereco_emissor(doc):
    partes = [
        _coalesce(doc.get("EnderecoEmpresa"), ""),
        _coalesce(doc.get("BairroEmpresa"), ""),
        f"{_coalesce(doc.get('MunicipioEmpresa'), '')}/{_coalesce(doc.get('UFEmpresa'), '')}".strip("/"),
        f"CEP {_coalesce(doc.get('CEPEmpresa'), '')}".strip(),
    ]
    return " - ".join([p for p in partes if p])

def _format_destinatario(doc):
    return _coalesce(doc.get("Fornecedor"), "")

def _format_doc_info(doc):
    numero = _coalesce(doc.get("Numero"), "")
    serie = _coalesce(doc.get("Serie"), "")
    emissao = _coalesce(doc.get("Data"), "")
    se = _coalesce(doc.get("SaidaEntrada"), "")
    return f"<b>Número:</b> {numero} &nbsp; <b>Série:</b> {serie}<br/><b>Emissão:</b> {emissao} &nbsp; <b>Saída/Entrada:</b> {se}"

def _format_moeda(val):
    try:
        if isinstance(val, str):
            val = val.replace('.', '').replace(',', '.')
        val = float(val)
        return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return "0,00"

# ------------- quadros -------------
def _sec_cabecalho(c, doc, y_top):
    c.setFont(BASE_FONT, 14)
    c.drawString(MARGIN, y_top - 8, "DANFE - Documento Auxiliar da NF-e")
    
    chave = _coalesce(doc.get("Chave"), "")
    w_col = (PAGE_W - 2 * MARGIN) / 2

    # MUDANÇA (CORREÇÃO DO ERRO): Substituído <para style='...'> por tags <font> válidas.
    font_bold_name = BASE_FONT + '-Bold'
    emitente_content_html = f"""
        <font name='{font_bold_name}' size='9.5'>{_format_emissor(doc)}</font><br/>
        <font size='8'>
        <b>CNPJ/CPF:</b> {_coalesce(doc.get('CNPJCPFEmpresa'), '')}<br/>
        {_format_endereco_emissor(doc)}
        </font>
    """
    p_emitente = Paragraph(emitente_content_html, STYLE_NORMAL)
    w_emit, h_emit = p_emitente.wrapOn(c, w_col - 10, 1000)

    # MUDANÇA (CORREÇÃO DO ERRO)
    nfe_info_html = f"""
        <b><font size='12'>NF-e</font></b><br/>
        <font size='8.5'>{_format_doc_info(doc)}</font>
    """
    p_nfe = Paragraph(nfe_info_html, STYLE_NORMAL)
    w_nfe, h_nfe = p_nfe.wrapOn(c, w_col - 10, 1000)

    h_barcode_area = 18 * mm
    h_conteudo = max(h_emit, h_nfe)
    h = h_conteudo + h_barcode_area + 5*mm
    y = y_top - h

    _box(c, MARGIN, y, w_col - 3, h)
    _box(c, MARGIN + w_col, y, w_col, h)

    _label(c, MARGIN + 3, y + h - 10, "Emitente", 8)
    _label(c, MARGIN + w_col + 3, y_top - 8, "Chave de Acesso:")

    p_emitente.drawOn(c, MARGIN + 5, y + h - h_emit - 15)
    p_nfe.drawOn(c, MARGIN + w_col + 5, y + h - h_nfe - 15)

    _barcode_code128(c, MARGIN + w_col + 3, y + 3, w_col - 6, h_barcode_area, chave)
    
    return y - 2 * mm

def _sec_qrcode_e_destinatario(c, doc, y_top, qr_text):
    # MUDANÇA (CORREÇÃO DO ERRO): Substituído <para style='...'> por tags <font> válidas.
    font_bold_name = BASE_FONT + '-Bold'
    dest_content_html = f"""
        <font name='{font_bold_name}' size='9.5'>{_format_destinatario(doc)}</font><br/>
        <font size='8'>
        <b>CNPJ/CPF:</b> {_coalesce(doc.get('CNPJCPF'), '')}
        </font>
    """
    p_dest = Paragraph(dest_content_html, STYLE_NORMAL)
    
    w_qr = 42 * mm
    w_dest = PAGE_W - MARGIN - (MARGIN + w_qr + 5)
    
    w_dest_p, h_dest_p = p_dest.wrapOn(c, w_dest - 10, 1000)
    
    h = max(30 * mm, h_dest_p + 15 * mm)
    y = y_top - h
    
    _barcode_qr(c, MARGIN, y, w_qr, h, qr_text)

    x_dest = MARGIN + w_qr + 5
    _box(c, x_dest, y, w_dest, h)
    _label(c, x_dest + 3, y + h - 10, "Destinatário/Remetente", 8)
    
    p_dest.drawOn(c, x_dest + 5, y + h - h_dest_p - 15)

    return y - 2 * mm

def _sec_calculo_imposto(c, totais, y_top):
    h = 24 * mm
    y = y_top - h
    _box(c, MARGIN, y, PAGE_W - 2*MARGIN, h)
    _label(c, MARGIN + 3, y + h - 10, "Cálculo do Imposto", 8)

    data = [
        ["Base de Cálculo do ICMS", _format_moeda(totais.get('BaseICMSTot')), "Valor do IPI", _format_moeda(totais.get('ValorIPI'))],
        ["Valor do ICMS", _format_moeda(totais.get('ValorICMS')), "Valor do PIS", _format_moeda(totais.get('ValorPIS'))],
        ["Base Cálculo ICMS ST", _format_moeda(totais.get('BaseSubst')), "Valor da COFINS", _format_moeda(totais.get('ValorCOFINS'))],
        ["Valor do ICMS ST", _format_moeda(totais.get('ValorSubst')), "Valor Total dos Produtos", _format_moeda(totais.get('ValorMercadoria'))],
        ["Valor do ISS", _format_moeda(totais.get('ValorISS')), "Valor Total da Nota", Paragraph(f"<b>{_format_moeda(totais.get('ValorTotal'))}</b>", STYLE_NORMAL)],
    ]

    col_widths = [40*mm, 25*mm, 45*mm, 25*mm, 45*mm, 0]
    col_widths[-1] = PAGE_W - (2*MARGIN) - sum(col_widths) - 2

    table = Table(data, colWidths=col_widths)
    
    style = TableStyle([
        ('FONT', (0, 0), (-1, -1), BASE_FONT, 7),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ('FONTNAME', (2, 4), (3, 4), BASE_FONT + '-Bold'),
        ('FONTSIZE', (2, 4), (3, 4), 9),
    ])

    table.setStyle(style)

    tw, th = table.wrapOn(c, 0, 0)
    table.drawOn(c, MARGIN + 1, y + (h - th) - 2)

    return y - 2*mm

def _sec_transportador_volumes(c, doc, y_top):
    data = [
        [
            Paragraph(f"<b>Transportador:</b> {_coalesce(doc.get('Transportadora'), '')}", STYLE_NORMAL),
            Paragraph(f"<b>Município/UF:</b> {_coalesce(doc.get('MunicipioTransp'), '')}/{_coalesce(doc.get('UFTransp'), '')}", STYLE_NORMAL)
        ],
        [
            Paragraph(f"<b>CNPJ/CPF:</b> {_coalesce(doc.get('CNPJCPFTransp'), '')} &nbsp; <b>Inscrição:</b> {_coalesce(doc.get('InscricaoTransp'), '')}", STYLE_NORMAL),
            Paragraph(f"<b>Frete:</b> {_coalesce(doc.get('TipoFrete'), '')} &nbsp; <b>Volumes:</b> {_coalesce(doc.get('QuantidadeEmb'), '0')} &nbsp; <b>Espécie:</b> {_coalesce(doc.get('Especie'), '')}", STYLE_NORMAL)
        ],
        [
            Paragraph(f"<b>Endereço:</b> {_coalesce(doc.get('EnderecoTransp'), '')}, {_coalesce(doc.get('BairroTransp'), '')}", STYLE_NORMAL),
            Paragraph(f"<b>Peso Bruto:</b> {_coalesce(doc.get('PesoBruto'), '0')} &nbsp; <b>Peso Líquido:</b> {_coalesce(doc.get('PesoLiquido'), '0')}", STYLE_NORMAL)
        ],
    ]
    
    available_width = PAGE_W - 2 * MARGIN - 4
    table = Table(data, colWidths=[available_width/2, available_width/2])
    table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 1),
        ('RIGHTPADDING', (0, 0), (-1, -1), 1),
    ]))
    
    tw, th = table.wrapOn(c, available_width, 1000)
    
    h = th + 12 * mm
    y = y_top - h
    _box(c, MARGIN, y, PAGE_W - 2*MARGIN, h)
    _label(c, MARGIN + 3, y + h - 10, "Transportador / Volumes Transportados", 8)
    
    table.drawOn(c, MARGIN + 2, y + (h - th) - 10)

    return y - 2*mm


def _mk_itens_table(itens):
    headers = ['Cód', 'Descrição', 'NCM', 'CFOP', 'UN', 'Qtd', 'Vlr Unit', 'Vlr Total']
    data = [headers]
    for it in itens or []:
        desc = Paragraph(_coalesce(it.get('DescProduto'), '')[:1000], STYLE_DESC_ITEM)
        data.append([
            _coalesce(it.get('CodProduto'), ''),
            desc,
            _coalesce(it.get('NCM'), ''),
            _coalesce(it.get('CFOP'), ''),
            _coalesce(it.get('SgUnd'), ''),
            _format_moeda(it.get('QtCpo')),
            _format_moeda(it.get('VrCpoUnt')),
            _format_moeda(it.get('VrCpoBru')),
        ])
        
    col_widths = [40*mm, 25*mm, 45*mm, PAGE_W - (2*MARGIN) - (40*mm + 25*mm + 45*mm)]
    col_widths[1] = (PAGE_W - 2 * MARGIN) - sum(col_widths)

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, 0), BASE_FONT + '-Bold', 8),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ('FONT', (0, 1), (-1, -1), BASE_FONT, 8),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        ('ALIGN', (2, 1), (4, -1), 'CENTER'),
        ('ALIGN', (5, 1), (-1, -1), 'RIGHT'),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ]))
    return table

def _sec_itens(c, itens, y_top):
    title_h = 8 * mm
    y = y_top - title_h
    _box(c, MARGIN, y, PAGE_W - 2*MARGIN, title_h)
    _label(c, MARGIN + 3, y + title_h - 10, "Dados dos Produtos/Serviços", 8)

    table = _mk_itens_table(itens)
    table_width = PAGE_W - 2 * MARGIN
    
    available_height_for_table = y - (35*mm)
    tw, th = table.wrapOn(c, table_width, available_height_for_table)
    
    y_tbl = y - th
    table.drawOn(c, MARGIN, y_tbl)
    return y_tbl - 2*mm

def _sec_dados_adicionais(c, dados_adicionais, y_bottom):
    h = 30 * mm
    y = 20 * mm
    
    _box(c, MARGIN, y, PAGE_W - 2*MARGIN, h)
    _label(c, MARGIN + 3, y + h - 10, "Dados Adicionais", 8)
    p = Paragraph(_coalesce(dados_adicionais, ""), STYLE_DESC_ITEM)
    p.wrapOn(c, PAGE_W - 2*MARGIN - 6, h - 12)
    p.drawOn(c, MARGIN + 3, y + h - 12 - p.height)
    return y - 2*mm

# ------------- principal -------------
def build_pdf(
    output_path: str,
    doc: dict,
    itens: list,
    totais: dict,
    duplicatas: list | None = None,
    qr_text: str | None = None,
    dados_adicionais: str | None = None,
):
    if 'ValorPIS' not in totais and 'ValorTotPis' in totais:
        totais['ValorPIS'] = totais.get('ValorTotPis')
    if 'ValorCOFINS' not in totais and 'ValorTotCofins' in totais:
        totais['ValorCOFINS'] = totais.get('ValorTotCofins')

    c = canvas.Canvas(output_path, pagesize=A4)

    y = PAGE_H - MARGIN
    y = _sec_cabecalho(c, doc, y)
    y = _sec_qrcode_e_destinatario(c, doc, y, qr_text)
    y = _sec_calculo_imposto(c, totais, y)
    y = _sec_transportador_volumes(c, doc, y)
    y_after_itens = _sec_itens(c, itens, y)
    
    dados_adic = _coalesce(dados_adicionais, doc.get("DadosAdicionais"), "")
    _sec_dados_adicionais(c, dados_adic, y_after_itens)

    c.setFont(BASE_FONT, 7)
    c.drawString(MARGIN, 12, "Este DANFE é uma representação; para fins fiscais, consulte a NF-e eletrônica.")
    c.drawRightString(PAGE_W - MARGIN, 12, "Gerado com ReportLab")

    c.showPage()
    c.save()