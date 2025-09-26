# app.py — UI mínima (só DOC) + tema dark purple + SQL parametrizada
import sys, os
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QFormLayout,
    QLineEdit, QPushButton, QFileDialog, QMessageBox, QLabel
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette, QColor

from db_config import DBConfig, get_conn
from danfe_report import build_pdf

# ----------------- THEME (dark purple) -----------------
DARK_BG   = QColor(23, 19, 33)     # roxo bem escuro
CARD_BG   = QColor(33, 28, 48)
ACCENT    = QColor(111, 66, 193)   # “bootstrap purple”
FG        = QColor(235, 235, 245)
MUTED_FG  = QColor(200, 200, 215)

STYLE_SHEET = f"""
    QWidget {{
        background-color: {CARD_BG.name()};
        color: {FG.name()};
        font-size: 14px;
    }}
    QLineEdit {{
        background-color: #1f1a30;
        padding: 8px 10px;
        border: 1px solid #3b2d63;
        border-radius: 8px;
        color: {FG.name()};
        selection-background-color: {ACCENT.name()};
    }}
    QLineEdit:focus {{ border: 1px solid {ACCENT.name()}; }}
    QLabel {{ color: {MUTED_FG.name()}; font-size: 12px; }}
    QPushButton {{
        background-color: {ACCENT.name()};
        color: white;
        padding: 9px 16px;
        border-radius: 10px;
        font-weight: 600;
    }}
    QPushButton:hover   {{ background-color: #7f54db; }}
    QPushButton:pressed {{ background-color: #6943b7; }}
"""

def apply_dark_theme(app: QApplication):
    pal = QPalette()
    pal.setColor(QPalette.Window, CARD_BG)
    pal.setColor(QPalette.WindowText, FG)
    pal.setColor(QPalette.Base, DARK_BG)
    pal.setColor(QPalette.AlternateBase, CARD_BG)
    pal.setColor(QPalette.ToolTipBase, CARD_BG)
    pal.setColor(QPalette.ToolTipText, FG)
    pal.setColor(QPalette.Text, FG)
    pal.setColor(QPalette.Button, ACCENT)
    pal.setColor(QPalette.ButtonText, QColor("white"))
    pal.setColor(QPalette.Highlight, ACCENT)
    pal.setColor(QPalette.HighlightedText, QColor("white"))
    app.setPalette(pal)
    app.setStyleSheet(STYLE_SHEET)

# ----------------- SUA QUERY “isso funciona” (parametrizada) -----------------
SQL_QUERY = r"""
SELECT
    Cpd.CdCpd
,   Numero                   = Ffm.NrFfm
,   NatOperacao              = Top1.NmTop

,   Empresa                  = PesUne.NmPes 
,   EnderecoEmpresa          = TlgUne.SgTlg + ' ' + LgrUne.NmLgr 
                               + CASE WHEN PesUne.NrPesEdr IS NULL THEN '' ELSE ', ' + PesUne.NrPesEdr END 
                               + ' ' + ISNULL(PesUne.NrPesEdrCom, '')
,   BairroEmpresa            = LocBaiUne.NmLoc
,   MunicipioEmpresa         = LocCidUne.NmLoc
,   UFEmpresa                = LocEstUne.SgLoc 
,   FoneEmpresa              = CONVERT(varchar, LocPaiUne.NrLocDDI) + ' (' + CONVERT(varchar, LocCidUne.NrLocDDD) +') ' + CONVERT(varchar, MctUne.NrMctTel)
,   CEPEmpresa               = PesUne.NrPesEdrCep
,   InscricaoEmpresa         = PesUne.NrPesCgf 
,   InscricaoEmpresaSubstTrib= ''

,   CNPJCPFEmpresa           = CASE WHEN PesUne.TpPes = 1
                                   THEN CONVERT(varchar, CASE WHEN PesUne.NrPesCpj IS NOT NULL 
                                          THEN SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesUne.NrPesCpj))),' ','0')+CONVERT(varchar,PesUne.NrPesCpj),1,2)+'.'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesUne.NrPesCpj))),' ','0')+CONVERT(varchar,PesUne.NrPesCpj),3,3)+'.'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesUne.NrPesCpj))),' ','0')+CONVERT(varchar,PesUne.NrPesCpj),6,3) +'/'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesUne.NrPesCpj))),' ','0')+CONVERT(varchar,PesUne.NrPesCpj),9,4) +'-'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesUne.NrPesCpj))),' ','0')+CONVERT(varchar,PesUne.NrPesCpj),13,2)
                                          ELSE '** *** *** **** **' END)
                                   ELSE CONVERT(varchar, CASE WHEN PesUne.NrPesCpj IS NOT NULL 
                                          THEN SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,PesUne.NrPesCpj))),' ','0')+CONVERT(varchar,PesUne.NrPesCpj),1,3)+'.'
                                             + SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,PesUne.NrPesCpj))),' ','0')+CONVERT(varchar,PesUne.NrPesCpj),4,3)+'.'
                                             + SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,PesUne.NrPesCpj))),' ','0')+CONVERT(varchar,PesUne.NrPesCpj),7,3) +'-'
                                             + SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,PesUne.NrPesCpj))),' ','0')+CONVERT(varchar,PesUne.NrPesCpj),10,2) 
                                          ELSE '*** *** *** **' END)
                               END

,   SaidaEntrada             = CASE WHEN Top1.TpTopSin = 1 THEN '1' WHEN Top1.TpTopSin = 3 THEN '2' END
,   CodigoDeBarras           = dbo.FnFormataCodBarraCode128(Nfe.Chave)
,   Chave                    = Nfe.Chave

,   Fornecedor               = CONVERT(varchar, Frn.CdFrn) + ' - ' + Pes.NmPes

,   CNPJCPF                  = CASE WHEN Pes.TpPes = 1
                                   THEN CONVERT(varchar, CASE WHEN Pes.NrPesCpj IS NOT NULL 
                                          THEN SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,Pes.NrPesCpj))),' ','0')+CONVERT(varchar,Pes.NrPesCpj),1,2)+'.'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,Pes.NrPesCpj))),' ','0')+CONVERT(varchar,Pes.NrPesCpj),3,3)+'.'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,Pes.NrPesCpj))),' ','0')+CONVERT(varchar,Pes.NrPesCpj),6,3) +'/'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,Pes.NrPesCpj))),' ','0')+CONVERT(varchar,Pes.NrPesCpj),9,4) +'-'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,Pes.NrPesCpj))),' ','0')+CONVERT(varchar,Pes.NrPesCpj),13,2)
                                          ELSE '** *** *** **** **' END)
                                   ELSE CONVERT(varchar, CASE WHEN Pes.NrPesCpj IS NOT NULL 
                                          THEN SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,Pes.NrPesCpj))),' ','0')+CONVERT(varchar,Pes.NrPesCpj),1,3)+'.'
                                             + SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,Pes.NrPesCpj))),' ','0')+CONVERT(varchar,Pes.NrPesCpj),4,3)+'.'
                                             + SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,Pes.NrPesCpj))),' ','0')+CONVERT(varchar,Pes.NrPesCpj),7,3) +'-'
                                             + SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,Pes.NrPesCpj))),' ','0')+CONVERT(varchar,Pes.NrPesCpj),10,2) 
                                          ELSE '*** *** *** **' END)
                               END

,   Data                     = Cpd.DtCpdEmi
,   DataEntrada              = Cpd.DtCpdCon
,   HoraSaida                = ISNULL(SUBSTRING(CONVERT(varchar,Nfe.DataeHoradoProcessamento,114),1,5),'00:00')

,   Fatura                   = dbo.FuFnFaturaCpd(Cpd.CdCpd)

,   Endereco                 = Tlg.SgTlg + ' ' + Lgr.NmLgr 
                               + CASE WHEN Pes.NrPesEdr IS NULL THEN '' ELSE ', ' + Pes.NrPesEdr END
                               + ' ' + ISNULL(Pes.NrPesEdrCom, '')
,   Bairro                   = LocBai.NmLoc
,   CEP                      = Pes.NrPesEdrCep
,   Municipio                = LocCid.NmLoc
,   Fone                     = CONVERT(varchar, LocCid.NrLocDdd) + ' ' + CONVERT(varchar, Mct.NrMctTel)
,   UF                       = LocEst.SgLoc
,   Inscricao                = Pes.NrPesCgf
,   InscricaoMunicipal       = ''

-- Totais do documento (uma vez via APPLY) - usando precisões do Psi
,   BaseICMSTot              = dbo.FnFormataValor(Tot.BaseICMSTot, '.', ',', Psi.QtPsiDecVal)
,   BaseSubst                = dbo.FnFormataValor(Tot.BaseSubst   , '.', ',', Psi.QtPsiDecVal)
,   BaseISS                  = dbo.FnFormataValor(Tot.BaseISS     , '.', ',', Psi.QtPsiDecVal)
,   ValorICMS                = dbo.FnFormataValor(Tot.ValorICMS   , '.', ',', Psi.QtPsiDecVal)
,   ValorSubst               = dbo.FnFormataValor(Tot.ValorSubst  , '.', ',', Psi.QtPsiDecVal)
,   ValorIPI                 = dbo.FnFormataValor(Tot.ValorIPI    , '.', ',', Psi.QtPsiDecVal)
,   ValorISS                 = dbo.FnFormataValor(Tot.ValorISS    , '.', ',', Psi.QtPsiDecVal)
,   ValorMercadoria          = dbo.FnFormataValor(ISNULL(Cpd.VrCpdMer,0), '.', ',', Psi.QtPsiDecVal)
,   ValorTotal               = dbo.FnFormataValor(ISNULL(Cpd.VrCpd,0),    '.', ',', Psi.QtPsiDecVal)
,   ValorFrete               = dbo.FnFormataValor(Tot.ValorFrete  , '.', ',', Psi.QtPsiDecVal)
,   ValorSeguro              = dbo.FnFormataValor(Tot.ValorSeguro , '.', ',', Psi.QtPsiDecVal)
,   ValorDesconto            = dbo.FnFormataValor(Tot.ValorDesconto,'.', ',', Psi.QtPsiDecVal)
,   ValorDespAce             = dbo.FnFormataValor(Tot.ValorDespAce, '.', ',', Psi.QtPsiDecVal)
,   ValorTotPis              = dbo.FnFormataValor(Tot.ValorPIS    , '.', ',', Psi.QtPsiDecVal)
,   ValorTotCofins           = dbo.FnFormataValor(Tot.ValorCOFINS , '.', ',', Psi.QtPsiDecVal)
,   VrTotTrib                = CONVERT(decimal(19,2), 0.00)

,   Transportadora           = ISNULL(PesTrp.NmPes, '')
,   CNPJCPFTransp            = CASE WHEN Pes.TpPes = 1
                                   THEN CONVERT(varchar, CASE WHEN PesTrp.NrPesCpj IS NOT NULL 
                                          THEN SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesTrp.NrPesCpj))),' ','0')+CONVERT(varchar,PesTrp.NrPesCpj),1,2)+'.'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesTrp.NrPesCpj))),' ','0')+CONVERT(varchar,PesTrp.NrPesCpj),3,3)+'.'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesTrp.NrPesCpj))),' ','0')+CONVERT(varchar,PesTrp.NrPesCpj),6,3) +'/'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesTrp.NrPesCpj))),' ','0')+CONVERT(varchar,PesTrp.NrPesCpj),9,4) +'-'
                                             + SUBSTRING(REPLACE(SPACE(14-LEN(CONVERT(varchar,PesTrp.NrPesCpj))),' ','0')+CONVERT(varchar,PesTrp.NrPesCpj),13,2)
                                          ELSE '** *** *** **** **' END)
                                   ELSE CONVERT(varchar, CASE WHEN PesTrp.NrPesCpj IS NOT NULL 
                                          THEN SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,PesTrp.NrPesCpj))),' ','0')+CONVERT(varchar,PesTrp.NrPesCpj),1,3)+'.'
                                             + SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,PesTrp.NrPesCpj))),' ','0')+CONVERT(varchar,PesTrp.NrPesCpj),4,3)+'.'
                                             + SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,PesTrp.NrPesCpj))),' ','0')+CONVERT(varchar,PesTrp.NrPesCpj),7,3) +'-'
                                             + SUBSTRING(REPLACE(SPACE(11-LEN(CONVERT(varchar,PesTrp.NrPesCpj))),' ','0')+CONVERT(varchar,PesTrp.NrPesCpj),10,2) 
                                          ELSE '*** *** *** **' END)
                               END
,   TipoFrete                = CASE WHEN Cpd.TpCpdFrt = 1 THEN '0' WHEN Cpd.TpCpdFrt = 2 THEN '1' WHEN Cpd.TpCpdFrt = 3 THEN '2' ELSE '9' END
,   EnderecoTransp           = TlgTrp.SgTlg + ' ' + LgrTrp.NmLgr 
                               + CASE WHEN PesTrp.NrPesEdr IS NULL THEN '' ELSE ', ' + PesTrp.NrPesEdr END
                               + ' ' + ISNULL(PesTrp.NrPesEdrCom, '')
,   BairroTransp             = LocBaiTrp.NmLoc
,   MunicipioTransp          = LocCidTrp.NmLoc
,   UFTransp                 = LocEstTrp.SgLoc
,   InscricaoTransp          = PesTrp.NrPesCgf
,   CodAntt                  = ''
,   Placa                    = ''
,   UFPlaca                  = ''
,   Marca                    = ''
,   Numeracao                = ''
,   QuantidadeEmb            = dbo.FnFormataValor(ISNULL(Cpd.QtCpdEmb, 0), '.', ',', Psi.QtPsiDecQtd)
,   Especie                  = UndEmb.SgUnd 
,   PesoBruto                = dbo.FnFormataValor(ISNULL(Cpd.QtCpdPsoBru, 0), '.', ',', Psi.QtPsiDecQtd)
,   PesoLiquido              = dbo.FnFormataValor(ISNULL(Cpd.QtCpdPsoLiq, 0), '.', ',', Psi.QtPsiDecQtd)

,   DadosAdicionais          = ISNULL(Obs.TtObs, '') + ISNULL(dbo.FuFnObsCpo(Cpo.CdCpo),'')
,   Serie                    = ISNULL(Frm.NrFrmSer,'001')
,   TpOpe                    = Top1.SgTop

-- INFORMAÇÕES DO OBJETO (por item)
,   CodProduto               = Obj.CdObj
,   DescProduto              = Obj.NmObj
,   NCM                      = ISNULL(OpoNCM.TtOpo, '')
,   CST                      = ISNULL(CpsCST.TtCps, '0')
,   CFOP                     = ISNULL(CpsCFOP.TtCps, '')
,   SgUnd                    = ISNULL(Und.SgUnd, '')
,   QtCpo                    = dbo.FnFormataValor(ISNULL(Cpo.QtCpo, 0), '.', ',', Psi.QtPsiDecQtd)
,   VrCpoUnt                 = dbo.FnFormataValor(ISNULL(Cpo.VrCpoUnt, 0), '.', ',', Psi.QtPsiDecPrc)
,   VrCpoBru                 = dbo.FnFormataValor(ISNULL(Cpo.VrCpoBru, 0), '.', ',', Psi.QtPsiDecVal)

-- valores por item (via APPLY usando OES do NFP)
,   BaseICMS                 = dbo.FnFormataValor(ISNULL(AggItem.BaseICMS,0),  '.', ',', Psi.QtPsiDecVal)
,   ICMS                     = dbo.FnFormataValor(ISNULL(AggItem.ValorICMS,0), '.', ',', Psi.QtPsiDecVal)
,   IPI                      = dbo.FnFormataValor(ISNULL(AggItem.ValorIPI,0),  '.', ',', Psi.QtPsiDecVal)
,   AliqICMS                 = dbo.FnFormataValor(ISNULL(AggItem.AliqICMS,0),  '.', ',', Psi.QtPsiDecVal)
,   AliqIPI                  = dbo.FnFormataValor(ISNULL(AggItem.AliqIPI,0),   '.', ',', Psi.QtPsiDecVal)

,   Cpo.NrCpoOrd

-- INFORMAÇÕES COMPLEMENTARES
,   NmFantasia               = ''
,   PtReferencia             = ''
,   ComplEndereco            = ''
,   NmVendedor               = ''
,   NrPedido                 = ''
,   Televendas               = ''
,   DispositivoLegal         = Dpl.TtDpl

-- AUTORIZADO
,   Aut                      = CASE WHEN Nfe.Situacao = 5 THEN 'DOCUMENTO NÃO AUTORIZADO' ELSE '' END
,   VrCpoDesRep              = ''
,   HistRep                  = ''

-- PESSOA DE ENTREGA
,   EnderecoEnt              = TlgEnt.SgTlg + ' ' + LgrEnt.NmLgr 
                               + CASE WHEN PesEnt.NrPesEdr IS NULL THEN '' ELSE ', ' + PesEnt.NrPesEdr END
                               + ' ' + ISNULL(PesEnt.NrPesEdrCom, '')
,   BairroEnt                = LocBaiEnt.NmLoc
,   CEPEnt                   = PesEnt.NrPesEdrCep
,   MunicipioEnt             = LocCidEnt.NmLoc
,   FoneEnt                  = CONVERT(varchar, LocCidEnt.NrLocDdd) + ' ' + CONVERT(varchar, MctEnt.NrMctTel)
,   UFEnt                    = LocEstEnt.SgLoc

,   situacao                 = ISNULL(NFe.Situacao,0)

FROM TbCpd           AS Cpd   WITH (NOLOCK)
JOIN TbCpo           AS Cpo   WITH (NOLOCK) ON Cpo.CdCpd = Cpd.CdCpd
JOIN TbFrn           AS Frn   WITH (NOLOCK) ON Frn.CdFrn = Cpd.CdFrn
JOIN TbPes           AS Pes   WITH (NOLOCK) ON Pes.CdPes = Frn.CdPes
LEFT JOIN TbLgr      AS Lgr   WITH (NOLOCK) ON Lgr.CdLgr = Pes.CdLgr
LEFT JOIN TbTlg      AS Tlg   WITH (NOLOCK) ON Tlg.CdTlg = Lgr.CdTlg
LEFT JOIN TbLoc      AS LocBai WITH (NOLOCK) ON LocBai.CdLoc = Pes.CdLoc AND LocBai.TpLoc = 5
LEFT JOIN TbLoc      AS LocCid WITH (NOLOCK) ON LocCid.CdLoc = LocBai.CdLocMae AND LocCid.TpLoc = 4
LEFT JOIN TbLoc      AS LocEst WITH (NOLOCK) ON LocEst.CdLoc = LocCid.CdLocMae AND LocEst.TpLoc = 3

-- Endereço de Entrega
LEFT JOIN TbPes      AS PesEnt   WITH (NOLOCK) ON PesEnt.CdPes = Cpd.CdPesOri
LEFT JOIN TbLgr      AS LgrEnt   WITH (NOLOCK) ON LgrEnt.CdLgr = PesEnt.CdLgr
LEFT JOIN TbTlg      AS TlgEnt   WITH (NOLOCK) ON TlgEnt.CdTlg = LgrEnt.CdTlg
LEFT JOIN TbLoc      AS LocBaiEnt WITH (NOLOCK) ON LocBaiEnt.CdLoc = PesEnt.CdLoc AND LocBaiEnt.TpLoc = 5
LEFT JOIN TbLoc      AS LocCidEnt WITH (NOLOCK) ON LocCidEnt.CdLoc = LocBaiEnt.CdLocMae AND LocCidEnt.TpLoc = 4
LEFT JOIN TbLoc      AS LocEstEnt WITH (NOLOCK) ON LocEstEnt.CdLoc = LocCidEnt.CdLocMae AND LocEstEnt.TpLoc = 3
LEFT JOIN TbMct      AS MctEnt WITH (NOLOCK) ON MctEnt.CdMct = PesEnt.CdMct

JOIN TbUne           AS Une   WITH (NOLOCK) ON Une.CdUne = Cpd.CdUne 
JOIN TbPes           AS PesUne WITH (NOLOCK) ON PesUne.CdPes = Une.CdPes
JOIN TbLgr           AS LgrUne WITH (NOLOCK) ON LgrUne.CdLgr = PesUne.CdLgr
JOIN TbTlg           AS TlgUne WITH (NOLOCK) ON TlgUne.CdTlg = LgrUne.CdTlg
JOIN TbLoc           AS LocBaiUne WITH (NOLOCK) ON LocBaiUne.CdLoc = PesUne.CdLoc AND LocBaiUne.TpLoc = 5
JOIN TbLoc           AS LocCidUne WITH (NOLOCK) ON LocCidUne.CdLoc = LocBaiUne.CdLocMae AND LocCidUne.TpLoc = 4
JOIN TbLoc           AS LocEstUne WITH (NOLOCK) ON LocEstUne.CdLoc = LocCidUne.CdLocMae AND LocEstUne.TpLoc = 3
JOIN TbLoc           AS LocPaiUne WITH (NOLOCK) ON LocPaiUne.CdLoc = LocCidUne.CdLoc001 AND LocPaiUne.TpLoc = 1

JOIN TbTop           AS Top1 WITH (NOLOCK) ON Top1.CdTop = Cpd.CdTop
JOIN TbObj           AS Obj  WITH (NOLOCK) ON Obj.CdObj = Cpo.CdObj

-- NCM por objeto
LEFT JOIN TbOpo      AS OpoNCM WITH (NOLOCK) ON OpoNCM.CdObj = Cpo.CdObj
LEFT JOIN TbNfp      AS NfpNCM WITH (NOLOCK) ON NfpNCM.CdCrcNcm = OpoNCM.CdCrc

LEFT JOIN TbUnd      AS Und    WITH (NOLOCK) ON Und.CdUnd = Cpo.CdUnd
LEFT JOIN TbUnd      AS UndEmb WITH (NOLOCK) ON UndEmb.CdUnd = Cpd.CdUndEmb

LEFT JOIN TbMct      AS Mct    WITH (NOLOCK) ON Mct.CdMct = Pes.CdMct
LEFT JOIN TbMct      AS MctUne WITH (NOLOCK) ON MctUne.CdMct = PesUne.CdMct

LEFT JOIN TbObs      AS Obs WITH (NOLOCK) ON Obs.CdObs = Cpd.CdObs

LEFT JOIN TbFrn      AS FrnTrp WITH (NOLOCK) ON FrnTrp.CdFrn = Cpd.CdFrnTrp
LEFT JOIN TbPes      AS PesTrp WITH (NOLOCK) ON PesTrp.CdPes = FrnTrp.CdPes
LEFT JOIN TbLgr      AS LgrTrp WITH (NOLOCK) ON LgrTrp.CdLgr = PesTrp.CdLgr
LEFT JOIN TbTlg      AS TlgTrp WITH (NOLOCK) ON TlgTrp.CdTlg = LgrTrp.CdTlg
LEFT JOIN TbLoc      AS LocBaiTrp WITH (NOLOCK) ON LocBaiTrp.CdLoc = PesTrp.CdLoc AND LocBaiTrp.TpLoc = 5
LEFT JOIN TbLoc      AS LocCidTrp WITH (NOLOCK) ON LocCidTrp.CdLoc = LocBaiTrp.CdLocMae AND LocCidTrp.TpLoc = 4
LEFT JOIN TbLoc      AS LocEstTrp WITH (NOLOCK) ON LocEstTrp.CdLoc = LocCidTrp.CdLocMae AND LocEstTrp.TpLoc = 3

LEFT JOIN TbFfm      AS Ffm WITH (NOLOCK) ON Ffm.CdFfm = Cpd.FolhaDeFormularioID_Nfe
LEFT JOIN NFe        AS Nfe WITH (NOLOCK) ON Nfe.FolhaDeFormularioID = Ffm.CdFfm AND Nfe.Situacao <> 5
LEFT JOIN TbFrm      AS Frm WITH (NOLOCK) ON Frm.CdFrm = Ffm.CdFrm

-- Precisões (UMA LINHA)
CROSS APPLY (
    SELECT TOP (1)
        QtPsiDecVal,
        QtPsiDecQtd,
        QtPsiDecPrc
    FROM TbPsi WITH (NOLOCK)
    ORDER BY QtPsiDecVal DESC
) AS Psi

-- Mapeamento NFP (UMA LINHA)
CROSS APPLY (
      SELECT TOP (1)
             CdOesIcmBas, CdOesIcmVal, CdOesIpiVal, CdOesIssqnBas, CdOesIssqnVal
           , CdOesFre, CdOesSeg, CdOesPisVal, CdOesCofVal, CdOesIIVal
           , CdOesIcmStBas, CdOesIcmStVal, CdOesAceDoc, CdOesIcmAli, CdOesIpiAli
           , CdOesCfo, CdOesIcmCst
      FROM TbNfp WITH (NOLOCK)
      ORDER BY CdNfp
) AS NfpMap

-- Dispositivo legal
LEFT JOIN TbDpl AS Dpl WITH (NOLOCK) 
       ON Dpl.CdDpl = (SELECT TOP 1 CdDpl FROM TbCps WITH (NOLOCK) 
                       WHERE CdCpo = Cpo.CdCpo AND CdOes = NfpMap.CdOesCfo)

-- Totais do documento (UMA LINHA)
CROSS APPLY (
      SELECT
          BaseICMSTot = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesIcmBas),
          BaseSubst   = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesIcmStBas),
          BaseISS     = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesIssqnBas),
          ValorICMS   = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesIcmVal),
          ValorSubst  = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesIcmStVal),
          ValorIPI    = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesIpiVal),
          ValorISS    = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesIssqnVal),
          ValorFrete  = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesFre),
          ValorSeguro = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesSeg),
          ValorPIS    = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesPisVal),
          ValorCOFINS = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesCofVal),
          ValorII     = (SELECT SUM(ISNULL(cps.VrCps,0))
                         FROM TbCpo cpo2 WITH (NOLOCK)
                         LEFT JOIN TbCps cps WITH (NOLOCK) ON cps.CdCpo = cpo2.CdCpo
                         WHERE cpo2.CdCpd = Cpd.CdCpd AND cps.CdOes = NfpMap.CdOesIIVal),
          ValorDesconto = (SELECT SUM(ISNULL(cpo2.VrCpoDes,0) + ISNULL(cpo2.VrCpoCpdDes,0))
                           FROM TbCpo cpo2 WITH (NOLOCK)
                           WHERE cpo2.CdCpd = Cpd.CdCpd),
          ValorDespAce  = (SELECT ISNULL(MAX(VrCps),0)
                           FROM TbCps WITH (NOLOCK)
                           WHERE CdCpd = Cpd.CdCpd AND CdOes = NfpMap.CdOesAceDoc)
) AS Tot

-- Agregados por item (1x/item)
OUTER APPLY (
      SELECT 
          BaseICMS  = (SELECT SUM(ISNULL(cps2.VrCps,0)) FROM TbCps cps2 WITH (NOLOCK)
                       WHERE cps2.CdCpo = Cpo.CdCpo AND cps2.CdOes = NfpMap.CdOesIcmBas),
          ValorICMS = (SELECT SUM(ISNULL(cps2.VrCps,0)) FROM TbCps cps2 WITH (NOLOCK)
                       WHERE cps2.CdCpo = Cpo.CdCpo AND cps2.CdOes = NfpMap.CdOesIcmVal),
          ValorIPI  = (SELECT SUM(ISNULL(cps2.VrCps,0)) FROM TbCps cps2 WITH (NOLOCK)
                       WHERE cps2.CdCpo = Cpo.CdCpo AND cps2.CdOes = NfpMap.CdOesIpiVal),
          AliqICMS  = (SELECT SUM(ISNULL(cps2.VrCps,0)) FROM TbCps cps2 WITH (NOLOCK)
                       WHERE cps2.CdCpo = Cpo.CdCpo AND cps2.CdOes = NfpMap.CdOesIcmAli),
          AliqIPI   = (SELECT SUM(ISNULL(cps2.VrCps,0)) FROM TbCps cps2 WITH (NOLOCK)
                       WHERE cps2.CdCpo = Cpo.CdCpo AND cps2.CdOes = NfpMap.CdOesIpiAli)
) AS AggItem

-- CFOP/CST (texto) por item
OUTER APPLY (SELECT TOP (1) TtCps 
             FROM TbCps WITH (NOLOCK)
             WHERE CdCpo = Cpo.CdCpo AND CdOes = NfpMap.CdOesCfo)    AS CpsCFOP
OUTER APPLY (SELECT TOP (1) TtCps 
             FROM TbCps WITH (NOLOCK)
             WHERE CdCpo = Cpo.CdCpo AND CdOes = NfpMap.CdOesIcmCst) AS CpsCST

WHERE Cpd.CdCpd = ?
"""

# ----------------- UI mínima -----------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerar DANFE (PDF)")
        self.setMinimumWidth(520)

        w = QWidget()
        v = QVBoxLayout(w)
        form = QFormLayout()

        self.ed_doc = QLineEdit()
        self.ed_doc.setPlaceholderText("Ex.: 5455")
        form.addRow("DOC:", self.ed_doc)
        v.addLayout(form)

        self.bt_pdf = QPushButton("Gerar DANFE (PDF)")
        self.bt_pdf.clicked.connect(self.on_generate)
        v.addWidget(self.bt_pdf, alignment=Qt.AlignLeft)

        self.setCentralWidget(w)

    def on_generate(self):
        cd_text = self.ed_doc.text().strip()
        if not cd_text.isdigit():
            QMessageBox.warning(self, "Atenção", "Informe o DOC (CdCpd) numérico.")
            return
        cdcpd = int(cd_text)

        # local para salvar
        suggested = os.path.join(os.path.expanduser("~"), "Documents", f"DANFE_{cdcpd}.pdf")
        out_path, _ = QFileDialog.getSaveFileName(self, "Salvar PDF", suggested, "PDF (*.pdf)")
        if not out_path:
            return

        try:
            # Conexão por defaults de db_config.py (sem mostrar na UI)
            cfg = DBConfig()
            with get_conn(cfg) as conn:
                cur = conn.cursor()
                cur.execute(SQL_QUERY, (cdcpd,))
                rows = cur.fetchall()
                if not rows:
                    QMessageBox.information(self, "Sem dados", f"Nenhum resultado para DOC={cdcpd}.")
                    return

                cols = [d[0] for d in cur.description]
                recs = [dict(zip(cols, r)) for r in rows]

                head = recs[0]
                totais = {
                    'BaseICMSTot': head.get('BaseICMSTot'),
                    'ValorICMS': head.get('ValorICMS'),
                    'BaseSubst': head.get('BaseSubst'),
                    'ValorSubst': head.get('ValorSubst'),
                    'ValorIPI': head.get('ValorIPI'),
                    'ValorISS': head.get('ValorISS'),
                    'ValorPIS': head.get('ValorTotPis') or head.get('ValorPIS'),
                    'ValorCOFINS': head.get('ValorTotCofins') or head.get('ValorCOFINS'),
                    'ValorFrete': head.get('ValorFrete'),
                    'ValorSeguro': head.get('ValorSeguro'),
                    'ValorDesconto': head.get('ValorDesconto'),
                    'ValorDespAce': head.get('ValorDespAce'),
                    'ValorMercadoria': head.get('ValorMercadoria'),
                    'ValorTotal': head.get('ValorTotal'),
                }

                item_fields = ['CodProduto','DescProduto','NCM','CFOP','SgUnd','QtCpo','VrCpoUnt','VrCpoBru']
                itens = [{k: r.get(k) for k in item_fields} for r in recs]

                qr_text = head.get("Chave")
                dados_adicionais = head.get("DadosAdicionais")
                duplicatas = None

                build_pdf(
                    out_path,
                    head,
                    itens,
                    totais,
                    duplicatas=duplicatas,
                    qr_text=qr_text,
                    dados_adicionais=dados_adicionais
                )

            # abre o PDF
            try:
                os.startfile(out_path)  # Windows
            except Exception:
                pass

            QMessageBox.information(self, "OK", f"PDF gerado em:\n{out_path}")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao gerar PDF:\n{e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_dark_theme(app)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
