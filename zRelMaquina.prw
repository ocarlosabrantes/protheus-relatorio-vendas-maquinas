#Include "Totvs.ch"
#Include "Protheus.ch"

/*/{Protheus.doc} User Function zRelMaquina
    Relatório de Vendas x Máquinas
    @author  Carlos.Abrantes
    @version 2.0  (refatorado)
    @since   04/03/2024
    @env     05
/*/
User Function zRelMaquina()
    Local aArea  := FWGetArea()
    Local aPergs := fMontaPergs()

    If ParamBox(aPergs, "Informe os parâmetros", , , , , , , , , .F., .F.)
        Processa({|| fGeraExcel()}, "Analisando Dados...")
    EndIf

    FWRestArea(aArea)
Return

/* ------------------------------------------------------------------ */

Static Function fMontaPergs()
    Local aPergs := {}
    Local cVenDe := Space(TamSX3("A3_COD")[1])
    Local cVenAte := StrTran(cVenDe, " ", "Z")

    aAdd(aPergs, {1, "Data De",       FirstDate(Date()), "", ".T.", "",    ".T.", 80, .F.})
    aAdd(aPergs, {1, "Data Até",      LastDate(Date()),  "", ".T.", "",    ".T.", 80, .T.})
    aAdd(aPergs, {1, "Vendedor De",   cVenDe,            "", ".T.", "SA3", ".T.", 60, .F.})
    aAdd(aPergs, {1, "Vendedor Até",  cVenAte,           "", ".T.", "SA3", ".T.", 60, .T.})

Return aPergs

/* ------------------------------------------------------------------ */

Static Function fGeraExcel()
    Local oSheet
    Local oExcel
    Local cArq      := GetTempPath() + "exporta_maquinas.xml"
    Local cSheet    := "Base"
    Local nAtual    := 0
    Local nTotal    := 0
    Local aColunas  := fDefineColunas()
    Local aLinha    := {}
    Local nCol      := 0

    If !fExecutaQuery()
        Return
    EndIf

    // Cria workbook
    oSheet := FWMSExcel():New()
    oSheet:AddWorkSheet(cSheet)
    oSheet:AddTable(cSheet, "")
    fCriaColunas(oSheet, cSheet, aColunas)

    // Conta registros para a régua
    Count To nTotal
    ProcRegua(nTotal)
    QRY_DAD->(DbGoTop())

    // Preenche linhas
    While !(QRY_DAD->(EoF()))
        nAtual++
        IncProc("Gerando planilha " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")

        aLinha := {;
            QRY_DAD->EMPRESA,;
            QRY_DAD->DATA_NF,;
            QRY_DAD->NUM_PED,;
            QRY_DAD->NUMERO_NF,;
            QRY_DAD->SERIE_NF,;
            QRY_DAD->VENDEDOR,;
            QRY_DAD->NOME_VEND,;
            QRY_DAD->CLIENTE,;
            QRY_DAD->LOJA,;
            QRY_DAD->PRODUTO,;
            QRY_DAD->DESCRICAO,;
            QRY_DAD->QUANTIDADE,;
            QRY_DAD->SERIE_PED,;
            QRY_DAD->GRUPO,;
            QRY_DAD->DESC_GRUPO,;
            QRY_DAD->MUNICIPIO,;
            QRY_DAD->ESTADO,;
            QRY_DAD->TIPO;
        }
    
        oSheet:AddRow(cSheet, "", aLinha)
        QRY_DAD->(DbSkip())
    EndDo

    QRY_DAD->(DbCloseArea())

    // Gera e abre o arquivo
    oSheet:Activate()
    oSheet:GetXMLFile(cArq)
    FreeObj(oSheet)

    oExcel := MsExcel():New()
    oExcel:WorkBooks:Open(cArq)
    oExcel:SetVisible(.T.)
    oExcel:Destroy()
    FreeObj(oExcel)

    MsgInfo("Planilha gerada com sucesso!", "Concluído")

Return

/* ------------------------------------------------------------------ */

Static Function fExecutaQuery()
    Local cSQL := ""

    // Filiais hardcoded em constante local para facilitar manutenção futura
    Local cFiliais := "'01','02'"

    cSQL += "SELECT"                                                              + CRLF
    cSQL += "    SD2.D2_FILIAL    AS EMPRESA,"                                    + CRLF
    cSQL += "    SD2.D2_EMISSAO   AS DATA_NF,"                                    + CRLF
    cSQL += "    SD2.D2_PEDIDO    AS NUM_PED,"                                    + CRLF
    cSQL += "    SD2.D2_DOC       AS NUMERO_NF,"                                  + CRLF
    cSQL += "    SD2.D2_SERIE     AS SERIE_NF,"                                   + CRLF
    cSQL += "    SA3.A3_COD       AS VENDEDOR,"                                   + CRLF
    cSQL += "    SA3.A3_NOME      AS NOME_VEND,"                                  + CRLF
    cSQL += "    SA1.A1_NOME      AS CLIENTE,"                                    + CRLF
    cSQL += "    SD2.D2_LOJA      AS LOJA,"                                       + CRLF
    cSQL += "    SD2.D2_COD       AS PRODUTO,"                                    + CRLF
    cSQL += "    SB1.B1_DESC      AS DESCRICAO,"                                  + CRLF
    cSQL += "    SD2.D2_QUANT     AS QUANTIDADE,"                                 + CRLF
    cSQL += "    SDB.DB_NUMSERI   AS SERIE_PED,"                                  + CRLF
    cSQL += "    SB1.B1_GRUPO     AS GRUPO,"                                      + CRLF
    cSQL += "    SB1.B1_XDSCGR    AS DESC_GRUPO,"                                 + CRLF
    cSQL += "    SA1.A1_MUN       AS MUNICIPIO,"                                  + CRLF
    cSQL += "    SA1.A1_EST       AS ESTADO,"                                     + CRLF
    cSQL += "    SF4.F4_TEXTO     AS TIPO"                                        + CRLF
    cSQL += " FROM "     + RetSQLName("SD2") + " SD2"                            + CRLF
    cSQL += " INNER JOIN " + RetSQLName("SF2") + " SF2"                          + CRLF
    cSQL += "    ON  SF2.F2_DOC    = SD2.D2_DOC"                                  + CRLF
    cSQL += "    AND SF2.F2_SERIE  = SD2.D2_SERIE"                                + CRLF
    cSQL += "    AND SF2.F2_FILIAL = SD2.D2_FILIAL"                               + CRLF
    cSQL += "    AND SF2.D_E_L_E_T_ = ' '"                                        + CRLF
    cSQL += " INNER JOIN " + RetSQLName("SA3") + " SA3"                          + CRLF
    cSQL += "    ON  SA3.A3_COD      = SF2.F2_VEND1"                              + CRLF
    cSQL += "    AND SA3.A3_FILIAL   = '" + FWxFilial("SA3") + "'"               + CRLF
    cSQL += "    AND SA3.D_E_L_E_T_ = ' '"                                        + CRLF
    cSQL += " INNER JOIN " + RetSQLName("SB1") + " SB1"                          + CRLF
    cSQL += "    ON  SB1.B1_COD      = SD2.D2_COD"                                + CRLF
    cSQL += "    AND SB1.B1_FILIAL   = '" + FWxFilial("SB1") + "'"               + CRLF
    cSQL += "    AND SB1.D_E_L_E_T_ = ' '"                                        + CRLF
    cSQL += " INNER JOIN " + RetSQLName("SA1") + " SA1"                          + CRLF
    cSQL += "    ON  SA1.A1_COD      = SD2.D2_CLIENTE"                            + CRLF
    cSQL += "    AND SA1.A1_LOJA     = SD2.D2_LOJA"                               + CRLF
    cSQL += "    AND SA1.A1_FILIAL   = '" + FWxFilial("SA1") + "'"               + CRLF
    cSQL += "    AND SA1.D_E_L_E_T_ = ' '"                                        + CRLF
    cSQL += " LEFT JOIN "  + RetSQLName("SDB") + " SDB"                          + CRLF
    cSQL += "    ON  SDB.DB_FILIAL   = SD2.D2_FILIAL"                             + CRLF
    cSQL += "    AND SDB.DB_DOC      = SD2.D2_DOC"                                + CRLF
    cSQL += "    AND SDB.DB_TM       = SD2.D2_TES"                                + CRLF
    cSQL += "    AND SDB.DB_CLIFOR   = SD2.D2_CLIENTE"                            + CRLF
    cSQL += "    AND SDB.DB_DATA     = SD2.D2_EMISSAO"                            + CRLF
    cSQL += "    AND SDB.DB_SERIE    = SD2.D2_SERIE"                              + CRLF
    cSQL += "    AND SDB.DB_PRODUTO  = SD2.D2_COD"                                + CRLF
    cSQL += "    AND SDB.DB_NUMSERI  = SD2.D2_NUMSERI"                            + CRLF
    cSQL += "    AND SDB.D_E_L_E_T_ = ' '"                                        + CRLF
    cSQL += " LEFT JOIN "  + RetSQLName("SF4") + " SF4"                          + CRLF
    cSQL += "    ON  SF4.F4_CODIGO   = SD2.D2_TES"                                + CRLF
    cSQL += "    AND SF4.D_E_L_E_T_ = ' '"                                        + CRLF
    cSQL += " WHERE SD2.D2_FILIAL    IN (" + cFiliais + ")"                      + CRLF
    cSQL += "   AND SD2.D2_EMISSAO BETWEEN '" + dToS(MV_PAR01) + "' AND '" + dToS(MV_PAR02) + "'" + CRLF
    cSQL += "   AND SA3.A3_COD     BETWEEN '" + MV_PAR03 + "' AND '" + MV_PAR04 + "'"             + CRLF
    cSQL += "   AND SD2.D2_SERIE    = '3'"                                        + CRLF
    cSQL += "   AND SB1.B1_LOCPAD   = 'MQ'"                                      + CRLF
    cSQL += "   AND SF4.F4_XRELFAT  = '1'"                                       + CRLF
    cSQL += "   AND SD2.D_E_L_E_T_ = ' '"                                        + CRLF
    cSQL += " ORDER BY SD2.D2_EMISSAO, SD2.D2_PEDIDO"                            + CRLF

    PlsQuery(cSQL, "QRY_DAD")
    TCSetField("QRY_DAD", "DATA_NF", "D")
    DbSelectArea("QRY_DAD")

    If QRY_DAD->(EoF())
        MsgStop("Nenhum registro encontrado para os filtros informados.", "Atenção")
        QRY_DAD->(DbCloseArea())
        Return .F.
    EndIf

Return .T.

/* ------------------------------------------------------------------ */
/*
    Colunas: { cTítulo, nTipo, nFormato }
    nTipo: 1=texto  2=data  nFormato: 1=geral  2=numérico
*/
Static Function fDefineColunas()
 Return {;
        {"Empresa",        1, 1},;
        {"Data NF",        2, 1},;
        {"Num. Pedido",    1, 1},;
        {"Num. Docto.",    1, 1},;
        {"Serie NF",       1, 1},;
        {"Cod. Vendedor",  1, 1},;
        {"Nome Vendedor",  1, 1},;
        {"Nome Cliente",   1, 1},;
        {"Loja",           1, 1},;
        {"Produto",        1, 1},;
        {"Desc. Produto",  1, 1},;
        {"Quantidade",     2, 2},;
        {"Num. Serie",     1, 1},;
        {"Grupo Estat.",   1, 1},;
        {"Desc. Grupo",    1, 1},;
        {"Municipio",      1, 1},;
        {"Estado",         1, 1},;
        {"Tipo",           1, 1};
    }

Static Function fCriaColunas(oSheet, cSheet, aColunas)
    Local nI := 0

    For nI := 1 To Len(aColunas)
         oSheet:AddColumn(cSheet, "", aColunas[nI][1], aColunas[nI][2], aColunas[nI][3], .F.)
    Next nI

Return
