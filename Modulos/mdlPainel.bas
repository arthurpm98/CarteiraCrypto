Attribute VB_Name = "mdlPainel"
Option Explicit
Private brl_, usdt_, paxg_, eth_, btc_ As Boolean
Private sMoedaUtilizada_, sMoedaCompra_ As String
Private cPreco_, cPrecoMaisBaixoDia_, cPrecoMaisBaixoSemana_, cPrecoMaisBaixoMes_, cPrecoMaisBaixoAno_ As Currency
Private dDataCotacao_ As String
Private tHoraCotacao_ As String

Public Sub ControlaMarcacaoFiltros(iValor As Boolean, _
                                   bOption As Boolean, _
                                   index As Integer)
        'Preenche as variáveis das moedas que serão utilizadas na pesquisa
100     If iValor = True Then
101         If index = 0 Then
102             If bOption = True Then
103                 BRL = True: USDT = False
104             Else
105                 BTC = True
106             End If
107         End If
108         If index = 1 Then
109             If bOption = True Then
110                 USDT = True: BRL = False
111             Else
112                 ETH = True
113             End If
114         End If
115         If index = 2 Then
116             If bOption = False Then
117                 PAXG = True
118             End If
119         End If
120     Else
121         If index = 0 Then
122             If bOption = True Then
123                 BRL = False
124             Else
125                 BTC = False
126             End If
127         End If
128         If index = 1 Then
129             If bOption = True Then
130                 USDT = False
131             Else
132                 ETH = False
133             End If
134         End If
135         If index = 2 Then
136             If bOption = False Then
137                 PAXG = False
138             End If
139         End If
140     End If
End Sub
Public Function VerificaOptionMarcado() As Boolean
100     iAux = 0
        'Verifica se existe pelo menos 1 option marcado
101     For iAux = 0 To FormPainel.optMoeda.Count - 1 'Quantidade de Index dos options
102         If FormPainel.optMoeda(iAux).Value = True Then VerificaOptionMarcado = True
103     Next
End Function
Public Function VerificaCheckboxMarcado() As Boolean
100     iAux = 0
        'Verifica se existe pelo menos 1 checkbox marcado
101     For iAux = 0 To FormPainel.chkMoeda.Count - 1 'Quantidade de Index dos checkboxs
102         If CBool(FormPainel.chkMoeda(iAux).Value) = True Then VerificaCheckboxMarcado = True
103     Next
End Function
Public Function MontaQueryFiltrosPainel() As String
100     Dim sMoedaUtilizada As String
        'Verifica qual será a moeda utilizada como base de preço
101     If BRL = True Then
102         sMoedaUtilizada = "?base=BRL"
103     ElseIf USDT = True Then
104         sMoedaUtilizada = "?base=USDT"
105     End If

        'Verifica quais moedas queremos comprar
106     If BTC = True Then MontaQueryFiltrosPainel = "/BTC" & sMoedaUtilizada & ";"
107     If ETH = True Then MontaQueryFiltrosPainel = MontaQueryFiltrosPainel & "/ETH" & sMoedaUtilizada & ";"
108     If PAXG = True Then MontaQueryFiltrosPainel = MontaQueryFiltrosPainel & "/PAXG" & sMoedaUtilizada & ";"
End Function

Public Sub RealizaRequisicao(sParametros As String)
100     Dim vQuery As Variant
101     vQuery = Split(sParametros, ";")
102     For iAux = 0 To UBound(vQuery)
103         If Trim(vQuery(iAux)) <> "" Then
104             Set rs = New WinHttpRequest
            
105             rs.Open "GET", URLAPI & vQuery(iAux)
106             rs.SetRequestHeader "X-RapidAPI-Key", APIKEY
107             rs.SetRequestHeader "X-RapidAPI-Host", APIHOST
108             rs.Send
                
109             If rs.Status <> 200 Then
110                 MsgBox "Erro de conexão com a API: " & rs.ResponseText, vbCritical
111                 Set rs = Nothing
112                 Exit Sub
113             End If

114             TratamentoRespostaRequisicao (rs.ResponseText)
115             GravaRespostaRequisicaoEmArquivo
116             Set rs = Nothing
117         End If
118     Next
End Sub

Private Sub TratamentoRespostaRequisicao(sResposta As String)
        'Preenche as variáveis que serão utilizadas pra salvar inserir dados no arquivo "valores.csv" e preencher o grid do painel.
100     Dim vResposta As Variant
101     sResposta = RemoveCaracterEspecial(sResposta)
102     vResposta = Split(sResposta, ",")
103     MOEDACOMPRA = Replace(vResposta(2), "datasymbol", "")
104     MOEDAUTILIZADA = Replace(vResposta(3), "base", "")
105     PRECO = CCur(Val(Replace$(vResposta(5), "price", "")))
106     DATACOTACAO = Format(Date, "Short Date")
107     HORACOTACAO = Format(Time, "hh:mm:ss")
108     PRECOBAIXODIA = VerificaPrecoMaisBaixo(Trim$(MOEDACOMPRA), Trim$((MOEDAUTILIZADA)), PRECO, DATACOTACAO, "D")
109     PRECOBAIXOSEMANA = VerificaPrecoMaisBaixo(Trim$(MOEDACOMPRA), Trim$((MOEDAUTILIZADA)), PRECO, DATACOTACAO, "W")
110     PRECOBAIXOMES = VerificaPrecoMaisBaixo(Trim$(MOEDACOMPRA), Trim$((MOEDAUTILIZADA)), PRECO, DATACOTACAO, "M")
111     PRECOBAIXOANO = VerificaPrecoMaisBaixo(Trim$(MOEDACOMPRA), Trim$((MOEDAUTILIZADA)), PRECO, DATACOTACAO, "YYYY")
End Sub

Private Function VerificaPrecoMaisBaixo(sMoedaCompra As String, _
                               sMoedaUtilizada As String, _
                               cPreco As Currency, _
                               sData As String, _
                               sTipoVerificacao As String) As Currency
100     Dim sArquivo As String
101     Dim iArquivo  As Integer
102     Dim sLinha    As String
103     Dim vLinha    As Variant
104     Dim lCodigo   As Long
        
        'Verifica o Preço mais baixo da moeda no arquivo valores.csv
106     cAux = 0

107     sArquivo = App.Path & "\Banco de Dados\valores.csv"
108     If Dir(sArquivo) <> "" Then
109         iArquivo = FreeFile
110         Open sArquivo For Input As #iArquivo
111         VerificaPrecoMaisBaixo = cPreco
112         Do While Not EOF(iArquivo)
113             Line Input #iArquivo, sLinha
114             vLinha = Split(sLinha, ";")
115             If Trim(vLinha(0)) <> "codigo" Then
116                 If Trim$(vLinha(1)) = sMoedaCompra And Trim$(vLinha(2)) = sMoedaUtilizada And DateDiff(sTipoVerificacao, Format(vLinha(4), "Short Date"), sData) = 0 Then
117                 cAux = IIf(CCur(Val(vLinha(3))) < cPreco, CCur(Val(vLinha(3))), cPreco)
118                 VerificaPrecoMaisBaixo = IIf(cAux < VerificaPrecoMaisBaixo, cAux, VerificaPrecoMaisBaixo)
119                 End If
120             End If
121             DoEvents
122         Loop
123         Close #iArquivo
124     Else
125         MsgBox "Arquivo bd.csv não encontrado!", vbCritical
126         Exit Function
127     End If
End Function

Private Sub GravaRespostaRequisicaoEmArquivo()
100     Dim sArquivo As String
101     Dim iArquivo As Integer
102     Dim sLinha   As String
103     Dim vLinha   As Variant
104     Dim lCodigo  As Long
        
        'Verifica o próximo código
105     sArquivo = App.Path & "\Banco de Dados\valores.csv"
106     If Dir(sArquivo) <> "" Then
107         iArquivo = FreeFile
108         Open sArquivo For Input As #iArquivo
109         Do While Not EOF(iArquivo)
110             Line Input #iArquivo, sLinha
111             vLinha = Split(sLinha, ";")
112             If Trim(vLinha(0)) <> "codigo" Then
113                 lCodigo = Val(vLinha(0))
114             End If
115             DoEvents
116         Loop
117         Close #iArquivo
118     Else
119         MsgBox "Arquivo bd.csv não encontrado!", vbCritical
120         Exit Sub
121     End If
        
122     lCodigo = lCodigo + 1
        
123     iArquivo = FreeFile
124     Open sArquivo For Append As #iArquivo
125     sLinha = lCodigo & ";" & MOEDACOMPRA & ";" & MOEDAUTILIZADA & ";" & PRECO & ";" & DATACOTACAO & ";" & HORACOTACAO & ";" & PRECOBAIXODIA & ";" & PRECOBAIXOSEMANA & ";" & PRECOBAIXOMES & ";" & PRECOBAIXOANO & ";"
126     Print #iArquivo, sLinha
127     Close #iArquivo
End Sub

Public Sub PreencheGridPainel()
100     Dim sArquivo As String
101     Dim iArquivo    As Integer
102     Dim sLinha      As String
103     Dim vLinha      As Variant
104     Dim vMatrizGrid As XArrayDB
105     Dim iContador   As Integer
        
106     sArquivo = App.Path & "\Banco de Dados\valores.csv"
107     If Dir(sArquivo) <> "" Then
108         iArquivo = FreeFile
109         iContador = 1
110         Set vMatrizGrid = New XArrayDB
111         Open sArquivo For Input As #iArquivo
112         Do While Not EOF(iArquivo)
113             Line Input #iArquivo, sLinha
114             vLinha = Split(sLinha, ";")
115             If Trim(vLinha(0)) <> "codigo" Then
116                 vMatrizGrid.ReDim 1, vMatrizGrid.Count(1) + 1, 0, 9
117                 vMatrizGrid(iContador, 0) = CStr(vLinha(1))
118                 vMatrizGrid(iContador, 1) = CStr(vLinha(2))
119                 vMatrizGrid(iContador, 2) = CCur(Val(vLinha(3)))
120                 vMatrizGrid(iContador, 3) = CStr(vLinha(4))
121                 vMatrizGrid(iContador, 4) = CStr(vLinha(5))
122                 vMatrizGrid(iContador, 5) = CCur(Val(vLinha(6)))
123                 vMatrizGrid(iContador, 6) = CCur(Val(vLinha(7)))
124                 vMatrizGrid(iContador, 7) = CCur(Val(vLinha(8)))
125                 vMatrizGrid(iContador, 8) = CCur(Val(vLinha(9)))
126                 iContador = iContador + 1
127             End If
128             DoEvents
129         Loop
130         FormPainel.GridPainel.Array = vMatrizGrid
131         FormPainel.GridPainel.ReBind
132         FormPainel.GridPainel.Refresh
            
133         Set vMatrizGrid = Nothing
134         Close #iArquivo
135     Else
136         MsgBox "Arquivo bd.csv não encontrado!", vbCritical
137         Exit Sub
138     End If
End Sub
Public Property Get BRL() As Boolean
100     BRL = brl_
End Property

Public Property Let BRL(ByVal valor As Boolean)
100     brl_ = valor
End Property

Public Property Get USDT() As Boolean
100     USDT = usdt_
End Property

Public Property Let USDT(ByVal valor As Boolean)
100     usdt_ = valor
End Property

Public Property Get PAXG() As Boolean
100     PAXG = paxg_
End Property

Public Property Let PAXG(ByVal valor As Boolean)
100     paxg_ = valor
End Property

Public Property Get ETH() As Boolean
100     ETH = eth_
End Property

Public Property Let ETH(ByVal valor As Boolean)
100     eth_ = valor
End Property

Public Property Get BTC() As Boolean
100     BTC = btc_
End Property

Public Property Let BTC(ByVal valor As Boolean)
100     btc_ = valor
End Property

Public Property Get MOEDAUTILIZADA() As String
100     MOEDAUTILIZADA = sMoedaUtilizada_
End Property

Public Property Let MOEDAUTILIZADA(ByVal valor As String)
100     sMoedaUtilizada_ = valor
End Property
Public Property Get MOEDACOMPRA() As String
100     MOEDACOMPRA = sMoedaCompra_
End Property

Public Property Let MOEDACOMPRA(ByVal valor As String)
100     sMoedaCompra_ = valor
End Property

Public Property Get PRECO() As Currency
100     PRECO = cPreco_
End Property

Public Property Let PRECO(ByVal valor As Currency)
100     cPreco_ = valor
End Property

Public Property Get PRECOBAIXODIA() As Currency
100     PRECOBAIXODIA = cPrecoMaisBaixoDia_
End Property

Public Property Let PRECOBAIXODIA(ByVal valor As Currency)
100     cPrecoMaisBaixoDia_ = valor
End Property

Public Property Get PRECOBAIXOSEMANA() As Currency
100     PRECOBAIXOSEMANA = cPrecoMaisBaixoSemana_
End Property

Public Property Let PRECOBAIXOSEMANA(ByVal valor As Currency)
100     cPrecoMaisBaixoSemana_ = valor
End Property

Public Property Get PRECOBAIXOMES() As Currency
100     PRECOBAIXOMES = cPrecoMaisBaixoMes_
End Property

Public Property Let PRECOBAIXOMES(ByVal valor As Currency)
100     cPrecoMaisBaixoMes_ = valor
End Property

Public Property Get PRECOBAIXOANO() As Currency
100     PRECOBAIXOANO = cPrecoMaisBaixoAno_
End Property

Public Property Let PRECOBAIXOANO(ByVal valor As Currency)
100     cPrecoMaisBaixoAno_ = valor
End Property

Public Property Get DATACOTACAO() As String
100     DATACOTACAO = dDataCotacao_
End Property

Public Property Let DATACOTACAO(ByVal valor As String)
100     dDataCotacao_ = valor
End Property

Public Property Get HORACOTACAO() As String
100     HORACOTACAO = tHoraCotacao_
End Property

Public Property Let HORACOTACAO(ByVal valor As String)
100     tHoraCotacao_ = valor
End Property
