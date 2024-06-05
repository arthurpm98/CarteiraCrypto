Attribute VB_Name = "mdlGlobal"
Option Explicit

Private urlAPI_ As String
Private APIHOST_ As String
Private APIKEY_ As String
Public rs      As WinHttpRequest
Public iAux As Integer
Public cAux As Currency

Private Sub Main()
        'Salva URL da API na variável "URLAPI"
100     urlAPI_ = "https://crypto-market-prices.p.rapidapi.com/tokens"
101     APIKEY_ = "ce5449884bmsh35da7354e60cfdep159ffbjsna6e13030ded9"
102     APIHOST_ = "crypto-market-prices.p.rapidapi.com"
        
        'Realiza um teste de conexão com a API e exibe uma mensagem de erro no StatusBarPainel caso a conexão falhe
103     If TestaConexaoAPI = False Then
104         FormPainel.StatusBarPainel.Panels.Item(2).Visible = True
105     End If
        
106     FormPainel.Show

End Sub

Public Function TestaConexaoAPI() As Boolean
100     Set rs = New WinHttpRequest
                   
        'Envia uma requisição
101     rs.Open "Get", URLAPI '& "/BTC?base=BRL"
102     rs.SetRequestHeader "X-RapidAPI-Key", APIKEY
103     rs.SetRequestHeader "X-RapidAPI-Host", APIHOST
104     rs.Send
        
105     If rs.STATUS <> 200 Then 'O status 200 é quando a requisição teve sucesso
106         TestaConexaoAPI = False
107         MsgBox "Erro de conexão com a API: " & rs.ResponseText, vbCritical
108         Exit Function
109     Else
110         TestaConexaoAPI = True
111     End If
        
112     Set rs = Nothing
End Function
Public Function RemoveCaracterEspecial(sTexto As String) As String
100     RemoveCaracterEspecial = Replace$(sTexto, "{", "")
101     RemoveCaracterEspecial = Replace$(RemoveCaracterEspecial, "}", "")
102     RemoveCaracterEspecial = Replace$(RemoveCaracterEspecial, """", "") 'Remove aspas
103     RemoveCaracterEspecial = Replace$(RemoveCaracterEspecial, ":", "")
104     RemoveCaracterEspecial = Replace$(RemoveCaracterEspecial, "[", "")
105     RemoveCaracterEspecial = Replace$(RemoveCaracterEspecial, "]", "")
End Function
Public Sub AguardeProcessamento(bAbre As Boolean)
        'Exibe ou Encerra a tela de Aguarde Processamento
100     If bAbre = True Then
101         FormAguardeProcessamento.Show
102     Else
103         Unload FormAguardeProcessamento
104     End If
End Sub
Public Property Get URLAPI() As String
    URLAPI = urlAPI_
End Property

Public Property Get APIHOST() As String
    APIHOST = APIHOST_
End Property

Public Property Get APIKEY() As String
    APIKEY = APIKEY_
End Property
