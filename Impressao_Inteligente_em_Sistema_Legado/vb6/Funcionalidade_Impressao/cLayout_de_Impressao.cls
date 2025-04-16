Option Explicit
'
'# cLayout
' Esta classe oferece uma "biblioteca de componentes" para compor o layout de impressão.
' Preparada para ser consumida pelas estratégias de impressão (strategy pattern).
' Aqui foram retirados os detalhes de implentação para mater a atenção no escopo de apresentar
' o uso do pattern design para a solução deste estudo de caso.
'
Private Montagem As Printer
Private A4 As Boolean
Private A5 As Boolean
Private Cupom As Boolean

Private Const MARGEM_ESQUERDA_PADRAO_CUPOM As Integer = 1
Private Const MARGEM_ESQUERDA_PADRAO_A5 As Integer = 1
Private Const MARGEM_ESQUERDA_PADRAO_A4 As Integer = 4
Private Const MARGEM_ESQUERDA_EXTENDIDA_A4 As Integer = 20

Private Const TAMANHO_FONTE_PADRAO_CUPOM  As Integer = 8
Private Const TAMANHO_FONTE_PADRAO_A5  As Integer = 9
Private Const TAMANHO_FONTE_PADRAO_A4  As Integer = 11
Private Fonte_Tamanho_Padrao As Integer

Private Const FONT_NAME_PADRAO As String = "Courier New"
Private Const FONT_BOLD_PADRAO As Boolean = False

'## métodos core
Private Sub Class_Initialize()
   Set Montagem = Printer
   Montagem.Orientation = vbPRORPortrait
   Montagem.ScaleMode = vbCentimeters
End Sub

Public Sub set_Formato(ByRef pFormato As Formato_de_Impressao_Enu)
   A4 = False
   A5 = False
   Cupom = False
   
   Select Case pFormato
      Case formatoA4
         A4 = True
         Fonte_Tamanho_Padrao = TAMANHO_FONTE_PADRAO_A4
      Case formatoA5
         A5 = True
         Montagem.PaperSize = vbPRPSA5
         Montagem.CurrentY = 0.5
         Fonte_Tamanho_Padrao = TAMANHO_FONTE_PADRAO_A5
      Case formatoCupom
         Cupom = True
         Fonte_Tamanho_Padrao = TAMANHO_FONTE_PADRAO_CUPOM
         Montagem.PaperSize = vbPRPSA4
         Montagem.CurrentY = 0.5
   End Select
   Usa_Estilo_Padrao
End Sub

'## métodos de  biblioteca de "componentes"
Public Sub logo()
   '--- implementa validação
   '--- define variáveis
   '--- monta print
      If Cupom Or A5 Then
      ElseIf A4 Then
      End If
End Sub

Public Sub nome_da_empresa()
      Usa_Estilo_Padrao
      '--- configura layout especificamente para este compomente
      '--- define variáveis
      '--- monta print
      If Cupom Or A5 Then
      ElseIf A4 Then
      End If
      Usa_Estilo_Padrao
End Sub

Public Sub Cnpj()
   Usa_Estilo_Padrao
   '--- implementa validação
   '--- define variáveis
   '--- padroniza texto
   '--- monta print
   If Cupom Or A5 Then
   ElseIf A4 Then
      Usa_Estilo_Padrao
   End If
End Sub

Public Sub endereco_empresa()
   '--- valida entradas para definir se monta print
   If Len(End_Loja_Str) = 0 _
   Or Len(Numero_Loja) = 0 _
   Or Len(Bairro_Loja_Str) = 0 _
   Or Len(Cidade_Loja_Str) = 0 _
   Or Len(Estado_Loja_Str) = 0 _
   Or Len(CEP_Loja_Str) = 0 _
   Then Exit Sub
   '--- define variáveis
   Dim Endereco_Numero As String
   Dim Bairro_Cidade_Estado As String
   '--- padroniza texto
   '--- monta print
      If Cupom Or A5 Then  
      ElseIf A4 Then
      End If
End Sub

Public Sub telefone_empresa()
   '--- valida entradas para definir se monta print
   '--- define variáveis
   '--- flags para definição de regras

   '--- aplica regra e padroniza texto
      '--- Considerando que uma das variáveis corresponde a string vazia, então
      '--- concatena as duas e o resultado será a impressão apenas daquela que tem conteúdo.
   '--- monta print
      If Cupom Or A5 Then
      ElseIf A4 Then
      End If
End Sub

Public Sub email_empresa()
   '--- valida entradas para definir se monta print
   '--- define variáveis
   '--- padroniza texto
   '--- monta print
      If Cupom Or A5 Then
      ElseIf A4 Then
      End If
End Sub

Public Sub separador()
   '--- reconfigura layout
   Usa_Estilo_Padrao  
   '--- define variáveis
   '--- monta print
   Usa_Estilo_Padrao
End Sub

Public Sub linha_em_branco()
   '--- reconfigura layout
   '--- monta print
   Usa_Estilo_Padrao
End Sub

Public Sub aviso_nao_eh_doc_fiscal()
   Usa_Estilo_Padrao
   '--- define variáveis
   '--- padroniza texto
   '--- monta print
End Sub

Public Sub data_da_impressao()
   '--- define variáveis
   '--- padroniza texto
   '--- monta print
      If Cupom Then
      ElseIf A5 Then
      ElseIf A4 Then
         Usa_Estilo_Padrao
      End If
End Sub

Public Sub vendedor_operador(ByRef pNome_Funcionario As String)
   '--- define variáveis
   '--- padroniza texto
   '--- monta print
      If Cupom Or A5 Then
      ElseIf A4 Then
      End If
End Sub

Public Sub titulo_menor(ByVal pTitulo As String)
   '--- monta print
   Usa_Estilo_Padrao
End Sub

Public Sub titulo(ByVal pTitulo As String)
   '--- monta print
   Usa_Estilo_Padrao
End Sub

Public Sub titulo_maior(ByVal pTitulo As String)
   Usa_Estilo_Padrao
   '--- monta print
   Usa_Estilo_Padrao
End Sub

'... assim segue implementando todos os estilos necessários...

'## RECURSOS DE FORMATAÇÃO

Public Function Escreve_com_Quebra_de_Linha( _
                  ByRef pTexto As String, _
                  ByRef pTamanho_Fonte As Integer, _
                  Optional ByRef pMargem_Esquerda As Integer = 0, _
                  Optional ByRef pMargem_Direita As Integer = 0 _
                  ) As String
   Dim Caracteres_por_Linha As Integer
   Dim Texto_Formatado As String
   Dim x As Integer
   Dim Ultima_Linha_Impressa As Double
   Dim Contador_de_Linhas_Impressas As Integer
   
   '--- garante que o tamanho da fonte se mantenha consistente
   Printer.FontSize = pTamanho_Fonte
   
   '--- inicia variáveis
   pTexto = Trim$(pTexto)
   Texto_Formatado = ""
   Caracteres_por_Linha = Obter_Quantidade_de_Caracteres_por_Linha(pTamanho_Fonte)
   Caracteres_por_Linha = Caracteres_por_Linha - pMargem_Esquerda - pMargem_Direita
   Contador_de_Linhas_Impressas = 0
   Ultima_Linha_Impressa = Len(pTexto) / Caracteres_por_Linha
   x = 1
   
   Do While x <= Len(pTexto)
      Contador_de_Linhas_Impressas = Contador_de_Linhas_Impressas + 1
      Texto_Formatado = _
                  Texto_Formatado & _
                  Space(pMargem_Esquerda) & _
                  Mid$(pTexto, x, Caracteres_por_Linha)
      If Contador_de_Linhas_Impressas < Ultima_Linha_Impressa Then
         Texto_Formatado = Texto_Formatado & vbCrLf
      End If
      x = x + Caracteres_por_Linha
   Loop
   Escreve_com_Quebra_de_Linha = Texto_Formatado
End Function

Private Sub Escreve_Centralizado(ByVal pTexto As String)
    Dim Larg_Texto As Long: Larg_Texto = Montagem.TextWidth(pTexto)
    Dim Larg_Pag   As Long: Larg_Pag = Montagem.ScaleWidth
         Montagem.CurrentX = (Larg_Pag - Larg_Texto) / 2
         Montagem.Print pTexto
End Sub

Private Function Obter_Quantidade_de_Caracteres_por_Linha(ByRef pTamanho_Fonte As Integer) As Integer
   If Cupom Then
      Select Case pTamanho_Fonte
         Case "8"
         ' Obter_Quantidade_de_Caracteres_por_Linha
         Case "9"
         ' Obter_Quantidade_de_Caracteres_por_Linha
         Case "11"
         ' Obter_Quantidade_de_Caracteres_por_Linha
         Case Else
         ' Obter_Quantidade_de_Caracteres_por_Linha
      End Select
   ElseIf A5 Then
      Select Case pTamanho_Fonte
         Case "8"
         ' Obter_Quantidade_de_Caracteres_por_Linha
         Case "9"
         ' Obter_Quantidade_de_Caracteres_por_Linha
         Case "11"
         ' Obter_Quantidade_de_Caracteres_por_Linha
         Case Else
          'Obter_Quantidade_de_Caracteres_por_Linha
      End Select
   ElseIf A4 Then
      Select Case pTamanho_Fonte
         Case "8"
         ' Obter_Quantidade_de_Caracteres_por_Linha
         Case "9"
         ' Obter_Quantidade_de_Caracteres_por_Linha
         Case "10"
         ' Obter_Quantidade_de_Caracteres_por_Linha
         Case "11"
         ' ... 
      End Select
   End If
End Function

Private Sub Usa_Estilo_Padrao()
' implementa as configurações padrão de estilo.
End Sub
