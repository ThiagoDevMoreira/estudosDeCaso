VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLayout_de_Impressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'# cLayout
' Esta classe oferece uma biblioteca de "componentes" para compor o layout de impressão.
' Preparada para ser consumida pelas estratégias de impressão (strategy pattern).
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
   '--- valida
   If Mostrar_Logotipo_Str <> 1 Then Exit Sub
   
   '--- define variáveis
   Dim Img           As PictureBox: Set Img = frmStatus.Picture5
   Dim Img_Largura   As Integer: Img_Largura = Img.width
   Dim Img_Altura    As Integer: Img_Altura = Img.height
   Dim i As Integer

   '--- monta print
      If Cupom Or A5 Then
         Montagem.PaintPicture Img.Picture, (Img_Largura / 2), 0, Img_Largura, Img_Altura
         For i = 1 To 10 ' de linhas em branco
            Montagem.Print
         Next i
      ElseIf A4 Then
         Printer.PaintPicture Img.Picture, 200, 250, Img.width, Img.height
      End If
End Sub

Public Sub nome_da_empresa()
      '--- reconfigura layout
      Usa_Estilo_Padrao
      Montagem.FontBold = True
      Montagem.FontSize = 12
      
      '--- define variáveis
      Dim nome_principal As String: nome_principal = Nome_Loja_Str
      Dim nome_secundario As String: nome_secundario = Nome_Loja2_Str
      
      '--- monta print
      If Cupom Or A5 Then
         Escreve_Centralizado nome_principal
         If Len(nome_secundario) > 0 Then Escreve_Centralizado nome_secundario
      ElseIf A4 Then
         If Mostrar_Logotipo_Str = 1 Then
            linha_em_branco
            Montagem.Print Space$(MARGEM_ESQUERDA_EXTENDIDA_A4) & nome_principal
            If Len(nome_secundario) > 0 Then
               Montagem.Print Space$(MARGEM_ESQUERDA_EXTENDIDA_A4) & nome_secundario
            End If
         Else
            PrinterCentralizado Nome_Loja_Str
            If Len(nome_secundario) > 0 Then PrinterCentralizado Nome_Loja2_Str
         End If
      End If
      Usa_Estilo_Padrao
End Sub

Public Sub Cnpj()
   Usa_Estilo_Padrao
   '--- valida entradas para definir se monta print
   If Len(CNPJ_Loja_Str) = 0 Then Exit Sub

   '--- define variáveis
   Dim Cnpj As String
   
   '--- padroniza texto
   Cnpj = "CNPJ: " & CNPJ_Loja_Str

   '--- monta print
   If Cupom Or A5 Then
      Escreve_Centralizado Cnpj
   ElseIf A4 Then
      If Mostrar_Logotipo_Str = 1 Then
         Montagem.Print Space$(MARGEM_ESQUERDA_EXTENDIDA_A4) & Cnpj
      Else
         Printer.FontSize = 9
         PrinterCentralizado Cnpj
      End If
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
   Endereco_Numero = End_Loja_Str & ", " & Numero_Loja & " | " & Bairro_Loja_Str
   Bairro_Cidade_Estado = Cidade_Loja_Str & " | " & Estado_Loja_Str & " | " & CEP_Loja_Str

   '--- monta print
      If Cupom Or A5 Then
         PrinterCentralizado Endereco_Numero
         PrinterCentralizado Bairro_Cidade_Estado
         
      ElseIf A4 Then
         If Mostrar_Logotipo_Str = 1 Then
            Montagem.Print Space$(MARGEM_ESQUERDA_EXTENDIDA_A4) & Endereco_Numero
            Montagem.Print Space$(MARGEM_ESQUERDA_EXTENDIDA_A4) & Bairro_Cidade_Estado
         Else
            PrinterCentralizado Endereco_Numero
            PrinterCentralizado Bairro_Cidade_Estado
         End If
      End If
End Sub

Public Sub telefone_empresa()
   '--- valida entradas para definir se monta print
   If Len(Fone_Loja_Str) = 0 _
   And Len(Cel_Loja_Str) = 0 _
   Then Exit Sub

   '--- define variáveis
   Dim Telefones As String
   Dim Tem_Telefone As Boolean
   Dim Tem_Celular As Boolean
   
   '--- flags para definição de regras
   Tem_Telefone = False
   Tem_Celular = False
   If Len(Fone_Loja_Str) > 0 Then Tem_Telefone = True
   If Len(Cel_Loja_Str) > 0 Then Tem_Celular = True

   '--- aplica regra e padroniza texto
   If Tem_Telefone And Tem_Celular Then
      Telefones = "TEL: " & Fone_Loja_Str & " | " & Cel_Loja_Str
   Else
      '--- Considerando que uma das variáceis corresponde a string vazia, então
      '--- concatena as duas e o resultado será a impressão apenas daquela que tem conteúdo.
      Telefones = "TEL: " & Fone_Loja_Str & Cel_Loja_Str
   End If
   
   '--- monta print
      If Cupom Or A5 Then
         PrinterCentralizado Telefones
      ElseIf A4 Then
         If Mostrar_Logotipo_Str = 1 Then
            Montagem.Print Space$(MARGEM_ESQUERDA_EXTENDIDA_A4) & Telefones
         Else
            PrinterCentralizado Telefones
         End If
      End If
End Sub

Public Sub email_empresa()
   '--- valida entradas para definir se monta print
   If Len(Email_Loja_Str) = 0 Then Exit Sub

   '--- define variáveis
   Dim Email As String
   
   '--- padroniza texto
   Email = "EMAIL: " & Email_Loja_Str
   
   '--- monta print
      If Cupom Or A5 Then
         PrinterCentralizado Email
      ElseIf A4 Then
         If Mostrar_Logotipo_Str = 1 Then
            Montagem.Print Space$(MARGEM_ESQUERDA_EXTENDIDA_A4) & Email
         Else
            PrinterCentralizado Email
         End If
      End If
End Sub

Public Sub separador()
   '--- reconfigura layout
   Usa_Estilo_Padrao
   Montagem.FontSize = 8
   
   '--- define variáveis
   Dim Comprimento As Integer
   Comprimento = Obter_Quantidade_de_Caracteres_por_Linha(Montagem.FontSize)
   
   '--- monta print
      Montagem.Print String$(Comprimento, "-")

   Usa_Estilo_Padrao
End Sub

Public Sub linha_em_branco()
   '--- reconfigura layout
   Usa_Estilo_Padrao
   Montagem.FontSize = 8
   
   '--- monta print
      Montagem.Print 'linha em branco
      
   Usa_Estilo_Padrao
End Sub

Public Sub aviso_nao_eh_doc_fiscal()
   Usa_Estilo_Padrao
   Montagem.FontSize = 12
   '--- define variáveis
   Dim Aviso As String
   
   '--- padroniza texto
   Aviso = "NÃO É DOCUMENTO FISCAL"

   '--- monta print
      Escreve_Centralizado Aviso
   Usa_Estilo_Padrao
End Sub

Public Sub data_da_impressao()
   '--- define variáveis
   Dim Data As String
   Dim Hora As String
   
   '--- padroniza texto
   Data = "DATA: " & Format$(Data_Sistema, "dd/MM/yy")
   Hora = "HORA: " & Format$(Time, "HH:mm")

   '--- monta print
      If Cupom Then
         Montagem.Print Data & Space$(16) & hora
      ElseIf A5 Then
         Escreve_Centralizado Data & Space$(10) & hora
      ElseIf A4 Then
         Usa_Estilo_Padrao
         Montagem.FontBold = True
         
         linha_em_branco
         Montagem.Print Space$(MARGEM_ESQUERDA_PADRAO_A4) & Data & Space$(4) & Hora
         
         Usa_Estilo_Padrao
      End If
End Sub

Public Sub vendedor_operador(ByRef pNome_Funcionario As String)
   '--- define variáveis
   Dim Vendedor As String
   Dim Operador As String
   
   '--- padroniza texto
   Vendedor = ""
   Operador = ""
   If Len(pNome_Funcionario) > 0 Then
      Vendedor = "VENDEDOR: " & pNome_Funcionario
   End If
   If Len(frmMDI.StatusBar1.Panels(2).Text) > 0 Then
      Operador = "OPERADOR: " & frmMDI.StatusBar1.Panels(2).Text
   End If
   
   '--- monta print
      If Cupom Or A5 Then
         If Len(pNome_Funcionario) > 0 Then Montagem.Print Vendedor
         If Len(frmMDI.StatusBar1.Panels(2).Text) > 0 Then Montagem.Print Operador
      ElseIf A4 Then
         Montagem.FontSize = 9
         Montagem.Print Space$(MARGEM_ESQUERDA_PADRAO_A4) & Vendedor & Space$(4) & Operador
         Usa_Estilo_Padrao
      End If
   
End Sub

Public Sub titulo_menor(ByVal pTitulo As String)
   Usa_Estilo_Padrao
   Montagem.FontSize = 11
   Montagem.FontBold = True
   
   '--- monta print
   Montagem.Print pTitulo
   
   Usa_Estilo_Padrao
End Sub

Public Sub titulo(ByVal pTitulo As String)
   Usa_Estilo_Padrao
   Montagem.FontSize = 13
   Montagem.FontBold = True
   
   '--- monta print
   Montagem.Print pTitulo
   
   Usa_Estilo_Padrao
End Sub

Public Sub titulo_maior(ByVal pTitulo As String)
   Usa_Estilo_Padrao
   Montagem.FontSize = 20
   Montagem.FontBold = True
   
   '--- monta print
   Montagem.Print pTitulo
   
   Usa_Estilo_Padrao
End Sub

Public Sub num_pedido(ByRef pNum_Pedido As Long)
   '--- reconfigura layout
   Usa_Estilo_Padrao
   Montagem.FontSize = 14
   Montagem.FontBold = True
   
   '--- define variáveis
   Dim num_pedido As String
   
   '--- padroniza texto
   num_pedido = "ORDEM: " & pNum_Pedido

   '--- monta print
      Montagem.Print num_pedido

   Usa_Estilo_Padrao
End Sub

Public Sub cod_e_nome_cliente(ByRef pCod_Cliente As Long, ByRef pNome_Cliente As String)
   '--- define variáveis
   Dim Cod_Cliente As String
   Dim Nome_Cliente As String
   Dim Codigo_e_Nome_Cliente As String
   
   '--- padroniza texto
   Cod_Cliente = "CLIENTE: " & Trim(pCod_Cliente)
   Nome_Cliente = pNome_Cliente
   Codigo_e_Nome_Cliente = "CLIENTE: " & Cod_Cliente & "-" & Nome_Cliente

   '--- monta print
      If Cupom Then
         Montagem.FontSize = 10
         
         Montagem.Print Cod_Cliente
         Montagem.Print Escreve_com_Quebra_de_Linha(Nome_Cliente, TAMANHO_FONTE_PADRAO_CUPOM)
      ElseIf A5 Then
         Montagem.Print Escreve_com_Quebra_de_Linha(Codigo_e_Nome_Cliente, 10)
      ElseIf A4 Then
         Montagem.FontBold = True
         linha_em_branco
         Montagem.Print Escreve_com_Quebra_de_Linha(Codigo_e_Nome_Cliente, 10, MARGEM_ESQUERDA_PADRAO_A4)
         linha_em_branco
      End If
   Usa_Estilo_Padrao
End Sub

Public Sub endereco_cliente( _
               ByRef pEndereco As String, _
               ByRef pNumero As String, _
               ByRef pBairro As String, _
               ByRef pCidade As String, _
               ByRef pEstado As String _
               )
   '--- define variáveis
   Dim Endereco_e_Num As String
   Dim Bairro_Cidade_Estado As String
   Dim Sem_Endereco As Boolean
   Dim Aviso_de_Erro_no_Endereco As String
   
   Sem_Endereco = False
   If Len(pEndereco) = 0 Then Sem_Endereco = True
   
   '--- padroniza texto
   Endereco_e_Num = pEndereco & ", " & pNumero
   Bairro_Cidade_Estado = pBairro & " - " & pCidade & "/" & pEstado
   Aviso_de_Erro_no_Endereco = "Cadastro do endereço está incompleto."
   
   '--- monta print
      If Sem_Endereco Then
         Montagem.Print Escreve_com_Quebra_de_Linha(Aviso_de_Erro_no_Endereco, TAMANHO_FONTE_PADRAO_CUPOM)
         Exit Sub
      End If
      If Cupom Or A5 Then
         Montagem.Print Escreve_com_Quebra_de_Linha(Endereco_e_Num, TAMANHO_FONTE_PADRAO_CUPOM)
         Montagem.Print Escreve_com_Quebra_de_Linha(Bairro_Cidade_Estado, TAMANHO_FONTE_PADRAO_CUPOM)
      ElseIf A4 Then
         Montagem.Print Escreve_com_Quebra_de_Linha(Endereco_e_Num, 9, MARGEM_ESQUERDA_PADRAO_A4)
         Montagem.Print Escreve_com_Quebra_de_Linha(Bairro_Cidade_Estado, 9, MARGEM_ESQUERDA_PADRAO_A4)
      End If
End Sub

Public Sub telefone_cliente(ByRef pTelefone As String, ByRef pCelular As String)
   '--- valida
   If Len(pTelefone) = 0 _
   And Len(pCelular) = 0 _
   Then Exit Sub

   '--- define variáveis
   Dim Tel As String
   Dim Cel As String
   Dim Tel_e_Cel As String
   Tel = ""
   Cel = ""
   If Len(pTelefone) > 0 Then Tel = pTelefone
   If Len(pCelular) > 0 Then Cel = pCelular
   
   '--- padroniza texto
   If Len(pTelefone) > 0 And Len(pCelular) > 0 Then
      Tel_e_Cel = "FONES: " & Tel & " | " & Cel
   Else
      Tel_e_Cel = "FONES: " & Tel & Cel
   End If
   
   '--- monta print
      Montagem.Print Tel_e_Cel
End Sub

Public Sub cpf_cliente(ByRef pCpf As String)
   '--- valida
   If Len(pCpf) = 0 Then Exit Sub
   
   '--- define variáveis
   Dim Cpf As String
   
   '--- padroniza texto
   Cpf = "CPF: " & pCpf
   
   '--- monta print
      If Cupom Or A5 Then
         Montagem.Print Cpf
      ElseIf A4 Then
         Montagem.Print Space$(MARGEM_ESQUERDA_PADRAO_A4) & Cpf
      End If
End Sub

Public Sub endereco_de_entrega(ByRef pEndereco_Entrega As String)
   Usa_Estilo_Padrao
   Montagem.FontBold = True
   
   '--- define variáveis
   Dim Endereco As String

   '--- padroniza texto
   Endereco = "END. ENTREGA: O MESMO"
   If pEndereco_Entrega <> "O MESMO" _
   And Len(pEndereco_Entrega) > 0 Then
      Endereco = "END. ENTREGA: " & pEndereco_Entrega
   End If

   '--- monta print
      Montagem.Print Escreve_com_Quebra_de_Linha(Endereco, Fonte_Tamanho_Padrao)

   Usa_Estilo_Padrao
End Sub

Public Sub observacao(ByRef pObservacao As String)
   '--- valida
   If Len(pObservacao) = 0 Then Exit Sub

   '--- define variáveis
   Dim Obs As String
   
   '--- padroniza texto
   Obs = "OBSERVAÇÃO: " & pObservacao
 
   '--- monta print
      linha_em_branco
      
      If Cupom Or A5 Then
         Montagem.Print Escreve_com_Quebra_de_Linha(Obs, Fonte_Tamanho_Padrao)
      ElseIf A4 Then
         Usa_Estilo_Padrao
         Montagem.FontSize = 9
         Montagem.FontBold = True
         Montagem.Print Space$(MARGEM_ESQUERDA_PADRAO_A4) & "OBSERVAÇÃO:"
         Printer.FontBold = False
         Montagem.Print Escreve_com_Quebra_de_Linha(pObservacao, Montagem.FontSize, MARGEM_ESQUERDA_PADRAO_A4)
         Usa_Estilo_Padrao
      End If
End Sub

Public Sub lista_produtos_venda(ByRef pLista_de_Itens() As ctlPedidos_Detalhes)
   '--- define variáveis conteúdo
   Dim Cabecario_da_Lista              As String
   Dim Item                            As Variant
   Dim Codigo                          As Long
   Dim Descricao                       As String
   Dim Qtd                             As Integer
   Dim total_a_pagar                   As Currency
   Dim Valor_Desconto                  As Currency
   Dim Desconto                        As Double
   Dim Unidade                         As String
   Dim valor_por_produto               As Currency
   Dim total_por_produto               As Currency
   Dim Item_Primeira_Linha             As String
   Dim Item_Segunda_Linha              As String
   Dim Item_Terceira_Linha             As String
   Dim mostrar_desconto                As Boolean
   '--- define veriáveis diagramação
   ' _ cabeçário
   Dim ESPACAMENTO_POS_COD             As Integer
   Dim ESPACAMENTO_POS_DESCRICAO            As Integer
   Dim ESPACAMENTO_POS_QTD             As Integer
   Dim ESPACAMENTO_POS_UNITARIO        As Integer
   ' _ primeira linha
   Dim ESPACAMENTO_LN_01_COL_01        As Integer
   Dim ESPACAMENTO_LN_01_COL_02        As Integer
   ' _ segunda linha
   Dim ESPACAMENTO_LN_02_COL_01        As Integer
   Dim ESPACAMENTO_LN_02_COL_02        As Integer
   Dim ESPACAMENTO_LN_02_COL_03        As Integer
   ' _ terceira linha
   Dim ESPACAMENTO_LN_03_MARGEM_ESQUERDA As Integer
   Dim ESPACAMENTO_LN_03_COL_01        As Integer
   Dim ESPACAMENTO_LN_03_COL_02        As Integer
   Dim ESPACAMENTO_LN_03_ESPACAMENTO   As Integer
   Dim ESPACAMENTO_LN_03_COL_03        As Integer
   
   '# CABEÇARIO DA LISTA
      '--- reconfigura layout
      Usa_Estilo_Padrao
      Montagem.FontBold = True
      
      '--- padroniza texto
      ' _ define espaçamento para cada formato
      If Cupom Then
         ESPACAMENTO_POS_COD = 2
         ESPACAMENTO_POS_DESCRICAO = 3
         ESPACAMENTO_POS_QTD = 4
         ESPACAMENTO_POS_UNITARIO = 5
      ElseIf A5 Then
         ESPACAMENTO_POS_COD = 3
         ESPACAMENTO_POS_DESCRICAO = 27
         ESPACAMENTO_POS_QTD = 8
         ESPACAMENTO_POS_UNITARIO = 12
      ElseIf A4 Then
         ESPACAMENTO_POS_COD = 5
         ESPACAMENTO_POS_DESCRICAO = 57
         ESPACAMENTO_POS_QTD = 5
         ESPACAMENTO_POS_UNITARIO = 6
      End If
      Cabecario_da_Lista = _
                  "CÓDIGO" & Space$(ESPACAMENTO_POS_COD) _
                  & "DESC." & Space$(ESPACAMENTO_POS_DESCRICAO) _
                  & "QTD" & Space$(ESPACAMENTO_POS_QTD) _
                  & "UNITÁRIO" & Space$(ESPACAMENTO_POS_UNITARIO) _
                  & "TOTAL"
      '--- printa cabecário
         Montagem.Print Cabecario_da_Lista

      Usa_Estilo_Padrao
   ' fim : # CABEÇARIO DA LISTA
   
   '# LISTA
      If Cupom Then
         ' _ primeira linha
         ESPACAMENTO_LN_01_COL_01 = 8
         ' _ segunda linha
         ESPACAMENTO_LN_02_COL_01 = 20
         ESPACAMENTO_LN_02_COL_02 = 8
         ESPACAMENTO_LN_02_COL_03 = 13
         ' _ terceira linha
         ESPACAMENTO_LN_03_MARGEM_ESQUERDA = 18
         ESPACAMENTO_LN_03_COL_01 = 10
         ESPACAMENTO_LN_03_COL_02 = 10
      ElseIf A5 Then
         ' _ primeira linha
         ESPACAMENTO_LN_01_COL_01 = 9
         ' _ segunda linha
         ESPACAMENTO_LN_02_COL_01 = 44
         ESPACAMENTO_LN_02_COL_02 = 16
         ESPACAMENTO_LN_02_COL_03 = 17
         ' _ terceira linha
         ESPACAMENTO_LN_03_MARGEM_ESQUERDA = 43 ' min=20
         ESPACAMENTO_LN_03_COL_01 = 13 ' min=13
         ESPACAMENTO_LN_03_COL_02 = 18 ' min=18
      ElseIf A4 Then
         ' _ primeira linha
         ESPACAMENTO_LN_01_COL_01 = 9
         ' _ segunda linha
         ESPACAMENTO_LN_02_COL_01 = 11
         ESPACAMENTO_LN_02_COL_02 = 19
         ' _ terceira linha
         ESPACAMENTO_LN_03_MARGEM_ESQUERDA = 0
         ESPACAMENTO_LN_03_COL_01 = 9
         ESPACAMENTO_LN_03_COL_02 = 76
      End If
   
      For Each Item In pLista_de_Itens
      '--- define variáveis
      With Item
         Codigo = .getField("Cod_Produto")
         Descricao = .getField("Descricao")
         Qtd = .getField("Qtd")
         total_a_pagar = .getField("Total_a_Pagar")
         Valor_Desconto = .getField("Valor_Desconto")
         Desconto = .getField("Desconto")
         Unidade = .getField("Unidade")
         mostrar_desconto = .getField("Mostrar_Desconto")
         valor_por_produto = .getField("Unitario")
         total_por_produto = .getField("Total_a_Pagar")
         If mostrar_desconto Then
            valor_por_produto = .getField("Unitario_Produto")
            total_por_produto = .getField("Total_Puro")
         End If
         
         Item_Primeira_Linha = _
                  Codigo & _
                  Space$(ESPACAMENTO_LN_01_COL_01 - Len(Codigo)) & _
                  Descricao
         If A4 Then
            Item_Primeira_Linha = Space$(2) & _
               Codigo & Space$(ESPACAMENTO_LN_01_COL_01 - Len(Codigo)) & Space$(2) & _
               Descricao & Space$(ESPACAMENTO_LN_01_COL_02 - Len(Descricao)) & Space$(2) & _
               Space$(ESPACAMENTO_LN_02_COL_01 - Len(Qtd & " " & Unidade)) & Qtd & " " & Unidade & Space$(4) & _
               Space$(ESPACAMENTO_LN_02_COL_02 - Len(FormatNumber(valor_por_produto, 2))) & FormatNumber(valor_por_produto, 2) & Space$(1) & _
               Space$(ESPACAMENTO_LN_02_COL_03 - Len(FormatNumber(total_por_produto, 2))) & FormatNumber(total_por_produto, 2)
         End If

         Item_Segunda_Linha = _
            Space$(ESPACAMENTO_LN_02_COL_01 - Len(Qtd & " " & Unidade)) & _
            Qtd & " " & Unidade & _
            Space$(ESPACAMENTO_LN_02_COL_02 - Len(FormatNumber(valor_por_produto, 2))) & _
            FormatNumber(valor_por_produto, 2) & _
            Space$(ESPACAMENTO_LN_02_COL_03 - Len(FormatNumber(total_por_produto, 2))) & _
            FormatNumber(total_por_produto, 2)
         
         If A4 Then
            Item_Segunda_Linha = _
               Space$(ESPACAMENTO_LN_03_MARGEM_ESQUERDA) & "DESCONTO: R$ " & FormatNumber(Valor_Desconto, 2) & _
               Space$(ESPACAMENTO_LN_03_COL_01 - Len(FormatNumber(Valor_Desconto, 2))) & Space$(1) & _
               "(" & FormatNumber(Desconto, 2) & "%)" & Space$(ESPACAMENTO_LN_03_COL_02 - Len(FormatNumber(Desconto, 2))) & Space$(1) & _
               String(ESPACAMENTO_LN_03_ESPACAMENTO - Len(FormatNumber(total_por_produto, 2)), ".") & Space$(1) & _
               FormatNumber(total_por_produto, 2)
         End If

         Item_Terceira_Linha = _
            Space$(ESPACAMENTO_LN_03_MARGEM_ESQUERDA - Len("DESCONTO: " & FormatNumber(Valor_Desconto, 2))) & _
            "DESCONTO: " & FormatNumber(Valor_Desconto, 2) & _
            Space$(ESPACAMENTO_LN_03_COL_01 - Len(FormatNumber(Desconto, 2))) & _
            "(" & FormatNumber(Desconto, 2) & "%)" & _
            Space$(ESPACAMENTO_LN_03_COL_02 - Len(FormatNumber(total_a_pagar, 2))) & _
            FormatNumber(total_a_pagar, 2)

         '--- Monta print
         If Cupom Or A5 Then
            Montagem.Print Escreve_com_Quebra_de_Linha(Item_Primeira_Linha, Montagem.FontSize)
            Montagem.Print Item_Segunda_Linha
            If mostrar_desconto Then Montagem.Print Item_Terceira_Linha
         ElseIf A4 Then
            Montagem.Print Escreve_com_Quebra_de_Linha(Item_Primeira_Linha, Montagem.FontSize)
            If mostrar_desconto Then Montagem.Print Item_Segunda_Linha
         End If
      End With
      Next Item
   Usa_Estilo_Padrao
End Sub

Public Sub subtotal_produtos(ByRef pSubtotal As Double)
   '--- reconfigura layout
   Usa_Estilo_Padrao
   Montagem.FontBold = True
      
   '--- define variáveis
   Dim SubTotal As String
   SubTotal = FormatNumber(pSubtotal, 2)
   
   '--- monta print
      If Cupom Then
         Montagem.FontSize = 9
         Montagem.Print _
               Space$(13) & "SUBTOTAL: " & _
               Space$(13 - Len(SubTotal)) & SubTotal

      ElseIf A5 Then
         Montagem.Print _
               Space$(54) & "SUBTOTAL: " & _
               Space$(13 - Len(SubTotal)) & SubTotal

      ElseIf A4 Then
         Montagem.Print _
               Space$(94) & "SUBTOTAL: " & _
               Space$(13 - Len(SubTotal)) & SubTotal
      End If
   Usa_Estilo_Padrao
End Sub

Public Sub soma_do_desconto(ByRef pSoma_de_Descontos As Double)
   '--- reconfigura layout
   Usa_Estilo_Padrao
   Montagem.FontBold = True
   
   '--- define variáveis
   Dim Soma_de_Descontos As Double
   Soma_de_Descontos = pSoma_de_Descontos
   
   '--- monta print
      If Cupom Then
         Montagem.FontSize = 9
         Montagem.Print _
               Space$(13) & "DESCONTO: " & _
               Space$(13 - Len(FormatNumber(Soma_de_Descontos, 2))) & FormatNumber(Soma_de_Descontos, 2)

      ElseIf A5 Then
         Montagem.Print _
               Space$(54) & "DESCONTO: " & _
               Space$(13 - Len(FormatNumber(Soma_de_Descontos, 2))) & FormatNumber(Soma_de_Descontos, 2)

      ElseIf A4 Then
         Montagem.Print _
               Space$(97) & "DESCONTO: " & _
               Space$(10 - Len(FormatNumber(Soma_de_Descontos, 2))) & FormatNumber(Soma_de_Descontos, 2)
      End If
   Usa_Estilo_Padrao
End Sub

Public Sub total_a_pagar(ByRef pTotal_a_Pagar As Double)
   '--- reconfigura layout
   Usa_Estilo_Padrao
   Montagem.FontBold = True
   Montagem.FontSize = 9
   
   '--- define variáveis
   Dim total_a_pagar As Double
   total_a_pagar = FormatNumber(pTotal_a_Pagar, 2)
   
   '--- monta print
      If Cupom Then
         Montagem.Print _
               Space$(7) & "TOTAL A PAGAR: " & _
               Space$(13 - Len(FormatNumber(total_a_pagar, 2))) & FormatNumber(total_a_pagar, 2)
      ElseIf A5 Then
         Montagem.Print _
               Space$(49) & "TOTAL A PAGAR: " & _
               Space$(13 - Len(FormatNumber(total_a_pagar, 2))) & FormatNumber(total_a_pagar, 2)
      ElseIf A4 Then
         Montagem.Print _
               Space$(92) & "TOTAL A PAGAR: " & _
               Space$(10 - Len(FormatNumber(total_a_pagar, 2)) & FormatNumber(total_a_pagar, 2))
      End If
End Sub

Public Sub recebimento_em_credito_loja(ByRef pCredito_Loja_Disponivel As Double, _
                        ByRef pCredito_Loja_Debitado As Double, _
                        ByRef pCredito_Loja_Remanescente As Double, _
                        ByRef pTotal_a_Pagar As Double)
   '--- define variáveis
   Dim Credito_Loja_Disponivel As Double
   Dim Credito_Loja_Debitado As Double
   Dim Credito_Loja_Remanescente As Double
   Dim total_a_pagar As Double
   Credito_Loja_Disponivel = FormatNumber(pCredito_Loja_Disponivel, 2)
   Credito_Loja_Debitado = FormatNumber(pCredito_Loja_Debitado, 2)
   Credito_Loja_Remanescente = FormatNumber(pCredito_Loja_Remanescente, 2)
   total_a_pagar = FormatNumber(pTotal_a_Pagar, 2)

   '--- monta print
      If Cupom Then
         
         Usa_Estilo_Padrao
         Montagem.Print "--- RECEBIDO EM CRÉDITO LOJA ------------"
         
         Montagem.FontBold = True
         Montagem.FontSize = 9
         
         Montagem.Print _
               Space$(1) & "CRÉDITO LOJA DISPONÍVEL: " & _
               Space$(10 - Len(Credito_Loja_Disponivel)) & Credito_Loja_Disponivel
         Montagem.Print _
               Space$(10) & "TOTAL DEBITADO: " & _
               Space$(10 - Len(Credito_Loja_Debitado)) & Credito_Loja_Debitado
         Montagem.Print _
               Space$(4) & "CRÉDITO REMANESCENTE: " & _
               Space$(10 - Len(Credito_Loja_Remanescente)) & Credito_Loja_Remanescente
         If total_a_pagar > Credito_Loja_Debitado Then
            Montagem.Print String$(41, "-")
            Montagem.Print _
                  Space$(20) & "TOTAL A PAGAR ATUALIZADO: " & _
                  Space$(10 - Len(total_a_pagar) - Len(Credito_Loja_Debitado)) & total_a_pagar - Credito_Loja_Debitado
         End If
         Montagem.Print String(41, "-")

      ElseIf A5 Then
         Montagem.Print "--- RECEBIDO EM CRÉDITO LOJA ---"
         Montagem.Print _
               Space$(17) & "CRÉDITO LOJA DISPONÍVEL: " & _
               Space$(10 - Len(Credito_Loja_Disponivel)) & Credito_Loja_Disponivel
         Montagem.Print _
               Space$(16) & "TOTAL DEBITADO: " & _
               Space$(10 - Len(Credito_Loja_Debitado)) & Credito_Loja_Debitado
         Montagem.Print _
               Space$(24) & "CRÉDITO REMANESCENTE: " & _
               Space$(10 - Len(Credito_Loja_Remanescente)) & Credito_Loja_Remanescente
         If total_a_pagar > Credito_Loja_Debitado Then
            Montagem.Print String$(80, "-")
            Montagem.Print _
                  Space$(16) & "TOTAL A PAGAR ATUALIZADO: " & _
                  Space$(10 - Len(total_a_pagar) - Len(Credito_Loja_Debitado)) & total_a_pagar - Credito_Loja_Debitado
         End If
         Montagem.Print String(80, "-")

      ElseIf A4 Then
         Montagem.Print "--- RECEBIDO EM CRÉDITO LOJA ---"
         Montagem.Print _
               Space$(78) & "CRÉDITO LOJA DISPONÍVEL: " & _
               Space$(11 - Len(Credito_Loja_Disponivel)) & Credito_Loja_Disponivel
         Montagem.Print _
               Space$(79) & "TOTAL DEBITADO: " & _
               Space$(11 - Len(Credito_Loja_Debitado)) & Credito_Loja_Debitado
         Montagem.Print _
               Space$(75) & "CRÉDITO REMANESCENTE: " & _
               Space$(11 - Len(Credito_Loja_Remanescente)) & Credito_Loja_Remanescente
         If total_a_pagar > Credito_Loja_Debitado Then
            Montagem.Print "--------------------------------"
            Montagem.Print _
                  Space$(70) & "TOTAL A PAGAR ATUALIZADO: " & _
                  Space$(11 - Len(total_a_pagar) - Len(Credito_Loja_Debitado)) & total_a_pagar - Credito_Loja_Debitado
         End If
         Montagem.FontSize = 8
         Montagem.Print String(116, "-")
         
      End If
      Usa_Estilo_Padrao
End Sub

Public Sub recebimento_em_dinheiro( _
                        ByRef pValor_Recebido_em_Dinheiro As String, _
                        ByRef pTroco As String)
   '--- reconfigura layout
   Usa_Estilo_Padrao
   Montagem.FontBold = True
   
   '--- define variáveis
   Dim Valor_Recebido_em_Dinheiro As String
   Dim Troco As String
   Valor_Recebido_em_Dinheiro = FormatNumber(pValor_Recebido_em_Dinheiro, 2)
   Troco = FormatNumber(pTroco, 2)

   '--- monta print
      If Cupom Then
         Montagem.Print _
               Space$(16) & "DINHEIRO: " & Space$(10 - Len(Valor_Recebido_em_Dinheiro)) & Valor_Recebido_em_Dinheiro
         Montagem.Print _
               Space$(19) & "TROCO: " & Space$(10 - Len(Troco)) & Troco

      ElseIf A5 Then
         Montagem.Print _
               Space$(60) & "DINHEIRO: " & Space$(10 - Len(Valor_Recebido_em_Dinheiro)) & Valor_Recebido_em_Dinheiro
         Montagem.Print _
               Space$(63) & "TROCO: " & Space$(10 - Len(Troco)) & Troco

      ElseIf A4 Then
         ' agurdando implementação
      End If
   Usa_Estilo_Padrao
End Sub

Public Sub parcelas_e_formas_de_pagamento(ByRef pLista_de_Itens() As ctlFicha, ByRef pNum_de_Parcelas As Integer)
   Usa_Estilo_Padrao
   '--- define variáveis
   Dim Parcelamento As Boolean
   Dim Esta_Pago As Boolean
   Dim Forma_de_Pagamento As String
   Dim Tipo_de_Item As String
   Dim N_Parcela As String
   Dim Valor As String
   Dim Vencimento As String
   Dim Primeira_Linha As String
   Dim Item As Variant
   Dim Eh_Parcela As Boolean

   '--- monta print
      If Cupom Then
         Montagem.FontSize = 10
         Escreve_Centralizado "FORMA DE PAGAMENTO" & Chr$(10)
         
         
      ElseIf A5 Then
         Montagem.FontSize = 10
         Escreve_Centralizado "FORMA DE PAGAMENTO" & Chr$(10)
         
      ElseIf A4 Then
         Montagem.FontSize = 12
         Escreve_Centralizado "FORMA DE PAGAMENTO" & Chr$(10)
         
      End If
      Usa_Estilo_Padrao
      
      ' percorre ficha
      For Each Item In pLista_de_Itens
         With Item
            '--- valida se é parcela
            Eh_Parcela = False
            Tipo_de_Item = "VENDA: "
            If pNum_de_Parcelas > 1 Then
               Eh_Parcela = True
               Tipo_de_Item = "PARCELA: "
            End If
   
            '--- valida se o item está pago
            Esta_Pago = False
            If Len(.getField("Data_Pago")) > 0 Then Esta_Pago = True
   
            '--- obtém dados
            Forma_de_Pagamento = .getField("Forma")
            N_Parcela = .getField("N_Parcela")
            Valor = FormatNumber(.getField("valor"), 2)
            Vencimento = Format$(.getField("Vencimento"), "dd/MM/yy")
   
               '--- define primeira linha: trata-se sobre está pago ou não
               Primeira_Linha = _
                  Tipo_de_Item & _
                  Left$(N_Parcela, 5) & Space$(5 - Len(Left$(N_Parcela, 5)))
   
               If Esta_Pago Then
                  Primeira_Linha = Primeira_Linha & " PAGO"
               Else
                  Primeira_Linha = Primeira_Linha & "VENCIMENTO: " & Vencimento
               End If
               
               '--- printa primeira linha
               Montagem.Print Primeira_Linha
   
               '--- define e printa segunda linha: trata-se sobre forma de pagamento e valor
               Montagem.Print _
                     "FORMA: " & Left$(Forma_de_Pagamento, 14) & Space$(14 - Len(Left$(Forma_de_Pagamento, 14))) & _
                     Space$(4) & _
                     "VALOR: " & Space$(9 - Len(Valor)) & Valor
   
         End With
      Next Item
      Usa_Estilo_Padrao
End Sub

Public Sub rodape()
      If Cupom Then
         If Len(MSG_Rodape_Venda_Str) > 41 Then
            Montagem.Print Escreve_com_Quebra_de_Linha(MSG_Rodape_Venda_Str, TAMANHO_FONTE_PADRAO_CUPOM)
         Else
            Escreve_Centralizado MSG_Rodape_Venda_Str
         End If
      ElseIf A5 Or A4 Then
         Montagem.Print MSG_Rodape_Venda_Str
      End If
End Sub

Public Sub assinatura()
   Usa_Estilo_Padrao
   Montagem.FontSize = 8
   Dim Comprimento_da_Linha As Integer
   Dim i As Byte
   
   If Cupom Then
      Comprimento_da_Linha = 41
   ElseIf A5 Then
      Comprimento_da_Linha = 81
   ElseIf A4 Then
      Comprimento_da_Linha = 116
   End If

   For i = 1 To 7
      Montagem.Print
   Next i
   Montagem.Print String$(Comprimento_da_Linha, "_")
   Montagem.Print
   Escreve_Centralizado "ASSINATURA"
   Montagem.Print
   Montagem.Print String$(Comprimento_da_Linha, "-")
   Usa_Estilo_Padrao
End Sub

Public Sub limite_e_saldo_de_credito_loja()
   Usa_Estilo_Padrao
   Montagem.FontSize = 10
   ' fix para tratamento de tipo das variáveis
   If Len(Limite_Compra_Str) = 0 Then Limite_Compra_Str = "0"
   If Len(Limite_Compra_Resta_Str) = 0 Then Limite_Compra_Resta_Str = "0"
   ' end fix
   Montagem.Print "LIMITE: " & CCur(Limite_Compra_Str) & Space$(2) & "SALDO: " & CCur(Limite_Compra_Resta_Str)

   Montagem.FontSize = 8
   Montagem.Print String(41, "-")
   Usa_Estilo_Padrao
End Sub

Public Sub agradecimento()
      Escreve_Centralizado "AGRADECEMOS A PREFERÊNCIA !"
End Sub

Public Sub assinatura_da_egos()
      Escreve_Centralizado "EGOS Soluções  |  Versão: " & Versao_Atual_Sistema_Str
End Sub

Public Sub mensagem_promocional(ByRef pMensagem_Promocional As String)
   Usa_Estilo_Padrao
   Montagem.FontSize = 10

   Montagem.Print Escreve_com_Quebra_de_Linha(pMensagem_Promocional, 8, 2, 2)
   
   Usa_Estilo_Padrao
End Sub

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
   
   '--- configura e inicia as variaveis
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
         Obter_Quantidade_de_Caracteres_por_Linha = 41
         Case "9"
         Obter_Quantidade_de_Caracteres_por_Linha = 37
         Case "11"
         Obter_Quantidade_de_Caracteres_por_Linha = 30
         Case Else
         Obter_Quantidade_de_Caracteres_por_Linha = 41
      End Select
   ElseIf A5 Then
      Select Case pTamanho_Fonte
         Case "8"
         Obter_Quantidade_de_Caracteres_por_Linha = 87
         Case "9"
         Obter_Quantidade_de_Caracteres_por_Linha = 80
         Case "11"
         Obter_Quantidade_de_Caracteres_por_Linha = 70
         Case Else
         Obter_Quantidade_de_Caracteres_por_Linha = 80
      End Select
   ElseIf A4 Then
      Select Case pTamanho_Fonte
         Case "8"
         Obter_Quantidade_de_Caracteres_por_Linha = 116
         Case "9"
         Obter_Quantidade_de_Caracteres_por_Linha = 105
         Case "10"
         Obter_Quantidade_de_Caracteres_por_Linha = 95
         Case "11"
         Obter_Quantidade_de_Caracteres_por_Linha = 86
         Case "12"
         Obter_Quantidade_de_Caracteres_por_Linha = 79
         Case Else
         Obter_Quantidade_de_Caracteres_por_Linha = 118
      End Select
   End If
End Function

Private Sub Usa_Estilo_Padrao()
   Montagem.FontName = FONT_NAME_PADRAO
   Montagem.FontBold = FONT_BOLD_PADRAO
   If Cupom Then Montagem.FontSize = TAMANHO_FONTE_PADRAO_CUPOM
   If A5 Then Montagem.FontSize = TAMANHO_FONTE_PADRAO_A5
   If A4 Then Montagem.FontSize = TAMANHO_FONTE_PADRAO_A4
End Sub

