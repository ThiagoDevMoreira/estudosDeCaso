VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEstrategia_Cupom_Promocional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'# Imprimir_Promocao()
Implements iEstrategia_de_Impressao

'# classe cEstrategia_Cupom_Promocional (estratégia concreta)
Public Sub iEstrategia_de_Impressao_prepara( _
            ByRef pDados As cPedido_Agregado, _
            ByRef pFormato As Formato_de_Impressao_Enu, _
            ByRef pLayout As cLayout_de_Impressao _
            )
   Dim Existe_End_Entrega  As Boolean: Existe_End_Entrega = False
   Dim A4                  As Boolean: A4 = False
   Dim A5                  As Boolean: A5 = False
   Dim Cupom               As Boolean: Cupom = False
   
   ' Lida com peculiaridade desta estratégia.
   On Error GoTo CASO_VALOR_OU_COMPONENTE_DA_TELA_NAO_ESTEJA_ACESSIVEL
   Dim Valor_Compra As Double
   Valor_Compra = Fix(CCur(frmVenda.lblFields(51).Caption / Valor_Promocional_Dbl))
   GoTo FLUXO_NORMAL
CASO_VALOR_OU_COMPONENTE_DA_TELA_NAO_ESTEJA_ACESSIVEL:
   MsgBox "falha em cEstrategia_Cupom_Promocional": Exit Sub
FLUXO_NORMAL:
   On Error GoTo 0

   '--- Seta flag de formato local e passa o formato para a biblioteca de layout
   Select Case pFormato
      Case FormatoA4
         A4 = True
      Case formatoA5
         A5 = True
      Case formatoCupom
         Cupom = True
   End Select
   pLayout.set_Formato pFormato

   '--- Descreve a estratégia
   With pLayout
      Dim i As Integer
   
      For i = 1 To Valor_Compra
         .nome_da_empresa
         .data_da_impressao
         .separador
         
         .linha_em_branco
         .titulo_maior "CUPOM"
         .titulo_maior "SORTEIO"
         .separador
         
         .linha_em_branco
         .linha_em_branco
         .titulo "CÓDIGO"
         .num_pedido pDados.getField("Ndoc")
         .separador
         
         .linha_em_branco
         .linha_em_branco
         .cod_e_nome_cliente _
                        pDados.obterCliente.getField("Cod_Cliente"), _
                        pDados.obterCliente.getField("Nome_Cliente")
         .endereco_cliente _
                        pDados.obterCliente.getField("Endereco"), _
                        pDados.obterCliente.getField("Numero"), _
                        pDados.obterCliente.getField("Bairro"), _
                        pDados.obterCliente.getField("Cidade"), _
                        pDados.obterCliente.getField("Estado")
         .telefone_cliente _
                        pDados.obterCliente.getField("Telefone"), _
                        pDados.obterCliente.getField("Celular")
         .cpf_cliente pDados.obterCliente.getField("Cpf")
         .separador
         
         .linha_em_branco
         .linha_em_branco
         .titulo_menor "TOTAL COMPRA"
         .titulo_menor "R$" & pDados.getField("Total_a_Pagar")
         .separador
         
         .linha_em_branco
         .linha_em_branco
         If Len(Mensagem_Promocional_Str) > 0 Then
            .mensagem_promocional Mensagem_Promocional_Str
         End If
      Next i
   End With
End Sub

Public Sub iEstrategia_de_Impressao_executa()
   Printer.EndDoc
End Sub



