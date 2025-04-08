VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEstrategia_Cupom_Troca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'# Imprimir_Cupom_Troca()
Implements iEstrategia_de_Impressao

'# classe cEstrategia_Cupom_Venda (estratégia concreta)
Public Sub iEstrategia_de_Impressao_prepara( _
            ByRef pDados As cPedido_Agregado, _
            ByRef pFormato As Formato_de_Impressao_Enu, _
            ByRef pLayout As cLayout_de_Impressao _
            )
   Dim Existe_End_Entrega  As Boolean: Existe_End_Entrega = False
   Dim A4                  As Boolean: A4 = False
   Dim A5                  As Boolean: A5 = False
   Dim Cupom               As Boolean: Cupom = False
   
   ' Lida c peculiaridade desta estratégia.
   On Error GoTo CASO_VALOR_OU_COMPONENTE_DA_TELA_NAO_ESTEJA_ACESSIVEL
   Dim Nome_Funcionario As String
   Nome_Funcionario = frmVenda.cboFields(8).Text
   GoTo FLUXO_NORMAL
CASO_VALOR_OU_COMPONENTE_DA_TELA_NAO_ESTEJA_ACESSIVEL:
   MsgBox "falha em cEstrategia_Cupom_Troca": Exit Sub
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
      .nome_da_empresa
      .titulo "COMPROVANTE PARA TROCA"
      .linha_em_branco
      .num_pedido pDados.getField("ndoc")
      .separador
      .data_da_impressao
      .vendedor_operador Nome_Funcionario
      .linha_em_branco
   End With
End Sub

Public Sub iEstrategia_de_Impressao_executa()
   Printer.EndDoc
End Sub

