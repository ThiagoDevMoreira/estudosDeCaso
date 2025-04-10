VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'# classe cImpressao
' do ponto de vista da arquitetura MVC, esta � a classe prncipal da estrutura do 'Controller'
' do ponto de vista do pattern esta � classe principal do strategy, que cumpre o papel de 'Context'
' esta � a classe que deve ser chamada pelo sistema sempre que precisar utilizar a funcionalidade
' de impress�o, salvo quando se trata de relat�rios vinculados � ferramentas de relat�rio externas.
'
'## propriedades
Private Estrategia_de_Impressao  As iEstrategia_de_Impressao
Private Formato                  As Formato_de_Impressao_Enu
Private Dados                    As cPedido_Agregado
Private Layout                   As cLayout_de_Impressao

'## m�todos p�blicos
Public Sub Imprimir( _
                     ByVal pTipo As Tipo_de_Cupom_Enu, _
                     ByVal pPedido As Integer, _
                     Optional ByVal pFormato As Formato_de_Impressao_Enu, _
                     Optional ByRef pVias As Integer = 1 _
                     )
   ' todas as impress�es que n�o sejam relat�rios devem chamar esta fun��o
   ' esta fun��o implementa uma estrutura "Strategy Pattern"
   ' referencia: https://refactoring.guru/pt-br/design-patterns/strategy
   '
   Dim Pedido_Obj                   As New cPedido_Agregado
   Dim Estrategia_de_Impressao_Obj  As iEstrategia_de_Impressao ' elemento "concrete strategy" do pattern
   Dim i As Integer
   
   '--- cria objeto que ser� fonte de dados para a funcionalidade de impressao
   Pedido_Obj.Ndoc = pPedido

   '--- entrega a fonte de dados para o contexto de impressao
   set_Pedido Pedido_Obj
   
   '--- permite o processamento de outros processos do sistema enquanto processa e realiza a solicita��o de impress�o
   DoEvents
   
   '--- define a estrat�gia de impress�o
   Select Case pTipo
      Case cupomTipoVenda: Set Estrategia_de_Impressao_Obj = New cEstrategia_Cupom_Venda
      Case cupomTipoCarneParcelaCupom: Set Estrategia_de_Impressao_Obj = New cEstrategia_CarneParcelaCupom
      Case cupomTipoCarneDuasParcelaCupom: Set Estrategia_de_Impressao_Obj = New cEstrategia_CarneDuasParcelaCupom
      Case cupomTipoTroca: Set Estrategia_de_Impressao_Obj = New cEstrategia_Cupom_Troca
      Case cupomTipoPromocional: Set Estrategia_de_Impressao_Obj = New cEstrategia_Cupom_Promocional
      Case Else: Err.Raise 5, , "Estrategia_de_Impressao_Obj inv�lido!"
   End Select
   
   '--- define o formato da impressao
   Select Case pFormato
      Case formatoCupom: set_Formato_de_Impressao (formatoCupom)
      Case formatoA4: set_Formato_de_Impressao (formatoA4)
      Case formatoA5: set_Formato_de_Impressao (formatoA5)
      Case Else: Err.Raise 5, , "set_Formato_de_Impressao inv�lido!"
   End Select

   '--- entrega estrat�gia de impress�o para o contexto
   set_Estrategia_de_Impressao Estrategia_de_Impressao_Obj
   DoEvents
   
   For i = 1 To pVias
      executa_Impressao
   Next i
   Screen.MousePointer = 0
End Sub

Private Sub executa_Impressao()
   DoEvents ' permite que o sistema continue respondendo enquanto processa a impress�o
   Set Layout = New cLayout_de_Impressao
   With Estrategia_de_Impressao
      
      .prepara Dados, Formato, Layout
      DoEvents
      DoEvents
      .executa ' Layout.get_Impressao
   
   End With
End Sub

Private Sub set_Pedido(ByRef pPedido As cPedido_Agregado)
   Set Dados = pPedido
End Sub

Private Sub set_Formato_de_Impressao(ByRef pFormato As Formato_de_Impressao_Enu)
   Formato = pFormato
End Sub

Private Sub set_Estrategia_de_Impressao(ByRef pEstrategia_de_Impressao As iEstrategia_de_Impressao)
   Set Estrategia_de_Impressao = pEstrategia_de_Impressao
End Sub
