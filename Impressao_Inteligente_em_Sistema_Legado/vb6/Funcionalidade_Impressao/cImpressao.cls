Option Explicit
'# classe cImpressao
'  Do ponto de vista da arquitetura MVC, esta � a classe principal da estrutura do 'Controller',
'  do ponto de vista de design pattern esta � classe principal do pattern strategy, que cumpre 
'  o papel 'Context'. Esta � a classe que deve ser chamada pelo app sempre que precisar utilizar
'  a funcionalidade de impress�o.
'  referencia: https://refactoring.guru/pt-br/design-patterns/strategy

'## propriedades
Private Estrategia_de_Impressao  As iEstrategia_de_Impressao
Private Formato                  As Formato_de_Impressao_Enu
Private Dados                    As Object
Private Layout                   As cLayout_de_Impressao

'## m�todo p�blicos, dando aten��o que seja apenas um m�todo exposto para o app.
Public Sub Imprimir( _
                     ByVal pTipo As Tipo_de_Impressao_Enu, _
                     ByVal pId As long, _
                     Optional ByVal pFormato As Formato_de_Impressao_Enu, _
                     Optional ByRef pVias As Integer = 1 _
                     )
   
   '--- cria objeto que ser� fonte de dados para a funcionalidade de impressao
   set Dados = obtem_Objeto_de_Dados(pTipo, pId)
   
   '--- permite o processamento de outros processos do sistema enquanto processa e realiza
   '    a solicita��o de impress�o
   DoEvents
   
   '--- define a estrat�gia de impress�o
   Select Case pTipo
      case tipoListaDeCompra: set Estrategia_de_Impressao_Obj = new cEstrategia_ListaDeCompra
      case tipoListaDeTroca: set Estrategia_de_Impressao_Obj = new cEstrategia_ListaDeTroca
      '...
   End Select
   
   '--- define o formato da impressao
   Formato = pFormato
   
   '--- solicita a impressao
   For i = 1 To pVias
      executa_Impressao
   Next i
End Sub

Private Sub executa_Impressao()
   DoEvents ' permite que o sistema continue respondendo enquanto processa a impress�o
   Set Layout = New cLayout_de_Impressao
   ' cada estrat�gia pode ter sua pr�pria forma de executar a impress�o.
   With Estrategia_de_Impressao
      .prepara Dados, Formato, Layout
      .executa
   End With
End Sub

'--- seleciona a fonte de dados adequada para cada tipo de impress�o.
private function obtem_Objeto_de_Dados(byref pTipo as Tipo_de_Impressao_Enu, byref pId as long) as object
   dim Dados as object
   select case pTipo
      case tipoListaDeCompra
         set dados = new cLista_De_Compras
         dados.get_Lista_de_Compras pId

      case tipoListaDeTroca
         set dados = new cLista_De_Troca
         dados.get_Lista_de_Troca pId
      '...
   end select
   set obtem_Objeto_de_Dados = Dados
end function