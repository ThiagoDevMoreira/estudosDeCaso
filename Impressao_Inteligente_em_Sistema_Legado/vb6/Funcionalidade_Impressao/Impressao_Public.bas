Option Explicit

Enum Formato_de_Impressao_Enu
   formatoCupom
   formatoA5
   FormatoA4
End Enum

Enum Tipo_de_Impressao_Enu
   tipoListaDeCompra
   tipoListaDeTroca
   tipoReciboDeVenda
   tipoCupomPromocional
End Enum

Public Function Seleciona_Formato_de_Impressao( _
                     ByRef pConfig_RS As adodb.recordSet) As Formato_de_Impressao_Enu
'--- valida retorno do banco de dados e a rela��o entre as flags
   With pConfig_RS
      If Not ( _
         !Tipo_Impressao_Cupom = 1 _
         Xor !Tipo_Impressao_A5 = 1 _
         Xor !Tipo_Impressao_A4 = 1 _
         ) Then MsgBox "N�o foi poss�vel selecionar o formato de impress�o.": Exit Function
   
      '--- define formato
      Seleciona_Formato_de_Impressao = formatoCupom
      If !Tipo_Impressao_A5 = 1 Then Seleciona_Formato_de_Impressao = formatoA5
      If !Tipo_Impressao_A4 = 1 Then Seleciona_Formato_de_Impressao = FormatoA4
   End With
End Function
