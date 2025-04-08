VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iEstrategia_de_Impressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'# interface comum a todas as etratégias de impressao
Public Sub prepara( _
            ByRef pDados As cPedido_Agregado, _
            ByRef pFormato As Formato_de_Impressao_Enu, _
            ByRef pLayout As cLayout_de_Impressao)
   ' Apenas a assinatura, sem implementação
End Sub

Public Sub executa()
' Apenas a assinatura, sem implementação
End Sub
