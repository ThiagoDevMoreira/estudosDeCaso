VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEstrategia_CarneParcelaCupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iEstrategia_de_Impressao

'# Carne_Parcela_Cupom
Public Sub iEstrategia_de_Impressao_prepara( _
            ByRef pDados As cPedido_Agregado, _
            ByRef pFormato As Formato_de_Impressao_Enu, _
            ByRef pLayout As cLayout_de_Impressao _
            )
   Dim Str_Ficha_Detalhes As New adodb.recordSet
   If CompileErro Then On Error GoTo ERRO
   If MsgBox("I M P R I M I R   C A R N Ê   E M   P A R C E L A S   ?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR") <> vbYes Then Exit Sub
   Printer.FontName = "Courier New"
   If sConn.State = 0 Then sConn.Open
   With Str_Ficha_Detalhes
      If .State = 1 Then .Close
      .Open "SELECT * " _
         & "FROM Ficha " _
         & "WHERE ndoc=" & frmVenda.txtFields(6), sConn, adOpenStatic, adLockReadOnly, adCmdText
      Set .ActiveConnection = Nothing
      frmImprimir_Tudo.Frame1.Visible = False
      frmImprimir_Tudo.Frame2.Visible = False
      frmImprimir_Tudo.Frame3.Visible = False
      frmImprimir_Tudo.Frame1.Top = 0
      frmImprimir_Tudo.Frame1.Left = 0
      frmImprimir_Tudo.Frame1.Visible = True
      frmImprimir_Tudo.Frame1.ZOrder 0
      frmImprimir_Tudo.Frame1.height = 200
      .MoveFirst
      Do Until .Eof
         If Nome_Loja2_Str <> "" Then
            frmImprimir_Tudo.Label10.Caption = Trim$(Nome_Loja_Str & Chr$(10) & Nome_Loja2_Str)
         Else
            frmImprimir_Tudo.Label10.Caption = Trim$(Nome_Loja_Str)
         End If
         frmImprimir_Tudo.Label1 = frmVenda.lblFields(2).Caption & "-" & frmVenda.lblFields(9).Caption
         frmImprimir_Tudo.Label2 = Format$(!Vencimento)
         frmImprimir_Tudo.Label3 = Format$(!N_Parcela, "@@@@@")
         frmImprimir_Tudo.Label4 = FormatNumber(!Valor, 2)
         frmImprimir_Tudo.Label5 = frmVenda.txtFields(6).Text
         frmImprimir_Tudo.Label12.Caption = "!" + (Format$(frmVenda.txtFields(6).Text, "#")) + "!"
         If Format$(!Data, "dd/MM/yy") = Format$(!Vencimento, "dd/MM/yy") Then
            frmImprimir_Tudo.Label14.Visible = True
         Else
            frmImprimir_Tudo.Label14.Visible = False
         End If
         SetParent frmImprimir_Tudo.hWnd, frmImprimir_Tudo.Frame1.hWnd
         On Error Resume Next
         frmImprimir_Tudo.PrintForm
         .MoveNext
      Loop
   End With
   Exit Sub
ERRO:
   MsgBox Err.Description, vbInformation, "Imprimir_Carne_Parcela_Cupom"
   Screen.MousePointer = 0
End Sub

Public Sub iEstrategia_de_Impressao_executa()
   'Printer.EndDoc
End Sub

