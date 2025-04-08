VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEstrategia_CarneDuasParcelaCupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iEstrategia_de_Impressao

'# Carne_Duas_Parcela_Cupom
Public Sub iEstrategia_de_Impressao_prepara( _
            ByRef pDados As cPedido_Agregado, _
            ByRef pFormato As Formato_de_Impressao_Enu, _
            ByRef pLayout As cLayout_de_Impressao _
            )
   Dim Str_Ficha_Detalhes As New adodb.recordSet
   If CompileErro Then On Error GoTo ERRO
   If MsgBox("Imprimir Carne em Parcelas ?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR") <> vbYes Then Exit Sub
   Printer.FontName = "Courier New"
   If sConn.State = 0 Then sConn.Open
   With Str_Ficha_Detalhes
      If .State = 1 Then .Close
      .Open "SELECT * " _
         & "FROM Ficha " _
         & "WHERE ndoc=" & pDados.getField("Ndoc"), sConn, adOpenStatic, adLockReadOnly, adCmdText
      Set .ActiveConnection = Nothing
   End With
   frmImprimir_Tudo.Frame1.Visible = False
   frmImprimir_Tudo.Frame2.Visible = False
   frmImprimir_Tudo.Frame3.Visible = False
   frmImprimir_Tudo.Frame1.Top = 0
   frmImprimir_Tudo.Frame1.Left = 0
   frmImprimir_Tudo.Frame1.Visible = True
   frmImprimir_Tudo.Frame1.ZOrder 0
   frmImprimir_Tudo.Frame1.height = 433
   With Str_Ficha_Detalhes
      .MoveFirst
      Do Until .Eof
         If Nome_Loja2_Str <> "" Then
            frmImprimir_Tudo.Label10.Caption = Trim$(Nome_Loja_Str & Chr$(10) & Nome_Loja2_Str)
            frmImprimir_Tudo.Label24.Caption = Trim$(Nome_Loja_Str & Chr$(10) & Nome_Loja2_Str)
         Else
            frmImprimir_Tudo.Label10.Caption = Trim$(Nome_Loja_Str)
            frmImprimir_Tudo.Label24.Caption = Trim$(Nome_Loja_Str)
         End If
         frmImprimir_Tudo.Label1 = frmVenda.lblFields(2).Caption & "-" & frmVenda.lblFields(9).Caption
         frmImprimir_Tudo.Label2 = Format$(!Vencimento, "dd/MM/yy")
         frmImprimir_Tudo.Label3 = Format$(!N_Parcela, "@@@@@")
         frmImprimir_Tudo.Label4 = FormatCurrency(!Valor, 2)
         frmImprimir_Tudo.Label5 = frmVenda.txtFields(6).Text
         frmImprimir_Tudo.Label12.Caption = "!" + (Format$(frmVenda.txtFields(6).Text, "#")) + "!"
         frmImprimir_Tudo.Label29 = frmVenda.lblFields(2).Caption & "-" & frmVenda.lblFields(9).Caption
         frmImprimir_Tudo.Label28 = Format$(!Vencimento, "dd/MM/yy")
         frmImprimir_Tudo.Label27 = Format$(!N_Parcela, "@@@@@")
         frmImprimir_Tudo.Label26 = FormatCurrency(!Valor, 2)
         frmImprimir_Tudo.Label25 = frmVenda.txtFields(6).Text
         frmImprimir_Tudo.Label37.Caption = "!" + (Format$(frmVenda.txtFields(6).Text, "#")) + "!"
         If Format$(!Data, "dd/MM/yy") = Format$(!Vencimento, "dd/MM/yy") Then
            frmImprimir_Tudo.Label14.Visible = True
            frmImprimir_Tudo.Label23.Visible = True
         Else
            frmImprimir_Tudo.Label14.Visible = False
            frmImprimir_Tudo.Label23.Visible = False
         End If
         SetParent frmImprimir_Tudo.hWnd, frmImprimir_Tudo.Frame1.hWnd
         On Error Resume Next
         frmImprimir_Tudo.PrintForm
         .MoveNext
      Loop
   End With
   Exit Sub
ERRO:
   MsgBox Err.Description, vbInformation, "Imprimir_Carne_Duas_Parcela_Cupom"
   Screen.MousePointer = 0
End Sub

Public Sub iEstrategia_de_Impressao_executa()
   'Printer.EndDoc
End Sub



