
Option Explicit

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
   
   ' Lida com as peculiaridades desta estratégia

   On Error GoTo CASO_VALOR_OU_COMPONENTE_DA_TELA_NAO_ESTEJA_ACESSIVEL
   Dim Troco As String
   Dim Valor_Recebido_em_Dinheiro As String
   Dim Nome_Funcionario As String
   Troco = frmVenda.lblFields(23).Caption
   Valor_Recebido_em_Dinheiro = frmVenda.txtFields(9).Text
   Nome_Funcionario = frmVenda.cboFields(8).Text
   GoTo FLUXO_NORMAL
CASO_VALOR_OU_COMPONENTE_DA_TELA_NAO_ESTEJA_ACESSIVEL:
   Troco = "0"
   Valor_Recebido_em_Dinheiro = "0"
FLUXO_NORMAL:
   On Error GoTo 0

   '--- Seta flag de formato local e passa o formato para a biblioteca de layout
   Select Case pFormato
      Case formatoA4
         A4 = True
      Case formatoA5
         A5 = True
      Case formatoCupom
         Cupom = True
   End Select
   pLayout.set_Formato pFormato

   '--- Descreve a estratégia
   With pLayout
      .logo
      .nome_da_empresa
      .Cnpj
      .endereco_empresa
      .telefone_empresa
      .email_empresa
      .separador
      
      .aviso_nao_eh_doc_fiscal
      .separador
      
      .data_da_impressao
      .vendedor_operador Nome_Funcionario
      .separador
      
      .num_pedido pDados.getField("Ndoc")
      .separador
      
      .cod_e_nome_cliente _
                  pDados.getField("Cod_Cliente"), _
                  pDados.getField("Nome_Cliente")
      .cpf_cliente pDados.obterCliente.getField("Cpf")
      .linha_em_branco
      .endereco_cliente _
                  pDados.obterCliente.getField("Endereco"), _
                  pDados.obterCliente.getField("Numero"), _
                  pDados.obterCliente.getField("Bairro"), _
                  pDados.obterCliente.getField("Cidade"), _
                  pDados.obterCliente.getField("Estado")
      .telefone_cliente _
                  pDados.obterCliente.getField("Telefone"), _
                  pDados.obterCliente.getField("Celular")
      .linha_em_branco
      .endereco_de_entrega pDados.obterCliente.getField("Endereco_Entrega")
      .observacao pDados.getField("Obs_Pedidos")
      .separador
      
      .lista_produtos_venda pDados.obterPedidos_Detalhes
      .separador
      
      If pDados.getField("flagF12") And pDados.getField("Tem_Desconto") Then
         .subtotal_produtos pDados.getField("Sub_Total")
         .soma_do_desconto pDados.getField("Somatorio_Descontos")
      End If
      
      .total_a_pagar pDados.getField("Valor_Pago") ' valor a cobrado pela venda;
      
      If pDados.getField("Forma_Pago") = "CREDITO LOJA" Then
         .recebimento_em_credito_loja _
                  pDados.getField("Credito_Loja_Disponivel"), _
                  pDados.getField("Credito_Loja_Debitado"), _
                  pDados.getField("Credito_Loja_Remanescente"), _
                  pDados.getField("Total_a_Pagar")
      End If
      
      If pDados.getField("Forma_Pago") = "DINHEIRO" Then
         .recebimento_em_dinheiro _
                  Valor_Recebido_em_Dinheiro, _
                  Troco
      End If
      .separador
      
      .parcelas_e_formas_de_pagamento _
                  pDados.obterFichas, _
                  pDados.getField("Parcelas_Count")
      .separador
      
      If MSG_Rodape_Venda_Str <> "" Then
         .linha_em_branco
         .rodape
         .separador
      End If
      
      If Via_Assinatura_Str = "1" Then
         .assinatura
      End If

      If pDados.getField("Mostrar_Limite_Cupom") = "1" _
         And pDados.getField("Cod_Cliente") <> "1" _
         And Limite_Compra_Str > "0" _
         Then
         .limite_e_saldo_de_credito_loja
      End If
      .linha_em_branco
      .agradecimento
      .assinatura_da_egos
   End With
End Sub

Public Sub iEstrategia_de_Impressao_executa()
   Printer.EndDoc
End Sub

