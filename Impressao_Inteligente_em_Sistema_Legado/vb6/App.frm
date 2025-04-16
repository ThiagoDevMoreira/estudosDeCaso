' app
' Suponha que o código a seguir é um trecho de um formulário
' que vai consumir a funcionalidade de impressao.
private Impressao_obj as cImpressao
private Lista as cListaDeCompras

public sub btnImprimir_Click()
    '--- Define variável para representar os dados de
    '    entrada do usuário por meio da UI.
    dim nome_Lista_De_Compra as string
    
    '--- Define variável que representa a identidade da lista
    '    de compra para a busca no banco de dados.
    dim id_Lista_De_Compra as long

    '--- Obtém dados da UI
    nome_Lista_De_Compra = trim(cboFields(0).text)
    
    '--- Busca o id da lista
    id_Lista_De_Compra = Lista.get_ID(nome_Lista_De_Compra)

    '--- Chama a funcionalidade de impressao
    Imprimir id_Lista_De_Compra
end sub

private sub Imprimir (ByRef pIdListaDeCompra as long)
    '--- Instancia a funcionalidade de impressão
    set Impressao_obj = new cImpressao
    
    '--- Chama o método imprimir
    Impressao_obj.imprimir( _
                     tipoListaDeCompra, _
                     pIdListaDeCompra, _
                     formatoA5, _
                     2 _
                     )
    
    '--- Libera a funcionalidade de impressão
    set Impressao_obj = Nothing
end sub
