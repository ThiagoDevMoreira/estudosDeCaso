' cListaDeCompras
'# classe cListaDeCompras
'  Esta é a classe controller da entidade "ListaDeCompras" pela qual são expostos à app
'  os dados e funcionalidades relacionadas à entidade ListaDeCompras.
'  A entidade em si estaria definida no model mListaDeCompras, esse model não será apresentado
'  aqui por não fazer parte do contexto deste estudo de caso, vou mater o foco na aplicação do
'  design pattern strategy
private mLista as mLista_de_Compras
Private Lista as cLista_De_Compras

public function busca_Id(byref pTermoDeBusca as variant) as long
    set Lista = new mLista_de_Compras
    with mLista
        if isnumeric(pTermoDeBusca) then
            pTermoDeBusca = trim(pTermoDeBusca)
            if .existe_Id_Lista(pTermoDeBusca) then
                get_Id = clng(pTermoDeBusca)
            end if
        elseif typename(pTermoDeBusca) = "string" then
            pTermoDeBusca = trim(lcase(pTermoDeBusca))
            if .existe_Nome_Lista(pTermoDeBusca) then
                get_Id = .obter_Id_Pelo_Nome(pTermoDeBusca)
            end if
        Else
            get_Id = -1
        end if
    end With
end function

public function get_Lista(byref pIdLista as long) as mLista_de_Compras
    set Lista = new mLista_de_Compras
end functiong