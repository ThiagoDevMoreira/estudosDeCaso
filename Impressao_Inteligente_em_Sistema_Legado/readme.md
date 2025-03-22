# Estudo de caso:
## Impressão Inteligente em Sistema Legado
Em VB6 temos o objeto `printer`: é nativo da linguagem, global e único. A própria linguagem nunca cria mais do que um único objeto `printer`. Na tentativa de `set este_eh_um_novo_printer = new printer`, a variável `este_eh_um_novo_printer` torna-se apenas uma nova referência ao único objeto `printer` do sistema.

É o `printer` que recebe todas as indicações de tamanho e características de estilo a ser utilizado (com muitas restrições! Considerando a capacidade de estilizar texto atualmente) e não tem métodos nativos para alinhar o texto nem para quebra de linhas.

> Se você pedir quebra de linha para o printer, ele vai responder: "faça suas próprias contas de quantos caracteres cabem na linha e se vira! Lhe desejo sorte ao tratar com separações de palavras e sílabas no fim da linha. hehehe"

Utiliza-se assim:

```vb6
' configurações de conexão com a impressora
printer.font_name = "arial"
printer.font_size = 12
printer.bold = True
printer.print "Esta é a linha impressa: " & com_esta_variavel
printer.end ' esta é a linha que manda o comando para a impressora
```

Note que são necessárias muitas linhas de código para obter uma única linha impressa.

Agora imagina o quanto pode escalar a complexidade, a repetição de código, o tempo de desenvolvimento e ainda mais o tempo de manutenção num sistema monolítico, procedural, que precisa imprimir muitos tipos de documentos comerciais (cupom de venda, orçamento, parcelamento...) e relatórios (histórico de compras dos clientes, cheques a compensar, estoque...), com layot diferente para cada formato de papel - cupom, A4, A5... -, solicitados pelas mais diversas partes do sistema local ou remoto sob o crivo da funcionalidade de permissões. E ainda enviados para diversas impressoras sejam locais na rede local ou via requisição web.

Para o VB6 existem recursos não nativos que podem ser adicionados para facilitar a funcionalidade de impressão, mas ainda assim igualmente legados e ainda assim seu uso é tão trabalhoso ou mais do que a abordagem nativa.

Lembre que ambos trabalhos: solução nativa ou solução externa, são enormes porém de perfís totalmente diferentes. Por um lado o trabalho de desenvolver uma solução com os recursos nativos, por outro as o trabalho de adaptar o sistema atual para integrar a solução exerna.

Acho que os mais novos não conseguem visualizar o tamanho da enrascada. Lembrem-se estamos na seara dos monolíticos legados e deteriorados pelo tempo...

Neste estudo de caso, vamos considerar a decisão do cliente de não comprar nenhuma solução externa e solicitar que a dedicação em tempo seja a menor possível (como sempre na maioria das empresas).

Evidentemente que o código em questão começou com um simples e ingênuo:

```vb6
printer.font_name = "arial"
printer.font_size = 12
printer.bold = True
printer.print "o valor deste produto é: " & valor_do_produo_com_desconto & "."
if exibir_desconto then
	printer.print "este produto tem: " & valor_do_desconto & "% de desconto."
end if
printer.end
```

Mas, ao longo dos anos acumularam-se as mudanças, soluções e ajustes simples e práticos.

Deparei-me com 2^3 condições `if-elseif-else` cada uma delas com outras 2^4 como "sub-condições", onde se emaranhavam os códigos de formatação (muitas vezes definidas no corpo principal do código, outras definidas em funções externas), com dados vindo diretamente do banco de dados, outros dados vindos de flags globais do sistema, outros dados vindos diretamente da interface gráfica... e ainda validações e matemáticas correspondentes às regras de negócio de cada uma das condições, onde a lógica, o propósitos e as responsabilidades eram bichos estranhos.

Motivação para o trabalho: cliente final reclamou que "as contas no papel impresso não batem com a transação comercial realizada"

Desafio aceito, vamos construir um refúgio no meio deste campo agreste. Adoro "sujar as mãos", prazo de 10 dias úteis, ótimo "bastante tempo" incuído nesse tempo o plantão de hotfixes, a realização de bugfixes prioritários e o tempo necessário para o pessoal de QA fazer os testes não-automatizados.

Objetivo principal: fazer que as "contas batam" no caso apontado pelo cliente.
Objetivo secundário: que a solução para o caso específico possa ser reutilizada para todos os casos de impressão do sistema, sem repetição de código e cuidando da sanidade emocional dos próximos desenvolvedores que visitem esta região.

Acredito que os desenvolvedores anteriores se preocuparam apenas com o objetivo principal, seja por falta de prazo adequado, capacidade ou pagamento...

Planejamento!

Dia 1: dia de avaliação, análise e mapeamento.
Dia 2: definição da abordagem: arquitetura MVC, estrutura inspirada no padrão Strategy.

	Obs sobre a implementaçaõ do MCV: Neste caso, o View é o próprio objeto VB.Printer (pessoalmente acho que aqui está o maior pulo do gato! A cereja do bolo! O tal do pensar fora da caixa). O Controller é a classe com função de "Contexto" do Strategy (as classes de estratégia devem ser consideradas conceitualmente como extenção do controller, pois vão consumir o model, coisa que é feita naturalmente pelo controller) e o Model é a classe que lida com as entidades do banco de dados e implementa a "matemática" das regras de negócio. Nesta estrutura específica o controller/contexto recebe as solicitações do sistema, conduz e organiza o fluxo e delega todo o processamento para as estratégias.

	Tarefas:
	A. Desenvolver uma nova estrutura de classes para a funcionalidade de impressão conforme o padrão Strategy para a gestão de das solicitações de impressão e ser consumida pelo sistema sempre que precisar de uma impressão.
	
	B. Desenvolver uma classe para cada uma das "estratégias" de impressão.
	
	C. Desenvolver uma classe que exponha as "peças" que vão compor o layout de impressão (como se fosse uma biblioteca de componentes de frontend) para a elaboração granlar do layout final que será impresso. Esta biblioteca é consumida pelas classes de "estratégias de impressão". A relação entre esta biblioteca e as estratégias se dá nos moldes do padrão Template.
		
	B.Desenvolver uma classe que lide com as entidades do banco de dados para incluir as validações e campos calculados necessários à funcionalidade de impressão, o model é consumido pelas estratégias, esta é a classe model (mvc) e é consumida pelas estratégias.
	
Dia 3 ao 10: Mãos na massa, fazendo vb6. Com um plano destes qualqer linguagem funciona! ;)
