# Ordens de Serviço, o Segredo do Baile:

## Apresentação

O desafio atual é (re)criar uma aplicação embutida num sistema maior, um projeto que chegou quase no fim, morreu na praia por abandono do desenvolvedor anterior e por ter empacado nos testes de usabilidade...
Trata-se de uma aplicação destinada a gestão de Ordens de Serviço (OS), desde a entrada da solicitação do cliente até a emissão de notas fiscais de produtos e de serviços.
Neste momento, o escopo da aplicação é atender oficinas mecânicas, mas faz parte dos requisitos preparar a arquitetura para que aceite facilmente ampliação para outros nichos.

## Ao analisar o que já estava pronto no sistema encontrei o enrosco

O sistema não estava lidando bem com uma das principais características, naturais, das ordens de serviço: o baile das idas e voltas da OS entre a empresa e cliente.

*Exemplo do baile:*

* Cliente leva o carro para o mecânico e relata o problema.

* Mecânico abre uma ordem de serviço, verifica o carro e...

* Manda um parecer técnico e uma proposta de serviço para o cliente aprovar (início do baile).

* Cliente responde com uma contraproposta (2o. passo do baile).

* Mecânico ajusta prazo e valores e manda para o cliente aprovar (3o. passo do baile).

* Vamos supor que o cliente aceite (4o. passo do baile).

* O mecânico começa a executar a manutenção e...

* Encontra outro problema atrelado à manutenção contratada.

* Então o mecânico avisa o cliente e propõem a manutenção adicional (reinicia o baile).

* ...

## O plano

Dentro de uma arquitetura gerral MVC, no sentido estrutural, vou adotar a abordagem do padrão Decorator para garantir uma estrutura que possa ser ampliada a aplicação para novos nichos num futuro próximo. (Porque Decorator em vez de outro padrão? Manda um alô, a gente conversa sobe isso thiago.dev.moreira@gmail.com)

Para o baile escolhi a abordagem do padrão State assim temos uma grande liberdade para refinar cada etapa do ciclo cliente-empresa-cliente. Ao mesmo tempo essa abordagem garante um lugar confortável para ter um fluxo "personalizado" para cada novo nicho, reaproveitando a base de código sem deteriorá-lo.

**Agora mãos à obra!...*
