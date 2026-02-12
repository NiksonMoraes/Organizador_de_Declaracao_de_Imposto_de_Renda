# Organizador_de_Declaracao_de_Imposto_de_Renda (Excel + VBA)

Formato: Excel .xlsm com macros e UI “app‑like” (menu lateral com botões, navegação entre abas, e ações para incluir/remover registros), pensado para organizar e consolidar informações da declaração do IRPF sem dar “cara de planilha”. [Organizado...o de Renda | Excel]

Sumário

#visão-geral
#principais-recursos
#arquitetura-do-projeto

#abas-e-conteúdos
#camada-vba-módulos-e-padrões


#requisitos
#instalação--configuração
#como-usar
#validações--regras-de-negócio
#boas-práticas-de-ux-app-like
#segurança-privacidade-e-log
#convenções-de-nome--organização
#testes--garantia-de-qualidade
#roadmap
#licença


Visão Geral

Objetivo: centralizar dados cadastrais, informes bancários e entradas de receita em um fluxo guiado, reduzindo erros e facilitando a preparação da declaração.
Experiência: menu lateral com ícones/botões, Anterior/Próximo em cada tela e ações rápidas para incluir/remover lançamentos.
Estrutura: abas TÍTULAR, INFORMES, NOTAS, TABELAS e _LOG_BANCOS, com exemplos preenchidos (TOTAL = 500.000; Banco 33 – Santander; anexo topazao_2025.pdf; lançamento HOLERITE com data serial 46061). [Organizado...o de Renda | Excel]


Principais Recursos

Menu lateral com imagens e botões que navegam entre as abas por link de referência ou macro (OnAction), minimizando a aparência “Excel”. [Organizado...o de Renda | Excel]
Botões “Anterior/Próximo” em cada aba para fluxo linear (wizard). [Organizado...o de Renda | Excel]
Aba INFORMES com ações:

Incluir novo banco: copia e cola o molde/seleção para a próxima linha, preservando estrutura.
Remover banco: exclui a seleção e registra no log.
(Implementado via módulos e funções VBA.)


Validações de dados por lista (ex.: SIM/NÃO e categorias), e link mailto no e‑mail do titular. [Organizado...o de Renda | Excel]
Catálogo de bancos em TABELAS para padronizar seleção (código + nome, incluindo Banco Topázio – 82). [Organizado...o de Renda | Excel]
Aba de log para trilha de auditoria das operações de bancos (oculta por padrão). [Organizado...o de Renda | Excel]


Arquitetura do Projeto
Abas e Conteúdos


TÍTULAR
Campos de PF (Nome, CPF, Nascimento, Título de Eleitor, Cônjuge, Endereço/CEP, Telefones, E‑mail) e três seletores SIM/NÃO (alterações da entrega anterior, dependente cônjuge, residente no exterior).

E‑mail com link mailto para abertura do cliente de e‑mail.
Validações tipo SIM/NÃO em seleção guiada. [Organizado...o de Renda | Excel]



INFORMES
Área para informes de rendimentos bancários por Banco (ex.: 33 – Banco Santander), Valor Atual (há totalização) e Anexo (ex.: topazao_2025.pdf).

Botões: Inserir novo banco e Remover banco (VBA).
Observação: há um cabeçalho com typo (“VALOR ALTUAL”) — preferível “Valor Atual”. [Organizado...o de Renda | Excel]



NOTAS
Lançamentos de entradas (ex.: HOLERITE) com Data, Categoria e Valor; algumas datas podem estar em número de série (p. ex., 46061 → 08/02/2026). [Organizado...o de Renda | Excel]


TABELAS
Catálogo de bancos (código + nome) que serve de referência para INFORMES (padronização e validação). [Organizado...o de Renda | Excel]


_LOG_BANCOS
Planilha técnica para auditar inclusões/remoções de bancos (estrutura de colunas: ID, DataHora, SheetName, DestAddress, C1…D3). Oculta por padrão. [Organizado...o de Renda | Excel]


Camada VBA (Módulos e Padrões)

Organização sugerida — refletindo o que seu projeto já faz e facilitando manutenção:



modNav (Navegação):

Navigate("ABA") para ir direto à aba pelo botão do menu.
NextSheet / PrevSheet para os botões Próximo/Anterior, baseados em um array de ordem de abas.



modUI (Menu & Layout):

Alinhamento automático de shapes do menu por nome (ex.: prefixo btnMenu_ + título da aba).
Destaque visual do item ativo (cor de fundo/texto) em Worksheet_Activate.



modInformes (Ações de Dados):

AddBankRow / RemoveSelectedBankRow para incluir/remover linhas em INFORMES (preferível usar Tabela do Excel/ListObject para robustez).
LogBank gravando data/hora, origem e conteúdo em _LOG_BANCOS.




Observação: como o arquivo é .xlsm, o projeto contém macros (vbaProject) que precisam estar com macros habilitadas no Excel para que o menu e os botões funcionem. [Organizado...o de Renda | Excel]


Requisitos

Microsoft Excel (desktop) — Office 2016+ recomendado.
Macros habilitadas (arquivo .xlsm).
Permissão para editar e executar macros no dispositivo.


Instalação & Configuração

Baixe/abra o arquivo .xlsm.
Ao abrir, habilite as macros (barra de segurança).
(Opcional) Adicione a pasta do projeto como Local Confiável no Centro de Confiabilidade.
Confira se TABELAS contém a lista de bancos atualizada; INFORMES deve ler essa lista (Validação de Dados). [Organizado...o de Renda | Excel]


Como Usar

Menu lateral: clique nos botões para ir às seções; use Anterior/Próximo para seguir o fluxo. [Organizado...o de Renda | Excel]
TÍTULAR: preencha os dados; os campos SIM/NÃO usam lista; o e‑mail é clicável (mailto). [Organizado...o de Renda | Excel]
INFORMES:

Incluir novo banco → cria uma linha com base no molde/seleção; preencha Banco (pelo catálogo), Valor Atual e Anexo.
Remover banco → exclui a linha selecionada e registra no log.
O TOTAL atualiza conforme os lançamentos. [Organizado...o de Renda | Excel]


NOTAS: registre Data (em dd/mm/aaaa), Categoria (lista) e Valor; se houver datas em número de série (ex.: 46061), converta para a data legível. [Organizado...o de Renda | Excel]


Validações & Regras de Negócio

SIM/NÃO: campos binários em TÍTULAR usam lista (evita digitação livre). [Organizado...o de Renda | Excel]
Categorias de NOTAS: controladas por lista (ex.: HOLERITE/CNPJ/FREELANCE), expansível. [Organizado...o de Renda | Excel]
Banco em INFORMES: seguir código + nome conforme TABELAS (padronização). [Organizado...o de Renda | Excel]
Totais: INFORMES possui totalização (e.g., soma dos valores atuais). [Organizado...o de Renda | Excel]


Boas Práticas de UX “App‑like”

Esconder linhas de grade/cabeçalhos e padronizar zoom para reforçar o visual de aplicativo.
Manter tema de cores consistente no menu e nas telas; usar ícones leves.
Proteção de planilha: impedir mover objetos; liberar apenas o necessário.
Feedback visual (botão ativo destacado) e mensagens objetivas para entradas inválidas.


Segurança, Privacidade e Log

Dados sensíveis (CPF/CEP/telefones) devem ser validados e mascarados em relatórios compartilháveis.
_LOG_BANCOS registra ações de inclusão/remoção com data/hora e contexto, facilitando auditoria (oculto por padrão). [Organizado...o de Renda | Excel]


Convenções de Nome & Organização

Shapes do menu: prefixo btnMenu_XX_<ABA> (ex.: btnMenu_02_TITULAR) para permitir alinhamento automático e destaque por nome.
Módulos VBA:

modNav (navegação),
modUI (menu & layout),
modInformes (incluir/remover + log).


Tabelas (ListObjects): recomendado nomear a área de INFORMES como tbInformes (simplifica o VBA).
Intervalos nomeados: ListaBancos apontando para TABELAS (validação de dados).


Testes & Garantia de Qualidade

Navegação: menu e Anterior/Próximo em todas as abas. [Organizado...o de Renda | Excel]
Validações: listas SIM/NÃO e Categorias funcionando (rejeitar texto livre). [Organizado...o de Renda | Excel]
Conversão de datas: garantir que datas tipo 46061 apareçam como dd/mm/aaaa após conversão. [Organizado...o de Renda | Excel]
Totalização: verificar atualização do TOTAL em INFORMES após incluir/remover linhas. [Organizado...o de Renda | Excel]
Log: confirmar registro em _LOG_BANCOS a cada ação (quando macros ativas). [Organizado...o de Renda | Excel]
Proteção: usuário só deve editar campos previstos.


Roadmap

Converter INFORMES em Tabela (ListObject) para robustez das macros e do log.
Checklist de documentos (anexos) por banco com status visual.
Resumo financeiro (por categoria/mês) com gráfico + exportação para PDF.
Validações adicionais (CPF/CEP/telefones) e mascaramento automático em exportações.
Dashboard inicial com KPIs e atalhos.

