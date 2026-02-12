# Organizador_de_Declaracao_de_Imposto_de_Renda (Excel + VBA)

Formato: Excel .xlsm com macros e UI “app‑like” (menu lateral com botões, navegação entre abas, e ações para incluir/remover registros), pensado para organizar e consolidar informações da declaração do IRPF sem dar “cara de planilha”. [Organizado...o de Renda | Excel]

Visão Geral

Objetivo: centralizar dados cadastrais, informes bancários e entradas de receita em um fluxo guiado, reduzindo erros e facilitando a preparação da declaração.
Experiência: menu lateral com ícones/botões, Anterior/Próximo em cada tela e ações rápidas para incluir/remover lançamentos.
Estrutura: abas TÍTULAR, INFORMES, NOTAS, TABELAS e _LOG_BANCOS, com exemplos preenchidos (TOTAL = 500.000; Banco 33 – Santander; anexo topazao_2025.pdf; lançamento HOLERITE com data serial 46061). [Organizado...o de Renda | Excel]


Principais Recursos

Menu lateral com imagens e botões que navegam entre as abas por link de referência ou macro (OnAction), minimizando a aparência “Excel”.
Botões “Anterior/Próximo” em cada aba para fluxo linear.

Aba INFORMES com ações:
Incluir novo banco: copia e cola o molde/seleção para a próxima linha, preservando estrutura.
Remover banco: exclui a seleção e registra no log.
(Implementado via módulos e funções VBA.)

Validações de dados por lista (ex.: SIM/NÃO e categorias), e link no e‑mail do titular.
Catálogo de bancos em TABELAS para padronizar seleção (código + nome).
Aba de log para trilha de auditoria das operações de bancos (oculta por padrão).


Como Usar

Menu lateral: clique nos botões para ir às seções; use Anterior/Próximo para seguir o fluxo.
TÍTULAR: preencha os dados; os campos SIM/NÃO usam lista; o e‑mail é clicável. 
INFORMES:
Incluir novo banco → cria uma linha com base no molde/seleção; preencha Banco (pelo catálogo), Valor Atual e Anexo.
Remover banco → exclui a linha selecionada e registra no log.
O TOTAL atualiza conforme os lançamentos.

NOTAS: registre Data (em dd/mm/aaaa), Categoria (lista) e Valor; se houver datas em número de série (ex.: 46061), converta para a data legível.

Instalação & Configuração

Baixe/abra o arquivo .xlsm.
Ao abrir, habilite as macros (barra de segurança).
(Opcional) Adicione a pasta do projeto como Local Confiável no Centro de Confiabilidade.
Confira se TABELAS contém a lista de bancos atualizada; INFORMES deve ler essa lista (Validação de Dados).


Arquitetura do Projeto
Abas e Conteúdos

TÍTULAR
Campos de PF (Nome, CPF, Nascimento, Título de Eleitor, Cônjuge, Endereço/CEP, Telefones, E‑mail) e três seletores SIM/NÃO (alterações da entrega anterior, dependente cônjuge, residente no exterior);
E‑mail com link mailto para abertura do cliente de e‑mail;
Validações tipo SIM/NÃO em seleção guiada.

INFORMES
Área para informes de rendimentos bancários por Banco (ex.: 33 – Banco Santander), Valor Atual (há totalização) e Anexo (ex.: topazao_2025.pdf).
Botões: Inserir novo banco e Remover banco (VBA).

NOTAS
Lançamentos de entradas (ex.: HOLERITE) com Data, Categoria e Valor; algumas datas podem estar em número de série (p. ex., 46061 → 08/02/2026).

TABELAS
Catálogo de bancos (código + nome) que serve de referência para INFORMES (padronização e validação).

_LOG_BANCOS
Planilha técnica para auditar inclusões/remoções de bancos (estrutura de colunas: ID, DataHora, SheetName, DestAddress, C1…D3). Oculta por padrão.

Observação: como o arquivo é .xlsm, o projeto contém macros (vbaProject) que precisam estar com macros habilitadas no Excel para que o menu e os botões funcionem.


Validações & Regras de Negócio

SIM/NÃO: campos binários em TÍTULAR usam lista (evita digitação livre).
Categorias de NOTAS: controladas por lista (ex.: HOLERITE/CNPJ/FREELANCE), expansível.
Banco em INFORMES: seguir código + nome conforme TABELAS (padronização). 
Totais: INFORMES possui totalização (e.g., soma dos valores atuais). 


Boas Práticas de UX 

Esconder linhas de grade/cabeçalhos e padronizar zoom para reforçar o visual de aplicativo.
Manter tema de cores consistente no menu e nas telas; usar ícones leves.
Proteção de planilha: impedir mover objetos; liberar apenas o necessário.
Feedback visual (botão ativo destacado) e mensagens objetivas para entradas inválidas.


Testes & Garantia de Qualidade

Navegação: menu e Anterior/Próximo em todas as abas.
Validações: listas SIM/NÃO e Categorias funcionando (rejeitar texto livre).
Conversão de datas: garantir que datas tipo 46061 apareçam como dd/mm/aaaa após conversão.
Totalização: verificar atualização do TOTAL em INFORMES após incluir/remover linhas.
Log: confirmar registro em _LOG_BANCOS a cada ação (quando macros ativas). 
Proteção: usuário só deve editar campos previstos.


Roadmap (Plano de evolução)

Converter INFORMES em Tabela (ListObject) para robustez das macros e do log.
Checklist de documentos (anexos) por banco com status visual.
Resumo financeiro (por categoria/mês) com gráfico + exportação para PDF.
Validações adicionais (CPF/CEP/telefones) e mascaramento automático em exportações.
Dashboard inicial com KPIs e atalhos.

