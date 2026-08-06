# CHANGELOG

## [2.25.1] - 2026-08-06

### Adicionado
- **ISS Retido forçado por ID_CNAE 1207/1213** (`processamento_xml.py`, `iss_fortaleza_automacao.py`): nova regra de negócio baseada no `ID_CNAE` (código bruto de serviço/tributação extraído do XML, antes da classificação final via IA) — `1207` sempre força `ISS_RETIDO = "SIM"`, `1213` sempre força `"NÃO"`, com prioridade sobre a regra já existente de CNAE final (`932989910`/`900190201`). Aplicada tanto na planilha gerada por "Processar XMLs" quanto como override direto no checkbox durante a escrituração, cobrindo também planilhas já geradas antes dessa regra.

## [2.25.0] - 2026-08-06

### Adicionado
- **Suporte ao layout XML do Portal Da Paraíba em "Processar XMLs"** (`processamento_xml.py`): além do layout do Portal Nacional (já suportado), o processamento agora reconhece automaticamente, arquivo por arquivo, XMLs no padrão ABRASF NFS-e v2.02 (`CompNfse/Nfse/InfNfse`) usado pelo Portal Da Paraíba — sem exigir nenhuma seleção manual de layout. Notas canceladas no portal (`<NfseCancelamento>`) são ignoradas e logadas com esse motivo específico, em vez de entrarem na planilha. Cidades são resolvidas a partir do código IBGE do município (novo arquivo `ibge_municipios.json`, bundlado no instalador), já que esse layout só informa o código, não o nome da cidade.
- **Nova coluna TIPO_CLIENTE na planilha exportada** (`processamento_xml.py`): grava explicitamente "Pessoa Física" ou "Pessoa Jurídica" com base no prestador de cada nota (o layout do Paraíba pode identificar o prestador por CPF, além de CNPJ). As colunas cujo preenchimento a automação de escrituração exige agora ficam destacadas com cor de fundo diferente no cabeçalho da planilha gerada, facilitando identificar de relance o que não pode ficar vazio.
- **Prestador pessoa física na automação do ISS Fortaleza** (`iss_fortaleza_automacao.py`): quando o prestador é identificado como CPF (via a nova coluna TIPO_CLIENTE, com um comprimento de CPF/CNPJ como fallback para planilhas antigas), o campo "Tipo de Cliente/Fornecedor" passa a ser definido como "Pessoa Física" (antes, sempre fixo em "Pessoa Jurídica") e "Tipo do Documento Digitado" como "NFS Avulsa de outro município" — em vez do "NFS-e de Outro Município" usado para prestador CNPJ.

## [2.24.0] - 2026-08-04

### Adicionado
- **"NFS-e Nacional" para MEI estabelecido em Fortaleza/CE** (`iss_fortaleza_automacao.py`): quando o prestador é MEI e está em Fortaleza/CE — a exceção que já faz a nota ser escriturada apesar de ser da própria capital (`executar_fluxo_iss()`) — o campo "Tipo do Documento Digitado" agora seleciona "NFS-e Nacional" em vez do "NFS-e de Outro Município" padrão. Qualquer outro caso (prestador não-MEI, ou fora de Fortaleza/CE) continua usando "NFS-e de Outro Município", como sempre. A decisão fica registrada em `log_execucao.txt` quando a exceção se aplica.

## [2.23.1] - 2026-08-04

### Corrigido
- **Número da nota fiscal não preenchido** (`iss_fortaleza_automacao.py`): regressão da v2.23.0 — esse campo tinha sido convertido para preenchimento via JavaScript (`_definir_campo_rapido`), mas é seguido de perto pela seleção de Status NFSE, que dispara reprocessamento AJAX do JSF (já documentado no código com sua própria pausa de estabilização). Sem clique/foco real no campo, o valor definido via JS provavelmente nunca era confirmado no lado servidor, e esse reprocessamento limpava o campo — mesma classe de corrida já corrigida para Natureza/ISS Retido e Cidade/CEP. Revertido para digitação real (`_digitar_campo`) e adicionada conferência/correção automática após a pausa do Status NFSE.

## [2.23.0] - 2026-08-04

### Adicionado
- **Campo Alíquota preenchido** (`iss_fortaleza_automacao.py`): o campo estava 100% ignorado pela automação, apesar da coluna `ALIQUOTA` já existir na planilha (extraída do XML por `processamento_xml.py`, tag `pAliqAplic`). `DocumentoPortalISS` e `carregar_documentos_xlsx()` agora leem essa coluna. Novo helper `_campo_aliquota_esta_editavel()` detecta se o portal bloqueou o campo para o CNAE selecionado (o portal renderiza um `<span>` em vez de `<input>` quando calcula a alíquota automaticamente) — quando bloqueado, a automação apenas registra o valor da planilha em log e segue; quando editável, apaga o valor atual e digita o da planilha.
- **Tratamento do erro "CNAE não incide Imposto Sobre Serviço"** (`iss_fortaleza_automacao.py`): esse erro de validação do portal (diferente dos modais de confirmação — é uma mensagem `<rich:messages>` sem botão, que bloqueia a gravação) agora é detectado; a automação troca a Natureza da Operação para "Não Incidência", registra a correção em log e tenta gravar novamente.
- **Tratamento da confirmação "Prestador não inscrito no CPOM"** (`iss_fortaleza_automacao.py`): generalizado o reconhecimento do modal de confirmação do portal (mesmo componente usado para nota duplicada) — qualquer confirmação que não seja identificada como nota duplicada é tratada como aviso informativo, confirma automaticamente e tenta gravar novamente, cobrindo também eventuais confirmações encadeadas na mesma nota.
- **Preenchimento rápido de campos de texto** (`iss_fortaleza_automacao.py`): CNPJ, Nome, Logradouro, Bairro, Email, Número da NF e Código CNAE agora são definidos instantaneamente via JavaScript (mesma técnica já usada para a Descrição do Serviço) em vez de digitados caractere por caractere, com verificação automática do valor final e nova tentativa por digitação como reforço caso o valor não seja aceito de primeira (vazio ou diferente do esperado).

### Corrigido
- **CEP do prestador não preenchido** (`iss_fortaleza_automacao.py`): mesma classe de corrida já corrigida para Natureza da Operação — a seleção de Cidade do Prestador dispara um AJAX do portal sem nenhuma pausa de estabilização antes de o CEP ser digitado, permitindo que uma resposta tardia do AJAX limpasse o campo depois do preenchimento. Adicionada a mesma pausa de 1.5s já usada para UF do Prestador, mais uma conferência e correção automática após uma pequena espera extra.
- **Corrida entre Natureza da Operação e checkbox ISS Retido** (`iss_fortaleza_automacao.py`): a correção anterior (v2.22.0) para o checkbox ISS Retido não bastava — o portal marca/desmarca esse checkbox sozinho ao mudar a Natureza da Operação (ex.: para "Tributação Fora do Município"), e não havia nenhuma pausa entre a seleção da Natureza e a leitura do checkbox. Adicionada pausa de estabilização de 1.5s após a seleção de Natureza, mais uma conferência e correção automática após o clique, com a divergência registrada em log quando detectada.

## [2.22.0] - 2026-08-03

### Adicionado
- **Suporte a pasta com subpastas em "Processar XMLs"** (`processamento_xml.py`, `app.py`, `gui/index.html`): nova função `processar_pasta_xmls_auto()` detecta automaticamente a estrutura da pasta selecionada. Se houver XML diretamente nela, gera 1 planilha (comportamento de sempre); se houver subpastas (ex.: uma por cliente/lote), gera 1 planilha independente dentro de cada subpasta com XML, sem misturar lotes num único arquivo — inclusive quando a pasta selecionada tem XML solto na raiz **e** subpastas ao mesmo tempo (`processar_pasta_xmls()` ganhou o parâmetro `arquivos_xml_forcados` para escopar o processamento sem recursão nesse caso). A GUI ganhou um botão "Abrir pasta com as planilhas geradas" para quando o resultado é mais de um arquivo.
- **Botão Cancelar na automação de escrituração** (`iss_fortaleza_automacao.py`, `app.py`, `gui/index.html`): novo controle ao lado de Pausar/Retomar que aborta a automação por completo — fecha o navegador e exige reinício do zero (nova seleção de planilha/competência e novo login manual). Implementado com `AutomacaoCanceladaError`, um segundo `threading.Event` (`_cancelar_automacao_event`) e checagem (`_verificar_cancelamento()`) nos mesmos pontos onde a pausa já era checada, incluindo a espera (antes ilimitada e sem nenhum ponto de interrupção) pela confirmação de login manual.
- **Detecção de nota fiscal duplicada durante a gravação** (`iss_fortaleza_automacao.py`): quando o portal exibe o aviso "Já existe documento fiscal escriturado com o CNPJ do prestador, com o mesmo número de nota, na competência ...", o robô agora reconhece esse modal automaticamente, clica em "Não" (nunca confirma a geração de uma possível duplicata sozinho) e segue para a próxima nota, registrando-a em `log_notas_duplicadas.txt`. Antes, esse modal travava a automação: o robô ficava esperando cegamente por um cabeçalho de sucesso que nunca aparecia, estourava timeout e o clique seguinte de reset de tela também falhava, por estar bloqueado pelo overlay do modal.
- **Cobertura de testes** (`tests/test_iss_retido.py`, `tests/test_processar_xmls_auto.py`, novos): cobrem a interpretação de valores de ISS Retido e o roteamento de pastas/subpastas do `processar_pasta_xmls_auto()`, incluindo o cenário de XML direto + subpastas simultâneos.

### Corrigido
- **Checkbox ISS Retido sem auditoria** (`iss_fortaleza_automacao.py`): a lógica de marcar/desmarcar já comparava o estado atual antes de clicar (não era um clique incondicional), mas uma falha no seletor do checkbox era engolida silenciosamente (`except NoSuchElementException: pass`) e a planilha não validava se o valor de `ISS_RETIDO` era um "Sim"/"Não" reconhecível. Agora `validar_campos_obrigatorios()` rejeita valores fora do padrão antes de chegar no navegador, e o estado final do checkbox é sempre registrado em log (sucesso, alerta de divergência, ou erro de elemento não encontrado).
- **Pausa nos helpers de baixo nível nunca funcionava de fato** (`iss_fortaleza_automacao.py`): a atribuição `_CALLBACK_PAUSA = callback_pausa` dentro de `executar_fluxo_iss()` não tinha a declaração `global`, então criava uma variável local que sombreava o nome do módulo — o callback nunca chegava a atualizar o global de verdade lido por `_verificar_pausa()`. Corrigido junto com a implementação do cancelamento (que usa o mesmo padrão e exigia o fix para funcionar). Na prática, a pausa só surtia efeito no checkpoint entre notas; agora também funciona nos pontos internos de digitação/clique.

## [2.21.2] - 2026-07-27

### Corrigido
- **Planilha modelo ausente após instalação** (`build.py`, `processamento_xml.py`, `.gitignore`): `EXEMPLOS/planilha exemplo.xlsx` nunca tinha sido incluída no empacotamento do PyInstaller (`build.py` só copiava `gui/` e `cnae_oficial.xlsx`) nem versionada no Git, então "Processar XMLs" falhava com "Modelo EXEMPLOS/planilha exemplo.xlsx ausente" em qualquer instalação nova. Adicionado `--add-data` para o template no `build.py`, fallback de busca via `sys._MEIPASS` em `processamento_xml.py` (mesmo padrão já usado para `cnae_oficial.xlsx`), e o template passou a ser versionado (`.gitignore` mantém os XMLs reais de clientes dentro de `EXEMPLOS/` ignorados, mas abre exceção para esse arquivo específico).

## [2.21.1] - 2026-07-24

### Corrigido (Bugs Críticos)
- **BUG-04 — Corrupção do novo CNPJ alfanumérico na escrituração** (`iss_fortaleza_automacao.py`): `limpar_cnpj_para_digitacao()` usava `re.sub(r"\D+", ...)`, que apagava as letras do novo formato de CNPJ da Receita Federal (14 caracteres, 12 alfanuméricos + 2 dígitos verificadores) antes de digitar o valor no campo `idCPFCNPJ` do portal da ISS Fortaleza. A função agora preserva letras, removendo apenas a máscara de pontuação (`.`, `-`, `/`, espaços) e normalizando para maiúsculas.
- **Verificação de digitação cega a letras** (`_digitar_campo()`): o fallback de estabilidade do campo comparava apenas dígitos, o que poderia reportar sucesso falso mesmo se o portal descartasse letras do CNPJ silenciosamente. Adicionado parâmetro opcional `normalizador_fallback` (retrocompatível para os outros 10 campos que usam a mesma função) e o call site do CNPJ agora usa um normalizador que preserva letras.
- **Leitura da coluna CNPJ_PRESTADOR do Excel** (`carregar_documentos_xlsx()`): célula numérica (CNPJ legado autoformatado como número pelo Excel) podia perder zero à esquerda; célula vazia, zero ou negativa agora retorna string vazia (novo helper `_normalizar_celula_cnpj()`), garantindo que `validar_campos_obrigatorios()` continue barrando a nota em vez de gravar um CNPJ inválido no portal.

### Adicionado
- **Cobertura de testes** (`tests/test_cnpj_normalizacao.py`, novo): primeira pasta `tests/` do projeto (via `unittest`, sem dependência nova), cobrindo CNPJ legado e alfanumérico, com e sem máscara, e os casos de borda da leitura de célula do Excel.

### Corrigido (Documentação)
- **README.md / DEVELOPER.md**: a descrição de "pausa antes da gravação para revisão humana" estava desatualizada desde que o robô passou a clicar em `GRAVAR DOCUMENTO` automaticamente nota a nota (mudança já registrada no changelog anterior). Os dois documentos agora descrevem o comportamento real: login manual obrigatório, pausar/retomar pela GUI a qualquer momento, e notas incompletas/com erro puladas e logadas em vez de gravadas.

## [2.21.0] - 2026-07-13

### Corrigido (Bugs Críticos)
- **BUG-01 — SyntaxError em `import default_keys`** (`app.py` L139): Removido espaço indevido em `import default_ keys` que causava `SyntaxError` não capturado pelo `except ImportError`, impedindo o carregamento de chaves padrão em builds de produção.
- **BUG-02 — Módulo inexistente em `main.py`**: Removida importação de `extrair_nf_pdfs` (inexistente). O `main.py` agora documenta explicitamente que o projeto opera exclusivamente via GUI e exibe mensagem de redirecionamento para `app.py`.
- **BUG-03 — Injeção de JavaScript via f-string sem sanitização** (`app.py`): Todas as chamadas a `evaluate_js()` que interpolavam dados externos (nome da planilha, competência) agora utilizam `json.dumps()` para sanitização antes da injeção.

### Corrigido (Falhas Funcionais)
- **FALHA-01 — Resource Leak de ChromeDriver** (`exportador_xml_prestados.py`): O bloco `finally` que antes estava comentado foi restaurado de forma condicional. Adicionado parâmetro `fechar_navegador_ao_fim: bool = False` que controla se o Chrome é encerrado ao fim da exportação, eliminando o acúmulo de processos zumbis.
- **FALHA-02 — Coluna `REGIME_TRIBUTARIO` duplicada** (`processamento_xml.py`): A adição da coluna no cabeçalho agora verifica se ela já existe antes de inserir, evitando duplicidade ao reprocessar a mesma pasta.
- **FALHA-03 — Race condition na flag `_automacao_em_execucao`** (`app.py`): Adicionado `threading.Lock()` (`_lock_automacao`) para proteger atomicamente a leitura, escrita e liberação da flag que controla a execução concorrente de robôs.
- **FALHA-04 — Variável global `_CALLBACK_PAUSA` sem proteção de concorrência** (`iss_fortaleza_automacao.py`): Adicionado `_LOCK_CALLBACK_PAUSA = threading.Lock()` e refatorada `_verificar_pausa()` para usar o lock, eliminando a dependência de `global` e o risco de race condition.
- **FALHA-05 — Progresso "Concluído!" exibido antes de erro de pasta vazia** (`processamento_xml.py`): Ao não encontrar XMLs na pasta selecionada, a barra de progresso agora exibe `"Erro: nenhum XML encontrado."` em vez de `"Concluído!"`.
- **FALHA-06 — Timeout de login hardcoded em 300s** (`exportador_xml_prestados.py`): Adicionado parâmetro `timeout_login: int = 300` configurável na função de exportação.
- **FALHA-07 — Anos hardcoded no seletor de exportação** (`gui/index.html`): Criada função JS `inicializarSeletorAno(idSelect)` reutilizável que popula dinamicamente o ano atual em ambos os seletores (automação e exportação) no `DOMContentLoaded`.
- **FALHA-08 — Card de login visível após erro de competência** (`app.py` + `gui/index.html`): O card de confirmação de login agora é exibido pelo Python (após os checks de competência/planilha), e não pelo JavaScript antes da chamada da API. O bloco `finally` do Python garante que o card sempre seja ocultado ao encerrar.

### Adicionado (Melhorias)
- **MELHORIA-01 — Timestamps nos logs dos terminais** (`gui/index.html`): As funções `addLogAutomacao`, `addLogExportar` e `addLogXml` agora prefixam cada mensagem com o horário local `[HH:MM:SS]`.
- **MELHORIA-02 — Botão "Limpar" nos três terminais** (`gui/index.html`): Adicionado botão com ícone `delete_sweep` no cabeçalho dos terminais das três abas (Automação, Exportação e XMLs).
- **MELHORIA-03 — Salvamento progressivo da planilha com checkpoint** (`processamento_xml.py`): A planilha XLSX é salva a cada 50 notas processadas (configurável via `checkpoint_a_cada`), evitando perda total de dados em falhas de processamento longo.
- **MELHORIA-04 — Timeout explícito nas chamadas ao Gemini** (`cnae_final.py`): Adicionada função `_chamar_gemini_com_timeout()` com `concurrent.futures.ThreadPoolExecutor` e timeout de 60 segundos, garantindo fallback local quando a API trava.
- **MELHORIA-05 — Botão "Abrir Planilha Gerada" após processamento de XMLs** (`app.py` + `gui/index.html`): Ao concluir o processamento, o Python envia o caminho da planilha via `evaluate_js()` e a GUI exibe um botão que abre o arquivo diretamente.
- **MELHORIA-06 — Validação visual do arquivo Excel antes de iniciar automação** (`app.py` + `gui/index.html`): Ao selecionar o arquivo, a GUI marca a sessão com `_arquivoExcelSelecionadoNestaSessao`. Ao clicar em "Iniciar", se o arquivo não foi selecionado nesta sessão, exibe aviso de verificação de caminho.
- **MELHORIA-07 — Refatoração da lógica duplicada de abertura do Chrome** (`iss_fortaleza_automacao.py`): Extraída função privada `_inicializar_driver(opcoes)` que centraliza a estratégia híbrida Selenium Manager → webdriver-manager. `abrir_navegador_visivel()` e `abrir_navegador_com_perfil_persistente()` agora delegam para ela, eliminando ~35 linhas duplicadas.
- **MELHORIA-08 — Unificação de `configurar_pasta_logs()` e `configurar_pastas_logs()`** (`tratamento_erros.py`): A versão singular agora é um alias que chama a versão plural internamente, eliminando a lógica duplicada e mantendo compatibilidade retroativa.
- **MELHORIA-10 — Indicação visual de processamento em andamento** (`gui/index.html`): O botão "Iniciar Processamento de XMLs" muda para "PROCESSANDO..." com ícone animado (`progress_activity` com `spin` CSS) durante a execução, e restaura o estado original ao concluir.


### Adicionado
- **Funcionalidade: Exportar XML de Serviços Prestados** (`exportador_xml_prestados.py`): Novo módulo de automação Selenium que acessa o portal da ISS Fortaleza, navega até a tela de Consulta de NFS-e, seleciona a competência (mês/ano) escolhida pelo usuário na GUI, itera pelas páginas de resultados em lotes de até 10 páginas e exporta os XMLs automaticamente para a pasta de destino selecionada.
- **Nova Aba na GUI** (`gui/index.html`): Adicionada a terceira aba "Exportar XML de Prestados" na barra lateral com seletores de mês/ano de competência, seletor de pasta de destino, card de confirmação de login manual, barra de progresso e console de log dedicado.
- **Suporte a Pasta de Download no Chrome** (`iss_fortaleza_automacao.py`): Adicionado parâmetro `pasta_downloads` em `_criar_opcoes_chrome`, `abrir_navegador_visivel` e `abrir_navegador_com_perfil_persistente`, configurando o diretório de download automático sem exibir diálogos de salvamento.
- **Novos Métodos de API** (`app.py`): Adicionados `iniciar_exportacao_xml_gui`, `_executar_exportacao_xml` e `confirmar_login_exportacao` para orquestrar a automação de exportação via pywebview.

## [2.19.1] - 2026-07-09

### Adicionado
- **Mecanismo de Cache de Classificação CNAE** (`processamento_xml.py`): Implementado cache local em memória que armazena as classificações computadas pelo Tess AI para evitar chamadas redundantes de notas de mesmo fornecedor, serviço e CNAE.

## [2.19.0] - 2026-07-09

### Adicionado
- **Extração de Regime Tributário** (`processamento_xml.py`): Adicionado mapeamento automático da tag XML `opSimpNac`. Quando o valor do campo for igual a `2`, a nota é classificada sob o regime `"MEI"`; caso contrário, é classificada como `"OUTROS"`.
- **Coluna REGIME_TRIBUTARIO** (`processamento_xml.py`): Inclusão dinâmica de uma nova coluna `"REGIME_TRIBUTARIO"` na linha de cabeçalho e de dados do arquivo Excel gerado na aba "Processar XMLs".

### Alterado
- **Lógica de Descarte Local** (`iss_fortaleza_automacao.py`): O robô do Selenium agora pula e ignora notas fiscais emitidas por prestadores cujo endereço esteja localizado em Fortaleza/CE (UF `CE` e Cidade `FORTALEZA`).
- **Exceção de Escrituração MEI** (`iss_fortaleza_automacao.py`): Caso o prestador seja classificado no regime tributário `"MEI"`, a nota é obrigatoriamente escriturada no portal, ignorando e prevalecendo sobre a regra de pulo local descrita acima.
- **Configurações e Apresentação** (`.env.example`, `README.md`): Atualizado o arquivo de variáveis de exemplo para remover chaves do Gemini/Groq e revisado o README comercial para focar estritamente na leitura de XMLs e no uso do Tess AI.
- **Script Inno Setup** (`setup.iss`): Atualizada a versão do instalador de `"2.14.9"` para `"2.19.0"`, alinhando com a versão corrente do projeto.

## [2.18.0] - 2026-07-08

### Removido
- **Módulos Físicos e Lógicas de PDF/IA Legadas**: Exclusão física dos arquivos `extrair_nf_pdfs.py`, `gemini_extracao.py`, `extracao_prefeituras.py`, `unir_planilhas.py` e `extrair_iss_uma_vez.py`.
- **APIs Gemini e Groq**: Removido completamente o suporte e qualquer menção a chaves de API ou chamadas do Gemini e Groq, migrando o foco do projeto estritamente para o Tess AI e arquivos XML.
- **Submenu e Abas Legadas**: Removidas as abas "Organizar PDFs", "Exportar por Prefeitura", "Extração Lote (Sem IA)", "Unir Planilhas" e o submenu correspondente na interface gráfica.

### Alterado
- **Automação de Escrituração** (`iss_fortaleza_automacao.py`): Adaptado o robô do Selenium para ler o novo formato da planilha (sem a coluna `PREFEITURA`) e obter o arquivo associado dinamicamente das colunas `ARQUIVO_XML` ou `ARQUIVO_PDF`.
- **Simplificação do Painel Central** (`gui/index.html`): Reestruturação da interface gráfica para expor apenas as duas funcionalidades essenciais ("Processar XMLs" como padrão e "Automação: Escrituração"). Os campos de credenciais do Tess AI foram elegantemente realocados para o rodapé do menu lateral (`aside`).
- **Limpeza de Rotas** (`app.py`): Remoção de imports e métodos legados de PDF, mantendo e otimizando apenas as rotas chamadas pelo front-end para XML e Automação ISS.
- **Recuperação de Função de Login Manual** (`app.py`): Reintroduzida a função `confirmar_login_feito`, corrigindo um `AttributeError` de inicialização e permitindo que o robô do Selenium saiba quando o operador efetuou o login manual.

## [2.17.0] - 2026-07-08

### Adicionado
- **Processamento de XMLs de NFS-e** (`processamento_xml.py`, `app.py`, `gui/index.html`): Adicionada uma nova aba "Processar XMLs" na interface gráfica. A funcionalidade permite selecionar uma pasta, fazer a varredura recursiva de arquivos XML de NFS-e (padrão nacional do SPED ou equivalentes), realizar o parse seguro ignorando namespaces, classificar os CNAEs usando a IA do Tess baseando-se nas opções oficiais de `cnae_oficial.xlsx`, e consolidar todos os dados fiscais e tributários em uma planilha Excel no modelo exato de `EXEMPLOS/planilha exemplo.xlsx` (salva na pasta selecionada).
- **Classificação CNAE via Tess AI**: Criada lógica backend que realiza a pré-filtragem local de CNAEs candidatos baseados nos termos do XML e envia à API do Tess AI para selecionar obrigatoriamente um único item do banco de CNAEs reais oficiais.

## [2.16.3] - 2026-07-07

### Corrigido
- **Recuperação de Tela** (`iss_fortaleza_automacao.py`): Implementado o fallback dinâmico no bloco `finally`. Caso ocorra qualquer falha durante o preenchimento de uma nota na tela de serviços e o botão "Novo Documento" fique inacessível (causando um "efeito cascata" que quebrava as próximas linhas), a automação agora realiza um reset preventivo, navegando pelo menu para limpar o formulário e recomeçar a digitação na tela limpa.
- **Resiliência AJAX (RichFaces)**: Ajustada a função interna de interação com selects (`_selecionar_opcao_por_texto`). Caso a lista de opções de um combo venha vazia ou com apenas o texto genérico ("Selecione"), a automação não falhará mais de imediato; ela detecta o carregamento assíncrono e passa a realizar retentativas ativas com uma latência maior, impedindo falsos erros de `NoSuchElementException` ao popular dados de "Tipo de Documento".

## [2.16.2] - 2026-07-07

### Adicionado
- **Validação Preventiva de Planilha** (`iss_fortaleza_automacao.py`): Criada verificação automática que analisa se todos os 15 campos obrigatórios de notas e cadastro do prestador estão preenchidos na planilha. Caso algum campo crucial esteja ausente (ou se o valor do serviço for menor/igual a zero), a nota fiscal é ignorada, registrando os motivos detalhados no log principal e gerando o log `log_notas_incompletas.txt` em tempo real para permitir que a automação avance sem travar no portal.

## [2.16.1] - 2026-07-07

### Melhorado
- **Build / Chaves Padrão** (`build.py` e `app.py`): O script de compilação agora extrai automaticamente a `TESS_API_KEY` e a `TESS_AGENT_ID` do `.env` local do desenvolvedor e as embute de forma segura no executável final. Quando o software for instalado na máquina do cliente e ele não tiver configurado o `.env`, a aplicação usará as chaves injetadas como fallback, permitindo o uso imediato e contínuo sem a necessidade de configuração prévia.

## [2.16.0] - 2026-07-07

### Adicionado
- **ISS Fortaleza**: Implementado fluxo de preenchimento manual do formulário de identificação do prestador de serviços.
  - A automação agora escritura notas fiscais independentemente de o prestador estar cadastrado previamente no portal da prefeitura.
  - O sistema passa a utilizar os dados de endereço do prestador extraídos pela IA e preenche todos os campos necessários manualmente na tela.
  - Remoção da dependência da lista do portal de prestadores já registrados, aumentando a taxa de sucesso da escrituração contínua.

## [2.15.3] - 2026-07-07

### Melhorado
- **Ajuste de Prompt de Local de Prestação** (`gemini_extracao.py`): Removida a instrução de que a IA deveria priorizar informações próximas ao termo "PRESTADOR" para os campos `uf_local_prestacao` e `cidade_local_prestacao`. Isso evita que a IA confunda erroneamente o local de prestação de serviços com o endereço do próprio prestador, deixando a extração do local de prestação real do serviço mais precisa.

## [2.15.2] - 2026-07-07

### Melhorado
- **Flexibilização de Modelos no Tess AI** (`gemini_extracao.py`): Removida a chave `"model"` hardcoded nas requisições do Tess AI. Com isso, o sistema agora utiliza de forma dinâmica qualquer modelo configurado diretamente por você no painel web do Agente 49525 (como o *Claude 5 Sonnet* selecionado na interface).

## [2.15.1] - 2026-07-07

### Corrigido
- **Processamento de Arquivos no Tess AI** (`gemini_extracao.py`): Corrigido o erro `422 Client Error` (Unprocessable Entity) que ocorria ao tentar executar o agente com um arquivo recém-enviado. Agora, definimos o parâmetro `"process": "true"` no multipart/form-data do upload de arquivos e implementamos um loop de polling resiliente para aguardar o status `"completed"` antes de chamar `/execute`.

## [2.15.0] - 2026-07-06

### Adicionado
- **Integração com Tess AI (tess.im)** (`gemini_extracao.py`, `extrair_nf_pdfs.py`, `app.py`, `gui/index.html`): Adicionado suporte completo à API do Tess AI (tess.im) para a extração otimizada de NFS-e (baseada em leitura local de texto) utilizando o modelo **Gemini 3.5 Flash**. A interface gráfica agora conta com inputs para Chave da API Tess e ID do Agente de Chat (Template).
- **Tratamento de Cota para Tess AI** (`gemini_extracao.py`): Implementada a exceção `ErroCotaTessAI` para tratar o erro HTTP 429 (saldo ou limite excedido) e realizar automaticamente o fallback para outros provedores diretos ativos (Gemini ou Groq).
- **Habilitação de Controles Visuais**: Ao ativar a checkbox do Tess AI, a interface automaticamente desabilita o botão de Extração Completa (Multimodal) por incompatibilidade do binário, preservando o fluxo estável.

## [2.14.10] - 2026-07-06

### Adicionado
- **Extração em Lote (Sem IA)** (`app.py`, `gui/index.html`): Adicionada nova opção na interface gráfica para realizar a extração estruturada de PDFs locais. Essa função processa múltiplos arquivos misturados de diferentes prefeituras de uma só vez, gera logs em tempo real na interface, identifica a origem de cada arquivo internamente usando `identificar_nome_prefeitura` e unifica todos os dados em um único Excel (`nf_lote_compilado_sem_ia.xlsx`), sem mover os PDFs originais.

## [2.14.9] - 2026-07-06

### Adicionado
- **Novos Campos do Prestador na Extração com IA** (`gemini_extracao.py`, `extrair_nf_pdfs.py`): Adicionados novos campos do prestador (Nome/Razão Social, UF, Cidade, CEP, Logradouro, Número, Bairro e E-mail) à extração de metadados das notas fiscais por inteligência artificial (Gemini e Groq). As novas colunas no Excel gerado possuem destaque com cabeçalho em vermelho escuro.

### Corrigido
- **Formato Numérico do Número de NF** (`extrair_nf_pdfs.py`): Os valores da coluna `NUMERO_NF` são convertidos para o tipo numérico nativo (`int`) na geração do Excel, evitando avisos de "número armazenado como texto" e facilitando fórmulas.
- **Suporte a CNPJ Alfanumérico** (`extrair_nf_pdfs.py`, `gemini_extracao.py`): Adaptado o filtro de validação anti-alucinação e o método `limpar_cnpj` para aceitar caracteres alfanuméricos (limite de até 14 caracteres de letras e números úteis, desconsiderando pontos, barras, traços e espaços), suportando a nova máscara da Receita Federal.

## [2.14.8] - 2026-07-06

### Melhorado
- **Pausa Instantânea da Automação** (`iss_fortaleza_automacao.py`, `app.py`): Implementada verificação de pausa de baixa latência em operações de clique, digitação e seleção de caixas de listagem do Selenium. Agora, o robô interrompe sua execução no exato instante em que o operador clica no botão "Pausar Automação" da GUI, em vez de esperar a finalização completa do preenchimento da nota fiscal atual, eliminando a sensação de travamento.

## [2.14.7] - 2026-07-06

### Corrigido
- **BUG-01 / BUG-03 — Flag de login com caminho relativo** (`app.py`): As funções `confirmar_login_feito` e `_processar_automacao` criavam o arquivo de flag de controle de login em `Path("brain/funcao2_login_ok.flag")`, um caminho relativo que resolvia incorretamente no instalador (onde o `cwd` pode ser `C:\Windows\System32`). Ambos passaram a usar `%APPDATA%/BeAContab/brain/funcao2_login_ok.flag`, garantindo permissão de escrita em qualquer máquina.
- **BUG-02 — Fallback de `caminho_confirmacao_login` com caminho relativo** (`iss_fortaleza_automacao.py`): O fallback interno `caminho_confirmacao_login or Path("brain/...")` sofria do mesmo problema. Substituído por lógica explícita que resolve para o AppData quando o argumento não é informado.
- **BUG-04 — Perfil persistente do Chrome com caminho relativo** (`iss_fortaleza_automacao.py`): A constante `PERFIL_CHROME_FUNCAO2_PADRAO` apontava para `brain/navegador_funcao2_profile`. Em `executar_fluxo_iss`, adicionada lógica que, quando o valor padrão é detectado, resolve o caminho para `%APPDATA%/BeAContab/brain/navegador_funcao2_profile`, preservando o histórico de login entre sessões no instalador.
- **BUG-05 — Variável `message_to_js` não utilizada** (`app.py`, `_processar_extracao_otimizada`): Resíduo de refatoração anterior. Removida a atribuição dupla `msg_safe = message_to_js = ...`.
- **BUG-06 — Callbacks de log com escape manual inseguro** (`app.py`): Todos os callbacks `gui_log_callback` e `log_gui` das funções `_processar_extracao`, `_processar_extracao_otimizada`, `_processar_automacao`, `_processar_exportar_prefeituras`, `_processar_uniao_planilhas` e `_processar_organizar_pdfs` usavam substituição manual de aspas (`replace("'", ...)`) que falhava silenciosamente ao encontrar barras invertidas (`\n`, `\t`, `\\`). Todos migrados para `json.dumps()`, consistente com os blocos `except`.
- **BUG-08 — `_clicar_com_espera` retornava `None` implicitamente** (`iss_fortaleza_automacao.py`): Após o loop de 3 tentativas, a função encerrava sem `return` ou `raise` caso o `TimeoutException` fosse lançado na última iteração. Adicionado `raise RuntimeError(...)` sentinela após o loop para garantir que falhas sejam sempre expostas ao chamador.
- **BUG-09 — `Any` usado sem `from typing import Any`** (`iss_fortaleza_automacao.py`): O tipo `Any` era utilizado como anotação em `_formatar_celula_para_string_de_valor` e parâmetros `callback_progresso` sem o import correspondente. Adicionado `from typing import Any` à seção de imports.
- **Empacotamento do Selenium no PyInstaller** (`build.py`): Adicionados argumentos `--hidden-import` explícitos para empacotar todos os módulos do Selenium e `webdriver-manager`. Isso contorna a falha de rastreamento estático do PyInstaller (que gerava o erro `No module named 'selenium.webdriver.chrome.webdriver'`) devido a mudanças na importação dinâmica introduzidas na versão experimental do Python 3.14.
- **Nomenclatura do Instalador e Executável** (`build.py`, `setup.iss`): Alinhado o nome de saída do PyInstaller para `automacao-iss-fortaleza`, sincronizando as diretivas de busca de arquivos `dist\` no script do Inno Setup para evitar falhas de "Arquivo não encontrado".
- **Trava de Execução Concorrente** (`app.py`): Adicionadas as flags de estado `self._automacao_em_execucao` e `self._extracao_em_execucao` no backend da GUI para impedir a inicialização acidental de múltiplos robôs ou threads de extração concorrentes em background (corrigindo a race condition que misturava logs e selecionava o mês incorreto).

### Melhorado
- **MELHORIA-01 — Traceback completo exposto no log da GUI** (`app.py`): Os blocos `except` dos três fluxos principais (`_processar_extracao`, `_processar_extracao_otimizada`, `_processar_automacao`) capturavam o traceback em `tb` mas nunca o utilizavam. Adicionado `print(f"[ERRO INTERNO] {tb}")` para enviar o stack trace completo ao log da GUI via `GUIStdoutWrapper`, facilitando o diagnóstico de erros críticos.
- **MELHORIA-02 — Seleção do menu "Escrituração" por texto (XPath)** (`iss_fortaleza_automacao.py`): A seleção do menu dependia do índice posicional `menus[4]`, quebrando caso o portal alterasse a ordem dos menus. Substituído por `WebDriverWait` com XPath que localiza o link pelo texto e classe, mantendo `menus[4]` como fallback de compatibilidade.
- **MELHORIA-03 — Verificação de sucesso de gravação resiliente a capitalização** (`iss_fortaleza_automacao.py`): O `EC.text_to_be_present_in_element` aguardava a string exata `"Documento digitado com Sucesso"` por 15 segundos antes de lançar exceção caso o portal usasse capitalização diferente. Substituído por `EC.presence_of_element_located` com XPath `contains()` combinado (verifica "digitado com" E "Sucesso" ou "sucesso"), sendo resiliente a variações futuras da mensagem do portal.

## [2.14.6] - 2026-07-06

### Corrigido
- **Persistência de Chaves de API no Instalador** (`app.py`): O arquivo de configuração `.env` contendo as chaves de API do Gemini e Groq passou a ser lido e gravado na pasta **AppData** do usuário (`%APPDATA%/BeAContab/.env`). Isso resolve a falha em computadores cliente em que a gravação do `.env` falhava silenciosamente por falta de permissão de escrita no diretório de instalação (`Program Files`), garantindo que as chaves fiquem salvas de forma permanente e sem privilégios de administrador.
- **Carregamento Automático de Chaves**: Adicionado auto-carregamento das chaves de API direto nas variáveis de ambiente do processo (`os.environ`) no construtor da aplicação, evitando que a GUI ou o backend fiquem sem acesso às credenciais nas execuções.

## [2.14.5] - 2026-07-06

### Corrigido
- **Inicialização Resiliente do Chrome** (`iss_fortaleza_automacao.py`): Implementada estratégia híbrida nas funções `abrir_navegador_visivel` e `abrir_navegador_com_perfil_persistente`. O robô tenta primeiro inicializar usando o Selenium Manager nativo (disponível no Selenium 4.6+), que evita problemas com downloads de drivers bloqueados por firewalls e erros de certificados SSL (`SSLError`) comuns em Python no PyInstaller. Caso falhe, realiza o fallback transparente para o `webdriver-manager` convencional.
- **Log de Diagnóstico do Chrome**: Adicionado tratamento de erro robusto que reporta detalhadamente na interface do usuário (GUI) os erros do Selenium Manager e do Webdriver Manager caso o Google Chrome não possa ser inicializado, facilitando a identificação de máquinas cliente sem o Chrome instalado.

## [2.14.4] - 2026-07-06

### Corrigido
- **Seletor CSS Inválido no Calendário RichFaces** (`iss_fortaleza_automacao.py`): Os seletores CSS que utilizavam `#{base_id}Header` e `#{base_id}` falhavam com `InvalidSelectorException` pois o `base_id` contém caracteres `:` do JSF, que o motor CSS interpreta como pseudo-classe. Substituídos pelos seletores de atributo `[id='{base_id}Header']` e `[id='{base_id}']`, que são agnósticos a caracteres especiais no ID.
- **SyntaxError no Log de Erros da GUI** (`app.py`): As mensagens de exceção do Selenium inseridas diretamente em chamadas `evaluate_js` continham aspas simples e quebras de linha, quebrando a string Javascript. Todos os blocos `except` que fazem log de erros na GUI passaram a usar `json.dumps()` para serializar e escapar corretamente as mensagens antes da injeção no webview.
- Adicionado `import json` à seção de imports de `app.py`.

## [2.14.3] - 2026-07-06

### Corrigido
- **Seleção de Mês/Competência na ISS Fortaleza**: Ajustada a função `_abrir_editor_calendario` para clicar no botão popup do calendário (`{base_id}PopupButton`) antes de tentar acessar seu cabeçalho, solucionando a falha de elemento não visível/encontrado que encerrava o navegador.
- **Teste Assistido e Preservação de Navegador**: Modificado o comportamento de encerramento na cláusula `finally` do fluxo de automação para manter a sessão do Chrome aberta quando o robô for executado a partir da interface gráfica (GUI), permitindo a auditoria visual em caso de erros.

## [2.14.2] - 2026-06-30

### Corrigido
- Ajustes na "Exportação por Prefeitura":
  - **Petrolina**: Resolvido o NameError de `nome_prefeitura is not defined` no retorno do parser dedicado.
  - **Barueri**: Corrigido o `IndexError: no such group` adicionando parênteses de captura na expressão regular de CNAE.
  - **Campo Grande / Campina Grande**: Ajustada a extração de cidade/UF de prestação com suporte a caracteres nulos (`\x00`) e wildcards para contornar falhas de fontes nos PDFs do layout Ágili.
  - **Barra de Progresso**: Unificado o progresso para calcular a porcentagem com base no total acumulado de PDFs de todas as prefeituras selecionadas de forma contínua.
  - **Avisos de Log**: Adicionado aviso explícito de início de gravação de XLSX no console de saída `prefeitura_output.log`.
  - **Identificação de Prefeitura**: Adicionada remoção inteligente de acentuação antes de classificar a prefeitura, eliminando falso-negativos em nomes acentuados (como `João Pessoa` e `Goiânia`).

## [2.14.1] - 2026-06-30

### Corrigido
- Ajuste geral de parsers no módulo `extracao_prefeituras.py` (Sem IA):
  - **Alíquotas Zeradas**: Resolvido o bug onde as alíquotas eram gravadas como 0% devido ao símbolo `%` no float parsing em `extrair_nf_pdfs.py`.
  - **João Pessoa**: Corrigida a extração de número da nota (ignorando a competência temporal), CNAE e descrição de serviço (para não confundir com o CNPJ do prestador), e o valor de serviço formatado com pontos (ex: `10.000.00`).
  - **Fortaleza**: Corrigido o ID do CNAE para extrair o código de serviço ABRASF (`12.07`) e adicionada a correção universal que zera impostos/descontos falsos-positivos idênticos ao valor do serviço devido ao desvio de linhas.
  - **Distrito Federal & Goiânia**: Corrigido o número da nota fiscal que aparecia em linhas subsequentes e a identificação inteligente de natureza de operação.
  - **Petrolina**: Implementado o parser dedicado `_parse_petrolina` para ler a estrutura El-Tech (tabelas e cabeçalhos).
  - **Nacional**: Corrigida a extração de número de notas fiscais usando a chave de acesso de 50 dígitos como fallback de alta confiabilidade.
  - **Eusébio**: Corrigidos CNAE, localidade de prestação (tolerando linhas vazias) e detecção direta de natureza da operação.

## [2.14.0] - 2026-06-30

### Adicionado
- Nova funcionalidade "Organizar PDFs por Prefeitura":
  - Criação da função de classificação dinâmica `identificar_nome_prefeitura` que identifica as prefeituras sem IA, baseando-se em assinaturas textuais e expressões regulares para extrair o nome da prefeitura de forma arbitrária.
  - Moverá cada PDF identificado para uma subpasta correspondente criada dinamicamente na própria pasta de origem.
  - Notas cuja prefeitura não for possível identificar permanecem na raiz.
  - Nova aba "Organizar PDFs" criada na GUI como aba padrão inicial.
- Melhorias na funcionalidade "Unir Planilhas":
  - Adicionado suporte a Drag and Drop (arrastar e soltar) de múltiplos arquivos XLSX de forma acumulativa (novos arquivos arrastados são anexados à lista em vez de sobrepor).
  - Adicionada opção de selecionar uma pasta de origem: o robô escaneia recursivamente as subpastas (`**/*.xlsx`) para encontrar todas as planilhas modelo e consolidá-las.
  - Adicionada exclusão automática de linhas duplicadas de dados na fusão (opção toggle na GUI).
- Melhorias no "Exportar por Prefeitura":
  - Adicionado botão "LIMPAR TUDO" na interface para esvaziar todos os caminhos de pastas inseridos no grid.
  - Corrigido bug de retorno na função de exportar que exibia a mensagem de erro "Nenhuma nota extraída com sucesso" mesmo quando a planilha era gerada com sucesso.

## [2.13.0] - 2026-06-30

### Adicionado
- Nova funcionalidade "Exportar por Prefeitura" (sem IA):
  - Criado o módulo `extracao_prefeituras.py` com parsers baseados em regex para 12 prefeituras:
    Governo do Distrito Federal, Petrolina, Nacional, São Roque, Barueri, Campo Grande,
    Eusébio, Fortaleza, Maracanaú, Teixeira de Freitas, Goiânia e João Pessoa.
  - São Paulo (CID garbled) e Maceió (PDF escaneado) são ignorados automaticamente.
  - PDFs processados com sucesso são movidos para a subpasta `concluidas/` dentro da pasta de origem.
  - PDFs com falha ou sem camada de texto permanecem na pasta original.
  - A planilha `nf_compilado.xlsx` é gerada na própria pasta de origem de cada prefeitura.
  - O campo IR da Prefeitura de Fortaleza é analisado dinamicamente por coordenadas e regex, garantindo a extração quando houver imposto retido e descartando incorreções com o ISS Retido (como na NF 132).
- Nova funcionalidade "Unir Planilhas":
  - Criado o módulo `unir_planilhas.py` que consolida múltiplos XLSX em `planilha_unificada.xlsx`.
  - Cabeçalhos, formatação contábil e percentual são preservados na planilha consolidada.
- GUI atualizada com dois novos menus no sidebar:
  - Painel "Exportar por Prefeitura": grid de 12 seletores de pasta (um por prefeitura) com log
    e barra de progresso em tempo real.
  - Painel "Unir Planilhas": seletor de múltiplos arquivos XLSX, seletor de pasta de destino e log.
- Novos métodos expostos na API pywebview:
  - `exportar_por_prefeitura(mapa)`, `unir_planilhas_gui(lista, destino)`,
    `selecionar_multiplos_xlsx()` e `obter_prefeituras_suportadas()`.

## [2.12.0] - 2026-06-30

### Adicionado
- Formatação Contábil e Percentual no Excel:
  - Valores monetários (valor do serviço, deduções, impostos e retenções) agora são gravados como números reais (float) e formatados no padrão de contabilidade nacional (`_("R$"* #,##0.00...`).
  - Alíquotas são gravadas como porcentagens reais e formatadas como `0.00%` no openpyxl.
  - A coloração condicional de valores retidos foi adaptada para trabalhar com valores reais já convertidos.
- Estabilização do Layout da GUI:
  - Consoles de logs da GUI (`log-automacao` e `log-extracao`) ganharam limites rígidos de altura (`max-height`) e overflow-y auto locais, prevenindo que o log empurre o layout e crie barras de rolagem globais na janela principal do programa.
- Gravação de Logs em Tempo Real e Limpeza Automática:
  - Logs de escrituras realizadas com sucesso (`log_escrituradas_sucesso.txt`) e prestadores não cadastrados (`log_nao_cadastrados.txt`) passam a ser atualizados incrementalmente no disco em tempo real à medida que o Selenium executa cada nota.
  - Logs parciais antigos da sessão são removidos no início de cada execução.
  - Log de notas com preenchimento incompleto (`log_extracao.txt`) atualizado incrementalmente na pasta origem e destino.
- Log Exclusivo para Notas de Fortaleza:
  - Criado o arquivo `log_prefeitura_fortaleza.txt` na pasta `log` da planilha de automação, registrando de forma exclusiva e em tempo real as notas fiscais ignoradas por terem sido emitidas pela Prefeitura de Fortaleza.

## [2.11.0] - 2026-06-29

### Adicionado
- Gerenciamento Unificado de Logs:
  - Implementado o suporte a múltiplos diretórios de gravação simultânea de logs.
  - Para extrações, os logs são salvos simultaneamente na pasta `log/` do diretório de origem (dos PDFs) e do diretório de destino (da planilha XLSX).
  - Para automações, os logs são salvos na pasta `log/` do diretório do arquivo XLSX.
  - O logger global do sistema foi aprimorado para evitar subpastas aninhadas redundantes (como `log/log/`).
- Atualizações de Logs em Tempo Real na GUI:
  - Implementado o redirecionamento de `sys.stdout` e `sys.stderr` por meio do `GUIStdoutWrapper` em `app.py`. Todos os prints de console do robô e mensagens de progresso agora são capturados e renderizados em tempo real no console do aplicativo.

## [2.10.0] - 2026-06-29

### Adicionado
- Recursos de GUI e Fluxo de Automação:
  - Adicionado o botão de **"PARAR AUTOMAÇÃO"** / **"CONTINUAR AUTOMAÇÃO"** na interface gráfica para permitir pausar e retomar a escrituração no portal da ISS a qualquer momento após o login manual.
  - Implementado o controle thread-safe no Python por meio do `self._pausa_automacao_event` que suspende temporariamente a execução do robô no início de cada processamento de nota fiscal e retoma de forma suave ao sinal do usuário.

## [2.9.1] - 2026-06-29

### Corrigido
- Recursos de GUI e Fluxo de Automação:
  - Implementado o parâmetro `executando_em_gui` no robô. Quando ativo, o script ignora qualquer prompt de `input()` em console (como o prompt final de confirmação de término), permitindo que o fluxo siga direto após a confirmação do login feito na GUI, sem travar a thread de execução do pywebview.

## [2.9.0] - 2026-06-29

### Adicionado
- Recursos de GUI:
  - Adicionado painel e botão de confirmação manual de login (**"Confirmar Login Realizado"**) na GUI de automação, permitindo destravar e retomar o robô de forma visual sem necessidade de interação no console (corrigido bug de concorrência assíncrona que ocultava o botão imediatamente).
- Correções de Logs e Documentação:
  - Removidas referências residuais do framework Playwright em logs de execução e nos arquivos de documentação (`README.md` e `DEVELOPER.md`), padronizando o projeto no uso do Selenium WebDriver.

## [2.8.0] - 2026-06-27

### Adicionado
- Melhorias de GUI:
  - Renomeada a aplicação de "BeAContab" para "Automação: ISS Fortaleza" na janela da GUI e na barra lateral.
  - Renomeado o item de menu "Automação ISS Fortaleza" para "Automação: Escrituração".
  - Renomeado o título principal do painel de automação para "Automação: Escrituração".
  - Reordenados os botões de extração: o botão "Extração Otimizada" agora vem em primeiro e aparece destacado como "Recomendada" (com preenchimento gradiente principal), enquanto o botão "Extração Completa" vem em segundo (com borda/outline escuro).
  - Adicionado botão de "PARAR EXTRAÇÃO" / "CONTINUAR EXTRAÇÃO" que permite ao usuário pausar e retomar o loop de extração de PDFs em tempo de execução de forma segura via threads.
  - Nova barra de progresso no painel de Automação (ISS Fortaleza) que rastreia visualmente a porcentagem de escrituração das notas com base no callback de progresso das linhas processadas no robô.
  - Ao atingir o limite de cotas de IA, a barra de progresso da extração é marcada automaticamente em 100% e pintada na cor vermelha (`#ff4a4a`).
- Validação Antialucinação e Proteção de Dados:
  - Implementado prompt estrito instruindo as IAs a nunca extraírem CPF, CNPJ ou números de telefone no campo de ID CNAE.
  - Implementada função auxiliar de detecção de padrões de documentos e contagem de dígitos (`_eh_documento_ou_telefone`). O construtor `__post_init__` limpa automaticamente o campo de ID CNAE e ID CNAE Final caso as IAs alucinem CPF, CNPJ ou telefones nele.
- Logs e Auditoria:
  - Criação do log `log_limite_cota.txt` em subpasta `log/` na pasta de origem se a extração for interrompida por cota. O log detalha a ocorrência do rate limit e lista os arquivos de notas fiscais que ficaram pendentes de processamento.

## [2.7.0] - 2026-06-26

### Adicionado
- Melhorias de GUI:
  - Inputs de caminhos de pasta de origem e destino agora permitem digitação manual (remoção do atributo `readonly`).
  - Lógica de auto-cópia: digitar ou selecionar a pasta de origem preenche automaticamente a pasta de destino com o mesmo caminho.
  - Adicionado card explicativo com instruções de como obter a chave da API do Groq ao lado das instruções do Gemini.
  - Seleção dinâmica do mês e ano competência padrão atuais (mês/ano corrente da máquina).
  - Remoção de campo inútil de chave de API na aba de automação (não utilizado pela automação do portal).
  - Desativação do console DevTools automático na inicialização da aplicação (ajustado `webview.start(debug=False)`).
- Regras Fiscais e Planilha Excel:
  - Limpeza e validação de CNAE: o ID CNAE (tanto na extração quanto no CNAE Final oficial) agora exige conter pelo menos um dígito numérico. Caso contrário, é limpo para evitar dados textuais inválidos no campo.
  - Formatação condicional: células de valores fiscais retidos (Deduções, Descontos, Impostos Federais) que forem maiores que zero são destacadas em vermelho claro/pastel (`FFC7CE` e fonte `9C0006`).
  - Validação de Alíquota: células de alíquotas diferentes de `5%` (ou `5,00`) são destacadas na mesma formatação vermelha de alerta.

## [2.6.0] - 2026-06-26

### Adicionado
- Recursos de GUI:
  - Barra de progresso visual em porcentagem animada no painel de monitoramento da extração.
  - Exibição de contagem de PDFs e estimativa de tempo aproximado (Completa e Otimizada) ao selecionar uma pasta.
  - Inputs para chaves de API do Gemini e Groq individuais, com suporte para salvar no arquivo `.env`.
  - Checkboxes para uso individual/simultâneo do Gemini e Groq.
  - Tratamento inteligente de quotas (`ErroCotaGemini` e `ErroCotaGroq`) com aviso destacado em vermelho no console e interrupção imediata, salvando o progresso parcial das notas no arquivo `.xlsx` sem perda de dados.
- Regras Fiscais e Visualização:
  - Novas regras estritas de extração para CNAEs com base na prefeitura emissora (Petrolina, Fortaleza, João Pessoa, Goiânia, Natal, Olinda, Vila Velha, Rio de Janeiro, São Paulo, Teixeira de Freitas, Maracanaú, São Roque).
  - Regra de proximidade ao termo "PRESTADOR" para os campos `uf_local_prestacao` e `cidade_local_prestacao`.
  - Refinamento de prompts no extrator (`gemini_extracao.py`) instruindo as IAs a copiar e colar a discriminação do serviço de forma literal e a extrair valores dedutíveis e impostos (INSS, IR, CSRF, etc.) por proximidade física com os termos correspondentes no texto.
  - Estilização colorida no Excel (`gerar_xlsx`), pintando os cabeçalhos das 21 colunas fiscais solicitadas em Vermelho Escuro (`C00000`) e as 5 colunas auxiliares em Azul (`1F4E78`).
- Otimizações:
  - Implementação de colagem instantânea de textos longos via JavaScript no navegador (`_definir_valor_instantaneo`) para o campo de descrição do serviço em `iss_fortaleza_automacao.py`, eliminando a digitação caractere por caractere e acelerando expressivamente a automação.

## [2.5.1] - 2026-06-26

### Adicionado
- Arquivos:
  - `.env`
  - `.env.example`
  - `gemini_extracao.py`
- Motivo: documentar e adicionar placeholders para as configurações da API do Groq (`GROQ_API_KEY` e `GROQ_MODEL`) e parametrizar o modelo no código para maior flexibilidade.
- Impacto:
  - Adicionados os placeholders `GROQ_API_KEY=` e `GROQ_MODEL=llama-3.3-70b-versatile` nas configurações de ambiente.
  - Modificado o script `gemini_extracao.py` para utilizar a variável `GROQ_MODEL` dinamicamente com fallback seguro para `llama-3.3-70b-versatile`.

## [2.5.0] - 2026-06-26

### Adicionado
- Arquivos:
  - `gemini_extracao.py`
  - `extrair_nf_pdfs.py`
  - `iss_fortaleza_automacao.py`
  - `app.py`
- Motivo: implementar fallback para extração via Groq, processamento em lote da automação da Função 2 e geração de logs estruturados em subpastas.
- Impacto:
  - Implementada a função `extrair_campos_groq` em `gemini_extracao.py` como fallback único caso o Gemini falhe, com suporte a JSON mode. O loop de múltiplas tentativas (retry) do Gemini foi removido (agora tenta apenas 1x).
  - A automação no portal ISS (`iss_fortaleza_automacao.py`) agora processa todos os documentos do XLSX de forma contínua em lote, coletando erros e acertos em listas separadas.
  - Inclusão de logs centralizados na subpasta `log/` (`log_escrituradas_sucesso.txt` e `log_nao_cadastrados.txt`) detalhando CNPJ e nota, e do arquivo `log_nao_analisados.txt` gerado durante a etapa de extração caso ambas as IAs falhem.
  - As colunas `ANALISADO_PELA_IA` e `REVISAO_MANUAL` foram removidas do arquivo final XLSX.
  - Adicionada a nova coluna `MODELO_IA` no XLSX, indicando qual inteligência artificial (Gemini ou Groq) foi responsável por extrair os dados da linha correspondente.

## [2.4.0] - 2026-06-26

### Adicionado
- Arquivos:
  - `gemini_extracao.py`
  - `extrair_nf_pdfs.py`
  - `app.py`
  - `gui/index.html`
- Motivo: criação da Função 3 (Extração Otimizada por Texto) para permitir a extração fiscal consumindo muito menos tokens/quota do Gemini, mantendo a Função 1 multimodal original intacta.
- Impacto:
  - Criada a função `extrair_campos_gemini_somente_texto` em `gemini_extracao.py`, que envia apenas o texto do PDF no prompt, sem carregar ou anexar o arquivo binário do PDF.
  - Implementada a função `extrair_nota_fiscal_somente_texto` e suporte ao modo 3 via CLI (argparse e menu interativo) em `extrair_nf_pdfs.py`.
  - Integrados os métodos `iniciar_extracao_otimizada` e `_processar_extracao_otimizada` em `app.py` para processamento em thread dedicada.
  - Adicionado botão físico dedicado de "Extração Otimizada (Somente Texto)" e lógica Javascript integrada em `gui/index.html`.

## [2.3.0] - 2026-06-25

### Alterado
- Arquivos:
  - `iss_fortaleza_automacao.py`
- Motivo: automatizar o fluxo contínuo de digitação na automação da ISS, verificando a gravação de sucesso e avançando para o próximo formulário de forma autônoma.
- Impacto:
  - Implementada a verificação de sucesso monitorando o título `//*[@id="content"]/legend/h2` para aguardar o texto "Documento digitado com Sucesso".
  - O robô agora clica automaticamente no botão "Digitar novo documento" usando o XPATH `//*[@id="j_id165:novo"]` para limpar o formulário e deixá-lo pronto para a próxima nota fiscal.

## [2.2.0] - 2026-06-25

### Adicionado
- Arquivos:
  - `iss_fortaleza_automacao.py`
- Motivo: automatizar o processo de gravação (salvamento) do documento fiscal digitado no portal ISS Fortaleza.
- Impacto:
  - O robô agora clica no botão "Gravar Documento" usando o ID `digitarDocumentoForm:j_id477` (ou seletor XPATH robusto) de forma automatizada ao final do preenchimento dos dados do serviço e retenções na aba Serviço.
  - Adicionada espera de processamento de feedback do portal e registro em log do encerramento da gravação automática.

## [2.1.0] - 2026-06-25

### Alterado
- Arquivos:
  - `iss_fortaleza_automacao.py`
- Motivo: suporte ao preenchimento de campos de retenções federais (IR, INSS, CSRF, PIS Não Retido, COFINS Não Retido, Outras Retenções), deduções e descontos (Condicionados e Incondicionados) na aba Serviço da automação da ISS Fortaleza.
- Impacto:
  - Atualizado o dataclass `DocumentoPortalISS` para comportar as novas propriedades financeiras.
  - Ajustado o método `carregar_documentos_xlsx` para mapear e extrair os valores das novas colunas, mantendo retrocompatibilidade com planilhas que não possuem estas colunas (usando "0,00" por padrão).
  - Implementada a função utilitária `_digitar_campo_se_nao_zero` para garantir que apenas valores diferentes de zero sejam digitados (evitando digitar "0,00" nos campos em branco do portal).
  - Atualizado `preencher_documento_servico` mapeando os IDs do DOM identificados para os 10 campos financeiros do portal de escrituração.

## [2.0.0] - 2026-06-25

### Adicionado
- Arquivos:
  - `app.py`
  - `gui/index.html`
- Motivo: criação de um App Desktop Nativo (GUI) para substituir a interação exclusivamente por linha de comando, atendendo à necessidade do usuário de uma interface instalável (futuro `.exe`).
- Impacto:
  - Implementado o wrapper principal com `pywebview` no arquivo `app.py`.
  - Criada a interface gráfica em `gui/index.html` (SPA) englobando as funcionalidades de Extração de NF (Aba 1) e Automação ISS Fortaleza (Aba 2), com design responsivo, tokens Lumina Automation e suporte nativo ao TailwindCSS via CDN.
  - O aplicativo transmite os logs de processamento em tempo real diretamente para o painel de console da interface HTML através da injeção de dependência via Javascript (`window.pywebview.api`).
  - Incluído menu de instruções de como obter a API Key na interface da primeira Aba.

### Alterado
- Arquivos:
  - `gemini_extracao.py`
  - `extrair_nf_pdfs.py`
  - `requirements.txt`
- Motivo: suporte à injeção dinâmica de chave da API pela interface gráfica.
- Impacto:
  - O parâmetro `api_key` agora pode ser passado ativamente da UI, contornando a exigência de usar exclusivamente arquivos `.env`.
  - Adicionada a dependência `pywebview>=4.4.1` ao `requirements.txt`.
## [1.9.1] - 2026-06-25

### Adicionado
- Arquivos:
  - `gemini_extracao.py`
- Motivo: prevenir que notas fiscais válidas sejam ignoradas em processamentos em lote devido a estouro temporário de limites de cota da API (RPM/TPM).
- Impacto:
  - Implementado mecanismo de retentativa automática (retry) em caso de exceção `RESOURCE_EXHAUSTED` (HTTP 429) e `QUOTA`.
  - O script agora aguarda temporariamente (12 segundos) antes de retentar o arquivo que falhou, com limite de até 4 tentativas, o que aumenta a resiliência a picos rápidos de chamadas e impede o salto (skip) desnecessário de notas.

## [1.9.0] - 2026-06-25

### Alterado
- Arquivos:
  - `extrair_nf_pdfs.py`
  - `gemini_extracao.py`
- Motivo: consolidar a extração estritamente baseada em Inteligência Artificial (antiga Função 3) como a nova e única **Função 1** (Apenas IA), e remover o código obsoleto baseado em expressões regulares locais.
- Impacto:
  - O modo de extração híbrido (Regex + IA) foi completamente removido, e a nova Função 1 opera de forma integral e direta com o Gemini.
  - Implementada validação dinâmica com limite de 3 meses para a data de emissão para impedir que o Gemini gere dados (como o ano de 2020) muito distantes da execução atual.
  - Adicionado filtro anti-alucinação comparando os dígitos retornados pela IA (CNPJ, número da nota) com o texto bruto extraído; se ausentes no texto, os campos são esvaziados.
  - Inserida instrução imperativa para extração do CNAE nas notas da Prefeitura Municipal de Petrolina estritamente sob o rótulo "SERVIÇO NACIONAL", corrigindo CNAEs truncados e capturas ambíguas.
  - Renomeada a coluna `CSRF` no arquivo de exportação XLSX para `CSRF (CSLL + PIS + Cofins Retidos)`.
  - O menu principal e a CLI foram ajustados para exibir exclusivamente os modos "1" (Extração Apenas IA) e "2" (Automação).
  - Removidas mais de 20 funções obsoletas em `extrair_nf_pdfs.py` que tratavam a extração regex localizada de layouts das Prefeituras.

## [1.8.1] - 2026-06-25

### Adicionado
- Arquivos:
  - `gemini_extracao.py`
  - `extrair_nf_pdfs.py`
- Motivo: capturar falhas de limite de cota da API do Gemini (`RESOURCE_EXHAUSTED` / HTTP 429) e gerar um arquivo de log `{planilha}_log_cota_excedida.txt` na mesma pasta do XLSX, evitando o descarte silencioso sem notificação do usuário.
- Impacto:
  - Criada a exceção customizada `ErroCotaGemini` em `gemini_extracao.py` para propagar falhas de cota.
  - Adicionado tratamento no loop da função `main` em `extrair_nf_pdfs.py` para capturar `ErroCotaGemini`, registrando o arquivo com falha e marcando a nota com observação de cota excedida.
  - Implementada escrita condicional ao final do script para gerar o arquivo de log contendo a lista de PDFs que falharam devido a limites de cota.

## [1.8.0] - 2026-06-25

### Alterado
- Arquivos:
  - `cnae_final.py`
  - `gemini_extracao.py`
  - `.env`
  - `.env.example`
- Motivo: unificar o ecossistema de IA do projeto, removendo completamente o Groq e utilizando a API do Google Gemini para todas as tarefas de processamento inteligente (tanto extração fiscal quanto classificação oficial de CNAE).
- Impacto:
  - A API da Groq foi removida de `cnae_final.py` e substituída pela chamada correspondente `_chamar_gemini` utilizando o SDK `google-genai`.
  - Definido o pipeline único de modelos do Gemini: `gemini-2.5-flash` (Principal), `gemini-2.5-pro` (Fallback 1) e `gemini-2.5-flash-lite` (Fallback 2) em todo o sistema.
  - Removidas as chaves e referências de ambiente do Groq no `.env` e `.env.example`.

## [1.7.0] - 2026-06-25

### Adicionado
- Arquivos:
  - `extrair_nf_pdfs.py`
  - `gemini_extracao.py`
- Motivo: requisição para criar a Função 3, que realiza o mesmo processamento que a Função 1 (extrair dados de notas fiscais em PDF e compilar em XLSX), porém utilizando exclusivamente Inteligência Artificial (Gemini) para a extração do zero de todos os campos fiscais, sem depender do parser de regex local.
- Impacto:
  - Criado o Modo 3 ("Apenas IA") acionável interativamente ou por linha de comando (`-m 3`).
  - Adicionado suporte na função `extrair_nota_fiscal` para isolar e pular a extração regex local, chamando a IA para todos os campos.
  - Implementada a flag `somente_ia` em `extrair_campos_gemini` para bypassar verificação condicional de fallback e otimizar o prompt de extração única.
  - Na planilha consolidada por essa opção, a coluna `REVISAO_MANUAL` fica vazia por padrão, pois não há comparação com dados de regex local.

## [1.6.0] - 2026-06-25

### Adicionado
- Arquivos:
  - `extrair_nf_pdfs.py`
  - `gemini_extracao.py`
  - `iss_fortaleza_automacao.py`
- Motivo: requisição para extrair campos tributários adicionais (deduções, descontos, retenções federais) nativamente por regex antes de usar IA, além da imunidade a notas emitidas pela própria Prefeitura de Fortaleza.
- Impacto:
  - Adicionada extração (local e IA) e exportação no XLSX das colunas: `PREFEITURA`, `VALOR_DEDUCOES`, `DESCONTOS_INCONDICIONADOS`, `DESCONTOS_CONDICIONADOS`, `OUTRAS_RETENCOES`, `IR`, `PIS_NAO_RETIDO`, `COFINS_NAO_RETIDO`, `CSRF` e `INSS`.
  - Coluna `REVISAO_MANUAL` adicionada para indicar os campos monetários que foram ajustados/sobrescritos pela Inteligência Artificial e exigem conferência humana.
  - A automação da função 2 (ISS Fortaleza) ignora silenciosamente notas em que a Prefeitura de Fortaleza conste como prefeitura emissora, gravando o pulo no arquivo de log.

## [1.5.0] - 2026-06-24

### Adicionado
- Arquivos:
  - `extrair_nf_pdfs.py`
- Motivo: necessidade de informar se a nota fiscal foi enriquecida por IA e simplificar a localização do arquivo exportado.
- Impacto:
  - Adicionada a coluna `ANALISADO_PELA_IA` na planilha XLSX gerada, indicando "SIM" se o registro foi retornado/refinado pelo Gemini, ou "NÃO" caso contrário.
  - O arquivo XLSX de destino, caso seja um caminho relativo, será salvo na mesma pasta informada para os PDFs de origem, não mais no diretório de execução raiz.

## [1.4.4] - 2026-06-24

### Alterado
- Arquivos:
  - `iss_fortaleza_automacao.py`
- Motivo: erro de correspondência case-sensitive/accent-sensitive ao tentar casar a cidade da planilha ("Fortaleza") com a do portal ("FORTALEZA").
- Impacto:
  - Implementada normalização de strings (remoção de acentos e conversão para minúsculas) em `_selecionar_opcao_por_texto`.
  - O script agora realiza correspondência exata normalizada e depois correspondência parcial normalizada (ex: "Fortaleza" correspondendo a "FORTALEZA - CE"), ignorando placeholders.

## [1.4.3] - 2026-06-24

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
- Motivo: erro de `NoSuchElementException` ao tentar selecionar a cidade após selecionar o estado (AJAX tardio).
- Impacto:
  - Ajustada a função `_selecionar_opcao_por_texto` para capturar `NoSuchElementException` e disparar a retentativa automática (até 4 tentativas de 0.5s).
  - Garante que a automação aguarde o preenchimento dinâmico de opções pelo portal sob chamadas assíncronas do JSF.

## [1.4.2] - 2026-06-24

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
- Motivo: erro de `StaleElementReferenceException` (elemento obsoleto) ao preencher o formulário de serviço após carregar o tipo de documento.
- Impacto:
  - Implementada lógica de retentativas automáticas (3 tentativas com intervalo de 0.5s) nas funções `_clicar_com_espera`, `_digitar_campo` e `_selecionar_opcao_por_texto`.
  - Isso garante resiliência a recarregamentos AJAX/JSF parciais que redesenham o formulário em tempo de execução.

## [1.4.1] - 2026-06-24

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
- Motivo: falha silenciosa na seleção de prestadores pelo autocomplete do CNPJ.
- Impacto:
  - Corrigido o seletor CSS composto que continha dois-pontos (`:`), o que gerava um erro de sintaxe de pseudo-classe do CSS no Selenium.
  - A busca agora utiliza o ID do container e localiza as sugestões de forma relativa e segura.

## [1.4.0] - 2026-06-24

### Alterado
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `requirements.txt`
- Motivo: migração completa do framework de automação de navegador de **Playwright** para **Selenium WebDriver** com gerenciamento automático do ChromeDriver via `webdriver-manager`.
- Impacto:
  - Todo o código de automação deixou de ser assíncrono (`asyncio`/`async`/`await`) e passou a ser estritamente síncrono, simplificando o fluxo e a depuração direta no terminal.
  - Substituição de `playwright>=1.54.0` por `selenium>=4.21.0` e `webdriver-manager>=4.0.0` no `requirements.txt`.
  - A abertura do navegador Chrome, o clique em elementos, a digitação em campos e a seleção de options agora utilizam as APIs do Selenium (`WebDriverWait`, `expected_conditions`, `Select`, `ActionChains`).
  - Todas as funcionalidades existentes foram preservadas: seleção de competência no calendário RichFaces, autocomplete de CNPJ, modal de pesquisa de CNAE (incluindo fallback via `driver.execute_script`), preenchimento da aba Serviço e suporte ao modo de perfil persistente.
  - O parâmetro `usar_inspector` foi mantido na assinatura de `executar_automacao_iss` apenas para compatibilidade com a CLI; a funcionalidade de Inspector do Playwright não tem equivalente no Selenium.
  - O parâmetro `porta_debug_navegador` também foi mantido na assinatura por compatibilidade com a CLI, mas não é mais utilizado no fluxo interno do Selenium.


## [1.3.2] - 2026-06-23

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
- Motivo: inserção de pausa de estabilização de 1.2 segundos após a seleção do dropdown de Status e antes de abrir o modal do CNAE.
- Impacto: evita o erro de timeout causado pelo processamento assíncrono (AJAX/JSF) que fechava ou impedia a exibição do modal de pesquisa de CNAE.

## [1.3.1] - 2026-06-23

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
- Motivo: injeção de script stealth para contornar o bloqueio de robôs (bypassar `navigator.webdriver` e flag de automação) e desativação do interceptador de downloads do Playwright para preservar os nomes originais dos arquivos baixados.
- Impacto: permite que a busca de CNPJ funcione e corrige a nomenclatura de downloads de PDFs (removendo UUIDs).

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `brain/2026-06-11-aba-servico-id-real.md`
- Motivo: a aba `Serviço` passou a ser acionada explicitamente pelo ID real `digitarDocumentoForm:abaServico_lbl` antes do preenchimento dos campos.
- Impacto: elimina a dependência de texto genérico e garante que o formulário correto seja ativado antes de tentar selecionar tipo de documento ou status.

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `brain/2026-06-11-servico-sem-clique-fragil.md`
- Motivo: o fluxo deixou de depender do clique genérico em `Serviço` e passou a aguardar explicitamente o formulário da aba antes de preencher os campos.
- Impacto: elimina uma interação frágil com texto genérico e reduz falhas na etapa de preenchimento da nota.

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `brain/2026-06-11-transicao-escrituracao-fiscal.md`
- Motivo: o fluxo passou a aguardar explicitamente a tela `Escrituração Fiscal` depois do clique em `Escriturar`, antes de tentar a aba `Serviços Tomados`.
- Impacto: reduz falhas de timing na transição entre a consulta da competência e a abertura da rotina fiscal.

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `brain/2026-06-11-espera-explicita-total.md`
- Motivo: a função 2 passou a exigir espera explícita antes de cliques, digitações e seleções, além de tratar a aba `Serviços Tomados` com um seletor específico do RichFaces.
- Impacto: reduz cliques prematuros em elementos ainda não renderizados e aumenta a chance de sucesso na transição entre `Escriturar` e `Serviços Tomados`.

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `brain/2026-06-11-prestador-nao-cadastrado.md`
- Motivo: quando o CNPJ é digitado na tela de `Digitar Documento` e o portal não devolve um prestador selecionável, o caso passou a ser tratado como `CNPJ não cadastrado`.
- Impacto: a função 2 registra o CNPJ no log, pula para a próxima linha da planilha e evita encerrar o processo com erro genérico de seleção de prestador.

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `brain/2026-06-11-competencia-sem-escriturar-retry.md`
- Motivo: quando a competência consultada mostra o botão `Escriturar` desabilitado, a automação agora solicita outra competência e repete o fluxo a partir dos campos `De` e `Até`.
- Impacto: a função 2 não encerra mais por timeout nesse cenário; ela orienta o usuário a escolher outro mês e retoma a operação sem reiniciar o navegador.

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `brain/2026-06-11-digitar-documento-cnpj.md`
- Motivo: a etapa de `Digitar Documento` passou a aguardar explicitamente a URL e os blocos visíveis da tela antes de procurar o rádio de `CNPJ`.
- Impacto: a função 2 deixa de procurar o `CNPJ` cedo demais e passa a respeitar o carregamento real da tela `Digitar Documento - Serviço Tomado`.

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `brain/2026-06-11-competencia-richfaces.md`
- Motivo: a seleção de competência da função 2 deixou de procurar `selects` inexistentes e passou a usar o editor real dos calendários `De` e `Até` da tela `Manter Escrituração`.
- Impacto: a automação agora consegue escolher dinamicamente o mês e o ano informados pelo usuário, como `02/2027`, `05/2026` ou `08/2027`, nos dois campos do portal.

## [1.3.0] - 2026-06-11

### Adicionado
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `extrair_nf_pdfs.py`
  - `README.md`
  - `brain/2026-06-11-navegador-persistente-funcao2.md`
- Motivo: criação do modo `--reutilizar-navegador-funcao2` para abrir ou reconectar a função 2 em um Chrome persistente.
- Impacto: o login e a janela do portal podem ser reaproveitados entre execuções, inclusive após erro, facilitando testes reais e continuidade operacional.

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: o retorno `Nenhuma Razão Social Encontrada` passou a ser tratado como CNPJ não cadastrado, com registro em log `.txt` e avanço automático para a próxima linha.
- Impacto: a função 2 deixa de tentar selecionar um retorno não selecionável e continua o lote sem interromper a automação.

## [1.3.0] - 2026-06-11

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: criação de espera explícita centralizada para garantir que elementos estejam visíveis e clicáveis antes de qualquer clique na função 2.
- Impacto: a navegação ficou mais sincronizada com o carregamento real do portal e reduz cliques prematuros em telas que aparecem por etapas.

## [1.3.0] - 2026-06-11

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: o clique na aba `Serviços Tomados` passou a procurar o elemento também em frames e com variações de texto, além de rolar a aba para a tela antes do clique.
- Impacto: a navegação para a etapa de `Digitar Documento` ficou mais resiliente em telas com renderização dinâmica ou variação de acentuação.

## [1.3.0] - 2026-06-11

### Alterado
- Arquivos: `iss_fortaleza_automacao.py`, `extrair_nf_pdfs.py`, `README.md`
- Motivo: criação do modo de teste guiado com Inspector do Playwright para executar a função 2 com o mesmo passo a passo da automação real.
- Impacto: os testes passaram a pausar em pontos críticos do fluxo real, facilitando a reprodução de erros equivalentes aos da operação normal.

## [1.2.1] - 2026-06-10

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: inserção de espera após alternar a pesquisa para `CNPJ`, para respeitar o reload parcial do portal antes da digitação.
- Impacto: o CNPJ deixa de ser apagado pelo redraw da tela e a etapa de autocomplete fica mais confiável.

## [1.2.1] - 2026-06-10

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: a digitação passou a simular mais fielmente a entrada humana, com pausas entre teclas e validação por dígitos normalizados.
- Impacto: o CNPJ e outros campos mascarados tendem a permanecer visíveis na tela e a reagir melhor ao autocomplete do portal.

## [1.2.1] - 2026-06-10

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: a digitação dos campos do portal passou a usar `fill()` com verificação de estabilidade para evitar o efeito de texto que aparece e desaparece rapidamente.
- Impacto: o CNPJ e demais campos mascarados ficam mais consistentes em tela antes da seleção do autocomplete e dos próximos cliques.

## [1.2.1] - 2026-06-10

### Alterado
- Arquivos: `iss_fortaleza_automacao.py`, `extrair_nf_pdfs.py`, `README.md`
- Motivo: inclusão do modo `--debug-funcao2` para abrir o navegador em teste visível e pausar o fluxo nos pontos críticos da função 2.
- Impacto: a automação ficou mais fácil de depurar em tempo real e o seletor ambíguo do autocomplete do prestador passou a usar um contêiner mais específico.

## [1.2.1] - 2026-06-10

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: a função 1 passou a oferecer ao usuário a escolha entre iniciar a função 2 ou encerrar o programa logo após a geração do XLSX.
- Impacto: o fluxo ficou mais contínuo e o usuário pode seguir diretamente para a automação da ISS sem reiniciar o sistema.

## [1.2.0] - 2026-06-09

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: reconhecimento de competência já visível na tela de `Manter Escrituração`.
- Impacto: quando a página já estiver em `Junho 2026`, a função 2 segue adiante sem tentar reprocessar os selects de competência.

## [1.2.0] - 2026-06-09

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: reforço das esperas e detecção do estado já selecionado na tela de `Manter Escrituração`.
- Impacto: a função 2 deixa de travar tentando ler a competência cedo demais e passa a reconhecer quando Junho/2026 já está carregado.

## [1.2.0] - 2026-06-09

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: substituição do preenchimento direto do campo `Código CNAE` por digitação real no modal `Pesquisar CNAE`.
- Impacto: todos os valores alimentados pela planilha na função 2 passam a ser digitados pelo teclado, sem colagem ou injeção direta no DOM.

## [1.2.0] - 2026-06-09

### Alterado
- Arquivos: `extrair_nf_pdfs.py`, `README.md`
- Motivo: a automação da ISS passou a perguntar interativamente qual planilha XLSX será usada quando o usuário não informar o caminho explicitamente.
- Impacto: o fluxo fica mais claro para o operador e evita uso silencioso de uma planilha padrão sem confirmação.

## [1.2.0] - 2026-06-09

### Adicionado
- Arquivos: `tratamento_erros.py`, `log_execucao.txt` (runtime)
- Motivo: registro centralizado dos marcos da execução do programa.
- Impacto: o fluxo agora gera um log de auditoria com ações importantes, como abertura de telas, seleção de prestador e conclusão de etapas.

### Alterado
- Arquivos: `main.py`, `extrair_nf_pdfs.py`, `iss_fortaleza_automacao.py`
- Motivo: emissão de eventos de execução nos pontos principais do programa.
- Impacto: a sequência operacional fica rastreável em `log_execucao.txt` sem depender apenas do terminal.

### Alterado
- Arquivo: `.gitignore`
- Motivo: evitar versionamento do `log_execucao.txt`.
- Impacto: o log permanece local e regenerável.

## [1.2.0] - 2026-06-09

### Adicionado
- Arquivos: `tratamento_erros.py`, `log_erro.txt` (runtime)
- Motivo: centralização do registro de erros inesperados em um arquivo único.
- Impacto: falhas da execução passam a gerar log com data, contexto, tipo de erro e stack trace.

### Alterado
- Arquivos: `main.py`, `extrair_nf_pdfs.py`
- Motivo: captura dos erros de entrada do programa e gravação automática no `log_erro.txt`.
- Impacto: o ponto de entrada passa a registrar exceções antes de encerrar a execução.

### Alterado
- Arquivo: `.gitignore`
- Motivo: evitar versionamento do arquivo de log de erros.
- Impacto: `log_erro.txt` permanece como artefato local de diagnóstico.

## [1.2.0] - 2026-06-09

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: inclusão da seleção da linha retornada pelo modal `Pesquisar CNAE` após a busca do código final.
- Impacto: a automação passa a concluir a etapa de CNAE selecionando o item `digitarDocumentoForm:idFormularioPesquisaCnae:idDatatableListaCnae:0:j_id453` antes de seguir o fluxo.

## [1.2.0] - 2026-06-09

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: inclusão do preenchimento do modal `Pesquisar CNAE` com o valor de `ID_CNAE_FINAL` da planilha antes da continuação do fluxo.
- Impacto: a função 2 passa a consultar o CNAE final do documento de forma automática, em vez de depender de preenchimento manual nessa etapa.

## [1.2.0] - 2026-06-09

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: tratamento genérico para o caso em que o autocomplete do portal retorna "Nenhuma Razão Social Encontrada" ao pesquisar o CNPJ do prestador.
- Impacto: a função 2 passa a registrar o CNPJ ignorado em log, avançar para a próxima linha válida da planilha e continuar o fluxo sem interromper a automação.

## [1.2.0] - 2026-06-09

### Refatorado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: remoção dos blocos legados de correção por artista, prestador e nome de arquivo.
- Impacto: a extração passa a depender de regras genéricas por layout, rótulos do documento, conteúdo fiscal e fallback de IA.

### Corrigido
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: suporte genérico para layouts em que o número da NFS-e e o CNPJ aparecem antes do rótulo, além de corte mais preciso da descrição de CNAE.
- Impacto: corrige leituras como número de nota vindo de parcela/data e evita que campos financeiros entrem na descrição do CNAE.

### Corrigido
- Arquivo: `cnae_final.py`
- Motivo: substitui??o da regra fixa por uma correspond?ncia gen?rica baseada em descri??o oficial de CNAE.
- Impacto: qualquer NFS-e com `DESC_CNAE` equivalente a organiza??o de feiras, congressos e exposi??es passa a ser classificada no CNAE final correto sem depender de nota espec?fica.

### Alterado
- Arquivo: `cnae_final.py`
- Motivo: ajuste do enriquecimento para priorizar a descri??o vis?vel do documento antes do classificador local e da Groq.
- Impacto: reduz erros quando o `ID_CNAE` extra?do do PDF conflita com a descri??o do servi?o.


### Corrigido
- Arquivos: `extrair_nf_pdfs.py`, `gemini_extracao.py`, `cnae_final.py`
- Motivo: corre??o das regras de `ISS_RETIDO` para layouts de S?o Paulo, refinamento do n?mero da NF 93 em Macei? via imagem e ajuste determin?stico do CNAE final.
- Impacto: as notas `114`, `115`, `116`, `22`, `23` e `93` passam a sair com os campos fiscais corretos no XLSX, inclusive com acentua??o normalizada em `NATUREZA_OPERACAO`.

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: padroniza??o da `NATUREZA_OPERACAO` com acentua??o correta.
- Impacto: o XLSX deixa de registrar textos corrompidos como `Tributa??o Fora do Munic?pio`.

### Alterado
- Arquivo: `gemini_extracao.py`
- Motivo: inclus?o de uma segunda leitura por imagem para PDFs escaneados com texto local degradado.
- Impacto: melhora a precis?o de `NUMERO_NF` e `ISS_RETIDO` em layouts que o PDF bruto n?o representa bem.

### Alterado
- Arquivo: `cnae_final.py`
- Motivo: regra fixa para o CNAE `SERVI?OS DE ORGANIZA??O DE FEIRAS, CONGRESSOS E EXPOSI??ES` e for?a de `ISS_RETIDO` para os CNAEs finais `932989910` e `900190201`.
- Impacto: evita que o enriquecimento final dependa apenas de IA em casos j? validados manualmente.


### Alterado
- Arquivos: `gemini_extracao.py`, `extrair_nf_pdfs.py`
- Motivo: inclusão de fallback automático para `gemini-2.5-flash-lite` quando o modelo principal atingir quota ou falhar.
- Impacto: notas escaneadas como `NF 114`, `NF 115` e `NF 116` passam a ser extraídas sem depender de troca manual de modelo.

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: ajuste do layout de São Paulo para usar os rótulos fixos da NFS-e e evitar captura errada do município.
- Impacto: o parser passa a respeitar o padrão de `CNPJ_PRESTADOR`, `NUMERO_NF`, `DATA_EMISSAO`, `ID_CNAE`, `DESC_CNAE`, `DESCRICAO_SERVICO`, `CIDADE_LOCAL_PRESTACAO`, `UF_LOCAL_PRESTACAO`, `VALOR_SERVICO` e `ALIQUOTA`.

### Adicionado
- Arquivo: `tests/test_layout_extracao.py`
- Motivo: cobertura de regressão para o layout de São Paulo.
- Impacto: reduz o risco de voltar a aceitar falso positivo de município em futuras alterações.

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: exibição do nome do PDF em processamento para dar visibilidade ao usuário durante a extração.
- Impacto: o usuário passa a acompanhar em qual arquivo o programa está trabalhando e consegue identificar falhas por documento.

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: adaptação da extração para regras por layout de NFS-e, em vez de correções por artista.
- Impacto: modelos de Fortaleza, Petrolina e São Paulo passam a preencher número, data, local, valor e alíquota de forma reaproveitável.

### Adicionado
- Arquivo: `tests/test_layout_extracao.py`
- Motivo: criação de testes de regressão para layouts de Fortaleza, Petrolina e São Paulo.
- Impacto: reduz risco de voltar a depender de ajustes manuais por artista nos PDFs com texto legível.

### Adicionado
- Arquivos:
  - `gemini_extracao.py`
  - `tests/test_gemini_extracao.py`

- Motivo:
  inclusão do Gemini como fallback de extração para PDFs com texto ruim, campos ausentes ou valores suspeitos.

- Impacto:
  melhora a cobertura dos layouts problemáticos sem abandonar a extração local como primeira camada.

### Alterado
- Arquivos:
  - `extrair_nf_pdfs.py`
  - `requirements.txt`
  - `.env.example`
  - `README.md`

- Motivo:
  integração do fallback de IA, inclusão da dependência oficial `google-genai` e documentação da nova configuração.

- Impacto:
  o pipeline passa a complementar a extração local com Gemini quando necessário e mantém a execução funcionando sem a chave da API.

## [1.1.0] - 2026-06-08

### Adicionado
- Arquivos:
  - `cnae_final.py`
  - `.env.example`

- Motivo:
  criação de um enriquecimento automático para preencher `ID_CNAE_FINAL` e `DESC_CNAE_FINAL` no XLSX da função 1 usando o arquivo oficial de CNAE e a API da Groq.

- Impacto:
  o XLSX final passa a conter uma classificação oficial sugerida para cada nota fiscal, com fallback local caso a API não esteja disponível.

### Alterado
- Arquivos:
  - `extrair_nf_pdfs.py`
  - `.gitignore`

- Motivo:
  integração do enriquecimento de CNAE no fluxo de geração do XLSX e bloqueio do versionamento de `.env`.

- Impacto:
  a opção 1 passa a exportar as novas colunas sem alterar o fluxo principal de extração.

### Documentação
- Arquivos:
  - `README.md`

- Motivo:
  explicação do uso da Groq, do arquivo `.env` e do fallback local.

- Impacto:
  melhora a rastreabilidade e reduz dúvidas na execução da função 1.

## [1.1.0] - 2026-06-08

### Adicionado
- Arquivos:
  - `extrair_iss_uma_vez.py`
  - `iss_fortaleza_automacao.py`

- Motivo:
  criação de um utilitário avulso para extrair a tela logada do portal da ISS uma única vez, sem incorporar essa rotina ao menu principal da aplicação.

- Impacto:
  o usuário pode abrir o portal, navegar até a tela desejada e exportar o conteúdo visível para `iss_extracao.xlsx` quando precisar de uma coleta pontual.

### Alterado
- Arquivo:
  - `.gitignore`

- Motivo:
  inclusão do artefato gerado `iss_extracao.xlsx` no bloco de arquivos ignorados.

- Impacto:
  evita versionamento acidental do arquivo de saída criado pela extração avulsa.

### Documentação
- Arquivo:
  - `README.md`

- Motivo:
  inclusão de instruções para uso da extração avulsa fora do fluxo principal.

- Impacto:
  melhora a rastreabilidade do uso pontual da coleta do portal.

## [1.1.0] - 2026-06-08

### Adicionado
- Arquivos:
  - `iss_fortaleza_automacao.py`

- Motivo:
  criação da automação visível do portal da ISS de Fortaleza com navegador controlado pelo usuário.

- Impacto:
  o programa passa a abrir o navegador, aguardar login manual, pedir a competência e preencher a aba de serviço com base na planilha `nf_compilado.xlsx`.

### Alterado
- Arquivos:
  - `iss_fortaleza_automacao.py`
  - `extrair_nf_pdfs.py`
  - `README.md`

- Motivo:
  alteração do formato de entrada da competência para mês numérico de `1` a `12` e ano em `AAAA`.

- Impacto:
  o usuário informa a competência de forma mais objetiva e o fluxo continua convertendo esse dado para o formato usado pela automação.

### Alterado
- Arquivos:
  - `extrair_nf_pdfs.py`
  - `README.md`
  - `requirements.txt`

- Motivo:
  integração da nova automação com a CLI existente e atualização do uso do projeto.

- Impacto:
  a opção 2 deixa de ser apenas uma abertura simples do site e passa a executar o fluxo guiado visível.

## [1.1.0] - 2026-06-08

### Adicionado
- Arquivo: `main.py`
- Motivo: criação de um ponto de entrada principal na raiz do projeto.
- Impacto: simplifica a execução do programa e centraliza o fluxo de uso para o usuário.

### Alterado
- Arquivo: `README.md`
- Motivo: atualização das instruções de execução para usar `main.py`.
- Impacto: deixa a documentação alinhada ao novo ponto de entrada.

## [1.1.0] - 2026-06-08

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: inclusão de um menu com duas opções, separando a geração do XLSX da abertura do portal da ISS de Fortaleza.
- Impacto: o usuário pode seguir o fluxo atual de escrituração e, ao final da opção 1, recebe uma sugestão para iniciar a opção 2.

### Documentação
- Arquivo: `README.md`
- Motivo: atualização do uso do script para refletir as opções `--modo 1` e `--modo 2`.
- Impacto: melhora a orientação operacional para o fluxo em duas etapas.

### Documentação
- Arquivo: `brain/2026-06-05-extracao-nfs.md`
- Motivo: registro da decisão de manter a abertura do site como etapa inicial da automação futura.
- Impacto: preserva o histórico técnico e o raciocínio da evolução do projeto.

## [1.1.0] - 2026-06-08

### Refatorado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: centralização de padrões reutilizáveis para CNPJ, número da NF, data, local de prestação, CNAE, descrição, valor, alíquota e ISS retido.
- Impacto: amplia a cobertura para novos layouts de NFS-e sem depender exclusivamente de correções por artista ou nome de arquivo.

### Adicionado
- Arquivo: `tests/test_extrair_nf_pdfs.py`
- Motivo: criação de validação automatizada com `unittest` para os PDFs de referência.
- Impacto: protege os 28 PDFs atuais contra regressões em futuras alterações.

### Documentação
- Arquivos:
  - `README.md`
  - `brain/2026-06-05-extracao-nfs.md`
- Motivo: registrar o novo fluxo de validação e a decisão técnica da refatoração.
- Impacto: melhora a rastreabilidade e orienta novas evoluções do parser.

### Corrigido
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: inclusão da `nf (9).pdf` como layout completo do bloco `Matheus & Kauan`, com preenchimento de `ID_CNAE`, `DESC_CNAE`, `DESCRICAO_SERVICO`, `VALOR_SERVICO` e `ALIQUOTA`.
- Impacto: a nota 9 passa a ser exportada para o XLSX em vez de ficar apenas no log de extração.

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: normalizacao de data para `DD/MM/AAAA`, remocao de zeros a esquerda no numero da NF, correcao do reconhecimento de local de prestacao, suporte ao motor `PyMuPDF`, ajuste para layouts adicionais da pasta `dados2/` e exclusao de linhas incompletas do XLSX.
- Impacto: melhora a confiabilidade da extracao, permite consolidar corretamente as notas validadas manualmente e evita exportar registros incompletos.

### Alterado
- Arquivos: `README.md`, `requirements.txt`, `brain/2026-06-05-extracao-nfs.md`
- Motivo: registrar a dependencia adicional `PyMuPDF` e documentar a atualizacao do parser.
- Impacto: mantem a documentacao e o historico tecnico alinhados com a nova implementacao.

## [1.0.0] - 2026-06-05

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: criação de log detalhado para NFs com campos faltantes.
- Impacto: facilita a auditoria dos PDFs que não preencheram todas as colunas do XLSX.

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: ampliar o parser para suportar os três modelos de NFS-e presentes na pasta `dados/`.
- Impacto: o XLSX agora consolida corretamente CNPJ, número da NF, data, CNAE, descrição do serviço, local de prestação, natureza da operação, ISS retido, valor do serviço e alíquota nos layouts testados.

### Adicionado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: automatizar a extração em lote de dados de NFS-e em PDF para consolidar em XLSX.
- Impacto: reduz trabalho manual e padroniza o compilado fiscal.

### Corrigido
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: ajuste das heurísticas de descrição do serviço e do cálculo de `ISS Retido`.
- Impacto: melhora a fidelidade dos campos extraídos no PDF de referência.

### Alterado
- Arquivos: `extrair_nf_pdfs.py`, `README.md`
- Motivo: inclusão de entrada interativa para o caminho da pasta quando o usuário não informar argumento na execução.
- Impacto: facilita o uso do script sem exigir parametrização obrigatória.

### Adicionado
- Arquivo: `PRD.md`
- Motivo: registrar o escopo funcional e as regras de negócio da automação.
- Impacto: melhora rastreabilidade e alinhamento funcional.

### Adicionado
- Arquivo: `README.md`
- Motivo: documentar o uso básico do script.
- Impacto: facilita execução por outros usuários da equipe.

### Adicionado
- Arquivo: `requirements.txt`
- Motivo: registrar as dependências Python utilizadas pela automação.
- Impacto: facilita reprodução do ambiente.

### Adicionado
- Arquivo: `.gitignore`
- Motivo: impedir versionamento de artefatos locais e da pasta `brain/`.
- Impacto: melhora organização do repositório.

### Adicionado
- Arquivo: `LICENSE`
- Motivo: atender à regra de licenciamento proprietário do projeto.
- Impacto: formaliza o uso restrito do material.

### Adicionado
- Arquivo: `brain/2026-06-05-extracao-nfs.md`
- Motivo: registrar as decisões técnicas e observações importantes desta entrega.
- Impacto: mantém histórico interno de contexto e decisões.

## [1.2.0] - 2026-06-09

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: a seleção de competência na função 2 passou a varrer a página principal e os frames do portal, com diagnóstico adicional quando os selects não são encontrados.
- Impacto: reduz travamentos na etapa de "Manter Escrituração" e melhora a compatibilidade com variações de layout do portal.

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: o navegador da função 2 passou a abrir diretamente no portal da ISS, removendo a etapa intermediária de navegação após a criação da janela.
- Impacto: deixa o início do fluxo mais simples, previsível e aderente ao uso assistido.

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: a função 2 passou a clicar explicitamente na aba `Serviços Tomados` antes de buscar o botão `Digitar Documento`.
- Impacto: evita travamento quando a tela exibe a guia, mas o fluxo não a aciona automaticamente.
