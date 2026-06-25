# CHANGELOG

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
