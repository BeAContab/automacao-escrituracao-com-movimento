# CHANGELOG

## [2026-06-10]

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: inserção de espera após alternar a pesquisa para `CNPJ`, para respeitar o reload parcial do portal antes da digitação.
- Impacto: o CNPJ deixa de ser apagado pelo redraw da tela e a etapa de autocomplete fica mais confiável.

## [2026-06-10]

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: a digitação passou a simular mais fielmente a entrada humana, com pausas entre teclas e validação por dígitos normalizados.
- Impacto: o CNPJ e outros campos mascarados tendem a permanecer visíveis na tela e a reagir melhor ao autocomplete do portal.

## [2026-06-10]

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: a digitação dos campos do portal passou a usar `fill()` com verificação de estabilidade para evitar o efeito de texto que aparece e desaparece rapidamente.
- Impacto: o CNPJ e demais campos mascarados ficam mais consistentes em tela antes da seleção do autocomplete e dos próximos cliques.

## [2026-06-10]

### Alterado
- Arquivos: `iss_fortaleza_automacao.py`, `extrair_nf_pdfs.py`, `README.md`
- Motivo: inclusão do modo `--debug-funcao2` para abrir o navegador em teste visível e pausar o fluxo nos pontos críticos da função 2.
- Impacto: a automação ficou mais fácil de depurar em tempo real e o seletor ambíguo do autocomplete do prestador passou a usar um contêiner mais específico.

## [2026-06-10]

### Alterado
- Arquivo: `extrair_nf_pdfs.py`
- Motivo: a função 1 passou a oferecer ao usuário a escolha entre iniciar a função 2 ou encerrar o programa logo após a geração do XLSX.
- Impacto: o fluxo ficou mais contínuo e o usuário pode seguir diretamente para a automação da ISS sem reiniciar o sistema.

## [2026-06-09]

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: reconhecimento de competência já visível na tela de `Manter Escrituração`.
- Impacto: quando a página já estiver em `Junho 2026`, a função 2 segue adiante sem tentar reprocessar os selects de competência.

## [2026-06-09]

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: reforço das esperas e detecção do estado já selecionado na tela de `Manter Escrituração`.
- Impacto: a função 2 deixa de travar tentando ler a competência cedo demais e passa a reconhecer quando Junho/2026 já está carregado.

## [2026-06-09]

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: substituição do preenchimento direto do campo `Código CNAE` por digitação real no modal `Pesquisar CNAE`.
- Impacto: todos os valores alimentados pela planilha na função 2 passam a ser digitados pelo teclado, sem colagem ou injeção direta no DOM.

## [2026-06-09]

### Alterado
- Arquivos: `extrair_nf_pdfs.py`, `README.md`
- Motivo: a automação da ISS passou a perguntar interativamente qual planilha XLSX será usada quando o usuário não informar o caminho explicitamente.
- Impacto: o fluxo fica mais claro para o operador e evita uso silencioso de uma planilha padrão sem confirmação.

## [2026-06-09]

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

## [2026-06-09]

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

## [2026-06-09]

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: inclusão da seleção da linha retornada pelo modal `Pesquisar CNAE` após a busca do código final.
- Impacto: a automação passa a concluir a etapa de CNAE selecionando o item `digitarDocumentoForm:idFormularioPesquisaCnae:idDatatableListaCnae:0:j_id453` antes de seguir o fluxo.

## [2026-06-09]

### Alterado
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: inclusão do preenchimento do modal `Pesquisar CNAE` com o valor de `ID_CNAE_FINAL` da planilha antes da continuação do fluxo.
- Impacto: a função 2 passa a consultar o CNAE final do documento de forma automática, em vez de depender de preenchimento manual nessa etapa.

## [2026-06-09]

### Corrigido
- Arquivo: `iss_fortaleza_automacao.py`
- Motivo: tratamento genérico para o caso em que o autocomplete do portal retorna "Nenhuma Razão Social Encontrada" ao pesquisar o CNPJ do prestador.
- Impacto: a função 2 passa a registrar o CNPJ ignorado em log, avançar para a próxima linha válida da planilha e continuar o fluxo sem interromper a automação.

## [2026-06-09]

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

## [2026-06-08]

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

## [2026-06-08]

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

## [2026-06-08]

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

## [2026-06-08]

### Adicionado
- Arquivo: `main.py`
- Motivo: criação de um ponto de entrada principal na raiz do projeto.
- Impacto: simplifica a execução do programa e centraliza o fluxo de uso para o usuário.

### Alterado
- Arquivo: `README.md`
- Motivo: atualização das instruções de execução para usar `main.py`.
- Impacto: deixa a documentação alinhada ao novo ponto de entrada.

## [2026-06-08]

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

## [2026-06-08]

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

## [2026-06-05]

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

## [2026-06-09]

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
