# Automações ISS

### Plataforma de Robótica de Processos Contábeis (RPA) para Escrituração de ISSQN

**Automações ISS** é a solução desenvolvida pela **BeAContab** para o escritório **Barreira & Associados**, que substitui a digitação manual de Notas Fiscais de Serviços Eletrônicas (NFS-e) por um fluxo automatizado, auditável e seguro — da leitura do XML original até a escrituração e o encerramento junto ao portal da Prefeitura de Fortaleza (ISS Fortaleza).

O sistema roda como um aplicativo desktop nativo do Windows, com interface 100% visual — sem linha de comando, sem planilhas macro e sem dependência de conhecimento técnico do operador.

---

## 🚀 Proposta de Valor

A conferência e digitação manual de notas fiscais é um dos gargalos operacionais mais caros de um departamento fiscal: é repetitiva, consome horas de um profissional qualificado e está sujeita a erros que podem gerar autuações, glosas e retrabalho. **Automações ISS** ataca esse gargalo em toda a sua extensão — não só na digitação, mas também na captura das notas de origem e no encerramento das empresas sem movimento.

| Ganho | Como o sistema entrega |
|---|---|
| **Tempo** | Centenas de XMLs são lidos, classificados e escriturados em minutos, sem intervenção linha a linha do operador. |
| **Precisão** | Os dados fiscais (CNPJ, valores, CNAE, retenções) saem direto do XML estruturado — não são redigitados, eliminando erro de transcrição. |
| **Conformidade** | Classificação de CNAE por Inteligência Artificial e aplicação automática das regras locais de retenção de ISS e natureza da operação vigentes em Fortaleza/CE. |
| **Rastreabilidade** | Cada execução gera um log narrativo, com timestamp, salvo automaticamente junto ao resultado — pronto para auditoria interna ou suporte técnico. |
| **Governança** | Nenhuma automação grava informação no portal da Prefeitura sem autenticação prévia do operador; ações irreversíveis (como encerramento de escrituração) exigem conferência humana antes de rodar. |

---

## 📦 As Automações do Sistema

O sistema reúne automações integradas, acessíveis por abas na interface, organizadas em três grupos: as **Funções Principais** (fluxo de um único CNPJ/login por vez), o grupo **Multi-CNPJ** (a mesma escrituração para várias empresas numa só execução) e **Outras Funções** (utilitários e casos auxiliares).

### Funções Principais

| Função | O que faz | Situação |
|---|---|---|
| **Baixar NFS-e — Portal Nacional** | Baixa em lote os XMLs de NFS-e direto da API do Portal Nacional (ADN), autenticando por certificado digital A1 | 🧪 Em testes |
| **Processar XMLs** | Lê uma pasta de XMLs de NFS-e e gera a planilha fiscal consolidada, com CNAE classificado por IA | ✅ Estável |
| **Escrituração** | Escritura as notas da planilha diretamente no portal ISS Fortaleza, nota a nota | 🧪 Em testes |

### Multi-CNPJ

| Função | O que faz | Situação |
|---|---|---|
| **Processar XMLs — Multi-CNPJ** | Mesmo processamento acima, mas consolidando várias pastas de CNPJ (ex.: saída de Baixar NFS-e — Portal Nacional) numa única planilha, com as colunas de login/senha do portal ISS por CNPJ | ✅ Estável |
| **Escrituração — Multi-CNPJ** | Escritura as notas de várias empresas numa só execução: para cada CNPJ, faz login com as credenciais da própria planilha, confere a inscrição certa e escritura, sem repetir login a cada empresa. Suporta simular (preenche os campos e não grava) antes de escriturar de verdade | 🧪 Em testes |

### Outras Funções

| Função | O que faz | Situação |
|---|---|---|
| **Exportar XML de Prestados** | Baixa em lote os XMLs de notas emitidas (serviços prestados) direto do portal | ✅ Estável |
| **Captura Escrituração (Com Movimento)** | Baixa em lote os certificados de escrituração de empresas encerradas com movimento | 🧪 Em testes |
| **Encerramento ISS (Sem Movimento)** | Encerra a escrituração de empresas sem movimento/inativas e emite relatório consolidado | 🧪 Em testes |
| **Encerramento ISS — Múltiplos Meses** | A mesma automação acima, mas para várias competências numa só execução — por empresa, processa todos os meses selecionados antes de trocar de empresa | 🧪 Em testes |
| **Importar para Athenas** | Converte a escrituração exportada do ISS Fortaleza para o layout de importação do sistema Athenas ERP, aplicando as regras de CFOP e retenções federais | 🧪 Em testes |

> As funções marcadas **Em testes** já operam ponta a ponta em produção; recomenda-se conferência manual dos resultados até a validação completa pela equipe técnica.

---

## ✨ Detalhamento das Funções

### Baixar NFS-e — Portal Nacional — Download Direto via API
Automatiza o download em lote de NFS-e do Portal Nacional (Ambiente de Dados Nacional/ADN), sem depender de navegador:
* **Seletor nativo de certificado do Windows:** o mesmo diálogo "Selecionar um Certificado" usado pelo Chrome/Edge — basta escolher entre os certificados já instalados; a chave privada nunca sai do repositório do Windows. Funciona com A1 e, potencialmente, com A3 (token/smartcard). Um modo alternativo por arquivo `.pfx`/`.p12` + senha continua disponível para quem não tem o certificado importado no Windows.
* **Varredura por NSU:** consulta a API oficial do governo a partir de qualquer ponto de partida, com retomada automática a partir do último NSU processado com sucesso.
* **Filtro por filial:** suporte a CNPJ completo de 14 dígitos para restringir a busca a um estabelecimento específico.
* **Organização automática:** os XMLs baixados já saem organizados por CNPJ, tipo de nota (emitidas/tomadas) e competência, prontos para alimentar a função de Processar XMLs sem etapas manuais.
* **Segurança:** a senha do certificado nunca é salva em disco; a chave privada só existe em texto plano durante a execução, em pasta temporária apagada ao final (com sucesso, erro ou cancelamento).

### Processar XMLs — Leitura e Classificação Inteligente
Transforma uma pasta de arquivos XML de NFS-e (Portal Nacional ou Portal Da Paraíba, com detecção automática de layout) em uma planilha Excel pronta para escrituração:
* **Parser resiliente:** leitura recursiva dos XMLs, robusta a namespaces e pequenas variações de layout entre municípios/emissores.
* **Classificação de CNAE via IA (Tess AI):** a descrição do serviço e a atividade original de cada nota são enviadas a um agente de IA treinado para apontar o CNAE oficial correspondente na base do escritório — com cache local para não reclassificar o mesmo código de serviço duas vezes.
* **Regras fiscais automáticas:** natureza da operação (tributação no município x fora do município) e retenção de ISS aplicadas de forma programática, incluindo overrides por CNAE e detecção de prestadores MEI.
* **Suporte a múltiplos lotes:** se a pasta selecionada tiver subpastas (ex.: uma por cliente), o sistema gera uma planilha independente por subpasta, sem misturar notas de origens diferentes.
* **Pausar, continuar e parar** a qualquer momento; ao parar, o que já foi processado fica salvo, e retomar a mesma pasta continua de onde parou.

### Escrituração — Lançamento Direto no Portal
Elimina a digitação manual no portal da Prefeitura:
* **Preenchimento guiado por Selenium:** todos os campos da escrituração (CNPJ, número, data, CNAE, alíquota, valores, retenção de ISS) são preenchidos automaticamente a partir da planilha gerada em Processar XMLs.
* **Resolução automática de bloqueios do portal:** nota duplicada, avisos de "prestador não inscrito no CPOM" e o erro "CNAE não incide ISS" são reconhecidos e tratados sem intervenção do operador.
* **Filtro inteligente de prestadores locais:** notas de prestadores já registrados no sistema da Prefeitura (Fortaleza/CE) são puladas automaticamente, com exceção aberta para o regime MEI.
* **Controle total do operador:** login sempre manual; pausa, retomada e cancelamento disponíveis a qualquer momento durante a execução.

### Processar XMLs — Multi-CNPJ
Mesma leitura/classificação de XMLs da função principal, mas pensada para quem administra várias empresas:
* Consolida várias pastas de CNPJ (uma por empresa) numa única planilha de saída, em vez de uma planilha por pasta.
* A planilha final ganha três colunas a mais — `CNPJ_TOMADOR`, `LOGIN` e `SENHA` — preenchidas manualmente pelo operador, que alimentam a Escrituração — Multi-CNPJ na sequência.
* Mesmo controle de pausar/continuar/parar da função individual.

### Escrituração — Multi-CNPJ
Estende a Escrituração para múltiplas empresas numa só execução, sem repetir login a cada uma:
* **Etapa de validação de acessos:** para cada empresa da planilha, faz login com o CPF/senha correspondente, seleciona a inscrição pelo CNPJ e confere que é a empresa certa antes de prosseguir — detectando login recusado, empresa não encontrada ou CNPJ divergente sem interromper as demais.
* **Simulação sem gravação:** antes de escriturar de verdade, é possível rodar em modo simulação — os campos são preenchidos exatamente como na escrituração real, mas nada é gravado no portal.
* **Escrituração real com retomada:** ao escriturar de fato, o sistema guarda o que já foi concluído por competência; reexecuções pulam empresas já concluídas e refazem só o que falhou, sem duplicar notas.
* **Empresa com problema não trava as demais:** login recusado, CNPJ divergente ou erro de portal marcam aquela empresa como pendência e o robô segue para a próxima.

### Exportar XML de Prestados — Captura de Notas Emitidas
Automatiza a coleta das notas fiscais emitidas pelo próprio escritório/cliente:
* **Download em lote** direto do portal ISS Fortaleza, sem necessidade de exportar página por página manualmente.
* **Filtro por competência:** basta informar mês e ano; o robô consulta, seleciona e exporta todas as notas do período automaticamente.
* **Pausar, continuar e parar** entre uma página e outra do laço de exportação, sem fechar o navegador nem perder os arquivos já baixados.

### Captura Escrituração (Com Movimento) — Certificados em Lote
Automatiza a coleta de comprovantes fiscais de empresas encerradas com movimento:
* Lê a planilha fiscal do escritório, filtra as empresas de Fortaleza encerradas numa data específica (excluindo as "sem movimento") e baixa o certificado de escrituração de cada uma, organizado automaticamente por ano/mês na pasta de saída.
* Login automático com CPF/senha informados na própria tela, sem etapas manuais no navegador.

### Encerramento ISS (Sem Movimento) — Baixa e Relatório Consolidado
Automatiza o encerramento em massa de empresas inativas:
* Identifica empresas sem movimento/inativas na planilha fiscal, verifica pendências de serviços prestados antes de agir, encerra a escrituração das aptas, baixa o certificado em PDF e grava a data de encerramento de volta na planilha.
* Ao final, apresenta um resumo objetivo (processadas / encerradas / com problema) e gera um relatório Excel detalhado para conferência.

### Encerramento ISS — Múltiplos Meses
Mesma automação acima, para quando é preciso encerrar mais de uma competência de uma vez:
* O operador monta uma lista de competências (mês/ano); para cada empresa da planilha, o robô processa todos os meses selecionados antes de trocar de empresa — evita repetir a busca da inscrição a cada mês.
* Gera um relatório de execução por competência, organizado em pastas por ano-mês.
* Um problema numa competência específica não impede as demais competências da mesma empresa, nem as demais empresas.

### Importar para Athenas
Converte a escrituração exportada do ISS Fortaleza para o layout de importação do sistema **Athenas ERP**:
* Lê as abas de Serviços Tomados e Serviços Prestados de um ou mais arquivos do ISS Fortaleza e gera as planilhas correspondentes já no layout do Athenas, com CFOP, split de PIS/COFINS/CSLL, ISS retido/próprio e ajuste automático de data pela competência.
* Regime tributário (Normal ou Simples Nacional) configurável por arquivo, já que um lote pode reunir empresas de regimes diferentes.
* Uma empresa com erro não interrompe as demais; ao final, um relatório de erros detalha o que falhou e por quê.
* Não usa portal nem navegador — é só leitura e escrita de planilha.

---

## 🛡️ Segurança, Conformidade e Auditoria

* **Login manual nas ações mais sensíveis:** Escrituração e Exportar XML de Prestados exigem autenticação manual do operador no portal antes de qualquer lançamento — a automação nunca conhece nem armazena a senha do usuário nesses fluxos. Já a Escrituração — Multi-CNPJ recebe login/senha diretamente na planilha (única forma viável para várias empresas numa execução só); a senha fica só em memória durante a execução e nunca aparece em log, mensagem de erro ou arquivo de retomada.
* **Certificado digital nunca persistido:** em Baixar NFS-e — Portal Nacional, a senha do certificado A1 não é salva em disco e a chave privada só existe em texto plano durante a execução, em pasta temporária apagada ao final.
* **Ações irreversíveis sinalizadas:** o encerramento de escrituração (Encerramento ISS, nas duas versões) é tratado como uma ação real e de difícil reversão, com aviso explícito na interface antes da execução; a Escrituração — Multi-CNPJ oferece um modo de simulação que preenche os campos sem gravar nada, para conferência antes de escriturar de verdade.
* **Log de auditoria por execução:** toda automação grava, junto ao resultado gerado, um arquivo de log com timestamp de cada etapa — histórico pronto para conferência interna ou suporte técnico.
* **Rede de segurança contra dados incompletos:** notas com campos obrigatórios ausentes ou inconsistentes são automaticamente puladas e registradas em log próprio, em vez de escrituradas incorretamente.
* **Credenciais protegidas localmente:** chaves de API e credenciais de portal ficam armazenadas apenas na máquina do operador, nunca em servidores de terceiros.

---

## 👥 Público-Alvo

* **Escritórios de contabilidade** que lidam com grande volume mensal de escrituração fiscal de múltiplos clientes.
* **Departamentos fiscais e financeiros** que precisam escriturar serviços tomados e prestados com agilidade e segurança.
* **Analistas tributários** que buscam centralizar e auditar a consistência dos dados de NFS-e antes do lançamento oficial.

---

## 📖 Como Funciona — Visão de Ponta a Ponta

**Fluxo de uma empresa por vez:**
1. **Captura:** o operador reúne os XMLs de serviços tomados (recebidos), usa Exportar XML de Prestados para os emitidos, ou baixa direto do Portal Nacional via certificado A1.
2. **Processamento:** Processar XMLs lê os arquivos, classifica os CNAEs com apoio de IA e gera a planilha fiscal consolidada.
3. **Escrituração:** essa planilha alimenta a Escrituração, que lança as notas no portal ISS Fortaleza, nota a nota, com supervisão do operador.
4. **Encerramento e certificação:** Captura Escrituração (Com Movimento) e Encerramento ISS (Sem Movimento) cuidam do ciclo de vida das empresas — capturando certificados de quem encerrou com movimento e encerrando/baixando certificado de quem está sem movimento.

**Fluxo de várias empresas de uma vez (Multi-CNPJ):**
5. Processar XMLs — Multi-CNPJ consolida várias pastas de CNPJ numa única planilha, já com as colunas de login/senha do portal por empresa.
6. Escrituração — Multi-CNPJ usa essa planilha para escriturar todas as empresas numa só execução, validando o acesso (e simulando, se preferir) antes de escriturar de verdade.
7. Encerramento ISS — Múltiplos Meses estende o encerramento de empresas sem movimento para várias competências de uma vez.

**Integração com outros sistemas:**
8. Importar para Athenas pega a escrituração já exportada do ISS Fortaleza e gera as planilhas prontas para importar no Athenas ERP.

---

## 💻 Requisitos

* Windows 10 ou superior (64 bits)
* Google Chrome instalado e atualizado
* Acesso à internet (portal ISS Fortaleza, API do Tess AI e API do Portal Nacional/ADN)
* Credenciais válidas do portal ISS Fortaleza
* Chave de API e ID do Agente do Tess AI (para Processar XMLs)
* Certificado digital A1 (ou A3) válido, instalado no repositório de certificados do Windows (ou em arquivo `.pfx`/`.p12`), para Baixar NFS-e — Portal Nacional

---

## 📄 Confidencialidade

Software de propriedade privada da **Barreira & Associados** — uso restrito às pessoas, empresas e equipes autorizadas pelo proprietário. Consulte o arquivo [LICENSE](LICENSE) para os termos completos.

---

> [!NOTE]
> Para o manual de operação completo (passo a passo de cada função, telas e solução de problemas), consulte o [Manual do Sistema](manual_sistema.html).
> Para instalação, configuração de chaves de API, variáveis de ambiente e instruções técnicas de execução do projeto, consulte o [Manual do Desenvolvedor (DEVELOPER.md)](DEVELOPER.md).
