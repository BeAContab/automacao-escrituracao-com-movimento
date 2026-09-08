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

## 📦 As 5 Automações do Sistema

O sistema reúne cinco automações integradas, acessíveis por abas na interface, cobrindo o ciclo fiscal completo de ISSQN de tomados e prestados:

| # | Função | O que faz | Situação |
|---|---|---|---|
| 1 | **Processar XMLs** | Lê uma pasta de XMLs de NFS-e e gera a planilha fiscal consolidada, com CNAE classificado por IA | ✅ Estável |
| 2 | **Automação: Escrituração** | Escritura as notas da planilha diretamente no portal ISS Fortaleza, nota a nota | 🧪 Em testes |
| 3 | **Exportar XML de Prestados** | Baixa em lote os XMLs de notas emitidas (serviços prestados) direto do portal | ✅ Estável |
| 4 | **Captura Escrituração (Com Movimento)** | Baixa em lote os certificados de escrituração de empresas encerradas com movimento | 🧪 Em testes |
| 5 | **Encerramento ISS (Sem Movimento)** | Encerra a escrituração de empresas sem movimento/inativas e emite relatório consolidado | 🧪 Em testes |

> As funções marcadas **Em testes** já operam ponta a ponta em produção; recomenda-se conferência manual dos resultados até a validação completa pela equipe técnica.

---

## ✨ Detalhamento das Funções

### 1. Processar XMLs — Leitura e Classificação Inteligente
Transforma uma pasta de arquivos XML de NFS-e (Portal Nacional ou Portal Da Paraíba, com detecção automática de layout) em uma planilha Excel pronta para escrituração:
* **Parser resiliente:** leitura recursiva dos XMLs, robusta a namespaces e pequenas variações de layout entre municípios/emissores.
* **Classificação de CNAE via IA (Tess AI):** a descrição do serviço e a atividade original de cada nota são enviadas a um agente de IA treinado para apontar o CNAE oficial correspondente na base do escritório — com cache local para não reclassificar o mesmo código de serviço duas vezes.
* **Regras fiscais automáticas:** natureza da operação (tributação no município x fora do município) e retenção de ISS aplicadas de forma programática, incluindo overrides por CNAE e detecção de prestadores MEI.
* **Suporte a múltiplos lotes:** se a pasta selecionada tiver subpastas (ex.: uma por cliente), o sistema gera uma planilha independente por subpasta, sem misturar notas de origens diferentes.

### 2. Automação: Escrituração — Lançamento Direto no Portal
Elimina a digitação manual no portal da Prefeitura:
* **Preenchimento guiado por Selenium:** todos os campos da escrituração (CNPJ, número, data, CNAE, alíquota, valores, retenção de ISS) são preenchidos automaticamente a partir da planilha gerada na Função 1.
* **Resolução automática de bloqueios do portal:** nota duplicada, avisos de "prestador não inscrito no CPOM" e o erro "CNAE não incide ISS" são reconhecidos e tratados sem intervenção do operador.
* **Filtro inteligente de prestadores locais:** notas de prestadores já registrados no sistema da Prefeitura (Fortaleza/CE) são puladas automaticamente, com exceção aberta para o regime MEI.
* **Controle total do operador:** login sempre manual; pausa, retomada e cancelamento disponíveis a qualquer momento durante a execução.

### 3. Exportar XML de Prestados — Captura de Notas Emitidas
Automatiza a coleta das notas fiscais emitidas pelo próprio escritório/cliente:
* **Download em lote** direto do portal ISS Fortaleza, sem necessidade de exportar página por página manualmente.
* **Filtro por competência:** basta informar mês e ano; o robô consulta, seleciona e exporta todas as notas do período automaticamente.

### 4. Captura Escrituração (Com Movimento) — Certificados em Lote
Automatiza a coleta de comprovantes fiscais de empresas encerradas com movimento:
* Lê a planilha fiscal do escritório, filtra as empresas de Fortaleza encerradas numa data específica (excluindo as "sem movimento") e baixa o certificado de escrituração de cada uma, organizado automaticamente por ano/mês na pasta de saída.
* Login automático com CPF/senha informados na própria tela, sem etapas manuais no navegador.

### 5. Encerramento ISS (Sem Movimento) — Baixa e Relatório Consolidado
Automatiza o encerramento em massa de empresas inativas:
* Identifica empresas sem movimento/inativas na planilha fiscal, verifica pendências de serviços prestados antes de agir, encerra a escrituração das aptas, baixa o certificado em PDF e grava a data de encerramento de volta na planilha.
* Ao final, apresenta um resumo objetivo (processadas / encerradas / com problema) e gera um relatório Excel detalhado para conferência.

---

## 🛡️ Segurança, Conformidade e Auditoria

* **Login manual nas ações mais sensíveis:** as Funções 2 e 3 exigem autenticação manual do operador no portal antes de qualquer lançamento — a automação nunca conhece nem armazena a senha do usuário nesses fluxos.
* **Ações irreversíveis sinalizadas:** o encerramento de escrituração (Função 5) é tratado como uma ação real e de difícil reversão, com aviso explícito na interface antes da execução.
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

1. **Captura:** o operador reúne os XMLs de serviços tomados (recebidos) ou usa a Função 3 para exportar os XMLs de serviços prestados direto do portal.
2. **Processamento:** a Função 1 lê os XMLs, classifica os CNAEs com apoio de IA e gera a planilha fiscal consolidada.
3. **Escrituração:** a Função 2 usa essa planilha para lançar as notas no portal ISS Fortaleza, nota a nota, com supervisão do operador.
4. **Encerramento e certificação:** as Funções 4 e 5 cuidam do ciclo de vida das empresas — capturando certificados de quem encerrou com movimento e encerrando/baixando certificado de quem está sem movimento.

---

## 💻 Requisitos

* Windows 10 ou superior (64 bits)
* Google Chrome instalado e atualizado
* Acesso à internet (portal ISS Fortaleza e API do Tess AI)
* Credenciais válidas do portal ISS Fortaleza
* Chave de API e ID do Agente do Tess AI (para a Função 1)

---

## 📄 Confidencialidade

Software de propriedade privada da **Barreira & Associados** — uso restrito às pessoas, empresas e equipes autorizadas pelo proprietário. Consulte o arquivo [LICENSE](LICENSE) para os termos completos.

---

> [!NOTE]
> Para o manual de operação completo (passo a passo de cada função, telas e solução de problemas), consulte o [Manual do Sistema](manual_sistema.html).
> Para instalação, configuração de chaves de API, variáveis de ambiente e instruções técnicas de execução do projeto, consulte o [Manual do Desenvolvedor (DEVELOPER.md)](DEVELOPER.md).
