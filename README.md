# Automação de Escrituração com Movimento

O **Automação de Escrituração com Movimento** é uma solução inteligente desenvolvida para otimizar, acelerar e garantir a segurança do processo de escrituração e conformidade fiscal de Notas Fiscais de Serviços Tomados (NFS-e) a partir de arquivos digitais XML.

Projetada com foco em eficiência operacional, a ferramenta opera **100% de forma visual e interativa (GUI)**, eliminando a digitação manual de notas fiscais, minimizando erros humanos e trazendo agilidade para os setores de contabilidade e fiscal.

---

## 🚀 Proposta de Valor

A digitação e conferência manual de NFS-e é um processo repetitivo, lento e altamente suscetível a falhas operacionais que podem gerar autuações e inconsistências fiscais. Esta solução automatiza o ciclo completo das notas e XMLs de serviços do município:

* **Economia de Tempo:** Processamento e consolidação de múltiplos arquivos XML em segundos.
* **Redução de Erros:** Captura direta dos dados fiscais estruturados do documento digital original (CNPJ, valores, CNAE, retenções).
* **Conformidade Tributária:** Mapeamento preciso de CNAEs oficiais de serviços tomados e aplicação automatizada das regras locais de retenção de ISS.
* **Segurança e Auditoria:** Registros detalhados de cada nota processada e logs de auditoria na tela do operador em tempo real.
* **Operação Visual:** Interface desktop amigável e segura, sem necessidade de lidar com telas de terminal ou comandos de linha (CLI).

---

## ✨ Funcionalidades Principais

A solução é dividida em três grandes módulos integrados acessíveis diretamente pelas abas da interface gráfica:

### 1. Processamento de XMLs e Inteligência Artificial (Tess AI)
* **Parser de XML Fiel ao Padrão:** Leitura recursiva resiliente de arquivos XML de NFS-e, parseando estruturas complexas com segurança e ignorando namespaces dinâmicos.
* **Classificação CNAE via Tess AI:** Enriquecimento inteligente das notas fiscais. O sistema extrai as atividades originais e a discriminação dos serviços e consulta a IA do Tess AI para selecionar obrigatoriamente e com exatidão a correspondente classificação oficial contida no banco de dados do `cnae_oficial.xlsx`.
* **Regras Fiscais de Natureza de Operação e Retenção:** Determinação programática da natureza da operação (Tributação no Município para prestadores de Fortaleza-CE e Tributação Fora para os demais) e da retenção estrita de ISSQN (baseado nas tags do XML e CNAEs específicos de eventos que forçam ISS retido).
* **Consolidação em Planilha:** Geração de relatórios consolidados em Excel contendo todos os dados estruturados por linha no modelo fiscal exato do escritório.

### 2. Automação de Escrituração (Integração com Portal ISS)
* **Navegação Inteligente:** Controle automatizado de navegador (Selenium WebDriver) para acessar o Portal da ISS de Fortaleza.
* **Preenchimento Guiado:** Alimentação automatizada de todos os campos da escrituração (CNPJ, número da nota, data de emissão, código do serviço, descrição, valores e retenção de ISS).
* **Filtragem de Prestadores Locais:** Pulagem automática preventiva de notas cujo prestador seja estabelecido em Fortaleza-CE (pois estas já constam do sistema da prefeitura), abrindo exceção para escriturar notas de prestadores classificados no regime tributário **MEI**.
* **Segurança na Gravação:** O login no portal é sempre manual, com confirmação explícita do operador antes de a automação iniciar. A partir daí, o preenchimento e a gravação de cada nota (botão "Gravar Documento") acontecem automaticamente nota a nota; o operador pode pausar e retomar o robô a qualquer momento pela interface, e notas com campos obrigatórios ausentes ou erros de preenchimento são automaticamente puladas e registradas em log (`log_notas_incompletas.txt`), em vez de gravadas no portal.

### 3. Exportar XML de Prestados
* **Download em Lote de XMLs:** Automatiza o download de arquivos XML de NFS-e (Serviços Prestados) diretamente do portal da ISS Fortaleza.
* **Filtro de Competência:** Permite selecionar o mês e o ano de interesse. O robô filtra as notas do período no painel de consulta e realiza a exportação sequencial por lote.

---

## 👥 Público-Target

* **Escritórios de Contabilidade:** Que lidam com grandes volumes mensais de escrituração fiscal de seus clientes.
* **Departamentos Fiscais e Financeiros:** Equipes internas de empresas que precisam escriturar serviços tomados de forma ágil e segura.
* **Analistas Tributários:** Profissionais que buscam auditar e conferir a consistência dos dados de suas NFS-e de maneira centralizada.

---

## 📖 Como Funciona? (Visão Geral)

1. **Recepção/Consulta de XMLs:** O operador seleciona a pasta de XMLs de serviços tomados ou usa o robô exportador na interface gráfica para extrair notas emitidas (prestadas).
2. **Processamento e Classificação:** O sistema processa os arquivos de entrada, classifica os CNAEs com apoio de Inteligência Artificial e gera uma planilha formatada contendo todas as notas fiscais do lote.
3. **Escrituração no Portal:** A partir da planilha gerada, o robô Selenium preenche os formulários do portal da prefeitura de Fortaleza-CE com segurança e precisão.

---

> [!NOTE]  
> Para obter detalhes de instalação, configuração de chaves de API, variáveis de ambiente e instruções técnicas de execução do projeto, consulte o [Manual do Desenvolvedor (DEVELOPER.md)](file:///c:/Users/gabriel.lima/Documents/GitHub/automacao-escrituracao-com-movimento/DEVELOPER.md).
