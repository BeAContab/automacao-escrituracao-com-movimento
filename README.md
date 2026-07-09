# Automação de Escrituração com Movimento

O **Automação de Escrituração com Movimento** é uma solução inteligente desenvolvida para otimizar, acelerar e garantir a segurança do processo de escrituração e conformidade fiscal de Notas Fiscais de Serviços Tomados (NFS-e) a partir de arquivos digitais XML.

Projetada com foco em eficiência operacional, a ferramenta elimina a necessidade de digitação manual de notas fiscais, minimizando erros humanos e trazendo agilidade para os setores de contabilidade e fiscal.

---

## 🚀 Proposta de Valor

A digitação e conferência manual de NFS-e é um processo repetitivo, lento e altamente suscetível a falhas operacionais que podem gerar autuações e inconsistências fiscais. Esta solução automatiza o ciclo completo da escrituração de serviços tomados:

* **Economia de Tempo:** Processamento e consolidação de múltiplos arquivos XML em segundos.
* **Redução de Erros:** Captura direta dos dados fiscais estruturados do documento digital original (CNPJ, valores, CNAE, retenções).
* **Conformidade Tributária:** Mapeamento preciso de CNAEs oficiais de serviços tomados e aplicação automatizada das regras locais de retenção de ISS.
* **Segurança e Auditoria:** Registros detalhados de cada nota processada e logs de auditoria para acompanhamento contínuo.

---

## ✨ Funcionalidades Principais

A solução é dividida em duas grandes etapas integradas que cobrem desde a recepção dos arquivos XML até a gravação definitiva no portal do município:

### 1. Processamento de XMLs e Inteligência Artificial (Tess AI)
* **Parser de XML Fiel ao Padrão:** Leitura recursiva resiliente de arquivos XML de NFS-e (DPS/NFS-e Padrão Nacional ou equivalentes), parseando estruturas complexas com segurança e ignorando namespaces dinâmicos.
* **Classificação CNAE via Tess AI:** Enriquecimento inteligente das notas fiscais. O sistema extrai as atividades originais e a discriminação dos serviços e consulta a IA do Tess AI para selecionar obrigatoriamente e com exatidão a correspondente classificação oficial contida no banco de dados do `cnae_oficial.xlsx`.
* **Regras Fiscais de Natureza de Operação e Retenção:** Determinação programática da natureza da operação (Tributação no Município para prestadores de Fortaleza-CE e Tributação Fora para os demais) e da retenção estrita de ISSQN (baseado nas tags do XML e CNAEs específicos de eventos que forçam ISS retido).
* **Consolidação em Planilha:** Geração de relatórios consolidados em Excel contendo todos os dados estruturados por linha no modelo fiscal exato do escritório.

### 2. Automação de Escrituração (Integração com Portal ISS)
* **Navegação Inteligente:** Controle automatizado de navegador (Selenium WebDriver) para acessar o Portal da ISS de Fortaleza.
* **Preenchimento Guiado:** Alimentação automatizada de todos os campos da escrituração (CNPJ, número da nota, data de emissão, código do serviço, descrição, valores e retenção de ISS).
* **Filtragem de Prestadores Locais:** Pulagem automática preventiva de notas cujo prestador seja estabelecido em Fortaleza-CE (pois estas já constam do sistema da prefeitura), abrindo exceção para escriturar notas de prestadores classificados no regime tributário **MEI**.
* **Segurança na Gravação:** A automação realiza o preenchimento dos formulários e aguarda na etapa final, permitindo que um operador humano valide visualmente a tela antes de confirmar a gravação definitiva no portal.

---

## 👥 Público-Alvo

* **Escritórios de Contabilidade:** Que lidam com grandes volumes mensais de escrituração fiscal de seus clientes.
* **Departamentos Fiscais e Financeiros:** Equipes internas de empresas que precisam escriturar serviços tomados de forma ágil e segura.
* **Analistas Tributários:** Profissionais que buscam auditar e conferir a consistência dos dados de suas NFS-e de maneira centralizada.

---

## 📖 Como Funciona? (Visão Geral)

1. **Coleta de XMLs:** O usuário armazena os arquivos XML de NFS-e recebidos em uma pasta do computador.
2. **Processamento e Classificação:** O sistema lê os arquivos, classifica os CNAEs com o Tess AI e cria a planilha consolidada sem a necessidade de intervenção manual.
3. **Escrituração no Portal:** A partir da planilha gerada, o robô abre o portal da prefeitura, realiza as buscas e preenche os dados do prestador e do serviço nota por nota, agilizando o fluxo fiscal diário.

---

> [!NOTE]  
> Para obter detalhes de instalação, configuração de chaves de API, variáveis de ambiente e instruções técnicas de execução do projeto, consulte o [Manual do Desenvolvedor (DEVELOPER.md)](file:///c:/Users/gabriel.lima/Documents/GitHub/automacao-escrituracao-com-movimento/DEVELOPER.md).
