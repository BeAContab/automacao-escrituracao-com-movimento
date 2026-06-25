# Automação de Escrituração com Movimento

O **Automação de Escrituração com Movimento** é uma solução inteligente desenvolvida para otimizar, acelerar e garantir a segurança do processo de escrituração e conformidade fiscal de Notas Fiscais de Serviços Tomados (NFS-e). 

Projetada com foco em eficiência operacional, a ferramenta elimina a necessidade de digitação manual de notas fiscais, minimizando erros humanos e trazendo agilidade para os setores de contabilidade e fiscal.

---

## 🚀 Proposta de Valor

A digitação e conferência manual de NFS-e é um processo repetitivo, lento e altamente suscetível a falhas operacionais que podem gerar autuações e inconsistências fiscais. Esta solução automatiza o ciclo completo da escrituração de serviços tomados:

* **Economia de Tempo:** Processamento em lote de múltiplos documentos em segundos.
* **Redução de Erros:** Captura direta dos dados fiscais do documento original (CNPJ, valores, CNAE, retenções).
* **Conformidade Tributária:** Mapeamento preciso de CNAEs oficiais e aplicação automatizada das regras locais de retenção de ISS.
* **Segurança e Auditoria:** Registros detalhados de cada nota processada e logs de auditoria para acompanhamento contínuo.

---

## ✨ Funcionalidades Principais

A solução é dividida em duas grandes etapas integradas que cobrem desde a recepção dos arquivos digitais até a interface do portal do município:

### 1. Extração Inteligente e Consolidação (Leitura de NFS-e)
* **Parser de Multi-Layout:** Reconhecimento automático de notas de diferentes prefeituras (como São Paulo, Fortaleza e Petrolina).
* **Processamento 100% via Inteligência Artificial:** Extração inteligente e direta de todos os campos fiscais de NFS-e diretamente através da IA do Google Gemini.
* **Resiliência e Retentativa (Retry):** Mecanismo automático para contornar oscilações de limites de cota de API (RPM/TPM), garantindo a conclusão de lotes com pausas inteligentes.
* **Validação Anti-Alucinação Dinâmica:** Regras no prompt e validações programáticas em Python (como verificação de dígitos de CNPJ e data de emissão dentro do limite de 3 meses da data corrente) para garantir consistência fiscal.
* **Enriquecimento e Classificação de CNAE:** Mapeamento inteligente de atividades contra a tabela oficial de CNAE utilizando correspondência semântica.
* **Consolidação em Planilha:** Geração de relatórios prontos em Excel contendo todos os dados auditados estruturados por linha.

### 2. Automação Assistida (Integração com Portal ISS)
* **Navegação Inteligente:** Controle automatizado de navegador (Playwright) para acessar o Portal da ISS de Fortaleza.
* **Preenchimento Guiado:** Alimentação automatizada dos formulários da escrituração (CNPJ, número, data, código do serviço, descrição, valores, local de prestação e retenção de ISS).
* **Segurança na Gravação:** A automação realiza todo o trabalho pesado de digitação e para o formulário na etapa final, permitindo que um operador humano valide visualmente a tela antes de confirmar a gravação definitiva.

---

## 👥 Público-Alvo

* **Escritórios de Contabilidade:** Que lidam com grandes volumes mensais de escrituração fiscal de seus clientes.
* **Departamentos Fiscais e Financeiros:** Equipes internas de empresas que precisam escriturar serviços tomados de forma ágil e segura.
* **Analistas Tributários:** Profissionais que buscam auditar e conferir a consistência dos dados de suas NFS-e de maneira centralizada.

---

## 📖 Como Funciona? (Visão Geral)

1. **Coleta de PDFs:** O usuário armazena os PDFs de NFS-e recebidos em uma pasta do computador.
2. **Geração do Compilado:** A ferramenta lê os arquivos e cria uma planilha consolidada com todas as informações fiscais estruturadas, gerando logs automáticos caso falte alguma informação em algum documento.
3. **Escrituração no Portal:** A partir da planilha gerada, o sistema abre o portal da prefeitura, realiza as buscas e preenche os dados do prestador e do serviço nota por nota, agilizando drasticamente o fluxo fiscal diário.

---

> [!NOTE]  
> Para obter detalhes de instalação, configuração de chaves de API, variáveis de ambiente e instruções técnicas de execução do projeto, consulte o [Manual do Desenvolvedor (DEVELOPER.md)](file:///c:/Users/gabriel.lima/Documents/GitHub/automacao-escrituracao-com-movimento/DEVELOPER.md).

