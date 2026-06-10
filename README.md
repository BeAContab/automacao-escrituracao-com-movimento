# Automação de Escrituração com Movimento

Script em Python com duas funções:

1. Extrair dados de NFS-e em PDF e consolidar tudo em uma planilha XLSX.
2. Executar a automação visível do portal da ISS de Fortaleza no navegador.

## Uso

O ponto de entrada padrão do projeto é o arquivo `main.py` na raiz.

Ao executar o projeto sem parâmetros, ele apresenta um menu com as duas opções.

### Opção 1 - Gerar XLSX

Modo interativo:

```bash
python main.py
```

O script solicita o caminho da pasta com os PDFs e, ao final, oferece a opção de iniciar a função 2 ou encerrar o programa.

Modo com argumento:

```bash
python main.py --modo 1 "C:\caminho\para\pasta\com\pdfs" -o nf_compilado.xlsx
```

### Opção 2 - Automação visível da ISS

```bash
python main.py --modo 2 --competencia "5/2026" --planilha nf_compilado.xlsx
```

Modo de teste visível da função 2, com pausas extras para inspeção manual:

```bash
python main.py --modo 2 --competencia "5/2026" --planilha nf_compilado.xlsx --debug-funcao2
```

## Validação

Execute a suíte de referência antes de evoluir novas regras de extração:

```bash
python -B -m unittest discover -s tests
```

Os testes validam os PDFs atuais da pasta `dados/` e protegem campos fiscais críticos em layouts representativos.

## Dependências

- `PyPDF2`
- `PyMuPDF`
- `openpyxl`
- `playwright`
- `google-genai`

## Enriquecimento de CNAE final

- Ao gerar o XLSX na opção 1, o projeto agora adiciona as colunas `ID_CNAE_FINAL` e `DESC_CNAE_FINAL`.
- O enriquecimento usa o `DESC_CNAE` extraído do PDF como referência principal e cruza esse dado com o arquivo `cnae_oficial.xlsx`.
- Quando a variável `GROQ_API_KEY` estiver disponível no arquivo `.env`, a classificação final usa a Groq para escolher o melhor CNAE oficial.
- Se a chave não estiver configurada ou a consulta falhar, o projeto usa um fallback local com a melhor correspondência disponível.
- O modelo padrão configurável é `llama-3.1-8b-instant`, com foco em custo-benefício.

## Fallback de IA para extração

- O parser local continua sendo a primeira camada de extração.
- Quando um PDF vier com texto ruim, campos ausentes ou valores suspeitos, o projeto pode acionar o Gemini como fallback.
- O Gemini recebe o PDF anexado e responde em JSON estruturado, ajudando a completar campos como número da nota, data, local, valores e descrição do serviço.
- Se `GEMINI_API_KEY` não estiver configurada, o fluxo continua somente com a extração local.

## Variáveis de ambiente

- Copie o arquivo `.env.example` para `.env` e preencha `GROQ_API_KEY`.
- Se quiser trocar o modelo, ajuste `GROQ_MODEL` no mesmo arquivo.
- Para habilitar o fallback de extração com IA, preencha `GEMINI_API_KEY` e, se quiser, ajuste `GEMINI_MODEL`.
- O arquivo `.env` é ignorado pelo Git para evitar o versionamento de credenciais.

## Observações de fluxo

- A opção 1 preserva o comportamento atual de extração e geração do XLSX.
- Ao concluir a opção 1, o programa sugere que o usuário siga para a opção 2.
- A opção 2 abre um navegador visível, pede o mês de `1` a `12` e o ano em `AAAA`, aguarda o login manual e preenche a tela até a etapa anterior a `GRAVAR DOCUMENTO`.

## Observações técnicas

- O script tenta usar `PyMuPDF` como motor principal de leitura dos PDFs.
- O parser usa padrões reutilizáveis e fallbacks por contexto para ampliar a cobertura de layouts de NFS-e.
- O parser foi validado com os layouts presentes em `dados/` e `dados2/`.
- Quando alguma NF não preencher todas as colunas, o script gera um log `*_log_extracao.txt` ao lado do XLSX e não exporta a linha para a planilha.
- Se o PDF for uma imagem escaneada, a extração pode falhar sem OCR.
- Erros inesperados da execução também são registrados em `log_erro.txt` na raiz do projeto.
- A execução normal do programa também gera `log_execucao.txt` na raiz, com os marcos principais do fluxo.
## Automação da ISS

- A opção 2 agora abre um navegador visível e conduz o fluxo guiado do portal da ISS de Fortaleza.
- Antes de iniciar o navegador, o programa pergunta o mês de `1` a `12` e o ano em `AAAA`.
- Em seguida, o programa pergunta qual planilha XLSX será usada na automação, usando `nf_compilado.xlsx` se você apenas pressionar `Enter`.
- O login continua manual por enquanto, e o programa aguarda confirmação do usuário antes de seguir.
- O fluxo preenche a tela até a etapa anterior a `GRAVAR DOCUMENTO`, conforme a orientação atual do projeto.
- Você também pode informar outro arquivo com `--planilha` ao chamar a CLI.

## Extração avulsa do portal

- Quando for necessário extrair a tela logada da ISS uma única vez, use o utilitário `extrair_iss_uma_vez.py`.
- Esse utilitário abre o portal, aguarda o login manual, permite que você navegue até a tela desejada e exporta o conteúdo visível para `iss_extracao.xlsx`.
- A extração avulsa fica fora do menu principal para não interferir no fluxo normal da aplicação.
