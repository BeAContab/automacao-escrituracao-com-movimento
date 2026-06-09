# PRD - Extração de NFS-e em PDF para XLSX

## 1. Objetivo
Automatizar a leitura de uma pasta contendo arquivos PDF de NFS-e e consolidar os dados fiscais relevantes em uma planilha XLSX.

## 2. Problema
A conferência manual de notas fiscais em PDF é lenta, sujeita a erro humano e dificulta a padronização dos dados para escrituração.

## 3. Solução Proposta
Criar um script Python que:
- Receba uma pasta com PDFs.
- Extraia os campos fiscais relevantes de cada documento.
- Gere um arquivo XLSX com uma linha por nota fiscal.

## 4. Campos de Saída
O XLSX deve conter:
- ARQUIVO_PDF
- CNPJ do prestador
- Número da NF
- Data de emissão / data fato gerador
- ID_CNAE
- DESC_CNAE
- Descrição do serviço
- UF do local de prestação
- Cidade do local de prestação
- Natureza da operação
- ISS Retido
- Valor do serviço
- Alíquota

## 5. Regras de Negócio
- Se a cidade do local de prestação for `Fortaleza`, a natureza da operação deve ser `Tributação no Município`.
- Caso contrário, a natureza da operação deve ser `Tributação Fora do Município`.
- O campo `ISS Retido` deve ser `SIM` quando `valor_líquido = valor_serviço - ISS`.
- Caso contrário, deve ser `NÃO`.
- O bloco `SERVIÇO NACIONAL` deve ser tratado como identificador do CNAE para fins de extração.

## 6. Premissas
- Os PDFs são textuais e permitem extração via biblioteca de leitura de PDF.
- O layout das notas segue um padrão semelhante ao documento de referência analisado.

## 7. Fora de Escopo
- OCR de PDFs escaneados como imagem.
- Integração com banco de dados.
- Classificação tributária avançada além das regras solicitadas.

## 8. Critérios de Aceite
- O script processa todos os PDFs de uma pasta.
- O XLSX é gerado sem intervenção manual.
- Os campos pedidos são preenchidos quando presentes no documento.
- O comportamento de `Natureza da operação` e `ISS Retido` segue as regras definidas.

## 9. Riscos
- Diferenças de layout entre prefeituras podem exigir ajustes de regex.
- PDFs sem texto extraível podem exigir OCR em uma etapa futura.

## 10. Observações
- O arquivo de saída inclui a coluna `ARQUIVO_PDF` para rastreabilidade do documento de origem.
