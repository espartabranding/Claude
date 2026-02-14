# Handwriting Recognition to Excel

Ferramenta de alta capacidade para reconhecimento de escrita manual (mesmo desorganizada/ruim) em imagens JPG, com exportação para planilha Excel.

## Funcionalidades

- **Pré-processamento avançado de imagem**: correção de rotação, remoção de ruído, binarização adaptativa, normalização de contraste
- **Duplo motor OCR**: Tesseract + EasyOCR trabalhando juntos para máxima precisão
- **Detecção automática de estrutura**: identifica se o conteúdo é tabular ou linear
- **Agrupamento inteligente**: organiza texto desorganizado em linhas e colunas
- **Exportação Excel formatada**: com cabeçalhos, bordas, filtros e largura automática de colunas
- **Processamento em lote**: processa múltiplas imagens de uma vez
- **Suporte multilíngue**: português, inglês, espanhol, francês, alemão

## Instalação

### 1. Instalar Tesseract OCR (sistema)

```bash
# Ubuntu/Debian
sudo apt-get install tesseract-ocr tesseract-ocr-por tesseract-ocr-eng

# macOS
brew install tesseract tesseract-lang

# Windows
# Baixe o instalador em: https://github.com/UB-Mannheim/tesseract/wiki
```

### 2. Instalar dependências Python

```bash
pip install -r requirements.txt
```

## Uso

### Comando básico

```bash
python handwriting_to_excel.py foto.jpg
```

Isso gera `foto.xlsx` no mesmo diretório.

### Opções completas

```bash
# Especificar arquivo de saída
python handwriting_to_excel.py foto.jpg -o resultado.xlsx

# Usar apenas EasyOCR (não precisa instalar Tesseract)
python handwriting_to_excel.py foto.jpg --engine easyocr

# Modo agressivo para escrita muito ruim
python handwriting_to_excel.py foto.jpg --enhance aggressive

# Forçar modo tabela
python handwriting_to_excel.py foto.jpg --mode table

# Forçar modo linhas (sem colunas)
python handwriting_to_excel.py foto.jpg --mode lines

# Primeira linha NÃO é cabeçalho
python handwriting_to_excel.py foto.jpg --no-header

# Salvar imagens de debug do pré-processamento
python handwriting_to_excel.py foto.jpg --debug

# Idiomas específicos
python handwriting_to_excel.py foto.jpg --lang pt en es
```

### Processamento em lote

```bash
# Processar pasta inteira (um Excel por imagem)
python batch_process.py pasta_de_imagens/ -o resultados/

# Combinar tudo em um único Excel (uma aba por imagem)
python batch_process.py pasta_de_imagens/ --single-file tudo.xlsx
```

## Parâmetros

| Parâmetro | Opções | Padrão | Descrição |
|-----------|--------|--------|-----------|
| `--engine` | `tesseract`, `easyocr`, `both` | `both` | Motor OCR |
| `--lang` | códigos de idioma | `pt en` | Idiomas para reconhecimento |
| `--enhance` | `standard`, `aggressive` | `standard` | Nível de enhancement |
| `--mode` | `auto`, `table`, `lines` | `auto` | Modo de detecção |
| `--no-header` | flag | - | Não tratar 1ª linha como cabeçalho |
| `--debug` | flag | - | Salvar imagens preprocessadas |
| `-v` | flag | - | Saída detalhada |

## Pipeline de Processamento

```
Imagem JPG
    │
    ▼
┌─────────────────────────────────┐
│  1. Pré-processamento           │
│  ├── Normalização de tamanho    │
│  ├── Conversão para escala de   │
│  │   cinza                      │
│  ├── Correção de rotação        │
│  ├── Remoção de ruído           │
│  ├── Melhoria de contraste      │
│  ├── Binarização adaptativa     │
│  └── Limpeza morfológica        │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  2. OCR (múltiplas versões)     │
│  ├── Tesseract OCR              │
│  ├── EasyOCR                    │
│  └── Merge e deduplicação       │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  3. Detecção de Estrutura       │
│  ├── Agrupamento em linhas      │
│  ├── Detecção de colunas        │
│  └── Montagem da tabela         │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  4. Exportação Excel            │
│  ├── Formatação de cabeçalho    │
│  ├── Auto-ajuste de colunas     │
│  ├── Filtros automáticos        │
│  └── Congelamento de painéis    │
└─────────────────────────────────┘
    │
    ▼
  arquivo.xlsx
```

## Dicas para Melhores Resultados

1. **Iluminação**: fotografe com boa iluminação e sem sombras
2. **Ângulo**: tente manter a câmera perpendicular ao papel
3. **Resolução**: use pelo menos 300 DPI ou fotos de alta resolução
4. **Contraste**: caneta escura em papel claro funciona melhor
5. **Modo agressivo**: use `--enhance aggressive` para escrita muito ruim
6. **Debug**: use `--debug` para ver como a imagem é pré-processada
