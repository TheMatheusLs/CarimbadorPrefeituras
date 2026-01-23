# Carimbador Automático - Prefeitura

Aplicação desktop desenvolvida em Python para automação de carimbos e numeração de páginas em documentos PDF. Focada no uso em prefeituras, permitindo a criação dinâmica de carimbos circulares e processamento em lote.

## 📋 Funcionalidades

- **Carimbar PDF Existente**:
  - Aplica o carimbo oficial e numeração sequencial em todas as páginas.
  - Detecta automaticamente orientação (Retrato/Paisagem) para posicionamento correto.
  - Opção para ignorar a capa (primeira página).
  
- **Gerar Folhas em Branco**:
  - Gera um novo PDF (A4) com quantidade definida de páginas, já carimbadas e numeradas.

- **Configuração e Personalização**:
  - **Gerador de Carimbos**: Criação automática da arte do carimbo (PNG) baseada no nome da cidade.
  - **Posicionamento**: Escolha entre 4 cantos padrão ou ajuste fino manual (X, Y).
  - **Persistência**: As configurações são salvas automaticamente em `config_v10.json`.

## 🛠️ Tecnologias Utilizadas

- **Interface Gráfica**: `tkinter` (Nativa)
- **Manipulação de PDF**: `pypdf`
- **Geração de PDF/Canvas**: `reportlab`
- **Processamento de Imagem**: `Pillow` (PIL)

## 🚀 Como Executar

### Pré-requisitos

Certifique-se de ter o Python instalado. Instale as dependências:

```bash
pip install pypdf reportlab Pillow
```

### Executando o Código Fonte

```bash
python main.py
```

## 📦 Compilação (Executável)

O projeto já possui um arquivo `.spec` configurado para uso com o **PyInstaller**. Para gerar o executável (`.exe`):

```bash
pyinstaller CarimbadorPrefeiturav2.spec
```
O executável será gerado na pasta `dist/`.

## 📂 Estrutura de Arquivos Importantes

- `main.py`: Código fonte principal.
- `config_v10.json`: Arquivo de configuração (gerado automaticamente).
- `Carimbos Prefeituras/`: Diretório onde as imagens dos carimbos gerados são armazenadas.

## 👤 Autor

Desenvolvido por **Matheus Lôbo**.
