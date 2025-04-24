# 🧾 Excel Formatter

Um projeto em Python que padroniza automaticamente arquivos Excel de acordo com o padrão definido pela empresa. Ideal para automatizar o tratamento de arquivos brutos recebidos em `.csv`, convertendo-os para `.xlsx`, aplicando formatações, e gerenciando tudo por meio de arquivos temporários e variáveis de ambiente.

---

## 🚀 Funcionalidades

- 🗃️ Conversão automática de `.csv` (delimitado por `;`) para `.xlsx`
- 🎨 Aplicação de formatação padrão definida pela empresa
- 📁 Organização e movimentação de arquivos entre pastas
- 🧪 Utilização de arquivos temporários para segurança no processamento
- 🔐 Suporte a `.env` para configuração de caminhos e variáveis sensíveis
- 📦 Gerenciado com [**uv**](https://github.com/astral-sh/uv) (ambiente virtual leve e rápido)

---

## ⚙️ Requisitos

- Python 3.11
- [uv](https://github.com/astral-sh/uv) instalado globalmente

---

## 📦 Instalação

```bash
# Clone o repositório
git clone https://github.com/AangeloGabriel/FormatadorDeExel.git
cd FormatadorDeExcel

# Crie o ambiente com uv
uv venv

# Ative o ambiente
# Linux/macOS:
source .venv/bin/activate
# Windows:
.venv\Scripts\activate

# Instale as dependências
pip install .