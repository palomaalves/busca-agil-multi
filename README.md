# 🔎 Busca Ágil em PDF — Multi-arquivos v2

Ferramenta web para buscar nomes ou termos em **um ou vários PDFs simultaneamente**, desenvolvida para a **Coordenadoria de Governança de Dados do TJPE**.

## ✨ Funcionalidades

- Upload de múltiplos PDFs de uma vez
- Busca ignorando acentuação e maiúsculas/minúsculas
- Destaque visual dos termos encontrados
- Estatísticas por arquivo e por termo
- **Exportação dos resultados em Excel**
- Processamento paralelo para maior velocidade
- Configurações salvas automaticamente no navegador

## 🚀 Como usar

### Localmente

```bash
pip install -r requirements.txt
python busca_agil_v2.py
```

Acesse `http://127.0.0.1:7860` no navegador.

### Hugging Face Spaces

Acesse diretamente: [Palu1006/busca-agil-v2](https://huggingface.co/spaces/Palu1006/busca-agil-v2)

## 📁 Estrutura

| Arquivo | Descrição |
|---|---|
| `busca_agil_v2.py` | App principal (v2 — self-contained, com Excel) |
| `app.py` | Versão anterior (multi-PDF, sem Excel) |
| `search_module.py` | Módulo de busca usado pelo `app.py` |
| `requirements.txt` | Dependências do projeto |

## 🛠️ Dependências

- [Gradio](https://gradio.app/) — interface web
- [pdfplumber](https://github.com/jsvine/pdfplumber) — extração de texto de PDF
- [openpyxl](https://openpyxl.readthedocs.io/) — exportação Excel

## 👩‍💻 Autora

Paloma Alves — Coordenadoria de Governança de Dados / TJPE
