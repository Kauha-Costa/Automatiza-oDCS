<div align="center">

# 📜 AutoDCS
### Automação de Certificados

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat-square&logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=flat-square&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)

Atualiza datas e imprime certificados `.docx` em lote — com um clique.

</div>

---

## ✨ Funcionalidades

- 📁 **Processamento em lote** — atualiza todos os `.docx` de uma pasta de uma vez
- 📅 **Troca de data automática** — substitui a data em todos os certificados com formatação preservada
- 🖨️ **Impressão integrada** — envia para a impressora logo após atualizar
- 💾 **Configuração persistente** — pasta e impressora são lembradas entre sessões
- ⛔ **Cancelamento seguro** — interrompe o processo entre arquivos sem corromper nada
- 🎨 **Interface moderna** — tema escuro com log colorido em tempo real

---

## 🖥️ Interface

> Tema escuro com feedback visual em tempo real, barra de progresso e log colorido por tipo de evento.

---

## 🚀 Como usar

**1.** Abra o aplicativo  
**2.** Clique em **"Selecionar..."** e escolha a pasta com os arquivos `.docx`  
**3.** Digite a nova data no formato `dd/mm/aaaa`  
**4.** *(Opcional)* Marque **"Imprimir certificados após atualizar"** e selecione a impressora  
**5.** Clique em **▶ Atualizar Certificados**

---

## ⚙️ Instalação

### Pré-requisitos

- Python **3.10+**
- Windows (necessário para impressão via `win32print`)

### Rodando em desenvolvimento

```bash
# Clone o repositório
git clone https://github.com/seu-usuario/autodcs.git
cd autodcs

# Instale as dependências
pip install -r requirements.txt

# Execute
python main.py
```

---

## 📦 Gerando o executável `.exe`

```bash
pyinstaller --noconsole --onefile --icon=icone.ico --name Certificados main.py ^
  --hidden-import=docx ^
  --hidden-import=ttkbootstrap ^
  --hidden-import=win32api ^
  --hidden-import=win32print ^
  --hidden-import=pywinauto ^
  --add-data "icone.ico;."
```

O executável será gerado em `dist\Certificados.exe` — basta distribuir esse arquivo.

---

## 📁 Estrutura do projeto

```
autodcs/
├── main.py                  # Ponto de entrada
├── icone.ico
├── requirements.txt
└── utils/
    ├── editor.py            # Lógica de atualização dos .docx
    ├── impressao.py         # Lógica de impressão
    └── config.py            # Leitura/escrita de config.json
```

---

## 📝 Observações

- A configuração é salva em `config.json` na mesma pasta do executável
- A formatação original dos documentos (negrito, itálico, fonte, cor) é **totalmente preservada**
- Janelas de aviso de margem do Word são **fechadas automaticamente**
- O botão **Cancelar** interrompe o processo com segurança entre arquivos

---

<div align="center">
  Feito com ☕ e Python
</div>
