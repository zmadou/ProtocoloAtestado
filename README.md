# Formulário → Planilha + PDF  

Sistema em Python que gera planilhas e documentos a partir de um formulário com interface gráfica.  
O projeto foi feito para agilizar processos acadêmicos e administrativos no **Senac**, evitando duplicidade de registros e garantindo padronização de protocolos.

---

## ✨ Funcionalidades

- **Interface gráfica (Tkinter)** com tema claro e barra de progresso.  
- **Configuração persistente** de sigla (MCI, MMD, IOT, etc.) e ano.  
- **Validações e máscaras**:  
  - CPF (11 dígitos, formatado no DOC)  
  - ID com tamanho fixo  
  - Código da oferta com tamanho fixo  
- **Geração automática** de planilha `MALA<SIGLA>.xlsx` com colunas padrão.  
- **Controle de protocolo**: sequência baseada apenas na planilha (`N req.`).  
- **Prevenção de duplicados**: não adiciona linha repetida se (Nome+ID+CPF+Data) já existir.  
- **Geração de documentos**:
  - Cria DOCX e tenta converter para PDF (usa `docx2pdf`, `comtypes` ou `win32com`).  
- **Personalização** via `config_form.json` ou variáveis de ambiente.  

---

## 📂 Estrutura de Saída

- **Planilha**:  
  - `MALASIGLA.xlsx` (ex.: `MALAMCI.xlsx`)  

- **Documentos**:  
  - Gerados em `Requerimentos/`  
  - Nomeados como:  
    ```
    01MCI2025 Nome do Aluno.docx
    01MCI2025 Nome do Aluno.pdf
    ```

---

## ⚙️ Requisitos

- Python 3.8+  
- Dependências (instalar via `pip install -r requirements.txt`):  
  - `pandas`  
  - `python-docx`  
  - `docx2pdf`  
  - `tkinter` (vem nativo com Python em Windows/Linux)  
  - `comtypes` (opcional, Windows)  
  - `pywin32` (opcional, Windows)  

---

## 🚀 Como usar

### 1. Clonar o repositório
```bash
git clone https://github.com/<seu-usuario>/<seu-repo>.git
cd <seu-repo>
```

### 2. Instalar dependências
```bash
pip install -r requirements.txt
```

### 3. Executar o formulário (modo gráfico recomendado)
```bash
python formulario.py --gui
```

### 4. Modo linha de comando (opcional)
```bash
python formulario.py --sigla MMD --ano 2025 --cpf-len 11 --id-len 10 --oferta-len 10
```

---

## 🔧 Configurações

- **Arquivo**: `config_form.json`  
- **Exemplo**:
```json
{
  "sigla": "MCI",
  "ano": "2025",
  "valid": {
    "cpf_len": 11,
    "id_len": 10,
    "oferta_len": 10
  },
  "last_req": 3
}
```

- **Variáveis de ambiente** (sobrepõem o config):  
  - `MCI_ANO` → define ano padrão  
  - `MCI_DOCX` → modelo DOCX base  
  - `MCI_XLSX` → força nome da planilha  
  - `MCI_CONFIG` → caminho alternativo do config  
  - `MCI_SAIDAS` → pasta de saída (default: `Requerimentos/`)  

---

## 📝 Colunas da Planilha

- N req.  
- N chamado  
- NOME  
- ID  
- CPF  
- CURSO  
- TURMA  
- Código da oferta  
- Data  
- retorno (Previsão)  

---

## 📌 Observações

- Funciona melhor no **Windows** (integração com Word para PDF).  
- No Linux/macOS, só gera o DOCX (PDF pode não funcionar se não houver Word instalado).  
- Se já existir registro idêntico, apenas regenera o DOCX/PDF sem duplicar a planilha.  

---

## 📄 Licença
Este projeto é de uso interno. Ajuste conforme necessário para o seu fluxo.
