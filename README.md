# Exportar Texto Longo dos Materiais no SAP

Este script automatiza o acesso ao SAP (transações **MMBE** e **MM03**) para extrair o **texto longo** dos materiais (campo "Texto do Pedido de Compras").  
Os códigos dos materiais são fornecidos por meio de um arquivo Excel e os textos extraídos são salvos nesse mesmo arquivo.

---

## 📋 Pré-requisitos

- Acesso ao SAP via **SAP Logon** (com login automático configurado — "login sem senha");
- Sistema operacional **Windows**;
- Microsoft Excel instalado;
- SAP GUI instalado e funcional.

---

## 🚀 Ordem de Execução

Execute os arquivos `.bat` na seguinte ordem:

### ✅ **Passo 00 – Instalar o Python**
Execute este passo apenas se você ainda **não tiver o Python instalado** em seu computador.  
O script abrirá a página oficial de download do Python para que você possa instalá-lo.

---

### ✅ **Passo 01 – Instalar as bibliotecas**
Executa a instalação automática das dependências necessárias via `pip`.

---

### ✅ **Passo 02 – Configurar o script**
Este passo abre uma interface gráfica (GUI) onde você deve configurar os 3 parâmetros principais:

- **Timeout (s):**  
  Define o tempo de espera entre as interações automatizadas com o SAP, simulando o tempo de resposta de um usuário real.  
  ⚠️ Valores mais baixos aceleram a execução, mas aumentam as chances de erro.

- **SAP Logon Path:**  
  Caminho completo até o executável do **SAP Logon** (`saplogon.exe`).  
  🔍 Dica: pesquise por `saplogon.exe` no menu Iniciar, clique com o botão direito, vá em "Abrir local do arquivo", e copie o caminho do navegador de arquivos manualmente (evite colar diretamente com `\`, pois o script pode não reconhecer corretamente).

- **Excel File Path:**  
  Caminho para o arquivo Excel `.xlsx` que contém a lista de materiais.  
  ⚠️ O arquivo deve conter uma coluna chamada exatamente **material_code**. Essa coluna será usada para buscar os materiais no SAP.

---

### ✅ **Passo 03 – Buscar textos longos no SAP**
Execute este passo para iniciar a automação.  
O script abrirá o SAP, acessará as transações necessárias e preencherá o Excel com os textos extraídos.

📌 **Importante:** Não utilize o computador durante a execução do script. O uso do teclado/mouse pode interferir no processo automatizado.

