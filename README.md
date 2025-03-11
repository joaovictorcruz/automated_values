# 📊 Atualização de Valores em Planilhas Excel

Este projeto tem como objetivo atualizar automaticamente os valores de uma planilha bancária (`PrimeiroArquivo.xlsx`) com base em um relatório externo (`SegundoArquivo.xlsx`).

O script **mantém a formatação original** da planilha ao inserir os valores corretos na coluna apropriada.

---

## ✨ Funcionalidades

✅ **Atualiza os valores automaticamente** na planilha original.  
✅ **Mantém a formatação** e estrutura da planilha bancária.  
✅ **Identifica a coluna correta para inserção dos valores** dinamicamente.  
✅ **Remove espaços em branco** dos nomes para evitar erros de correspondência.  
✅ **Cria uma cópia atualizada** sem alterar os arquivos originais.  

---

## 🛠️ Tecnologias Utilizadas

- **Python**: Linguagem principal para manipulação de dados.  
- **Pandas**: Biblioteca usada para leitura e manipulação dos dados do Excel.  
- **OpenPyXL**: Utilizada para manter a formatação da planilha original.  

---

## 📂 Estrutura do Projeto

```
📁 arquivos_excel/            # Contém as planilhas originais
   ├── PrimeiroArquivo.xlsx   # Planilha bancária a ser atualizada
   ├── SegundoArquivo.xlsx    # Relatório com os valores corretos

📁 arquivos_atualizados/      # Contém as planilhas atualizadas
   ├── ArquivoAtualizado.xlsx # Arquivo final com os valores corrigidos

📜 atualizar_valores_planilha.py  # Script principal para atualização dos valores
📜 verify.ipynb                   # Notebook opcional para conferência dos valores
```

---

## 📌 **Configuração e Uso**

### 🔹 1. Instalar as Dependências

Certifique-se de ter o Python instalado. Depois, instale as bibliotecas necessárias:
```sh
pip install pandas openpyxl
```

---

### 🔹 2. Adaptar os Arquivos

Para que o script funcione corretamente, **certifique-se de ajustar os nomes dos arquivos e colunas** conforme sua necessidade:

- O **PrimeiroArquivo.xlsx** (planilha bancária) deve conter a coluna **"Titular da Conta / Account Name"** e **"Valor / Amount (BRL)"**.
- O **SegundoArquivo.xlsx** (relatório) deve conter uma coluna com o nome dos titulares (**"Nome"**) e uma coluna com os valores corretos (**"Valor"**).

Se os nomes das colunas forem diferentes, altere as referências no código.

---

### 🔹 3. Executar o Script

Dentro da pasta do projeto, execute o seguinte comando:
```sh
python atualizar_valores_planilha.py
```

---

### 🔹 4. Verificar os Resultados

Após a execução, um arquivo atualizado será salvo em `arquivos_atualizados/ArquivoAtualizado.xlsx` com os valores corretos inseridos.

Para verificar se os valores estão corretos, você pode abrir o **verify.ipynb** e conferir os dados.

---

## 📢 **Observações**

- Se os valores estiverem sendo inseridos na coluna errada, **verifique os nomes das colunas** no seu Excel.
- Caso dê erro ao abrir o arquivo salvo, **verifique se o arquivo original não está aberto em outro programa**.
- Se o script não encontrar os valores, pode ser que existam **espaços invisíveis** nos nomes. O script já tenta corrigir isso automaticamente.

---

🚀 Feito para facilitar a atualização de dados em planilhas sem perder a formatação!  
📌 Se precisar de ajustes, modifique os nomes dos arquivos e colunas conforme sua necessidade!  

**Desenvolvido por [Seu Nome](https://github.com/seu-usuario)**

