# 🧾 Gerador de Relatórios Word com Base em Planilhas Excel

Este é um projeto em **Python** que automatiza a criação de relatórios personalizados em formato **.docx (Word)**, utilizando dados extraídos de uma planilha **.xlsx (Excel)**. A aplicação conta com uma **interface gráfica simples**, feita com `tkinter`, que facilita a seleção dos arquivos e a geração dos documentos.

## 🚀 Funcionalidades

- 📂 Seleção do arquivo Excel com os dados
- 📄 Seleção do modelo Word com placeholders (ex: `{{nome_cliente}}`)
- 🔁 Substituição automática de placeholders nos parágrafos e tabelas do Word
- 💾 Geração de um arquivo Word para cada linha da planilha
- 🖥️ Interface gráfica intuitiva e fácil de usar
- 📁 Salvamento automático dos relatórios em uma pasta definida pelo usuário

## 🛠️ Tecnologias Utilizadas

- [Python](https://www.python.org/)
- [pandas](https://pandas.pydata.org/)
- [python-docx](https://python-docx.readthedocs.io/)
- [tkinter](https://docs.python.org/3/library/tkinter.html)
- [os](https://docs.python.org/3/library/os.html)

## 📦 Instalação

1. Clone o repositório:
   ```bash
   cd gerador-relatorios-word
   ```

2. Instale as dependências:
   ```bash
   pip install pandas python-docx
   ```

3. Execute o script:
   ```bash
   python gerador_relatorios.py
   ```

## 📝 Como Usar

1. Prepare sua planilha Excel com colunas como `nome_cliente`, `data`, `endereco`, etc.
2. Crie um modelo Word (.docx) com os placeholders correspondentes, por exemplo:
   ```
   Prezado {{nome_cliente}},
   Agradecemos sua visita em {{data}}.
   ```
3. Execute o programa e:
   - Selecione o arquivo Excel
   - Selecione o modelo Word
   - Escolha a pasta de saída
4. O sistema gerará automaticamente um arquivo Word para cada linha da planilha, com os campos personalizados.

## ✅ Exemplo de Placeholder

| Coluna no Excel     | Placeholder no Word     |
|---------------------|--------------------------|
| nome_cliente        | `{{nome_cliente}}`       |
| data                | `{{data}}`               |
| endereco            | `{{endereco}}`           |

## 📌 Observações

- Os placeholders devem estar entre `{{` e `}}` e devem corresponder exatamente aos nomes das colunas da planilha.
- O nome dos arquivos gerados será baseado no campo `nome_cliente`, com espaços substituídos por `_`.

## 📄 Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

---

Desenvolvido com 💻 por [Alisson Dev](https://github.com/alissondev-btu)
