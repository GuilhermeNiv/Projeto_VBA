# 📊 Projeto VBA - Dashboard de Cotações Automatizado

Este projeto foi desenvolvido em VBA (Visual Basic for Applications) para Excel, com foco na automação de coleta de cotações de moedas via web. Ele extrai os dados do InfoMoney, e atualiza automaticamente uma planilha com:

- 💵 **Dólar**
- 💶 **Euro**
- 🪙 **Bitcoin**

Meu objetivo com este projeto é destacar minha habilidade de montar uma automação que busca dados em um site WEB e coloca-os em uma planilha em Excel para rápida visualização.

## 📌 Funcionalidades

✔️ Cria uma nova aba chamada `Cotação`  
✔️ Acessa páginas web e importa os dados
✔️ Exibe data e hora da última atualização  
✔️ Gera gráficos automáticos 
✔️ Cria botões interativos no Excel para:

- 🔁 Atualizar tudo
- 💵 Atualizar Dólar
- 💶 Atualizar Euro
- 🪙 Atualizar Bitcoin

---

## 🛠️ Como usar

1. Abra o Excel
2. Pressione `ALT + F11` para abrir o Editor VBA
3. Crie um novo módulo e cole o conteúdo do arquivo `automacao.vb`
4. Execute a macro `ConfigurarDashboard`

---

## 📁 Estrutura

Projeto_VBA/
├── automacao.vb ← Código da automação em VBA
└── README.md

---

## 💡 Dicas

- Se quiser adicionar mais moedas, basta duplicar a função `ImportarCotacao` com novos títulos e URLs.
- Os botões são criados dinamicamente, então é possível personalizá-los facilmente.

---

## 📅 Atualização

O script também insere a **data e hora** da última atualização diretamente na planilha.

---

👨‍💻 Autor:
Guilherme Laureano | 
Disponível para contratação | Uberlândia - MG
