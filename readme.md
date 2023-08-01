**RESUMO:**

**Aplicação em Python utilizando as bibliotecas openpyxl e tkinter para criar uma GUI com funções de CRUD em uma base de dados armazenada em um arquivo Excel (.xlsx).**

<presentation src="/presentation/janela da aplicação.png">

**DETALHAMENTO:**

Esse projeto consiste em uma aplicação de manipulação (CRUD) de uma base de dados em Excel (arquivos .xlsx). 

Através da aplicação, é possível inserir, visualizar, editar e deletar itens do cadastro. Utilizei um cadastro fictício de produtos (roupas).

Na janela da aplicação, existem quatro campos de preenchimento, uma Spinbox para escolher o código do produto, um campo de texto para o nome do produto, uma caixa de lista para escolher qual o status do estoque e uma caixa de seleção para marcar caso o produto já esteja fora de linha (variáveis escolhidas são genéricas e podem ser substituídas por outras relevantes sem grande dificuldade).

Além disso, existe um visor com scroll bar na janela, onde é possível ver todos produtos cadastrados. Ao selecionar um produto, é possível editar seu cadastro ou deletá-lo com os botões presentes.

**MANUAL DE USO:**

- Execute o programa com o comando “python app.py” ou “python3 app.py” caso esteja utilizando um Mac (a janela abrirá automaticamente, caso o arquivo .xlsx não existe, uma janela de erro será exibida).

- Cadastrar um novo produto:
Na janela da aplicação, preencha corretamente os quatro Campos de informação (Código, Nome, Status e Disponibilidade). Em seguida aperte “Salvar produto”.

- Editar produto: 
Selecione um produto no visor e aperte a tecla “selecionar produto já cadastrado”. As informações serão carregadas nos campos de preenchimento. Altere a informação que deseja e clique no botão “Salvar alterações do produto”. Caso tudo ocorra corretamente, uma caixa de confirmação se abrirá.

- Deletar produto:
Selecione um produto no visor e clique no botão “Deletar produto”. Caso tudo ocorra corretamente, uma janela informativa abrirá, confirmando a exclusão do produto.
