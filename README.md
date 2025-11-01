# Controle de Estoque Labpat

Aplicação desktop em Python/Tkinter para registrar movimentações de estoque do Labpat
em planilhas do Excel.

## Pré-requisitos

* Python 3.10 ou superior
* Dependências listadas em `requirements.txt`

Dependências opcionais para gerar o executável Windows estão listadas em
`requirements-dev.txt`.

Instale as dependências com:

```bash
pip install -r requirements.txt
```

## Executando

Na pasta raiz do projeto execute:

```bash
python main.py
```

Crie uma pasta chamada `planilhas/` (já incluída vazia no repositório) e coloque
ali as planilhas reais do laboratório (`ESTOQUE*.xlsx` e `LOG.xlsx`). Caso
esteja configurando o sistema pela primeira vez, você pode deixar a pasta vazia:
ao executar o aplicativo, os arquivos serão gerados automaticamente com o
formato correto.

Qualquer arquivo `.xlsx` cujo nome comece com `ESTOQUE` será exibido nas
comboboxes do sistema, permitindo trabalhar com estoques adicionais apenas
copiando-os para a pasta `planilhas/`.

Coloque também a imagem `LOGO.png` na pasta raiz do projeto (ao lado de
`main.py`). Ela será exibida automaticamente no topo das interfaces pública e
de gestor.

Se desejar personalizar o ícone do executável, salve a imagem `LOGO Estoque.png`
na mesma pasta.

> **Importante:** As planilhas de estoque são protegidas com senha para evitar
> edições diretas no Excel. O aplicativo aplica a senha configurada em
> `settings.json` (valor padrão `789`) sempre que atualiza os arquivos. Use o
> botão **"Senha das planilhas"** na interface do gestor para definir um novo
> valor quando desejar.

Ao lado de `main.py`, o aplicativo cria (ou atualiza) o arquivo
`settings.json`, onde ficam armazenados:

* **`manager_pin`** – PIN atual do gestor (apenas números, até 10 dígitos).
* **`manager_email`** – e-mail informado ao redefinir a senha.
* **`stock_password`** – senha utilizada para proteger as planilhas.
* **`email`** – parâmetros SMTP (`host`, `port`, `username`, `password`,
  `use_tls`, `use_ssl`, `from_address`) para envio automático da senha.

Configure os parâmetros SMTP conforme o serviço de e-mail corporativo. O
programa envia a nova senha para o e-mail cadastrado no momento da definição da
senha.

## Funcionalidades

* Interface pública com registro de saídas, scanner de código de barras e
  opção de recuperação de senha do gestor.
* Interface de gestor acessada com o PIN configurado, permitindo entradas,
  saídas, cadastro de novos produtos, redefinição do PIN e atualização da senha
  das planilhas.
* Integração com planilhas de estoque, respeitando a estrutura de colunas
  fornecida (cabeçalho na linha 6 e dados a partir da linha 7).
* Registro automático no arquivo `LOG.xlsx` de todas as movimentações.
* Suporte opcional a arquivos do Excel protegidos por senha utilizando a
  biblioteca `msoffcrypto-tool`.

## Gerando o executável "Controle de Estoque"

1. **Execute os comandos em um ambiente Windows.** O PyInstaller gera o `.exe`
   apenas no próprio Windows.
2. Instale as dependências de build (no ambiente virtual, se estiver usando):

   ```bash
   pip install -r requirements-dev.txt
   ```

3. Certifique-se de que os arquivos `LOGO.png` e `LOGO Estoque.png` estejam na
   pasta raiz do projeto, ao lado de `main.py`.
4. Gere o executável:

   ```bash
   python build_exe.py
   ```

   > O script reconhece o PyInstaller mesmo que o comando `pyinstaller` não
   > esteja disponível no `PATH`; ter o pacote instalado no ambiente Python já é
   > suficiente.

Ao final do processo, o arquivo `Controle de Estoque.exe` ficará disponível em
`dist/`. A pasta `dist/` pode ser distribuída aos usuários finais.
