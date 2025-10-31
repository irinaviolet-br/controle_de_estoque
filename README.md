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

> **Importante:** As planilhas de estoque são protegidas com a senha `789` para
> evitar edições diretas no Excel. O aplicativo aplica essa proteção
> automaticamente ao criar ou atualizar os arquivos.

## Funcionalidades

* Interface pública com registro de saídas e scanner de código de barras.
* Interface de gestor acessada com o PIN `2468`, permitindo entradas, saídas,
  cadastro de novos produtos e uso avançado do scanner.
* Integração com planilhas de estoque, respeitando a estrutura de colunas
  fornecida (cabeçalho na linha 6 e dados a partir da linha 7).
* Registro automático no arquivo `LOG.xlsx` de todas as movimentações.

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
