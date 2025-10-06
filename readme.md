# Automação BI - Instruções de Uso

## 1. Criar os Bookmarks no Power BI

Antes de rodar o sistema, acesse o relatório no Power BI e crie os bookmarks (marcadores) com os nomes desejados para exportação dos dados.  
Esses nomes serão usados pelo script para identificar e exportar cada conjunto de dados.

## 2. Alterar os Bookmarks no arquivo `.env`

Abra o arquivo `.env` com o Bloco de Notas ou outro editor de texto.  
No campo `BOOKMARKS`, coloque os nomes dos bookmarks separados por vírgula, exatamente como estão no Power BI.
Para funcionar corretamente, é necessário colocar na seguinte ordem:

1. A extração que mostra a lista total
2. A extração que mostra a linha total geral
3. A extração completa de todos os dados de cada representante, com os seus respectivos clientes

**Exemplo:**
```
BOOKMARKS="Material Coordenador Contas RT, Material Coordenador Geral RT, Materia Coordenador RT"
```
No campo `BOOKMARKS_META`, coloque os arquivos que voce deseja para que o sistema gere as metas
**Exemplo:**
```
BOOKMARKS_META="Meta Coordenador RT sem Grupo Econ.,Meta Coordenador RT"
```

## 3. Alterar o Username

No arquivo `.env`, altere o campo `USERNAME` para o seu nome de usuário do Windows (aquele que aparece na pasta `C:\Users\`).

**Exemplo:**
```
USERNAME="seu_usuario"
```


## 4. Alterar o codigo COLUNA_NOME nome da coluna do responsável que deseja buscar por exemplo `CONTAS REDE` COLUNA_GRUPO é o nome da coluna do cliente que precisa buscar por exemplo `GRUPO ECONOMICO`

## 5. Rodar o script `exe` desejado

Dê um duplo clique no arquivo `exe` que deseja fazer o tratamento para iniciar o processo.

<!-- ### 4.1 Fazer Login no Power BI

Quando o navegador abrir, faça login normalmente na sua conta do Power BI.

### 4.2 Esperar o sistema concluir a execução completa

Após o login, **não feche o navegador** **não mexa no computador até concluir a extração**.  
O sistema irá navegar, exportar os dados e processar os arquivos automaticamente.  
Aguarde até que o script finalize e a janela do navegador seja fechada.

--- -->

**Observações:**
- Todos os arquivos gerados e tratados serão salvos na pasta `C:\Users\SeuUsuario\Downloads\RT`.
- Se precisar alterar os bookmarks ou o usuário, edite o arquivo `.env` antes de rodar novamente.
<!-- - Não feche o navegador manualmente durante a execução. -->

---