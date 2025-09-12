# Automação BI - Instruções de Uso

## 1. Criar os Bookmarks no Power BI

Antes de rodar o sistema, acesse o relatório no Power BI e crie os bookmarks (marcadores) com os nomes desejados para exportação dos dados.  
Esses nomes serão usados pelo script para identificar e exportar cada conjunto de dados.

## 2. Alterar os Bookmarks no arquivo `.env`

Abra o arquivo `.env` com o Bloco de Notas ou outro editor de texto.  
No campo `BOOKMARKS`, coloque os nomes dos bookmarks separados por vírgula, exatamente como estão no Power BI.

**Exemplo:**
```
BOOKMARKS="Material Coordenador RT,Material Coordenador Geral RT,Material Coordenador Contas RT"
```

## 3. Alterar o Username

No arquivo `.env`, altere o campo `USERNAME` para o seu nome de usuário do Windows (aquele que aparece na pasta `C:\Users\`).

**Exemplo:**
```
USERNAME="seu_usuario"
```

## 4. Rodar o script `main.exe`

Dê um duplo clique no arquivo `main.exe` para iniciar o processo.

### 4.1 Fazer Login no Power BI

Quando o navegador abrir, faça login normalmente na sua conta do Power BI.

### 4.2 Esperar o sistema concluir a execução completa

Após o login, **não feche o navegador** **não mexa no computador até concluir a extração**.  
O sistema irá navegar, exportar os dados e processar os arquivos automaticamente.  
Aguarde até que o script finalize e a janela do navegador seja fechada.

---

**Observações:**
- Todos os arquivos gerados e tratados serão salvos na pasta `C:\Users\SeuUsuario\Downloads\RT`.
- Se precisar alterar os bookmarks ou o usuário, edite o arquivo `.env` antes de rodar novamente.
- Não feche o navegador manualmente durante a execução.

---