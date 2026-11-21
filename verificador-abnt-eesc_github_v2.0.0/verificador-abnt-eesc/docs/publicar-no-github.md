# Publicar e atualizar o projeto no GitHub

## Jeito mais fácil (Windows, script)
1. Extraia o projeto numa pasta (ex.: `Documentos\verificador-abnt-eesc`).
2. Dê dois cliques em **`scripts\enviar_para_github.bat`**. Na primeira vez ele:
   - instala o **Git** e o **GitHub CLI** se faltarem (pelo winget; o Windows pode pedir confirmação);
   - abre o navegador para você entrar na conta do GitHub (digite o código que aparece na janela);
   - cria o repositório **`verificador-abnt-eesc` privado** na sua conta e envia tudo.
3. Nas próximas vezes, o mesmo arquivo só registra e envia as mudanças.

Opções (rodar pelo PowerShell ou criar um atalho com os parâmetros):
| Parâmetro | Efeito |
|---|---|
| `-Mensagem "texto"` | descrição da mudança |
| `-Versao` | publica a versão de `verificador-interno\VERSION`: o GitHub testa e gera os `.zip` em **Releases** |
| `-Dono nome-da-organizacao` | cria dentro de uma organização (ex.: conta institucional da biblioteca) |
| `-Repo outro-nome` | outro nome de repositório |
| `-Publico` | repositório público |

O script se recusa a enviar `config.json`, logs e qualquer `.docx` que não seja exemplo ou template — trabalhos
de alunos nunca vão para o GitHub.

## Pelo site do GitHub (só o navegador)
1. Baixe o `.zip` do projeto e **extraia**. O envio é da pasta extraída, não do `.zip`.
2. No repositório, clique em **Add file > Upload files**.
3. Abra a pasta extraída, selecione **tudo o que está dentro dela** (pastas e arquivos, inclusive `.github`) e
   arraste para a página. Use Chrome ou Edge: eles mantêm as subpastas.
4. O site aceita até 100 arquivos por envio. O projeto tem menos que isso, mas se passar, envie em partes (por
   exemplo, primeiro `verificador-interno`, depois o resto).
5. Escreva a descrição da mudança e clique em **Commit changes**.
6. O envio pelo site **não apaga** arquivos que saíram do projeto. Quando uma versão remover arquivos, o
   `CHANGELOG.md` avisa. Nesse caso, apague-os pelo site: abra o arquivo, clique em **⋯** e depois em **Delete file**.
7. Para publicar a versão, vá em **Releases > Draft a new release**, crie a tag `vX.Y.Z` (igual ao
   `verificador-interno/VERSION`, ex.: `v2.0.0`) e clique em **Publish release**. O GitHub roda os testes e anexa
   os `.zip` de instalação.

## Sem linha de comando (GitHub Desktop)
1. Instale o [GitHub Desktop](https://desktop.github.com) e entre na conta.
2. *File > Add local repository* → escolha a pasta do projeto → *create a repository* (se pedir).
3. *Publish repository* → deixe **Keep this code private** marcado.
4. Depois de cada mudança: escreva um resumo, *Commit to main* e *Push origin*.

## Público ou privado?
Comece **privado**. Antes de abrir ao público:
- confirme com a Biblioteca/ABCD-USP se o PDF das Diretrizes e o template da EESC podem ser redistribuídos
  (ou troque o PDF por um link);
- escolha uma licença (ex.: MIT para o código) e adicione o arquivo `LICENSE`.

## O que roda sozinho no GitHub
- **Testes** (`.github/workflows/testes.yml`): a cada envio, em Linux e Windows, rodam o lint (`ruff`), o pytest
  (regras, robustez, cache e retrato dos resultados), a abertura do app e a checagem de sintaxe e formato dos
  scripts do Windows. O resultado aparece na aba *Actions* e nos *pull requests*.
- **Nova versão** (`.github/workflows/versao.yml`): ao enviar uma etiqueta `v*` (opção `-Versao`), gera
  `verificador_abnt_eesc_vX.zip` e `assistente_alunos_vX.zip` em *Releases*.
