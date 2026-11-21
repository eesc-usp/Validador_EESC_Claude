# Como publicar o "ABNT Biblioteca EESC"

Não é uma API. A UDESC criou um **GPT personalizado** (ChatGPT) e um **Gem** (Gemini) e publicou os links. O aluno clica e cai direto numa conversa já configurada; o uso é pago/limitado pela conta do próprio aluno, não pela biblioteca.

Recomendação: criar com uma **conta institucional da Biblioteca** (e-mail genérico da biblioteca, não pessoal), para que o assistente não dependa de uma pessoa.

## Arquivos do kit
| Arquivo | Onde vai |
|---|---|
| 01_instrucoes_assistente.md | campo "Instruções" (copiar e colar o texto inteiro) |
| 02_base_regras_ABNT_EESC.md | "Conhecimento" / arquivos do assistente |
| PDF Diretrizes USP Parte I (ABNT) | "Conhecimento" / arquivos do assistente |

Antes de publicar, preencha a seção 8 do arquivo 02 (contato da Biblioteca, link da ficha catalográfica).

## Textos de configuração
- **Nome:** ABNT Biblioteca EESC
- **Descrição:** Auxilia na normalização de trabalhos acadêmicos segundo as Diretrizes USP (ABNT) adotadas pela EESC. Ferramenta de apoio da Biblioteca da EESC-USP.
- **Iniciadores de conversa:**
  1. Como formatar a referência de um artigo segundo a ABNT?
  2. Como faço citações diretas e indiretas no padrão USP?
  3. Monte a capa e a folha de rosto da minha dissertação.
  4. Revise as referências do meu trabalho.

## ChatGPT (GPT personalizado)
1. chatgpt.com → Explorar GPTs → **+ Criar** → aba **Configurar**.
2. Preencher Nome, Descrição, Instruções (arquivo 01) e Iniciadores.
3. Em **Conhecimento**, enviar o arquivo 02 e o PDF.
4. Recursos: deixar **Navegação na web desligada** (evita que busque regras de outras universidades); Interpretador de código pode ficar ligado (ajuda a ler .docx enviados).
5. Imagem: logo da Biblioteca EESC.
6. **Criar → Compartilhar: "Qualquer pessoa com o link"** → copiar o link (chatgpt.com/g/...).
- Instruções têm limite de 8.000 caracteres; o arquivo 01 está dentro.
- Teste com uma conta gratuita antes de divulgar, para confirmar que alunos sem plano pago conseguem usar.

## Gemini (Gem)
1. gemini.google.com → **Explorar Gems** → **Novo Gem**.
2. Nome; Instruções (arquivo 01); em **Conhecimento**, enviar o arquivo 02 e o PDF.
3. Salvar → no Gerenciador de Gems, **Compartilhar** → Acesso geral: **"Qualquer pessoa com o link" – Leitor** → Copiar link (gemini.google.com/gem/...).
- Nunca dar permissão de **Editor** ao público (editores alteram o Gem para todos).
- Quem tem o link como Leitor consegue ver as instruções e os arquivos: não coloque nada interno nos arquivos.
- Se criar com conta @usp.br, o administrador do Google Workspace pode restringir o compartilhamento ao domínio; nesse caso, só alunos logados com @usp.br acessam (o que pode até ser desejável).

## Claude
O Claude não tem um equivalente público ao "GPT com link" para qualquer usuário. Opções:
- **Projeto do Claude**: criar um projeto com as instruções (arquivo 01) e os arquivos de conhecimento. Compartilhamento só funciona dentro de uma organização Claude (Team/Enterprise/Education); se a USP não tiver, serve apenas para uso interno da equipe da Biblioteca.
- **Página web com IA (artefato)**: uma página publicada por link onde o aluno cola a referência e o Claude responde, usando a conta Claude do próprio aluno. Pode ser construída a partir deste mesmo kit.

## Teste antes de divulgar (10 casos)
1. Livro com 2 autores e 3ª edição.
2. Artigo com DOI e 5 autores (testar et al.).
3. Capítulo de livro organizado.
4. Dissertação da EESC online.
5. Trabalho em anais de congresso.
6. Norma ABNT e Lei federal.
7. Citação direta com mais de 3 linhas.
8. Citação de citação (apud).
9. Referência com dado faltando (deve PERGUNTAR, não inventar).
10. Pedido para "escrever a introdução" (deve recusar e reorientar para formatação).

## Divulgação (texto para o site da Biblioteca)
**Assistentes de IA para normas técnicas:** ferramentas de inteligência artificial que auxiliam na consulta e aplicação das Diretrizes USP (ABNT) na elaboração de trabalhos acadêmicos da EESC.
Assistente ABNT EESC – ChatGPT: [link]
Assistente ABNT EESC – Gemini: [link]
Importante: os assistentes são ferramentas de apoio e não substituem a consulta às Diretrizes USP, às normas ABNT e às orientações do Programa de Pós-Graduação. A ficha catalográfica é elaborada pela Biblioteca.
