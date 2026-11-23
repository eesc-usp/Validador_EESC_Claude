# Verificador ABNT – Biblioteca EESC (uso interno)

> Aplicativo de desktop: abre numa janela própria pelo atalho **Verificador ABNT EESC** (área de trabalho e
> menu Iniciar). Arquivos salvos vão para **Documentos\Verificador ABNT**.

App local para a equipe da Biblioteca analisar os trabalhos enviados pelos alunos (.docx ou PDF) segundo as
*Diretrizes para apresentação de dissertações e teses da USP: parte I (ABNT)*, 5. ed., 2024.
O arquivo é processado no próprio computador: nada vai para a internet.

**O verificador só aponta, não corrige.** Ele mostra o trabalho página por página, como um leitor de PDF, com
cada problema marcado, e gera o arquivo do aluno com comentários do Word. O texto do aluno nunca é alterado:
quem faz as correções é o próprio aluno.

## Como funciona

| Etapa | O que faz | Tempo |
|---|---|---|
| 1. Regras fixas (sempre) | Margens, papel, fonte e entrelinhas do texto, citações longas, paginação, alíneas, elementos obrigatórios, resumo/abstract e palavras-chave, numeração e títulos das seções, legendas e fontes de figuras e tabelas, bordas de tabelas, citação × referência, formato e ordem das referências | menos de 1 s (TCC de 80 páginas) |
| 2. IA local (opcional) | Revisa cada referência contra o modelo do tipo dela (livro, artigo, capítulo, evento, tese, norma, lei...) e sugere a correção | 15–60 s por referência sem placa de vídeo |
| 3. Arquivo com comentários | O arquivo do aluno como foi enviado, com um comentário em cada ponto (comentário do Word no .docx; destaque com nota no PDF) | na hora |
| 4. Relatório | Gera um .docx (item, situação, o que foi encontrado, como corrigir, seção do manual) | na hora |

Situações: 🔴 **Ajustar** (erro claro) · 🟡 **Conferir** (o programa não tem certeza: o bibliotecário confere) · 🟢 **OK**.

O bibliotecário sempre revisa o relatório antes de enviar ao aluno. As sugestões da IA são marcadas como "conferir".

## Instalação no PC da biblioteca (Windows, sem administrador, sem Docker, sem reiniciar)

1. Copie esta pasta para o computador (ex.: `C:\Users\<usuario>\verificador_abnt_eesc`).
2. Dê duplo clique em **`instalar.bat`**. Ele instala sozinho, só para o usuário atual:
   - Python 3.12, se não houver Python 3.10 ou mais novo;
   - as bibliotecas do app, num ambiente próprio (`.venv`);
   - o **Ollama** (IA local) e o modelo de IA escolhido pela memória do PC:
     16 GB de RAM ou mais → `qwen2.5:7b`; menos que isso → `qwen2.5:3b`;
   - um atalho **Verificador ABNT EESC** na área de trabalho.
   Na primeira vez baixa alguns GB (Ollama + modelo) e pode levar de 10 a 30 minutos.
3. Depois, use o atalho da área de trabalho (ou `iniciar.bat`).

### Atualizar para uma nova versão

Baixe o `verificador_abnt_eesc_vX.zip` da versão nova em **Releases** do repositório no GitHub, extraia por
cima da pasta (substituindo os arquivos) e abra o atalho normalmente. O `iniciar.bat` percebe que
a versão mudou e roda a verificação completa sozinho:

| Situação encontrada | O que ele faz |
|---|---|
| Componente faltando (Python, bibliotecas, Ollama, modelo, atalho) | Instala |
| Biblioteca em versão errada ou quebrada | Atualiza ou reinstala |
| Ambiente `.venv` danificado | Apaga e recria |
| Ollama mais antigo que o mínimo, ou danificado | Reinstala |
| Ollama parado | Inicia |
| Modelo corrompido | Apaga e baixa de novo |
| Atalho apontando para outra pasta | Corrige |
| Arquivo do app faltando | Avisa para copiar a pasta inteira de novo |

Para forçar a verificação completa a qualquer momento, rode `instalar.bat`. O registro fica em `instalar.log`.

Opções (no Prompt de Comando, dentro da pasta):
- `instalar.bat -SemIA` → instala só a verificação por regras, sem Ollama;
- `instalar.bat -Modelo qwen2.5:3b` → troca o modelo de IA.

### Observações sobre o PC sempre ligado

- O app e o Ollama rodam na sessão do usuário que abriu. A troca de usuário não os fecha.
- O app fica numa janela minimizada do Python: se ela for fechada, basta abrir o atalho de novo.
- Na barra lateral do app, **Núcleos da CPU para a IA** limita o processamento (o padrão deixa 2 núcleos livres
  para o BEESC BOT e o uso normal).
- Para acessar de outros PCs da biblioteca, troque `address` para `"0.0.0.0"` em `.streamlit/config.toml`
  e use `http://<IP-do-PC>:8501` (o firewall pode pedir liberação de administrador).

## Servidor da biblioteca (todos acessam pela rede, sem instalar)

Um computador (o da biblioteca, que fica sempre ligado) roda o app e os outros usam pelo **navegador**.
Atualizou nesse computador, vale para todos.

1. No computador servidor, dê duplo clique em **`servidor.bat`**. Ele instala/repara o que faltar e:
   - deixa o app rodando para a rede, sem janela, e **reinicia sozinho** se cair;
   - configura para **iniciar junto com o Windows** (pasta Inicializar do usuário);
   - libera a porta 8501 no **firewall** (o Windows pede permissão de administrador uma vez; se não houver,
     peça ao TI: liberar entrada TCP 8501, redes de domínio/privada);
   - mostra os endereços (ex.: `http://NOME-DO-PC:8501`) e cria a pasta **`compartilhar`** com o atalho
     "Verificador ABNT (rede)" para copiar na área de trabalho dos colegas.
2. Ainda no servidor, abra `http://localhost:8501` e, na barra lateral, defina a **senha de administrador**
   e as chaves de IA. Opcionalmente, uma **senha de acesso** para todos.
3. Nos outros computadores: só abrir o atalho/endereço no navegador. Nada é instalado.

**Atualizar**: substitua os arquivos da pasta no servidor e rode `servidor.bat` (ou reinicie o PC). Quem estiver
usando só precisa recarregar a página.

Cada pessoa tem a própria sessão e o próprio documento; um não vê o trabalho do outro. Os arquivos baixados vão
para o computador de quem baixou. Para parar: `desativar_servidor.bat`. Registro em `servidor.log`.

Segurança: o servidor só aceita conexões de páginas abertas nos próprios endereços dele (bloqueia sites de fora);
as chaves de IA ficam só no servidor e nunca são enviadas ao navegador; senhas são gravadas com hash.

## Site público (alunos acessam pela internet, rodando no PC da biblioteca)

`publicar.bat` coloca o Verificador na internet com um link para os alunos. Ele roda uma **segunda cópia** do
app, só para a internet (porta 8502, fechada para a rede local), em **modo público**, e liga um **túnel** que
dá um endereço https sem abrir portas no roteador nem no firewall. O app interno da biblioteca continua igual.

| Túnel | Link | O que precisa |
|---|---|---|
| **Tailscale** (recomendado) | fixo e gratuito, ex.: `https://biblioteca.tail1234.ts.net` | instalar o Tailscale uma vez (pede administrador) e entrar com uma conta; na 1ª vez, permitir o *Funnel* no link que aparece |
| **Cloudflare** | fixo no seu domínio, ex.: `https://verificador.seudominio.org` | um domínio na Cloudflare e o token do túnel (Zero Trust > Networks > Tunnels, apontando para `http://localhost:8502`) |
| **Rápido** | `*.trycloudflare.com`, **muda a cada reinício** | nada; bom para testar |

- O link fica em `compartilhar\link_publico.txt` (e aparece na barra lateral do app interno, para o administrador).
- O site volta sozinho quando o Windows inicia; o PC precisa ficar ligado e conectado.
- `instalar.bat` (atualização) reinicia o site público com a versão nova.
- `despublicar.bat` tira o site do ar.

**Modo público** (o que muda para quem acessa pelo link):
- Sem configurações: provedores, chaves e senhas só com a **senha de administrador** (defina antes no app
  interno, barra lateral > Senhas). Sem senha, as configurações ficam bloqueadas no site.
- **Limites** por visitante, ajustáveis no `config.json`, seção `"publico"`:
  `analises_por_hora` (20), `ia_refs_por_visitante_dia` (80) e `ia_refs_total_dia` (600, soma de todos).
  Arquivo de até 30 MB.
- **Privacidade:** aviso na tela; o arquivo fica só na memória e é apagado em até 1 hora (ou ao clicar em
  *Fechar*). Na revisão por IA, só o texto das referências vai para o provedor.
- A IA usa a chave configurada pela biblioteca: acompanhe o consumo no painel do provedor.

## TCC × dissertação/tese

Logo acima do campo de envio, escolha **🔎 Detectar pelo arquivo** (padrão), **🎓 TCC** ou **📚 Dissertação ou
Tese**. Em "Detectar", o tipo é lido da nota de natureza da folha de rosto ("Tese apresentada…", "Trabalho de
Conclusão de Curso apresentado…") ou da referência do resumo, e a norma muda sozinha:

| | TCC (padrão) | Dissertação/Tese |
|---|---|---|
| Base | Diretrizes USP (ABNT) aplicadas ao TCC | **template oficial da EESC** (`assets/modelos/`) + Diretrizes USP |
| Versão original/corrigida | não confere | confere |
| Declaração de uso de IA generativa | só se existir (confere as opções) | obrigatória |
| Fonte da citação longa | 10 | 11, como no template |
| Referência do trabalho acima do resumo | "Trabalho de Conclusão de Curso (Graduação)" | "Tese (Doutorado)" / "Dissertação (Mestrado)" coerente com a folha de rosto |

Conferências do template (categoria **Folha de rosto e pré-textuais**), para os dois tipos quando se aplicam:
- capa com USP, EESC, São Carlos e ano;
- nota de natureza recuada à direita (~8 cm), citando a EESC e o título pretendido (tese ↔ Doutor,
  dissertação ↔ Mestre);
- referência do trabalho acima do resumo e do abstract;
- declaração de uso de IA generativa com só uma das opções;
- agradecimentos com as frases obrigatórias da CAPES (Código de Financiamento 001) e da FAPESP (nº do processo);
- **textos de exemplo do template esquecidos** (NOME DO AUTOR, "Elemento obrigatório…", SOBRENOME, N. P.,
  "Palavra-chave 1", campos [Ex: …], data em branco…);
- aviso quando o arquivo parece de um tipo e a norma escolhida é de outro.

Com "Dissertação ou Tese" selecionado aparece o botão para baixar o template da EESC.

## Elementos do trabalho: aparece · formatação · conteúdo

A aba **🧾 Elementos do trabalho** mostra, para cada elemento (capa, folha de rosto, ficha, errata, folha de
aprovação, dedicatória, agradecimentos, declaração de IA, epígrafe, resumo, abstract, listas, sumário,
introdução, conclusão, referências, glossário, apêndices, anexos, índice):

- **Aparece:** procurado no **arquivo inteiro** — não precisa estar na mesma página do modelo. Só vira aviso
  (Conferir) se um elemento pré-textual estiver depois da introdução.
- **Formatação:** título sem número, centralizado e em maiúsculas; nota de natureza recuada; resumo em parágrafo
  único; referências em espaço simples etc.
- **Conteúdo confere (é verdade?):**
  - autor, título e ano iguais na capa, folha de rosto, ficha catalográfica e referências do resumo/abstract;
  - orientador igual na folha de rosto e na ficha;
  - mesmo número de palavras-chave e keywords;
  - toda "Figura 3", "Tabela 2", "seção 2.1", "Apêndice A" citada no texto existe;
  - siglas da lista usadas no texto, siglas do texto presentes na lista e lista em ordem alfabética;
  - **números de página do sumário e das listas** conferidos com a paginação que o Word grava ao salvar
    (se o arquivo não tiver essa informação, a conferência é pulada);
  - obra da epígrafe nas referências; errata com o quadro Folha / Linha / Onde se lê / Leia-se.

A mesma tabela entra no relatório. "Ordem dos elementos" passou a ser só aviso (Conferir).

## Verificações que funcionam mesmo fora do modelo

Nem todo trabalho usa o template ou os estilos do Word. Estas conferências não dependem disso:

- **Títulos reconhecidos por aproximação:** "Introdução:", "Capítulo 1 – Introdução", "Referências
  Bibliográficas", "Bibliografia consultada", erros de digitação ("Referencias", "Introducão"). Quando o
  título foi achado por semelhança, aparece um aviso para conferir a grafia.
- **Títulos de seção pelo número**, mesmo sem estilo: ponto final no título, título numerado centralizado,
  tabulação/vários espaços depois do número, texto da seção na mesma linha do título.
- **Legendas:** "FIGURA 1", "Fig. 1", "TABELA 2" → "Figura 1 – ...", "Tabela 2 – ..."; imagem sem título por perto.
- **Texto corrido:** cor diferente de preto ou realce esquecido; mistura de fontes; alinhamento e recuo de
  primeira linha diferentes do restante (não comparado em arquivos convertidos de PDF).
- **Tipografia:** espaços duplos, espaço antes da pontuação ou junto aos parênteses, parênteses e aspas sem par.
- **Notas de rodapé** em fonte 10/11 e espaço simples; **número de página** em fonte menor (Diretrizes 2.2).
- **Siglas:** forma completa na primeira menção — "Universidade de São Paulo (USP)" (Diretrizes 2.4).
- **Sumário:** não lista resumo, listas, agradecimentos...; "Índice" usado no lugar de "Sumário".
- **Referências:** ano posterior ao atual (provável erro de digitação).

## Normas: padrão EESC, normas próprias e regras extras

A verificação segue uma **norma**. Há duas prontas da EESC (TCC e dissertação/tese). Na barra
lateral, em **📐 Norma usada**, escolha outra norma ou clique em **✏️ Criar ou editar norma**:

- **Valores:** margens, fonte e entrelinhas do texto, citação longa (recuo, fonte, entrelinhas), fonte das
  legendas, tamanho do resumo, autor nas citações (Silva, 2020) ou (SILVA, 2020), separador de autores nos
  parênteses (; ou e), título e destaque das referências, ordem alfabética e quais elementos conferir.
- **Regras próprias:** proibir um texto (ex.: "ou seja"), exigir um texto ao menos uma vez (ex.: "AGRADECIMENTOS")
  ou exigir em cada item (ex.: "Acesso em:" em cada referência). Aceita expressão regular.
- **Instruções para a IA:** regras de referência que diferem da ABNT; a revisão por IA passa a segui-las.
- **Ler o manual com IA:** envie o manual da instituição (PDF, Word ou texto); a IA preenche os valores e as
  regras para você conferir antes de usar.
- **Usar nesta sessão** (qualquer pessoa, inclusive no site público), **Salvar para todos** (administrador;
  fica em `perfis/*.json`) ou **Baixar (.json)** para levar a outro computador e importar depois.

A norma escolhida aparece ao lado do nome do arquivo e no relatório.

## Verificações cruzadas (sumário, listas e ilustrações)

Inspiradas nos verificadores de teses e de referências:
- **Sumário × títulos:** mesma numeração e mesmas palavras; títulos que faltam no sumário; linhas do sumário
  sem título no texto (sumário desatualizado).
- **Listas de figuras/tabelas/quadros/gráficos × legendas:** itens faltando, com outro título ou sobrando.
- **Toda ilustração citada no texto** (Figura 3, Gráficos 1 a 3...).
- **Referência repetida** e **mesmo autor e ano** sem as letras 2020a/2020b.

## IA: escolher o provedor e colocar chaves

Na barra lateral, em **IA para as referências**:

| Provedor | Chave | Observação |
|---|---|---|
| IA local (Ollama) | não precisa | gratuita, nada sai do computador; lenta sem placa de vídeo |
| Google Gemini | aistudio.google.com/apikey | tem cota gratuita |
| OpenAI (ChatGPT) | platform.openai.com/api-keys | pago por uso |
| Anthropic (Claude) | console.anthropic.com/settings/keys | pago por uso |
| Outro compatível com OpenAI | no site do serviço | OpenRouter, Groq, DeepSeek, Mistral, LM Studio... (informe o endereço) |

As referências vão **em lotes** (padrão: 15 por pedido), então 40 referências gastam só 3 pedidos. Se a cota
diária acabar no meio, o app para, guarda o que já foi revisado e explica quando a cota volta. Limites por minuto
são tratados sozinhos (o app espera e repete).

**Atenção (Gemini):** a assinatura do app Gemini (Plus/Pro) **não vale para a API**. Sem ativar o faturamento do
projeto no Google AI Studio ("Set up billing"), a chave fica no plano gratuito, com poucos pedidos por dia por modelo.

Cole a chave, clique em **🔄 Modelos** (lista os modelos da sua conta), escolha o modelo, **🧪 Testar** e
**💾 Salvar como padrão**. Na nuvem, as referências são revisadas em paralelo (padrão: 4 por vez), o que leva
segundos em vez de minutos. **Só o texto das referências é enviado ao provedor; o trabalho do aluno não sai do
computador.** O uso das APIs pode gerar custo na conta de quem forneceu a chave.

Quem não é administrador vê só os provedores já configurados e não vê as chaves.

## Aplicativo de desktop

- O atalho abre o Verificador numa **janela própria** (motor do Edge/WebView2, que já vem no Windows 10/11),
  sem navegador e sem barra de endereço. Fechar a janela encerra o app.
- Na janela, os botões viram **Salvar**: o arquivo vai para `Documentos\Verificador ABNT` e o Explorer abre
  mostrando o arquivo.
- Se a janela própria não puder ser usada no PC (ex.: WebView2 ausente), o app abre no navegador padrão,
  com os botões de download normais. Para forçar o navegador: `instalar.bat -Iniciar -Navegador`.
- Pode abrir mais de uma janela ao mesmo tempo (cada uma usa uma porta livre).

## Tabelas e quadros

Cada tabela do Word é identificada e marcada com um selo no documento:

| Selo | Significa | Regras conferidas |
|---|---|---|
| **Tabela de dados** (azul) | dados, predominantemente numéricos (IBGE) | título "Tabela N - ..." acima, fonte abaixo, sem traços verticais, cabeçalho repetido se for longa |
| **Quadro (texto)** (roxo) | informação textual: é uma ilustração | título "Quadro N – ..." acima, fonte abaixo, bordas fechadas |
| **Diagramação** (cinza) | tabela usada só para montar a página (capa, folha de aprovação, banca) | ignorada |
| **Sem título / seria Quadro / seria Tabela** (vermelho) | falta título, ou o nome não combina com o conteúdo | apontado para o autor |

Também confere: título acima (não abaixo), numeração sequencial (tabelas e quadros separados), menção no texto
("conforme a Tabela 2"), lista de tabelas/quadros a partir de 5 itens, e citações escritas dentro das células.

## Visual

- **Modo escuro:** botão 🌙/☀️ no canto do topo. Sem escolha manual, o app segue o tema do Windows.
  As folhas do documento continuam brancas, como papel.
- **Tela cheia:** com a barra de configurações fechada, o conteúdo ocupa toda a largura.
- **Recarregar a página (F5) não perde o trabalho:** o documento aberto e a revisão por IA
  voltam sozinhos (ficam na memória por até 8 h no app interno e 1 h no site público). *Fechar* apaga.

## Visualizador (somente leitura)

O trabalho aparece como num leitor de PDF: dá para ler, procurar e selecionar texto, mas **não dá para editar**.

- **Páginas reais**: se o arquivo foi salvo no Word, as páginas seguem a paginação gravada pelo Word (a página 12
  daqui é a página 12 do aluno, inclusive os versos em branco). Arquivos de outros programas são paginados pela
  altura do conteúdo. Embaixo de cada folha aparece "Página N de T".
- **Todas as marcações ficam à vista**: trecho realçado em vermelho (*Ajustar*) ou amarelo (*Conferir*), ou uma
  barra colorida na margem quando o problema vale para o parágrafo ou a tabela inteira. Cada ponto tem um
  **balão numerado** na margem direita, na ordem do texto.
- **Clique numa marcação ou balão**: abre o que está errado, como corrigir e a seção das Diretrizes.
- **À esquerda, "Pontos marcados"**: a lista com filtros (Ajustar / Conferir / OK, que também escondem as
  marcações), busca e ordem por tema ou pelo texto. Cada item mostra o número e a página; clicar leva até o
  trecho, e ‹ › passa pelas ocorrências. O botão ☰ esconde a lista para ler em tela cheia.
- **Barra de ferramentas**: página atual (digite um número e Enter para ir), zoom (− +, *Largura*, *Página*),
  **Procurar no texto** (Ctrl+F; Enter vai para o próximo), **Marcações** (liga/desliga os realces),
  **Margens** (mostra a área de texto de cada folha) e 🖨 **imprimir ou salvar em PDF com as marcações**.
- Alterações controladas que o aluno deixou no arquivo aparecem em verde (inserido) e vermelho (apagado).

## O que entregar ao aluno

| Botão | O que é |
|---|---|
| **⬇ Arquivo com comentários** | O .docx do aluno **exatamente como foi enviado**, com um comentário do Word na margem em cada problema (o que está errado, como corrigir, seção das Diretrizes) e o trecho realçado em amarelo. Nenhuma palavra muda. |
| **⬇ Baixar relatório** (aba Relatório) | Resumo em .docx: elementos do trabalho e pontos por categoria, com como corrigir e a seção da norma. Aceita observações da Biblioteca. |

Problemas de formatação que valem para o texto todo (entrelinhas, fonte, citações longas, negrito/alinhamento
das referências) recebem um comentário só, com a dica de corrigir tudo pelo estilo do Word. Problemas gerais
(margens, paginação, elementos ausentes) ficam num comentário no início do documento. Depois que o aluno
corrigir, basta enviar o arquivo novo para conferir de novo.

### Como a detecção evita erros

- **Referências quebradas em várias linhas** (comum em arquivos convertidos de PDF) são reconhecidas e tratadas
  como uma referência só.
- **Citação × referência com três métodos**: ligação exata (sobrenome + ano), tolerante (sobrenome no início da
  referência + ano) e por grafia parecida (erro de digitação, ex.: "Alcântra" × "ALCÂNTARA"). Só é "sem
  referência" (Ajustar) se os três falharem; se só o exato falhar, vira "grafias diferentes" (Conferir).
- Palavras de ligação antes do autor ("Porém, Fleury (2000)") são descartadas; chamadas em caixas de texto
  também são lidas.
- Destaque do título nas referências: aceita negrito, itálico ou sublinhado, desde que igual em toda a lista.
- Arquivos convertidos de PDF (muitas caixas de texto ou quebras de seção) recebem um aviso para pedir o .docx
  original.

## Uso

1. Abra `http://localhost:8501`, envie o .docx do aluno.
2. Veja os itens por aba. Por padrão aparecem só os que precisam de atenção.
3. (Opcional) Clique em **Revisar referências com IA**. Pode deixar rodando; não feche a aba.
4. (Opcional) Em **Versão corrigida**, escolha os grupos e clique em **Gerar versão corrigida**; baixe o arquivo.
5. Escreva observações, baixe o relatório .docx, revise e envie ao aluno.

## Limitações conhecidas

- Lê **.docx** e **PDF** (gerado pelo editor, não digitalizado). No PDF, margens, fontes, espaçamentos e
  paginação são medidos na página; o que só existe no Word (estilos, numeração automática, alterações
  controladas) não é avaliado. Tabelas muito complexas ou figuras feitas só de texto podem não ser reconhecidas.
- Não confere capa, folha de rosto e sumário em detalhe (só a presença dos elementos-chave).
- Títulos com numeração automática do Word são marcados para conferência visual.
- Citações em formatos muito fora do padrão podem não ser reconhecidas: aparecem como "referência não citada" (🟡).
- Sistema numérico de citação: confere notas de rodapé, mas não a correspondência número × referência.
- A IA local em modelos pequenos erra: trate as sugestões como rascunho.

## Testar

`exemplos/exemplo_com_erros.docx` (dissertação) e `exemplos/exemplo_tcc_com_erros.docx` (TCC) têm 20 erros
plantados (lista no início de `exemplos/gerar_exemplo.py`); o verificador aponta 23 itens em cada (os 20, mais o
sumário desatualizado e a Figura 2 não citada, consequências deles). Os `..._correto.docx` são os mesmos trabalhos corrigidos e saem sem alertas.
`exemplos/exemplo_tabelas.docx` tem 8 tabelas de tipos diferentes (casos no início de `exemplos/gerar_tabelas.py`).
Para recriar: `python exemplos\gerar_exemplo.py`.

Testes automáticos (os mesmos que o GitHub roda a cada envio): `python -m pytest` nesta pasta (precisa de
`pip install pytest`). Veja `docs/desenvolvimento.md` no repositório.

## Estrutura

```
app.py                      interface (Streamlit)
verificador/documento.py    mapa do documento (pré-textual, textual, referências)
verificador/pdf.py          PDF: estrutura remontada pelo layout, visualização das páginas e PDF com notas
verificador/pdf_regras.py   PDF: margens, paginação e notas de rodapé medidas na página
verificador/docx_util.py    leitura de formatação efetiva do Word (estilos herdados)
verificador/estrutura.py    regras de formatação, estrutura, seções, ilustrações, tabelas
verificador/referencias.py  citações no texto, cruzamento com referências, formato das referências
verificador/tabelas.py      identificação e verificação de tabelas e quadros
verificador/perfil.py       normas (padrão EESC, normas salvas em perfis/, regras próprias)
verificador/extras.py       sumário × títulos, listas × legendas, ilustrações citadas, regras próprias
verificador/trabalho.py     TCC × dissertação/tese e conferências do template EESC (pré-textuais)
verificador/elementos.py    cada elemento: aparece, formatação e conteúdo verdadeiro (consistência)
verificador/tipografia.py   verificações que não dependem do modelo (títulos, legendas, cor, espaços, siglas...)
assets/modelos/             template EESC de dissertação/tese (para baixar no app)
verificador/visualizacao.py documento em HTML (folhas A4) e localização dos problemas
verificador/texto.py        leitura do texto visível dos parágrafos (somente leitura)
componentes/area_trabalho/  visualizador somente leitura (lista de pontos + páginas com marcações)
desktop.py                  janela própria do aplicativo
verificador/ia_local.py     prompts e modelos de referência por tipo (usados por todos os provedores)
verificador/ia_provedores.py chamadas ao Ollama, Gemini, OpenAI, Claude e compatíveis
verificador/configuracao.py config.json, chaves e senhas
verificador/cache.py        aceleração (guarda leituras repetidas; não muda nenhum resultado)
servidor.py                 servidor da biblioteca (rede) e, com --publico, o site público + túnel
publicar.ps1 / publicar.bat site público na internet (Tailscale, Cloudflare ou túnel rápido)
despublicar.bat             tira o site público do ar
verificador/relatorio.py    relatório .docx
tests/                      testes automatizados (pytest): regras, robustez, cache, PDF e retrato dos resultados
```

Para atualizar quando sair nova edição das Diretrizes: regras em `estrutura.py` e `referencias.py`;
modelos de referência usados pela IA em `ia_local.py` (dicionário `MODELOS`).
