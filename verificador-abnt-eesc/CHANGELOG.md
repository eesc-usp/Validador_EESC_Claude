# Histórico de versões

Versões do **Verificador ABNT** (`verificador-interno/VERSION`). Para publicar uma versão nova, aumente o número
em `VERSION`, descreva aqui e rode `scripts\enviar_para_github.bat -Versao`.

## 3.0.0
O verificador passa a **só apontar**: ele não corrige nem edita mais o trabalho. O texto do aluno nunca é
alterado, e quem faz as correções é o próprio aluno, no Word.

**Removido**
- Correção automática (aba *Correções automáticas*, revisão dupla, aplicação das sugestões da IA).
- Edição de parágrafos pela tela, *Desfazer* e *Original*.
- Arquivos *Documento revisado* e *Versão limpa*, e a opção de controle de alterações nas configurações.
- Módulos `correcao.py` e `edicao.py`. O instalador apaga esses arquivos ao atualizar uma pasta antiga.

**Visualizador novo, somente leitura (como um leitor de PDF)**
- **Páginas reais**: segue a paginação que o Word grava ao salvar, inclusive os versos em branco. Por exemplo,
  o template da EESC sai com 52 páginas, e o Word mostra 53. Arquivos de outros programas são paginados pela
  altura do conteúdo. Cada folha mostra "Página N de T".
- **Todas as marcações à vista ao mesmo tempo**:
  - trechos realçados em vermelho (*Ajustar*) ou amarelo (*Conferir*);
  - barra na margem quando o problema vale para o parágrafo ou a tabela;
  - balões numerados na margem, na ordem do texto.
- Clique numa marcação para ver o problema, como corrigir e a seção das Diretrizes. Clique num item da lista
  para ir até o trecho.
- **Barra de ferramentas**:
  - página atual e ir para página;
  - zoom (largura, página inteira);
  - procurar no texto (Ctrl+F);
  - ligar e desligar as marcações;
  - mostrar as margens;
  - imprimir ou salvar em PDF com as marcações.
- **Lista**:
  - número e página de cada ponto;
  - ordem por tema ou pelo texto;
  - filtros que também escondem as marcações;
  - botão para esconder a lista e ler em tela cheia.

**Tela e arquivos**
- O botão principal passa a ser **⬇ Arquivo com comentários** (`nome_comentado.docx`): o arquivo do aluno
  exatamente como foi enviado, com um comentário do Word em cada ponto.
- O indicador "Alterações aplicadas" foi trocado por "Comentários" (quantos vão no arquivo do aluno).
- A revisão das referências por IA continua: as sugestões aparecem marcadas e entram como comentário.
- O relatório não tem mais a seção de alterações aplicadas.

**Testes**
- Novo `test_visualizador.py`:
  - confere que não existe correção nem edição;
  - confere que o visualizador não edita;
  - confere a paginação do Word;
  - confere que todo achado aponta para um trecho desenhado.
- Um teste confere que o arquivo com comentários tem exatamente o mesmo texto do enviado.
- Os achados de todos os documentos do retrato continuaram idênticos aos da 2.0. Mudaram só a página desenhada e
  a nota do relatório.

## 2.0.0
Versão de otimização: **os mesmos resultados, em bem menos tempo**. Todos os achados, correções, comentários,
páginas e relatórios continuam idênticos aos da 1.12, e os testes comparam isso a cada envio. A única
diferença de resultado é a correção descrita em "Detecção".

**Velocidade** (melhor de 3 execuções, mesmo computador):

| Documento | Análise | Página na tela | Correção automática (2 passadas) |
|---|---|---|---|
| TCC real de ~80 páginas | 2,5 s → 0,4 s | 2,9 s → 0,6 s | 8,2 s → 1,9 s |
| Template EESC de dissertação/tese | 1,0 s → 0,1 s | 1,0 s → 0,2 s | 4,2 s → 1,0 s |
| Trabalho longo sintético (~600 parágrafos) | 3,7 s → 0,2 s | 3,9 s → 0,1 s | 9,2 s → 1,0 s |

- Cache transparente (`verificador/cache.py`):
  - o estilo de cada parágrafo é guardado em vez de procurado de novo em todos os estilos do arquivo;
  - as expressões XPath são compiladas uma vez só;
  - durante uma análise, a formatação de cada parágrafo, a lista de referências e os títulos são calculados uma
    vez e reaproveitados por todas as verificações.
- Na correção, cada versão do arquivo é limpa e contada uma vez só, e o texto dos parágrafos é relido só depois
  de uma troca.
- Na tela, a versão limpa, a conferência de alterações controladas e o relatório não são mais refeitos a cada
  clique.

**Detecção**
- O **sumário automático do Word** (que fica dentro de um "controle de conteúdo") passa a ser lido. Antes, as
  linhas dele eram ignoradas, o que gerava "título não achado no sumário" por engano e deslocava a contagem de
  páginas. Ele também aparece na página da tela e pode ser editado.

**Robustez e qualidade**
- Seguro com várias pessoas usando ao mesmo tempo (site e rede): sem estado global na leitura das citações.
- Ordem das correções sempre igual entre execuções (antes podia variar a ordem no registro).
- Testes novos:
  - 17 documentos estranhos (vazio, só tabela, equações, campos, controles de conteúdo, estilos quebrados,
    tabelas aninhadas, parágrafo gigante...) passam por todas as funções;
  - os resultados com e sem cache são comparados;
  - um **retrato** dos resultados de todos os documentos públicos (`tests/retrato.py`) acusa qualquer mudança
    que não tenha sido feita de propósito.
- Lint automático (`ruff`) no GitHub; versões máximas nas dependências (`streamlit<2`, `python-docx<2`), para
  uma atualização grande de biblioteca não quebrar a instalação.

## 1.12.0
- Envio aceita **somente .docx**, com aviso claro e conferência de que o arquivo é mesmo um .docx do Word.
- Identificação tolerante a trabalhos fora do modelo: títulos com pontuação, "CAPÍTULO 1 – ...", grafia
  aproximada, títulos de seção sem estilo, legendas que são frases.
- Novas verificações: título de seção com ponto/centralizado/espaço após o número/texto na mesma linha;
  legendas "FIGURA"/"Fig."; imagem sem título; cor e realce no texto; mistura de fontes; alinhamento e recuo
  desiguais; espaços duplos e antes da pontuação; parênteses e aspas sem par; notas de rodapé e número de
  página em fonte menor; sigla por extenso na 1ª menção; sumário sem pré-textuais; "Índice" no lugar do
  sumário; ano de referência no futuro.
- Correção automática de espaços e de títulos de seção (sem alterar palavras).

## 1.11.0
- Aba **Elementos do trabalho**: cada elemento conferido em *aparece · formatação · conteúdo*, procurado no arquivo
  inteiro (sem exigir a posição/página do modelo).
- Conferências de consistência: autor, título e ano iguais na capa, folha de rosto, ficha e referências do
  resumo/abstract; orientador; nº de palavras-chave × keywords; figuras/tabelas/seções/apêndices citados existem;
  lista de siglas (uso e ordem); páginas do sumário e das listas pela paginação gravada pelo Word; epígrafe nas
  referências; errata.
- "Ordem dos elementos" passou a ser aviso. Tabela de elementos no relatório.

## 1.10.0
- Escolha inicial **TCC × dissertação/tese**, com detecção automática pela folha de rosto.
- Norma de dissertação/tese baseada no **template oficial da EESC** (nota de natureza, referência acima do resumo,
  declaração de uso de IA generativa, frases de fomento CAPES/FAPESP, textos de exemplo esquecidos). Botão para
  baixar o template.

## 1.9.0
- **Normas personalizáveis** (padrão EESC): valores, elementos, regras próprias (proibir/exigir texto), instruções
  para a IA, importação do manual da instituição com IA, salvar/baixar/importar.
- Verificações cruzadas: sumário × títulos, listas × legendas, ilustrações citadas, referências repetidas,
  mesmo autor e ano sem 2020a/2020b.

## 1.8.0
- **Site público** pelo PC da biblioteca (Tailscale Funnel, Cloudflare Tunnel ou túnel rápido), com modo público
  (limites de uso, privacidade, configurações só com senha).
- Modo escuro, tela cheia com a barra lateral fechada, trabalho preservado ao recarregar a página.

## 1.7.0
- Detecção mais robusta (referências quebradas em várias linhas, três métodos de ligação citação × referência,
  palavras de ligação, caixas de texto, citações longas × listas recuadas, arquivo convertido de PDF...).
- **Revisão dupla** na correção automática, com verificação de integridade do texto do aluno.

## 1.0 – 1.6
- Verificação por regras das Diretrizes USP, revisão das referências por IA (local com Ollama ou Gemini, ChatGPT,
  Claude e compatíveis), correção automática com controle de alterações, área de trabalho com visualização e
  edição do documento, aplicativo de desktop, identificação de tabelas e quadros, servidor na rede da
  biblioteca, original com apontamentos (comentários do Word) e relatório.
