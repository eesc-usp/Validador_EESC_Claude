# Histórico de versões

Versões do **Verificador ABNT** (`verificador-interno/VERSION`). Para publicar uma versão nova, aumente o número
em `VERSION`, descreva aqui e rode `scripts\enviar_para_github.bat -Versao`.

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
