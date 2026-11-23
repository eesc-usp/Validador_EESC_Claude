# Arquitetura do verificador

```
.docx ──► Documento (documento.py) ──► verificações ──► Achado[] ──► visualizador (somente leitura)
                                          │                         ├─► arquivo do aluno com comentários
                       perfil de norma ───┘                         └─► relatório .docx
                       (perfil.py)

O verificador nunca altera o trabalho: só lê, aponta e gera uma cópia com comentários.
```

## Fluxo
1. **`documento.py`** abre o `.docx` e monta a lista de **blocos** (parágrafos e tabelas, na ordem), as regiões
   (pré-textual, textual, pós-textual), o início da introdução e das referências, e agrupa referências
   quebradas em várias linhas.
2. **Verificações** — cada módulo devolve uma lista de `Achado` (`modelo.py`: categoria, item, status
   `OK`/`Ajustar`/`Verificar`, detalhe, como corrigir, seção da norma, trechos, blocos para localizar):
   | Módulo | Confere |
   |---|---|
   | `estrutura.py` | papel e margens, texto principal, citações longas, alíneas, elementos obrigatórios, resumo, títulos, ilustrações |
   | `tabelas.py` | classifica tabela × quadro × diagramação e confere título, fonte, bordas, menção no texto |
   | `referencias.py` | citações × referências (3 métodos), formato e ordem das referências |
   | `extras.py` | sumário × títulos, listas × legendas, ilustrações citadas, regras próprias do perfil |
   | `trabalho.py` | TCC × dissertação/tese e pré-textuais do template EESC |
   | `elementos.py` | cada elemento: aparece, formatação, conteúdo verdadeiro (consistência) + tabela-resumo |
3. **`perfil.py`** define a norma em uso (padrões EESC de TCC e de dissertação/tese, normas salvas em `perfis/`).
   As verificações leem `perfil.atual()`; `analisar(arquivo, perfil)` ativa a norma com um *context manager*.
4. **Saídas**: `visualizacao.py` (páginas em HTML, seguindo a paginação gravada pelo Word, + localização dos
   achados), `anotacao.py` (cópia do arquivo com comentários do Word; o texto não muda) e `relatorio.py`
   (relatório .docx). `texto.py` lê o texto visível de cada parágrafo.
5. **IA** (`ia_provedores.py`, `ia_local.py`): revisão das referências em lotes (Ollama, Gemini, OpenAI,
   Anthropic, compatíveis) e leitura de manuais para criar normas.
6. **`cache.py`**: aceleração transparente (estilos guardados, XPath compilado e, durante a análise, leituras
   reaproveitadas entre as verificações). Ver `docs/desenvolvimento.md`.

## Interface
- `app.py` (Streamlit): envio, escolha TCC × pós, norma, visualizador, download do arquivo comentado e abas de
  elementos, IA e relatório.
- `componentes/area_trabalho/index.html`: visualizador somente leitura (lista de achados + páginas com todas as
  marcações e balões numerados, ir para página, zoom, procurar, imprimir/PDF). Não devolve nada ao Python.
- Modos: navegador local, **desktop** (`desktop.py`, pywebview), **servidor** na rede (`servidor.py`) e **site
  público** (`servidor.py --publico` + túnel, `publicar.ps1`).
