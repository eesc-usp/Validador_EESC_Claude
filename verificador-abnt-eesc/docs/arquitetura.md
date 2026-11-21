# Arquitetura do verificador

```
.docx ──► Documento (documento.py) ──► verificações ──► Achado[] ──► interface / relatório / comentários
                                          │                              │
                       perfil de norma ───┘                              └──► correção automática (correcao.py)
                       (perfil.py)                                             revisão dupla + integridade
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
   As verificações leem `perfil.atual()`; `analisar(arquivo, perfil)` e `corrigir_duplo(..., perfil)` ativam a
   norma com um *context manager*.
4. **Saídas**: `visualizacao.py` (folhas A4 em HTML + localização dos achados), `anotacao.py` (original com
   comentários do Word), `relatorio.py` (relatório .docx), `correcao.py` (versão corrigida com controle de
   alterações; duas passadas com conferência de integridade das palavras), `edicao.py` (edição de parágrafos
   pela área de trabalho).
5. **IA** (`ia_provedores.py`, `ia_local.py`): revisão das referências em lotes (Ollama, Gemini, OpenAI,
   Anthropic, compatíveis) e leitura de manuais para criar normas.

## Interface
- `app.py` (Streamlit): envio, escolha TCC × pós, norma, área de trabalho, abas de elementos, correções, IA e relatório.
- `componentes/area_trabalho/index.html`: componente próprio (lista de achados + documento; clique leva ao trecho).
- Modos: navegador local, **desktop** (`desktop.py`, pywebview), **servidor** na rede (`servidor.py`) e **site
  público** (`servidor.py --publico` + túnel, `publicar.ps1`).
