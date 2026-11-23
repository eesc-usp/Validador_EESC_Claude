# Desenvolvimento

## Ambiente
```bash
cd verificador-interno
python -m venv .venv && . .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt pytest ruff
streamlit run app.py
```

## Testes
`python -m pytest` (pasta `verificador-interno`, cerca de 1 minuto) e `ruff check ..` (lint).

| Arquivo | O que garante |
|---|---|
| `test_exemplos.py`, `test_regras.py`, `test_fora_do_modelo.py` | exemplos com erros (23 pendências cada), exemplos corretos (0 pendências), arquivo com comentários com o texto idêntico ao enviado, normas personalizadas, template EESC, documento fora do modelo |
| `test_visualizador.py` | não existe correção nem edição; visualizador sem edição; páginas seguem a paginação do Word; todo achado aponta para um bloco desenhado |
| `test_robustez.py` | os documentos estranhos de `sinteticos.py` passam por análise, página, comentários e relatório sem erro |
| `test_cache.py` | resultados idênticos com o cache ligado e desligado (`cache.desligado()`) |
| `test_retrato.py` | tudo o que o verificador produz para os documentos públicos bate com `retrato_esperado.json` |

Se mudar os exemplos, regenere com `python exemplos/gerar_exemplo.py` e ajuste os números esperados em
`tests/test_exemplos.py`.

### Retrato dos resultados
`tests/retrato.py` registra, para cada documento público (exemplos, template e sintéticos), os achados nas duas
normas, as páginas, os locais, o arquivo com comentários e o relatório. Arquivos
`.docx` entram por um hash do XML canônico, sem datas e ids que mudam a cada execução.
- `python tests/retrato.py` mostra o que mudou em relação ao retrato guardado;
- `python tests/retrato.py --atualizar` grava o retrato novo. Use **só** quando a mudança for intencional (regra
  nova ou corrigida) e confira a diferença do `retrato_esperado.json` no commit: ela mostra exatamente quais
  achados mudaram.

## Desempenho e cache
`verificador/cache.py` acelera sem mudar resultados:
- **sempre ligado**: estilo por id guardado por documento; XPath compilado por thread; `qn()` guardado;
- **só dentro de `cache.leitura()`**, ou seja, enquanto o documento não muda (análise, página, localização dos
  achados): funções decoradas com `@por_elemento` (formatação efetiva de um parágrafo), `cache.memo(...)`
  (referências, títulos, listas de parágrafos), `Bloco.bruto` (texto cru do parágrafo).

Regras para quem mexe no código:
- função nova que **lê** um parágrafo e é chamada muitas vezes: decore com `@por_elemento` (em `docx_util.py`);
- **nunca** chame código que altera um documento dentro de `cache.leitura()`. A geração dos comentários
  (`anotacao.py`, que mexe numa cópia) roda fora dela de propósito;
- para medir, use `python -m cProfile -s tottime` sobre `verificador.analisar(...)`. Para depurar sem cache,
  defina `VERIFICADOR_SEM_CACHE=1`.

## Criar uma verificação nova
1. Escolha o módulo (veja `docs/arquitetura.md`) e escreva uma função que recebe `d` (Documento) e devolve
   `Achado`s. Use `PF.atual()` para os valores da norma — nunca números fixos.
2. Preencha `blocos=[...]` com os índices dos parágrafos: é o que faz o clique na interface levar ao trecho e o
   que gera os comentários no original.
3. Para tirar alarmes falsos, prefira status `VERIFICAR` quando a regra depende de interpretação.
4. Se o item pertencer a um elemento, inclua o nome em `elementos.catalogo()` para aparecer na tabela.
5. Acrescente um caso em `tests/` e, se fizer sentido, um erro plantado em `exemplos/gerar_exemplo.py`.
6. Rode `python tests/retrato.py`: devem aparecer só as mudanças que a regra nova deveria causar. Depois,
   `python tests/retrato.py --atualizar`.

## Publicar versão
1. Aumente `verificador-interno/VERSION` e escreva no `CHANGELOG.md`.
2. `scripts\enviar_para_github.bat -Versao` → o GitHub roda os testes, gera os `.zip` (`scripts/empacotar.py`)
   e publica em **Releases**. Na biblioteca: baixar o `.zip`, extrair por cima e rodar `instalar.bat`.
