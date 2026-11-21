# Desenvolvimento

## Ambiente
```bash
cd verificador-interno
python -m venv .venv && . .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt pytest
streamlit run app.py
```

## Testes
`python -m pytest` (pasta `verificador-interno`). Os testes usam os exemplos de `exemplos/` e o template EESC:
- exemplos com erros (dissertação e TCC) → 23 pendências cada; exemplos corretos → 0 pendências e 0 correções;
- a correção reduz as pendências e preserva o texto (integridade);
- normas personalizadas, saneamento de normas importadas, consistência (ano/título/menções), template EESC.

Se mudar os exemplos, regenere com `python exemplos/gerar_exemplo.py` e ajuste os números esperados em
`tests/test_exemplos.py`.

## Criar uma verificação nova
1. Escolha o módulo (veja `docs/arquitetura.md`) e escreva uma função que recebe `d` (Documento) e devolve
   `Achado`s. Use `PF.atual()` para os valores da norma — nunca números fixos.
2. Preencha `blocos=[...]` com os índices dos parágrafos: é o que faz o clique na interface levar ao trecho e o
   que gera os comentários no original.
3. Para tirar alarmes falsos, prefira status `VERIFICAR` quando a regra depende de interpretação.
4. Se o item pertencer a um elemento, inclua o nome em `elementos.catalogo()` para aparecer na tabela.
5. Acrescente um caso em `tests/` e, se fizer sentido, um erro plantado em `exemplos/gerar_exemplo.py`.

## Publicar versão
1. Aumente `verificador-interno/VERSION` e escreva no `CHANGELOG.md`.
2. `scripts\enviar_para_github.bat -Versao` → o GitHub roda os testes, gera os `.zip` (`scripts/empacotar.py`)
   e publica em **Releases**. Na biblioteca: baixar o `.zip`, extrair por cima e rodar `instalar.bat`.
