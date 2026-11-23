# Como contribuir

1. Crie uma *branch* a partir de `main`.
2. Faça a mudança em `verificador-interno/` (ou `assistente-alunos/`).
3. Rode os testes e o lint:
   ```bash
   cd verificador-interno
   pip install -r requirements.txt pytest ruff
   python -m pytest
   ruff check ..
   ```
   Se o teste do **retrato** acusar diferenças, confira se são as que a sua mudança deveria causar. Se forem,
   rode `python tests/retrato.py --atualizar` e inclua o `retrato_esperado.json` no commit.
4. Atualize `CHANGELOG.md` e, se o uso mudou, `verificador-interno/LEIA-ME.md`.
5. Abra um *pull request*. Os testes rodam sozinhos no GitHub (Linux e Windows).

## Regras da casa
- **Nunca** versione trabalhos de alunos, `config.json`, chaves de API ou logs. Para reproduzir um problema, crie
  um `.docx` mínimo ou acrescente um caso em `exemplos/gerar_exemplo.py`.
- Toda regra nova cita a seção das Diretrizes USP ou do template EESC (campo `secao` do `Achado`).
- Otimizações não podem mudar resultados: o retrato e o teste de cache precisam continuar sem diferenças.
- Os exemplos corretos (`exemplos/*correto.docx`) devem continuar com **zero** alertas e **zero** correções.
- Scripts `.ps1`/`.bat` em ASCII e com CRLF (o `.gitattributes` cuida do CRLF; o teste confere).
- Veja [`docs/desenvolvimento.md`](docs/desenvolvimento.md) para criar uma verificação nova.
