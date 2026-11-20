# Normalização ABNT — Biblioteca da EESC-USP

Ferramentas da Biblioteca da Escola de Engenharia de São Carlos (EESC-USP) para apoiar a normalização de
TCCs, dissertações e teses segundo as **Diretrizes para apresentação de dissertações e teses da USP:
parte I (ABNT), 5. ed., 2024** e o **template oficial da EESC**.

O repositório tem duas partes independentes:

| Pasta | O que é | Para quem |
|---|---|---|
| [`verificador-interno/`](verificador-interno/) | **Verificador ABNT**: app (Python/Streamlit) que lê o `.docx`, aponta o que está fora da norma, corrige o que é seguro com controle de alterações, gera o original com comentários e o relatório. Roda no PC da biblioteca, na rede local ou como site público. | Equipe da Biblioteca (e alunos, pelo site) |
| [`assistente-alunos/`](assistente-alunos/) | **Assistente "ABNT Biblioteca EESC"**: instruções e base de regras para criar um GPT (ChatGPT), Gem (Gemini) ou Projeto (Claude) que orienta os alunos. | Alunos |

Documentação de apoio em [`docs/`](docs/) e scripts de manutenção em [`scripts/`](scripts/).

## Começo rápido

### Verificador (Windows)

1. Baixe o `.zip` da versão mais recente em **Releases** (ou clone o repositório).
2. Abra a pasta `verificador-interno` e dê dois cliques em **`instalar.bat`**: instala Python, dependências,
   IA local (Ollama) e cria os atalhos. Não precisa de administrador.
3. Use o atalho **Verificador ABNT**.

Manual completo: [`verificador-interno/LEIA-ME.md`](verificador-interno/LEIA-ME.md) — instalação, servidor na
rede, site público, normas (TCC × dissertação/tese), IA, correções e limitações.

### Assistente para os alunos

Siga [`assistente-alunos/03_publicacao_GPT_Gemini_Claude.md`](assistente-alunos/03_publicacao_GPT_Gemini_Claude.md).

## Estrutura

```
.
├── verificador-interno/        app do verificador (é o que vai no .zip de instalação)
│   ├── app.py                  interface (Streamlit)
│   ├── verificador/            regras, correção, normas, IA, relatório
│   ├── componentes/            área de trabalho (visualizador/editor do documento)
│   ├── assets/                 ícone e template EESC de dissertação/tese
│   ├── exemplos/               documentos de exemplo (com erros e corretos) e geradores
│   ├── tests/                  testes automatizados (pytest)
│   ├── instalar.ps1/.bat       instalação e atualização no Windows
│   ├── servidor.py/.bat        servidor da biblioteca na rede
│   ├── publicar.ps1/.bat       site público (Tailscale/Cloudflare)
│   └── LEIA-ME.md              manual de uso
├── assistente-alunos/          GPT/Gem/Projeto para os alunos
├── docs/                       arquitetura, desenvolvimento, publicação no GitHub
├── scripts/                    enviar ao GitHub, gerar o pacote de instalação
├── .github/                    testes automáticos, geração de versões, modelos de issue/PR
└── CHANGELOG.md                histórico de versões
```

## Desenvolvimento

```bash
cd verificador-interno
pip install -r requirements.txt pytest
pytest                      # testes de regressão
streamlit run app.py        # abre o app no navegador
```

Detalhes em [`docs/desenvolvimento.md`](docs/desenvolvimento.md) e [`docs/arquitetura.md`](docs/arquitetura.md).
Contribuições: [`CONTRIBUTING.md`](CONTRIBUTING.md).

## Privacidade

Trabalhos de alunos **não** entram no repositório (o `.gitignore` bloqueia `.docx` fora de `exemplos/` e
`assets/`), assim como `config.json` (chaves de API e senhas), logs e normas salvas localmente.
