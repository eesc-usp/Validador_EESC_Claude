"""Revisão de referências com IA local (Ollama). Nada sai do computador.

Cada referência é enviada sozinha, junto com o modelo do tipo dela (livro, artigo...)
retirado das Diretrizes USP. Isso mantém o prompt curto, o que é essencial para um
modelo pequeno rodando em CPU.
"""
import json

import requests

from .modelo import AJUSTAR, OK, VERIFICAR, Achado

URL_PADRAO = "http://localhost:11434"

REGRAS_GERAIS = """Regras gerais (ABNT NBR 6023 / Diretrizes USP):
- Entrada pelo SOBRENOME em maiúsculas, seguido dos prenomes (abreviados ou por extenso): SILVA, J. A.
- Até 3 autores separados por ponto e vírgula; 4 ou mais: todos ou o primeiro seguido de et al.
- Título principal em negrito (marcado aqui como **texto**); subtítulo sem negrito, após dois-pontos.
- Edição só a partir da 2ª: "2. ed." (inglês "2nd ed.").
- Local: Editora, ano. Sem local: [S. l.]; sem editora: [s. n.].
- Online: DOI: ... Disponível em: URL. Acesso em: 5 abr. 2019. Meses: jan. fev. mar. abr. maio jun. jul. ago. set. out. nov. dez.
- Páginas: "p. 12-34". Termina com ponto."""

MODELOS = {
    "livro": """Livro:
SOBRENOME, Prenome. **Título**: subtítulo. Edição. Local: Editora, ano. Paginação (opcional).
Ex.: CHABAY, R. W. **Física básica**: matéria e interações. 4. ed. Rio de Janeiro: LTC, 2018.
Ex.: MOREIRA, J. R. S. (org.). **Energias renováveis, geração distribuída e eficiência energética**. Rio de Janeiro: LTC, 2017.""",
    "capitulo": """Capítulo de livro:
AUTOR DO CAPÍTULO. Título do capítulo. In: AUTOR OU ORGANIZADOR (org.). **Título do livro**. Local: Editora, ano. v., cap., p. inicial-final.
Ex.: MANFROI, V. Vinho branco. In: VENTURINI FILHO, W. G. (coord.). **Bebidas alcoólicas**: ciência e tecnologia. São Paulo: Blucher, 2010. v. 1, cap. 7, p. 143-163.""",
    "artigo": """Artigo de periódico:
AUTOR. Título do artigo. **Nome do Periódico**, Local, v. X, n. Y, p. inicial-final, mês ano. DOI: ... Disponível em: ... Acesso em: ...
Ex.: FELDMANN, A. E.; MERKE, F.; STUENKEL, O. Argentina, Brazil and Chile and democracy defence in Latin America: principled calculation. **International Affairs**, Oxford, v. 95, n. 2, p. 447-467, Mar. 2019.
Ex. com número eletrônico: GOMES, R. B. et al. Novos diálogos dos estudos feministas da deficiência. **Revista Estudos Feministas**, Florianópolis, v. 27, n. 1, e48155, 2019. DOI: https://doi.org/10.1590/1806-9584-2019v27n148155.""",
    "evento": """Trabalho em evento:
AUTOR. Título do trabalho. In: NOME DO EVENTO EM MAIÚSCULAS, número., ano, Cidade. **Anais** [...]. Local: Editora, ano. p. inicial-final.
Ex.: SILVA, E. J. F.; SANTOS, L. F. M. Método, objeto e problema em duas tradições epistemológicas. In: SEMINÁRIOS EM ADMINISTRAÇÃO - SEMEAD, 21., 2018, São Paulo. **Anais** [...]. São Paulo: EAD/FEA/USP, 2018.""",
    "academico": """Trabalho acadêmico (TCC, dissertação, tese):
AUTOR. **Título**: subtítulo. Ano de depósito. Tipo (Grau em Área) – Unidade, Universidade, Local, ano.
Ex.: BARNABE, R. F. **Análise do suprimento de frutas e sua sazonalidade**: um estudo de caso na compra de frutas para uma empresa de grande porte. 2011. Trabalho de Conclusão de Curso (Bacharelado em Engenharia de Produção Mecânica) – Escola de Engenharia de São Carlos, Universidade de São Paulo, São Carlos, 2011.""",
    "norma": """Norma técnica:
ENTIDADE. **Número da norma**: título. Local: Editora, ano. Paginação.
Ex.: ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR 6023**: informação e documentação: referências: elaboração. Rio de Janeiro: ABNT, 2018. 68 p.""",
    "legislacao": """Legislação:
JURISDIÇÃO. Título, numeração e data. Ementa. **Dados da publicação**, local, seção, página, data.
Ex.: BRASIL. Medida provisória nº 1.569-9, de 11 de dezembro de 1997. Estabelece multa em operações de importação, e dá outras providências. **Diário Oficial [da] República Federativa do Brasil**: seção 1, Brasília, DF, p. 29514, 14 dez. 1997.""",
    "patente": """Patente:
INVENTOR. **Título**. Depositante: Nome. Procurador: Nome. Número. Depósito: data. Concessão: data.
Ex.: VICENTE, Marcos Fernandes. **Reservatório para sabão em pó com suporte para escova**. Depositante: Marcos Fernandes Vicente. MU8802281-1U2. Depósito: 15 out. 2008. Concessão: 29 jun. 2010.""",
    "online": """Documento online (site, software, base de dados):
AUTOR OU ENTIDADE. **Título**. Local: Editora, ano. Disponível em: URL. Acesso em: dia mês. ano.
Ex.: UNIVERSIDADE DE SÃO PAULO. **Repositório da produção intelectual**: biblioteca digital da produção intelectual. São Paulo: USP, 2012. Disponível em: www.producao.usp.br. Acesso em: 10 abr. 2019.""",
}

PROMPT = """Você é um bibliotecário revisor de referências segundo as Diretrizes USP (ABNT NBR 6023:2018).
Avalie UMA referência. Use apenas as regras e o modelo abaixo. Não invente dados que não estão na referência:
se faltar um elemento (editora, local, páginas, volume), aponte a falta em vez de preencher.

{regras}

Modelo para este tipo:
{modelo}

Referência a avaliar (negrito marcado como **texto**):
{ref}

Responda SOMENTE com JSON neste formato:
{{"status": "OK" ou "AJUSTAR", "problemas": ["problema curto", ...], "sugestao": "referência corrigida, com **negrito**, sem inventar dados; vazio se OK"}}"""


def disponivel(url=URL_PADRAO):
    try:
        r = requests.get(f"{url}/api/tags", timeout=3)
        r.raise_for_status()
        return [m["name"] for m in r.json().get("models", [])]
    except Exception:
        return None


def revisar_referencia(ref, modelo, url=URL_PADRAO, timeout=600, nucleos=None):
    prompt = PROMPT.format(regras=REGRAS_GERAIS, modelo=MODELOS.get(ref.tipo, MODELOS["livro"]), ref=ref.marcado)
    r = requests.post(f"{url}/api/chat", timeout=timeout, json={
        "model": modelo,
        "messages": [{"role": "user", "content": prompt}],
        "format": "json",
        "stream": False,
        "options": {"temperature": 0, "num_ctx": 4096, **({"num_thread": int(nucleos)} if nucleos else {})},
    })
    r.raise_for_status()
    conteudo = r.json()["message"]["content"]
    try:
        dados = json.loads(conteudo)
    except json.JSONDecodeError:
        return Achado("Referências (IA)", ref.texto[:80], VERIFICAR, "resposta da IA ilegível",
                      origem="ia", trechos=[ref.texto])
    status = OK if str(dados.get("status", "")).upper().startswith("OK") else AJUSTAR
    probs = [str(p) for p in dados.get("problemas", []) if str(p).strip()]
    if status == OK and probs:
        status = VERIFICAR
    return Achado("Referências (IA)", ref.texto[:80] + ("…" if len(ref.texto) > 80 else ""), status,
                  "; ".join(probs) or "sem problemas",
                  str(dados.get("sugestao", "")).strip() if status != OK else "",
                  f"6 ({ref.tipo})", [ref.texto], origem="ia")
