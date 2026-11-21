"""Tipo de trabalho (TCC × dissertação/tese) e elementos pré-textuais do template oficial da EESC.

Baseado no TEMPLATE_Dissertacao_Tese_ABNT da Biblioteca EESC:
capa (USP / EESC / autor / título / São Carlos / ano) → folha de rosto com nota de natureza recuada à direita
("Tese apresentada à Escola de Engenharia de São Carlos da Universidade de São Paulo, como requisito para a
obtenção do título de Doutor em Ciências...", orientador, VERSÃO CORRIGIDA/ORIGINAL) → ficha catalográfica →
errata → folha de avaliação ou aprovação → dedicatória → agradecimentos (frases obrigatórias de CAPES/FAPESP)
→ declaração de uso de IA generativa → epígrafe → resumo e abstract precedidos da referência do trabalho
("SOBRENOME, N. P. Título. Ano. Tese (Doutorado) – Escola de Engenharia de São Carlos...") → listas → sumário.
"""
import re

from . import docx_util as du
from . import perfil as PF
from .documento import chave
from .estrutura import _amostra
from .modelo import AJUSTAR, OK, VERIFICAR, Achado

CAT = "Folha de rosto e pré-textuais"

RE_NATUREZA = re.compile(
    r"\b(TESE|DISSERTACAO|TRABALHO DE CONCLUSAO DE CURSO|MONOGRAFIA|TRABALHO DE GRADUACAO)\b.{0,40}"
    r"APRESENTAD[AO]\b")

# textos de exemplo do template que não podem ficar no trabalho entregue
SOBRAS_TEMPLATE = [
    (r"^NOME DO AUTOR$", "NOME DO AUTOR"),
    (r"TITULO: E SUBTITULO SE HOUVER|TITLE IN ENGLISH: SUBTITLE", "Título: e subtítulo se houver"),
    (r"PROF\(A\)\. DR\(A\)\.\s*$", "Orientador(a): Prof(a). Dr(a)."),
    (r"ESTA PAGINA DEVE CONTER A FICHA CATALOGRAFICA", "Esta página deve conter a ficha catalográfica"),
    (r"^SOBRENOME, N\. P\.", "SOBRENOME, N. P. (referência do resumo)"),
    (r"PALAVRA-CHAVE \d|KEYWORD \d", "Palavra-chave 1. Palavra-chave 2."),
    (r"^ELEMENTO (OBRIGATORIO|OPCIONAL)\b|^MANDATORY ELEMENT", "Elemento obrigatório/opcional (instrução)"),
    (r"\[EX:", "[Ex: ...] (instrução da declaração de IA)"),
    (r"^OPCAO [12]\s*[-–]\s*QUANDO", "Opção 1 / Opção 2 da declaração de IA"),
    (r"____ DE _{5,}", "São Carlos, ____ de ______ (data em branco)"),
    (r"PARTE (INICIAL|PRINCIPAL|FINAL) DO TEXTO", "Parte inicial/principal/final do texto (instrução)"),
]


def detectar(d):
    """('tcc'|'pos'|None, 'mestrado'|'doutorado'|None, bloco da nota de natureza ou None)."""
    pre = [b for b in d.blocos if b.tipo == "p" and b.texto and
           (d.i_intro is None or b.i < d.i_intro)][:200]
    for b in pre:
        k = chave(b.texto)
        if RE_NATUREZA.search(k):
            if re.search(r"\bTESE\b|DOUTOR", k):
                return "pos", "doutorado", b
            if re.search(r"DISSERTACAO|MESTR", k):
                return "pos", "mestrado", b
            return "tcc", None, b
    # sem nota de natureza: procura pela referência do resumo ou por menções soltas
    tudo = " ".join(chave(b.texto) for b in pre)
    if re.search(r"TESE \(DOUTORADO|TITULO DE DOUTOR", tudo):
        return "pos", "doutorado", None
    if re.search(r"DISSERTACAO \(MESTRADO|TITULO DE MESTRE", tudo):
        return "pos", "mestrado", None
    if re.search(r"TRABALHO DE CONCLUSAO DE CURSO|\(GRADUACAO\)|MONOGRAFIA", tudo):
        return "tcc", None, None
    return None, None, None


def verificar(d):
    P = PF.atual()
    EL = P["elementos"]
    tipo, grau, nat = detectar(d)
    out = []
    pre = [b for b in d.blocos if b.tipo == "p" and b.texto and (d.i_intro is None or b.i < d.i_intro)]
    k_pre = [(chave(b.texto), b) for b in pre]

    # ---- tipo de trabalho × norma escolhida
    if tipo and tipo != P.get("tipo", "tcc"):
        out.append(Achado(CAT, "Tipo de trabalho confere com a norma escolhida", VERIFICAR,
                          f"o documento parece ser {PF.TIPOS[tipo]}, mas a norma escolhida é de "
                          f"{PF.TIPOS[P.get('tipo', 'tcc')]}",
                          f"Escolha a norma de {PF.TIPOS[tipo]} no início da página (ou na barra lateral) e "
                          "verifique de novo.", "3.1", blocos=[nat.i] if nat else []))
    elif tipo:
        out.append(Achado(CAT, "Tipo de trabalho confere com a norma escolhida", OK,
                          PF.TIPOS[tipo] + (f" ({grau})" if grau else "")))

    # ---- capa
    if EL.get("capa", True):
        inicio = " ".join(k for k, _ in k_pre)   # em qualquer ponto da parte pré-textual
        falta = [n for n, rx in (("UNIVERSIDADE DE SÃO PAULO", r"UNIVERSIDADE DE SAO PAULO"),
                                 ("ESCOLA DE ENGENHARIA DE SÃO CARLOS", r"ESCOLA DE ENGENHARIA DE SAO CARLOS"),
                                 ("São Carlos", r"\bSAO CARLOS\b"), ("ano", r"\b(19|20)\d{2}\b"))
                 if not re.search(rx, inicio)]
        out.append(Achado(CAT, "Capa: instituição, local e ano", VERIFICAR if falta else OK,
                          "faltando: " + ", ".join(falta) if falta else "completa",
                          "A capa traz UNIVERSIDADE DE SÃO PAULO, ESCOLA DE ENGENHARIA DE SÃO CARLOS, nome do autor, "
                          "título, São Carlos e o ano (template EESC)." if falta else "", "3.1.1.1"))

    # ---- nota de natureza na folha de rosto
    if EL.get("natureza", True):
        if nat is None:
            exemplo = ("Tese/Dissertação apresentada à Escola de Engenharia de São Carlos da Universidade de São "
                       "Paulo, como requisito para a obtenção do título de Doutor/Mestre em Ciências – Programa "
                       "de Pós-Graduação em ..." if P.get("tipo") == "pos" else
                       "Trabalho de Conclusão de Curso apresentado à Escola de Engenharia de São Carlos da "
                       "Universidade de São Paulo, como requisito para obtenção do título de Engenheiro(a) ...")
            out.append(Achado(CAT, "Folha de rosto: nota de natureza do trabalho", VERIFICAR,
                              "não identificada", f"Incluir na folha de rosto, recuada à direita: “{exemplo}”.",
                              "3.1.2.1.1"))
        else:
            probs = []
            recuo = du.recuo_esquerdo_cm(nat.obj)
            if recuo < 6:
                probs.append(f"recuo esquerdo {recuo:.1f} cm (no template: cerca de 8 cm, do meio da mancha "
                             "para a direita)")
            k = chave(nat.texto)
            if "ESCOLA DE ENGENHARIA DE SAO CARLOS" not in k:
                probs.append("não cita a Escola de Engenharia de São Carlos")
            if grau == "doutorado" and re.search(r"TITULO DE MESTRE", k):
                probs.append("é tese, mas cita o título de Mestre")
            if grau == "mestrado" and re.search(r"TITULO DE DOUTOR", k):
                probs.append("é dissertação, mas cita o título de Doutor")
            if tipo == "pos" and not re.search(r"TITULO DE (MESTRE|DOUTOR)", k):
                probs.append("não indica o título pretendido (Mestre/Doutor em ...)")
            out.append(Achado(CAT, "Folha de rosto: nota de natureza do trabalho", AJUSTAR if probs else OK,
                              "; ".join(probs) if probs else "conforme o template",
                              "Nota de natureza recuada à direita (cerca de 8 cm), justificada, com o tipo de "
                              "trabalho, a EESC-USP e o título pretendido." if probs else "", "3.1.2.1.1",
                              [nat.texto[:150]] if probs else [], blocos=[nat.i]))

    # ---- referência do trabalho acima do resumo e do abstract
    if EL.get("referencia_resumo", True):
        for titulo, nome in ((r"RESUMO", "Resumo"), (r"ABSTRACT", "Abstract")):
            if nome == "Abstract" and not EL.get("abstract", True):
                continue
            tb = d.achar_titulo(titulo)
            if not tb:
                continue
            from .elementos import _ate_proximo_titulo
            secao = _ate_proximo_titulo(d, tb, 15)
            ref = next((b for b in secao if re.match(r"^[A-ZÀ-Ý][A-ZÀ-Ý' \-]+,\s*[A-ZÀ-Ý]", b.texto) and
                        re.search(r"\b(19|20)\d{2}\b", b.texto)), None)
            if ref is not None and secao and ref is not secao[0] and len(secao[0].texto) > 200:
                out.append(Achado(CAT, f"{nome}: referência antes do texto", VERIFICAR,
                                  "a referência do trabalho está depois do texto do resumo",
                                  "Coloque a referência do trabalho antes do texto.", "3.1.2.1.8", blocos=[ref.i]))
            if ref is None:
                out.append(Achado(CAT, f"{nome}: referência do trabalho acima do texto", AJUSTAR,
                                  "não encontrada",
                                  "Antes do texto do resumo, a referência do próprio trabalho, ex.: “SOBRENOME, N. P. "
                                  "Título: subtítulo. 2025. Tese (Doutorado) – Escola de Engenharia de São Carlos, "
                                  "Universidade de São Paulo, São Carlos, 2025.”", "3.1.2.1.8", blocos=[tb.i]))
                continue
            k = chave(ref.texto)
            probs = []
            if not re.search(r"\b(19|20)\d{2}\b", k):
                probs.append("sem o ano")
            esperado = {"doutorado": r"TESE \(DOUTORADO", "mestrado": r"DISSERTACAO \(MESTRADO"}.get(grau)
            if tipo == "pos" and esperado and not re.search(esperado, k):
                probs.append("tipo/grau diferente de " + ("Tese (Doutorado)" if grau == "doutorado"
                                                         else "Dissertação (Mestrado)"))
            if tipo == "tcc" and not re.search(r"TRABALHO DE CONCLUSAO DE CURSO|MONOGRAFIA", k):
                probs.append("sem 'Trabalho de Conclusão de Curso (Graduação)'")
            if "ESCOLA DE ENGENHARIA DE SAO CARLOS" not in k:
                probs.append("sem 'Escola de Engenharia de São Carlos'")
            out.append(Achado(CAT, f"{nome}: referência do trabalho acima do texto", AJUSTAR if probs else OK,
                              "; ".join(probs) if probs else "presente",
                              "Formato: SOBRENOME, N. P. Título. Ano. Tese (Doutorado) | Dissertação (Mestrado) | "
                              "Trabalho de Conclusão de Curso (Graduação) – Escola de Engenharia de São Carlos, "
                              "Universidade de São Paulo, São Carlos, ano." if probs else "", "3.1.2.1.8",
                              [ref.texto[:150]] if probs else [], blocos=[ref.i]))

    # ---- declaração de uso de IA generativa
    k_todos = [(chave(b.texto), b) for b in d.blocos if b.tipo == "p" and b.texto and len(b.texto) < 120]
    tem_decl = next((b for k, b in k_todos if re.search(r"USO DE INTELIGENCIA ARTIFICIAL|\bIAG\b.*DECLARA|"
                                                      r"DECLARACAO DE USO DE IA", k)), None)
    if EL.get("declaracao_ia") or tem_decl:
        if tem_decl is None:
            out.append(Achado(CAT, "Declaração de uso de IA generativa", VERIFICAR, "não encontrada",
                              "O template da EESC traz, depois dos agradecimentos, a 'DECLARAÇÃO DE USO DE "
                              "INTELIGÊNCIA ARTIFICIAL GENERATIVA', com a opção 1 (usou IAG: ferramentas, fase, "
                              "finalidade e responsabilidade) ou a opção 2 (não usou).", "template"))
        else:
            seg = [b for b in d.blocos[tem_decl.i + 1:tem_decl.i + 25] if b.tipo == "p" and b.texto]
            ks = " ".join(chave(b.texto) for b in seg)
            duas = bool(re.search(r"OPCAO 1", ks) and re.search(r"OPCAO 2", ks))
            out.append(Achado(CAT, "Declaração de uso de IA generativa", AJUSTAR if duas else OK,
                              "as duas opções do template continuam no texto" if duas else "presente",
                              "Mantenha só a opção que se aplica (usou ou não usou IAG), apague os rótulos "
                              "'Opção 1/Opção 2' e preencha os campos entre colchetes." if duas else "",
                              "template", blocos=[tem_decl.i]))

    # ---- agradecimentos: frases obrigatórias de fomento
    agr = d.achar_titulo(r"AGRADECIMENTOS?")
    if agr:
        seg = []
        for b in d.blocos[agr.i + 1:agr.i + 60]:
            if b.tipo != "p" or not b.texto:
                continue
            if re.fullmatch(r"[A-ZÀ-Ý ÇÃÕÉÊÍÓÚ]{6,}", b.texto.strip()) and not seg == []:
                break
            seg.append(b)
        k = " ".join(chave(b.texto) for b in seg)
        probs = []
        if re.search(r"\bCAPES\b|APERFEICOAMENTO DE PESSOAL", k) and \
                not re.search(r"CODIGO DE FINANCIAMENTO 001|FINANCE CODE 001", k):
            probs.append("cita a CAPES sem 'Código de Financiamento 001'")
        if re.search(r"\bFAPESP\b|AMPARO A PESQUISA DO ESTADO DE SAO PAULO", k) and \
                not re.search(r"PROCESSO\s*(N[ºO°.]*)?\s*\d{4}/\d{5}-\d", k):
            probs.append("cita a FAPESP sem o nº do processo (aaaa/nnnnn-n)")
        if probs:
            out.append(Achado(CAT, "Agradecimentos: frases obrigatórias de fomento", AJUSTAR, "; ".join(probs),
                              "CAPES: “O presente trabalho foi realizado com apoio da Coordenação de Aperfeiçoamento "
                              "de Pessoal de Nível Superior - Brasil (CAPES) - Código de Financiamento 001”. "
                              "FAPESP: “processo nº aaaa/nnnnn-n, Fundação de Amparo à Pesquisa do Estado de São "
                              "Paulo (FAPESP)”.", "template", blocos=[agr.i]))

    # ---- textos de exemplo do template esquecidos
    sobras = []
    for b in d.blocos:
        if b.tipo != "p" or not b.texto:
            continue
        k = chave(b.texto).strip()
        for rx, rot in SOBRAS_TEMPLATE:
            if re.search(rx, k):
                sobras.append((f"{rot}: “{b.texto[:70]}”", b.i))
                break
    out.append(Achado(CAT, "Textos de exemplo do template removidos", AJUSTAR if sobras else OK,
                      f"{len(sobras)} trecho(s) do modelo ainda no texto" if sobras else "nenhum",
                      "Substitua ou apague os textos de exemplo e as instruções do template antes de entregar."
                      if sobras else "", "template", _amostra([x for x, _ in sobras]),
                      blocos=[i for _, i in sobras[:300]]))
    return out
