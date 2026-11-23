"""Gera exemplo_com_erros.docx: um trabalho curto com erros conhecidos, para testar o verificador.

Erros plantados (o verificador deve apontar):
 1. margem direita 2,5 cm                      9. Tabela com bordas verticais
 2. parágrafo do texto em fonte 11            10. alínea iniciando com maiúscula / sem ';'
 3. palavras-chave separadas por ';'          11. citação (SOUZA, 2019) em caixa alta
 4. resumo com menos de 150 palavras          12. citação direta sem página (Lima, 2018)
 5. '2. MÉTODO' com ponto após o número       13. citação (Prado, 2021) sem referência
 6. seção '2.3' após '2.1' (salto)            14. COSTA (2015) nas referências e não citada
 7. 'Figura 2: ...' com dois-pontos, abaixo   15. 'Souza, M.' entrada em minúsculas
    da imagem e sem fonte                     16. '2ª edição', 'pp.', 'setembro', ordem alfabética
 8. alínea b) sem ponto final
"""
import io
import struct
import sys
import zlib
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def png(w=200, h=100):
    raw = b"".join(b"\x00" + b"\x80\x80\xc0" * w for _ in range(h))
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def campo_pagina(par):
    r = par.add_run()
    for tag, attr in (("w:fldChar", {"w:fldCharType": "begin"}), ("w:instrText", None),
                      ("w:fldChar", {"w:fldCharType": "end"})):
        el = OxmlElement(tag)
        if attr:
            for k, v in attr.items():
                el.set(qn(k), v)
        else:
            el.set(qn("xml:space"), "preserve")
            el.text = " PAGE "
        r._r.append(el)


def gerar(destino, correto=False, tcc=False):
    NAT = "Trabalho de Conclusão de Curso (Graduação)" if tcc else "Dissertação (Mestrado em Engenharia)"
    c = correto
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(12)
    st.paragraph_format.line_spacing = 1.5
    for nivel in (1, 2):
        h = doc.styles[f"Heading {nivel}"]
        h.font.size = Pt(12)
        h.font.bold = True
    doc.styles["Heading 1"].paragraph_format.page_break_before = True

    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin, s.top_margin, s.right_margin, s.bottom_margin = Cm(3), Cm(3), Cm(2 if c else 2.5), Cm(2)

    def p(texto="", centro=False, negrito=False, tam=None, esp=None, recuo=None, quebra=False):
        par = doc.add_paragraph()
        if quebra:
            par.add_run().add_break(WD_BREAK.PAGE)
        r = par.add_run(texto)
        r.bold = negrito or None
        if tam:
            r.font.size = Pt(tam)
        if centro:
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if esp:
            par.paragraph_format.line_spacing = esp
        if recuo:
            par.paragraph_format.left_indent = Cm(recuo)
        return par

    # ---------------- pré-textuais
    p("UNIVERSIDADE DE SÃO PAULO\nESCOLA DE ENGENHARIA DE SÃO CARLOS", centro=True)
    p("Maria Exemplo da Silva", centro=True)
    p("Estudo de caso sobre normalização", centro=True, negrito=True)
    p("São Carlos\n2026", centro=True)
    p("MARIA EXEMPLO DA SILVA", centro=True, quebra=True)
    if tcc:
        p("Trabalho de Conclusão de Curso apresentado à Escola de Engenharia de São Carlos da Universidade de "
          "São Paulo, como requisito para obtenção do título de Engenheira de Produção.", esp=1.0, recuo=8)
    else:
        p("Versão Original", centro=True)
        p("Dissertação apresentada à Escola de Engenharia de São Carlos da Universidade de São Paulo para "
          "obtenção do título de Mestre em Ciências.", esp=1.0, recuo=8)
    p("Orientador: Prof. Dr. Fulano de Tal", recuo=8)
    p("Catalogação na publicação – Serviço de Biblioteca EESC", quebra=True, tam=10)
    p("Aprovado em:", quebra=True)
    p("DECLARAÇÃO DE USO DE INTELIGÊNCIA ARTIFICIAL GENERATIVA", centro=True, negrito=True, quebra=True)
    p("O autor declara que não foram utilizadas ferramentas de Inteligência Artificial Generativa (IAG) em "
      "nenhuma etapa da elaboração deste trabalho.", esp=1.0)
    p("RESUMO", centro=True, negrito=True, quebra=True)
    p("SILVA, M. E. Estudo de caso sobre normalização. 2026. " + NAT + " – Escola de "
      "Engenharia de São Carlos, Universidade de São Paulo, São Carlos, 2026.", esp=1.0)
    p(" ".join(["Este resumo apresenta objetivos, métodos, resultados e conclusões do trabalho de exemplo."]
               * (14 if c else 1)), esp=1.0)
    p("Palavras-chave: Normalização. Bibliotecas. Engenharia." if c else
      "Palavras-chave: Normalização; Bibliotecas; Engenharia.", esp=1.0)
    p("ABSTRACT", centro=True, negrito=True, quebra=True)
    p("SILVA, M. E. Case study on standardization. 2026. " + NAT + " – Escola de "
      "Engenharia de São Carlos, Universidade de São Paulo, São Carlos, 2026.", esp=1.0)
    p(" ".join(["This abstract has enough words to pass the minimum length check of the verifier."] * 14), esp=1.0)
    p("Keywords: Standardization. Libraries. Engineering.", esp=1.0)
    p("SUMÁRIO", centro=True, negrito=True, quebra=True)
    for linha in ("1 INTRODUÇÃO\t5", "2 MÉTODO\t6", "2.1 Materiais\t6", "2.2 Procedimentos\t7", "REFERÊNCIAS\t9"):
        p(linha)
    doc.add_section(WD_SECTION.NEW_PAGE)
    sec2 = doc.sections[-1]
    sec2.header.is_linked_to_previous = False
    hp = sec2.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    campo_pagina(hp)

    # ---------------- textuais
    doc.add_heading("1 INTRODUÇÃO", level=1)
    p("A normalização de trabalhos acadêmicos é tema recorrente (Silva, 2020). Ramos e Pestana (2014) "
      "discutem a padronização das referências em bibliotecas universitárias brasileiras.")
    p("Para alguns autores a padronização é anterior à própria norma (" + ("Souza" if c else "SOUZA") + ", 2019), e outros indicam "
      "que \"a normalização facilita a recuperação da informação científica\" (Lima, 2018"
      + (", p. 12" if c else "") + ").", tam=None if c else 11)
    p("Outro estudo reforça a importância do tema para a comunicação científica ("
      + ("Costa, 2015" if c else "Prado, 2021") + ").")
    p("A análise e a síntese racionais só podem ser feitas mentalmente. Empregam-se principalmente na "
      "filosofia e na matemática. A análise é uma espécie de indução (Cervo; Bervian, 2007, p. 35).",
      tam=10, esp=1.0, recuo=4)
    p("Os objetivos deste trabalho são:")
    p("a) identificar os erros mais comuns;" if c else "a) Identificar os erros mais comuns;")
    p("b) propor um fluxo de verificação." if c else "b) propor um fluxo de verificação")
    doc.add_heading("2 MÉTODO" if c else "2. MÉTODO", level=1)
    p("O método combina verificação automática e revisão manual pela equipe da biblioteca, como descrito "
      "nas seções seguintes deste trabalho de exemplo.")
    doc.add_heading("2.1 Materiais", level=2)
    p("Foram analisados trabalhos depositados entre 2020 e 2025 na biblioteca, segundo critérios definidos "
      "previamente pela equipe" + (", como mostram as Figuras 1 e 2." if c else ", como mostra a Figura 1."))
    p("Figura 1 – Fluxo de verificação", tam=10)
    doc.add_paragraph().add_run().add_picture(io.BytesIO(png()), width=Cm(6))
    p("Fonte: elaborado pela autora.", tam=10, esp=1.0)
    if c:
        p("Figura 2 – Tela do sistema", tam=10)
        doc.add_paragraph().add_run().add_picture(io.BytesIO(png()), width=Cm(6))
        p("Fonte: elaborado pela autora.", tam=10, esp=1.0)
    else:
        doc.add_paragraph().add_run().add_picture(io.BytesIO(png()), width=Cm(6))
        p("Figura 2: Tela do sistema", tam=10)
    doc.add_heading("2.2 Procedimentos" if c else "2.3 Procedimentos", level=2)
    p("Os procedimentos seguiram as etapas apresentadas na Tabela 1, com conferência dupla dos resultados "
      "por dois bibliotecários.")
    p("Tabela 1 - Erros por tipo", tam=10)
    t = doc.add_table(rows=2, cols=2)
    if not c:
        t.style = "Table Grid"
    t.cell(0, 0).text, t.cell(0, 1).text = "Tipo", "Quantidade"
    t.cell(1, 0).text, t.cell(1, 1).text = "Referência", "42"
    p("Fonte: elaborado pela autora.", tam=10, esp=1.0)

    # ---------------- referências
    p("REFERÊNCIAS", centro=True, negrito=True, quebra=True)
    refs = [
        ("SILVA, J. A. ", "Normalização documentária", ". 2ª edição. São Paulo: Atlas, 2020."),
        ("RAMOS, A.; PESTANA, B. Padronização em bibliotecas. ", "Revista Brasileira de Biblioteconomia",
         ", São Paulo, v. 10, n. 2, pp. 15-30, 2014."),
        ("LIMA, C. ", "Recuperação da informação", ". Brasília, DF: IBICT, 2018. Disponível em: "
         "https://exemplo.br/lima. Acesso em: 5 setembro 2019."),
        ("Souza, M. ", "História da normalização", ". Rio de Janeiro: Interciência, 2019."),
        ("COSTA, P. ", "Referências bibliográficas", ". Campinas: Papirus, 2015."),
        ("CERVO, A. L.; BERVIAN, P. A. ", "Metodologia científica", ". 6. ed. São Paulo: Pearson, 2007."),
    ]
    if c:
        refs = [
            ("CERVO, A. L.; BERVIAN, P. A. ", "Metodologia científica", ". 6. ed. São Paulo: Pearson, 2007."),
            ("COSTA, P. ", "Referências bibliográficas", ". Campinas: Papirus, 2015."),
            ("LIMA, C. ", "Recuperação da informação", ". Brasília, DF: IBICT, 2018. Disponível em: "
             "https://exemplo.br/lima. Acesso em: 5 set. 2019."),
            ("RAMOS, A.; PESTANA, B. Padronização em bibliotecas. ", "Revista Brasileira de Biblioteconomia",
             ", São Paulo, v. 10, n. 2, p. 15-30, 2014."),
            ("SILVA, J. A. ", "Normalização documentária", ". 2. ed. São Paulo: Atlas, 2020."),
            ("SOUZA, M. ", "História da normalização", ". Rio de Janeiro: Interciência, 2019."),
        ]
    for antes, negr, depois in refs:
        par = doc.add_paragraph()
        par.paragraph_format.line_spacing = 1.0
        par.add_run(antes)
        par.add_run(negr).bold = True
        par.add_run(depois)
        doc.add_paragraph().paragraph_format.line_spacing = 1.0
    doc.save(destino)


if __name__ == "__main__":
    pasta = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent
    gerar(pasta / "exemplo_com_erros.docx")
    gerar(pasta / "exemplo_correto.docx", correto=True)
    gerar(pasta / "exemplo_tcc_com_erros.docx", tcc=True)
    gerar(pasta / "exemplo_tcc_correto.docx", correto=True, tcc=True)
    print("gerados em:", pasta)
