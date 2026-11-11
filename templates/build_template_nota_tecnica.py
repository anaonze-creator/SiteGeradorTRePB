"""Gera Template_Nota_Tecnica.docx com bookmarks nos pontos de
preenchimento identificados em js/dados-nota-tecnica.js. Mesmo esqueleto
de 14 seções da Nota de Demanda (art. 108/115 do RLC), com redação
tecnicamente orientada (art. 117), mais um bloco condicional de 4
seções extras (art. 123, V-VIII) para obras e serviços de engenharia.
Ver docs/superpowers/specs/2026-08-18-nota-demanda-nota-tecnica-design.md."""
import re
import unicodedata
import docx
from docx.shared import Pt

doc = docx.Document()
doc.styles["Normal"].font.name = "Arial"
doc.styles["Normal"].font.size = Pt(11)

_bookmark_id = [0]


def add_heading(texto):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(12)
    return p


def add_paragraph_with_bookmarks(partes):
    p = doc.add_paragraph()
    for parte in partes:
        if isinstance(parte, tuple):
            nome, texto_inicial = parte
            _bookmark_id[0] += 1
            bid = str(_bookmark_id[0])
            start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
            start.set(docx.oxml.ns.qn("w:id"), bid)
            start.set(docx.oxml.ns.qn("w:name"), nome)
            p._p.append(start)
            p.add_run(texto_inicial)
            end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
            end.set(docx.oxml.ns.qn("w:id"), bid)
            p._p.append(end)
        else:
            p.add_run(parte)
    return p


def normalizar_variante(valor):
    sem_acento = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Za-z0-9]+", "_", sem_acento).upper()


def add_bloco_variante(nome_base, variantes):
    for opcao, paragrafos in variantes.items():
        nome_bookmark = f"{nome_base}_BLOCO_{normalizar_variante(opcao)}"
        _bookmark_id[0] += 1
        bid = str(_bookmark_id[0])
        for i, texto in enumerate(paragrafos):
            p = doc.add_paragraph()
            if i == 0:
                start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
                start.set(docx.oxml.ns.qn("w:id"), bid)
                start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
                p._p.insert(0, start)
            p.add_run(texto)
            if i == len(paragrafos) - 1:
                end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
                end.set(docx.oxml.ns.qn("w:id"), bid)
                p._p.append(end)


def add_bloco_variante_sentinela_inicio(nome_bookmark):
    """Abre um bookmark que vai envolver vários parágrafos/seções
    inteiras (não cabe num único parágrafo). Combinar com
    add_bloco_variante_sentinela_fim(bid)."""
    _bookmark_id[0] += 1
    bid = str(_bookmark_id[0])
    p = doc.add_paragraph()
    start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
    start.set(docx.oxml.ns.qn("w:id"), bid)
    start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
    p._p.append(start)
    return bid


def add_bloco_variante_sentinela_fim(bid):
    p = doc.add_paragraph()
    end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
    end.set(docx.oxml.ns.qn("w:id"), bid)
    p._p.append(end)


add_heading("COMPANHIA NACIONAL DE ABASTECIMENTO")
add_heading("NOTA TÉCNICA")

add_heading("1. NECESSIDADE DE CONTRATAÇÃO (ART. 108, §1º, I, E ART. 115, I, DO RLC)")
add_paragraph_with_bookmarks([("NT01_Necessidade", "")])

add_heading("2. REFERÊNCIA A OUTROS INSTRUMENTOS DE PLANEJAMENTO DA COMPANHIA (ART. 108, §1º, II, DO RLC)")
add_paragraph_with_bookmarks([("NT02_ReferenciaPlanejamento", "")])

add_heading("3. REQUISITOS DA CONTRATAÇÃO (ART. 108, §1º, III, DO RLC)")
add_paragraph_with_bookmarks([("NT03_Requisitos", "")])

add_heading("4. ESTIMATIVA DAS QUANTIDADES, ACOMPANHADAS DAS MEMÓRIAS DE CÁLCULO E DOS DOCUMENTOS QUE LHE DÃO SUPORTE (ART. 108, §1º, IV, E ART. 115, V, DO RLC)")
add_paragraph_with_bookmarks([("NT04_EstimativaQuantidades", "")])

add_heading("5. LEVANTAMENTO DE MERCADO E JUSTIFICATIVA DE ESCOLHA DO TIPO DE SOLUÇÃO A CONTRATAR (ART. 108, §1º, V, E ART. 115, II, DO RLC)")
add_paragraph_with_bookmarks([("NT05_LevantamentoMercado", "")])

add_heading("6. ESTIMATIVAS DE PREÇOS OU PREÇOS REFERENCIAIS (ART. 108, §1º, VI, DO RLC)")
add_paragraph_with_bookmarks([("NT06_EstimativaPrecos", "")])

add_heading("7. DESCRIÇÃO DA SOLUÇÃO COMO UM TODO (ART. 108, §1º, VII, E ART. 115, IV, DO RLC)")
add_paragraph_with_bookmarks([("NT07_DescricaoSolucao", "")])

add_heading("8. JUSTIFICATIVAS PARA O PARCELAMENTO OU NÃO DA SOLUÇÃO, QUANDO NECESSÁRIA PARA INDIVIDUALIZAÇÃO DO OBJETO (ART. 108, §1º, VIII, DO RLC)")
add_bloco_variante("NT08_Parcelamento", {
    "Sim": ["8.1. A solução será parcelada, sendo a licitação realizada por item."],
    "Não": ["8.1. A solução não será parcelada, sendo a licitação realizada em lote único."],
})
add_paragraph_with_bookmarks([("NT08_Justificativa", "")])

add_heading("9. DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS (ART. 108, §1º, IX, E ART. 115, VI, DO RLC)")
add_paragraph_with_bookmarks([("NT09_ResultadosPretendidos", "")])

add_heading("10. PROVIDÊNCIAS PARA ADEQUAÇÃO DO AMBIENTE DA CONAB (ART. 108, §1º, X, DO RLC)")
add_bloco_variante("NT10_AdequacaoAmbiente", {
    "Sim": ["10.1. Será necessária a adequação do ambiente da Conab, conforme a seguir:"],
    "Não": ["10.1. Não há necessidade de adequação do ambiente da Conab para esta contratação."],
})
add_paragraph_with_bookmarks([("NT10_DetalheAdequacao", "")])

add_heading("11. CONTRATAÇÕES CORRELATAS OU INTERDEPENDENTES (ART. 108, §1º, XI, DO RLC)")
add_bloco_variante("NT11_ContratacoesCorrelatas", {
    "Sim": ["11.1. São contratos ou projetos técnicos correlatos ou interdependentes a esta contratação:"],
    "Não": ["11.1. Para este objeto, não há contratos ou projetos técnicos correlatos ou interdependentes."],
})
add_paragraph_with_bookmarks([("NT11_DetalheCorrelatas", "")])

add_heading("12. DECLARAÇÃO DA VIABILIDADE OU NÃO DA CONTRATAÇÃO (ART. 108, §1º, XII, DO RLC)")
add_bloco_variante("NT12_Viabilidade", {
    "Viável": ["12.1. A contratação mostra-se tecnicamente viável."],
    "Inviável": ["12.1. A contratação mostra-se tecnicamente inviável."],
})
add_paragraph_with_bookmarks([("NT12_Justificativa", "")])

add_heading("13. IDENTIFICAÇÃO DA MODALIDADE DE LICITAÇÃO OU DE CONTRATAÇÃO DIRETA (ART. 115, III, DO RLC)")
add_paragraph_with_bookmarks([
    "13.1. Considerando-se tudo o quanto foi exposto, a contratação será realizada na forma de ",
    ("NT13_Modalidade", ""), ".",
])
add_paragraph_with_bookmarks([("NT13_Justificativa", "")])

# Bloco condicional (art. 123, V-VIII do RLC) — só para obras e
# serviços de engenharia. Abrange 4 seções inteiras, por isso usa as
# sentinelas em vez de add_bloco_variante (que só envolve parágrafos
# soltos, não títulos com heading próprio).
add_heading("14. DA NATUREZA DO OBJETO PARA FINS DO ART. 123 DO RLC")
add_bloco_variante("NT_ObraServicoEngenharia", {
    "Sim": ["14.1. O objeto desta contratação é obra ou serviço de engenharia."],
    "Não": ["14.1. O objeto desta contratação não é obra ou serviço de engenharia."],
})

_bid_engenharia = add_bloco_variante_sentinela_inicio("NT_SecoesEngenharia_BLOCO_SIM")

add_heading("15. ADEQUAÇÃO TÉCNICA (ART. 123, V, DO RLC)")
add_paragraph_with_bookmarks([("NT15_AdequacaoTecnica", "")])

add_heading("16. FUNCIONALIDADE (ART. 123, VI, DO RLC)")
add_paragraph_with_bookmarks([("NT16_Funcionalidade", "")])

add_heading("17. REQUISITOS AMBIENTAIS (ART. 123, VII, DO RLC)")
add_paragraph_with_bookmarks([("NT17_RequisitosAmbientais", "")])

add_heading("18. ADEQUAÇÃO ÀS NORMAS VIGENTES (ART. 123, VIII, DO RLC)")
add_paragraph_with_bookmarks([("NT18_AdequacaoNormasVigentes", "")])

add_bloco_variante_sentinela_fim(_bid_engenharia)

add_heading("19. DISPOSIÇÕES FINAIS")
doc.add_paragraph("Responsável(is) pelo estudo — Área Técnica")
add_paragraph_with_bookmarks([("NT19_Responsaveis", "")])

doc.save("Template_Nota_Tecnica.docx")
print("Template_Nota_Tecnica.docx gerado com sucesso.")
