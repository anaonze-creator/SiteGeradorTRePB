"""Gera Template_Nota_Demanda.docx com bookmarks nos pontos de
preenchimento identificados em js/dados-nota-demanda.js. Estrutura de
14 seções (art. 108, §1º, I-XII, e art. 115, III e VII, do RLC), redigida
com viés funcional/administrativo (área demandante), conforme
docs/superpowers/specs/2026-08-18-nota-demanda-nota-tecnica-design.md."""
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
    """variantes: dict {valor_opcao: [paragrafo1, paragrafo2, ...]}. Cada
    variante fica dentro do seu próprio bookmark
    (<nome_base>_BLOCO_<VARIANTE_NORMALIZADA>) — o gerador remove todas
    as variantes não escolhidas."""
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


add_heading("COMPANHIA NACIONAL DE ABASTECIMENTO")
add_heading("NOTA DE DEMANDA")

add_heading("1. NECESSIDADE DE CONTRATAÇÃO (ART. 108, §1º, I, E ART. 115, I, DO RLC)")
add_paragraph_with_bookmarks([("ND01_Necessidade", "")])

add_heading("2. REFERÊNCIA A OUTROS INSTRUMENTOS DE PLANEJAMENTO DA COMPANHIA (ART. 108, §1º, II, DO RLC)")
add_paragraph_with_bookmarks([("ND02_ReferenciaPlanejamento", "")])

add_heading("3. REQUISITOS DA CONTRATAÇÃO (ART. 108, §1º, III, DO RLC)")
add_paragraph_with_bookmarks([("ND03_Requisitos", "")])

add_heading("4. ESTIMATIVA DAS QUANTIDADES, ACOMPANHADAS DAS MEMÓRIAS DE CÁLCULO E DOS DOCUMENTOS QUE LHE DÃO SUPORTE (ART. 108, §1º, IV, E ART. 115, V, DO RLC)")
add_paragraph_with_bookmarks([("ND04_EstimativaQuantidades", "")])

add_heading("5. LEVANTAMENTO DE MERCADO E JUSTIFICATIVA DE ESCOLHA DO TIPO DE SOLUÇÃO A CONTRATAR (ART. 108, §1º, V, E ART. 115, II, DO RLC)")
add_paragraph_with_bookmarks([("ND05_LevantamentoMercado", "")])

add_heading("6. ESTIMATIVAS DE PREÇOS OU PREÇOS REFERENCIAIS (ART. 108, §1º, VI, DO RLC)")
add_paragraph_with_bookmarks([("ND06_EstimativaPrecos", "")])

add_heading("7. DESCRIÇÃO DA SOLUÇÃO COMO UM TODO (ART. 108, §1º, VII, E ART. 115, IV, DO RLC)")
add_paragraph_with_bookmarks([("ND07_DescricaoSolucao", "")])

add_heading("8. JUSTIFICATIVAS PARA O PARCELAMENTO OU NÃO DA SOLUÇÃO, QUANDO NECESSÁRIA PARA INDIVIDUALIZAÇÃO DO OBJETO (ART. 108, §1º, VIII, DO RLC)")
add_bloco_variante("ND08_Parcelamento", {
    "Sim": ["8.1. A solução será parcelada, sendo a licitação realizada por item."],
    "Não": ["8.1. A solução não será parcelada, sendo a licitação realizada em lote único."],
})
add_paragraph_with_bookmarks([("ND08_Justificativa", "")])

add_heading("9. DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS (ART. 108, §1º, IX, E ART. 115, VI, DO RLC)")
add_paragraph_with_bookmarks([("ND09_ResultadosPretendidos", "")])

add_heading("10. PROVIDÊNCIAS PARA ADEQUAÇÃO DO AMBIENTE DA CONAB (ART. 108, §1º, X, DO RLC)")
add_bloco_variante("ND10_AdequacaoAmbiente", {
    "Sim": ["10.1. Será necessária a adequação do ambiente da Conab, conforme a seguir:"],
    "Não": ["10.1. Não há necessidade de adequação do ambiente da Conab para esta contratação."],
})
add_paragraph_with_bookmarks([("ND10_DetalheAdequacao", "")])

add_heading("11. CONTRATAÇÕES CORRELATAS OU INTERDEPENDENTES (ART. 108, §1º, XI, DO RLC)")
add_bloco_variante("ND11_ContratacoesCorrelatas", {
    "Sim": ["11.1. São contratos correlatos ou interdependentes a esta contratação:"],
    "Não": ["11.1. Para este objeto, não há contratos correlatos ou interdependentes."],
})
add_paragraph_with_bookmarks([("ND11_DetalheCorrelatas", "")])

add_heading("12. DECLARAÇÃO DA VIABILIDADE OU NÃO DA CONTRATAÇÃO (ART. 108, §1º, XII, DO RLC)")
add_bloco_variante("ND12_Viabilidade", {
    "Viável": ["12.1. A contratação mostra-se viável."],
    "Inviável": ["12.1. A contratação mostra-se inviável."],
})
add_paragraph_with_bookmarks([("ND12_Justificativa", "")])

add_heading("13. IDENTIFICAÇÃO DA MODALIDADE DE LICITAÇÃO OU DE CONTRATAÇÃO DIRETA (ART. 115, III, DO RLC)")
add_paragraph_with_bookmarks([
    "13.1. Considerando-se tudo o quanto foi exposto, a contratação será realizada na forma de ",
    ("ND13_Modalidade", ""), ".",
])
add_paragraph_with_bookmarks([("ND13_Justificativa", "")])

add_heading("14. DISPOSIÇÕES FINAIS (ART. 115, VII, DO RLC)")
doc.add_paragraph("Responsável(is) pelo estudo — Área Demandante")
add_paragraph_with_bookmarks([("ND14_Responsaveis", "")])

doc.save("Template_Nota_Demanda.docx")
print("Template_Nota_Demanda.docx gerado com sucesso.")
