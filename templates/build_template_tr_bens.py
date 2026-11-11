"""Gera Template_TR_Bens.docx a partir do texto oficial da minuta
SEI 55379667, com bookmarks nos pontos de preenchimento identificados
em js/dados-tr-bens.js. Rodar novamente sempre que o texto oficial mudar."""
import re
import unicodedata
import docx
from docx.shared import Pt, Cm

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
    """partes: lista de strings (texto fixo) e tuplas (bookmark_name, texto_inicial)."""
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


def _wrap_paragraph_in_bookmark(paragrafo, nome_bookmark):
    _bookmark_id[0] += 1
    bid = str(_bookmark_id[0])
    start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
    start.set(docx.oxml.ns.qn("w:id"), bid)
    start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
    paragrafo._p.insert(0, start)
    end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
    end.set(docx.oxml.ns.qn("w:id"), bid)
    paragrafo._p.append(end)


def add_tabela(cabecalhos, linhas=None, linhas_vazias=0):
    """Tabela simples "Table Grid": cabecalhos vira a 1ª linha; linhas é uma
    lista de tuplas com o texto de cada célula; linhas_vazias acrescenta N
    linhas em branco (mesma quantidade de colunas) após as linhas de dados."""
    tabela = doc.add_table(rows=1, cols=len(cabecalhos))
    tabela.style = "Table Grid"
    for celula, texto in zip(tabela.rows[0].cells, cabecalhos):
        celula.text = texto
    for linha_valores in (linhas or []):
        linha = tabela.add_row()
        for celula, texto in zip(linha.cells, linha_valores):
            celula.text = texto
    for _ in range(linhas_vazias):
        tabela.add_row()
    return tabela


def add_linha_modelo_tabela(tabela, nome_bookmark, textos_por_coluna):
    linha = tabela.add_row()
    for celula, texto in zip(linha.cells, textos_por_coluna):
        celula.text = texto
    _bookmark_id[0] += 1
    bid = str(_bookmark_id[0])
    paragrafo = linha.cells[0].paragraphs[0]
    start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
    start.set(docx.oxml.ns.qn("w:id"), bid)
    start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
    paragrafo._p.insert(0, start)
    end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
    end.set(docx.oxml.ns.qn("w:id"), bid)
    paragrafo._p.append(end)
    return linha


# Larguras de coluna da Tabela 1 (matriz principal) medidas diretamente do
# anexo oficial (anexos RLC/anexo_ii_minuta_de_matriz_de_riscosodt.odt,
# style:table-column-properties de Tabela1.A a Tabela1.J) — reproduzidas
# aqui em cm para que a tabela gerada bata com o layout do RLC.
LARGURAS_TABELA1_MATRIZ_CM = [0.69, 2.93, 1.59, 2.03, 2.47, 1.4, 1.38, 1.61, 1.9, 2.1]


def _definir_larguras_colunas(tabela, larguras_cm):
    tabela.autofit = False
    for linha in tabela.rows:
        for celula, largura in zip(linha.cells, larguras_cm):
            celula.width = Cm(largura)


def _definir_tamanho_fonte(tabela, tamanho_pt, linhas=None):
    """Aplica o tamanho de fonte a todas as células da tabela (ou só às
    linhas indicadas por índice, se `linhas` for informado)."""
    for i, linha in enumerate(tabela.rows):
        if linhas is not None and i not in linhas:
            continue
        for celula in linha.cells:
            for paragrafo in celula.paragraphs:
                for run in paragrafo.runs:
                    run.font.size = Pt(tamanho_pt)


def add_matriz_de_riscos_rlc(nome_bookmark_objeto):
    """Anexo II do RLC — Minuta de Matriz de Riscos, transcrita
    integralmente (cabeçalho de dois níveis, coluna de índice numérico e
    coluna de Objeto, conforme a estrutura real do anexo oficial) a
    partir de anexos RLC/anexo_ii_minuta_de_matriz_de_riscosodt.odt.
    Fontes e larguras de coluna reproduzem o anexo oficial (10pt no
    cabeçalho, 9pt no corpo — o odt usa fontes entre 8pt e 10pt)."""
    add_paragraph_with_bookmarks(["Objeto: ", (nome_bookmark_objeto, "")])
    tabela = doc.add_table(rows=2, cols=10)
    tabela.style = "Table Grid"

    tabela.cell(0, 1).merge(tabela.cell(0, 4))
    tabela.cell(0, 1).text = "Identificação"
    tabela.cell(0, 5).merge(tabela.cell(0, 6))
    tabela.cell(0, 5).text = "5. Avaliação"
    tabela.cell(0, 8).merge(tabela.cell(0, 9))
    tabela.cell(0, 8).text = "Tratamento ao Risco"

    cabecalhos = ["Nº", "1. Objeto", "2. Evento de Risco", "3. Causas",
                  "4. Consequências", "Probabilidade (P)", "Impacto (I)",
                  "6. Nível de Risco (P) x (I)", "7. Resposta ao evento do Risco",
                  "8. Responsável"]
    for i, texto in enumerate(cabecalhos):
        tabela.cell(1, i).text = texto

    for indice in ["1", "2", "3", "4", "5"]:
        linha = tabela.add_row()
        linha.cells[0].text = indice

    _definir_larguras_colunas(tabela, LARGURAS_TABELA1_MATRIZ_CM)
    _definir_tamanho_fonte(tabela, 10, linhas=[0, 1])
    _definir_tamanho_fonte(tabela, 9, linhas=range(2, len(tabela.rows)))

    doc.add_paragraph(
        "1. Descrição do objeto previsto para contratação. 2. O risco é um evento incerto "
        "que, se ocorrer, afeta a realização dos objetivos da contratação. 3. Condições que "
        "viabilizam a concretização de um evento de risco. 4. Identificação de quais são as "
        "consequências no caso da ocorrência do risco. 5. A avaliação da probabilidade e do "
        "impacto deverá ser analisada em uma escala de 1 a 5, conforme definida nas tabelas "
        "abaixo:")
    tabela_prob_titulo = add_tabela(
        ["Escala de Probabilidade (P) – Associada às chances do risco ocorrer", "", ""],
    )
    _definir_tamanho_fonte(tabela_prob_titulo, 9)
    tabela_prob = add_tabela(
        ["Descritor", "Descrição", "Nível"],
        linhas=[
            ("Muito Baixa", "Evento extraordinário, sem histórico de ocorrência.", "1"),
            ("Baixa", "Evento casual e inesperado, muito embora raro, há histórico de sua ocorrência.", "2"),
            ("Média", "Evento esperado, de frequência reduzida, e com histórico de ocorrência parcialmente conhecido.", "3"),
            ("Alta", "Evento usual, com histórico de ocorrência amplamente conhecido.", "4"),
            ("Muito Alta", "Evento repetitivo e constante.", "5"),
        ],
    )
    _definir_tamanho_fonte(tabela_prob, 9)
    tabela_impacto = add_tabela(
        ["Descritor", "Descrição", "Nível"],
        linhas=[
            ("Muito Baixo", "Impacto insignificante nos objetivos.", "1"),
            ("Baixo", "Impacto mínimo nos objetivos.", "2"),
            ("Médio", "Impacto mediano nos objetivos, com possibilidade de recuperação.", "3"),
            ("Alto", "Impacto significante nos objetivos, com possibilidade remota de recuperação.", "4"),
            ("Muito Alto", "Impacto máximo nos objetivos, sem possibilidade de recuperação.", "5"),
        ],
    )
    _definir_tamanho_fonte(tabela_impacto, 9)
    doc.add_paragraph(
        "6. Após o resultado do cálculo da probabilidade x impacto será obtido o nível do "
        "risco, que poderá ser classificado como baixo, médio, elevado e extremo, conforme "
        "tabela abaixo:")
    tabela_faixa = add_tabela(
        ["Faixa (P x I)", "Nível de Risco"],
        linhas=[
            ("1 a 2", "Baixo"),
            ("3 a 6", "Médio"),
            ("8 a 12", "Elevado"),
            ("15 a 25", "Extremo"),
        ],
    )
    _definir_tamanho_fonte(tabela_faixa, 9)
    doc.add_paragraph(
        "7. Tratar o risco consiste em propor ações para prevenir, transferir, mitigar ou "
        "aceitar o risco. Neste campo, deve-se descrever a ação/resposta mais adequada para "
        "tratamento do risco identificado. 8. Identificar o responsável/os responsáveis pela "
        "ação proposta.")
    tabela_elaborado = add_tabela(["Elaborado por: _________________________________________________ Nome e Assinatura"])
    _definir_tamanho_fonte(tabela_elaborado, 9)


def add_conditional_block(nome_base, texto_sim, texto_nao):
    """Dois parágrafos, cada um dentro do seu próprio bookmark
    (<nome_base>_BLOCO_SIM e <nome_base>_BLOCO_NAO). O gerador (Task 9)
    sempre remove o parágrafo do lado não escolhido pela resposta —
    nunca deixa os dois lados contraditórios no documento final."""
    p_sim = doc.add_paragraph()
    p_sim.add_run(texto_sim)
    _wrap_paragraph_in_bookmark(p_sim, f"{nome_base}_BLOCO_SIM")

    p_nao = doc.add_paragraph()
    p_nao.add_run(texto_nao)
    _wrap_paragraph_in_bookmark(p_nao, f"{nome_base}_BLOCO_NAO")


def normalizar_variante(valor):
    sem_acento = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Za-z0-9]+", "_", sem_acento).upper()


def add_bloco_variante(nome_base, variantes):
    """variantes: dict {valor_opcao: [paragrafo1, paragrafo2, ...]}. Cada
    item de paragrafos pode ser uma string (texto fixo) ou uma lista de
    partes no mesmo formato de add_paragraph_with_bookmarks (strings e
    tuplas (nome_campo, texto_inicial)), para permitir um bookmark de
    campo dentro de uma variante."""
    for opcao, paragrafos in variantes.items():
        nome_bookmark = f"{nome_base}_BLOCO_{normalizar_variante(opcao)}"
        _bookmark_id[0] += 1
        bid = str(_bookmark_id[0])
        primeiro = None
        ultimo = None
        for item in paragrafos:
            partes = item if isinstance(item, list) else [item]
            p = doc.add_paragraph()
            for parte in partes:
                if isinstance(parte, tuple):
                    nome_campo, texto_inicial = parte
                    _bookmark_id[0] += 1
                    bid_campo = str(_bookmark_id[0])
                    start_campo = docx.oxml.shared.OxmlElement("w:bookmarkStart")
                    start_campo.set(docx.oxml.ns.qn("w:id"), bid_campo)
                    start_campo.set(docx.oxml.ns.qn("w:name"), nome_campo)
                    p._p.append(start_campo)
                    p.add_run(texto_inicial)
                    end_campo = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
                    end_campo.set(docx.oxml.ns.qn("w:id"), bid_campo)
                    p._p.append(end_campo)
                else:
                    p.add_run(parte)
            if primeiro is None:
                primeiro = p
            ultimo = p
        start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
        start.set(docx.oxml.ns.qn("w:id"), bid)
        start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
        primeiro._p.insert(0, start)
        end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
        end.set(docx.oxml.ns.qn("w:id"), bid)
        ultimo._p.append(end)


def add_bloco_variante_sentinela_inicio(nome_bookmark):
    """Para blocos que misturam parágrafos e tabelas: marca só o INÍCIO
    do bloco num parágrafo vazio. Combinar com
    add_bloco_variante_sentinela_fim(bid) depois de todo o conteúdo já ter
    sido adicionado ao documento. O JS remove tudo entre as duas
    sentinelas por recorte de texto bruto — não importa o que tem no meio."""
    p = doc.add_paragraph()
    _bookmark_id[0] += 1
    bid = str(_bookmark_id[0])
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


# Capa
doc.add_paragraph("COMPANHIA NACIONAL DE ABASTECIMENTO")
titulo = doc.add_paragraph()
titulo_run = titulo.add_run("TERMO DE REFERÊNCIA")
titulo_run.bold = True

# 1. DO OBJETO
add_heading("1. DO OBJETO")
add_paragraph_with_bookmarks([
    "1.1. O presente Termo de Referência tem por objeto a aquisição de ",
    ("S01_DescricaoObjeto", "(descrição do material a ser adquirido)"),
    ", conforme especificações, condições, quantidades e exigências estabelecidas neste instrumento.",
])
add_paragraph_with_bookmarks([
    "1.2. O custo total estimado para aquisição dos produtos objetos deste Termo de Referência é de R$ ",
    ("S01_ValorEstimado", "________"),
    ". O critério utilizado para a formação da estimativa foi de ",
    ("S01_MetodologiaEstimativa", "(MÉDIA, MEDIANA OU MENOR PREÇO)"),
    " dos preços praticados no mercado.",
])
add_paragraph_with_bookmarks(["Metodologia detalhada: ", ("S01_MetodologiaEstimativaOutro", "____")])
add_bloco_variante("S01_CriterioJulgamentoSRP", {
    "Não": [[
        "1.3. O critério de julgamento das propostas será o de ",
        ("S01_CriterioJulgamento", "[menor preço ou maior desconto]"),
        " ofertado. Para fins de julgamento das propostas o valor estimado/valor de referência será ",
        ("S01_SigiloValor", "(público ou sigiloso)"),
        ". O modo de disputa adotado no certame será ",
        ("S01_ModoDisputa", "(aberto ou aberto e fechado)"),
        ".",
    ]],
    "Sim": [[
        "1.3. O critério de julgamento das propostas será o de menor preço ofertado, "
        "por se tratar de licitação para Sistema de Registro de Preços (art. 328 do RLC). "
        "Para fins de julgamento das propostas o valor estimado/valor de referência será ",
        ("S01_SigiloValor", "(público ou sigiloso)"),
        ". O modo de disputa adotado no certame será ",
        ("S01_ModoDisputa", "(aberto ou aberto e fechado)"),
        ".",
    ]],
})
add_paragraph_with_bookmarks([
    "1.4. Os recursos orçamentários decorrentes da aquisição do objeto deste Termo de Referência "
    "estão consignados no Orçamento da Conab para o ano 20XX, conforme segue: ND: (",
    ("S01_ND", "XXXX"),
    "), PTRES: (",
    ("S01_PTRES", "XXXX"),
    "), FONTE: (",
    ("S01_Fonte", "XXXX"),
    ").",
])
add_bloco_variante("S01_SRP", {
    "Sim": [
        "1.5. Na licitação para registro de preços não é necessária a indicação de dotação "
        "orçamentária, que somente será exigida quando da formalização do Contrato ou "
        "instrumento equivalente (art. 331 do RLC)."
    ],
    "Não": [""],
})
add_bloco_variante("S01_CriterioJulgamentoSRP", {
    "Sim": [["1.5.1. Justificativa da hipótese do art. 327 do RLC que ampara o uso do SRP: ",
             ("S01_JustificativaSRP", "____")]],
    "Não": [""],
})
add_bloco_variante("S01_TIC", {
    "Sim": [
        "1.6. Por se tratar de contratação de Tecnologia da Informação, esta contratação "
        "deve estar prevista no Plano Diretor de Tecnologia da Informação e Comunicação "
        "(PDTIC) ou ser previamente aprovada pelo Comitê Estratégico de Tecnologia da "
        "Informação, nos termos dos arts. 599-600 do RLC, observadas a Resolução CGPAR nº "
        "29/2022 e as orientações do Acórdão nº 980/2023-TCU."
    ],
    "Não": [""],
})

# 2. DA JUSTIFICATIVA
add_heading("2. DA JUSTIFICATIVA")
add_paragraph_with_bookmarks(["2.1. ", ("S02_Motivacao", "(motivação da contratação)")])
doc.add_paragraph(
    "2.2. O material em apreço será adquirido na modalidade Pregão Eletrônico, por ser "
    "caracterizado como bem comum, na forma do Inciso XIII do art. 3º do Regulamento de "
    "Licitações e Contratos da Conab - RLC e, portanto, possui padrões de desempenho e "
    "qualidade que podem ser objetivamente definidos no Edital, por meio de especificações "
    "usuais de mercado.")
add_conditional_block(
    "S02_AplicaExcecaoMEEPP",
    texto_sim=(
        "2.3. Em que pese o valor estimado da contratação ser inferior ao limite de "
        "R$ 80.000,00 (oitenta mil reais), o certame não será realizado com exclusividade "
        "de participação para microempresas e empresas de pequeno porte, na forma da Lei "
        "Complementar 123/2006 e do Decreto 8.538/2015, em razão da contratação em apreço "
        "se enquadrar na exceção prevista no art. 49 da Lei nº 123/2006, isto é: "
    ),
    texto_nao="",
)
add_paragraph_with_bookmarks(["", ("S02_MotivoExcecaoMEEPP", "")])
add_bloco_variante("S02_ImpactoOrcamentarioFinanceiro", {
    "Sim": [
        "2.4. Nos termos do art. 216, IX, do RLC e do art. 16, incisos I e II, da Lei "
        "Complementar nº 101, de 2000 (LRF), a presente contratação gera despesa "
        "obrigatória de caráter continuado ou impacto orçamentário-financeiro relevante, "
        "devendo constar dos autos a estimativa do impacto orçamentário-financeiro e a "
        "declaração do ordenador de despesa quanto à adequação orçamentária e financeira "
        "com a Lei Orçamentária Anual (LOA) e à compatibilidade com o Plano Plurianual "
        "(PPA) e a Lei de Diretrizes Orçamentárias (LDO)."
    ],
    "Não": [
        "2.4. Nos termos do art. 16, §3º, da Lei Complementar nº 101, de 2000 (LRF), a "
        "presente contratação não se enquadra no caput daquele artigo, dispensando a "
        "estimativa do impacto orçamentário-financeiro e a declaração de adequação "
        "orçamentária ali previstas."
    ],
})

# 3. DO OBJETIVO
add_heading("3. DO OBJETIVO")
add_paragraph_with_bookmarks(["3.1. ", ("S03_Objetivo", "(objetivo da aquisição)")])

# 4. DA VIGÊNCIA DO CONTRATO
add_heading("4. DA VIGÊNCIA DO CONTRATO")
add_paragraph_with_bookmarks([
    "4.1. O prazo de vigência do contrato será de ",
    ("S04_PrazoVigenciaMeses", "____"),
    " meses (duração do contrato, conforme art. 461 a 463 do RLC), contados da data de sua "
    "assinatura, prorrogável na forma do art. 497 e 498 do RLC.",
])
add_paragraph_with_bookmarks([
    "4.2. Quando cabível, a prorrogação poderá se estender por ",
    ("S04_PrazoProrrogacaoMeses", "____"),
    ", desde que observados os requisitos do art. 488 do RLC, entre eles a demonstração da "
    "vantajosidade na manutenção do ajuste (em regra, por pesquisa de preços), sob pena de a "
    "Conab não poder prorrogar o Contrato.",
])

# 5. DA SELEÇÃO DO FORNECEDOR
add_heading("5. DA SELEÇÃO DO FORNECEDOR")
doc.add_paragraph("5.1. DA HABILITAÇÃO E QUALIFICAÇÃO TÉCNICA")
doc.add_paragraph(
    "5.1.1. Para a habilitação, a licitante detentora da melhor oferta deverá apresentar toda "
    "a documentação relacionada a seguir. A licitante já regularmente cadastrada e habilitada "
    "parcialmente no Sistema de Cadastramento Unificado de Fornecedores (SICAF) ficará "
    "dispensada de apresentar os documentos correlatos, quando os registros no SICAF "
    "estiverem válidos, completos e regulares.")
add_paragraph_with_bookmarks([
    "5.1.2. Para efeito de caracterização de compatibilidade é definida como relevante "
    "a seguinte parcela do objeto: ",
    ("S05_ParcelaRelevante", "(descrever a parte mais relevante do objeto e o seu quantitativo)"),
])
add_bloco_variante("S05_QuantidadeMinimaCotadaPorItem", {
    "Sim": [["5.1.3. Quantidade mínima de unidades a ser cotada, por item (art. 335, III, do "
             "RLC): ", ("S05_QuantidadeMinimaCotadaPorItem", "____")]],
    "Não": [""],
})
doc.add_paragraph(
    "5.1.4. Relativos à Habilitação Jurídica: os documentos constitutivos compatíveis com a "
    "natureza jurídica da licitante (inscrição no Registro Público de Empresas Mercantis para "
    "o empresário individual; Certificado da Condição de Microempreendedor Individual para o "
    "MEI; ato constitutivo, estatuto ou contrato social em vigor, devidamente registrado, para "
    "sociedade empresária, EIRELI ou sociedade por ações; inscrição do ato constitutivo para "
    "sociedade simples; ata de fundação e estatuto social para cooperativa; decreto de "
    "autorização para empresa estrangeira em funcionamento no País), acompanhados de todas as "
    "alterações ou da consolidação respectiva.")
add_paragraph_with_bookmarks([
    "5.1.4.1. Quando a atividade relativa ao objeto exigir registro ou autorização de "
    "funcionamento de órgão regulador (ANVISA, ANP, Ministério da Agricultura, Polícia "
    "Federal, IBAMA etc.): ", ("S05_RegistroEspecialAtividade", ""),
])
doc.add_paragraph(
    "5.1.5. Relativos à Regularidade Fiscal e Trabalhista: prova de inscrição no CNPJ; prova "
    "de regularidade fiscal perante a Fazenda Nacional (RFB e PGFN); prova de inscrição no "
    "cadastro de contribuintes estadual ou municipal, quando houver; prova de regularidade "
    "perante a Fazenda Estadual e Municipal; prova de regularidade perante o INSS e o FGTS; e "
    "prova de inexistência de débitos inadimplidos perante a Justiça do Trabalho (CNDT). "
    "Havendo alguma restrição na comprovação da regularidade fiscal de microempresa, empresa "
    "de pequeno porte ou cooperativa, será assegurado prazo de 5 (cinco) dias úteis, "
    "prorrogável por igual período, para regularização, sob pena de decadência do direito à "
    "contratação.")
doc.add_paragraph(
    "5.1.6. Relativos à Qualificação Econômico-Financeira: certidão negativa de feitos sobre "
    "falência, expedida pelo distribuidor da sede da licitante, com prazo máximo de 60 "
    "(sessenta) dias quando não mencionar validade explícita; e balanço patrimonial e "
    "demonstrações contábeis do último exercício social, vedada a substituição por balancetes "
    "ou balanços provisórios. Será considerada boa a situação financeira da licitante quando "
    "os Índices de Liquidez Geral (LG), Solvência Geral (SG) e Liquidez Corrente (LC) forem "
    "maiores que 1 (um), apurados pelas fórmulas: LG = (Ativo Circulante + Realizável a Longo "
    "Prazo) / (Passivo Circulante + Exigível a Longo Prazo); SG = Ativo Total / (Passivo "
    "Circulante + Exigível a Longo Prazo); LC = Ativo Circulante / Passivo Circulante. A "
    "licitante com qualquer desses índices igual ou inferior a 1 (um) deverá comprovar "
    "patrimônio líquido mínimo de 10% (dez por cento) do valor total estimado da contratação.")
add_paragraph_with_bookmarks([
    "5.1.7. Relativos à Qualificação Técnica: atestado de capacidade técnica que comprove o "
    "fornecimento de bens com características e quantidades compatíveis com o objeto, por "
    "período não inferior a ", ("S05_PrazoExperienciaMinimaAnos", "____"),
    " anos, fornecido por pessoas jurídicas de direito público ou privado, tomando-se por "
    "parâmetro de compatibilidade a parcela relevante indicada no item 5.1.2.",
])
doc.add_paragraph(
    "5.1.8. Será aceito o somatório de atestados para fins de comprovação de capacidade "
    "técnica. A licitante vencedora, apenas por ocasião da assinatura do Contrato, deverá "
    "comprovar inscrição na entidade profissional competente, quando for o caso, e apresentar "
    "declaração de que é Microempresa, Empresa de Pequeno Porte ou Cooperativa, se for o caso.")
doc.add_paragraph(
    "5.1.9. Como condição prévia ao exame da documentação de habilitação, será verificado o "
    "eventual descumprimento das condições de participação mediante consulta ao SICAF, ao "
    "Cadastro Nacional de Empresas Inidôneas e Suspensas (CEIS), ao Cadastro Nacional de "
    "Condenações Cíveis por Atos de Improbidade Administrativa e à Lista de Inidôneos do TCU, "
    "em nome da licitante e de seu sócio majoritário.")
doc.add_paragraph(
    "5.1.10. Será inabilitada a licitante que não comprovar sua habilitação, ressalvado o "
    "saneamento de erros ou falhas que não alterem a substância das propostas e dos "
    "documentos, mediante despacho fundamentado. Será considerada vencedora a licitante que "
    "oferecer o menor preço e atender às exigências de habilitação deste Termo de Referência.")
doc.add_paragraph("5.2. DA AMOSTRA DOS BENS")
add_conditional_block(
    "S05_ExigeAmostra",
    texto_sim="5.2.1. Poderá ser solicitada à licitante provisoriamente classificada em primeiro lugar a apresentação de amostra dos bens.",
    texto_nao="5.2.1. Não será exigida amostra dos bens objeto desta licitação.",
)
add_paragraph_with_bookmarks([
    "5.2.1.1. Prazo para apresentação: ", ("S05_PrazoAmostraDias", "____"), " dias úteis.",
])
add_paragraph_with_bookmarks([
    "5.2.1.2. A amostra deverá ser apresentada em ", ("S05_LocalAmostra", "(local)"),
    ", no horário de ", ("S05_HorarioAmostra", "(horário)"), ".",
])

# 6. DA DESCRIÇÃO DA AQUISIÇÃO
add_heading("6. DA DESCRIÇÃO DA AQUISIÇÃO")
doc.add_paragraph("6.1. DA FORMA DE FORNECIMENTO, DO PRAZO, DO LOCAL DE ENTREGA E DA GARANTIA DOS BENS")
add_paragraph_with_bookmarks([
    "6.1.1. Os bens objetos deste Termo de Referência deverão ser entregues de forma ",
    ("S06_FormaFornecimento", "(integral/parcelada)"),
    ", devidamente embalados e lacrados, em ", ("S06_LocalEntrega", "(local)"),
    ", no horário de ", ("S06_HorarioEntrega", "(horário)"), ".",
])
_bid_enderecos = add_bloco_variante_sentinela_inicio("S06_EnderecosAdicionais_CONTAINER")
doc.add_paragraph(
    "6.1.2. Quando o objeto envolver a entrega em mais de um endereço, os locais adicionais "
    "são os seguintes:")
tabela_enderecos = add_tabela(["Endereço", "Observação"])
add_linha_modelo_tabela(tabela_enderecos, "S06_EnderecosAdicionais_LINHA_MODELO", ["((ENDERECO))", "((OBSERVACAO))"])
add_bloco_variante_sentinela_fim(_bid_enderecos)
add_paragraph_with_bookmarks([
    "6.1.3. O prazo de entrega dos bens é de ", ("S06_PrazoEntregaDias", "____"),
    " dias, contados a partir da assinatura do contrato, em remessa ",
    ("S06_TipoRemessa", "(única ou parcelada)"), ".",
])
add_paragraph_with_bookmarks([
    "6.1.4. Os bens deverão ter um período de garantia de, no mínimo, ",
    ("S06_GarantiaBensMeses", "____"), " meses, contados do recebimento definitivo pela Contratante.",
])
add_paragraph_with_bookmarks([
    "6.1.5. Durante o período de garantia, a Contratada deverá substituir o bem, no prazo "
    "máximo de ", ("S06_PrazoSubstituicaoDias", "____"), " dias, após ser comunicado.",
])
doc.add_paragraph("6.2. DO RECEBIMENTO DOS BENS")
add_paragraph_with_bookmarks([
    "6.2.1. Os bens serão recebidos provisoriamente no prazo de ",
    ("S06_PrazoRecebimentoProvisorioDias", "____"), " dias, pela fiscalização.",
])
add_paragraph_with_bookmarks([
    "6.2.2. Os bens serão recebidos definitivamente no prazo de ",
    ("S06_PrazoRecebimentoDefinitivoDias", "____"), " dias contados do recebimento provisório.",
])

# 7. DA GARANTIA CONTRATUAL
add_heading("7. DA GARANTIA CONTRATUAL")
add_bloco_variante("S07_ExigeGarantia", {
    "Sim": [[
        "7.1. A Contratada deverá apresentar, no prazo de 10 (dez) dias corridos após a "
        "assinatura do Contrato, prorrogáveis por igual período, comprovante de prestação de "
        "garantia correspondente a ", ("S07_PercentualGarantia_Inline", "____"),
        "%, conforme parâmetros do art. 439 do RLC, do ",
        ("S07_BaseCalculoGarantia", "(valor do Contrato ou valor anual do Contrato)"),
        ", mediante caução em dinheiro, seguro-garantia ou fiança bancária, cabendo à "
        "Contratada optar pela modalidade (art. 439, §1º, do RLC).",
    ]],
    "Não": ["7.1. Não será exigida garantia contratual."],
})
add_paragraph_with_bookmarks([
    "Percentual da garantia: ", ("S07_PercentualGarantia", "____"), "% do valor do Contrato "
    "— em regra, até 5% (art. 439, §2º, do RLC), podendo ser elevado para até 10% em "
    "contratações de grande vulto envolvendo complexidade técnica e riscos financeiros "
    "elevados (art. 439, §3º, do RLC).",
])
doc.add_paragraph(
    "7.2. A inobservância do prazo fixado para apresentação da garantia acarretará multa de "
    "0,08% (oito centésimos por cento) do valor do Contrato por dia de atraso, até o máximo de "
    "2% (dois por cento). O atraso superior a 25 (vinte e cinco) dias autoriza a rescisão do "
    "Contrato, conforme o inciso I do art. 569 do RLC. A garantia deverá ter validade de 90 "
    "(noventa) dias após o término da vigência do Contrato.")

# 8. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO (texto fixo)
add_heading("8. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO")
doc.add_paragraph(
    "8.1. Para a execução do ajuste, será adotado o método de trabalho baseado no conceito "
    "de delegação de responsabilidade. Nos termos dos art. 543 e 544 do RLC será designado "
    "fiscal, seu substituto, ou comissão de fiscalização para acompanhar e fiscalizar a "
    "entrega dos bens.")

# 9. DAS OBRIGAÇÕES DA CONTRATANTE (texto fixo)
add_heading("9. DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph("9.1. São obrigações da Contratante:")
doc.add_paragraph(
    "a) receber o objeto no prazo e condições estabelecidas no Edital e seus anexos;")
doc.add_paragraph(
    "b) verificar minuciosamente, no prazo fixado, a conformidade dos bens recebidos "
    "provisoriamente com as especificações constantes do Edital e seus anexos e da proposta, "
    "para fins de aceitação e recebimento definitivo;")
doc.add_paragraph(
    "c) comunicar à Contratada, por escrito, sobre imperfeições, falhas ou irregularidades "
    "verificadas no objeto fornecido, para que seja substituído, reparado ou corrigido;")
doc.add_paragraph(
    "d) rejeitar, no todo ou em parte, fornecimento de bens executado em desacordo com o "
    "previsto neste Termo de Referência;")
doc.add_paragraph(
    "e) acompanhar e fiscalizar o cumprimento das obrigações da Contratada, através de "
    "empregado ou comissão especialmente designados;")
doc.add_paragraph(
    "f) efetuar o pagamento à Contratada no valor correspondente ao fornecimento do objeto, "
    "no prazo e forma estabelecidos no Edital e seus anexos.")
doc.add_paragraph(
    "9.2. A Conab não responderá por quaisquer compromissos assumidos pela Contratada com "
    "terceiros, ainda que vinculados à execução do Contrato, bem como por qualquer dano "
    "causado a terceiros em decorrência de ato da Contratada, de seus empregados, prepostos "
    "ou subordinados.")

# 10. DAS OBRIGAÇÕES DA CONTRATADA
add_heading("10. DAS OBRIGAÇÕES DA CONTRATADA")
doc.add_paragraph(
    "10.1. A Contratada deve cumprir todas as obrigações constantes neste Termo, seus anexos "
    "e sua proposta, assumindo como exclusivamente seus os riscos e as despesas decorrentes "
    "da boa e perfeita execução do objeto e, ainda:")
doc.add_paragraph(
    "a) efetuar a entrega do objeto em perfeitas condições, conforme especificações, prazo e "
    "local constantes neste Termo e seus anexos, acompanhado da respectiva nota fiscal, na "
    "qual constarão as indicações referentes a: marca, fabricante, modelo, procedência e "
    "prazo de garantia ou validade;")
doc.add_paragraph(
    "a.1) O objeto deve estar acompanhado do manual do usuário, com uma versão em português "
    "e da relação da rede de assistência técnica autorizada;")
doc.add_paragraph(
    "b) responsabilizar-se pelos vícios e danos decorrentes do objeto, de acordo com os art. "
    "12, 13 e 17 a 27, do Código de Defesa do Consumidor (Lei nº 8.078, de 1990);")
doc.add_paragraph(
    "c) substituir, reparar ou corrigir, às suas expensas, no prazo fixado neste Termo de "
    "Referência, o objeto com avarias ou defeitos;")
doc.add_paragraph(
    "d) comunicar à Contratante, no prazo máximo de 24 (vinte e quatro) horas que antecede a "
    "data da entrega, os motivos que impossibilitem o cumprimento do prazo previsto, com a "
    "devida comprovação;")
doc.add_paragraph(
    "e) manter, durante toda a execução do contrato, em compatibilidade com as obrigações "
    "assumidas, todas as condições de habilitação e qualificação exigidas na licitação;")
doc.add_paragraph(
    "f) prestar todos os esclarecimentos que forem solicitados pela Conab, durante a "
    "realização do Contrato;")
doc.add_paragraph("g) indicar preposto para representá-la durante a execução do contrato;")
doc.add_paragraph(
    "h) garantir aos seus empregados, na execução do Contrato, política de respeito à "
    "diversidade, prevenção a práticas discriminatórias e alinhamento com os princípios "
    "institucionais de equidade de gênero e raça.")
add_bloco_variante("S10_ExigeProgramaIntegridade", {
    "Sim": [[
        "i) comprovar a implementação do Programa de Integridade, nos moldes do art. 603 e "
        "Anexo XXXII do RLC, no prazo de até ", ("S10_PrazoProgramaIntegridadeMeses", "____"),
        " meses após a assinatura do contrato.",
    ]],
    "Não": [""],
})

# 11. DO PAGAMENTO
add_heading("11. DO PAGAMENTO")
add_paragraph_with_bookmarks([
    "11.1. O pagamento será efetuado em até ", ("S11_PrazoPagamentoDiasUteis", "____"),
    " dias úteis após o recebimento definitivo do material, de acordo com a apresentação da "
    "Nota Fiscal/Fatura.",
])

# 12. DO REAJUSTE
add_heading("12. DO REAJUSTE")
add_conditional_block(
    "S12_Reajustavel",
    texto_sim="12.1. Dentro do prazo de vigência do contrato, os preços contratados poderão sofrer reajuste após o interregno de um ano.",
    texto_nao="12.1. Os preços são fixos e irreajustáveis no prazo de um ano, contado da data limite para a apresentação das propostas.",
)
add_paragraph_with_bookmarks(["Índice de referência: ", ("S12_IndiceReajuste", "")])
doc.add_paragraph(
    "12.2. O disposto nesta seção não afasta o direito ao reequilíbrio econômico-financeiro "
    "em sentido estrito, cabível a qualquer tempo, independentemente de previsão contratual, "
    "quando verificado fato imprevisível ou de consequências incalculáveis, anormal ou "
    "extraordinário, nos termos do art. 511 do RLC.")

# 13. DA SUBCONTRATAÇÃO
add_heading("13. DA SUBCONTRATAÇÃO")
add_conditional_block(
    "S13_PermiteSubcontratacao",
    texto_sim="13.1. É permitida a subcontratação parcial do objeto.",
    texto_nao="13.1. Não será admitida a subcontratação do objeto licitatório.",
)
add_paragraph_with_bookmarks([
    "Limites: entre ", ("S13_PercentualMinSubcontratacao", "____"), "% e ",
    ("S13_PercentualMaxSubcontratacao", "____"), "% do valor total do contrato.",
])
add_paragraph_with_bookmarks(["Detalhamento: ", ("S13_DetalheSubcontratacao", "")])
doc.add_paragraph("13.2. A subcontratação deverá observar o disposto no artigo 524 do RLC.")

# 14. DA ALTERAÇÃO DO CONTRATO E DA ALTERAÇÃO SUBJETIVA (texto fixo)
add_heading("14. DA ALTERAÇÃO DO CONTRATO E DA ALTERAÇÃO SUBJETIVA")
doc.add_paragraph(
    "14.1. O Contrato poderá ser alterado nas hipóteses previstas no art. 510 do RLC. A "
    "Contratada poderá aceitar, nas mesmas condições contratuais, os acréscimos ou "
    "supressões que se fizerem necessários, até 25% (vinte e cinco por cento) do valor "
    "inicial atualizado do Contrato, limite que sobe a 50% (cinquenta por cento) para os "
    "acréscimos, no caso particular de reforma de edifício ou de equipamento.")
doc.add_paragraph(
    "14.2. É admissível a fusão, cisão ou incorporação da Contratada com/em outra pessoa "
    "jurídica, desde que sejam observados pela nova pessoa jurídica todos os requisitos de "
    "habilitação exigidos na licitação, sejam mantidas as demais cláusulas e condições do "
    "Contrato, não haja prejuízo à execução do objeto pactuado e haja a anuência expressa "
    "da Conab à continuidade do Contrato.")

# 15. DAS VEDAÇÕES (texto fixo)
add_heading("15. DAS VEDAÇÕES")
doc.add_paragraph(
    "15.1. É vedado à Contratada: a) caucionar ou utilizar o Contrato para qualquer operação "
    "financeira; b) interromper a execução contratual sob alegação de inadimplemento por "
    "parte da Contratante, salvo nos casos previstos em lei; c) empregar menor de 18 "
    "(dezoito) anos em trabalho noturno, perigoso ou insalubre, e menor de 16 (dezesseis) "
    "anos, salvo na condição de aprendiz a partir de 14 (quatorze) anos; e d) subcontratar, "
    "total ou parcialmente, o objeto da contratação, ressalvadas as hipóteses de "
    "subcontratação parcial admitidas neste Termo de Referência.")

# 16. DAS SANÇÕES ADMINISTRATIVAS
add_heading("16. DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "16.1. A contratada, em caso de inadimplemento de suas obrigações, garantido o "
    "contraditório e ampla defesa, ficará sujeita às seguintes sanções previstas no RLC e "
    "na Lei nº 13.303, de 2016: advertência; multa moratória; multa compensatória; multa "
    "rescisória, para os casos de rescisão unilateral; suspensão do direito de participar "
    "de licitação e impedimento de contratar com a Conab, por até 02 (dois) anos.")
doc.add_paragraph(
    "a) advertência, aplicável nas infrações de menor gravidade, sem prejuízo à execução do "
    "objeto;")
add_paragraph_with_bookmarks([
    "b) multa por não regularização da documentação de habilitação: ",
    ("S15_MultaNaoRegularizacaoHabilitacao", "____"), "% sobre o valor estimado.",
])
doc.add_paragraph(
    "b.1) A documentação de habilitação compreende as certidões de regularidade fiscal e "
    "adimplência: SICAF; Cadastro Informativo de Créditos não Quitados do Setor Público "
    "Federal (CADIN); Certidão Negativa de Débitos Trabalhistas (CNDT); Cadastro Nacional de "
    "Empresas Inidôneas e Suspensas (CEIS); e Cadastro Nacional de Condenações Cíveis por Ato "
    "de Improbidade Administrativa e Inelegibilidade do Conselho Nacional de Justiça (CNJ).")
doc.add_paragraph(
    "b.2) Verificada irregularidade em qualquer dessas certidões, será concedido prazo de 5 "
    "(cinco) dias úteis para a regularização das pendências, assegurados o contraditório e a "
    "ampla defesa (art. 460, §1º, do RLC).")
add_paragraph_with_bookmarks([
    "c) multa pelas condutas dos art. 576 a 580 do RLC: ",
    ("S15_MultaCondutasArt576a580", "____"), "% sobre o valor estimado.",
])
add_paragraph_with_bookmarks([
    "d) multa pela recusa em assinar o Contrato: ",
    ("S15_MultaRecusaAssinarContrato", "____"), "% sobre o valor homologado.",
])
doc.add_paragraph(
    "e) multa moratória por atraso injustificado na entrega da garantia contratual, "
    "conforme item 7.2.")
add_paragraph_with_bookmarks([
    "f) multa moratória de ", ("S15_MultaMoratoriaPercentualDia", "____"),
    "% por dia de atraso injustificado, até o limite de ",
    ("S15_MultaMoratoriaLimiteDias", "____"), " dias.",
])
add_paragraph_with_bookmarks([
    "g) multa compensatória de ", ("S15_MultaCompensatoriaParcial", "____"),
    "% sobre o valor da parcela não executada, em caso de inexecução parcial.",
])
add_paragraph_with_bookmarks([
    "h) multa compensatória de ", ("S15_MultaCompensatoriaTotal", "____"),
    "% sobre o valor total do contrato, em caso de inexecução total.",
])
add_paragraph_with_bookmarks([
    "i) multa rescisória de ", ("S15_MultaRescisoria", "____"),
    "% sobre o valor total do contrato, em caso de rescisão unilateral.",
])

# 17. DA RESCISÃO CONTRATUAL (texto fixo)
add_heading("17. DA RESCISÃO CONTRATUAL")
doc.add_paragraph(
    "17.1. A inexecução total do Contrato ensejará a sua rescisão, enquanto a inexecução "
    "parcial poderá ensejar a sua rescisão, com as consequências cabíveis, conforme "
    "disposto nos arts. 568 a 572 do RLC.")
doc.add_paragraph(
    "17.2. A rescisão poderá ser: a) por ato unilateral e escrito da Conab; b) amigável, por "
    "acordo entre as partes, reduzida a termo no processo de contratação, desde que haja "
    "conveniência para a Conab; e c) judicial, por determinação judicial.")
doc.add_paragraph(
    "17.2.1. A rescisão amigável deverá ser precedida de autorização escrita e fundamentada "
    "da autoridade competente.")
doc.add_paragraph(
    "17.2.2. A rescisão amigável não será cabível nos casos em que forem constatados "
    "descumprimentos contratuais sem apuração de responsabilidade iniciada ou com apuração "
    "ainda em curso.")
doc.add_paragraph(
    "17.2.3. Os casos de rescisão contratual serão formalmente motivados, assegurando-se à "
    "Contratada o direito à prévia e ampla defesa, conforme procedimento previsto nos arts. "
    "582 a 593 do RLC.")
doc.add_paragraph(
    "17.3. A rescisão por ato unilateral da Conab acarreta as seguintes consequências, sem "
    "prejuízo das sanções previstas neste Termo de Referência: a) assunção imediata do "
    "objeto contratado, pela Conab, no estado e local em que se encontrar; b) execução da "
    "garantia contratual, quando houver, para ressarcimento pelos eventuais prejuízos "
    "sofridos pela Conab; e c) na hipótese de insuficiência da garantia contratual, quando "
    "houver, a retenção dos créditos decorrentes do Contrato até o limite dos prejuízos "
    "causados à Conab.")
doc.add_paragraph(
    "17.4. A rescisão deverá ser formalizada por termo de rescisão unilateral ou distrato, no "
    "caso de rescisão amigável, devendo o respectivo extrato ser publicado no Diário Oficial "
    "da União.")
doc.add_paragraph(
    "17.5. O termo de rescisão, sempre que possível, será precedido de: a) balanço dos "
    "eventos contratuais já cumpridos ou parcialmente cumpridos; b) relação dos pagamentos já "
    "efetuados e ainda devidos; e c) indenizações e multas.")

# 18. DA VEDAÇÃO AO NEPOTISMO (texto fixo)
add_heading("18. DA VEDAÇÃO AO NEPOTISMO")
doc.add_paragraph(
    "18.1. Conforme disposto no art. 12, §1º, inciso II, do RLC e do artigo 7º do "
    "Decreto nº 7.203, de 2010, fica vedada a contratação de empregado ou dirigente da "
    "Conab como pessoa física, ou de pessoa/empresa com relação de parentesco até o terceiro grau "
    "civil com autoridade do Ministério do Desenvolvimento Agrário e Agricultura Familiar (MDA), dirigente ou "
    "empregado da Conab com atribuições na área de licitação e contratação.")

# 19. DA SUSTENTABILIDADE AMBIENTAL
add_heading("19. DA SUSTENTABILIDADE AMBIENTAL")
doc.add_paragraph(
    "19.1. Compete à Contratada, no que couber, atender os critérios de sustentabilidade "
    "ambiental previstos no art. 10 do RLC.")
doc.add_paragraph(
    "19.2. A Contratada se responsabiliza administrativamente, civilmente e penalmente por "
    "qualquer dano causado pelo seu produto ao meio ambiente.")
add_paragraph_with_bookmarks(["19.3. ", ("S18_CriteriosEspecificos", "")])

# 20. DA MATRIZ DE RISCOS E RESPONSABILIDADES (texto fixo)
add_heading("20. DA MATRIZ DE RISCOS E RESPONSABILIDADES")
doc.add_paragraph(
    "20.1. A Matriz de Riscos é a cláusula contratual definidora dos riscos e das "
    "responsabilidades entre o CONTRATANTE e a CONTRATADA, conforme hipóteses elencadas na "
    "Matriz de Riscos - Anexo I do Termo de Referência.")

# 21. DAS DISPOSIÇÕES FINAIS
add_heading("21. DAS DISPOSIÇÕES FINAIS")
doc.add_paragraph(
    "21.1. Integram este Termo de Referência os seguintes Anexos: a) Anexo I - Matriz de "
    "Riscos; e b) Anexo II - Modelo de Apresentação de Proposta.")
add_paragraph_with_bookmarks(["Elaborador — Área Demandante: ", ("S20_ElabDemandanteNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S20_ElabDemandanteSetor", ""), " — Cargo: ", ("S20_ElabDemandanteCargo", "")])
add_paragraph_with_bookmarks(["Elaborador — Área Técnica: ", ("S20_ElabTecnicaNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S20_ElabTecnicaSetor", ""), " — Cargo: ", ("S20_ElabTecnicaCargo", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Demandante: ", ("S20_AprovDemandanteNome", ""), " — Setor: ", ("S20_AprovDemandanteSetor", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Técnica: ", ("S20_AprovTecnicaNome", ""), " — Setor: ", ("S20_AprovTecnicaSetor", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Demandante: ", ("S20_SuperDemandanteNome", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Técnica: ", ("S20_SuperTecnicaNome", "")])

# ANEXO I - MATRIZ DE RISCOS (Anexo II do RLC, transcrito integralmente)
doc.add_page_break()
add_heading("ANEXO I DO TERMO DE REFERÊNCIA")
add_heading("MATRIZ DE RISCOS")
add_matriz_de_riscos_rlc("S01_DescricaoObjeto_AnexoI")

# ANEXO II - MODELO DE APRESENTAÇÃO DE PROPOSTA (limpo, oficial)
doc.add_page_break()
add_heading("ANEXO II DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE APRESENTAÇÃO DE PROPOSTA")
doc.add_paragraph(
    "A proposta deverá conter: descrição detalhada do objeto, quantidades, valor unitário e "
    "valor total, em algarismos e por extenso; prazo de entrega e condições de entrega; "
    "prazo de vigência do fornecimento; prazo de garantia oferecido, não inferior a 60 "
    "(sessenta) dias corridos; prazo de validade da proposta, não inferior a 60 (sessenta) "
    "dias corridos; informação de que os preços cotados incluem todos os custos necessários "
    "(impostos, tributos, encargos trabalhistas, previdenciários, comerciais e fiscais); e os "
    "dados da proponente (razão social, endereço, telefone, CNPJ, dados bancários).")
tabela2 = doc.add_table(rows=1, cols=6)
tabela2.style = "Table Grid"
for celula, texto in zip(tabela2.rows[0].cells, ["Item", "Especificações", "CATMAT", "Quantidade", "Valor Unitário", "Valor Total"]):
    celula.text = texto
for _ in range(5):
    tabela2.add_row()

doc.save("Template_TR_Bens.docx")
print("Template_TR_Bens.docx gerado com sucesso.")
