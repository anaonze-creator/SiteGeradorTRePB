"""Gera Template_TR_Servicos.docx a partir do texto oficial da minuta
SEI 55379668, com bookmarks nos pontos de preenchimento identificados em
js/dados-tr-servicos.js. Rodar novamente sempre que o texto oficial mudar."""
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


# Capa
doc.add_paragraph("COMPANHIA NACIONAL DE ABASTECIMENTO")
titulo = doc.add_paragraph()
titulo_run = titulo.add_run("TERMO DE REFERÊNCIA")
titulo_run.bold = True

# 1. DO OBJETO
add_heading("1. DO OBJETO")
add_paragraph_with_bookmarks([
    "1.1. O presente Termo de Referência tem por objeto a contratação de ",
    ("S01_DescricaoObjeto", "(descrição do serviço a ser contratado)"),
    ", conforme especificações, condições, quantidades e exigências estabelecidas neste instrumento.",
])
add_paragraph_with_bookmarks([
    "1.2. Os serviços serão executados indiretamente no regime de ",
    ("S01_RegimeExecucao", "(regime de execução)"),
    ", conforme o inciso IV, art. 208 do RLC.",
])
add_paragraph_with_bookmarks([
    "1.3. O custo total estimado para a prestação dos serviços deste Termo de Referência é de R$ ",
    ("S01_ValorEstimado", "________"),
    ". O critério utilizado para a formação da estimativa foi de ",
    ("S01_MetodologiaEstimativa", "(MÉDIA, MEDIANA OU MENOR PREÇO)"),
    " dos preços praticados no mercado.",
])
add_paragraph_with_bookmarks(["Metodologia detalhada: ", ("S01_MetodologiaEstimativaOutro", "____")])
add_bloco_variante("S01_CriterioJulgamentoSRP", {
    "Não": [[
        "1.4. O critério de julgamento das propostas será o de ",
        ("S01_CriterioJulgamento", "[menor preço ou maior desconto]"),
        " ofertado. Para fins de julgamento das propostas o valor estimado/valor de referência será ",
        ("S01_SigiloValor", "(público ou sigiloso)"),
        ". O modo de disputa adotado no certame será ",
        ("S01_ModoDisputa", "(aberto ou aberto e fechado)"),
        ".",
    ]],
    "Sim": [[
        "1.4. O critério de julgamento das propostas será o de menor preço ofertado, "
        "por se tratar de licitação para Sistema de Registro de Preços (art. 328 do RLC). "
        "Para fins de julgamento das propostas o valor estimado/valor de referência será ",
        ("S01_SigiloValor", "(público ou sigiloso)"),
        ". O modo de disputa adotado no certame será ",
        ("S01_ModoDisputa", "(aberto ou aberto e fechado)"),
        ".",
    ]],
})
add_paragraph_with_bookmarks([
    "1.5. Os recursos orçamentários decorrentes da contratação do objeto deste Termo de Referência "
    "estão consignados no Orçamento da Conab para o ano 20XX, conforme segue: ND: (",
    ("S01_ND", "XXXX"), "), PTRES: (", ("S01_PTRES", "XXXX"), "), FONTE: (", ("S01_Fonte", "XXXX"), ").",
])
add_bloco_variante("S01_SRP", {
    "Sim": [
        "1.6. Na licitação para registro de preços não é necessária a indicação de dotação "
        "orçamentária, que somente será exigida quando da formalização do Contrato ou "
        "instrumento equivalente (art. 331 do RLC)."
    ],
    "Não": [""],
})
add_bloco_variante("S01_CriterioJulgamentoSRP", {
    "Sim": [["1.6.1. Justificativa da hipótese do art. 327 do RLC que ampara o uso do SRP: ",
             ("S01_JustificativaSRP", "____")]],
    "Não": [""],
})
add_bloco_variante("S01_TIC", {
    "Sim": [
        "1.7. Por se tratar de contratação de Tecnologia da Informação, esta contratação "
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
    "2.2. O serviço em apreço será contratado na modalidade Pregão Eletrônico, por ser "
    "caracterizado como serviço comum, na forma do Inciso XIII do art. 3º do Regulamento de "
    "Licitações e Contratos da Conab - RLC.")
add_bloco_variante("S02_AplicaExcecaoMEEPP", {
    "Sim": [
        "2.3. Em que pese o valor estimado da contratação ser inferior ao limite de "
        "R$ 80.000,00 (oitenta mil reais), o certame não será realizado com exclusividade de "
        "participação para microempresas e empresas de pequeno porte, na forma da Lei "
        "Complementar 123/2006 e do Decreto 8.538/2015, em razão da contratação em apreço se "
        "enquadrar na exceção prevista no art. 49 da Lei nº 123/2006, isto é: ",
    ],
    "Não": [""],
})
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
add_paragraph_with_bookmarks(["3.1. ", ("S03_Objetivo", "(objetivo da contratação)")])

# 4. DA VIGÊNCIA DO CONTRATO
add_heading("4. DA VIGÊNCIA DO CONTRATO")
add_bloco_variante("S04_TipoVigencia", {
    "Até 1 ano, prorrogável": [[
        "4.1. O prazo de vigência do Contrato será de ", ("S04_PrazoVigenciaMeses_Inline", "____"),
        " meses (duração do Contrato, "
        "conforme arts. 461 a 463 do RLC), contados da data de sua assinatura, prorrogável na "
        "forma dos arts. 497 e 498 do RLC, desde que observados os requisitos do art. 488 do "
        "RLC, entre eles a demonstração da vantajosidade na manutenção do ajuste (em regra, "
        "por pesquisa de preços), sob pena de a Conab não poder prorrogar o Contrato.",
    ]],
    "Plurianual (até 5 anos)": [[
        "4.1. O prazo de vigência do Contrato será de ", ("S04_PrazoVigenciaAnos_Inline", "____"),
        " anos (duração do Contrato, "
        "conforme arts. 461 a 462 do RLC), contados da data de sua assinatura, podendo ser "
        "prorrogado até o limite de 05 (cinco) anos, desde que observados os requisitos "
        "previstos no art. 488 do RLC, entre eles a demonstração da vantajosidade na "
        "manutenção do ajuste (em regra, por pesquisa de preços), sob pena de a Conab não "
        "poder prorrogar o Contrato.",
    ]],
})
add_paragraph_with_bookmarks(["Prazo de vigência (meses): ", ("S04_PrazoVigenciaMeses", "____")])
add_paragraph_with_bookmarks(["Prazo de vigência (anos): ", ("S04_PrazoVigenciaAnos", "____")])
add_paragraph_with_bookmarks([
    "4.2. Quando cabível, a prorrogação poderá se estender por ",
    ("S04_PrazoProrrogacaoMeses", "____"), ".",
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
    "5.1.2. a) Comprovação de aptidão para a prestação dos serviços em características, "
    "quantidades e prazos compatíveis com o objeto desta licitação, por período não inferior a ",
    ("S05_PrazoExperienciaMinimaAnos", "____"), " anos, mediante atestados fornecidos por "
    "pessoas jurídicas de direito público ou privado.",
])
add_paragraph_with_bookmarks([
    "a.1) Para fins da compatibilidade mencionada na alínea \"a\": ",
    ("S05_QuantitativoServicoPrincipal", "(quantitativo mínimo e serviço(s) principal(is))"),
])
add_bloco_variante("S05_ExigeInstalacaoEscritorio", {
    "Sim": [
        "5.1.3. Declaração de instalação de escritório: para a realização do objeto da "
        "licitação, a Contratada deverá entregar declaração de "
        "que instalará escritório na cidade em que está localizada a Matriz ou a Sureg da "
        "Conab, a ser comprovado no prazo máximo de 60 (sessenta) dias contados a partir da "
        "vigência do contrato.",
    ],
    "Não": [""],
})
add_paragraph_with_bookmarks(["Cidade do escritório: ", ("S05_MunicipioInstalacaoEscritorio", "")])
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
doc.add_paragraph(
    "5.1.7. Será aceito o somatório de atestados para fins de comprovação de capacidade "
    "técnica. A licitante vencedora, apenas por ocasião da assinatura do Contrato, deverá "
    "comprovar inscrição na entidade profissional competente, quando for o caso, e apresentar "
    "declaração de que é Microempresa, Empresa de Pequeno Porte ou Cooperativa, se for o caso.")
doc.add_paragraph(
    "5.1.8. Como condição prévia ao exame da documentação de habilitação, será verificado o "
    "eventual descumprimento das condições de participação mediante consulta ao SICAF, ao "
    "Cadastro Nacional de Empresas Inidôneas e Suspensas (CEIS), ao Cadastro Nacional de "
    "Condenações Cíveis por Atos de Improbidade Administrativa e à Lista de Inidôneos do TCU, "
    "em nome da licitante e de seu sócio majoritário.")
doc.add_paragraph(
    "5.1.9. Será inabilitada a licitante que não comprovar sua habilitação, ressalvado o "
    "saneamento de erros ou falhas que não alterem a substância das propostas e dos "
    "documentos, mediante despacho fundamentado. Será considerada vencedora a licitante que "
    "oferecer o menor preço e atender às exigências de habilitação deste Termo de Referência.")
doc.add_paragraph("5.2. DA VISTORIA")
add_bloco_variante("S05_TipoVistoria", {
    "Obrigatória": [
        "5.2.1. Para o correto dimensionamento e elaboração de sua proposta, o licitante "
        "deverá realizar vistoria nas instalações do local de execução dos serviços, "
        "acompanhado por empregado designado para esse fim, devendo o agendamento ser "
        "efetuado previamente, podendo sua realização ser comprovada por declaração assinada "
        "pelo licitante de que realizou a vistoria e tem pleno conhecimento das condições e "
        "peculiaridades do local, na forma do Anexo III deste TR.",
    ],
    "Facultativa": [
        "5.2.1. Para o correto dimensionamento e elaboração de sua proposta, recomenda-se que "
        "o licitante realize vistoria nas instalações do local de execução dos serviços, "
        "sendo a realização facultativa. Optando por realizá-la, deverá estar acompanhado por "
        "empregado designado para esse fim, devendo o agendamento ser "
        "efetuado previamente, podendo sua realização ser comprovada por declaração assinada "
        "pelo licitante de que realizou a vistoria e tem pleno conhecimento das condições e "
        "peculiaridades do local, na forma do Anexo III deste TR.",
        "O licitante que optar por não realizar a vistoria deverá apresentar declaração de que "
        "tem pleno conhecimento das condições e peculiaridades do local, assumindo total "
        "responsabilidade por sua opção, na forma do Anexo III deste TR.",
    ],
})
add_paragraph_with_bookmarks([
    "5.2.1.1. O agendamento da vistoria deverá ser feito de segunda a sexta-feira, das ",
    ("S05_HorarioVistoriaInicio", "____"), " às ", ("S05_HorarioVistoriaFim", "____"),
    " horas, pelo telefone (", ("S05_TelefoneAgendamentoVistoria", "____"), ").",
])
doc.add_paragraph(
    "5.2.2. O prazo para vistoria iniciar-se-á no dia útil seguinte ao da publicação do "
    "Edital, estendendo-se até o dia útil anterior à data prevista para a abertura da sessão "
    "pública.")

# 6. DA DESCRIÇÃO DOS SERVIÇOS
add_heading("6. DA DESCRIÇÃO DOS SERVIÇOS")
doc.add_paragraph(
    "6.1. A prestação dos serviços não gera vínculo empregatício entre os empregados da "
    "Contratada e a Conab, vedando-se qualquer relação que caracterize pessoalidade e "
    "subordinação direta.")
doc.add_paragraph("6.2. DO DETALHAMENTO DOS SERVIÇOS")
add_paragraph_with_bookmarks(["6.2.1. ", ("S06_DetalhamentoServicos", "")])
doc.add_paragraph("6.3. DO LOCAL E PRAZO DE EXECUÇÃO")
add_paragraph_with_bookmarks([
    "6.3.1. Os serviços objeto deste Termo de Referência deverão ser realizados em ",
    ("S06_LocalExecucao", "(local)"), ", de segunda a sexta-feira, no horário de ",
    ("S06_HorarioExecucao", "(horário)"), ".",
])
_bid_enderecos = add_bloco_variante_sentinela_inicio("S06_EnderecosAdicionais_CONTAINER")
doc.add_paragraph(
    "6.3.2. Quando o objeto envolver a execução em mais de um endereço, os locais adicionais "
    "são os seguintes:")
tabela_enderecos = add_tabela(["Endereço", "Observação"])
add_linha_modelo_tabela(tabela_enderecos, "S06_EnderecosAdicionais_LINHA_MODELO", ["((ENDERECO))", "((OBSERVACAO))"])
add_bloco_variante_sentinela_fim(_bid_enderecos)
add_paragraph_with_bookmarks([
    "6.3.3. O prazo de execução dos serviços será de ", ("S06_PrazoExecucaoDias", "____"),
    " dias, contados a partir da assinatura do Contrato.",
])
doc.add_paragraph("6.4. DA ROTINA DE EXECUÇÃO DOS SERVIÇOS")
add_paragraph_with_bookmarks(["6.4.1. ", ("S06_RotinaExecucao", "")])
doc.add_paragraph("6.5. DA GARANTIA DOS SERVIÇOS")
add_bloco_variante("S06_ExigeGarantiaServicos", {
    "Sim": ["6.5.1. Os serviços deverão ter um período de garantia de, no mínimo, ____ meses, "
            "contados do recebimento definitivo pela Contratante."],
    "Não": ["6.5.1. Não será exigida garantia dos serviços."],
})
add_paragraph_with_bookmarks(["Período de garantia (meses): ", ("S06_GarantiaServicosMeses", "____")])
add_paragraph_with_bookmarks([
    "6.5.2. Durante o período de garantia, a Contratada deverá refazer o serviço, no prazo "
    "máximo de ", ("S06_PrazoRefazimentoDias", "____"), " dias, após ser comunicada.",
])
doc.add_paragraph("6.6. DOS MATERIAIS A SEREM DISPONIBILIZADOS")
add_bloco_variante("S06_ExigeMateriaisDisponibilizados", {
    "Sim": ["6.6.1. Para a perfeita execução dos serviços, a Contratada deverá disponibilizar "
            "os materiais, equipamentos, ferramentas e utensílios necessários, promovendo sua "
            "substituição quando necessário."],
    "Não": [""],
})
add_paragraph_with_bookmarks(["", ("S06_MateriaisDisponibilizados", "")])
doc.add_paragraph("6.7. DO INSTRUMENTO DE MEDIÇÃO DE RESULTADOS")
doc.add_paragraph(
    "6.7.1. Durante a vigência do Contrato, a Conab, para avaliar a qualidade na execução dos "
    "serviços, adotará o Instrumento de Medição de Resultados (IMR), constante no Anexo V "
    "deste Termo de Referência.")
doc.add_paragraph("6.8. DO RECEBIMENTO DOS SERVIÇOS")
doc.add_paragraph(
    "6.8.1. O recebimento provisório será realizado pela fiscalização; o recebimento "
    "definitivo, pelo empregado ou Comissão designada, nos termos do art. 526 e 527 do RLC.")
add_paragraph_with_bookmarks([
    "6.8.2. Prazo para recebimento provisório: ", ("S06_PrazoRecebimentoProvisorioDias", "____"),
    " dias. Prazo para recebimento definitivo, contado do provisório: ",
    ("S06_PrazoRecebimentoDefinitivoDias", "____"), " dias.",
])

# 7. DA GARANTIA CONTRATUAL
add_heading("7. DA GARANTIA CONTRATUAL")
add_bloco_variante("S07_ExigeGarantia", {
    "Sim": [
        [
            "7.1. A Contratada deverá apresentar, no prazo de 10 (dez) dias corridos após a "
            "assinatura do Contrato, prorrogáveis por igual período, comprovante de prestação de "
            "garantia mediante caução em dinheiro, seguro-garantia ou fiança bancária, cabendo à "
            "Contratada optar pela modalidade (art. 439, §1º, do RLC), correspondente a ",
            ("S07_PercentualGarantia_Inline", "____"),
            "%, conforme parâmetros do art. 439 do RLC, do ",
            ("S07_BaseCalculoGarantia", "(valor do Contrato ou valor anual do Contrato)"),
        ],
        "7.2. Nos casos de contratação de serviços com dedicação exclusiva de mão de obra, o "
        "valor da garantia deverá corresponder a 5% do valor do contrato, limitado ao "
        "equivalente a 2 (dois) meses do custo da folha de pagamento dos empregados da "
        "contratada que venham a participar da execução dos serviços contratados (art. 439, "
        "§4º, do RLC).",
    ],
    "Não": ["7.1. Não será exigida garantia contratual."],
})
add_paragraph_with_bookmarks([
    "Percentual da garantia: ", ("S07_PercentualGarantia", "____"), "% do valor do Contrato "
    "— em regra, até 5% (art. 439, §2º, do RLC), podendo ser elevado para até 10% em "
    "contratações de grande vulto envolvendo complexidade técnica e riscos financeiros "
    "elevados (art. 439, §3º, do RLC).",
])
add_paragraph_with_bookmarks([
    "7.3. A inobservância do prazo fixado para apresentação da garantia acarretará a "
    "aplicação de multa de 0,08% (oito centésimos por cento) do ",
    ("S07_BaseCalculoGarantia_74", "(valor do Contrato ou valor anual do Contrato)"),
    " por dia de atraso, até o máximo de 2% (dois por cento). O atraso superior a 25 (vinte "
    "e cinco) dias autoriza a rescisão do Contrato, conforme o inciso I do art. 569 do RLC. "
    "A garantia deverá ter validade de 90 (noventa) dias após o término da vigência do "
    "Contrato.",
])

# 8. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO
add_heading("8. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO")
doc.add_paragraph(
    "8.1. Para a execução do ajuste, será adotado o método de trabalho baseado no conceito de "
    "delegação de responsabilidade. Nos termos dos arts. 543 e 544 do RLC, será designado "
    "Fiscal, seu substituto, ou Comissão de Fiscalização para acompanhar e fiscalizar a "
    "execução dos serviços.")
add_paragraph_with_bookmarks(["8.2. Rotinas de fiscalização específicas: ", ("S08_RotinasFiscalizacaoEspecificas", "")])

# 9. DAS OBRIGAÇÕES DA CONTRATANTE (texto fixo)
add_heading("9. DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph("9.1. São obrigações da Contratante:")
doc.add_paragraph(
    "a) exigir o cumprimento de todas as obrigações assumidas pela Contratada, de acordo "
    "com as cláusulas contratuais e os termos de sua proposta;")
doc.add_paragraph(
    "b) receber o objeto no prazo e condições estabelecidas no Edital e seus anexos;")
doc.add_paragraph(
    "c) verificar minuciosamente, no prazo fixado, a conformidade dos serviços recebidos "
    "provisoriamente com as especificações constantes do Edital e da proposta, para fins de "
    "aceitação e recebimento definitivo;")
doc.add_paragraph(
    "d) comunicar à Contratada, por escrito, sobre imperfeições, falhas ou irregularidades "
    "verificadas no curso da execução dos serviços, para que seja reparado ou corrigido;")
doc.add_paragraph(
    "e) rejeitar, no todo ou em parte, serviço executado em desacordo com o previsto neste "
    "Termo de Referência;")
doc.add_paragraph(
    "f) acompanhar e fiscalizar o cumprimento das obrigações da Contratada, através de "
    "empregado ou comissão especialmente designado;")
doc.add_paragraph(
    "g) efetuar o pagamento à Contratada no valor correspondente a prestação de serviço, no "
    "prazo e forma estabelecidos no Edital e seus anexos.")
doc.add_paragraph(
    "9.2. A Conab não responderá por quaisquer compromissos assumidos pela Contratada com "
    "terceiros, ainda que vinculados à execução do Contrato, bem como por qualquer dano "
    "causado a terceiros em decorrência de ato da Contratada, de seus empregados, prepostos "
    "ou subordinados.")

# 10. DAS OBRIGAÇÕES DA CONTRATADA
add_heading("10. DAS OBRIGAÇÕES DA CONTRATADA")
doc.add_paragraph(
    "10.1. A Contratada deve cumprir todas as obrigações constantes neste Termo, seus anexos "
    "e sua proposta, assumindo como exclusivamente seus os riscos e as despesas decorrentes da "
    "boa e perfeita execução do objeto e, ainda:")
doc.add_paragraph(
    "a) executar os serviços conforme especificações deste Termo de Referência e de sua "
    "proposta, com a alocação dos empregados necessários ao perfeito cumprimento das "
    "cláusulas contratuais, além de fornecer os materiais e equipamentos, ferramentas e "
    "utensílios necessários, na qualidade e quantidade especificadas neste Termo de "
    "Referência e em sua proposta;")
doc.add_paragraph(
    "b) responsabilizar-se pelos vícios e danos decorrentes do objeto, de acordo com os arts. "
    "12, 13 e 17 a 27, do Código de Defesa do Consumidor (Lei nº 8.078, de 1990);")
doc.add_paragraph(
    "c) substituir, reparar, remover ou corrigir, às suas expensas, no total ou em parte, no "
    "prazo estipulado pela fiscalização, os serviços efetuados em que se verificarem vícios, "
    "defeitos ou incorreções resultantes da execução ou dos materiais empregados;")
doc.add_paragraph(
    "d) comunicar à Contratante, no prazo máximo de 24 (vinte e quatro) horas que antecede a "
    "data da finalização dos serviços, os motivos que impossibilitem o cumprimento do prazo "
    "previsto, com a devida comprovação;")
doc.add_paragraph(
    "e) utilizar empregados habilitados e com conhecimentos básicos dos serviços a serem "
    "executados, em conformidade com as normas e determinações em vigor;")
doc.add_paragraph(
    "f) apresentar os empregados devidamente uniformizados e identificados por meio de "
    "crachá, além de provê-los com os Equipamentos de Proteção Individual - EPI, quando for "
    "o caso;")
doc.add_paragraph(
    "g) apresentar à Contratante, quando for o caso, a relação nominal dos empregados que "
    "adentrarão a Conab para a execução do serviço;")
doc.add_paragraph(
    "h) responsabilizar-se por todas as obrigações trabalhistas, sociais, previdenciárias, "
    "tributárias e as demais previstas na legislação específica, cuja inadimplência não "
    "transfere responsabilidade à Contratante;")
doc.add_paragraph(
    "i) atender as solicitações da Contratante quanto à substituição dos empregados "
    "alocados, no prazo fixado pelo fiscal do Contrato, nos casos em que ficar constatado "
    "descumprimento das obrigações relativas à execução do serviço, conforme descrito neste "
    "Termo de Referência;")
doc.add_paragraph(
    "j) instruir seus empregados quanto à necessidade de acatar as normas internas da "
    "Conab;")
doc.add_paragraph(
    "k) instruir seus empregados a respeito das atividades a serem desempenhadas, "
    "alertando-os a não executar atividades não abrangidas pelo Contrato, devendo a "
    "Contratada relatar à Contratante toda e qualquer ocorrência neste sentido, a fim de "
    "evitar desvio de função;")
doc.add_paragraph(
    "l) relatar à Contratante toda e qualquer irregularidade verificada no decorrer da "
    "prestação dos serviços;")
doc.add_paragraph(
    "m) não permitir a utilização de qualquer trabalho do menor de dezesseis anos, exceto na "
    "condição de aprendiz para os maiores de quatorze anos; nem permitir a utilização do "
    "trabalho do menor de dezoito anos em trabalho noturno, perigoso ou insalubre;")
doc.add_paragraph(
    "n) guardar sigilo sobre todas as informações obtidas em decorrência do cumprimento do "
    "Contrato;")
doc.add_paragraph(
    "o) arcar com o ônus decorrente de eventual equívoco no dimensionamento dos "
    "quantitativos de sua proposta, devendo complementá-los, caso o previsto inicialmente em "
    "sua proposta não seja satisfatório para o atendimento ao objeto da licitação, exceto "
    "quando ocorrer algum dos eventos arrolados no Art. 497 do RLC da Conab;")
doc.add_paragraph(
    "p) manter, durante toda a execução do Contrato, em compatibilidade com as obrigações "
    "assumidas, todas as condições de habilitação e qualificação exigidas na licitação;")
doc.add_paragraph(
    "q) prestar todos os esclarecimentos que forem solicitados pela Conab, durante a "
    "realização do Contrato;")
doc.add_paragraph(
    "r) indicar preposto para representá-la durante a execução do Contrato;")
doc.add_paragraph(
    "s) garantir aos seus empregados, na execução do Contrato, política de respeito à "
    "diversidade, prevenção a práticas discriminatórias e alinhamento com os princípios "
    "institucionais de equidade de gênero e raça.")
doc.add_paragraph("Declaração de instalação de escritório em município(s) específico(s):")
add_bloco_variante("S10_ExigeEscritorioMunicipio", {
    "Sim": [
        "A Contratada deverá entregar declaração de que instalará escritório no(s) "
        "município(s) indicado(s), a ser comprovado no prazo máximo de 60 (sessenta) dias "
        "contados a partir da vigência do Contrato, dispondo de capacidade operacional para "
        "receber e solucionar qualquer demanda da Contratante.",
    ],
    "Não": [""],
})
add_paragraph_with_bookmarks(["Município(s): ", ("S10_MunicipiosEscritorio", "")])
doc.add_paragraph(
    "t) deter instalações, aparelhamento e pessoal técnico adequados e disponíveis para a "
    "realização do objeto da licitação.")
add_bloco_variante("S10_NaturezaIntelectual", {
    "Sim": [
        "u) ceder os direitos patrimoniais relativos ao projeto ou serviço técnico "
        "especializado, para que a Conab possa utilizá-lo, de acordo com o art. 525 do RLC. "
        "Quando o projeto referir-se a obra imaterial de caráter tecnológico, insuscetível de "
        "privilégio, a cessão dos direitos incluirá o fornecimento de todos os dados, "
        "documentos e elementos de informação pertinentes à tecnologia de concepção, "
        "desenvolvimento, fixação em suporte físico e aplicação da obra;",
        "v) assegurar à Contratante o direito de propriedade intelectual dos produtos "
        "desenvolvidos, inclusive sobre adequações e atualizações futuras, de forma "
        "permanente, e os direitos autorais da solução, do projeto, de suas especificações "
        "técnicas e da documentação produzida, inclusive os produzidos por terceiros "
        "subcontratados, ficando proibida sua utilização sem autorização expressa da "
        "Contratante, sob pena de multa, sem prejuízo das sanções civis e penais cabíveis;",
        "w) realizar a transição contratual com transferência de conhecimento, tecnologia e "
        "técnicas empregadas, sem perda de informações, podendo exigir, inclusive, a "
        "capacitação dos técnicos da Contratante ou da nova empresa que continuará a execução "
        "dos serviços, nos termos do art. 17 do RLC.",
    ],
    "Não": [""],
})
add_bloco_variante("S10_ExigeProgramaIntegridade", {
    "Sim": [[
        "x) comprovar a implementação do Programa de Integridade, nos moldes do art. 603 e "
        "Anexo XXXII do RLC, no prazo de até ", ("S10_PrazoProgramaIntegridadeMeses", "____"),
        " meses após a assinatura do contrato.",
    ]],
    "Não": [""],
})

# 11. DO PAGAMENTO
add_heading("11. DO PAGAMENTO")
doc.add_paragraph(
    "11.1. O pagamento será efetuado pela Conab no prazo de até 10 (dez) dias corridos, "
    "contados do recebimento da Nota Fiscal/Fatura.")
doc.add_paragraph(
    "11.2. A emissão da Nota Fiscal/Fatura será precedida do recebimento provisório e "
    "definitivo do serviço, nos seguintes termos: a) no prazo de até 05 (cinco) dias corridos "
    "do adimplemento da parcela, a Contratada deverá entregar toda a documentação "
    "comprobatória do cumprimento da obrigação contratual;")
add_bloco_variante("S08_TipoFiscalizacao", {
    "Fiscal Funcional": [
        "b) no prazo de até 07 (sete) dias corridos a partir do recebimento dos documentos da "
        "Contratada, o Fiscal Funcional deverá realizar a análise de toda a documentação "
        "apresentada;",
        "c) no prazo de até 03 (três) dias úteis a partir da análise da documentação, o Fiscal "
        "Funcional avaliará a execução por meio do Instrumento de Medição de Resultado (IMR) e "
        "emitirá Termo de Recebimento Provisório, com detalhamento da execução contratual, "
        "encaminhando-o ao empregado ou comissão designada para o recebimento definitivo.",
    ],
    "Comissão de Fiscalização": [
        "b) no prazo de até 07 (sete) dias corridos a partir do recebimento dos documentos da "
        "Contratada:",
        "b.1) o Fiscal Funcional deverá realizar a análise de toda a documentação apresentada "
        "e emitir relatório, com detalhamento da execução contratual;",
        "b.2) o Fiscal Técnico deverá realizar a análise de toda a documentação apresentada e "
        "emitir relatório, com detalhamento da execução contratual;",
        "b.3) o Fiscal Setorial deverá realizar a análise de toda a documentação apresentada e "
        "emitir relatório, com detalhamento da execução contratual;",
        "c) o Fiscal Funcional, no prazo de até 03 (três) dias úteis a partir do recebimento do "
        "último relatório, avaliará a execução por meio do Instrumento de Medição de Resultado "
        "(IMR) e emitirá Termo de Recebimento Provisório com base nos relatórios emitidos pelos "
        "fiscais, encaminhando-o, caso não haja irregularidades, ao empregado ou comissão "
        "designada para o recebimento definitivo.",
    ],
})
doc.add_paragraph(
    "d) constatadas impropriedades ou irregularidades que impeçam a liquidação e o pagamento "
    "da despesa, estas serão registradas no Termo de Recebimento Provisório, com prazo de até "
    "05 (cinco) dias úteis para regularização; e) sanadas as impropriedades, o Fiscal Funcional "
    "ou a Comissão de Fiscalização, em 05 (cinco) dias corridos, elaborará relatório detalhado "
    "e encaminhará o Termo de Recebimento Provisório ao empregado ou comissão designada para o "
    "recebimento definitivo; f) no prazo de até 05 (cinco) dias corridos, o empregado ou "
    "Comissão designada providenciará o recebimento definitivo, emitindo Termo Circunstanciado "
    "e comunicando a Contratada para emissão da Nota Fiscal ou Fatura, com o valor exato "
    "dimensionado pela fiscalização com base no IMR.")
doc.add_paragraph(
    "11.3. Antes de cada pagamento será realizada consulta ao SICAF; havendo resultado "
    "desfavorável, será concedido prazo de 05 (cinco) dias úteis para regularização ou defesa, "
    "prorrogável uma vez por igual período.")

# 12. DO REAJUSTE
add_heading("12. DO REAJUSTE")
add_bloco_variante("S12_TipoCorrecao", {
    "Preço fixo (irreajustável)": ["12.1. O preço é fixo e irreajustável."],
    "Reajuste (sem dedicação exclusiva de mão de obra)": [[
        "12.1. O preço consignado no Contrato será corrigido anualmente, observado o "
        "interregno mínimo de um ano, contado da data limite para apresentação da "
        "proposta, pela variação do índice de referência: ",
        ("S12_IndiceReajuste", "____"), ".",
    ]],
    "Repactuação (com dedicação exclusiva de mão de obra)": [[
        "12.1. Por se tratar de serviço com dedicação exclusiva de mão de obra, o preço "
        "consignado no Contrato será objeto de repactuação (art. 3º, LXXXIII, e art. 501 do "
        "RLC), observado o interregno mínimo de um ano contado da data-base da proposta "
        "para os custos decorrentes do mercado, e do Acordo, Convenção ou Dissídio "
        "Coletivo de Trabalho para os custos decorrentes da mão de obra. Acordo, Convenção "
        "ou Dissídio Coletivo de Trabalho de referência: ",
        ("S12_AcordoConvencaoDissidio", "____"), ".",
    ]],
})
doc.add_paragraph(
    "12.2. O reajuste ou a repactuação pressupõe interregno mínimo de um ano contado da data "
    "limite para apresentação da proposta, só sendo cabível, portanto, quando a vigência do "
    "Contrato — inclusive prorrogações — alcançar 12 (doze) meses.")
doc.add_paragraph(
    "12.3. Serão objeto de preclusão os reajustes ou repactuações a que a Contratada fizer "
    "jus durante a vigência do Contrato e que não forem solicitados até o implemento dos "
    "seguintes eventos: a) assinatura do termo aditivo de prorrogação contratual; b) data em "
    "que o Contrato completa 12 (doze) meses, 24 (vinte e quatro) meses e assim "
    "sucessivamente; ou c) encerramento do Contrato.")
doc.add_paragraph(
    "12.4. O disposto nesta seção não afasta o direito ao reequilíbrio econômico-financeiro "
    "em sentido estrito, cabível a qualquer tempo, independentemente de previsão contratual, "
    "quando verificado fato imprevisível ou de consequências incalculáveis, anormal ou "
    "extraordinário, nos termos do art. 511 do RLC.")

# 13. DA SUBCONTRATAÇÃO
add_heading("13. DA SUBCONTRATAÇÃO")
add_bloco_variante("S13_PermiteSubcontratacao", {
    "Sim": ["13.1. É permitida a subcontratação parcial do objeto entre os limites mínimo e "
            "máximo, respectivamente, do valor total do Contrato."],
    "Não": ["13.1. Não será admitida a subcontratação do objeto licitatório."],
})
add_paragraph_with_bookmarks([
    "Limites: entre ", ("S13_PercentualMinSubcontratacao", "____"), "% e ",
    ("S13_PercentualMaxSubcontratacao", "____"), "% do valor total do Contrato.",
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
    "jurídica, desde que observados os requisitos de habilitação, mantidas as demais "
    "cláusulas e condições do Contrato, não haja prejuízo à execução do objeto pactuado e "
    "haja a anuência expressa da Conab.")

# 15. DAS VEDAÇÕES (texto fixo)
add_heading("15. DAS VEDAÇÕES")
doc.add_paragraph(
    "15.1. É vedado à Contratada: a) caucionar ou utilizar o Contrato para qualquer operação "
    "financeira; b) interromper a execução contratual sob alegação de inadimplemento por "
    "parte da Contratante, salvo nos casos previstos em lei; c) empregar menor de 18 "
    "(dezoito) anos em trabalho noturno, perigoso ou insalubre, e menor de 16 (dezesseis) "
    "anos, salvo na condição de aprendiz a partir de 14 (quatorze) anos; d) subcontratar, "
    "total ou parcialmente, o objeto da contratação, ressalvadas as hipóteses de "
    "subcontratação parcial admitidas neste Termo de Referência; e e) solicitar, por meio de "
    "seus empregados, serviços, materiais ou equipamentos a empresas terceirizadas que "
    "prestem serviços à Conab (art. 155 do RLC).")

# 16. DAS SANÇÕES ADMINISTRATIVAS
add_heading("16. DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "16.1. A contratada, em caso de inadimplemento de suas obrigações, garantido o "
    "contraditório e ampla defesa, ficará sujeita às seguintes sanções previstas no RLC e na "
    "Lei nº 13.303, de 2016: advertência; multa moratória; multa compensatória; multa "
    "rescisória, para os casos de rescisão unilateral por descumprimento contratual; e "
    "suspensão do direito de participar de licitação e impedimento de contratar com a Conab, "
    "por até 02 (dois) anos.")
doc.add_paragraph(
    "a) advertência, aplicável nas infrações de menor gravidade, sem prejuízo à execução do "
    "objeto;")
add_paragraph_with_bookmarks([
    "b) multa por não regularização da documentação de habilitação: ",
    ("S15_MultaNaoRegularizacaoHabilitacao", "____"), "% sobre o valor estimado para a licitação.",
])
add_paragraph_with_bookmarks([
    "c) multa pela prática das condutas elencadas nos arts. 576 a 580 do RLC: ",
    ("S15_MultaCondutasArt576a580", "____"), "% sobre o valor estimado para a licitação.",
])
add_paragraph_with_bookmarks([
    "d) multa pela recusa em assinar o Contrato dentro do prazo estabelecido: ",
    ("S15_MultaRecusaAssinarContrato", "____"), "% sobre o valor homologado para a licitação.",
])
doc.add_paragraph("e) multa moratória por atraso injustificado na entrega da garantia contratual, conforme item 7.3.")
add_paragraph_with_bookmarks([
    "f) multa moratória de 0,2% (dois décimos por cento) sobre o ",
    ("S15_BaseCalculoMulta_E", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso na execução dos serviços, até o limite de 15 (quinze) dias.",
])
add_paragraph_with_bookmarks([
    "g) multa moratória de 0,3% (três décimos por cento) sobre o ",
    ("S15_BaseCalculoMulta_F", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso na execução dos serviços, por período superior ao previsto na alínea "
    "anterior, até o limite de 15 (quinze) dias. Esgotado o prazo, poderá ocorrer a não "
    "aceitação do objeto, sem prejuízo da rescisão unilateral da avença.",
])
add_paragraph_with_bookmarks([
    "h) multa compensatória no percentual de 10% (dez por cento) sobre o ",
    ("S15_BaseCalculoMulta_G", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", no caso de inexecução parcial do Contrato.",
])
add_paragraph_with_bookmarks([
    "i) multa compensatória de ", ("S15_MultaCompensatoriaTotal", "____"),
    "% sobre o valor total do Contrato, no caso de inexecução total do Contrato.",
])
add_paragraph_with_bookmarks([
    "j) multa rescisória de ", ("S15_MultaRescisoria", "____"),
    "% sobre o valor total do Contrato, no caso de rescisão contratual unilateral.",
])
add_paragraph_with_bookmarks([
    "k) multa de 0,2% a 3,2% por dia sobre o ",
    ("S15_BaseCalculoMulta_TABELA2", "(valor do Contrato, valor anual ou valor mensal)"),
    ", conforme grau da infração cometida (Tabelas 1 e 2 abaixo).",
])

tabela1 = doc.add_table(rows=1, cols=2)
tabela1.style = "Table Grid"
tabela1.rows[0].cells[0].text = "TABELA 1 — DESCRIÇÃO DA INFRAÇÃO"
tabela1.rows[0].cells[1].text = "GRAU"
infracoes_fixas = [
    ("Permitir situação que crie a possibilidade de causar dano físico, lesão corporal ou "
     "consequências letais, por ocorrência", "5"),
    ("Suspender ou interromper, salvo motivo de força maior ou caso fortuito, os serviços "
     "contratuais, por dia e por unidade de atendimento", "4"),
    ("Servir-se de funcionário sem qualificação para executar os serviços contratados, por "
     "empregado e por dia", "3"),
    ("Recusar-se a executar serviço determinado pela fiscalização, por serviço e por dia", "2"),
    ("Deixar de cumprir determinação formal ou instrução complementar da fiscalização, por "
     "ocorrência", "2"),
    ("Deixar de substituir empregado alocado que não atenda às necessidades do serviço, por "
     "funcionário e por dia", "1"),
    ("Deixar de cumprir quaisquer itens do Edital e seus Anexos não previstos nesta tabela, "
     "após reincidência formalmente notificada pela fiscalização, por item e por ocorrência", "3"),
    ("Deixar de indicar e manter durante a execução do Contrato os prepostos previstos no "
     "Edital/Contrato, por dia", "1"),
    ("Deixar de providenciar treinamento para seus funcionários conforme previsto na relação "
     "de obrigações da Contratada, por funcionário", "1"),
]
for descricao, grau in infracoes_fixas:
    linha = tabela1.add_row()
    linha.cells[0].text = descricao
    linha.cells[1].text = grau
add_linha_modelo_tabela(tabela1, "S15_InfracoesExtras_LINHA_MODELO", ["((DESCRICAO))", "((GRAU))"])

tabela2 = doc.add_table(rows=1, cols=2)
tabela2.style = "Table Grid"
tabela2.rows[0].cells[0].text = "GRAU"
tabela2.rows[0].cells[1].text = "CORRESPONDÊNCIA (percentual por dia)"
for grau, percentual in [("1", "0,2%"), ("2", "0,4%"), ("3", "0,8%"), ("4", "1,6%"), ("5", "3,2%")]:
    linha = tabela2.add_row()
    linha.cells[0].text = grau
    linha.cells[1].text = percentual
doc.add_paragraph(
    "16.2. As multas moratória, compensatória e rescisória possuem fatos geradores "
    "distintos; se aplicadas duas multas sobre o mesmo fato gerador, configura bis in idem.")
doc.add_paragraph(
    "16.3. Da sanção de suspensão: cabe em razão de ação ou omissão capaz de causar, ou que "
    "tenha causado, prejuízo à Conab, suas instalações, pessoas, imagem, meio ambiente, ou em "
    "decorrência de determinação legal, aplicada por até 02 (dois) anos, nos termos dos arts. "
    "579 a 580 do RLC.")

# 17. DA RESCISÃO CONTRATUAL (texto fixo)
add_heading("17. DA RESCISÃO CONTRATUAL")
doc.add_paragraph(
    "17.1. A inexecução total do Contrato ensejará a sua rescisão, enquanto a inexecução "
    "parcial poderá ensejar a sua rescisão, com as consequências cabíveis, conforme disposto "
    "nos arts. 568 a 572 do RLC.")
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
    "prejuízo das sanções previstas no título anterior: a) assunção imediata do objeto "
    "contratado, pela Conab, no estado e local em que se encontrar; b) execução da garantia "
    "contratual, quando houver, para ressarcimento pelos eventuais prejuízos sofridos pela "
    "Conab; e c) na hipótese de insuficiência da garantia contratual, quando houver, a "
    "retenção dos créditos decorrentes do Contrato até o limite dos prejuízos causados à "
    "Conab.")
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
    "18.1. Conforme disposto no art. 12, §1º, inciso II, do RLC e no art. 7º do Decreto nº "
    "7.203, de 2010, fica vedada a contratação de empregado ou dirigente da Conab como pessoa "
    "física, ou de pessoa/empresa com relação de parentesco até o terceiro grau civil com "
    "autoridade do Ministério do Desenvolvimento Agrário e Agricultura Familiar (MDA), dirigente ou empregado da "
    "Conab com atribuições na área de licitação e contratação.")

# 19. DA SUSTENTABILIDADE AMBIENTAL
add_heading("19. DA SUSTENTABILIDADE AMBIENTAL")
doc.add_paragraph(
    "19.1. Compete à Contratada, no que couber, atender os critérios de sustentabilidade "
    "ambiental previstos no art. 10 do RLC.")
doc.add_paragraph(
    "19.2. A Contratada se responsabiliza administrativamente, civilmente e penalmente por "
    "qualquer dano causado pelo seu serviço ao meio ambiente.")
add_paragraph_with_bookmarks(["19.3. ", ("S18_CriteriosEspecificos", "")])

# 20. DA MATRIZ DE RISCOS E RESPONSABILIDADES (texto fixo)
add_heading("20. DA MATRIZ DE RISCOS E RESPONSABILIDADES")
doc.add_paragraph(
    "20.1. A Matriz de Riscos é a cláusula contratual definidora dos riscos e das "
    "responsabilidades entre o Contratante e a Contratada, conforme hipóteses elencadas na "
    "Matriz de Riscos - Anexo I do Termo de Referência.")

# 21. DAS DISPOSIÇÕES FINAIS
add_heading("21. DAS DISPOSIÇÕES FINAIS")
doc.add_paragraph(
    "21.1. Integram este Termo de Referência os seguintes Anexos: a) Anexo I - Matriz de "
    "Riscos; b) Anexo II - Modelo de Apresentação de Proposta; c) Anexo III - Declaração de "
    "Vistoria Técnica; d) Anexo IV - Modelo de Declaração de Contratos Firmados com a "
    "Iniciativa Privada e a Administração Pública; e e) Anexo V - Modelo de Instrumento de "
    "Medição de Resultados (IMR).")
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

# ANEXO II - MODELO DE APRESENTAÇÃO DE PROPOSTA
doc.add_page_break()
add_heading("ANEXO II DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE APRESENTAÇÃO DE PROPOSTA")
doc.add_paragraph(
    "A proposta deverá conter: descrição detalhada do objeto e valor total, em algarismos e "
    "por extenso; prazo de execução e condições de execução dos serviços; prazo de vigência "
    "do Contrato; prazo de garantia oferecido, quando aplicável, não inferior a 60 (sessenta) "
    "dias corridos; prazo de validade da proposta, não inferior a 60 (sessenta) dias corridos; "
    "informação de que os preços cotados incluem todos os custos necessários (impostos, "
    "tributos, encargos trabalhistas, previdenciários, comerciais e fiscais); e os dados da "
    "proponente (razão social, endereço, telefone, CNPJ, dados bancários).")
tabela_proposta = doc.add_table(rows=1, cols=3)
tabela_proposta.style = "Table Grid"
for celula, texto in zip(tabela_proposta.rows[0].cells, ["Item", "Especificações", "Valor Total"]):
    celula.text = texto
for _ in range(5):
    tabela_proposta.add_row()
doc.add_paragraph("VALOR TOTAL DOS ITENS: R$")

# ANEXO III - MODELO DE DECLARAÇÃO DE VISTORIA
doc.add_page_break()
add_heading("ANEXO III DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE DECLARAÇÃO DE VISTORIA")
doc.add_paragraph("À Companhia Nacional de Abastecimento – CONAB")
doc.add_paragraph("Ref.: Pregão Eletrônico – nº ___/20XX")
doc.add_paragraph(
    "Declaramos que, em atendimento ao previsto no Pregão Eletrônico nº ___/_____, "
    "eu, ______________________________________, portador(a) da CI/RG nº______________ e do "
    "CPF nº _____________________, responsável da empresa "
    "___________________________________________________, como seu representante legal para "
    "os fins da presente declaração, compareci perante a CONAB/Sureg/XX, situada à "
    "_____________________________________________________________________________________, "
    "e vistoriei os locais onde serão executados os serviços objeto da licitação em apreço, "
    "tomando plena ciência das condições e graus de dificuldade existentes.")
doc.add_paragraph("Cidade-UF, ___/___/20XX.")
doc.add_paragraph("Assinatura e carimbo (Responsável Técnico da Empresa)")
doc.add_paragraph("Visto: Responsável pela CONAB")

# ANEXO IV - MODELO DE DECLARAÇÃO DE CONTRATOS FIRMADOS (Anexo IV do RLC,
# transcrito integralmente a partir de
# anexos RLC/anexo_iv_minuta_de_declaracao_de_contratos_firmados.odt)
doc.add_page_break()
add_heading("ANEXO IV DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE DECLARAÇÃO DE CONTRATOS FIRMADOS COM A INICIATIVA PRIVADA E A ADMINISTRAÇÃO PÚBLICA")
doc.add_paragraph(
    "Declaro que a empresa ________________________, inscrita no CNPJ (MF) nº ________, "
    "inscrição estadual nº _____________, estabelecida em _____________, possui os seguintes "
    "contratos firmados com a iniciativa privada e a Administração Pública:")
tabela_contratos = add_tabela(
    ["Nome do Órgão/Empresa", "Vigência do Contrato", "Valor Total do Contrato*"],
    linhas_vazias=10,
)
linha_total = tabela_contratos.add_row()
linha_total.cells[0].text = "Valor Total dos Contratos"
linha_total.cells[2].text = "R$"
add_tabela(["Assinatura e Carimbo do Emissor"])
doc.add_paragraph("Observação:")
doc.add_paragraph(
    "Nota 1: Além dos nomes dos órgãos/empresas, o licitante deverá informar também o "
    "endereço completo dos órgãos/empresas, com os quais têm contratos vigentes.")
doc.add_paragraph(
    "Nota 2: *Considera-se o valor remanescente do contrato, excluindo o já executado.")
doc.add_paragraph(
    "Fórmula exemplificativa, para fins de atendimento ao disposto no artigo 131, § 6º, "
    "inciso IV, alíneas a e b, do Regulamento de Licitações e Contratos da Conab.")
doc.add_paragraph(
    "1. A Declaração de Compromissos Assumidos deve informar que 1/12 (um doze avos) dos "
    "contratos firmados pela licitante não é superior ao Patrimônio Líquido da licitante.")
doc.add_paragraph("FÓRMULA DE CÁLCULO:")
add_tabela(["Valor do Patrimônio Líquido x 12", "Valor total dos contratos *", "> 1"])
doc.add_paragraph("Observação:")
doc.add_paragraph("Nota 1: Esse resultado deverá ser superior a 1 (um).")
doc.add_paragraph(
    "Nota 2: Considera-se o valor remanescente do contrato, excluindo o já executado*.")
doc.add_paragraph(
    "2. Caso a diferença entre a receita bruta discriminada na Demonstração do Resultado "
    "do Exercício (DRE) e a declaração apresentada seja maior que 10% (dez por cento) "
    "positivo ou negativo em relação à receita bruta, o licitante deverá apresentar "
    "justificativas.")
doc.add_paragraph("FÓRMULA DE CÁLCULO:")
add_tabela(["(Valor da Receita Bruta - Valor total dos Contratos) x 100", "Valor da Receita Bruta"])

# ANEXO V - MODELO DE INSTRUMENTO DE MEDIÇÃO DE RESULTADO (IMR)
doc.add_page_break()
add_heading("ANEXO V DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE INSTRUMENTO DE MEDIÇÃO DE RESULTADO (IMR)")

doc.add_paragraph("Indicador Nº ____ – ____ (título do indicador)")
tabela_imr = doc.add_table(rows=1, cols=2)
tabela_imr.style = "Table Grid"
tabela_imr.rows[0].cells[0].text = "Item"
tabela_imr.rows[0].cells[1].text = "Descrição"
for item in ["Finalidade", "Meta a cumprir", "Instrumento de medição", "Forma de acompanhamento",
             "Periodicidade", "Mecanismo de cálculo", "Início de vigência",
             "Faixas de ajuste no pagamento", "Sanções", "Observações"]:
    linha = tabela_imr.add_row()
    linha.cells[0].text = item

doc.add_paragraph("Exemplo de indicador")
doc.add_paragraph("Nº 01 – Prazo de atendimento de demandas (OS)")
tabela_imr_exemplo = doc.add_table(rows=1, cols=2)
tabela_imr_exemplo.style = "Table Grid"
tabela_imr_exemplo.rows[0].cells[0].text = "Item"
tabela_imr_exemplo.rows[0].cells[1].text = "Descrição"
exemplo_imr = [
    ("Finalidade", "Garantir um atendimento célere as demandas da Conab."),
    ("Meta a cumprir", "24 h."),
    ("Instrumento de medição", "Sistema informatizado de solicitação de serviços – Ordem de Serviços (OS) eletrônica."),
    ("Forma de acompanhamento", "Pelo sistema."),
    ("Periodicidade", "Mensal."),
    ("Mecanismo de cálculo", "Cada OS será verificada e valorada individualmente. N.º de horas no atendimento/24 = X"),
    ("Início de vigência", "Data da assinatura do contrato."),
    ("Faixas de ajuste no pagamento",
     "X até 1 – 100% do valor da OS. De 1 a 1,5 – 90% do valor da OS. De 1,5 a 2 – 80% do valor da OS."),
    ("Sanções",
     "20% das OS acima de 2 – multa de XX. 30% das OS acima de 2 – multa de XX + rescisão contratual."),
    ("Observações", "RLC"),
]
for item, descricao in exemplo_imr:
    linha = tabela_imr_exemplo.add_row()
    linha.cells[0].text = item
    linha.cells[1].text = descricao

doc.save("Template_TR_Servicos.docx")
print("Template_TR_Servicos.docx gerado com sucesso.")
