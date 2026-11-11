"""Gera Template_TR_Dispensa.docx a partir do texto oficial da minuta
SEI 55379661, com bookmarks nos pontos de preenchimento identificados em
js/dados-tr-dispensa.js. Rodar novamente sempre que o texto oficial mudar."""
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


def add_bloco_variante_sentinela_inicio(nome_bookmark):
    """Para blocos que misturam parágrafos e tabelas (Seções 21-25
    combinadas, e cada um dos dois anexos de contrato): marca só o INÍCIO
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
    "1.1. O presente Termo de Referência tem por objeto a contratação de ",
    ("S01_DescricaoObjeto", "(descrição do serviço a ser contratado)"),
    ", conforme especificações, condições, quantidades e exigências estabelecidas neste instrumento, por meio de contratação direta por dispensa de licitação, com fundamento legal no art. 416, inciso II, do RLC.",
])
add_paragraph_with_bookmarks([
    "1.2. O custo total estimado para a prestação dos serviços deste Termo de Referência é de R$ ",
    ("S01_ValorEstimado", "________"),
    ". O critério utilizado para a formação da estimativa foi de ",
    ("S01_MetodologiaEstimativa", "(MÉDIA, MEDIANA OU MENOR PREÇO)"),
    " dos preços praticados no mercado.",
])
add_paragraph_with_bookmarks([
    "1.3. O critério de julgamento das propostas será o de ",
    ("S01_CriterioJulgamento", "[menor preço ou maior desconto]"), " ofertado.",
])
add_paragraph_with_bookmarks([
    "1.4. Os recursos orçamentários decorrentes da contratação do objeto deste Termo de Referência "
    "estão consignados no Orçamento da Conab para o ano 20XX, conforme segue: ND: (",
    ("S01_ND", "XXXX"), "), PTRES: (", ("S01_PTRES", "XXXX"), "), FONTE: (", ("S01_Fonte", "XXXX"), ").",
])

# 2. DA JUSTIFICATIVA
add_heading("2. DA JUSTIFICATIVA")
add_paragraph_with_bookmarks(["2.1. ", ("S02_Motivacao", "(motivação da contratação)")])
doc.add_paragraph(
    "2.2. O serviço em apreço será contratado por meio de dispensa de licitação, com fundamento "
    "legal no art. 416, inciso II, do Regulamento de Licitações e Contratos da Conab.")
add_bloco_variante("S02_ExcecaoExclusividadeMEEPP", {
    "Sim": [
        "2.3. Em atendimento à Lei Complementar 123/2016 e ao Decreto 8.538/2015, que regulamentam "
        "o tratamento favorecido, diferenciado e simplificado para as microempresas e empresas de "
        "pequeno porte, este processo será destinado exclusivamente à participação de "
        "microempresas e empresas de pequeno porte, tendo em vista o valor estimado da contratação "
        "ser de até R$ 80.000,00 (oitenta mil reais), estabelecido em lei.",
    ],
    "Não": [
        "2.3. Em que pese o valor estimado da contratação seja até R$ 80.000,00 (oitenta mil "
        "reais), a dispensa de licitação não será destinada exclusivamente à participação de "
        "microempresas e empresas de pequeno porte, na forma da Lei Complementar 123/2016 e do "
        "inciso III do art. 10 do Decreto 8.538/2015, em razão da contratação em apreço se "
        "enquadrar na exceção prevista no art. 10 do Decreto 8.538/2015, isto é: ",
    ],
})
add_paragraph_with_bookmarks(["", ("S02_MotivoExcecaoExclusividade", "")])
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
add_paragraph_with_bookmarks([
    "2.5. Nos termos do art. 416, incisos I e II, do RLC, foi observado o limite de valor "
    "considerando o somatório do valor desta contratação com o valor de outros objetos de "
    "mesma natureza contratados pela mesma unidade gestora no mesmo exercício financeiro, "
    "não havendo fracionamento de despesa: ",
    ("S02_SomatorioContratacoesMesmaNatureza", "(descrever a apuração do somatório e a conclusão quanto ao limite)"),
])
doc.add_paragraph(
    "2.6. A presente contratação será precedida de divulgação de aviso em sítio eletrônico "
    "oficial, pelo prazo mínimo de 3 (três) dias úteis, com atendimento à Instrução "
    "Normativa SEGES/ME nº 67, de 2021, para busca da proposta mais vantajosa.")
add_bloco_variante("S02_PagamentoCartao", {
    "Sim": [
        "2.7. A contratação será paga por meio de cartão de pagamento, com divulgação do "
        "respectivo extrato no Portal Nacional de Contratações Públicas (PNCP)."
    ],
    "Não": [[
        "2.7. A presente contratação não será paga por meio de cartão de pagamento, pelas "
        "seguintes razões: ", ("S02_JustificativaNaoCartao", ""),
    ]],
})

# 3. DO OBJETIVO
add_heading("3. DO OBJETIVO")
add_paragraph_with_bookmarks(["3.1. ", ("S03_Objetivo", "(objetivo da contratação)")])

# 4. DA VIGÊNCIA DO CONTRATO
add_heading("4. DA VIGÊNCIA DO CONTRATO")
add_bloco_variante("S04_TipoVigencia", {
    "Até 1 ano, prorrogável": [[
        "4.1. O prazo de vigência do Contrato será de ", ("S04_PrazoVigenciaMeses_Inline", "____"),
        " meses (duração do Contrato, conforme "
        "arts. 461 a 463 do RLC), contados da data de sua assinatura, prorrogável desde que "
        "observados os requisitos do art. 488 do RLC, entre eles a demonstração da "
        "vantajosidade na manutenção do ajuste (em regra, por pesquisa de preços), sob pena de "
        "a Conab não poder prorrogar o Contrato.",
    ]],
    "Plurianual (até 5 anos)": [[
        "4.1. O prazo de vigência do Contrato será de ", ("S04_PrazoVigenciaAnos_Inline", "____"),
        " anos (duração do Contrato, conforme "
        "arts. 461 a 462 do RLC), contados da data de sua assinatura, podendo ser prorrogado até "
        "o limite de 05 (cinco) anos, desde que observados os requisitos previstos no art. 488 "
        "do RLC, entre eles a demonstração da vantajosidade na manutenção do ajuste (em regra, "
        "por pesquisa de preços), sob pena de a Conab não poder prorrogar o Contrato.",
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
doc.add_paragraph("5.1. DA APRESENTAÇÃO DA PROPOSTA")
add_paragraph_with_bookmarks([
    "5.1.1. A proponente deverá enviar a sua proposta à Conab até ",
    ("S05_PrazoLimiteEnvioProposta", "(data e horário)"),
    " contendo: descrição detalhada do objeto; preço unitário e global, em algarismos e por "
    "extenso; prazo de execução e condições de execução dos serviços; prazo de vigência do "
    "Contrato; prazo de garantia oferecido, quando aplicável, não inferior a 60 (sessenta) "
    "dias corridos; prazo de validade da proposta, não inferior a 60 (sessenta) dias corridos; "
    "informação de que os preços cotados incluem todos os custos necessários (impostos, tributos, "
    "encargos trabalhistas, previdenciários, comerciais e fiscais); e os dados da proponente "
    "(razão social, endereço, telefone, CNPJ, dados bancários).",
])
doc.add_paragraph(
    "5.1.2. A proposta de preços deverá ser redigida em língua portuguesa, sem emendas, rasuras "
    "ou entrelinhas, devidamente datada e assinada pelo representante legal da proponente. Todas "
    "as especificações do objeto contidas na proposta vinculam a Contratada.")
doc.add_paragraph("5.2. DA HABILITAÇÃO E QUALIFICAÇÃO TÉCNICA")
doc.add_paragraph(
    "5.2.1. Para a habilitação, a proponente detentora da melhor oferta deverá apresentar toda a "
    "documentação relacionada a seguir. A proponente já regularmente cadastrada e habilitada "
    "parcialmente no Sistema de Cadastramento Unificado de Fornecedores (SICAF) ficará dispensada "
    "de apresentar os documentos correlatos, quando os registros no SICAF estiverem válidos, "
    "completos e regulares.")
doc.add_paragraph(
    "5.2.2. Relativos à Habilitação Jurídica: os documentos constitutivos compatíveis com a "
    "natureza jurídica da proponente (inscrição no Registro Público de Empresas Mercantis para o "
    "empresário individual; Certificado da Condição de Microempreendedor Individual para o MEI; "
    "ato constitutivo, estatuto ou contrato social em vigor, devidamente registrado, para "
    "sociedade empresária, EIRELI ou sociedade por ações; inscrição do ato constitutivo para "
    "sociedade simples; ata de fundação e estatuto social para cooperativa; decreto de "
    "autorização para empresa estrangeira em funcionamento no País), acompanhados de todas as "
    "alterações ou da consolidação respectiva.")
add_paragraph_with_bookmarks([
    "5.2.2.1. Quando a atividade relativa ao objeto exigir registro ou autorização de "
    "funcionamento de órgão regulador (ANVISA, ANP, Ministério da Agricultura, Polícia Federal, "
    "IBAMA etc.): ", ("S05_RegistroEspecialAtividade", ""),
])
doc.add_paragraph(
    "5.2.3. Relativos à Regularidade Fiscal e Trabalhista: prova de inscrição no CNPJ; prova de "
    "regularidade fiscal perante a Fazenda Nacional (RFB e PGFN); prova de inscrição no cadastro "
    "de contribuintes estadual ou municipal, quando houver; prova de regularidade perante a "
    "Fazenda Estadual e Municipal; prova de regularidade perante o INSS e o FGTS; e prova de "
    "inexistência de débitos inadimplidos perante a Justiça do Trabalho (CNDT). Havendo alguma "
    "restrição na comprovação da regularidade fiscal de microempresa, empresa de pequeno porte ou "
    "cooperativa, será assegurado prazo de 5 (cinco) dias úteis, prorrogável por igual período, "
    "para regularização, sob pena de decadência do direito à contratação.")
doc.add_paragraph(
    "5.2.4. Relativos à Qualificação Econômico-Financeira: certidão negativa de feitos sobre "
    "falência, expedida pelo distribuidor da sede da proponente, com prazo máximo de 60 "
    "(sessenta) dias quando não mencionar validade explícita; e balanço patrimonial e "
    "demonstrações contábeis do último exercício social, vedada a substituição por balancetes ou "
    "balanços provisórios. Será considerada boa a situação financeira da proponente quando os "
    "Índices de Liquidez Geral (LG), Solvência Geral (SG) e Liquidez Corrente (LC) forem maiores "
    "que 1 (um), apurados pelas fórmulas: LG = (Ativo Circulante + Realizável a Longo Prazo) / "
    "(Passivo Circulante + Exigível a Longo Prazo); SG = Ativo Total / (Passivo Circulante + "
    "Exigível a Longo Prazo); LC = Ativo Circulante / Passivo Circulante. A proponente com "
    "qualquer desses índices igual ou inferior a 1 (um) deverá comprovar patrimônio líquido "
    "mínimo de 10% (dez por cento) do valor total estimado da contratação.")
add_paragraph_with_bookmarks([
    "5.2.5. Relativos à Qualificação Técnica: atestado de capacidade técnica que comprove a "
    "execução de serviços com características, quantidades e prazos compatíveis com o objeto, "
    "por período não inferior a ", ("S05_PrazoExperienciaMinimaAnos", "____"),
    " anos, fornecido por pessoas jurídicas de direito público ou privado. Para fins de "
    "comprovação de compatibilidade: ",
    ("S05_QuantitativoServicoPrincipal", "(quantitativo mínimo e serviço(s) principal(is))"),
])
doc.add_paragraph(
    "5.2.6. Será aceito o somatório de atestados para fins de comprovação de capacidade técnica. "
    "A proponente vencedora, apenas por ocasião da assinatura do Contrato, deverá comprovar "
    "inscrição na entidade profissional competente, quando for o caso, e apresentar declaração de "
    "que é Microempresa, Empresa de Pequeno Porte ou Cooperativa, se for o caso.")
doc.add_paragraph(
    "5.2.7. Como condição prévia ao exame da documentação de habilitação, será verificado o "
    "eventual descumprimento das condições de participação mediante consulta ao SICAF, ao "
    "Cadastro Nacional de Empresas Inidôneas e Suspensas (CEIS), ao Cadastro Nacional de "
    "Condenações Cíveis por Atos de Improbidade Administrativa e à Lista de Inidôneos do TCU, em "
    "nome da proponente e de seu sócio majoritário.")
doc.add_paragraph(
    "5.2.8. Será inabilitada a proponente que não comprovar sua habilitação, ressalvado o "
    "saneamento de erros ou falhas que não alterem a substância das propostas e dos documentos, "
    "mediante despacho fundamentado. Será considerada vencedora a proponente que oferecer o menor "
    "preço ou maior desconto e atender às exigências de habilitação deste Termo de Referência.")
doc.add_paragraph("5.3. DA VISTORIA")
add_bloco_variante("S05_TipoVistoria", {
    "Obrigatória": [
        "5.3.1. Para o correto dimensionamento e elaboração de sua proposta, o proponente deverá "
        "realizar vistoria nas instalações do local de execução dos serviços, acompanhado por "
        "empregado designado para esse fim, devendo o agendamento ser efetuado previamente, "
        "podendo sua realização ser comprovada por declaração assinada pelo proponente de que "
        "realizou a vistoria e tem pleno conhecimento das condições e peculiaridades do local, na "
        "forma do Anexo III deste TR.",
    ],
    "Facultativa": [
        "5.3.1. Para o correto dimensionamento e elaboração de sua proposta, recomenda-se que "
        "o proponente realize vistoria nas instalações do local de execução dos serviços, "
        "sendo a realização facultativa. Optando por realizá-la, deverá estar acompanhado por "
        "empregado designado para esse fim, devendo o agendamento ser efetuado previamente, "
        "podendo sua realização ser comprovada por declaração assinada pelo proponente de que "
        "realizou a vistoria e tem pleno conhecimento das condições e peculiaridades do local, na "
        "forma do Anexo III deste TR.",
        "O proponente que optar por não realizar a vistoria deverá apresentar declaração de que "
        "tem pleno conhecimento das condições e peculiaridades do local, assumindo total "
        "responsabilidade por sua opção, na forma do Anexo III deste TR.",
    ],
})
add_paragraph_with_bookmarks([
    "5.3.1.1. O agendamento da vistoria deverá ser feito de segunda a sexta-feira, das ",
    ("S05_HorarioVistoriaInicio", "____"), " às ", ("S05_HorarioVistoriaFim", "____"),
    " horas, pelo telefone (", ("S05_TelefoneAgendamentoVistoria", "____"), ").",
])
doc.add_paragraph("5.3.2. O prazo para vistoria será igual ao prazo estabelecido para envio das propostas.")

# 6. DA DESCRIÇÃO DOS SERVIÇOS
add_heading("6. DA DESCRIÇÃO DOS SERVIÇOS")
doc.add_paragraph(
    "6.1. A prestação dos serviços não gera vínculo empregatício entre os empregados da "
    "Contratada e a Conab, vedando-se qualquer relação que caracterize pessoalidade e "
    "subordinação direta.")
add_paragraph_with_bookmarks(["6.2. ", ("S06_DetalhamentoServicos", "")])
add_paragraph_with_bookmarks([
    "6.3. Os serviços objeto deste Termo de Referência deverão ser realizados em ",
    ("S06_LocalExecucao", "(local)"), ", de segunda a sexta-feira, no horário de ",
    ("S06_HorarioExecucao", "(horário)"), ".",
])
_bid_enderecos = add_bloco_variante_sentinela_inicio("S06_EnderecosAdicionais_CONTAINER")
doc.add_paragraph(
    "6.3.1. Quando o objeto envolver a execução em mais de um endereço, os locais adicionais "
    "são os seguintes:")
tabela_enderecos = add_tabela(["Endereço", "Observação"])
add_linha_modelo_tabela(tabela_enderecos, "S06_EnderecosAdicionais_LINHA_MODELO", ["((ENDERECO))", "((OBSERVACAO))"])
add_bloco_variante_sentinela_fim(_bid_enderecos)
add_paragraph_with_bookmarks([
    "6.4. O prazo de execução dos serviços será de ", ("S06_PrazoExecucaoDias", "____"),
    " dias, contados a partir da assinatura do Contrato.",
])
add_paragraph_with_bookmarks(["6.5. Rotina de execução: ", ("S06_RotinaExecucao", "")])
add_bloco_variante("S06_ExigeGarantiaServicos", {
    "Sim": ["6.6. Os serviços deverão ter um período de garantia de, no mínimo, ____ meses, "
            "contados do recebimento definitivo pela Contratante."],
    "Não": ["6.6. Não será exigida garantia dos serviços na contratação em apreço."],
})
add_paragraph_with_bookmarks(["Período de garantia (meses): ", ("S06_GarantiaServicosMeses", "____")])
add_bloco_variante("S06_ExigeMateriaisDisponibilizados", {
    "Sim": ["6.7. Para a perfeita execução dos serviços, a Contratada deverá disponibilizar os "
            "materiais, equipamentos, ferramentas e utensílios necessários, promovendo sua "
            "substituição quando necessário."],
    "Não": [""],
})
add_paragraph_with_bookmarks(["", ("S06_MateriaisDisponibilizados", "")])
add_paragraph_with_bookmarks([
    "6.8. Prazo para recebimento provisório: ", ("S06_PrazoRecebimentoProvisorioDias", "____"),
    " dias. Prazo para recebimento definitivo, contado do provisório: ",
    ("S06_PrazoRecebimentoDefinitivoDias", "____"), " dias.",
])
doc.add_paragraph(
    "6.9. Durante a vigência do Contrato, a Conab, para avaliar a qualidade na execução dos "
    "serviços, adotará o Instrumento de Medição de Resultados (IMR), constante no Anexo IV deste "
    "Termo de Referência.")
doc.add_paragraph(
    "6.10. O recebimento provisório será realizado pela fiscalização; o recebimento definitivo, "
    "pelo empregado ou Comissão designada, nos termos dos arts. 526 e 527 do RLC.")

# 7. DO INSTRUMENTO CONTRATUAL
add_heading("7. DO INSTRUMENTO CONTRATUAL")
doc.add_paragraph(
    "7.1. As obrigações decorrentes da presente dispensa de licitação serão formalizadas por meio "
    "de contrato simplificado ou contrato, conforme escolha consignada neste Termo de Referência, "
    "celebrado entre a Conab e a proponente vencedora, observando os termos do RLC.")
doc.add_paragraph(
    "7.2. A proponente vencedora será convocada para, no prazo de até 05 (cinco) dias úteis, "
    "celebrar o instrumento contratual, mediante prévia comprovação das condições de habilitação "
    "consignadas neste TR, mantidas durante a vigência contratual. Antes da contratação será "
    "feita consulta ao CADIN, ao SICAF, ao CEIS, à CNDT, ao Cadastro Nacional de Condenações "
    "Cíveis por Atos de Improbidade Administrativa e à Lista de Inidôneos do TCU.")

# 8. DA GARANTIA CONTRATUAL
add_heading("8. DA GARANTIA CONTRATUAL")
add_bloco_variante("S08_ExigeGarantia", {
    "Sim": [[
        "8.1. A Contratada deverá apresentar, no prazo de 10 (dez) dias corridos após a "
        "assinatura do Contrato, prorrogáveis por igual período, comprovante de prestação de "
        "garantia correspondente a ", ("S08_PercentualGarantia_Inline", "____"),
        "%, conforme parâmetros do art. 439 do RLC, do ",
        ("S08_BaseCalculoGarantia", "(valor do Contrato ou valor anual do Contrato)"),
        ", mediante caução em dinheiro, seguro-garantia ou fiança bancária.",
    ]],
    "Não": ["8.1. Não será exigida garantia contratual, para a prestação dos serviços em apreço."],
})
add_paragraph_with_bookmarks(["Percentual da garantia: ", ("S08_PercentualGarantia", "____"), "%"])
doc.add_paragraph(
    "8.1.1. O percentual da garantia é, em regra, de até 5% do valor do Contrato (art. 439, "
    "§2º, do RLC), podendo ser elevado para até 10% em contratações de grande vulto envolvendo "
    "complexidade técnica e riscos financeiros elevados (art. 439, §3º, do RLC). Nos casos de "
    "contratação de serviços com dedicação exclusiva de mão de obra, o valor da garantia "
    "deverá corresponder a 5% do valor do contrato, limitado ao equivalente a 2 (dois) meses "
    "do custo da folha de pagamento dos empregados da contratada que venham a participar da "
    "execução dos serviços contratados (art. 439, §4º, do RLC).")
doc.add_paragraph(
    "8.2. A inobservância do prazo fixado para apresentação da garantia acarretará multa de "
    "0,08% (oito centésimos por cento) do valor do Contrato por dia de atraso, até o máximo de 2% "
    "(dois por cento). O atraso superior a 25 (vinte e cinco) dias autoriza a rescisão do "
    "Contrato, conforme o inciso I do art. 569 do RLC. A garantia deverá ter validade de 90 "
    "(noventa) dias após o término da vigência do Contrato.")

# 9. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO
add_heading("9. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO")
doc.add_paragraph(
    "9.1. Para a execução do ajuste, será adotado o método de trabalho baseado no conceito de "
    "delegação de responsabilidade. Nos termos dos arts. 543 e 544 do RLC, será designado Fiscal, "
    "seu substituto, ou Comissão de Fiscalização para acompanhar e fiscalizar a execução dos "
    "serviços, avaliando-a constantemente por meio do Instrumento de Medição de Resultado (IMR), "
    "nos termos dos arts. 545 a 548 do RLC.")
add_paragraph_with_bookmarks(["9.2. Rotinas de fiscalização específicas: ", ("S09_RotinasFiscalizacaoEspecificas", "")])

# 10. DAS OBRIGAÇÕES DA CONTRATANTE (texto fixo)
add_heading("10. DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph("10.1. São obrigações da Contratante:")
doc.add_paragraph(
    "a) exigir o cumprimento de todas as obrigações assumidas pela Contratada, de acordo com as "
    "cláusulas contratuais e os termos de sua proposta;")
doc.add_paragraph(
    "b) receber o objeto no prazo e condições estabelecidas neste Termo de Referência e seus "
    "anexos;")
doc.add_paragraph(
    "c) verificar minuciosamente, no prazo fixado, a conformidade dos serviços recebidos "
    "provisoriamente com as especificações constantes neste Termo de Referência e da proposta, "
    "para fins de aceitação e recebimento definitivo;")
doc.add_paragraph(
    "d) comunicar à Contratada, por escrito, sobre imperfeições, falhas ou irregularidades "
    "verificadas no curso da execução dos serviços, para que seja reparado ou corrigido;")
doc.add_paragraph(
    "e) rejeitar, no todo ou em parte, serviço executado em desacordo com o previsto neste Termo "
    "de Referência;")
doc.add_paragraph(
    "f) acompanhar e fiscalizar o cumprimento das obrigações da Contratada, através de empregado "
    "ou comissão especialmente designados;")
doc.add_paragraph(
    "g) efetuar o pagamento à Contratada no valor correspondente à prestação de serviço, no "
    "prazo e forma estabelecidos neste Termo de Referência e seus anexos.")
doc.add_paragraph(
    "10.2. A Conab não responderá por quaisquer compromissos assumidos pela Contratada com "
    "terceiros, ainda que vinculados à execução do Contrato, bem como por qualquer dano causado "
    "a terceiros em decorrência de ato da Contratada, de seus empregados, prepostos ou "
    "subordinados.")

# 11. DAS OBRIGAÇÕES DA CONTRATADA
add_heading("11. DAS OBRIGAÇÕES DA CONTRATADA")
doc.add_paragraph(
    "11.1. A Contratada deve cumprir todas as obrigações constantes neste Termo, seus anexos e "
    "sua proposta, assumindo como exclusivamente seus os riscos e as despesas decorrentes da boa "
    "e perfeita execução do objeto e, ainda:")
doc.add_paragraph(
    "a) executar os serviços conforme especificações deste Termo de Referência e de sua "
    "proposta, com a alocação dos empregados necessários ao perfeito cumprimento das cláusulas "
    "contratuais, além de fornecer os materiais e equipamentos, ferramentas e utensílios "
    "necessários, na qualidade e quantidade especificadas neste Termo de Referência e em sua "
    "proposta;")
doc.add_paragraph(
    "b) responsabilizar-se pelos vícios e danos decorrentes do objeto, de acordo com os artigos "
    "12, 13 e 17 a 27, do Código de Defesa do Consumidor (Lei nº 8.078, de 1990);")
doc.add_paragraph(
    "c) substituir, reparar, remover ou corrigir, às suas expensas, no total ou em parte, no "
    "prazo estabelecido pela Conab, os serviços efetuados em que se verificarem vícios, defeitos "
    "ou incorreções resultantes da execução ou dos materiais empregados;")
doc.add_paragraph(
    "d) comunicar à Contratante, no prazo máximo de 24 (vinte e quatro) horas que antecede a "
    "data da finalização dos serviços, os motivos que impossibilitem o cumprimento do prazo "
    "previsto, com a devida comprovação;")
doc.add_paragraph(
    "e) utilizar empregados habilitados e com conhecimentos básicos dos serviços a serem "
    "executados, em conformidade com as normas e determinações em vigor;")
doc.add_paragraph(
    "f) apresentar os empregados devidamente uniformizados e identificados por meio de crachá, "
    "além de provê-los com os Equipamentos de Proteção Individual - EPI, quando for o caso;")
doc.add_paragraph(
    "g) apresentar à Contratante, quando for o caso, a relação nominal dos empregados que "
    "adentrarão a Conab para a execução do serviço;")
doc.add_paragraph(
    "h) responsabilizar-se por todas as obrigações trabalhistas, sociais, previdenciárias, "
    "tributárias e as demais previstas na legislação específica, cuja inadimplência não "
    "transfere responsabilidade à Contratante;")
doc.add_paragraph(
    "i) atender as solicitações da Contratante quanto à substituição dos empregados alocados, no "
    "prazo fixado pelo fiscal do Contrato, nos casos em que ficar constatado descumprimento das "
    "obrigações relativas à execução do serviço, conforme descrito neste Termo de Referência;")
doc.add_paragraph(
    "j) instruir seus empregados quanto à necessidade de acatar as normas internas da Conab;")
doc.add_paragraph(
    "k) instruir seus empregados a respeito das atividades a serem desempenhadas, alertando-os a "
    "não executar atividades não abrangidas pelo Contrato, devendo a Contratada relatar à "
    "Contratante toda e qualquer ocorrência neste sentido, a fim de evitar desvio de função;")
doc.add_paragraph(
    "l) relatar à Contratante toda e qualquer irregularidade verificada no decorrer da prestação "
    "dos serviços;")
doc.add_paragraph(
    "m) não permitir a utilização de qualquer trabalho do menor de dezesseis anos, exceto na "
    "condição de aprendiz para os maiores de quatorze anos; nem permitir a utilização do "
    "trabalho do menor de dezoito anos em trabalho noturno, perigoso ou insalubre;")
doc.add_paragraph(
    "n) guardar sigilo sobre todas as informações obtidas em decorrência do cumprimento do "
    "Contrato;")
doc.add_paragraph(
    "o) arcar com o ônus decorrente de eventual equívoco no dimensionamento dos quantitativos de "
    "sua proposta, devendo complementá-los, caso o previsto inicialmente em sua proposta não "
    "seja satisfatório para o atendimento ao objeto da contratação, exceto quando ocorrer algum "
    "dos eventos arrolados no Art. 497 do RLC da Conab;")
doc.add_paragraph(
    "p) manter, durante toda a execução do Contrato, em compatibilidade com as obrigações "
    "assumidas, todas as condições de habilitação e qualificação exigidas na contratação;")
doc.add_paragraph(
    "q) prestar todos os esclarecimentos que forem solicitados pela Conab, durante a realização "
    "do Contrato;")
doc.add_paragraph(
    "r) indicar preposto para representá-la durante a execução do Contrato;")
doc.add_paragraph(
    "s) deter instalações, aparelhamento e pessoal técnico adequados e disponíveis para a "
    "realização do objeto da contratação.")
doc.add_paragraph(
    "Garantir aos seus empregados, na execução do Contrato, política de respeito à "
    "diversidade, prevenção a práticas discriminatórias e alinhamento com os princípios "
    "institucionais de equidade de gênero e raça.")
add_bloco_variante("S11_NaturezaIntelectual", {
    "Sim": [
        "11.2. Ceder os direitos patrimoniais relativos ao projeto ou serviço técnico "
        "especializado, para que a Conab possa utilizá-lo, de acordo com o art. 525 do RLC. "
        "Quando o projeto referir-se a obra imaterial de caráter tecnológico, insuscetível de "
        "privilégio, a cessão dos direitos incluirá o fornecimento de todos os dados, documentos "
        "e elementos de informação pertinentes à tecnologia de concepção, desenvolvimento, "
        "fixação em suporte físico e aplicação da obra.",
        "11.3. Assegurar à Contratante o direito de propriedade intelectual dos produtos "
        "desenvolvidos, inclusive sobre adequações e atualizações futuras, de forma permanente, "
        "e os direitos autorais da solução, do projeto, de suas especificações técnicas e da "
        "documentação produzida, inclusive os produzidos por terceiros subcontratados, ficando "
        "proibida sua utilização sem autorização expressa da Contratante, sob pena de multa, sem "
        "prejuízo das sanções civis e penais cabíveis.",
        "11.4. Realizar a transição contratual com transferência de conhecimento, tecnologia e "
        "técnicas empregadas, sem perda de informações, podendo exigir, inclusive, a capacitação "
        "dos técnicos da Contratante ou da nova empresa que continuará a execução dos serviços, "
        "nos termos do art. 17 do RLC.",
    ],
    "Não": [""],
})

# 12. DO PAGAMENTO
add_heading("12. DO PAGAMENTO")
doc.add_paragraph(
    "12.1. O pagamento será efetuado pela Conab no prazo de até 10 (dez) dias corridos, contados "
    "do recebimento da Nota Fiscal/Fatura.")
doc.add_paragraph(
    "12.2. A emissão da Nota Fiscal/Fatura será precedida do recebimento provisório e definitivo "
    "do serviço: no prazo de até 05 (cinco) dias corridos do adimplemento da parcela, a "
    "Contratada deverá entregar a documentação comprobatória do cumprimento da obrigação "
    "contratual;")
add_bloco_variante("S09_TipoFiscalizacao", {
    "Fiscal Funcional": [
        "no prazo de até 07 (sete) dias corridos a partir do recebimento dos documentos da "
        "Contratada, o Fiscal Funcional deverá realizar a análise de toda a documentação "
        "apresentada, avaliar a execução por meio do Instrumento de Medição de Resultado (IMR) e "
        "emitir Termo de Recebimento Provisório, encaminhando-o ao empregado ou comissão "
        "designada para o recebimento definitivo.",
    ],
    "Comissão de Fiscalização": [
        "no prazo de até 07 (sete) dias corridos a partir do recebimento dos documentos da "
        "Contratada: o Fiscal Funcional deverá realizar a análise de toda a documentação "
        "apresentada e emitir relatório; o Fiscal Técnico deverá realizar a análise de toda a "
        "documentação apresentada e emitir relatório; o Fiscal Setorial deverá realizar a análise "
        "de toda a documentação apresentada e emitir relatório;",
        "o Fiscal Funcional, no prazo de até 03 (três) dias úteis a partir do recebimento do "
        "último relatório, avaliará a execução por meio do Instrumento de Medição de Resultado "
        "(IMR) e emitirá Termo de Recebimento Provisório com base nos relatórios emitidos pelos "
        "fiscais, encaminhando-o, caso não haja irregularidades, ao empregado ou comissão "
        "designada para o recebimento definitivo.",
    ],
})
doc.add_paragraph(
    "12.3. Constatadas impropriedades que impeçam a liquidação e o pagamento da despesa, estas "
    "serão registradas no Termo de Recebimento Provisório, com prazo de até 05 (cinco) dias úteis "
    "para regularização; sanadas as impropriedades, será elaborado relatório detalhado e "
    "encaminhado o Termo de Recebimento Provisório ao empregado ou Comissão designada, que "
    "providenciará o recebimento definitivo em até 05 (cinco) dias corridos, emitindo Termo "
    "Circunstanciado e comunicando a Contratada para emissão da Nota Fiscal ou Fatura, com o "
    "valor exato dimensionado pela fiscalização com base no IMR.")
doc.add_paragraph(
    "12.4. Antes de cada pagamento será realizada consulta ao SICAF; havendo resultado "
    "desfavorável, será concedido prazo de 05 (cinco) dias úteis para regularização ou defesa, "
    "prorrogável uma vez por igual período.")

# 13. DO REAJUSTE
add_heading("13. DO REAJUSTE")
add_bloco_variante("S13_TipoCorrecao", {
    "Preço fixo (irreajustável)": ["13.1. O preço é fixo e irreajustável."],
    "Reajuste (sem dedicação exclusiva de mão de obra)": [[
        "13.1. O preço consignado no Contrato será corrigido anualmente, observado o "
        "interregno mínimo de um ano, contado da data limite para apresentação da "
        "proposta, pela variação do índice de referência: ",
        ("S13_IndiceReajuste", "____"), ".",
    ]],
    "Repactuação (com dedicação exclusiva de mão de obra)": [[
        "13.1. Por se tratar de serviço com dedicação exclusiva de mão de obra, o preço "
        "consignado no Contrato será objeto de repactuação (art. 3º, LXXXIII, e art. 501 do "
        "RLC), observado o interregno mínimo de um ano contado da data-base da proposta "
        "para os custos decorrentes do mercado, e do Acordo, Convenção ou Dissídio "
        "Coletivo de Trabalho para os custos decorrentes da mão de obra. Acordo, Convenção "
        "ou Dissídio Coletivo de Trabalho de referência: ",
        ("S13_AcordoConvencaoDissidio", "____"), ".",
    ]],
})
doc.add_paragraph(
    "13.2. O reajuste ou a repactuação pressupõe interregno mínimo de um ano contado da data "
    "limite para apresentação da proposta, só sendo cabível, portanto, quando a vigência do "
    "Contrato — inclusive prorrogações — alcançar 12 (doze) meses.")
doc.add_paragraph(
    "13.3. Serão objeto de preclusão os reajustes ou repactuações a que a Contratada fizer "
    "jus durante a vigência do Contrato e que não forem solicitados até o implemento dos "
    "seguintes eventos: a) assinatura do termo aditivo de prorrogação contratual; b) data em "
    "que o Contrato completa 12 (doze) meses, 24 (vinte e quatro) meses e assim "
    "sucessivamente; ou c) encerramento do Contrato.")
doc.add_paragraph(
    "13.4. O disposto nesta seção não afasta o direito ao reequilíbrio econômico-financeiro "
    "em sentido estrito, cabível a qualquer tempo, independentemente de previsão contratual, "
    "quando verificado fato imprevisível ou de consequências incalculáveis, anormal ou "
    "extraordinário, nos termos do art. 511 do RLC.")

# 14. DA SUBCONTRATAÇÃO
add_heading("14. DA SUBCONTRATAÇÃO")
add_bloco_variante("S14_PermiteSubcontratacao", {
    "Sim": ["14.1. É permitida a subcontratação parcial do objeto entre os limites mínimo e "
            "máximo, respectivamente, do valor total do Contrato."],
    "Não": ["14.1. Não será admitida a subcontratação do objeto desta dispensa de licitação."],
})
add_paragraph_with_bookmarks([
    "Limites: entre ", ("S14_PercentualMinSubcontratacao", "____"), "% e ",
    ("S14_PercentualMaxSubcontratacao", "____"), "% do valor total do Contrato.",
])
add_paragraph_with_bookmarks(["Detalhamento: ", ("S14_DetalheSubcontratacao", "")])
doc.add_paragraph("14.2. A subcontratação deverá observar o disposto no art. 524 do RLC.")

# 15. DA ALTERAÇÃO SUBJETIVA (texto fixo)
add_heading("15. DA ALTERAÇÃO SUBJETIVA")
doc.add_paragraph(
    "15.1. É admissível a fusão, cisão ou incorporação da Contratada com/em outra pessoa "
    "jurídica, desde que observados os requisitos de habilitação, mantidas as demais cláusulas e "
    "condições do Contrato, não haja prejuízo à execução do objeto pactuado e haja a anuência "
    "expressa da Conab.")

# 16. DAS SANÇÕES ADMINISTRATIVAS
add_heading("16. DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "16.1. A contratada, em caso de inadimplemento de suas obrigações, garantido o contraditório "
    "e ampla defesa, ficará sujeita às seguintes sanções previstas no RLC e na Lei nº 13.303, de "
    "2016: advertência; multa moratória; multa compensatória; multa rescisória, para os casos de "
    "rescisão unilateral por descumprimento contratual; e suspensão do direito de participar de "
    "licitação e impedimento de contratar com a Conab, por até 02 (dois) anos.")
doc.add_paragraph(
    "a) advertência, aplicável nas infrações de menor gravidade, sem prejuízo à execução do "
    "objeto;")
add_paragraph_with_bookmarks([
    "b) multa por não regularização da documentação de habilitação: ",
    ("S16_MultaNaoRegularizacaoHabilitacao", "____"), "% sobre o valor estimado para a contratação.",
])
add_paragraph_with_bookmarks([
    "c) multa pela prática das condutas elencadas nos arts. 576 a 580 do RLC: ",
    ("S16_MultaCondutasArt576a580", "____"), "% sobre o valor estimado para a contratação.",
])
add_paragraph_with_bookmarks([
    "d) multa pela recusa em assinar o Contrato dentro do prazo estabelecido: ",
    ("S16_MultaRecusaAssinarContrato", "____"), "% sobre o valor da contratação.",
])
doc.add_paragraph("e) multa moratória por atraso injustificado na entrega da garantia contratual, conforme item 8.2.")
add_paragraph_with_bookmarks([
    "f) multa moratória de 0,2% (dois décimos por cento) sobre o ",
    ("S16_BaseCalculoMulta_E", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso na execução dos serviços, até o limite de 15 (quinze) dias.",
])
add_paragraph_with_bookmarks([
    "g) multa moratória de 0,3% (três décimos por cento) sobre o ",
    ("S16_BaseCalculoMulta_F", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso na execução dos serviços, por período superior ao previsto na alínea "
    "anterior, até o limite de 15 (quinze) dias. Esgotado o prazo, poderá ocorrer a não "
    "aceitação do objeto, sem prejuízo da rescisão unilateral da avença.",
])
add_paragraph_with_bookmarks([
    "h) multa compensatória no percentual de 10% (dez por cento) sobre o ",
    ("S16_BaseCalculoMulta_G", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", no caso de inexecução parcial do Contrato.",
])
add_paragraph_with_bookmarks([
    "i) multa compensatória de ", ("S16_MultaCompensatoriaTotal", "____"),
    "% sobre o valor total do Contrato, no caso de inexecução total do Contrato.",
])
add_paragraph_with_bookmarks([
    "j) multa rescisória de ", ("S16_MultaRescisoria", "____"),
    "% sobre o valor total do Contrato, no caso de rescisão contratual unilateral.",
])
add_paragraph_with_bookmarks([
    "k) multa de 0,2% a 3,2% por dia sobre o ",
    ("S16_BaseCalculoMulta_TABELA2", "(valor do Contrato, valor anual ou valor mensal)"),
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
    ("Deixar de cumprir quaisquer itens do Termo de Referência e seus Anexos não previstos nesta "
     "tabela, após reincidência formalmente notificada pela fiscalização, por item e por "
     "ocorrência", "3"),
    ("Deixar de indicar e manter durante a execução do Contrato os prepostos previstos no "
     "Contrato, por dia", "1"),
    ("Deixar de providenciar treinamento para seus funcionários conforme previsto na relação de "
     "obrigações da Contratada, por funcionário", "1"),
]
for descricao, grau in infracoes_fixas:
    linha = tabela1.add_row()
    linha.cells[0].text = descricao
    linha.cells[1].text = grau
add_linha_modelo_tabela(tabela1, "S16_InfracoesExtras_LINHA_MODELO", ["((DESCRICAO))", "((GRAU))"])

tabela2 = doc.add_table(rows=1, cols=2)
tabela2.style = "Table Grid"
tabela2.rows[0].cells[0].text = "GRAU"
tabela2.rows[0].cells[1].text = "CORRESPONDÊNCIA (percentual por dia)"
for grau, percentual in [("1", "0,2%"), ("2", "0,4%"), ("3", "0,8%"), ("4", "1,6%"), ("5", "3,2%")]:
    linha = tabela2.add_row()
    linha.cells[0].text = grau
    linha.cells[1].text = percentual
doc.add_paragraph(
    "16.2. As multas moratória, compensatória e rescisória possuem fatos geradores distintos; se "
    "aplicadas duas multas sobre o mesmo fato gerador, configura bis in idem.")
doc.add_paragraph(
    "16.3. Da sanção de suspensão: cabe em razão de ação ou omissão capaz de causar, ou que "
    "tenha causado, prejuízo à Conab, suas instalações, pessoas, imagem, meio ambiente, ou em "
    "decorrência de determinação legal, aplicada por até 02 (dois) anos, nos termos dos arts. "
    "579 a 580 do RLC.")

# 17. DA RESCISÃO CONTRATUAL (texto fixo)
add_heading("17. DA RESCISÃO CONTRATUAL")
doc.add_paragraph(
    "17.1. A inexecução total do Contrato ensejará a sua rescisão, enquanto a inexecução parcial "
    "poderá ensejar a sua rescisão, com as consequências cabíveis, conforme disposto nos arts. "
    "568 a 572 do RLC.")
doc.add_paragraph(
    "17.2. A rescisão poderá ser: a) por ato unilateral e escrito da Conab; b) amigável, por "
    "acordo entre as partes, reduzida a termo no processo de contratação, desde que haja "
    "conveniência para a Conab; e c) judicial, por determinação judicial.")
doc.add_paragraph(
    "17.2.1. A rescisão amigável deverá ser precedida de autorização escrita e fundamentada da "
    "autoridade competente.")
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
    "prejuízo das sanções previstas neste Termo de Referência: a) assunção imediata do objeto "
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
add_paragraph_with_bookmarks(["19.3. ", ("S19_CriteriosEspecificos", "")])

# 20. DA MATRIZ DE RISCOS E RESPONSABILIDADES
add_heading("20. DA MATRIZ DE RISCOS E RESPONSABILIDADES")
add_bloco_variante("S20_ExigeMatrizRiscos", {
    "Sim": [
        "20.1. A Matriz de Riscos é a cláusula contratual definidora dos riscos e das "
        "responsabilidades entre o Contratante e a Contratada e caracterizadora do equilíbrio "
        "econômico financeiro na execução do Contrato, em termos de ônus financeiros decorrentes "
        "de eventos supervenientes à contratação.",
        "20.2. A Contratada é integral e exclusivamente responsável por todos os riscos e "
        "responsabilidades relacionados ao objeto do ajuste, conforme hipóteses não-exaustivas "
        "elencadas na Matriz de Riscos — Anexo I do Termo de Referência, e não é responsável "
        "pelos riscos e responsabilidades quando estes competirem à Contratante, conforme "
        "estabelecido no mesmo Anexo.",
    ],
    "Não": [
        "20.1. A presente contratação é composta por serviços de pronto pagamento, os quais não "
        "apresentam nenhum tipo de risco nas fases de planejamento, seleção e gestão do contrato, "
        "ficando assim dispensada a apresentação de matriz de riscos para este Termo de "
        "Referência.",
    ],
})

# 21. DA ALTERAÇÃO DO CONTRATO / 22. DAS VEDAÇÕES / 23. DA VINCULAÇÃO / 24. LEGISLAÇÃO / 25. FORO
# Bloco inteiro (5 seções, sem tabela no meio) só existe quando o
# instrumento contratual é "Contrato Simplificado" — ver Seção 7.
_bid_secoes_21a25 = add_bloco_variante_sentinela_inicio(
    "S07_TipoInstrumentoContratual_SECOES21A25_BLOCO_CONTRATO_SIMPLIFICADO")

add_heading("21. DA ALTERAÇÃO DO CONTRATO")
doc.add_paragraph(
    "21.1. O Contrato Simplificado decorrente deste Termo de Referência poderá ser alterado nas "
    "hipóteses previstas no art. 510 do RLC, vedando-se alterações que resultem em violação ao "
    "dever de licitar. A Contratada poderá aceitar, nas mesmas condições contratuais, os "
    "acréscimos ou supressões que se fizerem necessários, até 25% (vinte e cinco por cento) do "
    "valor inicial atualizado do contrato. Fica vedada a celebração de termos aditivos "
    "decorrentes de eventos supervenientes alocados, na Matriz de Riscos, como de "
    "responsabilidade da Contratada, que somente poderá solicitar reequilíbrio "
    "econômico-financeiro nas hipóteses excluídas de suas responsabilidades.")

add_heading("22. DAS VEDAÇÕES")
doc.add_paragraph(
    "22.1. É vedado à Contratada: a) caucionar ou utilizar o Contrato Simplificado decorrente "
    "deste Termo de Referência para qualquer operação financeira; b) interromper a execução "
    "contratual sob alegação de inadimplemento por parte da Contratante, salvo nos casos "
    "previstos em lei; c) empregar menor de 18 anos em trabalho noturno, perigoso ou "
    "insalubre e não empregar menor de 16 anos, salvo na condição de aprendiz a partir de 14 "
    "anos; d) subcontratar, na íntegra ou parcialmente, o objeto da contratação; e e) "
    "solicitar, por meio de seus empregados, serviços, materiais ou equipamentos a empresas "
    "terceirizadas que prestem serviços à Conab (art. 154 do RLC).")

add_heading("23. DA VINCULAÇÃO E DA MANUTENÇÃO DA QUALIFICAÇÃO E HABILITAÇÃO")
doc.add_paragraph(
    "23.1. Consideram-se integrantes do Contrato Simplificado o Termo de Referência da Dispensa "
    "de Licitação e seus Anexos, a Proposta da Contratada, no que couber, e demais documentos "
    "pertinentes, independentemente de transcrição. A Contratada se obriga a manter, durante toda "
    "a execução do objeto, as condições de habilitação e qualificação exigidas neste Termo de "
    "Referência.")

add_heading("24. DA LEGISLAÇÃO APLICÁVEL")
doc.add_paragraph(
    "24.1. Os casos omissos serão decididos pela Contratante, segundo as disposições contidas no "
    "RLC, na Lei nº 13.303, de 2016, e demais normas federais de licitações e contratos "
    "administrativos aplicáveis às empresas públicas e, subsidiariamente, segundo as disposições "
    "do Código de Defesa do Consumidor e normas e princípios gerais dos contratos.")

add_heading("25. DO FORO")
doc.add_paragraph(
    "25.1. As partes elegem o foro da Justiça Federal, Seção Judiciária do Estado da Sureg ou de "
    "Brasília-DF, no caso da Matriz, competente para dirimir quaisquer dúvidas suscitadas em "
    "razão do Contrato Simplificado decorrente deste Termo de Referência, que não puderem ser "
    "resolvidas de comum acordo.")

add_bloco_variante_sentinela_fim(_bid_secoes_21a25)

# 26. DAS DISPOSIÇÕES FINAIS
add_heading("26. DAS DISPOSIÇÕES FINAIS")
doc.add_paragraph(
    "26.1. Integram este Termo de Referência os seguintes Anexos: a) Anexo I - Matriz de Riscos "
    "(quando aplicável); b) Anexo II - Modelo de Apresentação de Proposta; c) Anexo III - "
    "Declaração de Vistoria Técnica; d) Anexo IV - Modelo de Instrumento de Medição de Resultados "
    "(IMR); e e) o instrumento contratual (Contrato Simplificado ou Minuta de Contrato, conforme "
    "escolhido na Seção 7).")
add_paragraph_with_bookmarks(["Elaborador — Área Demandante: ", ("S26_ElabDemandanteNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S26_ElabDemandanteSetor", ""), " — Cargo: ", ("S26_ElabDemandanteCargo", "")])
add_paragraph_with_bookmarks(["Elaborador — Área Técnica: ", ("S26_ElabTecnicaNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S26_ElabTecnicaSetor", ""), " — Cargo: ", ("S26_ElabTecnicaCargo", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Demandante: ", ("S26_AprovDemandanteNome", ""), " — Setor: ", ("S26_AprovDemandanteSetor", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Técnica: ", ("S26_AprovTecnicaNome", ""), " — Setor: ", ("S26_AprovTecnicaSetor", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Demandante: ", ("S26_SuperDemandanteNome", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Técnica: ", ("S26_SuperTecnicaNome", "")])

# ANEXO I - MATRIZ DE RISCOS (só existe se S20_ExigeMatrizRiscos = "Sim";
# Anexo II do RLC, transcrito integralmente)
doc.add_page_break()
_bid_anexo_i = add_bloco_variante_sentinela_inicio("S20_ExigeMatrizRiscos_ANEXO_I_BLOCO_SIM")
add_heading("ANEXO I DO TERMO DE REFERÊNCIA")
add_heading("MATRIZ DE RISCOS")
add_matriz_de_riscos_rlc("S01_DescricaoObjeto_AnexoI")
add_bloco_variante_sentinela_fim(_bid_anexo_i)

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
for _ in range(4):
    tabela_proposta.add_row()
doc.add_paragraph("VALOR TOTAL DOS ITENS: R$")

# ANEXO III - MODELO DE DECLARAÇÃO DE VISTORIA
doc.add_page_break()
add_heading("ANEXO III DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE DECLARAÇÃO DE VISTORIA")
doc.add_paragraph("À Companhia Nacional de Abastecimento – CONAB")
doc.add_paragraph("Ref.: Dispensa de Licitação – nº ___/20XX")
doc.add_paragraph(
    "Declaramos que, em atendimento ao previsto na Dispensa de Licitação nº ___/_____, eu, "
    "______________________________________, portador(a) da CI/RG nº______________ e do CPF nº "
    "_____________________, responsável da empresa "
    "___________________________________________________, como seu representante legal para os "
    "fins da presente declaração, compareci perante a CONAB/Sureg/XX, situada à "
    "_____________________________________________________________________________________, e "
    "vistoriei os locais onde serão executados os serviços objeto da Dispensa de Licitação em "
    "apreço, tomando plena ciência das condições e graus de dificuldade existentes.")
doc.add_paragraph("Cidade-UF, ___/___/20XX.")
doc.add_paragraph("Assinatura e carimbo (Responsável Técnico da Empresa)")
doc.add_paragraph("Visto: Responsável pela CONAB")

# ANEXO IV - MODELO DE IMR
doc.add_page_break()
add_heading("ANEXO IV DO TERMO DE REFERÊNCIA")
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

# ANEXO V - CONTRATO SIMPLIFICADO (só existe se Contrato Simplificado)
doc.add_page_break()
_bid_anexo_v = add_bloco_variante_sentinela_inicio(
    "S07_TipoInstrumentoContratual_ANEXO_BLOCO_CONTRATO_SIMPLIFICADO")
add_heading("ANEXO V DO TERMO DE REFERÊNCIA")
add_heading("MINUTA DE CONTRATO SIMPLIFICADO")
# Formulário do Anexo XI do RLC (Contrato Simplificado), com as 11 cláusulas
# de termos e condições transcritas integralmente a partir de
# anexos RLC/anexo_xi_contrato_simplificadoodt.odt.
# Os 32 campos numerados são preenchidos no momento da emissão do contrato
# (dados do fornecedor vencedor, itens, valores) — não no momento da
# elaboração do TR — por isso permanecem em branco no modelo, como no
# formulário oficial.
add_tabela(["CONTRATO SIMPLIFICADO", "1. Número", "2. Data"])
add_tabela(["3. Área Demandante", "4. Processo N.º", "5. CNPJ", "6. Inscrição Estadual"])
doc.add_paragraph("DADOS DO FORNECEDOR")
add_tabela(["7. Razão Social"])
add_tabela(["8. Endereço", "9. CEP", "10. (DDD) Telefone"])
add_tabela(["11. Município/Cidade", "12. UF", "13. CNPJ/CPF", "14. Inscrição Estadual"])
doc.add_paragraph(
    "Na forma estabelecida no processo licitatório e/ou condições constantes no verso "
    "deste Contrato Simplificado, solicitamos fornecimento dos materiais/serviços abaixo "
    "discriminados:")
doc.add_paragraph("MATERIAIS/SERVIÇOS")
tabela_itens_cs = add_tabela(
    ["15. Item", "16. Especificação Detalhada do Material ou Serviço", "17. Unid.",
     "18. Quant.", "19. Valor Unitário (R$)", "Valor Total (R$)"],
    linhas_vazias=10,
)
linha_total_cs = tabela_itens_cs.add_row()
linha_total_cs.cells[0].text = "20. Total Geral (R$) ="
add_tabela(["21. Amparo Legal"])
doc.add_paragraph("CONDIÇÕES")
add_tabela(["22. Local de Entrega"])
add_tabela(["23. Prazo de Entrega", "24. Garantia", "25. Condições de Pagamento", "26. Frete ( ) CIF ( ) FOB"])
doc.add_paragraph("DESPESA EMPENHADA")
add_tabela(["27. Número", "28. Data"])
add_tabela(
    ["29. Responsável por", "30. Aprovado por", "31. Autorizado por",
     "32. Estamos de acordo com as condições estabelecidas. Data: ___/___/______ "
     "Assinatura do Fornecedor: _________________"],
)
doc.add_paragraph(
    "60.000/066 — 1.ª via - Fornecedor; 2.ª via - Emitente; 3.ª via - Área Financeira; "
    "4.ª via - Almoxarifado; 5.ª via - Processo")
doc.add_paragraph("O presente CONTRATO SIMPLIFICADO obedecerá aos seguintes termos e condições:")
doc.add_paragraph(
    "1. O Fornecedor deverá acusar, prontamente, o recebimento deste, comprometendo-se, para "
    "todos os fins, a aceitar os termos e condições do presente Contrato Simplificado e do "
    "respectivo Termo de Referência ou Projeto Básico.")
doc.add_paragraph(
    "2. Concorda o Fornecedor que todas as mercadorias ou serviços especificados no presente "
    "Contrato Simplificado são de primeira qualidade, por ele garantidas como tais. "
    "Responsabiliza-se, ainda, pelos vícios e defeitos, visíveis ou ocultos, que desde já são "
    "considerados redibitórios na forma da lei, obrigando-se, se assim desejar a Conab, a repor "
    "as mercadorias ou refazer os serviços que não satisfaçam, plenamente, as especificações e "
    "indicações apresentadas pela Conab, livres de qualquer pagamento ou despesa extra e no prazo "
    "que lhe for designado.")
doc.add_paragraph(
    "3. O Fornecedor assume toda a responsabilidade por qualquer interpelação ou reivindicação "
    "que venha a ser feita, envolvendo patentes ou direitos autorais, correndo, igualmente, por "
    "conta do Fornecedor quaisquer pagamentos, multas ou obrigações exigidas por alegados ou "
    "possíveis titulares desses direitos.")
doc.add_paragraph(
    "4. As mercadorias ou serviços a serem fornecidos, segundo desenhos e especificações da "
    "Conab, não poderão levar marca comercial do Fornecedor por meio de inscrição, gravura ou "
    "outra qualquer forma de identificação, a não ser que, expressamente e por escrito, seja isso "
    "convencionado.")
doc.add_paragraph(
    "5. A responsabilidade da Conab na execução do presente limita-se e extingue-se ao preço e "
    "seu pagamento, sendo da exclusiva responsabilidade do Fornecedor todas e quaisquer despesas "
    "ou ônus, inclusive trabalhistas e de previdência social, não se compreendendo a operação "
    "realizada através desta como subcontratação, ou subempreitada.")
doc.add_paragraph(
    "6. Os serviços serão prestados diretamente à Conab nos locais por ela indicados. A prestação "
    "coberta pelo presente Contrato Simplificado somente se entende realizada após a verificação "
    "de que os serviços satisfazem os requisitos, especificações e indicações da Conab. O "
    "Fornecedor responsabiliza-se pelos serviços realizados e pelas mercadorias que oferece para "
    "a sua execução, bem como pelos seguros e embalagens, sempre que expressamente e por escrito "
    "não se convencione o contrário.")
doc.add_paragraph(
    "7. Nenhum adicional de preço será devido pela Conab a título de frete ou embalagem, a não "
    "ser que, expressamente e por escrito, seja convencionado.")
doc.add_paragraph(
    "8. Todos os documentos de embarque, conhecimentos de transporte e listas acaso expedidos "
    "para a execução do serviço deverão conter o número do presente Contrato Simplificado e ser "
    "enviados à Conab na data do embarque da mercadoria, ou entregues com esta, no que couber. "
    "Todos os sacos, caixas ou embalagens deverão levar o número do presente Contrato Simplificado "
    "e o nome do transportador e, sempre que possível, papeletas de material forte, seguramente "
    "presas aos volumes.")
doc.add_paragraph(
    "9. A data do recebimento dos serviços prestados ou das mercadorias pela Conab fixará o termo "
    "inicial para forma de pagamento que haja sido contratada, salvo quando expressamente "
    "convencionado de outra forma.")
doc.add_paragraph(
    "10. O prazo de entrega deverá ser rigorosamente observado, sob pena de aplicação das "
    "sanções administrativas previstas no Termo de Referência respectivo.")
doc.add_paragraph(
    "11. O pagamento do presente Contrato Simplificado far-se-á via Ordem Bancária ou Ordem de "
    "Pagamento.")
add_bloco_variante_sentinela_fim(_bid_anexo_v)

# ANEXO VI - MINUTA DE CONTRATO (só existe se Contrato completo)
doc.add_page_break()
_bid_anexo_vi = add_bloco_variante_sentinela_inicio(
    "S07_TipoInstrumentoContratual_ANEXO_BLOCO_CONTRATO")
add_heading("ANEXO VI DO TERMO DE REFERÊNCIA")
add_heading("MINUTA DE CONTRATO")
add_paragraph_with_bookmarks([
    "Processo Administrativo nº ", ("S_AnexoContrato_NumeroProcesso", "____"),
    ". Contrato nº ", ("S_AnexoContrato_NumeroContrato", "____"), ".",
])
doc.add_paragraph(
    "CONTRATANTE: (unidade da Conab — Matriz ou Sureg —, CNPJ, endereço e representante "
    "legal com nome, cargo e CPF)")
doc.add_paragraph(
    "CONTRATADA: (razão social), (CNPJ, endereço e representante legal com nome e CPF)")
add_paragraph_with_bookmarks([
    "CONTRATO QUE ENTRE SI CELEBRAM A COMPANHIA NACIONAL DE ABASTECIMENTO - CONAB E A EMPRESA "
    "___________________ PARA A PRESTAÇÃO DE SERVIÇOS DE ___________________, de conformidade "
    "com o Processo Administrativo nº ", ("S_AnexoContrato_NumeroProcesso_Inline", "____"),
    " referente à Dispensa de Licitação, regendo-se pelo Termo de "
    "Referência e seus anexos e pela proposta da Contratada, pelo RLC, pela Lei nº 13.303, de "
    "2016, e demais legislações pertinentes.",
])
doc.add_paragraph("CLÁUSULA PRIMEIRA - DO OBJETO")
doc.add_paragraph(
    "O objeto do presente instrumento é a contratação de serviços, que serão prestados nas "
    "condições estabelecidas no Termo de Referência, ao qual este Termo de Contrato se vincula, "
    "independentemente de transcrição, juntamente com a proposta vencedora.")
doc.add_paragraph("CLÁUSULA SEGUNDA - DA VIGÊNCIA")
add_bloco_variante("S04_TipoVigencia_ANEXO_VI", {
    "Até 1 ano, prorrogável": [[
        "O prazo de vigência do Contrato será de ", ("S04_PrazoVigenciaMeses_ANEXO_VI_Inline", "____"),
        " meses, contados da data de sua "
        "assinatura, prorrogável desde que observados os requisitos do art. 488 do RLC, entre "
        "eles a demonstração da vantajosidade na manutenção do ajuste (em regra, por pesquisa "
        "de preços), sob pena de a Conab não poder prorrogar o Contrato.",
    ]],
    "Plurianual (até 5 anos)": [[
        "O prazo de vigência do Contrato será de ", ("S04_PrazoVigenciaAnos_ANEXO_VI_Inline", "____"),
        " anos, contados da data de sua "
        "assinatura, podendo ser prorrogado até o limite de 05 (cinco) anos, observados os "
        "requisitos do art. 488 do RLC, entre eles a demonstração da vantajosidade na "
        "manutenção do ajuste (em regra, por pesquisa de preços), sob pena de a Conab não "
        "poder prorrogar o Contrato.",
    ]],
})
add_paragraph_with_bookmarks(["Prazo de vigência (meses): ", ("S04_PrazoVigenciaMeses_ANEXO_VI", "____")])
add_paragraph_with_bookmarks(["Prazo de vigência (anos): ", ("S04_PrazoVigenciaAnos_ANEXO_VI", "____")])
doc.add_paragraph("CLÁUSULA TERCEIRA - DO REGIME DE EXECUÇÃO, DOS PRAZOS DE EXECUÇÃO E DA GARANTIA DOS SERVIÇOS")
doc.add_paragraph(
    "Os serviços serão executados conforme o regime, os prazos de execução e as condições de "
    "garantia estabelecidos no Termo de Referência, admitida prorrogação do prazo de execução "
    "observado o disposto nos arts. 497 e 498 do RLC.")
doc.add_paragraph("CLÁUSULA QUARTA - DO RECEBIMENTO DOS SERVIÇOS")
doc.add_paragraph(
    "O recebimento provisório e definitivo dos serviços observará o disposto no Termo de "
    "Referência e nos arts. 526 e 527 do RLC, não excluindo a responsabilidade da Contratada "
    "pelos prejuízos resultantes da incorreta execução do Contrato.")
add_paragraph_with_bookmarks([
    "Prazo para recebimento provisório: ", ("S06_PrazoRecebimentoProvisorioDias_ANEXO_VI", "____"),
    " dias. Prazo para recebimento definitivo, contado do provisório: ",
    ("S06_PrazoRecebimentoDefinitivoDias_ANEXO_VI", "____"), " dias.",
])
doc.add_paragraph("CLÁUSULA QUINTA - DO VALOR DO CONTRATO")
doc.add_paragraph(
    "O valor do Contrato corresponde ao valor proposto pela Contratada vencedora, incluídas todas "
    "as despesas ordinárias diretas e indiretas decorrentes da execução do objeto, inclusive "
    "tributos, encargos sociais, trabalhistas e previdenciários.")
doc.add_paragraph("CLÁUSULA SEXTA - DA GARANTIA CONTRATUAL")
add_bloco_variante("S08_ExigeGarantia_ANEXO_VI", {
    "Sim": [[
        "A Contratada deverá apresentar, no prazo de 10 (dez) dias corridos após a assinatura "
        "do Contrato, prorrogáveis por igual período, comprovante de prestação de garantia "
        "correspondente a ", ("S08_PercentualGarantia_ANEXO_VI_Inline", "____"),
        "%, conforme parâmetros do art. 439 do RLC, "
        "mediante caução em dinheiro, seguro-garantia ou fiança bancária.",
    ]],
    "Não": ["Não será exigida garantia contratual, para a prestação dos serviços em apreço."],
})
add_paragraph_with_bookmarks(["Percentual da garantia: ", ("S08_PercentualGarantia_ANEXO_VI", "____"), "%"])
add_paragraph_with_bookmarks(["A garantia incide sobre: ", ("S08_BaseCalculoGarantia_ANEXO_VI", "")])
doc.add_paragraph(
    "O percentual da garantia é, em regra, de até 5% do valor do Contrato (art. 439, §2º, do "
    "RLC), podendo ser elevado para até 10% em contratações de grande vulto envolvendo "
    "complexidade técnica e riscos financeiros elevados (art. 439, §3º, do RLC). Nos casos de "
    "contratação de serviços com dedicação exclusiva de mão de obra, o valor da garantia "
    "deverá corresponder a 5% do valor do contrato, limitado ao equivalente a 2 (dois) meses "
    "do custo da folha de pagamento dos empregados da contratada que venham a participar da "
    "execução dos serviços contratados (art. 439, §4º, do RLC).")
doc.add_paragraph("CLÁUSULA SÉTIMA - DA DOTAÇÃO ORÇAMENTÁRIA")
add_paragraph_with_bookmarks([
    "A despesa orçamentária da execução deste Contrato correrá à conta da Natureza da Despesa (",
    ("S01_ND_ANEXO_VI", "XXXX"), "), PTRES (", ("S01_PTRES_ANEXO_VI", "XXXX"), ") e Fonte (",
    ("S01_Fonte_ANEXO_VI", "XXXX"), "), conforme Nota de Empenho a ser emitida.",
])
doc.add_paragraph("CLÁUSULA OITAVA - DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph(
    "São obrigações da Contratante as estabelecidas no Termo de Referência, não respondendo a "
    "Conab por compromissos assumidos pela Contratada com terceiros.")
doc.add_paragraph("CLÁUSULA NONA - DAS OBRIGAÇÕES DA CONTRATADA")
doc.add_paragraph("São obrigações da Contratada as estabelecidas no Termo de Referência.")
add_bloco_variante("S11_NaturezaIntelectual_ANEXO_VI", {
    "Sim": [
        "Ceder os direitos patrimoniais relativos ao projeto ou serviço técnico especializado, "
        "para que a Conab possa utilizá-lo, de acordo com o art. 525 do RLC, e assegurar à "
        "Contratante o direito de propriedade intelectual dos produtos desenvolvidos e os "
        "direitos autorais da solução, do projeto e da documentação produzida, realizando a "
        "transição contratual com transferência de conhecimento e tecnologia, nos termos do "
        "art. 17 do RLC.",
    ],
    "Não": [""],
})
doc.add_paragraph("CLÁUSULA DÉCIMA - DA SUSTENTABILIDADE AMBIENTAL")
doc.add_paragraph(
    "A Contratada observará os critérios de sustentabilidade ambiental previstos no Termo de "
    "Referência e no art. 10 do RLC, respondendo administrativa, civil e penalmente por danos "
    "causados ao meio ambiente.")
add_paragraph_with_bookmarks(["Critérios específicos: ", ("S19_CriteriosEspecificos_ANEXO_VI", "")])
doc.add_paragraph("CLÁUSULA DÉCIMA PRIMEIRA - DA MANUTENÇÃO DA QUALIFICAÇÃO E HABILITAÇÃO")
doc.add_paragraph(
    "A Contratada se obriga a manter, durante toda a execução do objeto, as condições de "
    "habilitação e qualificação exigidas para esta contratação.")
doc.add_paragraph("CLÁUSULA DÉCIMA SEGUNDA - DO ACOMPANHAMENTO E DA FISCALIZAÇÃO DO CONTRATO")
doc.add_paragraph("O acompanhamento e a fiscalização do Contrato observarão o previsto no Termo de Referência.")
doc.add_paragraph("CLÁUSULA DÉCIMA TERCEIRA - DO PAGAMENTO")
doc.add_paragraph("O prazo para pagamento e demais condições a ele referentes se encontram definidos no Termo de Referência.")
doc.add_paragraph(
    "Na inexistência de outra regra contratual, em caso de atraso de pagamento provocado "
    "exclusivamente pela Conab, o valor devido será acrescido de atualização financeira desde a "
    "data do vencimento até a do efetivo pagamento, com juros de mora de 0,5% (meio por cento) "
    "ao mês, ou 6% (seis por cento) ao ano, conforme o art. 565 do RLC.")
doc.add_paragraph("CLÁUSULA DÉCIMA QUARTA - DO REAJUSTE")
add_bloco_variante("S13_TipoCorrecao_ANEXO_VI", {
    "Preço fixo (irreajustável)": ["O preço é fixo e irreajustável."],
    "Reajuste (sem dedicação exclusiva de mão de obra)": [[
        "O preço consignado no Contrato será corrigido anualmente, observado o interregno "
        "mínimo de um ano, contado da data limite para apresentação da proposta, pela "
        "variação do índice de referência: ",
        ("S13_IndiceReajuste_ANEXO_VI", "____"), ".",
    ]],
    "Repactuação (com dedicação exclusiva de mão de obra)": [[
        "Por se tratar de serviço com dedicação exclusiva de mão de obra, o preço "
        "consignado no Contrato será objeto de repactuação (art. 3º, LXXXIII, e art. 501 do "
        "RLC), observado o interregno mínimo de um ano contado da data-base da proposta "
        "para os custos decorrentes do mercado, e do Acordo, Convenção ou Dissídio "
        "Coletivo de Trabalho para os custos decorrentes da mão de obra. Acordo, Convenção "
        "ou Dissídio Coletivo de Trabalho de referência: ",
        ("S13_AcordoConvencaoDissidio_ANEXO_VI", "____"), ".",
    ]],
})
doc.add_paragraph(
    "O reajuste ou a repactuação pressupõe interregno mínimo de um ano contado da data "
    "limite para apresentação da proposta, só sendo cabível, portanto, quando a vigência do "
    "Contrato — inclusive prorrogações — alcançar 12 (doze) meses.")
doc.add_paragraph(
    "Serão objeto de preclusão os reajustes ou repactuações a que a Contratada fizer jus "
    "durante a vigência do Contrato e que não forem solicitados até o implemento dos "
    "seguintes eventos: a) assinatura do termo aditivo de prorrogação contratual; b) data em "
    "que o Contrato completa 12 (doze) meses, 24 (vinte e quatro) meses e assim "
    "sucessivamente; ou c) encerramento do Contrato.")
doc.add_paragraph(
    "O disposto nesta cláusula não afasta o direito ao reequilíbrio econômico-financeiro em "
    "sentido estrito, cabível a qualquer tempo, independentemente de previsão contratual, "
    "quando verificado fato imprevisível ou de consequências incalculáveis, anormal ou "
    "extraordinário, nos termos do art. 511 do RLC.")
doc.add_paragraph("CLÁUSULA DÉCIMA QUINTA - DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "A Contratada, em caso de inadimplemento de suas obrigações, garantido o contraditório e "
    "ampla defesa, ficará sujeita às sanções previstas no RLC, na Lei nº 13.303, de 2016, e no "
    "Termo de Referência: advertência; multa moratória; multa compensatória; multa rescisória; e "
    "suspensão do direito de participar de licitação e impedimento de contratar com a Conab, por "
    "até 02 (dois) anos.")
doc.add_paragraph(
    "a) advertência, aplicável nas infrações de menor gravidade, sem prejuízo à execução do "
    "objeto;")
add_paragraph_with_bookmarks([
    "b) multa por não regularização da documentação de habilitação: ",
    ("S16_MultaNaoRegularizacaoHabilitacao_ANEXO_VI", "____"), "% sobre o valor estimado.",
])
add_paragraph_with_bookmarks([
    "c) multa pela prática das condutas elencadas nos arts. 576 a 580 do RLC: ",
    ("S16_MultaCondutasArt576a580_ANEXO_VI", "____"), "% sobre o valor estimado.",
])
add_paragraph_with_bookmarks([
    "d) multa pela recusa em assinar o Contrato dentro do prazo estabelecido: ",
    ("S16_MultaRecusaAssinarContrato_ANEXO_VI", "____"), "% sobre o valor da contratação.",
])
doc.add_paragraph(
    "e) multa moratória por atraso injustificado na entrega da garantia contratual, conforme "
    "Cláusula Sexta.")
add_paragraph_with_bookmarks([
    "f) multa moratória de 0,2% sobre o ",
    ("S16_BaseCalculoMulta_ANEXO_VI_E", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso, até o limite de 15 (quinze) dias.",
])
add_paragraph_with_bookmarks([
    "g) multa moratória de 0,3% sobre o ",
    ("S16_BaseCalculoMulta_ANEXO_VI_F", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso, por período superior ao previsto na alínea anterior, até o limite de "
    "15 (quinze) dias.",
])
add_paragraph_with_bookmarks([
    "h) multa compensatória de 10% sobre o ",
    ("S16_BaseCalculoMulta_ANEXO_VI_G", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", no caso de inexecução parcial do Contrato.",
])
add_paragraph_with_bookmarks([
    "i) multa compensatória de ", ("S16_MultaCompensatoriaTotal_ANEXO_VI", "____"),
    "% sobre o valor total do Contrato, no caso de inexecução total do Contrato.",
])
add_paragraph_with_bookmarks([
    "j) multa rescisória de ", ("S16_MultaRescisoria_ANEXO_VI", "____"),
    "% sobre o valor total do Contrato, no caso de rescisão contratual unilateral.",
])
add_paragraph_with_bookmarks([
    "k) multa de 0,2% a 3,2% por dia sobre o ",
    ("S16_BaseCalculoMulta_ANEXO_VI_TABELA2", "(valor do Contrato, valor anual ou valor mensal)"),
    ", conforme grau da infração cometida (Tabelas 1 e 2 abaixo).",
])
tabela1_anexo_vi = doc.add_table(rows=1, cols=2)
tabela1_anexo_vi.style = "Table Grid"
tabela1_anexo_vi.rows[0].cells[0].text = "TABELA 1 — DESCRIÇÃO DA INFRAÇÃO"
tabela1_anexo_vi.rows[0].cells[1].text = "GRAU"
for descricao, grau in infracoes_fixas:
    linha = tabela1_anexo_vi.add_row()
    linha.cells[0].text = descricao
    linha.cells[1].text = grau
add_linha_modelo_tabela(tabela1_anexo_vi, "S16_InfracoesExtras_LINHA_MODELO", ["((DESCRICAO))", "((GRAU))"])
tabela2_anexo_vi = doc.add_table(rows=1, cols=2)
tabela2_anexo_vi.style = "Table Grid"
tabela2_anexo_vi.rows[0].cells[0].text = "GRAU"
tabela2_anexo_vi.rows[0].cells[1].text = "CORRESPONDÊNCIA (percentual por dia)"
for grau, percentual in [("1", "0,2%"), ("2", "0,4%"), ("3", "0,8%"), ("4", "1,6%"), ("5", "3,2%")]:
    linha = tabela2_anexo_vi.add_row()
    linha.cells[0].text = grau
    linha.cells[1].text = percentual
doc.add_paragraph("CLÁUSULA DÉCIMA SEXTA - DA RESCISÃO")
doc.add_paragraph(
    "A rescisão observará o disposto no Termo de Referência e nos arts. 568 a 572 e 582 a 593 do "
    "RLC.")
doc.add_paragraph("CLÁUSULA DÉCIMA SÉTIMA - DA SUBCONTRATAÇÃO")
doc.add_paragraph("As condições de subcontratação, quando admitidas, observam o disposto no Termo de Referência.")
doc.add_paragraph("CLÁUSULA DÉCIMA OITAVA - DA MATRIZ DE RISCOS E RESPONSABILIDADES")
doc.add_paragraph(
    "A Matriz de Riscos, quando aplicável — Anexo I do Termo de Referência —, constitui peça "
    "integrante deste Contrato, independentemente de transcrição, definindo os riscos e "
    "responsabilidades entre Contratante e Contratada.")
doc.add_paragraph("CLÁUSULA DÉCIMA NONA - DA ALTERAÇÃO DO CONTRATO")
doc.add_paragraph(
    "O presente Contrato poderá ser alterado nas hipóteses do art. 510 do RLC, vedando-se "
    "alterações que resultem em violação ao dever de licitar. A Contratada poderá aceitar "
    "acréscimos ou supressões de até 25% (vinte e cinco por cento) do valor inicial atualizado do "
    "contrato.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA - DA ALTERAÇÃO SUBJETIVA")
doc.add_paragraph(
    "É admissível a fusão, cisão ou incorporação da Contratada com/em outra pessoa jurídica, "
    "observados os requisitos de habilitação e a anuência expressa da Conab.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA PRIMEIRA - DAS VEDAÇÕES")
doc.add_paragraph(
    "É vedado à Contratada: a) caucionar ou utilizar este Contrato para qualquer operação "
    "financeira; b) interromper a execução contratual sob alegação de inadimplemento da "
    "Contratante, salvo nos casos previstos em lei; c) empregar menor de 18 anos em trabalho "
    "noturno, perigoso ou insalubre e não empregar menor de 16 anos, salvo na condição de "
    "aprendiz a partir de 14 anos; d) subcontratar, na íntegra ou parcialmente, o objeto da "
    "contratação, salvo quando a subcontratação parcial for admitida no Termo de Referência; "
    "e e) solicitar, por meio de seus empregados, serviços, materiais ou equipamentos a "
    "empresas terceirizadas que prestem serviços à Conab (art. 154 do RLC).")
doc.add_paragraph("CLÁUSULA VIGÉSIMA SEGUNDA - DA VEDAÇÃO AO NEPOTISMO")
doc.add_paragraph(
    "Conforme disposto no art. 12, §1º, inciso II, do RLC e no art. 7º do Decreto nº 7.203, de "
    "2010, fica vedada a contratação de empregado ou dirigente da Conab como pessoa física, ou de "
    "pessoa/empresa com relação de parentesco até o terceiro grau civil com autoridade do "
    "Ministério do Desenvolvimento Agrário e Agricultura Familiar (MDA), dirigente ou empregado "
    "da Conab com atribuições na área de licitação e contratação.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA TERCEIRA - DA VINCULAÇÃO")
doc.add_paragraph(
    "Consideram-se integrantes do presente Instrumento Contratual o Termo de Referência da "
    "Dispensa de Licitação e seus Anexos, a Proposta da Contratada, no que couber, e demais "
    "documentos pertinentes, independentemente de transcrição.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA QUARTA - DA LEGISLAÇÃO APLICÁVEL")
doc.add_paragraph(
    "Os casos omissos serão decididos pela Contratante, segundo as disposições do RLC, da Lei nº "
    "13.303, de 2016, e demais normas aplicáveis, e, subsidiariamente, do Código de Defesa do "
    "Consumidor.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA QUINTA - DA PROTEÇÃO DE DADOS PESSOAIS")
doc.add_paragraph(
    "As Partes obrigam-se a cumprir a legislação vigente sobre proteção de dados pessoais, "
    "em especial a Lei nº 13.709, de 2018 (Lei Geral de Proteção de Dados Pessoais), no que "
    "for aplicável à execução deste Contrato.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA SEXTA - DA PUBLICAÇÃO")
doc.add_paragraph(
    "A publicação do extrato do presente Contrato deverá ser providenciada pela Contratante até o "
    "5º (quinto) dia útil do mês subsequente ao da sua assinatura no Diário Oficial da União, "
    "conforme o art. 480 do RLC.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA SÉTIMA - DO FORO")
doc.add_paragraph(
    "As partes elegem o foro da Justiça Federal, Seção Judiciária do Estado da Sureg ou de "
    "Brasília-DF, no caso da Matriz, competente para dirimir quaisquer dúvidas suscitadas em "
    "razão deste Contrato, que não puderem ser resolvidas de comum acordo.")
doc.add_paragraph(
    "Por estarem justas e acordadas, firmam o presente instrumento, em 03 (três) vias, de igual "
    "teor e forma, diante das testemunhas abaixo identificadas.")
doc.add_paragraph("Pela Contratante: ___________________     Pela Contratada: ___________________")
doc.add_paragraph("Testemunha 1: ___________________     Testemunha 2: ___________________")
add_bloco_variante_sentinela_fim(_bid_anexo_vi)

doc.save("Template_TR_Dispensa.docx")
print("Template_TR_Dispensa.docx gerado com sucesso.")
