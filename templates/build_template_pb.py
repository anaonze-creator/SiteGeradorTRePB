"""Gera Template_PB.docx a partir do modelo institucional de Projeto Básico
da GEPEA (skill gepea-instrumentos, modelo_institucional_projeto_basico.txt),
com bookmarks nos pontos de preenchimento identificados em js/dados-pb.js.
Rodar novamente sempre que o modelo institucional mudar."""
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
    """Marca o INÍCIO de um bloco que mistura parágrafos e tabelas (ex.:
    Anexo II-A inteiro). Combinar com add_bloco_variante_sentinela_fim(bid)
    depois de todo o conteúdo ter sido adicionado."""
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
titulo_run = titulo.add_run("PROJETO BÁSICO")
titulo_run.bold = True

# 1. DA DESCRIÇÃO DO OBJETO DA CONTRATAÇÃO
add_heading("1. DA DESCRIÇÃO DO OBJETO DA CONTRATAÇÃO (ART. 157, I, RLC)")
add_paragraph_with_bookmarks([
    "1.1. O presente Projeto Básico tem por objeto a contratação de ",
    ("PB01_DescricaoObjeto", "(descrição do objeto)"),
    ", conforme especificações, condições, quantidades e exigências estabelecidas neste instrumento.",
])
add_bloco_variante("PB01_TipoObjeto", {
    "Obra": ["1.2. Os serviços, materiais e/ou obras referenciados neste Projeto Básico "
             "enquadram-se como obra, cujos padrões de desempenho e qualidade podem ser "
             "objetivamente definidos por meio de especificações usuais de mercado."],
    "Serviço comum de engenharia": [
        "1.2. Os serviços, materiais e/ou obras referenciados neste Projeto Básico "
        "enquadram-se como serviço comum de engenharia, cujos padrões de desempenho e "
        "qualidade podem ser objetivamente definidos por meio de especificações usuais de "
        "mercado.",
    ],
    "Serviço não comum de engenharia": [
        "1.2. Os serviços, materiais e/ou obras referenciados neste Projeto Básico "
        "enquadram-se como serviço não comum de engenharia.",
    ],
})
add_paragraph_with_bookmarks([
    "1.3. Os serviços serão executados indiretamente no regime de ",
    ("PB01_RegimeExecucao", "(regime de execução)"), ".",
])
add_paragraph_with_bookmarks([
    "1.4. O critério de julgamento das propostas será o de ",
    ("PB01_CriterioJulgamento", "(menor preço global ou maior desconto)"),
    ". O valor estimado/valor de referência será ", ("PB01_SigiloValor", "(público ou sigiloso)"),
    ". O modo de disputa adotado será ", ("PB01_ModoDisputa", "(aberto, fechado ou aberto e fechado)"), ".",
])
add_bloco_variante("PB01_TIC", {
    "Sim": [
        "Por se tratar de contratação de Tecnologia da Informação, esta contratação deve "
        "estar prevista no Plano Diretor de Tecnologia da Informação e Comunicação (PDTIC) "
        "ou ser previamente aprovada pelo Comitê Estratégico de Tecnologia da Informação, "
        "nos termos dos arts. 599-600 do RLC, observadas a Resolução CGPAR nº 29/2022 e as "
        "orientações do Acórdão nº 980/2023-TCU."
    ],
    "Não": [""],
})
add_paragraph_with_bookmarks([
    "1.5. O custo total estimado para a prestação dos serviços deste Projeto Básico é de R$ ",
    ("PB01_ValorEstimado", "(valor)"), " (", ("PB01_ValorPorExtenso", "valor por extenso"), ").",
])

# 2. DO OBJETIVO DA CONTRATAÇÃO E SUA JUSTIFICATIVA
add_heading("2. DO OBJETIVO DA CONTRATAÇÃO E SUA JUSTIFICATIVA (ART. 157, II, RLC)")
add_paragraph_with_bookmarks(["2.1. A contratação tem por objetivo ", ("PB02_Objetivo", "")])
add_paragraph_with_bookmarks(["2.2. A necessidade da contratação justifica-se por ", ("PB02_Necessidade", "")])
add_paragraph_with_bookmarks(["2.3. A justificativa das especificações técnicas adotadas decorre de ", ("PB02_JustificativaEspecificacoes", "")])
add_paragraph_with_bookmarks(["2.4. O quantitativo demandado foi calculado com base em ", ("PB02_CalculoQuantitativo", "")])
add_bloco_variante("PB02_ModalidadeOuFundamento", {
    "Licitação": [[
        "2.5. O serviço em apreço será contratado na modalidade ",
        ("PB02_ModalidadeLicitatoria", "(modalidade)"), ", nos termos do art. 3º do RLC.",
    ]],
    "Dispensa de Licitação": [
        "2.5. A contratação será realizada por dispensa de licitação, com fundamento no "
        "art. 416, I, do RLC, por se tratar de obra ou serviço de engenharia de valor "
        "global até R$ 140.000,00 (cento e quarenta mil reais).",
    ],
})
add_bloco_variante("PB02_ImpactoOrcamentarioFinanceiro", {
    "Sim": [
        "2.6. Nos termos do art. 216, IX, do RLC e do art. 16, incisos I e II, da Lei "
        "Complementar nº 101, de 2000 (LRF), a presente contratação gera despesa "
        "obrigatória de caráter continuado ou impacto orçamentário-financeiro relevante, "
        "devendo constar dos autos a estimativa do impacto orçamentário-financeiro e a "
        "declaração do ordenador de despesa quanto à adequação orçamentária e financeira "
        "com a Lei Orçamentária Anual (LOA) e à compatibilidade com o Plano Plurianual "
        "(PPA) e a Lei de Diretrizes Orçamentárias (LDO)."
    ],
    "Não": [
        "2.6. Nos termos do art. 16, §3º, da Lei Complementar nº 101, de 2000 (LRF), a "
        "presente contratação não se enquadra no caput daquele artigo, dispensando a "
        "estimativa do impacto orçamentário-financeiro e a declaração de adequação "
        "orçamentária ali previstas."
    ],
})
_bid_controles_dispensa = add_bloco_variante_sentinela_inicio(
    "PB02_ControlesDispensa_BLOCO_DISPENSA_DE_LICITACAO")
add_paragraph_with_bookmarks([
    "2.7. Nos termos do art. 416, inciso I, do RLC, foi observado o limite de valor "
    "considerando o somatório do valor desta contratação com o valor de outros objetos de "
    "mesma natureza contratados pela mesma unidade gestora no mesmo exercício financeiro, "
    "não havendo fracionamento de despesa: ",
    ("PB02_SomatorioContratacoesMesmaNatureza", "(descrever a apuração do somatório e a conclusão quanto ao limite)"),
])
doc.add_paragraph(
    "2.8. A presente contratação será precedida de divulgação de aviso em sítio eletrônico "
    "oficial, pelo prazo mínimo de 3 (três) dias úteis, com atendimento à Instrução "
    "Normativa SEGES/ME nº 67, de 2021, para busca da proposta mais vantajosa.")
add_bloco_variante("PB02_PagamentoCartao", {
    "Sim": [
        "2.9. A contratação será paga por meio de cartão de pagamento, com divulgação do "
        "respectivo extrato no Portal Nacional de Contratações Públicas (PNCP)."
    ],
    "Não": [[
        "2.9. A presente contratação não será paga por meio de cartão de pagamento, pelas "
        "seguintes razões: ", ("PB02_JustificativaNaoCartao", ""),
    ]],
})
add_bloco_variante_sentinela_fim(_bid_controles_dispensa)

# 3. DA CARACTERIZAÇÃO DO OBJETO
add_heading("3. DA CARACTERIZAÇÃO DO OBJETO (ART. 157, III, RLC)")
doc.add_paragraph("3.1. As características essenciais da contratação são as seguintes:")
add_bloco_variante("PB03_ModalidadeOuFundamento", {
    "Licitação": [["Modalidade licitatória: ", ("PB03_ModalidadeLicitatoria", "")]],
    "Dispensa de Licitação": [
        "Fundamento da dispensa: art. 416, I, do RLC (obras e serviços de engenharia de "
        "valor global até R$ 140.000,00).",
    ],
})
add_paragraph_with_bookmarks(["Regime de execução: ", ("PB03_RegimeExecucao", "")])
add_paragraph_with_bookmarks(["Critério de julgamento: ", ("PB03_CriterioJulgamento", "")])
add_paragraph_with_bookmarks(["Serviço continuado: ", ("PB03_ServicoContinuado", "")])
add_paragraph_with_bookmarks(["Participação ME/EPP: ", ("PB03_ParticipacaoMEEPP", "")])
add_paragraph_with_bookmarks(["Tipo de objeto: ", ("PB03_TipoObjeto", "")])
add_paragraph_with_bookmarks([
    "3.2. A forma de contratação, o regime adotado e o critério de julgamento justificam-se, "
    "ainda, por ", ("PB03_JustificativaForma", ""),
])

# 4. DAS ESPECIFICAÇÕES DO OBJETO
add_heading("4. DAS ESPECIFICAÇÕES DO OBJETO (ART. 157, IV, RLC)")
add_paragraph_with_bookmarks(["4.1. A solução técnica adotada compreende ", ("PB04_SolucaoTecnica", "")])
doc.add_paragraph(
    "4.2. As especificações são suficientes para evitar alterações, no decorrer do "
    "empreendimento, que desvirtuem a solução escolhida, nos termos do art. 157, IV, \"c\", do "
    "RLC.")
doc.add_paragraph("4.3. As especificações técnicas detalhadas por sistema constam do Anexo IX deste Projeto Básico.")
add_bloco_variante("PB04_AdotaBIM", {
    "Sim": [
        "4.4. O Decreto nº 10.306/2020 e a Estratégia BIM BR (Decreto nº 11.888/2024) "
        "estabelecem a adoção progressiva da metodologia BIM na Administração Pública "
        "Federal. Quando aplicável, o BIM abrangerá, no mínimo: modelos com extração de "
        "quantitativos e documentação gráfica; orçamentação e planejamento; e atualização "
        "as-built ao término. Os modelos serão entregues em formato IFC; os direitos "
        "autorais pertencem à União; os custos de capacitação da Contratada são de sua "
        "exclusiva responsabilidade.",
    ],
    "Não": [""],
})
add_paragraph_with_bookmarks(["4.5. Locais de execução: ", ("PB04_LocaisExecucao", "")])

# 5. DAS ESPECIFICAÇÕES TÉCNICAS E ANEXOS TÉCNICOS
add_heading("5. DAS ESPECIFICAÇÕES TÉCNICAS E ANEXOS TÉCNICOS (ART. 157, IV, RLC)")
doc.add_paragraph(
    "5.1. Na execução dos serviços deverão ser observadas, além das especificações deste "
    "Projeto Básico e dos projetos, as normas e especificações da ABNT.")
doc.add_paragraph(
    "5.2. Demais referências normativas eventualmente propostas deverão atender, "
    "primeiramente, às normas internas da Conab; na sua ausência, às normas da ABNT; e, "
    "subsidiariamente, às normas estrangeiras.")
doc.add_paragraph(
    "5.3. Integram este Projeto Básico, para todos os fins, os anexos técnicos aplicáveis à "
    "natureza e à complexidade do objeto: projeto técnico (plantas, cortes, diagramas, "
    "especificações de materiais); estudos, levantamentos e diagnósticos (cadastrais, "
    "fotográficos, laudos); quantitativos e base de orçamento (memória de quantitativos, "
    "memória de cálculo, critérios de medição); e demais documentos de suporte.")

# 6. DA DEFINIÇÃO DE MÉTODOS E ESTRATÉGIA DE SUPRIMENTO
add_heading("6. DA DEFINIÇÃO DE MÉTODOS E ESTRATÉGIA DE SUPRIMENTO (ART. 157, V, RLC)")
doc.add_paragraph(
    "6.1. Nos termos do art. 157, V, do RLC, este item apresenta os subsídios para a "
    "montagem do plano de licitação e gestão da obra ou serviço de engenharia, "
    "compreendendo a programação da execução, a estratégia de suprimentos de materiais e "
    "mão de obra, as normas de fiscalização aplicáveis e demais dados necessários ao caso "
    "concreto.")
add_paragraph_with_bookmarks(["6.2. ", ("PB06_MetodoConstrutivo", "")])

# 7. DO PRAZO DE EXECUÇÃO
add_heading("7. DO PRAZO DE EXECUÇÃO (ART. 157, VI, RLC)")
add_paragraph_with_bookmarks([
    "7.1. Os serviços somente poderão ser iniciados após o recebimento da Ordem de Serviço "
    "emitida pela Conab, nos termos do art. 155 do RLC, devendo a execução ter início em até ",
    ("PB07_PrazoInicioDias", "____"), " dias, contados do recebimento da Ordem de Serviço.",
])
add_paragraph_with_bookmarks([
    "7.2. O prazo total de execução é de ", ("PB07_PrazoTotalExecucao", "(dias corridos ou meses)"),
    ". Justificativa: ", ("PB07_JustificativaPrazo", ""),
])
doc.add_paragraph(
    "7.3. Para os serviços por demanda/eventuais, o prazo de cada execução será definido "
    "conforme o cronograma de cada demanda, não podendo ultrapassar o limite máximo da "
    "vigência contratual.")

# 8. DA VIGÊNCIA CONTRATUAL
add_heading("8. DA VIGÊNCIA CONTRATUAL (ART. 157, VII, RLC)")
add_paragraph_with_bookmarks([
    "8.1. O prazo de vigência do Contrato será de ", ("PB08_PrazoVigencia", "(meses ou dias)"),
    ", contados da data de sua assinatura, nos termos dos arts. 461 e 462 do RLC.",
])
doc.add_paragraph(
    "8.2. A vigência contempla o prazo de execução acrescido dos prazos de recebimento "
    "provisório e definitivo, bem como de eventual prorrogação prevista no RLC.")
doc.add_paragraph(
    "8.3. Eventual prorrogação da vigência deverá observar os requisitos do art. 488 do RLC, "
    "entre eles a demonstração da vantajosidade na manutenção do ajuste (em regra, por "
    "pesquisa de preços), sob pena de a Conab não poder prorrogar o Contrato.")
add_paragraph_with_bookmarks([
    "8.4. Quando cabível, a prorrogação poderá se estender por ",
    ("PB08_PrazoProrrogacao", "____"), ".",
])

# 9. DA VISTORIA TÉCNICA
add_heading("9. DA VISTORIA TÉCNICA (ART. 157, VIII, RLC)")
add_bloco_variante("PB09_TipoVistoria", {
    "Obrigatória": [
        "9.1. Para o correto dimensionamento e elaboração da proposta, o licitante deverá "
        "realizar vistoria nas instalações do local de execução dos serviços, tomando "
        "ciência das características e peculiaridades existentes.",
    ],
    "Facultativa": [
        "9.1. Para o correto dimensionamento e elaboração da proposta, recomenda-se que o "
        "licitante realize vistoria nas instalações do local de execução dos serviços, "
        "tomando ciência das características e peculiaridades existentes, sendo a "
        "realização facultativa.",
    ],
})
add_paragraph_with_bookmarks([
    "9.2. A vistoria deverá ser marcada e realizada em dias úteis, das ", ("PB09_HorarioVistoria", "(horário)"),
    ", mediante agendamento prévio pelo telefone ", ("PB09_TelefoneVistoria", "(telefone)"),
    " ou pelo e-mail ", ("PB09_EmailVistoria", "(e-mail)"), ".",
])
add_bloco_variante("PB09_VistoriaPrazo", {
    "Licitação": [
        "9.3. O prazo para vistoria iniciar-se-á no dia útil seguinte ao da publicação do "
        "Edital e estender-se-á até o dia útil anterior à data prevista para a abertura da "
        "sessão pública.",
    ],
    "Dispensa de Licitação": [
        "9.3. O prazo para a vistoria será definido no processo de dispensa de licitação, "
        "conforme cronograma estabelecido pela área demandante.",
    ],
})
add_bloco_variante("PB09_TipoVistoria_Anexo", {
    "Obrigatória": [
        "9.4. O licitante apresentará a Declaração de Vistoria Técnica (Anexo IV), comprovando "
        "a realização da vistoria obrigatória de que trata esta Seção.",
    ],
    "Facultativa": [
        "9.4. Optando por realizar a vistoria, o licitante apresentará a Declaração de Vistoria "
        "Técnica (Anexo IV); optando por não realizá-la, apresentará a Declaração de Dispensa de "
        "Vistoria Técnica (Anexo V), assumindo integral responsabilidade por eventuais erros de "
        "dimensionamento da proposta.",
    ],
})

# 10. DA AMOSTRA OU DEMONSTRAÇÃO PRÉVIA DOS SERVIÇOS
add_heading("10. DA AMOSTRA OU DEMONSTRAÇÃO PRÉVIA DOS SERVIÇOS (ART. 157, IX, RLC)")
add_bloco_variante("PB10_ExigeAmostra", {
    "Sim": [""],
    "Não": ["10.1. Considerando a natureza da contratação, não será exigida amostra ou demonstração prévia."],
})
add_paragraph_with_bookmarks(["10.1. ", ("PB10_DetalheAmostra", "")])

# 11. DOS CRITÉRIOS DE SIMILARIDADE
add_heading("11. DOS CRITÉRIOS DE SIMILARIDADE (ART. 157, X, RLC)")
add_bloco_variante("PB11_ExigeSimilaridade", {
    "Sim": [""],
    "Não": ["11.1. Considerando a natureza da contratação, item não aplicável."],
})
add_paragraph_with_bookmarks(["11.1. ", ("PB11_DetalheSimilaridade", "")])

# 12. DA ANOTAÇÃO DE RESPONSABILIDADE TÉCNICA
add_heading("12. DA ANOTAÇÃO DE RESPONSABILIDADE TÉCNICA (ART. 157, XI, RLC)")
add_paragraph_with_bookmarks([
    "12.1. A Contratada deverá apresentar à Conab, em até ", ("PB12_PrazoARTDias", "____"),
    " dias corridos da assinatura do Contrato, as Anotações de Responsabilidade Técnica "
    "(ART) ou Registros de Responsabilidade Técnica (RRT) registrados no CREA/CAU, em nome "
    "do(s) responsável(is) técnico(s).",
])
doc.add_paragraph("12.2. Os custos das ART/RRT emitidas pela Contratada durante a execução são de sua exclusiva responsabilidade.")
add_paragraph_with_bookmarks([
    "12.3. A Contratada apresentará a baixa da ART/RRT em até ", ("PB12_PrazoBaixaARTDias", "____"),
    " dias úteis após a finalização das obras ou serviços.",
])
doc.add_paragraph(
    "12.4. É dever do gestor exigir a ART/RRT do empregado da Conab responsável pela "
    "elaboração do Projeto Básico/Executivo, pela execução, supervisão e fiscalização. Os "
    "custos dessas ART/RRT, emitidas a serviço da Companhia, são por ela suportados.")

# 13. DA HABILITAÇÃO — QUALIFICAÇÃO TÉCNICA E OPERACIONAL
add_heading("13. DA HABILITAÇÃO — QUALIFICAÇÃO TÉCNICA E OPERACIONAL (ART. 157, XII, RLC)")
doc.add_paragraph("13.1. Qualificação técnica")
doc.add_paragraph(
    "13.1.1. Certidão de registro de pessoa jurídica no CREA/CAU, em nome da licitante, com "
    "validade na data de recebimento dos documentos de habilitação.")
add_paragraph_with_bookmarks([
    "13.1.2. Certidão de Acervo Técnico (CAT) em nome do(s) responsável(is) técnico(s), "
    "comprovando responsabilidade por serviços de características semelhantes às parcelas "
    "de maior relevância técnica e econômica do objeto: ", ("PB13_ParcelasRelevantes", ""), ".",
])
add_paragraph_with_bookmarks([
    "13.1.3. Atestado(s) de capacidade técnica emitido(s) por pessoa jurídica de direito "
    "público ou privado, comprovando execução de serviços semelhantes: ", ("PB13_QuantitativoMinimo", ""),
    " Admite-se o somatório de atestados.",
])
doc.add_paragraph(
    "13.1.4. As exigências observam as vedações do art. 157, XII, do RLC, sendo expressamente "
    "vedadas: exigência de certificados ISO, cartas de solidariedade ou documentos de "
    "garantia de qualidade; exigência de visto/registro local de CREA/CAU no momento da "
    "habilitação; fixação de número mínimo de atestados ou limitação de tempo de "
    "experiência; percentuais mínimos acima de 50% dos itens de maior relevância; e "
    "restrição a tipologias específicas de obras/serviços.")
doc.add_paragraph("13.2. Qualificação econômico-financeira")
add_paragraph_with_bookmarks([
    "13.2.1. Certidão negativa de efeitos de falência, recuperação judicial ou "
    "extrajudicial, expedida pelo distribuidor da sede do licitante, com prazo de validade "
    "de ", ("PB13_PrazoValidadeCertidaoFalencia", "____"), " dias da emissão.",
])
doc.add_paragraph(
    "13.2.2. Balanço Patrimonial e demonstrações contábeis do último exercício social, "
    "comprovando índices de Liquidez Geral (LG), Solvência Geral (SG) e Liquidez Corrente "
    "(LC) superiores a 1 (um).")
add_paragraph_with_bookmarks([
    "13.2.3. Capital Circulante Líquido de, no mínimo, ", ("PB13_CapitalCirculanteMinimo", "____"),
    "% do valor estimado da contratação.",
])
add_paragraph_with_bookmarks([
    "13.2.4. Patrimônio Líquido de, no mínimo, ", ("PB13_PatrimonioLiquidoMinimo", "____"),
    "% do valor estimado da contratação.",
])

# 14. DOS CRITÉRIOS E PRÁTICAS DE SUSTENTABILIDADE
add_heading("14. DOS CRITÉRIOS E PRÁTICAS DE SUSTENTABILIDADE (ART. 157, XIII, RLC)")
doc.add_paragraph(
    "14.1. A Contratada observará os critérios de sustentabilidade do art. 157, XIII, do "
    "RLC, especialmente: especificações que visem à economia de manutenção e operação, à "
    "redução do consumo de energia e água e ao uso de tecnologias e materiais de menor "
    "impacto ambiental; uso de produtos de limpeza e conservação conformes às classificações "
    "da ANVISA; medidas para evitar o desperdício de água tratada; e fornecimento e "
    "exigência de uso dos Equipamentos de Proteção Individual necessários.")
add_bloco_variante("PB14_CriteriosObra", {
    "Execução de obra ou serviço já projetado": [
        "14.1.1. Adicionalmente, por se tratar de execução de obra ou serviço de engenharia: "
        "uso obrigatório de agregados reciclados, sempre que houver oferta, capacidade de "
        "suprimento e custo inferior aos naturais; e cumprimento do PGRCC (Resolução CONAMA "
        "nº 307/2002), com Controle de Transporte de Resíduos para todos os resíduos "
        "removidos.",
    ],
})
doc.add_paragraph(
    "14.2. A Contratada é responsável, administrativa, civil e penalmente, por qualquer dano "
    "ambiental decorrente da execução, sendo vedado o descarte irregular de resíduos.")
add_paragraph_with_bookmarks(["14.3. ", ("PB14_CriteriosEspecificos", "")])

# 15. DOS CRITÉRIO DE ACEITAÇÃO DO OBJETO E IMR
add_heading("15. DOS CRITÉRIOS DE ACEITAÇÃO DO OBJETO E INSTRUMENTO DE MEDIÇÃO DE RESULTADOS (ART. 157, XIV, RLC)")
doc.add_paragraph("15.1. Somente será aceito o objeto em consonância com este Projeto Básico quanto a quantidade, qualidade e funcionamento.")
add_paragraph_with_bookmarks([
    "15.2. Os serviços serão executados conforme as normas técnicas vigentes, em especial: ",
    ("PB15_NormasTecnicas", ""), "; a Lei nº 13.303, de 2016; e o RLC 10.901 da Conab.",
])
doc.add_paragraph(
    "15.3. Durante a vigência, a Conab adotará o Instrumento de Medição de Resultados (IMR), "
    "constante do Anexo VII, com critérios objetivos para efeito de pagamento, podendo "
    "ensejar multa e rescisão.")

# 16. DO RECEBIMENTO DO OBJETO CONTRATUAL
add_heading("16. DO RECEBIMENTO DO OBJETO CONTRATUAL (ART. 157, XV, RLC)")
doc.add_paragraph("16.1. O recebimento provisório ou definitivo não exclui a responsabilidade da Contratada pelos prejuízos resultantes da incorreta execução do Contrato.")
add_paragraph_with_bookmarks([
    "16.2. O recebimento provisório será realizado pela Fiscalização em até ", ("PB16_PrazoRecebimentoProvisorioDias", "____"),
    " dias corridos do adimplemento da parcela/entrega, formalizado em Termo de Recebimento "
    "Provisório detalhado, encaminhado ao responsável pelo recebimento definitivo.",
])
add_paragraph_with_bookmarks([
    "16.3. O recebimento definitivo, que concretiza o ateste da execução, será realizado "
    "pelo empregado ou comissão designada em até ", ("PB16_PrazoRecebimentoDefinitivoDias", "____"),
    " dias corridos do recebimento provisório, após verificação de toda a documentação.",
])
doc.add_paragraph("16.4. O recebimento definitivo não exclui a responsabilidade civil por vícios e defeitos ocultos, na forma do Código Civil.")
doc.add_paragraph(
    "16.5. O recebimento definitivo ficará condicionado à verificação, pela Fiscalização ou "
    "pela comissão designada, da completa entrega da documentação técnica e administrativa "
    "aplicável ao objeto: ART/RRT de execução e respectiva baixa, quando exigível; projetos "
    "e desenhos as built; manuais de operação, uso e manutenção; certificados de ensaios, "
    "comissionamento, calibração, conformidade ou garantia; relatórios técnicos; "
    "documentação exigida pelas concessionárias ou órgãos competentes; e demais documentos "
    "previstos neste Projeto Básico, seus anexos ou nas normas técnicas aplicáveis.")

# 17. DAS OBRIGAÇÕES DA CONTRATANTE
add_heading("17. DAS OBRIGAÇÕES DA CONTRATANTE (ART. 157, XVI, RLC)")
doc.add_paragraph("17.1. São obrigações da Contratante:")
doc.add_paragraph("a) exigir o cumprimento de todas as obrigações assumidas pela Contratada, conforme as cláusulas contratuais e os termos da proposta;")
doc.add_paragraph("b) receber o objeto no prazo e nas condições estabelecidas;")
doc.add_paragraph("c) verificar minuciosamente, no prazo fixado, a conformidade dos serviços recebidos provisoriamente, para fins de aceitação e recebimento definitivo;")
doc.add_paragraph("d) comunicar à Contratada, por escrito, imperfeições, falhas ou irregularidades, para reparo ou correção;")
doc.add_paragraph("e) rejeitar, no todo ou em parte, serviço executado em desacordo com este Projeto Básico;")
doc.add_paragraph("f) acompanhar e fiscalizar a execução por meio de empregado ou comissão designados;")
doc.add_paragraph("g) efetuar o pagamento no prazo e na forma estabelecidos;")
doc.add_paragraph("h) não praticar atos de ingerência na administração da Contratada (poder de mando sobre empregados, direcionamento de contratações, desvio de função, etc.); e")
doc.add_paragraph("i) providenciar as ART dos empregados da Conab responsáveis pela supervisão e fiscalização, arcando com os respectivos custos.")
doc.add_paragraph("17.2. A Conab não responderá por compromissos assumidos pela Contratada com terceiros, ainda que vinculados à execução do Contrato.")

# 18. DAS OBRIGAÇÕES DA CONTRATADA
add_heading("18. DAS OBRIGAÇÕES DA CONTRATADA (ART. 157, XVII, RLC)")
add_paragraph_with_bookmarks([
    "18.1. A Contratada deve cumprir todas as obrigações constantes neste Projeto, seus "
    "Anexos e sua proposta, assumindo como exclusivamente seus os riscos e as despesas "
    "decorrentes da boa e perfeita execução do objeto, e ainda:"
])
add_paragraph_with_bookmarks([
    "a) apresentar as ART/RRT registradas no CREA/CAU em até ", ("PB18_PrazoARTContratoDias", "____"),
    " dias corridos da assinatura do Contrato;",
])
doc.add_paragraph("b) executar os serviços conforme este Projeto Básico e sua proposta, com a alocação dos profissionais necessários;")
doc.add_paragraph("c) iniciar os serviços somente após o recebimento da Ordem de Serviço emitida pela Conab;")
doc.add_paragraph("d) seguir o Cronograma Físico-Financeiro, comunicando imediatamente qualquer impossibilidade de cumprimento;")
doc.add_paragraph("e) responsabilizar-se pelos vícios e danos decorrentes da execução, na forma da legislação aplicável;")
doc.add_paragraph("f) proteger adequadamente os bens da Conab no local e restituí-los ao término;")
doc.add_paragraph("g) submeter à aprovação prévia da Conab toda substituição de materiais e equipamentos de referência;")
doc.add_paragraph("h) manter sigilo sobre assuntos de interesse da Conab ou de terceiros acessados na execução;")
doc.add_paragraph("i) manter, durante toda a execução, as condições de habilitação e qualificação exigidas, inclusive a regularidade no SICAF;")
doc.add_paragraph("j) cumprir as obrigações trabalhistas, sociais, previdenciárias e tributárias, inclusive as previstas em Acordo/Convenção/Dissídio Coletivo;")
doc.add_paragraph("k) vedar o trabalho de menor de 16 anos (salvo aprendiz a partir de 14) e o trabalho noturno, perigoso ou insalubre de menor de 18 anos;")
doc.add_paragraph("l) indicar e manter preposto para representá-la durante a execução;")
doc.add_paragraph("m) arcar com o ônus de eventual equívoco no dimensionamento dos quantitativos de sua proposta, ressalvadas as hipóteses do RLC;")
doc.add_paragraph("n) obedecer às normas oficiais e às recomendações dos fabricantes, fornecendo materiais, mão de obra qualificada, ferramentas, equipamentos e EPIs;")
doc.add_paragraph("o) não se prevalecer de erro ou omissão deste Projeto Básico para eximir-se de responsabilidades;")
doc.add_paragraph("p) garantir aos seus empregados na execução do contrato, política de respeito à diversidade, prevenção a práticas discriminatórias e de alinhamento com os princípios institucionais de equidade de gênero e raça.")
add_bloco_variante("PB_CessaoDireitos", {
    "Elaboração de projeto, laudo ou anteprojeto por terceiros": [
        "Ceder os direitos patrimoniais relativos ao projeto ou serviço técnico especializado, "
        "para que a Conab possa utilizá-lo, de acordo com o art. 525 do RLC. Quando o projeto "
        "referir-se a obra imaterial de caráter tecnológico, insuscetível de privilégio, a "
        "cessão dos direitos incluirá o fornecimento de todos os dados, documentos e elementos "
        "de informação pertinentes à tecnologia de concepção, desenvolvimento, fixação em "
        "suporte físico e aplicação da obra.",
        "Assegurar à Contratante o direito de propriedade intelectual dos produtos "
        "desenvolvidos, inclusive sobre adequações e atualizações futuras, de forma "
        "permanente, e os direitos autorais da solução, do projeto, de suas especificações "
        "técnicas e da documentação produzida, inclusive os produzidos por terceiros "
        "subcontratados, ficando proibida sua utilização sem autorização expressa da "
        "Contratante, sob pena de multa, sem prejuízo das sanções civis e penais cabíveis.",
        "Realizar a transição contratual com transferência de conhecimento, tecnologia e "
        "técnicas empregadas, sem perda de informações, podendo exigir, inclusive, a "
        "capacitação dos técnicos da Contratante ou da nova empresa que continuará a "
        "execução dos serviços, nos termos do art. 17 do RLC.",
    ],
    "Execução de obra ou serviço já projetado": [""],
})
add_bloco_variante("PB18_ObrigacaoBIM", {
    "Sim": [
        "18.2. Adotar a metodologia BIM na execução dos serviços, nos termos do item 4.4 "
        "deste Projeto Básico, disponibilizando os modelos à Fiscalização no formato IFC ao "
        "final de cada etapa e, obrigatoriamente, no recebimento provisório.",
    ],
    "Não": [""],
})
add_paragraph_with_bookmarks(["18.3. ", ("PB18_ObrigacoesEspecificas", "")])

# 19. DO PAGAMENTO
add_heading("19. DO PAGAMENTO (ART. 157, XVIII, RLC)")
doc.add_paragraph("19.1. A emissão da Nota Fiscal/Fatura será precedida do recebimento provisório ou definitivo dos serviços, conforme o caso.")
add_paragraph_with_bookmarks([
    "19.2. O pagamento será efetuado pela Conab em até ", ("PB19_PrazoPagamentoDiasUteis", "____"),
    " dias úteis, contados do recebimento da Nota Fiscal/Fatura devidamente atestada, "
    "observado o teto legal do RLC.",
])
doc.add_paragraph(
    "19.3. A Nota Fiscal/Fatura conterá, no mínimo: prazo de validade; data de emissão; "
    "dados do Contrato e da Conab; período de prestação; valor a pagar; e tributos a reter "
    "na fonte.")
doc.add_paragraph("19.4. Antes de cada pagamento será realizada consulta ao SICAF. Resultado desfavorável concederá prazo de 5 (cinco) dias úteis para regularização ou defesa.")
doc.add_paragraph("19.5. Dos pagamentos serão retidos os impostos e contribuições devidos na fonte, conforme a legislação vigente.")
doc.add_paragraph(
    "19.6. No atraso de pagamento provocado exclusivamente pela Conab, o valor devido será "
    "acrescido de atualização financeira por meio da fórmula EM = I × N × VP, com I = "
    "0,00016438 (correspondente à taxa de 6% a.a.).")

# 20. DO REAJUSTAMENTO / REPACTUAÇÃO DO CONTRATO
add_heading("20. DO REAJUSTAMENTO / REPACTUAÇÃO DO CONTRATO (ART. 157, XIX, RLC)")
add_bloco_variante("PB20_TipoCorrecao", {
    "Sim": [
        "20.1. Observado o interregno mínimo de 12 (doze) meses, contado da data-base da "
        "proposta, os preços serão reajustados, para os custos decorrentes do mercado, e "
        "repactuados (art. 3º, LXXXIII, e art. 501 do RLC), quando se tratar de contratação "
        "com dedicação exclusiva de mão de obra, por solicitação da Contratada.",
        "20.2. O pedido será fundamentado por novo Acordo/Dissídio/Convenção Coletiva, para "
        "custos de mão de obra, e/ou planilha de custos com variação comprovada, para os "
        "demais custos.",
        "20.3. Serão objeto de preclusão os reajustes ou repactuações a que a Contratada "
        "fizer jus durante a vigência do Contrato e que não forem solicitados até o "
        "implemento dos seguintes eventos: a) assinatura do termo aditivo de prorrogação "
        "contratual; b) data em que o Contrato completa 12 (doze) meses, 24 (vinte e quatro) "
        "meses e assim sucessivamente; ou c) encerramento do Contrato.",
        ["20.4. Para os custos sujeitos à variação de mercado, o índice de reajuste será o ",
         ("PB20_IndiceReajuste", "")],
        "20.5. As repactuações/reajustes serão formalizados por apostilamento, exceto "
        "quando coincidirem com prorrogação contratual, caso em que serão formalizados por "
        "termo aditivo.",
        "20.6. O disposto nesta seção não afasta o direito ao reequilíbrio "
        "econômico-financeiro em sentido estrito, cabível a qualquer tempo, "
        "independentemente de previsão contratual, quando verificado fato imprevisível ou "
        "de consequências incalculáveis, anormal ou extraordinário, nos termos do art. 511 "
        "do RLC.",
    ],
    "Não": [
        "20.1. Não haverá reajuste ou repactuação de preços neste Contrato.",
        "20.2. O disposto nesta seção não afasta o direito ao reequilíbrio "
        "econômico-financeiro em sentido estrito, cabível a qualquer tempo, "
        "independentemente de previsão contratual, quando verificado fato imprevisível ou "
        "de consequências incalculáveis, anormal ou extraordinário, nos termos do art. 511 "
        "do RLC.",
    ],
})

# 21. DO PRAZO DE GARANTIA DA OBRA OU DO SERVIÇO DE ENGENHARIA
add_heading("21. DO PRAZO DE GARANTIA DA OBRA OU DO SERVIÇO DE ENGENHARIA (ART. 157, XX, RLC)")
add_paragraph_with_bookmarks([
    "21.1. Os serviços terão garantia de, no mínimo, ", ("PB21_PrazoGarantiaServico", "(meses ou anos)"),
    ", contados do recebimento definitivo.",
])
add_bloco_variante("PB21_GarantiaObra", {
    "Execução de obra ou serviço já projetado": [
        "21.2. Para obras de engenharia, a garantia mínima contra vícios e defeitos ocultos "
        "é de 5 (cinco) anos (art. 618 do Código Civil), independentemente de estar prevista "
        "neste Projeto Básico ou no Contrato, por decorrer de lei.",
    ],
})
doc.add_paragraph("21.3. Materiais, peças e componentes substituídos serão garantidos pelo prazo do fabricante, quando superior ao mínimo legal.")
doc.add_paragraph("21.4. Durante a garantia, a Contratada refará o serviço, ou parte dele, no prazo estabelecido pela Conab, após comunicada.")

# 22. DA GARANTIA CONTRATUAL
add_heading("22. DA GARANTIA CONTRATUAL (ART. 157, XXI, RLC)")
add_paragraph_with_bookmarks([
    "22.1. A Contratada apresentará, em até ", ("PB22_PrazoGarantiaDias", "____"),
    " dias corridos após a assinatura, garantia de ", ("PB22_PercentualGarantia", "____"),
    "% do valor total do Contrato, em uma das modalidades (à sua escolha): caução em "
    "dinheiro; seguro-garantia; ou fiança bancária.",
])
doc.add_paragraph(
    "22.2. A garantia assegurará o pagamento de: prejuízos do não cumprimento do objeto; "
    "prejuízos diretos à Conab ou a terceiros; multas contratuais; e obrigações trabalhistas "
    "e previdenciárias não adimplidas.")
doc.add_paragraph(
    "22.3. A garantia terá validade de 90 (noventa) dias após o término da vigência. O "
    "atraso na entrega acarretará multa de 0,08% ao dia, até o máximo de 2%; atraso superior "
    "a 25 dias autoriza a rescisão.")
doc.add_paragraph(
    "22.4. O percentual da garantia é, em regra, de até 5% do valor do Contrato (art. 439, "
    "§2º, do RLC), podendo ser elevado para até 10% em contratações de grande vulto envolvendo "
    "complexidade técnica e riscos financeiros elevados (art. 439, §3º, do RLC). Nos casos de "
    "contratação de serviços com dedicação exclusiva de mão de obra, o valor da garantia "
    "deverá corresponder a 5% do valor do contrato, limitado ao equivalente a 2 (dois) meses "
    "do custo da folha de pagamento dos empregados da contratada que venham a participar da "
    "execução dos serviços contratados (art. 439, §4º, do RLC).")

# 23. DA ALTERAÇÃO DO CONTRATO E DA PREVISÃO DE ALTERAÇÃO SUBJETIVA
add_heading("23. DA ALTERAÇÃO DO CONTRATO E DA PREVISÃO DE ALTERAÇÃO SUBJETIVA (ART. 157, XXII, RLC)")
doc.add_paragraph(
    "23.1. O Contrato poderá ser alterado nas hipóteses previstas no art. 510 do RLC. A "
    "Contratada poderá aceitar, nas mesmas condições contratuais, os acréscimos ou "
    "supressões que se fizerem necessários, até 25% (vinte e cinco por cento) do valor "
    "inicial atualizado do Contrato, limite que sobe a 50% (cinquenta por cento) para os "
    "acréscimos, no caso particular de reforma de edifício ou de equipamento.")
doc.add_paragraph(
    "23.2. É admissível a fusão, cisão ou incorporação da Contratada, desde que: observados "
    "todos os requisitos de habilitação da licitação; mantidas as demais cláusulas e "
    "condições; não haja prejuízo à execução; e haja anuência expressa da Conab.")

# 24. DA FISCALIZAÇÃO E GERENCIAMENTO DO CONTRATO
add_heading("24. DA FISCALIZAÇÃO E GERENCIAMENTO DO CONTRATO (ART. 157, XXIII, RLC)")
doc.add_paragraph(
    "24.1. Adotar-se-á o método baseado em delegação de responsabilidade, com os seguintes "
    "papéis: Fiscal Técnico, que acompanha e fiscaliza tecnicamente a execução, verifica os "
    "resultados e realiza o recebimento provisório; Fiscal Administrativo/Funcional, que "
    "verifica as obrigações funcionais e trabalhistas, afere a conformidade documental e "
    "emite o IMR; e Preposto da Contratada, representante responsável por acompanhar a "
    "execução e atuar como interlocutor com a Conab.")
doc.add_paragraph("24.2. A gestão e a fiscalização observarão os artigos do RLC sobre gestão e fiscalização de contratos de obras e serviços de engenharia.")
doc.add_paragraph("24.3. O Contrato só será considerado integralmente cumprido após a comprovação do pagamento de todas as obrigações trabalhistas, sociais, previdenciárias e de FGTS da mão de obra alocada.")

# 25. DAS SANÇÕES ADMINISTRATIVAS
add_heading("25. DAS SANÇÕES ADMINISTRATIVAS (ART. 157, XXIV, RLC)")
doc.add_paragraph("25.1. Garantidos o contraditório e a ampla defesa, o inadimplemento sujeitará a Contratada às sanções previstas no RLC e na Lei nº 13.303/2016:")
add_paragraph_with_bookmarks(["a) Advertência, para infrações de menor gravidade, sem prejuízo à execução do objeto;"])
add_paragraph_with_bookmarks([
    "b) multa por não regularização da documentação de habilitação: ",
    ("PB25_MultaNaoRegularizacaoHabilitacao", "____"), "% sobre o valor estimado para a contratação;",
])
add_paragraph_with_bookmarks([
    "c) multa pela prática das condutas elencadas nos arts. 576 a 580 do RLC: ",
    ("PB25_MultaCondutasArt576a580", "____"), "% sobre o valor estimado para a contratação;",
])
add_paragraph_with_bookmarks([
    "d) multa pela recusa em assinar o Contrato dentro do prazo estabelecido: ",
    ("PB25_MultaRecusaAssinarContrato", "____"), "% sobre o valor da contratação;",
])
doc.add_paragraph(
    "e) multa moratória de 0,2% (dois décimos por cento) sobre o valor do Contrato (ou "
    "valor anual do Contrato), por dia de atraso na execução, até o limite de 15 (quinze) "
    "dias;")
doc.add_paragraph(
    "f) multa moratória de 0,3% (três décimos por cento) sobre o valor do Contrato (ou "
    "valor anual do Contrato), por dia de atraso na execução, por período superior ao "
    "previsto na alínea anterior, até o limite de 15 (quinze) dias. Esgotado o prazo, "
    "poderá ocorrer a não aceitação do objeto, sem prejuízo da rescisão unilateral da "
    "avença;")
add_paragraph_with_bookmarks([
    "g) Multa compensatória de ", ("PB25_MultaCompensatoria", "____"),
    "% do valor total do Contrato, pelo descumprimento grave ou rescisão por culpa da Contratada;",
])
add_paragraph_with_bookmarks([
    "h) Multa rescisória de ", ("PB25_MultaRescisoria", "____"),
    "% nos casos de rescisão unilateral por descumprimento contratual;",
])
doc.add_paragraph("i) Suspensão do direito de licitar e impedimento de contratar com a Conab pelo prazo de até 2 (dois) anos.")
doc.add_paragraph("25.2. As multas são autônomas, podendo ser aplicadas cumulativa ou isoladamente, e serão descontadas da garantia ou dos pagamentos; o saldo excedente será cobrado judicialmente. A aplicação das sanções será registrada no SICAF.")

# 26. DA MATRIZ DE RISCOS E QUADRO DE INOVAÇÕES
add_heading("26. DA MATRIZ DE RISCOS E QUADRO DE INOVAÇÕES (ART. 157, XXV, RLC)")
doc.add_paragraph(
    "26.1. A Matriz de Riscos é a cláusula definidora dos riscos e das responsabilidades "
    "entre Contratante e Contratada e caracterizadora do equilíbrio econômico-financeiro do "
    "Contrato, constando do Anexo II deste Projeto Básico.")
doc.add_paragraph("26.2. A Contratada só poderá solicitar reequilíbrio econômico-financeiro nas hipóteses alocadas à Contratante na Matriz de Riscos; fica vedada a celebração de aditivos por eventos alocados à Contratada.")
add_bloco_variante("PB26_ExigeQuadroInovacoes", {
    "Sim": ["26.3. Em obras e serviços de engenharia, além da Matriz de Riscos, é necessário "
            "o preenchimento do Quadro de Inovações (art. 157, XXV, c/c art. 210, III, do "
            "RLC), constante do Anexo II-A."],
    "Não": [""],
})

# 27. DA ORÇAMENTAÇÃO
add_heading("27. DA ORÇAMENTAÇÃO (ART. 157, XXVI, RLC)")
doc.add_paragraph(
    "27.1. As planilhas orçamentárias da contratação, elaboradas conforme o RLC, constam do "
    "Anexo de Memória de Cálculo e do Anexo de Orçamento Analítico (documentos técnicos "
    "produzidos no sistema oficial de orçamentação e anexados ao processo separadamente), "
    "com indicação da fonte de cada item (SINAPI como referência primária; SICRO para obras "
    "rodoviárias; demais fontes apenas mediante justificativa de inviabilidade).")

# 28. DO CRITÉRIO DE AVALIAÇÃO DO CUSTO
add_heading("28. DO CRITÉRIO DE AVALIAÇÃO DO CUSTO (ART. 157, XXVII, RLC)")

_bid_28_execucao = add_bloco_variante_sentinela_inicio(
    f"PB28_MetodologiaCusto_BLOCO_{normalizar_variante('Execução de obra ou serviço já projetado')}")
add_paragraph_with_bookmarks([
    "28.1. A estimativa adotou o critério de ", ("PB28_CriterioEstimativa", "(média, mediana ou menor preço)"),
    " dos preços, obtidos por ", ("PB28_FonteReferencia", "(fonte)"), ".",
])
add_paragraph_with_bookmarks([
    "28.2. Os custos diretos sem BDI totalizam R$ ", ("PB28_CustoSemBDI", "____"),
    ". O BDI adotado é de ", ("PB28_PercentualBDI", "____"),
    "%, resultando no valor estimado total de R$ ", ("PB28_ValorTotalComBDI", "____"), ".",
])
add_bloco_variante_sentinela_fim(_bid_28_execucao)

_bid_28_elaboracao = add_bloco_variante_sentinela_inicio(
    f"PB28_MetodologiaCusto_BLOCO_{normalizar_variante('Elaboração de projeto, laudo ou anteprojeto por terceiros')}")
doc.add_paragraph(
    "28.1. Por se tratar de serviço técnico de natureza intelectual, a orçamentação não "
    "segue a equação de BDI do art. 165 do RLC, adotando metodologia de Fator K e TRDE, "
    "própria da engenharia consultiva.")
add_paragraph_with_bookmarks(["28.2. ", ("PB28_MetodologiaCustoIntelectual", "")])
add_bloco_variante_sentinela_fim(_bid_28_elaboracao)

# 29. DO CRONOGRAMA FÍSICO-FINANCEIRO
add_heading("29. DO CRONOGRAMA FÍSICO-FINANCEIRO (ART. 157, XXVIII, RLC)")
doc.add_paragraph(
    "29.1. O Cronograma Físico-Financeiro detalhado, com etapas, percentuais físicos e "
    "desembolsos mensais, é documento técnico produzido separadamente e anexado ao processo. "
    "Quando, em razão da natureza do objeto, não for possível estabelecê-lo, a omissão e sua "
    "justificativa constam expressamente nos autos, nos termos do art. 157, §1º, do RLC.")

# 30. DO VALOR ESTIMADO DA CONTRATAÇÃO
add_heading("30. DO VALOR ESTIMADO DA CONTRATAÇÃO (ART. 157, XXIX, RLC)")
add_paragraph_with_bookmarks([
    "30.1. O custo total estimado é de R$ ", ("PB30_ValorEstimado", "(valor)"),
    " (", ("PB30_ValorPorExtenso", "valor por extenso"), ").",
])
doc.add_paragraph("30.2. Os recursos correrão à conta do Orçamento da União para a Conab, conforme dados orçamentários fornecidos pela área financeira nos autos do processo.")

# 31. DA SUBCONTRATAÇÃO
add_heading("31. DA SUBCONTRATAÇÃO (ART. 157, XXX, RLC)")
add_bloco_variante("PB31_PermiteSubcontratacao", {
    "Sim": ["31.1. É permitida a subcontratação parcial do objeto, vedada, em qualquer caso, "
            "a subcontratação total (art. 524 do RLC)."],
    "Não": ["31.1. Não será admitida a subcontratação do objeto, vedada, em qualquer caso, "
            "a subcontratação total (art. 524 do RLC)."],
})
add_paragraph_with_bookmarks(["Percentual máximo de subcontratação: ", ("PB31_PercentualSubcontratacao", "____"), "%"])
doc.add_paragraph(
    "31.2. A subcontratação não transfere responsabilidade: a Contratada permanece "
    "integralmente responsável pela perfeita execução, padronização, compatibilidade, "
    "gerenciamento centralizado e qualidade. A subcontratada deverá comprovar habilitação "
    "jurídica, regularidade fiscal e qualificação técnica da parcela. É vedada a "
    "subcontratação de quem tenha participado da licitação ou da elaboração deste Projeto "
    "Básico.")

# 32. DA RESCISÃO CONTRATUAL
add_heading("32. DA RESCISÃO CONTRATUAL")
doc.add_paragraph(
    "32.1. A inexecução total do Contrato ensejará a sua rescisão, enquanto a inexecução "
    "parcial poderá ensejar a sua rescisão, com as consequências cabíveis, conforme disposto "
    "nos arts. 568 a 572 do RLC, incluindo as hipóteses de não apresentação ou recomposição "
    "da garantia, falência ou insolvência, perda das condições de habilitação e razões de "
    "interesse público mediante despacho motivado.")
doc.add_paragraph(
    "32.2. A rescisão poderá ser: a) por ato unilateral e escrito da Conab; b) amigável, por "
    "acordo entre as partes, reduzida a termo no processo de contratação, desde que haja "
    "conveniência para a Conab; e c) judicial, por determinação judicial.")
doc.add_paragraph(
    "32.2.1. A rescisão amigável deverá ser precedida de autorização escrita e fundamentada "
    "da autoridade competente.")
doc.add_paragraph(
    "32.2.2. A rescisão amigável não será cabível nos casos em que forem constatados "
    "descumprimentos contratuais sem apuração de responsabilidade iniciada ou com apuração "
    "ainda em curso.")
doc.add_paragraph(
    "32.2.3. Os casos de rescisão contratual serão formalmente motivados, assegurando-se à "
    "Contratada o direito à prévia e ampla defesa, conforme procedimento previsto nos arts. "
    "582 a 593 do RLC.")
doc.add_paragraph(
    "32.3. A rescisão por ato unilateral da Conab acarreta as seguintes consequências, sem "
    "prejuízo das sanções previstas neste Projeto Básico: a) assunção imediata do objeto "
    "contratado, pela Conab, no estado e local em que se encontrar; b) execução da garantia "
    "contratual, quando houver, para ressarcimento pelos eventuais prejuízos sofridos pela "
    "Conab; e c) na hipótese de insuficiência da garantia contratual, quando houver, a "
    "retenção dos créditos decorrentes do Contrato até o limite dos prejuízos causados à "
    "Conab.")
doc.add_paragraph(
    "32.4. A rescisão deverá ser formalizada por termo de rescisão unilateral ou distrato, no "
    "caso de rescisão amigável, devendo o respectivo extrato ser publicado no Diário Oficial "
    "da União.")
doc.add_paragraph(
    "32.5. O termo de rescisão, sempre que possível, será precedido de: a) balanço dos "
    "eventos contratuais já cumpridos ou parcialmente cumpridos; b) relação dos pagamentos já "
    "efetuados e ainda devidos; e c) indenizações e multas.")

# 33. DA VEDAÇÃO AO NEPOTISMO
add_heading("33. DA VEDAÇÃO AO NEPOTISMO")
doc.add_paragraph(
    "33.1. São vedadas as contratações que configurem nepotismo, notadamente aquelas que "
    "envolvam relação de parentesco, até o terceiro grau civil, com dirigente da Conab, com "
    "empregado cujas atribuições envolvam a atuação na área responsável pela "
    "licitação/contratação ou com autoridade do Ministério do Desenvolvimento Agrário e "
    "Agricultura Familiar (MDA), nos termos do art. 12, §1º, inciso II, do RLC e do art. 7º "
    "do Decreto nº 7.203/2010.")

# 34. DAS VEDAÇÕES (texto fixo)
add_heading("34. DAS VEDAÇÕES")
doc.add_paragraph(
    "34.1. É vedado à Contratada: a) caucionar ou utilizar o Contrato para qualquer operação "
    "financeira; b) interromper a execução contratual sob alegação de inadimplemento por "
    "parte da Contratante, salvo nos casos previstos em lei; c) empregar menor de 18 "
    "(dezoito) anos em trabalho noturno, perigoso ou insalubre, e menor de 16 (dezesseis) "
    "anos, salvo na condição de aprendiz a partir de 14 (quatorze) anos; d) subcontratar, "
    "total ou parcialmente, o objeto da contratação, ressalvadas as hipóteses de "
    "subcontratação parcial admitidas neste Projeto Básico; e e) solicitar, por meio de seus "
    "empregados, serviços, materiais ou equipamentos a empresas terceirizadas que prestem "
    "serviços à Conab (art. 155 do RLC).")

# 35. DAS DISPOSIÇÕES FINAIS E RELAÇÃO DE ANEXOS
add_heading("35. DAS DISPOSIÇÕES FINAIS E RELAÇÃO DE ANEXOS")
doc.add_paragraph(
    "35.1. Este Projeto Básico foi elaborado com base no RLC 10.901 e nas normas técnicas "
    "aplicáveis, constituindo o conjunto de elementos necessários e suficientes à realização "
    "do certame e à contratação do objeto.")
add_bloco_variante("PB34_RevogacaoOuAnulacao", {
    "Licitação": [
        "35.2. A Conab poderá revogar a licitação, por interesse público, ou anulá-la por "
        "ilegalidade, mediante despacho fundamentado.",
    ],
    "Dispensa de Licitação": [
        "35.2. A Conab poderá anular o processo de dispensa por ilegalidade, mediante "
        "despacho fundamentado, assegurados o contraditório e a ampla defesa.",
    ],
})
doc.add_paragraph(
    "35.3. Integram este Projeto Básico os seguintes Anexos: a) Anexo I — Caracterização do "
    "Objeto e Condições Gerais; b) Anexo II — Matriz de Riscos (e Anexo II-A — Quadro de "
    "Inovações, quando aplicável); c) Anexo III — Modelo de Apresentação de Proposta; d) "
    "Anexo IV — Modelo de Declaração de Vistoria; e) Anexo V — Modelo de Declaração de "
    "Dispensa de Vistoria; f) Anexo VI — Modelo de Declaração de Contratos Firmados; g) "
    "Anexo VII — Modelo de Instrumento de Medição de Resultado (IMR); h) Anexo VIII — "
    "Modelo de Placa de Obra/Serviço de Engenharia. As Especificações Técnicas, a Memória de "
    "Cálculo, a Curva ABC, o Orçamento Analítico e o Cronograma Físico-Financeiro são "
    "documentos técnicos produzidos no sistema oficial de orçamentação e anexados ao "
    "processo SEI separadamente.")
doc.add_paragraph("Aprovo o presente Projeto Básico e seus Anexos em todo o seu teor, tendo em vista a coerência das justificativas e dos objetivos em relação à contratação (arts. 100, V, e 131 do RLC).")
doc.add_paragraph("Elaborado por:")
add_paragraph_with_bookmarks(["1) ", ("PB34_Elab1NomeMatricula", ""), " — CREA/CAU: ", ("PB34_Elab1CreaCau", ""), " — ", ("PB34_Elab1Area", "")])
add_paragraph_with_bookmarks(["2) ", ("PB34_Elab2NomeMatricula", ""), " — CREA/CAU: ", ("PB34_Elab2CreaCau", ""), " — ", ("PB34_Elab2Area", "")])
add_paragraph_with_bookmarks(["3) ", ("PB34_Elab3NomeMatricula", ""), " — CREA/CAU: ", ("PB34_Elab3CreaCau", ""), " — ", ("PB34_Elab3Area", "")])
doc.add_paragraph("Aprovado por:")
add_paragraph_with_bookmarks(["1) ", ("PB34_Aprov1NomeMatricula", ""), " — ", ("PB34_Aprov1Area", "")])
add_paragraph_with_bookmarks(["2) ", ("PB34_Aprov2NomeMatricula", ""), " — ", ("PB34_Aprov2Area", "")])
add_paragraph_with_bookmarks(["3) ", ("PB34_Aprov3NomeMatricula", ""), " — ", ("PB34_Aprov3Area", "")])

# ANEXO I - CARACTERIZAÇÃO DO OBJETO E CONDIÇÕES GERAIS
doc.add_page_break()
add_heading("ANEXO I DO PROJETO BÁSICO")
add_heading("CARACTERIZAÇÃO DO OBJETO E CONDIÇÕES GERAIS")
add_paragraph_with_bookmarks(["Objeto: ", ("PB01_DescricaoObjeto_AnexoI", "")])
doc.add_paragraph(
    "1. Objeto e localização: denominação completa do objeto e número do processo SEI; tipo "
    "de intervenção (obra nova, reforma, ampliação, manutenção, adequação etc.); regime de "
    "execução adotado; e, para cada local, nome da unidade Conab, endereço completo e CEP.")
doc.add_paragraph(
    "2. Diagnóstico da situação atual e solução proposta: situação atual das "
    "instalações/sistemas, com identificação das deficiências que motivam a contratação; "
    "solução técnica adotada; resultado esperado.")
doc.add_paragraph(
    "3. Escopo dos serviços: relação de todos os sistemas/especialidades abrangidos "
    "(elétrico, hidrossanitário, combate a incêndio, climatização, instalações civis, "
    "marcenaria, lógica/dados, CFTV, gás, infraestrutura externa etc.), com a natureza da "
    "intervenção e o local de cada um.")
doc.add_paragraph(
    "4. Limitações, exclusões e informações complementares: serviços/sistemas excluídos do "
    "escopo; interfaces com outros contratos; restrições operacionais; área construída, "
    "número de pavimentos, ano de construção/última reforma, laudos e projetos existentes.")
doc.add_paragraph(
    "5. Mobilização, canteiro e dispositivos iniciais: layout do canteiro para aprovação "
    "prévia da Fiscalização; delimitação do canteiro; Diário de Obra; registro fotográfico.")
doc.add_paragraph(
    "6. Segurança e saúde no trabalho: cumprimento das Normas Regulamentadoras do MTE "
    "aplicáveis (NR-4, NR-6, NR-18, NR-24, NBR 7678/1983); PGR e PCMSO; lista de EPIs "
    "obrigatórios; vedação de permanência sem EPI e uniforme.")
doc.add_paragraph(
    "7. Condições operacionais de execução: horário ordinário de execução; condições para "
    "trabalho extraordinário ou em dias não úteis; manutenção do canteiro em condições de "
    "limpeza e organização.")

# ANEXO II - MATRIZ DE RISCOS (Anexo II do RLC, adaptado pelo modelo
# institucional GEPEA com a coluna Fase — transcrito integralmente)
doc.add_page_break()
add_heading("ANEXO II DO PROJETO BÁSICO")
add_heading("MATRIZ DE RISCOS")
add_paragraph_with_bookmarks(["Objeto: ", ("PB01_DescricaoObjeto_AnexoII", "")])
tabela_riscos = add_tabela(
    ["Item", "Fase", "Evento de Risco", "Causas", "Consequências",
     "Probabilidade (P)", "Impacto (I)", "Nível de Risco (P x I)",
     "Resposta ao Risco", "Responsável"],
)
for i, fase in enumerate(["Planejamento da Contratação", "Seleção de Fornecedores", "Gestão do Contrato"], start=1):
    linha = tabela_riscos.add_row()
    linha.cells[0].text = str(i)
    linha.cells[1].text = fase
doc.add_paragraph(
    "Fases: planejamento da contratação, seleção de fornecedores e gestão do Contrato. "
    "Evento de risco: evento incerto que, se ocorrer, afeta os objetivos da contratação. "
    "Causas: condições que viabilizam o evento. Consequências: consequências da "
    "ocorrência. Resposta ao risco: estratégia adotada (mitigar o impacto, evitar a "
    "causa, transferir via seguro/garantia ou aceitar o risco), detalhando a ação "
    "prática preventiva ou de contingência. Responsável: Contratante ou Contratada.")
doc.add_paragraph(
    "A avaliação da probabilidade deverá ser analisada em uma escala de 1 a 5, "
    "conforme a tabela abaixo:")
add_tabela(
    ["Descritor", "Descrição", "Nível"],
    linhas=[
        ("Muito Baixa", "Evento extraordinário, sem histórico de ocorrência.", "1"),
        ("Baixa", "Evento casual e inesperado; raro, mas com histórico de ocorrência.", "2"),
        ("Média", "Evento esperado, de frequência reduzida, com histórico parcialmente conhecido.", "3"),
        ("Alta", "Evento usual, com histórico amplamente conhecido.", "4"),
        ("Muito Alta", "Evento repetitivo e constante.", "5"),
    ],
)
doc.add_paragraph("A avaliação do impacto deverá ser analisada em uma escala de 1 a 5, conforme a tabela abaixo:")
add_tabela(
    ["Descritor", "Descrição", "Nível"],
    linhas=[
        ("Muito Baixo", "Impacto insignificante nos objetivos.", "1"),
        ("Baixo", "Impacto mínimo nos objetivos.", "2"),
        ("Médio", "Impacto mediano, com possibilidade de recuperação.", "3"),
        ("Alto", "Impacto significante, com possibilidade remota de recuperação.", "4"),
        ("Muito Alto", "Impacto máximo, sem possibilidade de recuperação.", "5"),
    ],
)
doc.add_paragraph("Nível de risco (P x I):")
add_tabela(
    ["Faixa (P x I)", "Classificação"],
    linhas=[
        ("1 a 2", "Baixo"),
        ("3 a 6", "Médio"),
        ("8 a 12", "Elevado"),
        ("15 a 25", "Extremo"),
    ],
)
doc.add_paragraph("Observação: as hipóteses são exemplificativas e não exaustivas.")

# ANEXO II-A - QUADRO DE INOVAÇÕES (só existe se PB26_ExigeQuadroInovacoes = "Sim")
# Anexo III do RLC — Minuta de Matriz de Riscos das Contratações de Obras e
# Serviços de Engenharia, transcrito integralmente (checklist de 14 frações
# de obra) a partir de
# anexos RLC/anexo_iii_minuta_matriz_riscos_contratacoes_obras_servicos_engenhariaodt.odt
_bid_anexo_iia = add_bloco_variante_sentinela_inicio("PB26_ExigeQuadroInovacoes_ANEXO_IIA_BLOCO_SIM")
doc.add_page_break()
add_heading("ANEXO II-A DO PROJETO BÁSICO")
add_heading("QUADRO DE INOVAÇÕES")
doc.add_paragraph(
    "Em caso de contratação de obras e serviços, além do Anexo II — Matriz de "
    "Riscos, é necessário o preenchimento da tabela abaixo (art. 157, XXV, c/c "
    "art. 210, III, do RLC).")
add_tabela(
    ["Frações do objeto (Obras e serviços de engenharia)",
     "Permissão para inovar em soluções metodológicas e tecnológicas, desde que "
     "atento ao que foi delineado no projeto básico. (Sim/Não)"],
    linhas=[
        ("Serviços preliminares", ""),
        ("Terraplanagem", ""),
        ("Infraestrutura (fundação e laje)", ""),
        ("Supraestrutura (pilares, vigas e vedações)", ""),
        ("Cobertura", ""),
        ("Piso e pavimentação", ""),
        ("Instalações", ""),
        ("Equipamentos", ""),
        ("Esquadrias", ""),
        ("Revestimentos", ""),
        ("Sinalização", ""),
        ("Obras complementares", ""),
        ("Interferências", ""),
        ("Meio ambiente e paisagismo", ""),
    ],
    linhas_vazias=6,
)
doc.add_paragraph("Este Rol não é exaustivo.")
add_tabela(["Elaborado por: _________________________________________________ Nome e Assinatura"])
add_bloco_variante_sentinela_fim(_bid_anexo_iia)

# ANEXO III - MODELO DE APRESENTAÇÃO DE PROPOSTA
doc.add_page_break()
add_heading("ANEXO III DO PROJETO BÁSICO")
add_heading("MODELO DE APRESENTAÇÃO DE PROPOSTA")
doc.add_paragraph("Identificação do licitante: razão social/CNPJ, endereço, telefone/e-mail, representante legal.")
doc.add_paragraph("1. Objeto: reproduzir, com fidelidade, a descrição do objeto conforme o item 1.1 do Projeto Básico.")
tabela_proposta = doc.add_table(rows=1, cols=6)
tabela_proposta.style = "Table Grid"
for celula, texto in zip(tabela_proposta.rows[0].cells, ["Item", "Código/Banco", "Descrição", "Und.", "Quant.", "Vlr. unit. (c/ BDI)"]):
    celula.text = texto
for _ in range(4):
    tabela_proposta.add_row()
doc.add_paragraph("Total sem BDI (R$): ____     Total do BDI (R$): ____     Total Geral (R$): ____     Total Geral por extenso: ____")
doc.add_paragraph(
    "Declaramos que os preços propostos foram calculados com estrita observância das regras "
    "de BDI vigentes, incluindo todos os tributos diretos e indiretos, encargos "
    "trabalhistas, previdenciários, fiscais e comerciais, bem como fretes e seguros "
    "necessários para a perfeita execução do objeto, mantendo-se firmes pelo prazo de "
    "validade indicado abaixo.")
doc.add_paragraph(
    "Prazo de execução e condições de execução: ____     Prazo de vigência do Contrato: "
    "____     Prazo de garantia oferecido (mínimo de 60 dias, quando aplicável): ____")
doc.add_paragraph("Validade da proposta: ____ dias da abertura do certame (mínimo de 60 dias).")

# ANEXO IV - MODELO DE DECLARAÇÃO DE VISTORIA TÉCNICA
doc.add_page_break()
add_heading("ANEXO IV DO PROJETO BÁSICO")
add_heading("MODELO DE DECLARAÇÃO DE VISTORIA TÉCNICA")
add_paragraph_with_bookmarks(["Objeto: ", ("PB01_DescricaoObjeto_AnexoIV", "")])
doc.add_paragraph("À Companhia Nacional de Abastecimento — Conab.")
doc.add_paragraph(
    "Declaramos que, na data indicada, o(a) representante da empresa compareceu perante a "
    "Conab/Unidade/Sureg, no endereço indicado, e vistoriou os locais de execução dos "
    "serviços objeto da contratação, tomando plena ciência das condições e graus de "
    "dificuldade existentes.")
doc.add_paragraph(
    "Declaramos, ainda, que a realização da vistoria técnica e o pleno conhecimento das "
    "condições locais impedem a alegação futura de desconhecimento das condições de "
    "execução, não podendo servir de argumento para pleitos de reequilíbrio "
    "econômico-financeiro ou para justificar o descumprimento de obrigações contratuais "
    "junto à Conab.")
doc.add_paragraph("Local e data: ____     Assinatura/carimbo do representante da empresa: ____     Visto do responsável pela Conab: ____")

# ANEXO V - MODELO DE DECLARAÇÃO DE DISPENSA DE VISTORIA TÉCNICA
doc.add_page_break()
add_heading("ANEXO V DO PROJETO BÁSICO")
add_heading("MODELO DE DECLARAÇÃO DE DISPENSA DE VISTORIA TÉCNICA")
add_paragraph_with_bookmarks(["Objeto: ", ("PB01_DescricaoObjeto_AnexoV", "")])
add_bloco_variante("PB09_TipoVistoria_AnexoV", {
    "Obrigatória": [
        "Não aplicável a esta contratação: a vistoria técnica é obrigatória (Seção 9 deste "
        "Projeto Básico), não havendo hipótese de dispensa a ser declarada pelo licitante.",
    ],
    "Facultativa": [
        "À Companhia Nacional de Abastecimento — Conab.",
        "A empresa declara que optou por não realizar a vistoria aos locais de execução, que "
        "conhece as condições locais e o grau de dificuldade para a perfeita execução do objeto, "
        "que assume todo e qualquer risco por essa decisão, conforme o item 9 do Projeto Básico, "
        "e que se compromete a prestar fielmente os serviços nos termos do Edital, do Projeto "
        "Básico e dos demais anexos.",
        "Declaramos, ainda, que a não realização da vistoria técnica não poderá ser utilizada "
        "como argumento para futuros pleitos de reequilíbrio econômico-financeiro ou para "
        "justificar o descumprimento de obrigações contratuais junto à Conab.",
        "Local e data: ____     Assinatura/carimbo do representante da empresa: ____",
    ],
})

# ANEXO VI - MODELO DE DECLARAÇÃO DE CONTRATOS FIRMADOS (ANEXO IV, RLC) —
# transcrito integralmente a partir do modelo institucional GEPEA
doc.add_page_break()
add_heading("ANEXO VI DO PROJETO BÁSICO")
add_heading("MODELO DE DECLARAÇÃO DE CONTRATOS FIRMADOS")
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

# ANEXO VII - MODELO DE INSTRUMENTO DE MEDIÇÃO DE RESULTADO (IMR)
# Estrutura transcrita integralmente do anexo oficial do RLC
# (anexos RLC/anexo_xv_instrumento_de_medicao_resultadoodt.odt): uma
# tabela-modelo em branco (Indicador Nº + Título, e os itens Item a
# Observações) seguida de um exemplo real preenchido. Cada contratação
# deve definir seus próprios indicadores fora do site (Word ou SEI) —
# não inventamos critérios específicos por objeto.
doc.add_page_break()
add_heading("ANEXO VII DO PROJETO BÁSICO")
add_heading("MODELO DE INSTRUMENTO DE MEDIÇÃO DE RESULTADO (IMR)")

ITENS_IMR = [
    "Finalidade", "Meta a cumprir", "Instrumento de medição",
    "Forma de acompanhamento", "Periodicidade", "Mecanismo de cálculo",
    "Início de vigência", "Faixas de ajuste no pagamento", "Sanções",
    "Observações",
]

doc.add_paragraph("Indicador Nº ____ – ____ (título do indicador)")
tabela_imr_modelo = doc.add_table(rows=1, cols=2)
tabela_imr_modelo.style = "Table Grid"
tabela_imr_modelo.rows[0].cells[0].text = "Item"
tabela_imr_modelo.rows[0].cells[1].text = "Descrição"
for item in ITENS_IMR:
    linha = tabela_imr_modelo.add_row()
    linha.cells[0].text = item
    linha.cells[1].text = ""

doc.add_paragraph("Exemplo de indicador")
doc.add_paragraph("Nº 01 – Prazo de atendimento de demandas (OS)")
tabela_imr_exemplo = doc.add_table(rows=1, cols=2)
tabela_imr_exemplo.style = "Table Grid"
tabela_imr_exemplo.rows[0].cells[0].text = "Item"
tabela_imr_exemplo.rows[0].cells[1].text = "Descrição"
for item, descricao in [
    ("Finalidade", "Garantir um atendimento célere às demandas da Conab."),
    ("Meta a cumprir", "24 h."),
    ("Instrumento de medição", "Sistema informatizado de solicitação de serviços – Ordem de Serviços (OS) eletrônica."),
    ("Forma de acompanhamento", "Pelo sistema."),
    ("Periodicidade", "Mensal."),
    ("Mecanismo de cálculo", "Cada OS será verificada e valorada individualmente. Nº de horas no atendimento / 24 = X"),
    ("Início de vigência", "Data da assinatura do contrato."),
    ("Faixas de ajuste no pagamento",
     "X até 1 – 100% do valor da OS\nDe 1 a 1,5 – 90% do valor da OS\nDe 1,5 a 2 – 80% do valor da OS"),
    ("Sanções",
     "20% das OS acima de 2 – multa de XX\n30% das OS acima de 2 – multa de XX + rescisão contratual"),
    ("Observações", "RLC"),
]:
    linha = tabela_imr_exemplo.add_row()
    linha.cells[0].text = item
    linha.cells[1].text = descricao

# ANEXO VIII - MODELO DE PLACA DE OBRA/SERVIÇO DE ENGENHARIA (só existe para
# execução de obra/serviço — ver PB01_NaturezaObjeto. O exemplar dos
# Hortomercados exclui este anexo, com justificativa expressa, quando a
# contratação não envolve execução física visível ao público.)
doc.add_page_break()
_bid_anexo_viii = add_bloco_variante_sentinela_inicio(
    f"PB_AnexoVIII_BLOCO_{normalizar_variante('Execução de obra ou serviço já projetado')}")
add_heading("ANEXO VIII DO PROJETO BÁSICO")
add_heading("MODELO DE PLACA DE OBRA/SERVIÇO DE ENGENHARIA")
doc.add_paragraph(
    "A placa de obra/serviço de engenharia deverá seguir o modelo institucional vigente à "
    "época da contratação, observando as dimensões mínimas, materiais, identidade visual, "
    "logomarcas, informações do contrato, identificação do responsável técnico e número da "
    "ART/RRT, conforme os padrões da Conab, a ABNT NBR 12284 e demais normas aplicáveis "
    "(inclusive eventuais restrições da legislação eleitoral vigentes à época).")
add_bloco_variante_sentinela_fim(_bid_anexo_viii)

# ANEXO — MINUTA DE CONTRATO (só existe se Dispensa de Licitação; art.
# 438 do RLC — 30 partes essenciais da minuta de Contrato — e art. 435,
# II, que impede o Contrato Simplificado para obras/serviços de
# engenharia). Sem minuta SEI de referência para esta variante — texto
# derivado diretamente do RLC, reaproveitando a redação já validada do
# Anexo VI do TR-Dispensa para as cláusulas RLC-genéricas. Ver
# docs/superpowers/specs/2026-08-18-pb-dispensa-design.md.
_bid_anexo_contrato = add_bloco_variante_sentinela_inicio("PB_AnexoContrato_BLOCO_DISPENSA_DE_LICITACAO")
doc.add_page_break()
add_heading("ANEXO DO PROJETO BÁSICO")
add_heading("MINUTA DE CONTRATO")
add_paragraph_with_bookmarks([
    "Processo Administrativo nº ", ("PB_AnexoContrato_NumeroProcesso", "____"),
    ". Contrato nº ", ("PB_AnexoContrato_NumeroContrato", "____"), ".",
])
doc.add_paragraph(
    "CONTRATANTE: (unidade da Conab — Matriz ou Sureg —, CNPJ, endereço e representante "
    "legal com nome, cargo e CPF)")
doc.add_paragraph(
    "CONTRATADA: (razão social), (CNPJ, endereço e representante legal com nome e CPF)")
add_paragraph_with_bookmarks([
    "CONTRATO QUE ENTRE SI CELEBRAM A COMPANHIA NACIONAL DE ABASTECIMENTO - CONAB E A EMPRESA "
    "___________________ PARA A CONTRATAÇÃO DE ", ("PB01_TipoObjeto_AnexoContrato", "(obra/serviço)"),
    ", de conformidade com o Processo Administrativo nº ",
    ("PB_AnexoContrato_NumeroProcesso_Inline", "____"),
    " referente à Dispensa de Licitação, "
    "regendo-se pelo Projeto Básico e seus anexos e pela proposta da Contratada, pelo RLC, "
    "pela Lei nº 13.303, de 2016, e demais legislações pertinentes.",
])

doc.add_paragraph("CLÁUSULA PRIMEIRA - DO OBJETO")
add_paragraph_with_bookmarks([
    "O objeto do presente instrumento é a contratação de ",
    ("PB01_DescricaoObjeto_AnexoContrato", ""),
    ", nas condições estabelecidas no Projeto Básico, ao qual este Termo de Contrato se "
    "vincula, independentemente de transcrição, juntamente com a proposta vencedora.",
])

doc.add_paragraph("CLÁUSULA SEGUNDA - DA VIGÊNCIA")
add_paragraph_with_bookmarks([
    "O prazo de vigência do Contrato será de ", ("PB08_PrazoVigencia_AnexoContrato", "____"),
    ", contado da data de sua assinatura, admitida prorrogação nos termos do RLC.",
])

doc.add_paragraph("CLÁUSULA TERCEIRA - DO REGIME DE EXECUÇÃO, DOS PRAZOS E DO RECEBIMENTO")
add_paragraph_with_bookmarks([
    "Os serviços/obras serão executados no regime de ", ("PB01_RegimeExecucao_AnexoContrato", ""),
    ", com recebimento provisório em até ", ("PB16_PrazoRecebimentoProvisorioDias_AnexoContrato", "____"),
    " dias corridos do adimplemento da parcela/entrega, e recebimento definitivo em até ",
    ("PB16_PrazoRecebimentoDefinitivoDias_AnexoContrato", "____"),
    " dias corridos do recebimento provisório, observado o Instrumento de Medição de "
    "Resultado (IMR), Anexo VII do Projeto Básico, quando for o caso.",
])

doc.add_paragraph("CLÁUSULA QUARTA - DO VALOR DO CONTRATO")
doc.add_paragraph(
    "O valor do Contrato corresponde ao valor proposto pela Contratada vencedora, incluídas "
    "todas as despesas ordinárias diretas e indiretas decorrentes da execução do objeto, "
    "inclusive tributos, encargos sociais, trabalhistas e previdenciários.")

doc.add_paragraph("CLÁUSULA QUINTA - DA GARANTIA CONTRATUAL")
add_paragraph_with_bookmarks([
    "A Contratada deverá apresentar, no prazo de ", ("PB22_PrazoGarantiaDias_AnexoContrato", "____"),
    " dias corridos após a assinatura do Contrato, comprovante de prestação de garantia "
    "correspondente a ", ("PB22_PercentualGarantia_AnexoContrato", "____"),
    "% (percentual), conforme parâmetros do art. 439 do RLC, mediante caução em dinheiro, "
    "seguro-garantia ou fiança bancária.",
])
doc.add_paragraph(
    "O percentual da garantia é, em regra, de até 5% do valor do Contrato (art. 439, §2º, do "
    "RLC), podendo ser elevado para até 10% em contratações de grande vulto envolvendo "
    "complexidade técnica e riscos financeiros elevados (art. 439, §3º, do RLC). Nos casos de "
    "contratação de serviços com dedicação exclusiva de mão de obra, o valor da garantia "
    "deverá corresponder a 5% do valor do contrato, limitado ao equivalente a 2 (dois) meses "
    "do custo da folha de pagamento dos empregados da contratada que venham a participar da "
    "execução dos serviços contratados (art. 439, §4º, do RLC).")

doc.add_paragraph("CLÁUSULA SEXTA - DO PRAZO DE GARANTIA DA OBRA OU DO SERVIÇO DE ENGENHARIA")
add_paragraph_with_bookmarks([
    "O prazo de garantia dos serviços/obra, contado do recebimento definitivo, é de ",
    ("PB21_PrazoGarantiaServico_AnexoContrato", "____"),
    ". Para obras de engenharia, o Código Civil já garante 5 (cinco) anos contra vícios "
    "ocultos (art. 618), independentemente deste prazo.",
])

doc.add_paragraph("CLÁUSULA SÉTIMA - DA DOTAÇÃO ORÇAMENTÁRIA")
add_paragraph_with_bookmarks([
    "A despesa orçamentária da execução deste Contrato correrá à conta da Natureza da "
    "Despesa (", ("PB_ND_AnexoContrato", "XXXX"), "), PTRES (", ("PB_PTRES_AnexoContrato", "XXXX"),
    ") e Fonte (", ("PB_Fonte_AnexoContrato", "XXXX"), "), conforme Nota de Empenho a ser emitida.",
])

doc.add_paragraph("CLÁUSULA OITAVA - DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph(
    "São obrigações da Contratante as estabelecidas no Projeto Básico, não respondendo a "
    "Conab por compromissos assumidos pela Contratada com terceiros.")

doc.add_paragraph("CLÁUSULA NONA - DAS OBRIGAÇÕES DA CONTRATADA")
add_paragraph_with_bookmarks([
    "São obrigações da Contratada as estabelecidas no Projeto Básico, incluindo a "
    "apresentação das ART/RRT registradas no CREA/CAU em até ",
    ("PB18_PrazoARTContratoDias_AnexoContrato", "____"), " dias corridos da assinatura.",
])
doc.add_paragraph(
    "A Contratada obriga-se a manter, durante toda a execução do Contrato, em "
    "compatibilidade com as obrigações por ela assumidas, todas as condições de "
    "habilitação e qualificação exigidas no curso do procedimento licitatório ou de "
    "dispensa de licitação.")

doc.add_paragraph("CLÁUSULA DÉCIMA - DA ANOTAÇÃO DE RESPONSABILIDADE TÉCNICA (ART/RRT)")
add_paragraph_with_bookmarks([
    "A Contratada apresentará a ART/RRT referente ao objeto em até ",
    ("PB12_PrazoARTDias_AnexoContrato", "____"),
    " dias corridos da assinatura do Contrato, e a respectiva baixa em até ",
    ("PB12_PrazoBaixaARTDias_AnexoContrato", "____"), " dias úteis após a finalização.",
])

doc.add_paragraph("CLÁUSULA DÉCIMA PRIMEIRA - DO PAGAMENTO")
add_paragraph_with_bookmarks([
    "O pagamento será efetuado em até ", ("PB19_PrazoPagamentoDiasUteis_AnexoContrato", "____"),
    " dias úteis após o recebimento definitivo e atesto da Nota Fiscal/Fatura.",
])

doc.add_paragraph("CLÁUSULA DÉCIMA SEGUNDA - DO REAJUSTAMENTO / REPACTUAÇÃO")
add_bloco_variante("PB20_TipoCorrecao_AnexoContrato", {
    "Sim": [[
        "Observado o interregno mínimo de 12 (doze) meses, os preços serão reajustados/"
        "repactuados por solicitação da Contratada, pela variação do índice ",
        ("PB20_IndiceReajuste_AnexoContrato", ""), ".",
    ]],
    "Não": ["Não haverá reajuste ou repactuação de preços neste Contrato."],
})

doc.add_paragraph("CLÁUSULA DÉCIMA TERCEIRA - DA FISCALIZAÇÃO E GERENCIAMENTO DO CONTRATO")
doc.add_paragraph(
    "O acompanhamento e a fiscalização do Contrato observarão o previsto no Projeto Básico.")

doc.add_paragraph("CLÁUSULA DÉCIMA QUARTA - DA MATRIZ DE RISCOS E QUADRO DE INOVAÇÕES")
doc.add_paragraph(
    "A Matriz de Riscos — Anexo II do Projeto Básico — e, quando aplicável, o Quadro de "
    "Inovações — Anexo II-A —, constituem peças integrantes deste Contrato, "
    "independentemente de transcrição, definindo os riscos e responsabilidades entre "
    "Contratante e Contratada.")

doc.add_paragraph("CLÁUSULA DÉCIMA QUINTA - DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "A Contratada, em caso de inadimplemento de suas obrigações, garantido o contraditório e "
    "ampla defesa, ficará sujeita às sanções previstas no RLC, na Lei nº 13.303, de 2016, e no "
    "Projeto Básico: advertência; multa moratória; multa compensatória; multa rescisória; e "
    "suspensão do direito de participar de licitação e impedimento de contratar com a Conab, "
    "por até 02 (dois) anos.")
doc.add_paragraph(
    "a) advertência, aplicável nas infrações de menor gravidade, sem prejuízo à execução do "
    "objeto;")
add_paragraph_with_bookmarks([
    "b) multa por não regularização da documentação de habilitação: ",
    ("PB25_MultaNaoRegularizacaoHabilitacao_AnexoContrato", "____"), "% sobre o valor estimado para a contratação.",
])
add_paragraph_with_bookmarks([
    "c) multa pela prática das condutas elencadas nos arts. 576 a 580 do RLC: ",
    ("PB25_MultaCondutasArt576a580_AnexoContrato", "____"), "% sobre o valor estimado para a contratação.",
])
add_paragraph_with_bookmarks([
    "d) multa pela recusa em assinar o Contrato dentro do prazo estabelecido: ",
    ("PB25_MultaRecusaAssinarContrato_AnexoContrato", "____"), "% sobre o valor da contratação.",
])
doc.add_paragraph(
    "e) multa moratória de 0,2% (dois décimos por cento) sobre o valor do Contrato (ou "
    "valor anual do Contrato), por dia de atraso, até o limite de 15 (quinze) dias.")
doc.add_paragraph(
    "f) multa moratória de 0,3% (três décimos por cento) sobre o valor do Contrato (ou "
    "valor anual do Contrato), por dia de atraso, por período superior ao previsto na "
    "alínea anterior, até o limite de 15 (quinze) dias. Esgotado o prazo, poderá ocorrer a "
    "não aceitação do objeto, sem prejuízo da rescisão unilateral da avença.")
add_paragraph_with_bookmarks([
    "g) multa compensatória: ", ("PB25_MultaCompensatoria_AnexoContrato", "____"),
    "% do valor total do Contrato, pelo descumprimento grave ou rescisão por culpa da "
    "Contratada.",
])
add_paragraph_with_bookmarks([
    "h) multa rescisória: ", ("PB25_MultaRescisoria_AnexoContrato", "____"),
    "% nos casos de rescisão unilateral por descumprimento contratual.",
])
doc.add_paragraph(
    "i) multa de 0,2% a 3,2% por dia sobre o valor do Contrato (ou valor anual/mensal do "
    "Contrato), conforme grau da infração cometida (Tabelas 1 e 2 abaixo).")
tabela1_pb_anexo_vi = doc.add_table(rows=1, cols=2)
tabela1_pb_anexo_vi.style = "Table Grid"
tabela1_pb_anexo_vi.rows[0].cells[0].text = "TABELA 1 — DESCRIÇÃO DA INFRAÇÃO"
tabela1_pb_anexo_vi.rows[0].cells[1].text = "GRAU"
for descricao, grau in [
    ("Permitir situação que crie a possibilidade de causar dano físico, lesão corporal ou "
     "consequências letais, por ocorrência", "5"),
    ("Suspender ou interromper, salvo motivo de força maior ou caso fortuito, a execução "
     "contratual, por dia e por unidade de atendimento", "4"),
    ("Servir-se de empregado sem qualificação para executar os serviços contratados, por "
     "empregado e por dia", "3"),
    ("Recusar-se a executar serviço determinado pela fiscalização, por serviço e por dia", "2"),
    ("Deixar de cumprir determinação formal ou instrução complementar da fiscalização, por "
     "ocorrência", "2"),
    ("Deixar de substituir empregado alocado que não atenda às necessidades do serviço, por "
     "funcionário e por dia", "1"),
    ("Deixar de cumprir quaisquer itens do Projeto Básico e seus Anexos não previstos nesta "
     "tabela, após reincidência formalmente notificada pela fiscalização, por item e por "
     "ocorrência", "3"),
    ("Deixar de indicar e manter durante a execução do Contrato os prepostos previstos no "
     "Contrato, por dia", "1"),
    ("Deixar de providenciar treinamento para seus empregados conforme previsto na relação "
     "de obrigações da Contratada, por empregado", "1"),
]:
    linha = tabela1_pb_anexo_vi.add_row()
    linha.cells[0].text = descricao
    linha.cells[1].text = grau
tabela2_pb_anexo_vi = doc.add_table(rows=1, cols=2)
tabela2_pb_anexo_vi.style = "Table Grid"
tabela2_pb_anexo_vi.rows[0].cells[0].text = "GRAU"
tabela2_pb_anexo_vi.rows[0].cells[1].text = "CORRESPONDÊNCIA (percentual por dia)"
for grau, percentual in [("1", "0,2%"), ("2", "0,4%"), ("3", "0,8%"), ("4", "1,6%"), ("5", "3,2%")]:
    linha = tabela2_pb_anexo_vi.add_row()
    linha.cells[0].text = grau
    linha.cells[1].text = percentual

doc.add_paragraph("CLÁUSULA DÉCIMA SEXTA - DA RESCISÃO")
doc.add_paragraph(
    "A rescisão observará o disposto no Projeto Básico e nos arts. 568 a 572 e 582 a 593 do "
    "RLC.")

doc.add_paragraph("CLÁUSULA DÉCIMA SÉTIMA - DA SUBCONTRATAÇÃO")
doc.add_paragraph("As condições de subcontratação, quando admitidas, observam o disposto no Projeto Básico.")

doc.add_paragraph("CLÁUSULA DÉCIMA OITAVA - DA ALTERAÇÃO DO CONTRATO")
doc.add_paragraph(
    "O presente Contrato poderá ser alterado nas hipóteses do art. 510 do RLC, vedando-se "
    "alterações que resultem em violação ao dever de licitar. A Contratada poderá aceitar "
    "acréscimos ou supressões de até 25% (vinte e cinco por cento) do valor inicial "
    "atualizado do contrato.")

doc.add_paragraph("CLÁUSULA DÉCIMA NONA - DA ALTERAÇÃO SUBJETIVA")
doc.add_paragraph(
    "É admissível a fusão, cisão ou incorporação da Contratada com/em outra pessoa jurídica, "
    "observados os requisitos de habilitação e a anuência expressa da Conab.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA - DAS VEDAÇÕES")
doc.add_paragraph(
    "É vedado à Contratada: a) caucionar ou utilizar este Contrato para qualquer operação "
    "financeira; b) interromper a execução contratual sob alegação de inadimplemento da "
    "Contratante, salvo nos casos previstos em lei; c) empregar menor de 18 anos em trabalho "
    "noturno, perigoso ou insalubre e não empregar menor de 16 anos, salvo na condição de "
    "aprendiz a partir de 14 anos; d) subcontratar, na íntegra, o objeto da contratação, "
    "salvo quando a subcontratação parcial for admitida no Projeto Básico; e e) solicitar, "
    "por meio de seus empregados, serviços, materiais ou equipamentos a empresas "
    "terceirizadas que prestem serviços à Conab (art. 155 do RLC).")

doc.add_paragraph("CLÁUSULA VIGÉSIMA PRIMEIRA - DA VEDAÇÃO AO NEPOTISMO")
doc.add_paragraph(
    "Conforme disposto no art. 12, §1º, inciso II, do RLC e no art. 7º do Decreto nº 7.203, "
    "de 2010, fica vedada a contratação de empregado ou dirigente da Conab como pessoa "
    "física, ou de pessoa/empresa com relação de parentesco até o terceiro grau civil com "
    "autoridade do Ministério do Desenvolvimento Agrário e Agricultura Familiar (MDA), dirigente ou empregado da "
    "Conab com atribuições na área de licitação e contratação.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA SEGUNDA - DA VINCULAÇÃO")
doc.add_paragraph(
    "Consideram-se integrantes do presente Instrumento Contratual o Projeto Básico da "
    "Dispensa de Licitação e seus Anexos, a Proposta da Contratada, no que couber, e demais "
    "documentos pertinentes, independentemente de transcrição.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA TERCEIRA - DA LEGISLAÇÃO APLICÁVEL")
doc.add_paragraph(
    "Os casos omissos serão decididos pela Contratante, segundo as disposições do RLC, da "
    "Lei nº 13.303, de 2016, e demais normas aplicáveis, e, subsidiariamente, do Código de "
    "Defesa do Consumidor.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA QUARTA - DA PROTEÇÃO DE DADOS PESSOAIS")
doc.add_paragraph(
    "As Partes obrigam-se a cumprir a legislação vigente sobre proteção de dados pessoais, "
    "em especial a Lei nº 13.709, de 2018 (Lei Geral de Proteção de Dados Pessoais), no que "
    "for aplicável à execução deste Contrato.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA QUINTA - DA PUBLICAÇÃO")
doc.add_paragraph(
    "A publicação do extrato do presente Contrato deverá ser providenciada pela Contratante "
    "até o 5º (quinto) dia útil do mês subsequente ao da sua assinatura no Diário Oficial da "
    "União, conforme o art. 480 do RLC.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA SEXTA - DO FORO")
doc.add_paragraph(
    "As partes elegem o foro da Justiça Federal, Seção Judiciária do Estado da Sureg ou de "
    "Brasília-DF, no caso da Matriz, competente para dirimir quaisquer dúvidas suscitadas em "
    "razão deste Contrato, que não puderem ser resolvidas de comum acordo.")

doc.add_paragraph(
    "Por estarem justas e acordadas, firmam o presente instrumento, em 03 (três) vias, de "
    "igual teor e forma, diante das testemunhas abaixo identificadas.")
doc.add_paragraph("Pela Contratante: ___________________     Pela Contratada: ___________________")
doc.add_paragraph("Testemunha 1: ___________________     Testemunha 2: ___________________")
add_bloco_variante_sentinela_fim(_bid_anexo_contrato)

doc.save("Template_PB.docx")
print("Template_PB.docx gerado com sucesso.")
