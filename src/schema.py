"""Módulo de schema e estrutura de dados do Excel.

Define a estrutura dos dados extraídos e organização das planilhas.
"""

from dataclasses import dataclass

@dataclass(frozen=True)
class ColSpec:
    aba: str
    secao: str
    campo: str
    key: str


# Schema do Excel alinhado com "Sistema/Dados que precisam ser extraídos.txt" (ordem e nomes).
SCHEMA_ESTRUTURA: list[tuple[str, list[tuple[str, list[str]]]]] = [
    (
        'Titular',
        [
            ('Dados do Empreendimento', ['Tipo de Cadastro', 'Loteamento']),
            (
                'Minha Casa, Minha Vida',
                [
                    'Você ou alguém da sua família possui imóvel (como herança, doação ou aluguel vitalício) em qualquer lugar do Brasil?',
                    'A renda familiar mensal ultrapassa R$ 2.850,00?',
                    'Mora no Município há pelo menos 2 anos na mesma cidade?',
                ],
            ),
            (
                'Dados Gerais',
                [
                    'Nome',
                    'CPF/CNPJ',
                    'Tipo de Pessoa',
                    'Tipo de Documento',
                    'Nº Documento',
                    'Data de Emissão',
                    'Órgão Emissor',
                    'UF',
                    'Data de Vencimento',
                    'Naturalidade (Cidade)',
                    'UF Naturalidade (Estado)',
                    'Nacionalidade (País)',
                    'Data de Nascimento',
                    'Raça',
                    'Gênero',
                    'Contrato Distratado ou Rescindido Involutariamente',
                    'Recebe Atendimento Socio-assistencial do Município',
                    'Segurança frequente no bairro (Ronda GCM, PM)',
                    'Situação Atual do Domicílio',
                    'Tipo de Beneficiário Programa Social',
                    'Proprietário do Imóvel?',
                    'Recebeu/Participou de REURB?',
                ],
            ),
            (
                'Informações do CadÚnico',
                [
                    'A família esta inscrita no CAD Único?',
                    'O responsável familiar é mulher monoparental e responsável por maior número de filhos de 0 a 17 anos declarado no Cadúnico?',
                    'O responsável familiar é homem monoparental e responsável por maior número de filhos de 0 a 17 anos declarado no Cadúnico?',
                    'Família unipessoal, sem dependente, declarada no Cadúnico residente no domícilio?',
                    'A requerente é mulher em situação de violência doméstica/familiar, com medida protetiva de urgência declarada?',
                    'Família com pessoa(s) negra(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) LGBT na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) Quilombola(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) de povos tradicionais na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) em situação de rua na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Programas CAD Único',
                    'Mulheres Empreendedoras',
                    'Tratamento Oncológico/Hemodialise',
                ],
            ),
            (
                'Cadastro Preferencial',
                ['Cadastro Preferencial', 'Deficiência', 'Doença', 'Possui Microcefalia', 'Tem Veículo Próprio'],
            ),
            ('Dados dos Pais', ['Nome do Pai', 'Nome da Mãe']),
            (
                'Educação e Profissão',
                ['Escolaridade', 'Escola', 'Situação de Emprego', 'Profissão', 'Profissão (se Outro)', 'Tempo de Serviço'],
            ),
            (
                'Situação Marital',
                ['Estado Civil', 'Data do Casamento/União', 'Regime do Casamento', 'Regime do Casamento (se Outro)'],
            ),
            (
                'CAD',
                [
                    'NIS (PIS/PASEP)',
                    'Nº da Carteira de Trabalho',
                    'Nº de Série da Carteira de Trabalho',
                    'UF da Carteira de Trabalho',
                    'Nº do Título de Eleitor',
                    'Zona Eleitoral',
                    'Seção Eleitoral',
                    'Data de Emissão',
                    'UF Título de Eleitor',
                    'Cidade do Título de Eleitor',
                ],
            ),
            (
                'Dados de Contato',
                ['Email', 'Telefone', 'Telefone Celular', 'Telefone Comercial', 'Telefone Para Recados'],
            ),
            ('Parecer Social', ['Diagnóstico Social', 'Observações']),
            (
                'Documentos',
                [
                    'CPF/RG Frente',
                    'CPF/RG Verso',
                    'CPF/RG (Cônjuge) Frente',
                    'CPF/RG (Cônjuge) Verso',
                    'Comprovante de Residência',
                    'Certidão de Nascimento',
                    'Certidão de Casamento/União Estável',
                    'Foto Moradia (Frente)',
                    'Foto Moradia (Lateral)',
                    'Conferência Topografica',
                    'Foto do Candidato',
                    'Comprovante de Renda',
                    'Contrato de Compra e Venda',
                    'Outras Comprovações de Posse',
                ],
            ),
        ],
    ),
    (
        'Segundo Titular',
        [
            (
                'Perguntas de Triagem',
                [
                    'Você ou alguém da sua família possui imóvel (como herança, doação ou aluguel vitalício) em qualquer lugar do Brasil?',
                    'A renda familiar mensal ultrapassa R$ 2.850,00?',
                    'Mora no Município há pelo menos 2 anos na mesma cidade?',
                ],
            ),
            (
                'Dados Gerais',
                [
                    'Nome',
                    'CPF/CNPJ',
                    'Tipo de Pessoa',
                    'Tipo de Documento',
                    'Nº Documento',
                    'Data de Emissão',
                    'Órgão Emissor',
                    'UF',
                    'Data de Vencimento',
                    'Naturalidade (Cidade)',
                    'UF Naturalidade (Estado)',
                    'Nacionalidade (País)',
                    'Data de Nascimento',
                    'Raça',
                    'Gênero',
                    'Contrato Distratado ou Rescindido Involutariamente',
                    'Recebe Atendimento Socio-assistencial do Município',
                    'Segurança frequente no bairro (Ronda GCM, PM)',
                    'Situação Atual do Domicílio',
                    'Tipo de Beneficiário Programa Social',
                    'Proprietário do Imóvel?',
                    'Recebeu/Participou de REURB?',
                ],
            ),
            (
                'Informações do CadÚnico',
                [
                    'A família esta inscrita no CAD Único?',
                    'O responsável familiar é mulher monoparental e responsável por maior número de filhos de 0 a 17 anos declarado no Cadúnico?',
                    'O responsável familiar é homem monoparental e responsável por maior número de filhos de 0 a 17 anos declarado no Cadúnico?',
                    'Família unipessoal, sem dependente, declarada no Cadúnico residente no domícilio?',
                    'A requerente é mulher em situação de violência doméstica/familiar, com medida protetiva de urgência declarada?',
                    'Família com pessoa(s) negra(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) LGBT na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) Quilombola(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) de povos tradicionais na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) em situação de rua na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Programas CAD Único',
                    'Mulheres Empreendedoras',
                    'Tratamento Oncológico/Hemodialise',
                ],
            ),
            (
                'Cadastro Preferencial',
                ['Cadastro Preferencial', 'Deficiência', 'Doença', 'Possui Microcefalia', 'Tem Veículo Próprio'],
            ),
            ('Dados dos Pais', ['Nome do Pai', 'Nome da Mãe']),
            (
                'Educação e Profissão',
                ['Escolaridade', 'Escola', 'Situação de Emprego', 'Profissão', 'Profissão (se Outro)', 'Tempo de Serviço'],
            ),
            (
                'Situação Marital',
                ['Estado Civil', 'Data do Casamento/União', 'Regime do Casamento', 'Regime do Casamento (se Outro)'],
            ),
            (
                'CAD',
                [
                    'NIS (PIS/PASEP)',
                    'Nº da Carteira de Trabalho',
                    'Nº de Série da Carteira de Trabalho',
                    'UF da Carteira de Trabalho',
                    'Nº do Título de Eleitor',
                    'Zona Eleitoral',
                    'Seção Eleitoral',
                    'Data de Emissão',
                    'UF Título de Eleitor',
                    'Cidade do Título de Eleitor',
                ],
            ),
            (
                'Dados de Contato',
                ['Email', 'Telefone', 'Telefone Celular', 'Telefone Comercial', 'Telefone Para Recados'],
            ),
            ('Parecer Social', ['Diagnóstico Social', 'Observações']),
            (
                'Documentos',
                [
                    'CPF/RG Frente',
                    'CPF/RG Verso',
                    'CPF/RG (Cônjuge) Frente',
                    'CPF/RG (Cônjuge) Verso',
                    'Certidão de Nascimento',
                    'Certidão de Casamento/União Estável',
                    'Foto do Candidato',
                    'Comprovante de Renda',
                ],
            ),
        ],
    ),
    (
        'Composição Familiar',
        [
            (
                'Dados Gerais',
                [
                    'Tipo',
                    'Nome',
                    'CPF',
                    'Tem CPF?',
                    'Tipo de Pessoa',
                    'RG',
                    'Data de Emissão do RG',
                    'Órgão Emissor do RG',
                    'UF do RG',
                    'Naturalidade',
                    'UF Naturalidade',
                    'Nacionalidade',
                    'Data de Nascimento',
                    'Nº da Certidão de Nascimento',
                    'Raça',
                    'Gênero',
                    'Telefone',
                    'Estado Civil',
                    'Pessoa Negra?',
                    'Pessoa Quilombola?',
                    'Pessoa LGBTQIAPN+?',
                    'Situação de Rua?',
                ],
            ),
            (
                'Educação e Profissão',
                ['Escolaridade', 'Escola', 'Situação de Emprego', 'Profissão', 'Profissão (se Outro)', 'Tempo de Serviço'],
            ),
            ('Cadastro Preferencial', ['Cadastro Preferencial', 'Deficiência', 'Doença']),
            (
                'CAD',
                [
                    'NIS (PIS/PASEP)',
                    'Nº da Carteira de Trabalho',
                    'Nº de Série da Carteira de Trabalho',
                    'UF da Carteira de Trabalho',
                    'Família Inscrita no CAD Único',
                    'Programas CAD Único',
                    'Nº do Título de Eleitor',
                    'Zona Eleitoral',
                    'Seção Eleitoral',
                    'Data de Emissão',
                    'UF Título de Eleitor',
                    'Cidade do Título de Eleitor',
                ],
            ),
            (
                'Documentos',
                [
                    'CPF/RG Frente',
                    'CPF/RG Verso',
                    'CPF/RG (Cônjuge) Frente',
                    'CPF/RG (Cônjuge) Verso',
                    'Certidão de Nascimento',
                    'Certidão de Casamento/União Estável',
                    'Foto do Candidato',
                    'Comprovante de Renda',
                ],
            ),
        ],
    ),
    (
        'Renda',
        [
            ('Dados da Renda', ['Tipo de Renda', 'Pessoa', 'Valor da Renda', 'Tipo se Apto Para Reurb']),
        ],
    ),
    (
        'Endereço',
        [
            (
                'Dados do Endereço',
                [
                    'Lote',
                    'Quadra',
                    'CEP',
                    'UF',
                    'Cidade',
                    'Bairro',
                    'Endereço',
                    'Número',
                    'Complemento',
                    'Tipo de Moradia',
                    'Possui outro imóvel?',
                    'Forma de aquisição',
                    'Relação com o imóvel',
                    'Uso do imóvel',
                    'Há famílias que habitam ou trabalham a, no máximo, 28 km de distância do Centro do Empreendimento?',
                    'Reside no município desde (ano)',
                ],
            ),
            (
                'Documentos',
                ['Comprovante de Residência', 'Foto Moradia (Frente)', 'Foto Moradia (Lateral)', 'Conferência Topografica'],
            ),
        ],
    ),
    (
        'Imóvel',
        [
            (
                'Aquisição e Infraestrutura',
                [
                    'Lote',
                    'Quadra',
                    'Área total (m²)',
                    'Área construída (m²)',
                    'Tipo de terreno',
                    'Tipo de imóvel',
                    'Nº de cadastro imobiliário (IPTU)',
                    'CEP',
                    'UF',
                    'Cidade',
                    'Bairro',
                    'Endereço',
                    'Número',
                    'Complemento',
                    'Infraestrutura do entorno',
                ],
            ),
            (
                'Informações do Imóvel',
                [
                    'Data do início do contrato',
                    'Data do fim do contrato',
                    'Data de aquisição',
                    'Valor de aquisição',
                    'Nome do edifício',
                    'Nº de matrícula',
                    'Nome do cartório',
                ],
            ),
            (
                'Documentos',
                ['Comprovante de Residência', 'Foto Moradia (Frente)', 'Foto Moradia (Lateral)', 'Conferência Topografica'],
            ),
        ],
    ),
    (
        'Documentos',
        [
            (
                'Titular',
                [
                    'CPF/RG Frente',
                    'CPF/RG Verso',
                    'CPF/RG (Cônjuge) Frente',
                    'CPF/RG (Cônjuge) Verso',
                    'Comprovante de Residência',
                    'Certidão de Nascimento',
                    'Certidão de Casamento/União Estável',
                    'Foto Moradia (Frente)',
                    'Foto Moradia (Lateral)',
                    'Conferência Topografica',
                    'Foto do Candidato',
                    'Comprovante de Renda',
                    'Contrato de Compra e Venda',
                    'Outras Comprovações de Posse',
                    'MCMV',
                ],
            ),
            (
                'Segundo Titular',
                [
                    'CPF/RG Frente',
                    'CPF/RG Verso',
                    'CPF/RG (Cônjuge) Frente',
                    'CPF/RG (Cônjuge) Verso',
                    'Certidão de Nascimento',
                    'Certidão de Casamento/União Estável',
                    'Foto do Candidato',
                    'Comprovante de Renda',
                ],
            ),
            (
                'Dependentes',
                [
                    'CPF/RG Frente',
                    'CPF/RG Verso',
                    'CPF/RG (Cônjuge) Frente',
                    'CPF/RG (Cônjuge) Verso',
                    'Certidão de Nascimento',
                    'Certidão de Casamento/União Estável',
                    'Foto do Candidato',
                    'Comprovante de Renda',
                ],
            ),
        ],
    ),
    (
        'Questionário',
        [
            (
                'Perguntas',
                [
                    '1. A mulher é a responsável pela unidade familiar?',
                    '2. Há pessoa negra na composição familiar?',
                    '3. Há pessoa com deficiência na composição familiar, comprovada por avaliação biopsicossocial (Lei nº 13.146/2015 e Decreto nº 11.063/2022)?',
                    '4. Há idoso na composição familiar, comprovado por documento civil com data de nascimento?',
                    '5. Há criança ou adolescente na composição familiar, comprovado por certidão de nascimento, guarda ou tutela?',
                    '6. Há pessoa com câncer ou doença rara crônica e degenerativa na família, comprovado por laudo médico?',
                    '7. Há mulheres vítimas de violência doméstica/familiar na família, comprovado por registro no Cadastro Nacional de Violência Doméstica (Lei Maria da Penha)?',
                    '8. Há integrantes de povos indígenas ou quilombolas na família, declarados no CadÚnico?',
                    '9. A família reside em área de risco (deslizamentos, inundações etc.), conforme mapeamento do PMRR, CPRM ou Defesa Civil?',
                    '10. O beneficiário teve contrato distratado ou rescindido involuntariamente, conforme normativo do Ente Público?',
                    '11. Atualmente é atendido pelas redes Socioassistenciais do Município?',
                ],
            ),
        ],
    ),
]


def _flatten_schema(schema: list[tuple[str, list[tuple[str, list[str]]]]]) -> list[ColSpec]:
    specs: list[ColSpec] = []
    for aba, secoes in schema:
        for secao, campos in secoes:
            for campo in campos:
                key = f"{aba}|||{secao}|||{campo}"
                specs.append(ColSpec(aba=aba, secao=secao, campo=campo, key=key))
    return specs


COL_SPECS: list[ColSpec] = _flatten_schema(SCHEMA_ESTRUTURA)
COL_KEY_BY_TRIPLE: dict[tuple[str, str, str], str] = {(c.aba, c.secao, c.campo): c.key for c in COL_SPECS}


_ABORT_REQUESTED = False


def _request_abort():
    global _ABORT_REQUESTED
    _ABORT_REQUESTED = True


def _check_abort():
    if _ABORT_REQUESTED:
        raise KeyboardInterrupt


def _col_key(aba: str, secao: str, campo: str) -> str:
    return COL_KEY_BY_TRIPLE.get((aba, secao, campo), f"{aba}|||{secao}|||{campo}")


