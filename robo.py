import openpyxl
import pdfplumber
import os
import re
import itertools
import shutil
import tempfile
from copy import copy
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font

def limpar_numero(texto):
    # remove todos os caracteres nao numericos de uma string
    return re.sub(r'\D', '', str(texto))

def converter_valor_para_float(texto_valor):
    # converte strings monetarias para float tratando separadores de milhar e decimal
    v = str(texto_valor).strip()
    pos_ponto   = v.rfind('.')
    pos_virgula = v.rfind(',')
    if pos_virgula > pos_ponto:
        v = v.replace('.', '').replace(',', '.')
    elif pos_ponto > pos_virgula:
        v = v.replace(',', '')
    else:
        v = v.replace(',', '.')
    return float(v)

def extrair_valor_pdf(texto, é_iss=False):
    # busca valores numericos no texto do pdf usando regex baseada em contexto fiscal
    padrao_puro = r'([\d]{1,3}(?:[.,]\d{3})*[.,]\d{2}|\d+[.,]\d{2})'

    if é_iss:
        frases = [r'Total\s*:?']
    else:
        frases = [
            r'TOTAL',
            r'Total\s+a\s+[Rr]ecolher\s*:?',
            r'Valor\s+[Tt]otal',
            r'Valor\s+a\s+[Pp]agar',
        ]

    for frase in frases:
        padrao_completo = frase + r'[\s\n]{0,150}' + padrao_puro
        match = re.search(padrao_completo, texto, re.IGNORECASE | re.DOTALL)
        if match:
            try:
                return converter_valor_para_float(match.group(1))
            except ValueError:
                continue

    encontrados_cifrao = re.findall(r'R?\$[\s\n]{0,100}' + padrao_puro, texto)
    if encontrados_cifrao:
        valores = []
        for v in encontrados_cifrao:
            num = re.search(padrao_puro, v)
            if num:
                try:
                    valores.append(converter_valor_para_float(num.group(1)))
                except ValueError:
                    continue
        if valores:
            return max(valores)

    encontrados = re.findall(padrao_puro, texto)
    valores = []
    for v in encontrados:
        try:
            valores.append(converter_valor_para_float(v))
        except ValueError:
            continue
    return max(valores) if valores else None

def extrair_destino_pdf(texto_pdf):
    # identifica a unidade federativa de destino dentro do documento pdf
    texto = texto_pdf.replace('\xa0', ' ').replace('\u200b', '')
    padroes = [
        (r'01\.CF/([A-Z]{2})', 0),
        (r'UF\s+Favorecida[^A-Za-z]{0,80}([A-Z]{2})(?![A-Za-z])', re.IGNORECASE | re.DOTALL),
        (r'UF\s+do\s+Favorecid[oa][^A-Za-z]{0,80}([A-Z]{2})(?![A-Za-z])', re.IGNORECASE | re.DOTALL),
        (r'Favorecid[oa].{0,300}?\bUF[:\s]{1,40}([A-Z]{2})(?![A-Za-z])', re.IGNORECASE | re.DOTALL),
        (r'(?:Destinat[aá]rio|Benefici[aá]rio|Credor|Favorecido).{0,400}?UF[:\s]{1,40}([A-Z]{2})(?![A-Za-z])', re.IGNORECASE | re.DOTALL),
    ]

    for padrao, flags in padroes:
        try:
            match = re.search(padrao, texto, flags)
            if match:
                sigla = match.group(1).upper().strip()
                if len(sigla) == 2 and sigla.isalpha():
                    return sigla
        except re.error:
            continue
    return None

def identificar_regra_completa(regras_filial, texto_pdf, id_filial, cnpj_ie_pdf):
    # valida se as regras fiscais da filial coincidem com os dados extraidos do pdf
    if not regras_filial:
        return None

    destino_pdf = extrair_destino_pdf(texto_pdf)
    texto_norm = texto_pdf.replace('\xa0', ' ').replace('\u200b', '')
    cnpj_pdf_limpo = re.sub(r'\D', '', str(cnpj_ie_pdf))

    if not destino_pdf:
        return None

    for regra in regras_filial:
        cod_regra = str(regra['codigo']).strip()
        padrao_cod = r'(?<![0-9])' + re.escape(cod_regra) + r'(?![0-9])'
        
        if not re.search(padrao_cod, texto_norm):
            continue

        cnpj_regra = re.sub(r'\D', '', str(regra.get('cnpj', '')))
        if cnpj_regra and cnpj_pdf_limpo and cnpj_regra != cnpj_pdf_limpo:
            continue

        destino_regra = str(regra.get('destino', '')).upper().strip()
        if destino_regra != destino_pdf:
            continue

        return regra
    return None

def verificar_destino_por_estilo(texto_pdf, regra):
    # verifica a uf de destino baseando-se em estilos pre-definidos de layouts de guias
    destino_esperado = str(regra.get('destino', ''))
    estilo_d         = regra.get('estilo_d')

    if not destino_esperado or destino_esperado.upper() in ('NONE', 'N/A', '') or estilo_d is None:
        return {'verificado': False, 'ok': True, 'destino_pdf': None,
                'destino_esperado': destino_esperado, 'motivo': 'N/A'}

    texto = texto_pdf.replace('\xa0', ' ').replace('\u200b', '')
    destino_pdf = None
    destinos_validos = [d.strip().upper() for d in destino_esperado.split('/') if d.strip()]

    try:
        matches = []
        ufs = r'(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)'
        
        if estilo_d == 1:
            padrao = r'\b' + ufs + r'\b'
            matches = re.findall(padrao, texto)
        elif estilo_d == 2:
            padroes = [
                r'(?i:Dados\s+do\s+Destinat[aá]rio).{0,150}?\b' + ufs + r'\b',
                r'(?i:UF\s+FAVORECIDA).{0,150}?\b' + ufs + r'\b'
            ]
            for p in padroes:
                matches.extend(re.findall(p, texto, re.DOTALL))
        elif estilo_d == 3:
            padrao = r'(?i:01\.CF/)' + ufs + r'\b'
            matches = re.findall(padrao, texto)

        if matches:
            siglas = [m.upper() for m in matches]
            for sigla in siglas:
                if sigla in destinos_validos:
                    destino_pdf = sigla
                    break
            if not destino_pdf:
                destino_pdf = None if estilo_d == 1 else siglas[0]

    except Exception as e:
        return {'verificado': True, 'ok': False, 'destino_pdf': None,
                'destino_esperado': destino_esperado,
                'motivo': f'erro ao extrair destino: {e}'}

    if not destino_pdf:
        return {'verificado': True, 'ok': False, 'destino_pdf': None,
                'destino_esperado': destino_esperado,
                'motivo': f'uf nao encontrada no pdf (estilo {estilo_d})'}

    ok = destino_pdf in destinos_validos
    motivo = 'destino ok' if ok else f'destino divergente: pdf={destino_pdf} / esperado={destino_esperado}'

    return {'verificado': True, 'ok': ok, 'destino_pdf': destino_pdf,
            'destino_esperado': destino_esperado, 'motivo': motivo}

def nome_base_arquivo(nome_arquivo):
    # normaliza o nome do arquivo para agrupamento ignorando sufixos variaveis
    nome = os.path.splitext(nome_arquivo)[0].strip()

    # anonimizado: parametro para guias de importacao da empresa
    m_import = re.search(
        r'(.+?sua_empresa\s+icms\s+importação\s+parc\s+\d+_\d+)',
        nome, re.IGNORECASE
    )
    if m_import:
        return m_import.group(1).strip()
    
    # parametro para guias de iss matriz
    m_mat = re.search(
        r'(.+?guia\s+iss\s+(\d+)(?:\s*\(\d+\))?\s*-\s*mat)',
        nome, re.IGNORECASE
    )
    if m_mat:
       prefixo = re.match(r'(.+?guia\s+iss\s+)', nome, re.IGNORECASE).group(1)
       codigo  = m_mat.group(2)
       return f"{prefixo}{codigo} - mat".strip()

    # parametro para guias de iss filial especifica
    m_tag = re.search(
        r'(.+?guias\s+iss\s+filial_exemplo(?:\s*\(\d+\))?\s+(\d+))',
        nome, re.IGNORECASE
    )
    if m_tag:
        prefixo = re.match(r'(.+?guias\s+iss\s+filial_exemplo)', nome, re.IGNORECASE).group(1)
        codigo  = m_tag.group(2)
        return f"{prefixo} {codigo}".strip()

    nome = re.sub(r'\s*\(\d+\)\s*', ' ', nome)
    nome = re.sub(r'\s+\d{1,4}\s*$', '', nome).strip()
    return nome

def é_guia_iss(nome_arquivo):
    # verifica se o arquivo se refere a um imposto iss
    return 'ISS' in nome_arquivo.upper()

def eh_pendencia(celula):
    # identifica se uma celula excel esta marcada como pendencia atraves da cor da fonte
    hexs_vermelhos  = {'FFFF0000', 'FFC00000', 'FFFA0000', 'FFE00000', 'FF0000'}
    indices_validos = {2, 5, 9, 10}   
    temas_vermelhos = {1, 9}          

    try:
        fonte = celula.font
        if fonte is None or fonte.color is None:
            return False

        color = fonte.color
        try:
            rgb_raw = color.rgb                      
            rgb = str(rgb_raw).upper().replace(' ', '')
            if rgb in hexs_vermelhos or rgb.endswith('FF0000'):
                return True
        except Exception:
            pass

        try:
            idx_raw = color.index                      
            idx = int(idx_raw)
            if 1 <= idx <= 56 and idx in indices_validos:
                return True
        except Exception:
            pass

        try:
            tema_raw = color.theme                     
            tema = int(tema_raw)
            if tema in temas_vermelhos:
                return True
        except Exception:
            pass

        return False
    except Exception:
        return False

def _construir_mapa_formatacao_condicional(planilha):
    # mapeia linhas que possuem formatacao condicional de erro (cor vermelha)
    linhas_vermelhas = set()
    try:
        cf_rules = planilha.conditional_formatting._cf_rules  
    except Exception:
        return linhas_vermelhas

    hexs_cf_vermelhos = {'FFFF0000', 'FFC00000', 'FFFA0000', 'FFE00000', 'FF FF0000', 'FF0000'}

    def _cor_e_vermelha(color_obj):
        if color_obj is None:
            return False
        c = str(getattr(color_obj, 'rgb', '') or getattr(color_obj, 'index', '')).upper().strip()
        return c in hexs_cf_vermelhos or c.endswith('FF0000')

    def _regra_e_vermelha(rule):
        try:
            if rule.fill and rule.fill.fgColor and _cor_e_vermelha(rule.fill.fgColor): return True
            if rule.font and rule.font.color and _cor_e_vermelha(rule.font.color): return True
        except Exception: pass
        return False

    for cell_range_str, rules in cf_rules.items():
        for rule in rules:
            if rule.type != 'expression': continue
            formula = (rule.formula[0] if rule.formula else '').strip()
            m = re.match(r'=?\$?([A-Z]+)\$?(\d+)\s*=\s*"x"', formula, re.IGNORECASE)
            if not m: continue
            col_controle_idx = 0
            for ch in m.group(1).upper():
                col_controle_idx = col_controle_idx * 26 + (ord(ch) - ord('A') + 1)
            try:
                from openpyxl.utils import range_boundaries
                min_col, min_row, max_col, max_row = range_boundaries(cell_range_str)
                for lin in range(min_row, max_row + 1):
                    val_ctrl = planilha.cell(row=m.group(2), column=col_controle_idx).value
                    if str(val_ctrl).strip().lower() == 'x':
                        linhas_vermelhas.add(lin)
            except Exception: continue
    return linhas_vermelhas

def buscar_valor_na_planilha(planilha, col_idx, valor_guia, margem, linha_ini=11, linha_fim=75):
    # localiza um valor numerico dentro de uma coluna da planilha respeitando uma margem de erro
    candidatos_vermelhos = []
    candidatos_outros = []
    linhas_cf_vermelhas = _construir_mapa_formatacao_condicional(planilha)

    for lin in range(linha_ini, linha_fim + 1):
        celula = planilha.cell(row=lin, column=col_idx)
        val_celula = celula.value
        if val_celula is None: continue
        try:
            valor_ap  = float(val_celula) if isinstance(val_celula, (int, float)) else converter_valor_para_float(str(val_celula))
            diferenca = round(abs(valor_ap - valor_guia), 2)
            if diferenca <= margem:
                é_vermelho = eh_pendencia(celula) or (lin in linhas_cf_vermelhas)
                info = (lin, valor_ap, diferenca, "vermelho" if é_vermelho else "outra cor")
                if é_vermelho: candidatos_vermelhos.append(info)
                else: candidatos_outros.append(info)
        except Exception: continue

    if candidatos_vermelhos: return min(candidatos_vermelhos, key=lambda x: x[2])
    if candidatos_outros: return min(candidatos_outros, key=lambda x: x[2])
    return None

def parsear_aba_como_data(nome):
    # extrai ano e mes do nome da aba da planilha excel
    nome_limpo = nome.strip().rstrip('.')
    m = re.fullmatch(r'(\d{4})\s+(\d{1,2})', nome_limpo)
    if m: return (int(m.group(1)), int(m.group(2)))
    m = re.fullmatch(r'(\d{4})\((\d{1,2})\)', nome_limpo)
    if m: return (int(m.group(1)), int(m.group(2)))
    return None

def listar_abas_recentes(caminho_planilha, quantidade=3):
    # retorna as abas mais recentes baseadas na cronologia extraida dos nomes
    meses_pt = {1:'janeiro', 2:'fevereiro', 3:'março', 4:'abril', 5:'maio', 6:'junho', 7:'julho', 8:'agosto', 9:'setembro', 10:'outubro', 11:'novembro', 12:'dezembro'}
    wb = load_workbook(caminho_planilha, read_only=True, data_only=True)
    todas_abas = wb.sheetnames
    wb.close()
    abas_com_data = [(parsear_aba_como_data(n), n) for n in todas_abas if parsear_aba_como_data(n)]
    if not abas_com_data: raise ValueError("nenhuma aba com formato de periodo encontrada.")
    abas_com_data.sort(key=lambda x: x[0])
    recentes = abas_com_data[-quantidade:][::-1]
    return [{'nome': n, 'ano': a, 'mes': m, 'label': f"{meses_pt.get(m, m)} / {a} — aba: '{n}'"} for (a, m), n in recentes]

def _copiar_cor(color):
    # utilitario para clonar objetos de cor entre celulas de diferentes planilhas
    from openpyxl.styles.colors import Color as OxlColor
    try:
        rgb = color.rgb
        if rgb and str(rgb).upper() not in ('00000000', 'NONE', 'FFFFFFFF'): return OxlColor(rgb=rgb)
    except Exception: pass
    try:
        tema = color.theme
        tint = getattr(color, 'tint', 0.0) or 0.0
        if tema is not None: return OxlColor(theme=tema, tint=tint)
    except Exception: pass
    return copy(color)

def extrair_aba_para_temp(caminho_origem, nome_aba, pasta_temp):
    # cria uma copia temporaria de uma aba especifica para evitar locks no arquivo original
    import zipfile, xml.etree.ElementTree as ET
    wb_origem = load_workbook(caminho_origem, read_only=True, data_only=True)
    ws_origem = wb_origem[nome_aba]
    wb_novo = Workbook()
    ws_novo = wb_novo.active
    ws_novo.title = nome_aba

    for row in ws_origem.iter_rows(values_only=False):
        for cell in row:
            if cell.value is None: continue
            nova_celula = ws_novo.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.font and cell.font.color:
                try: nova_celula.font = Font(color=_copiar_cor(cell.font.color))
                except Exception: pass
    wb_origem.close()
    caminho_aba = os.path.join(pasta_temp, f"aba_{nome_aba.replace(' ', '_')}.xlsx")
    wb_novo.save(caminho_aba)
    return caminho_aba

def carregar_codigos_fiscais(caminho_planilha):
    # carrega o mapeamento de codigos fiscais e regras de negocios da aba 'codigos'
    wb = load_workbook(caminho_planilha, data_only=True, read_only=True)
    if 'CODIGOS' not in wb.sheetnames:
        wb.close()
        return {}
    aba = wb['CODIGOS']
    mapa = {}
    
    for linha in range(2, aba.max_row + 1):
        f_id = str(aba.cell(row=linha, column=1).value).strip()
        if f_id == 'None': continue
        
        mapa.setdefault(f_id, []).append({
            'codigo': str(aba.cell(row=linha, column=2).value).strip(),
            'nomenclatura': str(aba.cell(row=linha, column=3).value).strip(),
            'cnpj': str(aba.cell(row=linha, column=4).value).strip(),
            'destino': str(aba.cell(row=linha, column=5).value).strip().upper(),
            'linhas': [int(x) for x in str(aba.cell(row=linha, column=6).value).split('/') if x.isdigit()],
            'estilo_d': aba.cell(row=linha, column=7).value
        })
    wb.close()
    return mapa

def extrair_codigo_fiscal(texto_pdf, regras_filial, cnpj_ie_encontrado=None):
    # tenta encontrar o codigo fiscal no pdf baseando-se em busca exata ou proximidade
    if not regras_filial: return None
    texto_norm = texto_pdf.replace('\xa0', ' ').replace('\u200b', '')
    candidatas_exatas = []
    
    for regra in regras_filial:
        codigo_raw = regra['codigo'].strip()
        padrao_exato = r'(?<![0-9])' + re.escape(codigo_raw) + r'(?![0-9])'
        if re.search(padrao_exato, texto_norm, re.IGNORECASE):
            candidatas_exatas.append(regra)
            
    if len(candidatas_exatas) == 1: return candidatas_exatas[0]
    return candidatas_exatas[0] if candidatas_exatas else None

def carregar_cadastro(caminho_planilha):
    # mapeia cnpjs e ies para ids internos de filiais a partir da aba 'cadastro'
    wb = load_workbook(caminho_planilha, data_only=True, read_only=True)
    aba = wb['CADASTRO']
    mapa = {}
    for linha in range(2, aba.max_row + 1):
        f_id = str(aba.cell(row=linha, column=1).value).strip()
        cnpj = limpar_numero(aba.cell(row=linha, column=4).value)
        ie = limpar_numero(aba.cell(row=linha, column=5).value)
        if cnpj: mapa[cnpj] = f_id
        if ie: mapa[ie] = f_id
    wb.close()
    return mapa

def mapear_colunas_filiais(planilha, linha_cabecalho=9, max_col=100):
    # identifica em qual coluna da planilha de apuracao cada filial esta localizada
    colunas = {}
    for c in range(1, max_col + 1):
        val = planilha.cell(row=linha_cabecalho, column=c).value
        if val: colunas[str(val).strip()] = c
    return colunas

def conferir(caminho_planilha_guias, caminho_apuracao, pasta_pdfs, aba_selecionada, margem=1.00, callback_progresso=None, callback_log=None):
    # executa o fluxo principal de conferencia entre pdfs fiscais e planilha de apuracao
    def log(tipo, msg):
        if callback_log: callback_log(tipo, msg)
    
    pasta_temp = tempfile.mkdtemp(prefix="robo_fiscal_")
    try:
        log('info', f"iniciando processamento na aba '{aba_selecionada}'...")
        caminho_aba = extrair_aba_para_temp(caminho_apuracao, aba_selecionada, pasta_temp)
        wb_ap = load_workbook(caminho_aba, data_only=True)
        planilha_ap = wb_ap.active
        
        mapa_documentos = carregar_cadastro(caminho_planilha_guias)
        colunas_filiais = mapear_colunas_filiais(planilha_ap)
        
        pdfs = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith('.pdf')]
        log('info', f"analisando {len(pdfs)} arquivos...")

        # o processamento segue a logica de extracao de texto e comparacao de valores
        # os dados sensiveis de filiais foram substituidos por 'filial_exemplo'
        # nomes de departamentos foram anonimizados para 'seu_departamento'
        
        return {"status": "concluido", "processados": len(pdfs)}
    finally:
        shutil.rmtree(pasta_temp, ignore_errors=True)