"""Extração determinística do perfil Recebíveis Power BI detalhado, sem rede.

Não é um extrator universal: recusa layouts não reconhecidos. O perfil usa
marcadores textuais e coordenadas das colunas, nunca valores fixos do exemplo.
"""
import re
from pathlib import Path
import pdfplumber
from extrair_recebiveis_gemini import (
    DocumentExtraction, ExtractedTable, ColumnDefinition, TableRow, UserFacingError,
)

NUM = r"[+-]?[\d.]+,\d{2}"
PCT = NUM + '%'
INT = r"[\d.]+"


def clean(text):
    return re.sub(r'\s+', ' ', re.sub(r'[\ue000-\uf8ff]', '', text)).strip()


def extract_local(pdf_path: Path, expected_pages: int) -> DocumentExtraction:
    tables = []
    warnings = [
        'Extração sem IA pelo perfil Recebíveis Power BI detalhado. Confira os valores antes do uso.',
        'Somente linhas e colunas presentes no PDF são extraídas. Totais podem incluir registros ocultos; não são recalculados.',
    ]
    with pdfplumber.open(pdf_path) as pdf:
        pages = [(page, [dict(line, text=clean(line['text'])) for line in page.extract_text_lines()]) for page in pdf.pages]
        detail = [(i, page, lines) for i, (page, lines) in enumerate(pages) if any('vlr_principal_atualizado_aberto' in x['text'] for x in lines)]
        if len(detail) != 1:
            raise UserFacingError('O modo sem IA requer o layout detalhado de Recebíveis Power BI do modelo validado. Use GPT para outro layout ou PDF digitalizado.')
        index, page, lines = detail[0]

        def section(all_lines, start, end=None, prefix=False):
            starts = [i for i,x in enumerate(all_lines) if (x['text'].startswith(start) if prefix else x['text'] == start)]
            if len(starts) != 1:
                raise UserFacingError(f'Layout não reconhecido na seção {start}. Nenhum arquivo parcial foi gerado.')
            a = starts[0] + 1
            b = next((i for i in range(a,len(all_lines)) if all_lines[i]['text'] == end),len(all_lines)) if end else len(all_lines)
            if end and b == len(all_lines):
                raise UserFacingError(f'Não foi possível delimitar a seção {start}.')
            return all_lines[a:b]

        def add(title, headers, types, rows, page_no, notes=None):
            if not rows:
                raise UserFacingError(f'Nenhuma linha reconhecida em {title}; o layout precisa de revisão.')
            tables.append(ExtractedTable(table_id=f'T{len(tables)+1}', sheet_name=title, title=title,
                source_pages=[page_no], source_note='Extração determinística de texto do PDF', limitations=notes or [],
                columns=[ColumnDefinition(name=h,data_type=t) for h,t in zip(headers,types)],
                rows=[TableRow(row_kind='total' if r[0].startswith(('Total','Totais')) else 'data',values=r) for r in rows]))

        def matched(block, pattern):
            return [list(m.groups()) for line in block if (m := re.fullmatch(pattern,line['text']))]

        pn = index + 1
        summary = section(lines,'Total Processos Ativos','Inadimplência',True)
        add('Resumo Processos',['Total Processos Ativos','Prorrogados','% Prorrogados','Reparcelados','% Reparcelados','Antecipados','% Antecipados'],
            ['integer','integer','percentage','integer','percentage','integer','percentage'],
            matched(summary,rf'({INT}) ({INT}) ({PCT}) ({INT}) ({PCT}) ({INT}) ({PCT})'),pn)
        block = section(lines,'Inadimplência','Recuperação')
        add('Inadimplência',['Inadimplência','Valor','Qtd','Observacao'],['text','percentage','percentage','text'],
            matched(block,rf'(Inadimplência .+?) ({PCT}) ({PCT}) (.+)'),pn)

        # Posições horizontais preservam percentuais vazios, sem deslocar colunas.
        block = section(lines,'Recuperação','Repasse')
        data = [x for x in block if re.match(r'^(\d{4} \d{1,2}|Total) ',x['text'])]
        all_words = page.extract_words()
        def words_for(line):
            return [w for w in all_words if abs(w['top']-line['top']) < 3]
        anchor_row = next((x for x in data if len(re.findall(PCT,x['text'])) == 6),None)
        if not anchor_row:
            raise UserFacingError('Recuperação não possui linha de referência completa para alinhar as colunas.')
        anchors = [w['x1'] for w in words_for(anchor_row) if re.fullmatch(PCT,w['text'])]
        recovered = []
        for line in data:
            prefix = re.match(r'^(\d{4}) (\d{1,2})',line['text'])
            row = [prefix[1],prefix[2]] if prefix else ['Total','']
            cells = ['']*6
            for w in words_for(line):
                if re.fullmatch(PCT,w['text']):
                    col = min(range(6),key=lambda k:abs(anchors[k]-w['x1']))
                    if cells[col] or abs(anchors[col]-w['x1'])>8:
                        raise UserFacingError('Alinhamento ambíguo em Recuperação. Use GPT ou revise o layout.')
                    cells[col] = w['text']
            recovered.append(row+cells)
        add('Recuperação',['Ano','Mes_Numero','% Realizado Valores Total','% Realizado Valores No Mês','% Realizado Valores Até Mês Seguinte','% Realizado Parcelas','% Realizado Parcelas No Mês','% Realizado Parcelas Até Mês Seguinte'],['integer','integer']+['percentage']*6,recovered,pn)
        add('Repasse',['Periodo','R$'],['text','currency'],matched(section(lines,'Repasse','Recebíveis'),rf'([A-Za-zçÇ]{{3}}/\d{{4}}|Total) ({NUM})'),pn)
        add('Recebíveis',['Descricao','Recebíveis','Média','Mínimo'],['text']+['currency']*3,
            matched(section(lines,'Recebíveis','Resumo 1'),rf'(De \d{{2}}/\d{{2}}/\d{{4}} até \d{{2}}/\d{{2}}/\d{{4}}) ({NUM}) ({NUM}) ({NUM})'),pn)

        for title,end,first in [('Resumo 1','Resumo 2','Periodo'),('Resumo 2','Valores a Vencer/Vencidos','Período'),('Valores a Vencer/Vencidos','Agregado Mensal','Ano')]:
            block = section(lines,title,end)
            header = next((x for x in block if 'Cobrança Administrativa' in x['text'] and 'Juridico' in x['text']),None)
            if not header:
                raise UserFacingError(f'Cabeçalho incompatível em {title}.')
            hw = words_for(header)
            right_edges = [(next(w['x0'] for w in hw if w['text']==left)+next(w['x1'] for w in hw if w['text']==right))/2 for left,right in [('Cobrança','Administrativa'),('Juridico','Juridico'),('Total','Total')]]
            rows = []
            for line in block:
                if not re.search(NUM,line['text']):
                    continue
                ws = words_for(line)
                amounts = [w for w in ws if re.fullmatch(NUM,w['text'])]
                if not amounts:
                    continue
                label = clean(' '.join(w['text'] for w in ws if w['x0'] < amounts[0]['x0']-1))
                vals = ['']*3
                for w in amounts:
                    center = (w['x0']+w['x1'])/2
                    col = min(range(3),key=lambda k:abs(right_edges[k]-center))
                    if vals[col] or abs(right_edges[col]-center)>25:
                        raise UserFacingError(f'Colunas ambíguas em {title}.')
                    vals[col]=w['text']
                rows.append([label]+vals)
            add(title,[first,'Cobrança Administrativa','Juridico','Total'],['text']+['currency']*3,rows,pn)
        block = section(lines,'Agregado Mensal')
        rows = matched(block,rf'(\d{{4}}) (\d{{1,2}}) ({NUM}) (Sim|Não)')
        rows += [[m[0],'',m[1],''] for m in matched(block,rf'(Total) ({NUM})')]
        add('Agregado Mensal',['Ano','Mes_Numero','vlr_principal_atualizado_aberto','a vencer'],['integer','integer','currency','text'],rows,pn)

        covered = {index}
        for i, (_,ls) in enumerate(pages):
            text = '\n'.join(x['text'] for x in ls)
            if 'Solucao_Contrato ID_Contrato Numero_Contrato' in text:
                block=section(ls,'Contratos')
                rows=matched(block,rf'(.+?) (GRADUAÇÃO|PÓS-GRADUAÇÃO|TÉCNICO) (\S+) (\S+) (\d+) (\d+) ({NUM}) (\S+)')
                add('Contratos',['Sigla Convênio','Grau','Modalidade','Solucao_Contrato','ID_Contrato','Numero_Contrato','Taxa_Convenio_Contrato','Situacao_Contrato'],['text']*6+['decimal','text'],rows,i+1,['Colunas cortadas à direita (como Ano_Con) não foram inferidas.'])
                covered.add(i)
            elif 'N° Processo Fundacred ID Contrato' in text:
                rows=matched(section(ls,'Parcelas'),r'(\d+) (.+?) (\S+) (PRESENCIAL|EAD|HÍBRIDO) (GRADUAÇÃO|PÓS-GRADUAÇÃO|TÉCNICO) (\d+) (\d+) (\d+) (\d+) (\d{2}/\d{2}/\d{4})')
                add('Parcelas',['Código Convênio','Nome Amigável','Solução Contrato','Modalidade','Grau','N° Processo','N° Processo Fundacred','ID Contrato','N° Contrato','Data Efetivação'],['text']*9+['date'],rows,i+1,['Somente as dez colunas completas visíveis; colunas cortadas e linhas fora da rolagem não foram inferidas.'])
                covered.add(i)
            elif 'Detalhamento Estudantes' in text:
                summary=section(ls,'Quantidade de Estudantes','Detalhamento Estudantes',True)
                add('Resumo Exposição',['Quantidade de Estudantes','Quantidade de Contratos','Quantidade de Parcelas','Valor Total em Aberto','% Maior exposição'],['integer']*3+['currency','percentage'],matched(summary,rf'({INT}) ({INT}) ({INT}) R\$ ({NUM}) ({PCT})'),i+1)
                add('Detalhamento Estudantes',['ID Beneficiário','Contratos','Parcelas','Valor em Aberto','% Exposição'],['text','integer','integer','currency','percentage'],matched(section(ls,'Detalhamento Estudantes'),rf'(\S+) ({INT}) ({INT}) R\$ ({NUM}) ({PCT})'),i+1,['O total do relatório pode abranger estudantes não visíveis nesta página.'])
                covered.add(i)
        for i,(_,ls) in enumerate(pages):
            if i not in covered:
                warnings.append(f'Página {i+1} não extraída no perfil detalhado; pode conter resumo duplicado ou conteúdo adicional. Revise a página separadamente.')
        warnings.extend(f'{t.title}: {note}' for t in tables for note in t.limitations)
    return DocumentExtraction(source_file=pdf_path.name,page_count=expected_pages,warnings=warnings,tables=tables)
