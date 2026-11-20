import json

import morph_kgc
import pytest
from openpyxl import Workbook
from rdflib import Dataset, URIRef

PREFIX = '''@prefix rr: <http://www.w3.org/ns/r2rml#> .
@prefix rml: <http://semweb.mmlab.be/ns/rml#> .
@prefix ql: <http://semweb.mmlab.be/ns/ql#> .
@prefix ex: <https://example.test/> .
'''


def mapping(tmp_path, name, source, body):
    path = tmp_path / (name + '.ttl')
    path.write_text(PREFIX + f'''ex:{name} a rr:TriplesMap;
 rml:logicalSource [ rml:source {json.dumps(str(source))}; rml:referenceFormulation ql:CSV ];
 {body} .''')
    return path


def config(*sections):
    return '[CONFIGURATION]\nnumber_of_processes=1\noutput_format=N-QUADS\nlogging_level=ERROR\n' + '\n'.join(sections)


def simple(tmp_path, source, options=''):
    path = mapping(tmp_path, 'Loans', source, '''rr:subjectMap [ rr:template "https://example.test/loan/{id}"; rr:graph ex:source ];
 rr:predicateObjectMap [ rr:predicate ex:amount; rr:objectMap [ rml:reference "amount"; rr:datatype <http://www.w3.org/2001/XMLSchema#decimal> ] ]''')
    return config(f'[Loans]\nmappings={path}\n{options}')


def workbook(tmp_path):
    book = Workbook()
    book.active.title = 'Cover'
    book.active.append(['not a table'])
    loans = book.create_sheet('Loans')
    loans.append(['Report'])
    loans.append(['id', 'borrower', 'amount'])
    loans.append(['001', 'B01', '123.4500'])
    loans.append(['001', 'B01', '123.4500'])
    loans.append(['002', 'missing', '0.00'])
    borrowers = book.create_sheet('Borrowers')
    borrowers.append(['borrower', 'name'])
    borrowers.append(['B01', 'Synthetic'])
    borrowers.append(['B01', 'Synthetic'])
    path = tmp_path / 'source.xlsx'
    book.save(path)
    return path


def joined(tmp_path):
    source = workbook(tmp_path)
    loans = mapping(tmp_path, 'Loans', source, '''rr:subjectMap [ rr:template "https://example.test/loan/{id}"; rr:graph ex:source ];
 rr:predicateObjectMap [ rr:predicate ex:borrower; rr:objectMap [ rr:parentTriplesMap ex:Borrowers;
 rr:joinCondition [ rr:child "borrower"; rr:parent "borrower" ] ] ]''')
    borrowers = mapping(tmp_path, 'Borrowers', source, '''rr:subjectMap [ rr:template "https://example.test/person/{borrower}" ];
 rr:predicateObjectMap [ rr:predicate ex:name; rr:objectMap [ rml:reference "name" ] ]''')
    return config(f'[Loans]\nmappings={loans}\nexcel_sheet="Loans"\nexcel_header=1',
                  f'[Borrowers]\nmappings={borrowers}\nexcel_sheet="Borrowers"')


def normalized(statements):
    return {s.rstrip() for s in statements}


def test_sheet_header_named_graph_and_duplicate_receipts(tmp_path):
    cfg = simple(tmp_path, workbook(tmp_path), 'excel_sheet="Loans"\nexcel_header=1')
    result = morph_kgc.materialize_with_provenance(cfg)
    assert result.statements == normalized(morph_kgc.materialize_set(cfg))
    assert len(result.statements) == 2
    assert len(result.receipts) == 3
    assert any('"123.4500"^^' in s for s in result.statements)
    receipt = next(r for r in result.receipts if 'loan/001>' in r['statement'])
    assert {(c['sheet'], c['row'], c['column']) for c in receipt['cells']} == {('Loans', 3, 1), ('Loans', 3, 3)}
    graph = Dataset()
    graph.parse(data=' .\n'.join(sorted(result.statements)) + ' .', format='nquads')
    assert len(graph.graph(URIRef('https://example.test/source'))) == 2


def test_join_keeps_both_sides_and_every_duplicate_support(tmp_path):
    cfg = joined(tmp_path)
    result = morph_kgc.materialize_with_provenance(cfg)
    assert result.statements == normalized(morph_kgc.materialize_set(cfg))
    links = [r for r in result.receipts if '<https://example.test/borrower>' in r['statement']]
    assert len(links) == 4  # Two child rows x two parent rows; only one RDF assertion.
    assert len({r['statement'] for r in links}) == 1
    assert {(next(c['row'] for c in r['cells'] if c['sheet'] == 'Loans'),
             next(c['row'] for c in r['cells'] if c['sheet'] == 'Borrowers')) for r in links} == {(3,2),(3,3),(4,2),(4,3)}
    assert all({(c['sheet'], c['column']) for c in r['cells']} == {('Loans',1),('Loans',2),('Borrowers',1)} for r in links)
    assert not any('loan/002' in s for s in result.statements)
    assert result == morph_kgc.materialize_with_provenance(cfg)


def test_default_excel_behavior_and_index_selection(tmp_path):
    book = Workbook()
    book.active.append(['id', 'amount'])
    book.active.append(['001', '123.4500'])
    source = tmp_path / 'simple.xlsx'
    book.save(source)
    assert normalized(morph_kgc.materialize_set(simple(tmp_path, source))) == {
        '<https://example.test/loan/001> <https://example.test/amount> "123.4500"^^<http://www.w3.org/2001/XMLSchema#decimal> <https://example.test/source>'}
    cfg = simple(tmp_path, workbook(tmp_path), 'excel_sheet=1\nexcel_header=1')
    assert len(morph_kgc.materialize_set(cfg)) == 2


def test_xls_with_auto_engine(tmp_path):
    import xlwt
    book = xlwt.Workbook()
    sheet = book.add_sheet('Loans')
    for row, values in enumerate([['id','amount'],['001','123.4500']]):
        for column, value in enumerate(values):
            sheet.write(row, column, value)
    source = tmp_path / 'legacy.xls'
    book.save(str(source))
    result = morph_kgc.materialize_with_provenance(simple(tmp_path, source, 'excel_engine=auto'))
    assert len(result.statements) == 1
    assert {c['row'] for c in result.receipts[0]['cells']} == {2}


def test_csv_record_coordinates_and_null_filtering(tmp_path):
    source = tmp_path / 'source.csv'
    source.write_text('id,amount\n001,123.4500\n001,123.4500\n002,\n003,nan\n')
    cfg = simple(tmp_path, source)
    result = morph_kgc.materialize_with_provenance(cfg)
    assert result.statements == normalized(morph_kgc.materialize_set(cfg))
    assert len(result.statements) == 1
    assert len(result.receipts) == 2
    assert {c['record'] for r in result.receipts for c in r['cells']} == {1,2}
    assert all('row' not in c for r in result.receipts for c in r['cells'])


@pytest.mark.parametrize('options', ['excel_sheet=-1','excel_sheet=true','excel_sheet=[]',
                                     'excel_header=-1','excel_header=null','excel_engine=bogus'])
def test_invalid_reader_options_fail(tmp_path, options):
    with pytest.raises(ValueError):
        morph_kgc.materialize_set(simple(tmp_path, workbook(tmp_path), options))


def test_duplicate_excel_headers_fail_instead_of_inventing_coordinates(tmp_path):
    book = Workbook()
    book.active.append(['id','amount','amount'])
    book.active.append(['001','1','2'])
    source = tmp_path / 'duplicate.xlsx'
    book.save(source)
    with pytest.raises(ValueError, match='unique'):
        morph_kgc.materialize_with_provenance(simple(tmp_path, source))


def test_csv_multiline_uses_record_ordinal_not_line_number(tmp_path):
    source = tmp_path / 'source.csv'
    source.write_text('id,amount,unused\n001,1.00,"two\nlines"\n002,2.00,next\n')
    result = morph_kgc.materialize_with_provenance(simple(tmp_path, source))
    second = next(r for r in result.receipts if 'loan/002>' in r['statement'])
    assert {c['record'] for c in second['cells']} == {2}


def test_missing_formula_cache_retains_upstream_behavior(tmp_path):
    book = Workbook()
    book.active.append(['id','amount'])
    book.active.append(['001','=1+1'])
    source = tmp_path / 'formula.xlsx'
    book.save(source)
    result = morph_kgc.materialize_with_provenance(simple(tmp_path, source))
    assert result.statements == set()
    assert result.receipts == []


def test_other_sources_fail_closed(tmp_path):
    source = tmp_path / 'source.json'
    source.write_text('[{"id":"001","amount":"1"}]')
    with pytest.raises(ValueError, match='only CSV/TSV and Excel'):
        morph_kgc.materialize_with_provenance(simple(tmp_path, source))
