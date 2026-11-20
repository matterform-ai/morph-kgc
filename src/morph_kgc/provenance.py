"""Experimental, opt-in source coordinates for flat CSV/Excel RML mappings.

Coordinates describe all references read for an emission (including join keys),
not minimal causal explanations. Values are deliberately excluded from receipts.
"""
import json
from dataclasses import dataclass

from .args_parser import load_config_from_argument
from .constants import (CSV, TSV, EXCEL, RML_SOURCE, RML_EXECUTION,
                        RML_QUOTED_TRIPLES_MAP, RML_TRIPLES_MAP_CLASS, RML_CONSTANT)
from .mapping.mapping_parser import retrieve_mappings
from .materializer import _materialize_rml_rule

COLUMN = '__morph_source_coordinates__'


@dataclass
class MaterializationResult:
    statements: set
    receipts: list


def attach_source_coordinates(data, rule, references):
    if any(COLUMN in str(column) for column in data.columns):
        raise ValueError('Source uses the reserved provenance column name')
    coordinates = data.attrs.get('source_coordinates')
    records = []
    for index in range(len(data)):
        cells = []
        for reference in sorted(references):
            cell = dict(source=str(rule['logical_source_value']),
                        source_name=str(rule['source_name']), reference=reference)
            if coordinates:
                cell.update(sheet=coordinates['sheet'],
                            row=coordinates['first_row'] + index,
                            column=coordinates['columns'][reference])
            else:
                # CSV record ordinal, NOT a physical line (quoted fields may span lines).
                cell['record'] = index + 1
            cells.append(cell)
        records.append(json.dumps(cells, sort_keys=True))
    data[COLUMN] = records
    return data


def materialize_with_provenance(config):
    """Return RDF statements and one receipt per emission before RDF deduplication.

    Sequential experimental API. Rejects unsupported sources/functions/RDF-star
    rather than returning incomplete lineage. Existing materialize APIs are unchanged.
    """
    config = load_config_from_argument(config)
    config._capture_provenance = True
    config.set_number_of_processes('1')
    mappings, functions, _ = retrieve_mappings(config)
    positions = ('subject', 'predicate', 'object', 'graph', 'lang_datatype')
    for _, rule in mappings.iterrows():
        if rule['source_type'] not in [CSV, TSV, *EXCEL] or rule['logical_source_type'] != RML_SOURCE:
            raise ValueError('Provenance supports only CSV/TSV and Excel file sources')
        if any(rule[f'{p}_map_type'] in (RML_EXECUTION, RML_QUOTED_TRIPLES_MAP) for p in positions):
            raise ValueError('Provenance does not yet support functions or quoted triples')
        if all(rule[f'{p}_map_type'] == RML_CONSTANT for p in ('subject', 'predicate', 'object', 'graph')):
            raise ValueError('Source-independent constant mappings have no cell provenance')
    statements, receipts = set(), []
    for _, rule in mappings.loc[mappings['triples_map_type'] == RML_TRIPLES_MAP_CLASS].iterrows():
        data = _materialize_rml_rule(rule, mappings, functions, config)
        for _, row in data.iterrows():
            cells = json.loads(row[COLUMN])
            if 'parent_' + COLUMN in row:
                cells += json.loads(row['parent_' + COLUMN])
            statement = row['triple'].rstrip()
            statements.add(statement)
            receipts.append(dict(statement=statement, triples_map=str(rule['triples_map_id']), cells=cells))
    receipts.sort(key=lambda receipt: json.dumps(receipt, sort_keys=True))
    return MaterializationResult(statements, receipts)
