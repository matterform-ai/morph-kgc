# Tabular ingestion spike

Based on upstream tag `2.10.0`, commit
`203b1ff717c9ebb280b87e50bea55577d5cd386c`. Apache-2.0; see LICENSE.
Changes to upstream files add Excel configuration, an experimental provenance API,
and prevent self-join elimination across separately configured sources.

## Excel reader options

Each data source section can set:

```ini
[Loans]
mappings=loans.ttl
excel_sheet="Loans"
excel_header=1
excel_engine=auto
```

`excel_sheet` is a JSON string (sheet name) or nonnegative integer (zero-based
index). `excel_header` is a zero-based row index. Defaults preserve the upstream
first sheet / first header / openpyxl behavior. Engines: openpyxl, xlrd, auto.
Install `[excel]` for XLS/XLSX support. Each sheet uses its own mapping file and
configuration section, even when sheets belong to one workbook.

## Experimental provenance API

```python
from morph_kgc import materialize_with_provenance
result = materialize_with_provenance("config.ini")
# result.statements: set of serialized statements without the terminating dot
# result.receipts: list of {statement, triples_map, cells}
```

Runs sequentially. Preserves each contributing source row through flat parent
joins and emits receipts before RDF deduplication. Each cell has source,
source_name, reference, and either sheet/row/column (Excel, 1-based coordinates)
or record (CSV/TSV, 1-based parsed data-record ordinal, not a physical line).
Join receipts include both sides' join keys and term-map references. Cell values
are omitted, but output statements may contain confidential data: protect receipts
like the input dataset. Coordinates are relative to the input bytes; the caller
must bind those bytes and mappings to hashes for durable replay.

This is conservative reference lineage, not minimal causal provenance. It does
not explain missing/rejected rows, compute formula dependencies, retain formula
text, validate datatypes, or preserve display formats. Formula cells use cached
values as in upstream. No change to existing materialize APIs except reader
options and the cross-source join fix.

Only CSV/TSV and Excel file sources, ordinary term maps, and flat parent joins are
supported by the provenance API. Functions, SQL views, other source kinds,
quoted triples and all-constant mappings are rejected. Excel columns referenced
by mappings must have unique literal headers. Provenance may expand memory and
join cardinality substantially; use this API for bounded experiments.

## Tests

```sh
uv venv --python 3.12
uv pip install -e '.[test,excel]' xlwt
uv run --no-project python -m pytest test/fork_spike -q
```

`xlwt` only writes synthetic legacy XLS fixtures. Tests cover default reader
behavior, sheet names/indices, header offsets, XLS, exact lexical strings, named
graphs, null filtering, duplicate support, cross-sheet joins and invalid options.

Keep changes on a branch based on the pinned release. Rebase/cherry-pick patches
onto upstream main and run upstream tests before proposing separate reader and
provenance PRs. Do not submit Matterform mappings or customer data upstream.
