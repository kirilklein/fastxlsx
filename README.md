# fastxlsx

Fast, forward-only `.xlsx` export for Python — a thin [PyO3](https://pyo3.rs) binding
over [`rust_xlsxwriter`](https://crates.io/crates/rust_xlsxwriter)'s `constant_memory`
mode. Built for **large tabular exports** with flat memory usage, not for reading or
editing existing workbooks.

```python
import fastxlsx

wb = fastxlsx.Workbook()
ws = wb.add_sheet("Patients")
ws.append(["Patient_ID", "Score", "Eligible", "Note"])
ws.append([123, 4.5, True, "NSQIP"])
wb.save("patients.xlsx")
```

The API is deliberately append-only: forward streaming, one row at a time, no random
cell access and no mutation of existing files. Supported cell types today: `str`, `int`,
`float`, `bool`, `None` (blank).

## Benchmarks

### What we measure and why

`.xlsx` writing is a CPU- and allocation-heavy inner loop: for every cell the writer
produces a chunk of XML, registers strings/styles in lookup tables, and zip-compresses
the result. The question is whether doing that loop in Rust beats the best existing
Python writers on **both wall time and peak memory**.

We compare against the real competitors, not a strawman:

- **`xlsxwriter_constant`** — `XlsxWriter` in `constant_memory` mode. This is the fairest
  comparison: `rust_xlsxwriter` is the Rust port of `XlsxWriter`, so this isolates
  *Python vs Rust* with the format strategy held constant.
- **`openpyxl_writeonly`** — `openpyxl` in `write_only=True` (its streaming mode).
- **`xlsxwriter`** — `XlsxWriter` default (buffers the workbook in memory).
- **`pandas`** — `DataFrame.to_excel` with the default engine.

### Method

- Each engine writes the **same dataset** in its **own process**; peak RSS is taken from
  `getrusage(RUSAGE_SELF).ru_maxrss` after the write (so the ~14–15 MB Python interpreter
  floor is included in every Python row — fastxlsx's *marginal* footprint is single-digit
  MB; the pure-Rust core measures ~4 MB).
- Each profile holds ~1,000,000 cells so engines are compared on equal work.
- Reproduce with: `python benchmarks/run_matrix.py`

Machine: macOS 15.7.4, Apple Silicon (arm64), Python 3.13.5. Single run per cell;
numbers are directional, not statistically averaged.

### Results

#### mixed — 200k rows × 5 cols, mixed types (str/int/float/bool)

| Engine | Time (s) | Peak RSS (MB) |
|---|---:|---:|
| **fastxlsx** | **0.60** | **18** |
| xlsxwriter_constant | 2.53 | 25 |
| xlsxwriter | 2.96 | 193 |
| openpyxl_writeonly | 4.26 | 38 |
| pandas | 5.26 | 264 |

#### numeric — 200k rows × 5 cols, all numeric

| Engine | Time (s) | Peak RSS (MB) |
|---|---:|---:|
| **fastxlsx** | **0.57** | **18** |
| xlsxwriter_constant | 2.05 | 25 |
| xlsxwriter | 2.20 | 188 |
| openpyxl_writeonly | 3.75 | 38 |
| pandas | 4.12 | 232 |

#### strings — 200k rows × 5 cols, all unique strings (stresses the shared-string table)

| Engine | Time (s) | Peak RSS (MB) |
|---|---:|---:|
| **fastxlsx** | **0.90** | **17** |
| xlsxwriter_constant | 3.46 | 24 |
| xlsxwriter | 3.63 | 285 |
| openpyxl_writeonly | 5.01 | 39 |
| pandas | 6.47 | 345 |

#### wide — 20k rows × 50 cols, mixed types (wide rows)

| Engine | Time (s) | Peak RSS (MB) |
|---|---:|---:|
| **fastxlsx** | **0.59** | **17** |
| xlsxwriter_constant | 2.26 | 24 |
| xlsxwriter | 2.51 | 202 |
| openpyxl_writeonly | 3.56 | 38 |
| pandas | 4.54 | 285 |

### Takeaways

- **fastxlsx wins on both axes in every profile.** Against the fairest competitor
  (`xlsxwriter_constant`) it is **~3.6–4.2× faster** and uses **~1.4× less memory** —
  that gap is the pure Python-vs-Rust win, with the streaming strategy held constant.
- **Memory stays flat (~17–18 MB) regardless of dataset shape**, confirming the
  constant-memory streaming core. Non-streaming writers (`xlsxwriter` default, `pandas`)
  scale RSS with data size, reaching 190–345 MB.
- **High-cardinality strings are the hardest case for fastxlsx** (0.90 s vs 0.57 s for
  numeric) because shared-string handling dominates — still 3.8× faster than the next
  best.
- Each fastxlsx run here pays per-row FFI overhead (one `append` call per row); a planned
  bulk-column path should widen the gap further.

### Caveats

- Single run per data point; treat ratios as directional.
- All cell types are simple scalars; styled cells, dates, and formulas are not yet
  measured (and not yet supported).
- Results are from one machine/OS; absolute numbers will vary.

## Development

```sh
# Build + install into the local venv
uv run --with maturin maturin build --release --out dist
uv pip install --python .venv dist/*.whl python-calamine

# Correctness round-trip tests
.venv/bin/python -m unittest discover -s tests -v

# Benchmarks
.venv/bin/python benchmarks/run_matrix.py
```
