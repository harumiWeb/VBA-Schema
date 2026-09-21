# Runtime benchmark contract

`task benchmark` is the M7 runtime performance gate for Windows 64-bit Excel.
It is intentionally separate from `xlflow metrics`, which measures static
complexity only.

## Fixtures

| ID | Workload | Required result |
| --- | --- | --- |
| `object_1000_fields_success` | 1,000 text fields, valid dictionary, 10 validations per measured sample | `issue_count = 0`, median `< 500 ms` |
| `object_1000_fields_issues` | 1,000 text fields, every tenth value invalid, 10 validations per measured sample | `issue_count = 5,000` across five runs; baseline comparison only |
| `scalar_10000_success` | 10,000 whole-number validations | `issue_count = 0`, median `< 1,000 ms` |
| `scalar_10000_issues` | 10,000 non-integral values | `issue_count = 50,000` across five runs; baseline comparison only |

Each fixture performs one warm-up followed by five measured iterations. The
object fixtures repeat the same validation ten times inside each measured
sample so the short object workload is not dominated by `Timer` quantization.
The report stores all raw samples and their median, together with the actual
validation count, issue count, and RegExp creation count (`0` for these
fixtures).

## Report and baseline

- Current tracked baseline: `benchmarks/windows-x64-baseline.json`.
- Local timestamped reports: `artifacts/benchmarks/<timestamp>-windows-x64.json`.
- A success target must pass absolutely. Every fixture must also remain within
  25% of the corresponding baseline median when a baseline is available.
- Baseline updates use `tools/run-benchmark.ps1 -UpdateBaseline` only after a
  reviewer confirms the fixture contract. If fixture IDs, iteration count, or
  counters change, update this spec/ADR first; the explicit update is then
  allowed to replace the old baseline while absolute targets still gate.

## Environment and safety

The report records machine, Office bitness, Excel version, and the exact
command. The task starts only a clean managed session, refuses to attach to a
user-owned or busy workbook, and discards only the session it owns. macOS and
32-bit Excel are not covered by the baseline and remain unverified.
