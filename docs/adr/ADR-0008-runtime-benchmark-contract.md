# ADR-0008: Runtime benchmark contract and baseline policy

## Status

`accepted`

## Background

Static procedure metrics describe code shape but do not establish runtime
performance. M7 needs a repeatable Windows/Excel measurement that exercises the
same validation workload before and after a change, while keeping benchmark
helpers out of the three-file VBE import payload.

## Decision

- Runtime measurements are implemented by
  `src/modules/Benchmarks/ValidationBenchmarks.bas` and are invoked through
  `task benchmark`.
- The benchmark owns a fresh managed xlflow session. It must refuse an active,
  busy, recovery-required, or user-owned session and must stop/discard only the
  session it started.
- The fixture set is stable: 1,000-field object success and issue workloads,
  plus 10,000 scalar success and issue workloads. Each fixture has one warm-up
  and five measured iterations; the median and raw samples are recorded. The
  object workload runs ten validations inside each measured sample so its
  short elapsed time is not dominated by `Timer` quantization; counters report
  those actual validations.
- Success fixtures have absolute release targets: object success `< 500 ms` and
  scalar success `< 1,000 ms`. Issue fixtures are measured for regression
  visibility but have no absolute target because issue construction is
  intentionally more expensive.
- Every report records validation count, issue count, RegExp creation count,
  machine, Office bitness, Excel version, command, fixture IDs, and the target
  values. A tracked Windows x64 baseline lives at
  `benchmarks/windows-x64-baseline.json`; a new result may not regress a
  baseline median by more than 25 percent.
- Timestamped reports under `artifacts/benchmarks/` are local evidence and are
  ignored by Git. Benchmark modules are excluded from build and static metric
  collection, so they cannot enter the release payload or distort production
  complexity metrics.
- Windows 64-bit Excel is the only measured environment. macOS and 32-bit
  Excel remain source-compatible by design but unverified and are not implied
  by this baseline.

## Consequences

- Performance decisions can compare deterministic fixture and counter sets,
  not only wall-clock numbers.
- The benchmark requires a local Windows Excel installation and does not run in
  Excel-free CI. CI verifies the source and release payload separately.
- Issue-heavy validation remains visible without making a noisy absolute SLA
  for diagnostics allocation.
- A baseline update is a release decision: it must use the same fixture and
  iteration contract, or first update this ADR/spec when that contract changes,
  and it must be documented in `tasks/todo.md`. `-UpdateBaseline` is the only
  explicit escape hatch; absolute targets still gate the update.

## Rationale

- Evidence: `src/modules/Benchmarks/ValidationBenchmarks.bas` defines the
  fixture IDs, counters, warm-up, and five samples.
- Evidence: `tools/run-benchmark.ps1` owns session safety, target checks,
  baseline comparison, and report output.
- Evidence: `benchmarks/windows-x64-baseline.json` is the current Windows x64
  baseline, and `Taskfile.yml` exposes the reproducible command.
- Related docs: `docs/specs/benchmark-contract.md`, `docs/adr/ADR-0007-release-and-ci-boundary.md`.

## Supersedes

- None

## Superseded by

- None
