# xlflow issue: changed-only push can skip a fresh managed session

## Observed behavior

`xlflow push --fast --session --no-save --json` may report `source state unchanged; skipped workbook import` immediately after starting a new managed session. The push state cache is shared with the project, but the newly started session can still contain an older workbook. A subsequent `xlflow run` then reports that a source macro is unavailable even though the source file exists.

## Reproduction

1. Push a source tree containing a new standard module and let xlflow write `.xlflow/state/push.json`.
2. Stop the managed session without saving.
3. Start a fresh managed session.
4. Run `xlflow push --fast --session --no-save --json`.
5. Observe `source state unchanged; skipped workbook import` and run the new macro.

## Workaround used here

The benchmark-owned session uses `xlflow push --session --no-save --json` without `--fast`, which forces the source import. The benchmark task refuses to attach to an existing session, so the import and discard lifecycle remain owned by the task. The release smoke probe uses the same non-fast push because it is also a fresh workbook.

## Evidence

- `tools/run-benchmark.ps1`
- `src/modules/Benchmarks/ValidationBenchmarks.bas`
- `tools/release-smoke.ps1`
