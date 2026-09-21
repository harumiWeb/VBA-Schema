# Guide for AI Agents

## 0. Project Overview

### VBA-Schema

A lightweight schema validation library for VBA.

This project is a declarative runtime validation library for VBA inspired by TypeScript's Zod.

### Project Structure

Below is the structure of this project.
When making changes to the project structure, workers must update the following tree to the latest version sequentially so that subsequent AI agents can understand them.

```txt
root: .
├── .github/
│   └── workflows/
│       ├── release.yml
│       └── source-check.yml
├── build/
│   └── Book.xlsm (tracked development workbook)
├── dist/
│   └── VBA-Schema/ (generated, ignored)
├── benchmarks/
│   └── windows-x64-baseline.json
├── docs/
│   ├── adr/
│   │   ├── ADR-0008-runtime-benchmark-contract.md
│   │   ├── ADR-0005-literal-enum-pattern-email-boundary.md
│   │   ├── ADR-0006-union-branch-ownership-and-error-boundary.md
│   │   ├── ADR-0007-release-and-ci-boundary.md
│   │   ├── ADR-0001-small-distribution-and-portable-core.md
│   │   ├── ADR-0002-vba-safe-api-and-error-boundary.md
│   │   ├── ADR-0003-object-dictionary-and-schema-cycle-boundary.md
│   │   └── ADR-0004-array-collection-sequence-boundary.md
│   ├── specs/
│   │   ├── benchmark-contract.md
│   │   └── v1-contract.md
│   ├── xlflow-issues/
│   │   ├── 20260921-fmt-parser-recovery.md
│   │   └── 20260921-push-state-cache-fresh-session.md
│   ├── release-checklist.md
│   └── design.md
├── src/
│   ├── classes/
│   │   ├── VSchema.cls
│   │   └── VValidationResult.cls
│   ├── modules/
│   │   ├── Benchmarks/
│   │   │   └── ValidationBenchmarks.bas
│   │   ├── Tests/
│   │   │   ├── PublicApiCompile.bas
│   │   │   ├── TestLiteral.bas
│   │   │   ├── TestPattern.bas
│   │   │   ├── TestUnion.bas
│   │   │   ├── SampleTests.bas
│   │   │   ├── TestArray.bas
│   │   │   ├── TestAnyValue.bas
│   │   │   ├── TestBool.bas
│   │   │   ├── TestDateTime.bas
│   │   │   ├── TestErrors.bas
│   │   │   ├── TestNullable.bas
│   │   │   ├── TestNumber.bas
│   │   │   ├── TestResult.bas
│   │   │   ├── TestObject.bas
│   │   │   └── TestText.bas
│   │   ├── Xlflow/
│   │   │   ├── XlflowAssert.bas
│   │   │   ├── XlflowDebug.bas
│   │   │   ├── XlflowRuntime.bas
│   │   │   └── XlflowUI.bas
│   │   ├── App.bas
│   │   ├── Main.bas
│   │   ├── Schema.bas
│   │   └── Ui.bas
│   └── workbook/
│       ├── Sheet1.bas
│       └── ThisWorkbook.bas
├── sample/
│   ├── 01-order-import/
│   │   ├── README.md
│   │   └── SampleOrderImport.bas
│   ├── 02-settings-validation/
│   │   ├── README.md
│   │   └── SampleSettingsValidation.bas
│   ├── 03-api-payload/
│   │   ├── README.md
│   │   └── SampleApiPayload.bas
│   └── README.md
├── tools/
│   ├── benchmark-environment.ps1
│   ├── check-format.ps1
│   ├── check-production-hygiene.ps1
│   ├── run-benchmark.ps1
│   ├── release-stage.ps1
│   ├── package-release.ps1
│   ├── release-smoke.ps1
│   ├── release-verify.ps1
│   ├── test-benchmark-environment.ps1
│   ├── test-release-stage-safety.ps1
│   ├── verify-samples.ps1
│   └── provision-workbook.ps1
├── tasks/
│   ├── lessons.md
│   └── todo.md
├── AGENTS.md
├── CHANGELOG.md
├── LICENSE
├── README.md
├── Taskfile.yml
└── xlflow.toml
```

## 1. Workflow Design

### 1. Basic Approach: Work in Plan Mode First

- For tasks involving three or more steps or those affecting the overall architecture, always begin in Plan mode.
- If progress stalls at any point, do not force continuation - stop and replan instead.
- Use the Plan mode not only for implementation but also for designing verification procedures.
- As early as possible, refine specifications to reduce ambiguity.

### 2. Multi-Agent Strategy

- Make active use of subagents to avoid contaminating the main context.
- Delegate tasks such as research, verification, and parallel analysis to subagents.
- For complex problems, utilize subagents even when they require significant computational resources.
- Assign each subagent only one task to maintain focused execution.
- Use an explorer for codebase exploration (primarily reading activities).
- Use a worker for implementation and modifications.
- Use a reviewer for code reviews.

### 3. Self-Improvement Loop

- When receiving correction instructions from users, document these patterns in `tasks/lessons.md`.
- Formulate clear rules for yourself to prevent repeating the same mistakes.
- Continuously refine these rules until error rates decrease significantly.
- At the beginning of each session, review relevant lessons related to the project.

### 4. Always Verify Before Finalizing

- Do not mark tasks as complete until you can demonstrate their functionality.
- When necessary, compare your changes against the main branch for verification.
- Always ask yourself: "Would a staff engineer approve this?"
- Complete the process by running tests, reviewing logs, and demonstrating proper operation.

### 5. Maintain Balance While Pursuing More Elegant Solutions

- Before implementing major changes, pause to first consider: "Is there a more elegant way to do this?"
- If your fix feels ad hoc, reframe it as: "How can I implement this in a more refined manner based on what I know now?"
- However, do not overthink simple and obvious fixes - avoid excessive design.
- Before delivering any deliverable, thoroughly review your own implementation with a critical eye.

### 6. Handle Bug Fixes Autonomously

- When receiving bug reports, investigate them independently without waiting for instructions, then proceed directly to resolution.
- Use logs, errors, and failing tests to autonomously identify and resolve the issue.
- Avoid forcing users into unnecessary context switches.

## - Even without explicit instructions, if the CI pipeline is down, take initiative to resolve it.

## 2. Required Workflow Procedures

Before generating or modifying code, perform the following steps according to the scale of your work:

1. Understand the requirements: Review relevant specification documents, ADR documentation, and existing implementations.
2. Consider the design implications: Assess impact scope, compatibility with current designs, and alternative approaches.
3. If necessary, create working notes:
   - For recurrence prevention: `tasks/lessons.md`
4. Add or update tests as needed.
5. Implement changes.
6. Verify functionality.
7. Run tests.
8. Conduct self-review.
9. Update documentation, ADR documents, specifications, and the CHANGELOG as appropriate.

- Any updates to ADR documents or specifications must be recorded in the respective directories:
- For ADR documents: `docs/adr/`
- For specification documents: `docs/specs/`
- If changes affect public APIs, they may require recording in the following documentation:
- Specification documents within `docs/specs/`
- Overview descriptions in the `README.md` file

## - In final reports, retain the absolute paths of used `tmp_workspaces`, the execution commands, test results, and any unverified items.

## 3. Documentation Retention Policy

### Role Separation Guidelines

- The `tasks/lessons.md` should exclusively serve as a repository for recording recurrence prevention rules - it must not be used for storing design decisions or actual specifications themselves.
- Design judgments and trade-offs should be recorded in the `docs/adr/` directory, while current internal specifications and constraints should be moved to the `docs/specs/` directory.

### Distinction Between ADR Documents and Specification Documentation

- ADR documentation should capture the reasoning behind decisions and document the rationale for choosing one approach over others—information that will be valuable for future implementers facing similar challenges.
- When editing ADR documents, use the `adr-manager` skill.
- Specification documents should record: permanent rules established through review processes, continuous integration testing, and failure resolution; as well as CLI specifications, validation requirements, and compatibility agreements.
- If additional regression tests were added due to specific design considerations where forgetting the rationale could lead to recurrence, document these in the specification documentation.

### Information to Preserve

- The reasoning behind decisions that will be useful for future implementers facing similar issues
- The chosen approach after comparing multiple alternatives
- Permanent rules established through review processes, CI testing, and failure resolution
- CLI specifications, validation requirements, and compatibility agreements
- Documentation of additional regression tests where forgetting the rationale could cause recurrence due to design context

### Information to Discard

- Single-step task notes
- Aborted hypotheses or intermediate notes
- Progress logs that lose value after completion

## 4. Core Principles

- **Keep it simple first**: All changes should maintain maximum simplicity with minimal scope impact.
- **Do no harm**: Identify the root cause. Do not resort to quick fixes. Maintain professional engineering standards.
- **Minimize impact**: Only modify necessary components without introducing new bugs.

## Dogfooding xlflow

As a sub-objective, this project also serves as a dogfooding project to test whether it is possible to develop a full-scale VBA library using only xlflow and AI agents.

Therefore, any bugs on the xlflow side encountered during the development of this project will be reported as issues to the [official xlflow repository](https://github.com/harumiWeb/xlflow). As such, any problems found should be documented in `docs\xlflow-issues` as soon as they are discovered. Whenever possible, it is desirable to include conditions that allow for reproduction later.
