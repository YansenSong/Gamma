# CLAUDE.md · Practical Coding-Agent Guidelines

A compact set of behavioral rules for coding agents such as Claude Code, Codex, and other repository-aware assistants.

The goal is simple: **reduce unnecessary changes, hidden assumptions, over-engineering, and unverified “looks good” implementations.**

These rules deliberately bias toward careful, reviewable engineering rather than maximum speed. For trivial tasks, use judgment.

## How to use this repository

This repository is intentionally tiny. Treat the document below as a reusable baseline and merge it with project-specific instructions.

Typical options:

```text
Project root
├── CLAUDE.md      # Claude Code instructions
├── AGENTS.md      # Generic coding-agent instructions
└── ...
```

You can copy or adapt these rules into either file, then add repository-specific details such as:

- build and test commands;
- architecture boundaries;
- generated files that must not be edited;
- formatting / linting conventions;
- deployment constraints;
- files containing secrets or local configuration;
- project-specific verification steps.

A good project instruction file should describe **how to work safely in this repository**, not attempt to explain every line of the codebase.

## Recommended project-specific section

After the generic rules, add a short block like this:

```markdown
## Project Notes

- Python 3.12, managed with uv.
- Run tests with `uv run pytest`.
- Run lint with `uv run ruff check .`.
- Do not edit generated files under `dist/`.
- API entry point: `src/app/main.py`.
- Database migrations live under `migrations/`.
- Before finishing, run the smallest relevant tests first, then the full suite when practical.
```

This keeps stable agent behavior separate from fast-changing project facts.

---

# Guidelines

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:

- State assumptions explicitly when they materially affect the solution.
- If multiple interpretations exist, make the ambiguity visible instead of silently choosing one.
- If a simpler approach solves the problem, prefer it.
- If something important is unclear, identify exactly what is unclear before making broad changes.
- Read the relevant code before proposing an architecture that may already exist.

The objective is not to produce a long plan for every change. It is to avoid confidently implementing the wrong thing.

## 2. Simplicity First

**Write the minimum code that solves the requested problem. Nothing speculative.**

- Do not add features that were not requested.
- Do not create abstractions for one-off behavior without a real need.
- Do not add “flexibility” or configuration solely because it might be useful someday.
- Do not add defensive branches for impossible states unless the codebase genuinely requires them.
- Prefer an obvious 50-line solution over a clever 200-line framework.

Ask:

> Would a senior engineer reviewing this diff say it became more complicated than the problem required?

If yes, simplify it.

## 3. Surgical Changes

**Touch only what the task requires. Clean up only the mess your change creates.**

When editing existing code:

- Do not reformat unrelated files.
- Do not rename adjacent functions merely because you prefer another style.
- Do not refactor unrelated code while fixing a focused bug.
- Match the repository's established patterns unless changing them is part of the task.
- If you notice unrelated dead code or technical debt, mention it separately rather than silently expanding scope.

When your own changes create unused imports, variables, functions, or files, clean those up.

A useful review question is:

> Can every changed line be traced back to the requested outcome?

If not, reconsider the diff.

## 4. Goal-Driven Execution

**Turn vague work into verifiable outcomes, then loop until those outcomes are checked.**

Examples:

```text
“Add validation”
→ define invalid cases
→ add or identify tests for them
→ implement validation
→ verify the tests pass
```

```text
“Fix the bug”
→ reproduce the failure
→ identify the smallest root cause
→ make the change
→ verify the original failure no longer occurs
```

```text
“Refactor X”
→ establish current behavior
→ make the structural change
→ verify behavior remains unchanged
```

For larger tasks, a brief plan is enough:

```text
1. Inspect the relevant path → verify the real entry points and constraints.
2. Make the smallest implementation → verify focused tests.
3. Review the diff → run broader checks when practical.
```

Strong success criteria let an agent work independently without turning every small decision into another question.

## 5. Verify Before Declaring Success

A change is not complete just because the code looks plausible.

Before saying a task is done, use the strongest practical verification available:

- unit or integration tests;
- type checking;
- linting;
- compilation / build;
- a reproducible manual check;
- inspection of the final diff.

If verification cannot be run, say exactly what was and was not verified.

Avoid vague completion statements such as:

```text
“This should work.”
```

Prefer concrete evidence:

```text
“Updated the parser and ran the three related tests; all pass. The full integration suite was not run.”
```

## 6. Preserve Repository Intent

Before making non-trivial changes, understand what kind of repository you are in.

A production service, research prototype, reusable library, one-off script, ROS workspace, and personal automation project should not all receive the same engineering treatment.

Check for existing instructions and project signals first:

```text
README.md
CLAUDE.md
AGENTS.md
CONTRIBUTING.md
pyproject.toml / package.json / Cargo.toml
Makefile
CI workflows
existing tests
```

The repository itself is usually the best source of truth for how the code should be changed and verified.

## 7. Do Not Fabricate Evidence

Never invent:

- test results;
- benchmark numbers;
- command output;
- API behavior not verified in code or documentation;
- dependency versions;
- citations or links;
- files that do not exist.

When uncertain, distinguish clearly between:

```text
Observed in the repository
Inferred from the code
Recommended as a next step
```

That distinction makes agent-generated work much easier to trust and review.

---

## What good agent work looks like

These guidelines are working when you see:

- smaller, more focused diffs;
- fewer unnecessary abstractions;
- assumptions surfaced before they become bugs;
- tests tied directly to requested behavior;
- less collateral formatting or refactoring;
- clear statements about what was actually verified;
- project-specific conventions respected instead of overwritten.

## Suggested maintenance

Keep this baseline short enough that an agent will actually follow it. Put highly specific instructions in the target repository rather than continuously expanding the global rule set.

When a recurring failure pattern appears, add one concrete rule that prevents it. If a rule stops being useful, remove it.

The best instruction file is not the longest one; it is the one that consistently produces smaller, safer, easier-to-review changes.
