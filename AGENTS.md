# work-wrap

<!-- BEGIN DEV-KNOWLEDGE:shared-ai-workflow -->
## Shared AI workflow

- Treat this `AGENTS.md` as the cross-agent source of truth; Claude-specific notes may remain in `CLAUDE.md`.
- Codex must obtain an independent Claude PASS before mutation and another PASS for the actual diff before completion, commit, or push. On `NEEDS_USER_INPUT`, ask the user immediately.
- Use the project `*-context` skill as a router to authoritative specifications and verification commands.
- Preserve existing changes. Never let the central synchronizer commit or push this child repository.
- Shared tooling is addressed through `DEV_KNOWLEDGE_ROOT`; do not add machine-specific absolute paths to tracked files.
<!-- END DEV-KNOWLEDGE:shared-ai-workflow -->
