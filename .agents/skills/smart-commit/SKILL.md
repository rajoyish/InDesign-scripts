---
name: smart-commit
description: Trigger when the user says "commit", "save changes", or runs "/commit".
---

# Smart Commit Workflow

When triggered, follow this exact workflow:

1. **Pre-flight Check**: Use the `run_command` tool to execute `git status --porcelain`. If the output is completely empty, inform the user there are no changes to commit and stop execution.
2. **Stage Changes**: Use the `run_command` tool to execute `git add .` to stage all modified, deleted, and untracked files, then execute `git diff --staged` to read the changes.
3. **Generate Commit Message**: Read the staged diff output. Generate a commit message based on the diff in this exact format:

   type(scope): short subject

   - bullet of what changed
   - bullet of why

4. **Commit**: Use the `run_command` tool to execute `git commit -m "<generated message>"`. Output a brief success message once the commit is finalized.

## Commit Message Rules:
- **Types**: Use one of `feat`, `fix`, `refactor`, `chore`, `docs`, `style`, `test`.
- **Scope**: Keep scopes relevant to the architecture (e.g., `ui`, `api`, `deps`, `auth`, `seo`).
- **Subject**: Must be under 60 characters and in the imperative mood (e.g., "add eSewa QR", not "added eSewa QR").
- **Body bullets**: Mandatory if the code change spans multiple files or complex logic.
- **NEVER** include a Co-Authored-By trailer.
