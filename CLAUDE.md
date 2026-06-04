# CLAUDE.md

Guidance for AI assistants (Claude Code and others) working in this repository.

## Project status: greenfield

**NewCreative** is a brand-new project. As of this writing the repository
contains only:

- `README.md` — currently a one-line stub (`# NewCreative` / `Truth`).
- `LICENSE` — The Unlicense (public domain dedication).
- `CLAUDE.md` — this file.

There is **no source code, build system, test suite, dependency manifest, or
CI configuration yet.** Do not assume a language, framework, or toolchain is in
place — none has been chosen. Treat statements below about "where things go" as
conventions to establish as the project grows, not as descriptions of existing
structure.

> When you add the first real code, **update this file in the same change** so
> it stays an accurate map of the codebase. An out-of-date CLAUDE.md is worse
> than none.

## License: public domain

This project is released into the public domain under The Unlicense. Practical
implications for contributions:

- Only add code/content that is itself public-domain-compatible. Do **not**
  paste in code under restrictive licenses (GPL, proprietary, etc.) or content
  that would impose attribution or copyleft obligations.
- Be cautious with third-party dependencies — prefer permissively licensed ones
  (MIT, BSD, Apache-2.0, ISC, Unlicense, CC0) and note any license constraints
  the dependency carries.

## Git workflow & conventions

- **Default branch:** `main`.
- **Do not commit directly to `main`.** Develop on a feature branch and open a
  pull request for review. When a session specifies a designated working
  branch, use exactly that branch.
- **Pushing:** use `git push -u origin <branch-name>`. After pushing, open a PR
  (ready for review, not a draft) if one does not already exist for the branch.
- **Commit messages:** write clear, descriptive messages in the imperative mood
  (e.g. "Add user auth module", not "added stuff"). Keep the subject line
  concise and explain the *why* in the body when it isn't obvious.
- **Keep changes focused.** One logical change per PR where practical.

## Conventions to follow as the project grows

These are the defaults to apply once code is introduced. Update this section
with concrete commands and paths the moment the toolchain is chosen.

1. **Pick and document the stack.** When the first language/framework is added,
   record here: how to install dependencies, how to build, how to run, and how
   to test. Include the exact commands.
2. **Match existing style.** Once a codebase exists, mirror its formatting,
   naming, and structural patterns rather than importing outside conventions.
   Prefer the project's configured formatter/linter over manual styling.
3. **Add tests with features.** Establish a test directory and runner early;
   new behavior should ship with tests.
4. **Keep the root tidy.** Source under a conventional directory (e.g. `src/`),
   tests alongside or under `tests/`, docs under `docs/`. Add a `.gitignore`
   appropriate to the chosen stack before committing build artifacts or
   dependencies.

## Working agreement for AI assistants

- **Verify, don't assume.** This file describes a near-empty repo. Before
  acting on any assumption about structure or tooling, check the actual files
  on disk — the project may have grown since this was written.
- **Be honest about state.** If something doesn't exist yet, say so rather than
  inventing it. Don't fabricate build/test commands that haven't been set up.
- **Ask when genuinely ambiguous.** Foundational decisions (language, framework,
  architecture) belong to the maintainer. Surface the choice instead of picking
  silently.
- **Leave the repo better documented.** Whenever you make a structural change,
  reflect it here.

---
*Maintainers: keep this file current. When in doubt, the working tree is the
source of truth — this document should describe it accurately.*
