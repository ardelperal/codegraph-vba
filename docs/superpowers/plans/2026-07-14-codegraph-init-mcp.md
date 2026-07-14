# CodeGraph Init MCP Tool Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Expose `codegraph init` as an opt-in MCP tool without widening the default MCP surface.

**Architecture:** Add one static tool definition and route it through a synchronous CLI subprocess helper in `ToolHandler`. Enforce the optional mutation-path allowlist before spawning the built CLI, then return the subprocess output and exit status unchanged as an MCP result.

**Tech Stack:** TypeScript, Node.js `child_process.spawnSync`, MCP tool definitions, Vitest.

---

### Task 1: Verify the RED baseline

**Files:**
- Test: `__tests__/mcp-tools-lifecycle.test.ts`

- [ ] **Step 1: Run the focused lifecycle tests with portable Node 22**

Run: `corepack pnpm exec vitest run mcp-tools-lifecycle -t codegraph_init`

Expected: the existing RED-first baseline documents the init behavior before MCP dispatch exists.

### Task 2: Add the opt-in init tool

**Files:**
- Modify: `src/mcp/tools.ts`

- [ ] **Step 1: Add the `codegraph_init` schema and mutating annotations**

Define required `path`, optional `force`, `verbose`, and `projectPath`; keep `DEFAULT_MCP_TOOLS` and `READ_ONLY_ANNOTATIONS` unchanged.

- [ ] **Step 2: Add minimal mutating dispatch**

Validate `CODEGRAPH_MCP_ALLOWLIST`, invoke `process.execPath dist/bin/codegraph.js init [--force] [--verbose] <path>` with `spawnSync`, combine stdout/stderr, and set `isError` from the exit status.

- [ ] **Step 3: Build and run the focused tests**

Run: `corepack pnpm run build`

Run: `corepack pnpm exec vitest run mcp-tools-lifecycle -t codegraph_init`

Expected: build succeeds and the init subset is green.

### Task 3: Document and verify the work unit

**Files:**
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Add the Unreleased changelog entry**

Document that `codegraph_init` is a CLI subprocess wrapper and remains off by default.

- [ ] **Step 2: Run lifecycle and VBA regression gates**

Run: `corepack pnpm exec vitest run mcp-tools-lifecycle`

Expected: only this issue's init tests are green; the other 17 lifecycle tests remain RED.

Run: `corepack pnpm exec vitest run vba extraction-sql-query sql-query-discovery`

Expected: all Windows-VBA baseline tests pass.

- [ ] **Step 3: Inspect the final diff without committing**

Run: `git diff --check` and `git status --short`.

Expected: only the plan, implementation, and changelog are modified; `.tooling/` remains untracked.
