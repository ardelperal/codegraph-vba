# Project-Scoped Daemon Release

## Decision

Provide one explicit release lifecycle per canonical project root across daemon control, watchdog supervision, MCP coordination, and the CLI.

## Scope

The chain adds one explicit project-scoped release operation across daemon lifecycle control, watchdog supervision, MCP proxy coordination, and the non-interactive CLI.

## Non-Negotiable Invariants

- **Generation-safe ownership**: release targets a canonical root and one authenticated daemon generation. It never signals an unverified PID.
- **Startup exclusion**: startup and release arbitrate ownership before publishing or deleting daemon artifacts.
- **Intentional release is stable**: watchdog supervision does not immediately respawn a deliberately released root. A later normal CodeGraph launch or query may establish a new lifecycle.
- **Ordered requests**: a release waits for earlier requests and rejects later project calls on the owning proxy.
- **Explicit operator access**: MCP release remains opt-in and destructive; CLI release requires an explicit path.
- **Project isolation**: releasing one root leaves unrelated projects available. Process-wide cleanup is not part of this design.
- **Idempotent outcomes**: repeated release returns a typed outcome: `released`, `not-running`, `no-daemon`, `identity-mismatch`, `unreachable`, or `termination-failed`.
- **Resource release only**: release does not remove a worktree or project directory.
- **Tests travel with behavior**: each child carries focused behavior-first tests and builds independently on Node 22.

## Chain Units

1. Lifecycle arbitration, authenticated generation-safe release, leases, heartbeat recovery, startup exclusion, and core tests.
2. Watchdog release tombstones, canonical aliases, explicit resume, and focused tests.
3. Proxy request barriers, per-root coordination, opt-in MCP surface, annotations, and focused tests.
4. Non-interactive CLI command, daemon manager presentation, parser and integration tests, and final cross-surface coverage.

## Operator Surfaces

The non-interactive CLI surface is `codegraph daemon stop --path <project-root>`. The destructive MCP surface is `codegraph_release`, disabled by default and requiring an explicit `path`.

## Consequences

Each child targets its immediate parent. Reviewers can validate and roll back one behavior boundary at a time while the tracker remains a durable integration map.

## Contributor Checklist

- [ ] Keep each child diff limited to its declared unit.
- [ ] Run the focused tests and build under Node 22.
- [ ] Preserve authenticated ownership, startup exclusion, and request ordering.
- [ ] Do not merge the tracker before all children are integrated.
