# Project-Scoped Daemon Release

## Decision

Deliver issue #234 as a four-unit feature branch chain behind a tracker branch. The tracker records integration boundaries without describing child behavior as available on `main`.

## Scope

The chain adds one explicit project-scoped release operation across daemon lifecycle control, watchdog supervision, MCP proxy coordination, and the non-interactive CLI.

## Non-Negotiable Invariants

- **Generation-safe ownership**: release targets a canonical root and one authenticated daemon generation. It never signals an unverified PID.
- **Startup exclusion**: startup and release arbitrate ownership before publishing or deleting daemon artifacts.
- **Intentional release persists**: watchdog supervision cannot respawn a released root until an explicit resume.
- **Ordered requests**: a release waits for earlier requests and rejects later project calls on the owning proxy.
- **Explicit operator access**: MCP release remains opt-in and destructive; CLI release requires an explicit path.
- **Tests travel with behavior**: each child carries focused behavior-first tests and builds independently on Node 22.

## Chain Units

1. Lifecycle arbitration, authenticated generation-safe release, leases, heartbeat recovery, startup exclusion, and core tests.
2. Watchdog release tombstones, canonical aliases, explicit resume, and focused tests.
3. Proxy request barriers, per-root coordination, opt-in MCP surface, annotations, and focused tests.
4. Non-interactive CLI command, daemon manager presentation, parser and integration tests, and final cross-surface coverage.

## Final Operator Surfaces

The completed chain is intended to expose an opt-in `codegraph_release` MCP tool and a non-interactive `codegraph daemon stop --path <root>` command. These surfaces remain unavailable from `main` until the child chain is integrated.

## Consequences

Each child targets its immediate parent. Reviewers can validate and roll back one behavior boundary at a time while the tracker remains a durable integration map.

## Contributor Checklist

- [ ] Keep each child diff limited to its declared unit.
- [ ] Run the focused tests and build under Node 22.
- [ ] Preserve authenticated ownership, startup exclusion, and request ordering.
- [ ] Do not merge the tracker before all children are integrated.
