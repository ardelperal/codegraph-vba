/**
 * Server-level instructions emitted in the MCP `initialize` response.
 *
 * MCP clients (Claude Code, Cursor, opencode, LangChain, OpenAI Agent
 * SDK, …) surface this text in the agent's system prompt automatically,
 * giving the agent a high-level playbook for the codegraph toolset
 * before it sees individual tool descriptions.
 *
 * Goals when editing this:
 *   - Lead the agent to codegraph_explore for any structural/flow question
 *   - Reinforce "explore instead of Read/Grep" for indexed code
 *   - Anti-patterns (don't re-verify with grep; don't hand-reconstruct flows)
 *
 * Keep it tight. The agent reads this every session — long instructions
 * burn tokens. The DEFAULT MCP surface is `codegraph_explore` ALONE (see
 * DEFAULT_MCP_TOOLS in tools.ts) — reference only that tool here. The other
 * tools (node/search/callers/…) stay defined and are re-enablable via
 * CODEGRAPH_MCP_TOOLS, but they are NOT listed to agents, so don't name them.
 */
export const SERVER_INSTRUCTIONS = `# Codegraph — code intelligence over an indexed knowledge graph

Codegraph is a SQLite knowledge graph of every symbol, edge, and file in
the workspace — pre-computed structure you would otherwise re-derive by
reading files (cached intelligence: thousands of parse/trace decisions you
don't pay to re-reason each run). Reads are sub-millisecond; the index lags
writes by ~1s through the file watcher. Reach for it BEFORE *and* while
writing or editing code — not just for questions: one call returns the
verbatim source PLUS who calls it and what it affects, so you edit with the
blast radius in view. More accurate context, in far fewer tokens and
round-trips than reading files yourself.

## One tool: codegraph_explore — use it instead of reading files

There is a single tool, \`codegraph_explore\`, and it is Read-equivalent. It
takes either a natural-language question or a bag of symbol/file names and
returns the **verbatim, line-numbered source** of the relevant symbols
grouped by file — the same \`<n>\\t<line>\` shape \`Read\` gives you, safe to
\`Edit\` from — PLUS the call path among them (including dynamic-dispatch hops
like callbacks, React re-render, and JSX children that grep can't follow) and
a blast-radius summary of what depends on them.

Whether you're answering "how does X work" or implementing a change (fixing a
bug, adding a feature), call \`codegraph_explore\` before you Read. ONE call
usually answers the whole question. Codegraph IS the pre-built search index —
so running your own grep + read loop, or delegating the lookup to a separate
file-reading sub-task/agent, repeats work codegraph already did and costs more
for the same answer. A direct codegraph answer is typically one to a few
calls; a grep/read exploration is dozens.

## How to query

- **Almost any question — "how does X work", architecture, a bug, "what/where is X", or surveying an area** → \`codegraph_explore\` with a natural-language question or the relevant names. ONE capped call returns the verbatim source grouped by file; most often the ONLY call you need.
- **"How does X reach/become Y? / the flow / the path from X to Y"** → \`codegraph_explore\`, naming the symbols that span the flow (e.g. \`mutateElement renderScene\`) — it surfaces the call path among them, riding dynamic-dispatch hops, and returns their source.
- **Reading or editing a file/symbol you can name** → put its name or file path in the \`codegraph_explore\` query — it returns that current line-numbered source (safe to \`Edit\` from) with the call path and blast radius attached, so you don't Read it separately. For an overloaded name it returns every matching definition's body in one call.
- **Need more?** Call \`codegraph_explore\` again with more specific names — treat the source it returns as already Read.

## Anti-patterns

- **Trust codegraph's results — don't re-verify them with grep.** They come from a full AST parse; re-checking with grep is slower, less accurate, and wastes context.
- **Don't grep or Read first** to find or understand indexed code — ONE \`codegraph_explore\` returns the relevant symbols' source together in a single round-trip. Reach for raw \`Read\`/\`Grep\` only to confirm a specific detail codegraph didn't cover, or for what codegraph doesn't index (configs, docs).
- **Don't reconstruct a flow by hand** — name the endpoints in one \`codegraph_explore\` and it surfaces the path between them, dynamic-dispatch hops included.
- **After editing, check the staleness banner.** When a tool response starts with "⚠️ Some files referenced below were edited since the last index sync…", the listed files are pending re-index — Read those specific files for accurate content. Every file NOT in that banner is fresh, so still trust codegraph. A different, rarer banner — "⚠️ CodeGraph auto-sync is DISABLED…" — means live watching stopped entirely (the whole index is frozen, not just a few files); until it's resolved, Read files directly to confirm anything that may have changed.

## Limitations

- If a tool reports a project isn't indexed (no \`.codegraph/\`), stop calling codegraph tools for that project for the rest of the session and use your built-in tools there instead. Indexing is the user's decision — mention they can run \`codegraph init\` if it comes up, but don't run it yourself.
- Index lags file writes by ~1 second.
- Cross-file resolution is best-effort name matching; ambiguous calls may return multiple candidates.
- No live correctness validation — that's still the TypeScript compiler / test suite / linter's job. Codegraph supplements those with structural context they don't have.

## Supported Languages

The indexer recognizes a fixed set of languages; if you ask about symbols in a
file with an unsupported extension, codegraph will report the project isn't
indexed for that file and you should fall back to Read/Grep. The fork-specific
addition beyond upstream codegraph is **VBA / Access** (Dysflow export
format):

- **VBA / Access** - Dysflow exports Access/VBA source as \`.bas\`/\`.cls\`/
  \`.form.txt\`/\`.report.txt\`. Codegraph extracts \`.bas\`/\`.cls\` as \`module\`/
  \`class\`/\`function\` nodes with \`calls\`/\`implements\`/\`references\` edges
  (procedural-level; regex-based, not full AST). Cross-module calls, qualified
  \`Dim As\`, \`WithEvents\`, and SQL table references inside string literals
  emit synthesized edges tagged \`metadata.synthesizedBy\` (\`vba-name-resolution\`,
  \`vba-withevents\`, \`vba-sql-table\`). \`.form.txt\` and \`.report.txt\` are
  extracted as a \`module\` plus one \`property\` per Access control - **no**
  \`function\`/\`sub\`/\`class\` nodes come from form files; the canonical code
  lives in the sibling \`.cls\`, parsed by the same extractor on that file.
  Dysflow test manifests (\`tests.*.json\`) link each registered \`Test_*\`
  procedure to its manifest with a \`references\` edge tagged
  \`vba-test-manifest\` carrying the test name + tags, so \`getCallers\` of a
  production symbol reaches its covering test atoms with the manifest and tags
  to run.
  Pass \`projectPath\` to a codegraph index that includes VBA files.
- **VBA unresolved refs carry syntactic shape** (v1.7+). \`unresolved_refs.reference_kind\` is no longer the literal string \`"references"\` — it reports what the syntactic shape actually was. Values: \`call\` (paren-form or statement-form call site), \`qualified-call\` (\`obj.Foo(...)\` with runtime receiver), \`property-get\` / \`property-set\` (\`Me.Name\`, \`obj.Prop = value\`), \`bang-get\` / \`bang-set\` (\`Me!SubCtl\`, \`obj!Field = value\`), \`unqualified-ident\` (bare identifier like \`HayErrorEnRiesgo\` in an \`If\` condition), \`member-with\` (\`.Member\` inside a \`With\` block), \`dao-query\` (\`DoCmd.OpenQuery "X"\` argument). The legacy value \`references\` is retained on any path the round did not reclassify, so older SQL filters that key on it keep working. To find real missing callees, filter \`WHERE reference_kind IN ('call','qualified-call','unqualified-ident','member-with','bang-get')\` — that set has <10% false positives (DAO-field accesses, form-property reads, and bang refs no longer pollute the bucket).
- **Post-extraction stub resolver** (v1.7+). Edges with \`metadata.synthesizedBy='vba-name-resolution'\` start life pointing at a synthetic function node; the resolver at \`src/resolution/index.ts:resolveVbaCallStubs\` (invoked from \`indexAll\` and \`sync\`) walks them and repoints each \`target\` to the real \`nodes.id\` when one exists. Runtime-object calls (\`DAO.*\`, \`fso.*\`, \`ListBox.*\`, \`Collection.*\`, \`err.*\`, \`VBA.*\`, \`Application.*\`, \`Screen.*\`, \`DoCmd.*\`, \`CurrentDb.*\`, \`Forms\`, \`Reports\`, \`Debug\`, \`Modules\`, \`References\`, \`CommandBars\`, \`SysCmd\`, \`CreateObject\`, \`GetObject\`, \`Fields\`) are explicitly declined — they remain \`stub:true\` because they can never link to user code. Shadow user classes (e.g. a user class actually named \`DAO\` with an \`Execute\` method) are preserved and linked normally. Every stub edge carries \`metadata.repointDecision\` with one of \`reponted-to-real\`, \`declined-runtime\`, \`declined-ambiguous\`, \`declined-not-found\`, so consumers can tell apart a runtime-object decline from a genuinely-missing callee at the SQL layer.
`;

/**
 * Instructions variant sent when the server's own root has NO codegraph index.
 *
 * The tools are still exposed (gating tool availability on whether `./` has an
 * index is the bug behind #964: it breaks monorepos where only sub-projects are
 * indexed, and a server that started before `codegraph init` never surfaces the
 * tools afterward). Instead of an "inactive" note, this variant tells the agent
 * codegraph works **per project**: there's no default project to query, so pass
 * a `projectPath` to any project that HAS a `.codegraph/`. The full single-
 * project playbook ({@link SERVER_INSTRUCTIONS}) is sent instead when the root
 * IS indexed, so the common case stays tight.
 */
export const SERVER_INSTRUCTIONS_NO_ROOT_INDEX = `# Codegraph — available (per-project; pass projectPath)

Codegraph is a SQLite knowledge graph of a codebase's symbols, edges, and
files: one \`codegraph_explore\` call returns the verbatim, line-numbered source
of the relevant symbols PLUS the call paths between them and a blast-radius
summary — replacing a grep + Read loop with one round-trip.

This server started somewhere with no \`.codegraph/\` of its own, so there is no
default project — but the tools are available and work **per project**:

- To query a project that HAS a \`.codegraph/\` index (e.g. a service inside a
  monorepo, or a second repo), pass its path as \`projectPath\` to
  \`codegraph_explore\` (and any other codegraph tool). Codegraph resolves the
  nearest \`.codegraph/\` at or above that path and answers from it — for as many
  projects as you like in one session.
- For a project with no \`.codegraph/\`, use your built-in tools (Read/Grep/Glob)
  for that project. Indexing is the user's decision — don't run it yourself, but
  if it comes up they can run \`codegraph init\` in a project to enable codegraph
  there (a new index is picked up live, no restart).
`;
