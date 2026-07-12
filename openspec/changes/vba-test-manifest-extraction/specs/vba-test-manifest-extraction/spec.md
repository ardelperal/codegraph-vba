# Delta for vba-test-manifest-extraction

## ADDED Requirements

### Requirement: Detect Dysflow VBA test manifests

The system MUST route a JSON file to the test-manifest extractor when, and only
when, its basename matches `^tests(\.[\w-]+)*\.json$` (case-insensitive) AND its
parsed content has a top-level `tests` array whose items carry a string
`procedure`. All other `.json` files MUST be ignored by this extractor.

#### Scenario: A tests.vba.*.json manifest is recognized

- GIVEN a file `tests/tests.vba.smoke.json` with `{ "tests": [ { "procedure": "Test_X_RunAll" } ] }`
- WHEN the extractor pipeline processes it
- THEN it is handled by `VbaTestManifestExtractor`
- AND a `file` node is emitted for the manifest

#### Scenario: A non-manifest JSON is ignored

- GIVEN a `package.json` or a `tsconfig.json`
- WHEN the extractor pipeline processes it
- THEN `VbaTestManifestExtractor` emits no nodes or references for it

#### Scenario: A tests.*.json without a tests array is ignored

- GIVEN a file `tests.config.json` with `{ "runnerPolicy": {}, "procedures": ["A"] }` (no `tests` array of objects)
- WHEN the extractor pipeline processes it
- THEN no `vba-test-manifest` references are emitted (out of scope: sequences)

### Requirement: Link registered test atoms to manifest metadata

For each `tests[]` entry carrying a string `procedure`, the system MUST emit one
`UnresolvedReference` toward that procedure name with
`metadata.synthesizedBy === 'vba-test-manifest'` and metadata `testName`
(the entry `name`, or the procedure name when `name` is absent), `tags`
(the entry `tags`, or `[]`), and `manifestFile` (the manifest path). It MUST NOT
emit a duplicate node for the procedure.

#### Scenario: Each test entry emits one manifest reference

- GIVEN a manifest with three entries each carrying a `procedure`
- WHEN the extractor processes it
- THEN exactly three `UnresolvedReference`s tagged `vba-test-manifest` are emitted
- AND each carries its `testName`, `tags`, and `manifestFile`

#### Scenario: Tags and name default when absent

- GIVEN an entry `{ "procedure": "Test_Y" }` with no `name` and no `tags`
- WHEN the extractor processes it
- THEN the reference `metadata.testName === 'Test_Y'` and `metadata.tags` is `[]`

#### Scenario: Malformed manifest does not crash the index

- GIVEN a `tests.vba.broken.json` containing invalid JSON
- WHEN the extractor processes it
- THEN a low-severity `ExtractionError` is recorded
- AND zero references are emitted
- AND no exception propagates out of extraction

### Requirement: Resolve manifest references to the test function nodes

The `ReferenceResolver` MUST bind each `vba-test-manifest` reference to the
existing `function` node whose name equals the referenced `procedure`,
producing a `references` edge from the manifest `file` node to that function
node, carrying the reference metadata. When no such function node exists, the
reference MUST remain unresolved (never silently dropped).

#### Scenario: A manifest reference resolves to its Test_ procedure

- GIVEN an indexed module defining `Public Sub Test_X_RunAll()` and a manifest naming `procedure: "Test_X_RunAll"`
- WHEN resolution runs
- THEN a `references` edge exists from the manifest file node to the `Test_X_RunAll` function node
- AND `edge.metadata.synthesizedBy === 'vba-test-manifest'`

#### Scenario: A manifest naming a missing procedure stays unresolved

- GIVEN a manifest naming `procedure: "Test_Removed"` with no such procedure indexed
- WHEN resolution runs
- THEN no `references` edge is created for it
- AND the unresolved reference is retained (drift surfaced, not hidden)
