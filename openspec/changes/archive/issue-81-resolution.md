# SDD Change Resolution: VBA Preprocessor Comments & Line-Continuation (Issue #81)

- **Issue**: ardelperal/codegraph#81
- **Repository**: `codegraph-vba`
- **Date**: 2026-07-07

---

## 1. Problem Description

The VBA preprocessor pipeline inside `VbaExtractor` was executing in the following order:
1. `joinLineContinuations`
2. `preprocessConditionalCompilation`
3. `stripVbaComments`

This order led to two critical failures due to the interaction of line continuations (` _`) and comments:
- **False Line Continuations**: A comment line (either a full line or mid-line comment) that ended with an underscore `_` (e.g. `' This ends in _`) was treated as a continuation marker. This joined the comment line with the subsequent line of code, effectively commenting out and deleting valid source code.
- **Blocked Valid Continuations**: Legitimate line continuations that had a comment appended at the end of the physical line (e.g. `Call MySub(a, _ ' trailing comment`) were not joined because the continuation regex expected the newline character immediately after the `_` character.

---

## 2. Implementation & Resolution

The preprocessor pipeline order was updated in `src/extraction/vba-extractor.ts` to execute `stripVbaComments` **first**:
1. `stripVbaComments`: Removes all comments, converting fully commented lines into blank lines (preserving line-count parity).
2. `joinLineContinuations`: Resolves valid continuations. Since comments are already gone, any `_` characters inside comments have been removed, and valid continuations now end cleanly at the line boundary.
3. `preprocessConditionalCompilation`: Resolves active branches.

### Tests Added
We wrote two test cases in `__tests__/extraction-vba.test.ts` to cover both regression scenarios:
- **Test 1**: Verify that comments ending in `_` do not comment out subsequent code lines.
- **Test 2**: Verify that calls spanning multiple lines with a trailing comment after the continuation character are successfully extracted.

Both tests, along with the full VBA test suite, pass successfully.

---

## 3. Deployment Status

- **Commits**: `fix(vba): change preprocessing order to strip comments before joining line continuations`
- **Branch**: Merged into `main` and pushed to both `origin` (https://github.com/ardelperal/codegraph-vba.git) and `release-origin` (https://github.com/ardelperal/codegraph.git).
- **CI Status**: Passed successfully on GitHub Actions.
- **Issue Status**: Issue ardelperal/codegraph#81 has been successfully closed.
