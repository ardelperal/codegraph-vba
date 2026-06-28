# vba-form-ui-extraction

## Purpose

Indexes Dysflow-exported `.form.txt` and `.report.txt` UI files into the codegraph as a single `module` node per file plus one `property` node per control. The extractor emits exactly one `references` edge from the form/report module node to the sibling `.cls` of the same basename, so `codegraph_explore` can trace form → class. The extractor MUST NOT emit any `function`, `sub`, or non-form `class` nodes; that responsibility lives in `vba-code-extraction`, which reads the canonical `.cls`.

## Requirements

### Requirement: Form Module Node With Class Binding

The system MUST emit one `module` node per `.form.txt` source, named per `Attribute VB_Name` when present, else the file basename without extension. The system MUST emit one `references` edge from that module node to a node whose name matches the sibling `.cls` basename (so the graph can resolve the form → class binding at lookup time).

#### Scenario: Form module named from VB_Name

- GIVEN a `.form.txt` source containing `Attribute VB_Name = "Form_Main"` and one `TextBox` control block
- WHEN the extractor processes the source with filePath `src/forms/Form_Main.form.txt`
- THEN it emits one `module` node named `Form_Main`
- AND one `references` edge whose target node name equals `Form_Main`

#### Scenario: Form module named from filename when VB_Name absent

- GIVEN a `.form.txt` source with no `Attribute VB_Name` line and one `TextBox` control block
- WHEN the extractor processes the source with filePath `src/forms/Form_Main.form.txt`
- THEN it emits one `module` node named `Form_Main`
- AND one `references` edge whose target node name equals `Form_Main`

### Requirement: Controls Emit Property Nodes

The system MUST emit one `property` node per control declaration in a `.form.txt` source, with `metadata.controlType` set to the Access control type (e.g. `'TextBox'`, `'CommandButton'`, `'Label'`, `'ComboBox'`).

#### Scenario: Single textbox control

- GIVEN a `.form.txt` source containing one `TextBox` control block
- WHEN the extractor processes the source
- THEN it emits one `property` node
- AND `metadata.controlType` equals `'TextBox'`

#### Scenario: Multiple controls produce multiple property nodes

- GIVEN a `.form.txt` source containing one `TextBox` and one `CommandButton` control block
- WHEN the extractor processes the source
- THEN it emits two `property` nodes
- AND exactly one has `metadata.controlType === 'TextBox'`
- AND exactly one has `metadata.controlType === 'CommandButton'`

### Requirement: Reports Behave Like Forms

The system MUST treat `.report.txt` files identically to `.form.txt`: one `module` node (named per `Attribute VB_Name` or filename), one `references` edge to the sibling `.cls` basename, and one `property` node per control.

#### Scenario: Report module and properties

- GIVEN a `.report.txt` source containing `Attribute VB_Name = "Report_Orders"` and one `TextBox` control block
- WHEN the extractor processes the source with filePath `src/reports/Report_Orders.report.txt`
- THEN it emits one `module` node named `Report_Orders`
- AND one `references` edge whose target node name equals `Report_Orders`
- AND one `property` node with `metadata.controlType === 'TextBox'`

### Requirement: No Code Nodes From Form UI

The system MUST NOT emit any `function`, `sub`, or non-form `class` nodes from a `.form.txt` or `.report.txt` source. The only `module` node allowed is the form/report one above; the only other nodes allowed are `property` nodes for controls.

#### Scenario: Form source containing literal Sub keyword still produces no function nodes

- GIVEN a `.form.txt` source containing the literal text `Sub Form_Load()` in any context (comment, section header, embedded code block)
- WHEN the extractor processes the source
- THEN the result contains zero `function` nodes
- AND the result contains zero `class` nodes
- AND the result contains zero non-form `module` nodes