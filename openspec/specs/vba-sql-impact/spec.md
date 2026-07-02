# Specification: VBA SQL Impact Analysis

## 1. Overview
The `vba-sql-impact` capability traces Microsoft Access saved queries (QueryDefs) to their VBA callers, extracts form/report data bindings, and resolves SQL table/column usages (including alias resolution) to identify downstream change impacts.

## 2. Requirements and Scenarios

### 2.1 Caller Tracing
**Requirement:** The capability MUST identify VBA modules and lines that reference a saved query either via literal string arguments in `OpenRecordset` calls or by accessing a query through the `QueryDefs` collection.

**Scenario 1: OpenRecordset Literal Reference**
- **GIVEN** a VBA module containing `Set rs = db.OpenRecordset("qryGetRiesgos", dbOpenSnapshot)`
- **WHEN** tracing references for the query `qryGetRiesgos`
- **THEN** the system MUST detect this line and file as a caller.

**Scenario 2: QueryDef Reference**
- **GIVEN** a VBA module containing `Set qdf = db.QueryDefs("qryGetRiesgos")`
- **WHEN** tracing references for the query `qryGetRiesgos`
- **THEN** the system MUST detect this line and file as a caller.

---

### 2.2 Form and Report Bindings Extraction
**Requirement:** The capability SHALL parse form (`.form.txt`) and report (`.report.txt`) configuration files to extract `RecordSource` and `RowSource` properties and map them to their corresponding query or table.

**Scenario 1: RecordSource Extraction**
- **GIVEN** a form file `frmRiesgos.form.txt` containing the property:
  `RecordSource = "qryGetRiesgos"`
- **WHEN** extracting bindings for `frmRiesgos`
- **THEN** the system MUST identify `RecordSource` as `qryGetRiesgos`.

**Scenario 2: RowSource Extraction**
- **GIVEN** a form file `frmRiesgos.form.txt` containing:
  `RowSource = "SELECT id, nombre FROM tblUsuarios"`
- **WHEN** extracting bindings for controls in `frmRiesgos`
- **THEN** the system MUST identify `RowSource` as the query/table `tblUsuarios`.

---

### 2.3 Column & Table Alias Resolution
**Requirement:** The capability MUST parse SQL queries, identify tables and columns touched, and resolve table aliases (defined via `AS` or implicit space) to map column names back to their original database tables.

**Scenario 1: Column Alias Resolution in JOIN**
- **GIVEN** a saved query file `qryGetRiesgos.sql` containing:
  `SELECT r.estado, u.nombre FROM tblRiesgos AS r INNER JOIN tblUsuarios u ON r.user_id = u.id`
- **WHEN** analyzing the column lineage
- **THEN** the system SHALL map:
  - `r.estado` -> `tblRiesgos.estado`
  - `u.nombre` -> `tblUsuarios.nombre`
  - `r.user_id` -> `tblRiesgos.user_id`
  - `u.id` -> `tblUsuarios.id`

---

### 2.4 Downstream Impact Reporting
**Requirement:** The capability SHALL detect all tables read/written by a saved query and list downstream warnings if a target table is modified.

**Scenario 1: Table Schema Change Impact**
- **GIVEN** a saved query `qryGetRiesgos` referencing `tblRiesgos`
- **AND** a form `frmRiesgos` bound to `qryGetRiesgos`
- **AND** a VBA module calling `qryGetRiesgos`
- **WHEN** assessing impact for a schema change in `tblRiesgos`
- **THEN** the system SHOULD list the query `qryGetRiesgos`, the form `frmRiesgos`, and the VBA caller as impacted downstream.
