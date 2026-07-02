import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { DatabaseConnection } from '../src/db';
import { QueryBuilder } from '../src/db/queries';
import {
  traceVbaCallers,
  extractFormBindings,
  resolveSqlLineage,
  runImpactAnalysis
} from '../src/utils/sql-impact-helpers';

describe('VBA SQL Impact Helpers', () => {
  describe('traceVbaCallers', () => {
    it('should trace callers referencing a query in OpenRecordset', () => {
      const content = `
        Sub Test1()
          Set rs = db.OpenRecordset("qryGetRiesgos", dbOpenSnapshot)
        End Sub
        Sub Test2()
          Set rs = db.OpenRecordset("otherQuery")
        End Sub
        Sub Test3()
          Set rs = db.OpenRecordset("qryGetRiesgos")
        End Sub
      `;
      const result = traceVbaCallers(content, 'qryGetRiesgos');
      expect(result).toEqual([3, 9]);
    });

    it('should trace callers referencing a query in QueryDefs', () => {
      const content = `
        Set qdf = db.QueryDefs("qryGetRiesgos")
        Set qdf = db.QueryDefs("otherQuery")
        Set qdf = db.querydefs("qrygetriesgos")
      `;
      const result = traceVbaCallers(content, 'qryGetRiesgos');
      expect(result).toEqual([2, 4]);
    });

    it('should trace callers without parentheses', () => {
      const content = `
        db.OpenRecordset "qryGetRiesgos"
        db.OpenRecordset "otherQuery"
      `;
      const result = traceVbaCallers(content, 'qryGetRiesgos');
      expect(result).toEqual([2]);
    });
  });

  describe('extractFormBindings', () => {
    it('should extract RecordSource from form definition', () => {
      const content = `
        Begin Form
          RecordSource ="qryGetRiesgos"
          Caption ="Riesgos Form"
        End
      `;
      const result = extractFormBindings(content);
      expect(result.recordSource).toBe('qryGetRiesgos');
      expect(result.rowSources).toEqual([]);
    });

    it('should extract RowSource from ComboBox and ListBox controls', () => {
      const content = `
        Begin Form
          RecordSource ="qryGetRiesgos"
          Begin ComboBox
            Name ="cboUsuario"
            RowSource ="tblUsuarios"
          End
          Begin TextBox
            Name ="txtNombre"
          End
          Begin ListBox
            Name ="lstRoles"
            RowSource ="SELECT id FROM tblRoles"
          End
        End
      `;
      const result = extractFormBindings(content);
      expect(result.recordSource).toBe('qryGetRiesgos');
      expect(result.rowSources).toEqual([
        { control: 'cboUsuario', target: 'tblUsuarios' },
        { control: 'lstRoles', target: 'tblRoles' }
      ]);
    });
  });

  describe('resolveSqlLineage', () => {
    it('should resolve tables and column aliases in a JOIN query', () => {
      const sql = 'SELECT r.estado, u.nombre FROM tblRiesgos AS r INNER JOIN tblUsuarios u ON r.user_id = u.id';
      const result = resolveSqlLineage(sql);
      expect(result.tables.sort()).toEqual(['tblRiesgos', 'tblUsuarios'].sort());
      expect(result.lineage).toEqual([
        { source: 'r.estado', resolved: 'tblRiesgos.estado' },
        { source: 'u.nombre', resolved: 'tblUsuarios.nombre' },
        { source: 'r.user_id', resolved: 'tblRiesgos.user_id' },
        { source: 'u.id', resolved: 'tblUsuarios.id' }
      ]);
    });

    it('should handle complex casing and spacing in SQL', () => {
      const sql = '  SELECT   a.col1  ,  b.col2   FROM   tableA   AS   a   JOIN   tableB   b   ON   a.id = b.id ';
      const result = resolveSqlLineage(sql);
      expect(result.tables.sort()).toEqual(['tableA', 'tableB'].sort());
      expect(result.lineage).toEqual([
        { source: 'a.col1', resolved: 'tableA.col1' },
        { source: 'b.col2', resolved: 'tableB.col2' },
        { source: 'a.id', resolved: 'tableA.id' },
        { source: 'b.id', resolved: 'tableB.id' }
      ]);
    });
  });

  describe('VBA SQL Impact Integration', () => {
    let tempDir: string;
    let dbConn: DatabaseConnection;

    beforeEach(() => {
      tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-sql-impact-'));
      const dbPath = path.join(tempDir, 'test_codegraph.db');
      dbConn = DatabaseConnection.initialize(dbPath);
    });

    afterEach(() => {
      dbConn.close();
      if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      }
    });

    it('should run full-flow lineage extraction and downstream impact assessment', () => {
      const q = new QueryBuilder(dbConn.getDb());
      q.insertNodes([
        {
          id: 'btn1',
          kind: 'control',
          name: 'btnSave',
          qualifiedName: 'btnSave',
          filePath: 'src/forms/frmRiesgos.form.txt',
          language: 'vba',
          startLine: 1, endLine: 1, startColumn: 0, endColumn: 0,
          updatedAt: Date.now()
        },
        {
          id: 'btn1_click',
          kind: 'event',
          name: 'btnSave_Click',
          qualifiedName: 'btnSave_Click',
          filePath: 'src/modules/ModRiesgos.bas',
          language: 'vba',
          startLine: 10, endLine: 15, startColumn: 0, endColumn: 0,
          updatedAt: Date.now()
        },
        {
          id: 'load_data_func',
          kind: 'function',
          name: 'LoadData',
          qualifiedName: 'LoadData',
          filePath: 'src/modules/ModRiesgos.bas',
          language: 'vba',
          startLine: 20, endLine: 25, startColumn: 0, endColumn: 0,
          updatedAt: Date.now()
        }
      ]);

      q.insertEdges([
        { source: 'btn1', target: 'btn1_click', kind: 'defines-event' },
        { source: 'btn1_click', target: 'load_data_func', kind: 'calls' }
      ]);

      const modulesDir = path.join(tempDir, 'src', 'modules');
      const formsDir = path.join(tempDir, 'src', 'forms');
      const queriesDir = path.join(tempDir, 'queries');
      fs.mkdirSync(modulesDir, { recursive: true });
      fs.mkdirSync(formsDir, { recursive: true });
      fs.mkdirSync(queriesDir, { recursive: true });

      const vbaContent = `
Attribute VB_Name = "ModRiesgos"
Public Sub btnSave_Click()
    Call LoadData
End Sub

Public Sub LoadData()
    ' Line 20
    ' Line 21
    ' Line 22
    Set rs = db.OpenRecordset("qryGetRiesgos", dbOpenSnapshot)
    ' Line 24
End Sub
`;
      fs.writeFileSync(path.join(modulesDir, 'ModRiesgos.bas'), vbaContent.trim());

      const formContent = `
Begin Form
  RecordSource ="qryGetRiesgos"
  Begin ComboBox
    Name ="cboUsuario"
    RowSource ="tblUsuarios"
  End
End
`;
      fs.writeFileSync(path.join(formsDir, 'frmRiesgos.form.txt'), formContent.trim());

      const sqlContent = `
SELECT r.estado, u.nombre
FROM tblRiesgos AS r
INNER JOIN tblUsuarios u ON r.user_id = u.id
`;
      fs.writeFileSync(path.join(queriesDir, 'qryGetRiesgos.sql'), sqlContent.trim());

      const result = runImpactAnalysis(dbConn.getDb(), tempDir, 'qryGetRiesgos');

      expect(result.query_name).toBe('qryGetRiesgos');
      expect(result.callers).toEqual([
        {
          file: 'src/modules/ModRiesgos.bas',
          line: 10,
          context: '    Set rs = db.OpenRecordset("qryGetRiesgos", dbOpenSnapshot)'
        }
      ]);
      expect(result.form_bindings).toEqual([
        {
          file: 'src/forms/frmRiesgos.form.txt',
          control: 'Form',
          property: 'RecordSource',
          target: 'qryGetRiesgos'
        },
        {
          file: 'src/forms/frmRiesgos.form.txt',
          control: 'cboUsuario',
          property: 'RowSource',
          target: 'tblUsuarios'
        }
      ]);
      expect(result.tables_touched.sort()).toEqual(['tblRiesgos', 'tblUsuarios'].sort());
      expect(result.lineage).toEqual([
        { source: 'r.estado', resolved: 'tblRiesgos.estado' },
        { source: 'u.nombre', resolved: 'tblUsuarios.nombre' },
        { source: 'r.user_id', resolved: 'tblRiesgos.user_id' },
        { source: 'u.id', resolved: 'tblUsuarios.id' }
      ]);
      expect(result.downstream_impact.queries).toEqual(['qryGetRiesgos']);
      expect(result.downstream_impact.forms).toEqual(['frmRiesgos']);
      expect(result.downstream_impact.vba_callers).toContain('ModRiesgos.bas');
    });
  });
});
