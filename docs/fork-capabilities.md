# Capacidades del fork hasta v1.17.1

[Volver al README](../README.md)

Esta página distingue la base heredada, las capacidades publicadas del fork y los límites de su evidencia estática. Los contratos detallados permanecen en las referencias enlazadas.

## Versión y linaje verificado

| Dato | Evidencia |
|---|---|
| Versión del fork | `v1.17.1`, en [package.json](../package.json) y [CHANGELOG.md](../CHANGELOG.md) |
| Base upstream integrada | CodeGraph `v1.4.1`, commit [`ecc8b30`](https://github.com/colbymchenry/codegraph/commit/ecc8b307ac2f8a7d06bff02ee513c4ea2380b2f8) |
| Merge en el fork | [`910c7bb`](https://github.com/ardelperal/codegraph-vba/commit/910c7bb09fcc1e681e86803f93fffdb758086a22), antecesor de la versión del fork |
| Upstream `v1.6.0` | [`dfccdf6`](https://github.com/colbymchenry/codegraph/commit/dfccdf62547fcd76d343344d823a0e1998d3a89f) no es antecesor del fork `v1.17.1`; no está integrado |

El fork `v1.6.0` integró upstream `v1.4.1`. La coincidencia de números con upstream `v1.6.0` no implica igualdad de código.

Una rama de planificación o un intento de merge no demuestra integración. Tampoco se deben atribuir al fork funciones por aparecer en la web de upstream.

## Orientación rápida

Dysflow exporta el proyecto Access; CodeGraph indexa ese árbol de texto. Para empezar, siga la [instalación](../README.md#get-started) y el [ejemplo ejecutable de Access](../README.md#vba--access--dysflow-integration).

La base integrada aporta grafo local SQLite, búsqueda de símbolos y texto, consultas de llamadas e impacto, indexación incremental, watcher y MCP para varios lenguajes.

Estas superficies se describen en el [README](../README.md#how-it-works). No prometen paridad con versiones posteriores de upstream.

## Capacidades específicas del fork

| Área | Alcance publicado | Referencia verificable |
|---|---|---|
| Lenguaje VBA | `Sub`, `Function`, propiedades, parámetros, variables de módulo, tipos, constantes, declaraciones API y relaciones de llamadas | [Pruebas VBA](../__tests__/extraction-vba.test.ts), [parámetros](../__tests__/extraction-vba-parameters.test.ts), [variables](../__tests__/extraction-vba-module-variables.test.ts) |
| UI de Access | Layouts de formularios e informes, controles, secciones, subformularios, fuentes de datos y código asociado | [Modelo UI](../__tests__/extraction-vba-control-modeling.test.ts), [subformularios](../__tests__/extraction-vba-source-object.test.ts) |
| Eventos | Handlers de controles y ciclo de vida; expresiones de evento; `WithEvents`, eventos personalizados y `RaiseEvent` | [Ciclo de vida](../__tests__/extraction-vba-form-lifecycle-events.test.ts), [expresiones](../__tests__/extraction-vba-expression-events.test.ts), [eventos personalizados](../__tests__/extraction-vba-event-synth.test.ts) |
| SQL y datos | QueryDefs, lecturas/escrituras, DDL, bases externas, SQL en bindings, funciones de dominio y esquema backend con columnas | [SQL](../__tests__/extraction-sql-query.test.ts), [bindings](../__tests__/extraction-vba-runtime-binding-sql.test.ts), [dominio](../__tests__/extraction-vba-domain-functions.test.ts), [backend](../__tests__/extraction-access-erd.test.ts) |
| Operaciones Access | Apertura/cierre de objetos, macros, tablas, exportaciones, correo, recordsets y consultas con nombre | [DoCmd](../__tests__/extraction-vba-docmd-object-verbs.test.ts), [cierre](../__tests__/extraction-vba-docmd-close.test.ts), [consultas](../__tests__/extraction-vba-query-name.test.ts) |
| Errores | Política por procedimiento, etiquetas, rutas de error y canales de registro, presentación o relanzamiento | [Guía de errores](vba-error-handling.md), [canales](../__tests__/extraction-vba-error-channel.test.ts) |
| Artefactos Dysflow | Manifiestos y secuencias JSON compatibles enlazan pruebas con procedimientos, etiquetas y políticas | [Manifiestos](../__tests__/extraction-vba-test-manifest.test.ts), [secuencias](../__tests__/extraction-vba-test-sequence.test.ts) |
| Integración | CLI headless, pruebas afectadas mediante `affected`, esquema del índice y herramientas MCP adicionales habilitadas explícitamente | [Contrato externo](external-integration.md), [esquema](index-schema.md), [MCP](../README.md#mcp-tools) |

El proyecto puede desactivar la expansión de UI, manifiestos y secuencias Dysflow con `vba.dysflowExport: false`; véase el [resolvedor específico](../src/extraction/frameworks/dysflow-export.ts).

### Evidencia de comportamiento en v1.17

`getBehaviorEvidence` y la herramienta opcional `codegraph_behavior_evidence` reúnen handler, rutas de llamadas y tablas alcanzadas para un control o handler concreto.

Mantienen las ramas separadas, declaran ambigüedades y evidencia sin resolver. La ausencia de tablas no significa que el procedimiento sea inocuo.

En `v1.17.1`, el presupuesto limita tanto resultados como trabajo de recorrido; el resultado indica si queda incompleto. Los handlers admiten selección por nombre o archivo del layout.

El [contrato de uso](../README.md#behavior-evidence-for-one-control) tiene pruebas de [ensamblado](../__tests__/vba-behavior-evidence.test.ts), [presupuesto](../__tests__/vba-behavior-evidence-budget.test.ts) y [selección](../__tests__/vba-behavior-evidence-layout.test.ts).

La [aceptación semántica](../README.md#semantic-acceptance) indexa una copia fresca del corpus y compara respuestas con expectativas escritas desde el código fuente.

## Invariantes y límites

- **Fuentes canónicas:** `.bas` y `.cls` aportan código; `.form.txt` y `.report.txt`, UI; `.sql`, consultas guardadas. El código incrustado en SaveAsText no se duplica como procedimientos.
- **Propiedad y dirección:** cada control pertenece a su layout. La relación de evento se almacena handler → control/layout; recorrer desde el control requiere seguirla en sentido inverso.
- **Análisis estático:** no abre ni ejecuta el `.accdb`, no verifica compilación ni demuestra que el binario coincida con la exportación.
- **Evidencia incompleta:** referencias dinámicas o no resueltas y relaciones heurísticas requieren comprobación del consumidor. Ninguna relación ausente prueba ausencia de efectos.
- **Formato acotado:** `.frm` y `.dsr` heredados no forman parte del contrato de exportación Dysflow documentado.

Las [pruebas del ejemplo](../__tests__/vba-documented-example.test.ts) y de [semántica del consumidor](../__tests__/vba-consumer-graph-semantics.test.ts) respaldan el modelo estático, no ejecución real de Access.

## Comprobación antes de actualizar esta página

- [ ] Contrastar versión y ascendencia con Git; no equiparar números de release del fork y upstream.
- [ ] Comprobar las afirmaciones de comportamiento contra pruebas y contratos actuales, no solo contra este resumen.
- [ ] Mantener explícitos los límites de fuentes, resolución y presupuesto al añadir capacidades.

## Navegación

[Notas del fork](../FORK_NOTES.md) · [Historial](../CHANGELOG.md) · [Referencias VBA](vba-reference-kinds.md) · [Decisiones del resolvedor](vba-stub-repoint-decision.md)
