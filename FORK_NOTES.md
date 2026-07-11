# Fork notes

Este fork de [`colbymchenry/codegraph`](https://github.com/colbymchenry/codegraph)
añade soporte para indexar proyectos Microsoft Access / VBA
(`.bas`, `.cls`, `.form.txt`, `.report.txt`).

## Por qué

El upstream no incluye VBA en su mapa de extensiones
(`src/extraction/grammars.ts`), así que los proyectos de esta organización no
entran al índice. La lógica de negocio vive en los `.bas` y `.cls`, no en el
resto del repo, por lo que el grafo queda inútil sin VBA.

## Convención de código de formularios (no negociable)

- `.cls` es la **fuente de verdad** del código de un formulario.
- `.form.txt` contiene **solo** UI (controles, layout, propiedades).
- El extractor UI no debe emitir nodos `function` / `sub` / `module` desde
  `.form.txt`: cualquier código que el editor Access serializa ahí se ignora y
  se sobrescribe en el próximo import por Dysflow.

## Plan de cambios (en rama aparte, no en `main`)

1. `src/types.ts`: añadir `'vba'` al array `LANGUAGES`.
2. `src/extraction/grammars.ts`:
   - Mapear `.bas` / `.cls` / `.frm` / `.dsr` / `.form.txt` / `.report.txt` a
     `'vba'`.
   - Ampliar `isLanguageSupported` y `getLanguageDisplayName`.
3. `src/extraction/vba-extractor.ts` (nuevo): extractor regex para `.bas` y
   `.cls`. Detecta:
   - `Public Function / Private Function / Public Sub / Private Sub`.
   - `Property Get / Let / Set`.
   - `Dim … As New Clase`, `New Clase`.
   - `WithEvents m_X As Form_Foo` + handlers `m_X_Evento`.
   - Call sites `Identificador.Identificador(`.
   - SQL embebido: `FROM <tabla>`, `INTO <tabla>`, `UPDATE <tabla>`.
4. `src/extraction/vba-form-extractor.ts` (nuevo): extractor UI para
   `.form.txt` y `.report.txt`. Solo nodos `property` por control y un edge
   `references` al `.cls` del mismo `basename`.
5. `src/extraction/index.ts`: registrar ambos extractores donde hoy se enrutan
   `liquid`, `razor`, `vue`, `svelte`.
6. Tests en `__tests__/` con fixtures de `00_EXPEDIENTES_staging` y
   `00_NO_CONFORMIDADES_staging`.

Todas las edges sintéticas deben llevar `provenance: 'heuristic'` y
`metadata.synthesizedBy` (`vba-name-resolution`, `vba-new-binding`,
`vba-withevents`, `vba-sql-table`, `vba-form-binding`).

## Remotes

- `origin`: `ardelperal/codegraph-vba` (este fork).
- `upstream`: `colbymchenry/codegraph` (fuente original).

## Sincronización con upstream

```bash
git fetch upstream
git checkout main
git merge upstream/main
```
