# Notas del fork

`codegraph-vba` amplía CodeGraph para analizar proyectos VBA, Microsoft Access y exportaciones de Dysflow.

El alcance publicado hasta `v1.17.1` y el linaje upstream verificado se recogen en [Capacidades del fork](docs/fork-capabilities.md).

## Trabajo ya implementado

El plan inicial de extractores está entregado. El modelo incluye código VBA, UI y eventos de Access, consultas y datos, análisis de errores y evidencia de comportamiento.

Consulte el [modelo y los ejemplos actuales](README.md#vba--access--dysflow-integration) y el [historial de cambios](CHANGELOG.md), no el antiguo plan de implementación.

## Invariantes de las fuentes

- **Código canónico:** `.cls` contiene el código asociado a formularios e informes; `.bas` contiene los módulos estándar.
- **UI separada:** `.form.txt` y `.report.txt` aportan layouts, controles y propiedades, nunca procedimientos duplicados del código incrustado.
- **Evidencia estática:** CodeGraph indexa texto exportado. Dysflow gestiona el ciclo con el binario Access; el grafo no demuestra ejecución ni sincronización del `.accdb`.
- **Inferencias explícitas:** las relaciones heurísticas conservan su procedencia; una relación ausente no prueba ausencia de efectos en ejecución.

Las pruebas del [extractor UI](__tests__/extraction-vba-form.test.ts) y del [ejemplo documentado](__tests__/vba-documented-example.test.ts) respaldan esta separación.

## Sincronización con upstream

`origin` corresponde a `ardelperal/codegraph-vba`; `upstream`, a `colbymchenry/codegraph`.

La base integrada es upstream `v1.4.1`, no upstream `v1.6.0`. Las dos líneas de versión son independientes; la planificación de un merge no equivale a código integrado.

Una sincronización debe identificar el tag y commit upstream objetivo, medir la divergencia y preservar la identidad y las pruebas VBA/Access del fork.

## Comprobación antes de actualizar estas notas

- [ ] Verificar versión y ascendencia en Git antes de afirmar que upstream está integrado.
- [ ] Mantener los contratos detallados en el README y la página de capacidades, con enlaces a pruebas.

## Navegación

[Capacidades del fork](docs/fork-capabilities.md) · [Instalación](README.md#get-started) · [Historial](CHANGELOG.md)
