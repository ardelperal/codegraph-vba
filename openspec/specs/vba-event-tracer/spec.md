# vba-event-tracer

## Purpose

Specifies the extraction of VBA event signatures, the query-time dynamic resolution of `WithEvents` event handlers, and the execution schema of the `vba-event-tracer` skill.

## Requirements

### Requirement: Event Signature Extraction

The extractor MUST populate the `signature` column on `event` nodes with the trimmed source declaration line.

#### Scenario: Event signature parsing

- GIVEN a class module `PedidoPublisher.cls` containing:
  `Public Event PedidoGuardado(ByVal id As Long, ByVal total As Double)`
- WHEN the extractor processes the class module
- THEN it MUST create an `event` node named `PedidoGuardado`
- AND set `signature` to `"Public Event PedidoGuardado(ByVal id As Long, ByVal total As Double)"`

---

### Requirement: Dynamic WithEvents Handler Resolution

At query time, the system MUST resolve handlers by matching `<VariableName>_<EventName>` procedures within subscriber modules that hold `subscribes-event` edges targeting the event's class.
Circular references MUST NOT cause infinite recursion.
Unqualified queries with multiple matching event names MUST return `EVENT_AMBIGUOUS`.

#### Scenario: Happy path WithEvents resolution

- GIVEN an `event` node `PedidoGuardado` in class `PedidoPublisher`
- AND a subscriber class `PedidoSubscriber` with a `subscribes-event` edge to `PedidoPublisher` (metadata `variableName` is `m_Publisher`)
- AND `PedidoSubscriber` contains a function node named `m_Publisher_PedidoGuardado`
- WHEN the query engine resolves handlers for event `PedidoPublisher.PedidoGuardado`
- THEN it MUST return the procedure `PedidoSubscriber.m_Publisher_PedidoGuardado`

#### Scenario: Circular reference resolution

- GIVEN `ClassA` contains a `WithEvents` variable subscribing to `ClassB`
- AND `ClassB` contains a `WithEvents` variable subscribing to `ClassA`
- WHEN the query engine traces events and handlers
- THEN the resolution MUST terminate without infinite recursion loops

#### Scenario: Ambiguity handling

- GIVEN class `PublisherA` and class `PublisherB` both define an event named `DataChanged`
- WHEN querying handlers for `DataChanged` without a class prefix
- THEN the system MUST return `EVENT_AMBIGUOUS` with candidate qualifiers `PublisherA.DataChanged` and `PublisherB.DataChanged`

---

### Requirement: vba-event-tracer Skill Execution

The `vba-event-tracer` skill MUST query the graph to trace event declarations, raise sites, and handlers under 100ms, returning the trace JSON schema.

#### Scenario: Tracer skill execution and schema

- GIVEN a trace request for event `PedidoPublisher.PedidoGuardado`
- WHEN the `vba-event-tracer` skill executes
- THEN it MUST return a JSON object with:
  - `event_declarations`: array of matching events (`module`, `line`, `signature`).
  - `raise_sites`: array of calling sites (`module`, `line`, `context`).
  - `handlers`: array of subscriber procedures (`form` or `module`, `handler`, `via`).
  - `warnings`: list of warnings (e.g. `NO_RAISERS` if the event is never raised).
- AND finish execution in less than 100ms.
