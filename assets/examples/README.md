# examples

This directory stores document examples used as style and structure references.

Important rules:

- Public examples are for format, structure, tone, and wording patterns only.
- Public examples must not be treated as company facts or copied into company output.
- Internal examples can stay empty at first. Add them only after confirming they are safe to reuse.
- Each example should have metadata: department, document_type, source, use_for, avoid_learning, and notes.

Suggested layers:

```text
examples/
  public/    Public official references and style cards.
  internal/  Company-approved examples, initially empty.
```

When a skill needs examples, load by:

```text
department + document_type + task
```

Example:

```text
admin + notice + polish_docx
marketing + customer_solution + generate_pptx
product + release_note + polish_docx
```
