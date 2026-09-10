---
name: docx-document-engineering
description: Read, create, edit, format, audit, convert, and template Microsoft Word DOCX documents using the safest available MCP server, CLI, or programming library. Use for document automation, report generation, template filling, structural inspection, comments, tracked revisions, tables, images, and PDF conversion.
---

# DOCX Document Engineering Skill

## Purpose

Act as a reliable Microsoft Word document engineer. Handle `.docx` files end-to-end:

- Read and extract document content.
- Inspect structure before editing.
- Create new DOCX documents.
- Edit existing DOCX documents while preserving structure and formatting.
- Fill Word templates with dynamic data.
- Insert or update tables, images, headings, lists, and page breaks.
- Work with comments and, where the selected tool supports it, Track Changes.
- Audit documents for placeholders, TODO/DRAFT markers, missing content, and structural problems.
- Convert Markdown/HTML/text to DOCX.
- Convert DOCX to PDF when a compatible conversion engine is available.
- Validate the output after every meaningful mutation.

## Core Principle

**Do not treat DOCX as plain text.**

A Word document is a structured OOXML package containing paragraphs, runs, styles, sections, tables, relationships, media, comments, headers/footers, and other XML parts.

Therefore:

1. Inspect first.
2. Choose the least destructive editing path.
3. Preserve the existing template whenever possible.
4. Make only the requested changes.
5. Validate the resulting DOCX.
6. Report limitations instead of claiming unsupported features worked.

---

# 1. Tool Selection and Routing

Choose tools according to the task.

## A. MCP / AI-agent workflows

### Preferred: `docx-mcp`

Use when the agent needs a Word-specific MCP interface for document lifecycle operations.

Typical capabilities include:

- `create_docx`
- `read_docx`
- `write_docx`
- `append_docx`
- `list_docx`
- `delete_docx`
- `copy_docx`
- merge-field discovery/filling
- content-control discovery
- document properties
- style listing/application
- bullets and numbering
- image insertion and captions

Reference implementation:
`aiexplorations/docx-mcp`

Important:
- Verify the installed version's actual tool list before relying on a feature.
- Do not infer that every Word feature is supported merely because it exists in another DOCX tool.

### Alternative: `python-office-mcp-server`

Use when the workflow needs a broader OOXML/Office-document MCP server or when deterministic inspection, patching, templates, comments, tables, images, auditing, or Track Changes workflows are available.

Reference implementation:
`rcarmo/python-office-mcp-server`

Useful primary workflow:

1. `office_help` for workflow discovery.
2. `office_inspect` / `word_document_map` for structure.
3. `office_template` for template analysis.
4. `office_patch` for targeted changes.
5. `office_comment` for comments.
6. `office_table` for table operations.
7. `office_image` for images.
8. `office_audit` for verification.

For supported mutation paths, prefer:
- `dry_run` when available to preview matches.
- `safe` when a separate output path is appropriate.
- `strict` when all requested targets must match.
- `best_effort` only when partial compatibility is acceptable.

### `office-word-mcp-server`

Use as a reference or fallback for Word-specific MCP workflows, especially where its available operations fit the task.

Reference implementation:
`GongRzhe/Office-Word-MCP-Server`

Caution:
The upstream repository was archived on 3 March 2026. Treat it as a reference/legacy option rather than assuming active maintenance.

---

# 2. CLI and Conversion Routing

## Pandoc

Use Pandoc when the primary operation is format conversion or generating a DOCX from Markdown/HTML.

Typical pattern:

`pandoc input.md -o output.docx --reference-doc=custom-template.docx`

Use a reference DOCX when consistent Word styles are required.

Pandoc is especially suitable for:

- Markdown → DOCX
- HTML → DOCX
- structured report generation
- PDF generation when the installed Pandoc toolchain supports the required PDF engine

Do not use Pandoc as the default editor for an existing complex DOCX when preserving advanced Word-native structure is important.

## `docx2pdf`

Use for DOCX → PDF when the environment has the required Microsoft Word-compatible conversion dependency.

Typical command:

`docx2pdf document.docx`

Before using it:
- verify the operating system and prerequisites;
- confirm the conversion executable/Word installation is available;
- validate the generated PDF.

---

# 3. Python Library Routing

## `python-docx`

Use for custom Python automation and deterministic DOCX manipulation.

Reference implementation:
`python-openxml/python-docx`

Typical use:

```python
from docx import Document

document = Document("input.docx")
document.add_paragraph("New content")
document.save("output.docx")
```

Good for:

- paragraphs
- runs
- headings
- tables
- sections
- headers/footers
- images
- page breaks
- styles
- document creation and modification

Use it when an MCP server is unnecessary or when a custom Python workflow gives better control.

### Important limitation

`python-docx` should not be treated as a complete implementation of every Word-native feature. Advanced OOXML features may require direct XML manipulation or a more specialized tool.

When a requested feature is not exposed by the library:
- inspect the DOCX XML structure;
- use a supported OOXML approach if safe;
- otherwise use a specialized tool;
- never silently downgrade the requested behavior.

---

# 4. Template Engines

## `docxtpl`

Use Python `docxtpl` when the user has a reusable Word template and wants dynamic data inserted through Jinja-style variables.

Best for:

- certificates
- invoices
- letters
- quotations
- forms
- reports
- repeated structured documents

Recommended pattern:

1. Start from an approved `.docx` template.
2. Define placeholders.
3. Validate the placeholder schema.
4. Render with structured data.
5. Save to a new output file.
6. Re-open and validate the generated document.

Do not rebuild a complex approved template from scratch merely to insert data.

## `docxtemplater`

Use Node.js/TypeScript `docxtemplater` when the surrounding application is JavaScript/TypeScript.

It supports template replacement, loops, and conditions. Optional paid modules add capabilities such as images, HTML, tables, charts, QR codes, styling, footnotes, and metadata.

Reference implementation:
`open-xml-templating/docxtemplater`

Do not assume optional/paid modules are available in an environment.

---

# 5. JavaScript DOCX Generation

## `docx` npm package

Use the JavaScript/TypeScript `docx` package when a web or Node.js application must generate DOCX programmatically.

Good for:

- generating reports
- creating structured documents
- headings and paragraphs
- tables
- images
- page layout
- programmatic document generation

Prefer template-based generation when an existing corporate Word template must be preserved.

---

# 6. Standard Workflow

For every DOCX task, follow this workflow unless the user explicitly asks for a different process.

## Step 1 — Identify the operation

Classify the request as one or more of:

- READ
- CREATE
- EDIT
- TEMPLATE
- FORMAT
- COMMENT
- TRACK_CHANGES
- AUDIT
- CONVERT
- BATCH

## Step 2 — Inspect the input

For an existing DOCX, inspect:

- filename and extension
- document properties
- paragraph count
- headings
- tables
- sections
- headers/footers
- images
- comments
- placeholders/content controls
- DRAFT/TODO markers
- obvious formatting inconsistencies

Do not start destructive editing before understanding the document.

## Step 3 — Preserve the source

Unless the user explicitly asks for in-place editing:

- keep the original untouched;
- write to a distinct output path;
- use a predictable filename such as:
  `document_v2.docx`
  `document_review.docx`
  `document_final.docx`

## Step 4 — Apply minimal targeted changes

Prefer:

- placeholder replacement
- targeted paragraph edits
- targeted table edits
- style application
- anchor-based insertion

over recreating the entire document.

## Step 5 — Validate

At minimum verify:

- file exists;
- DOCX can be opened/re-parsed;
- requested text exists;
- unintended old text is absent where replacement was requested;
- table dimensions/content remain correct;
- headings/styles remain sensible;
- no unresolved placeholders remain if the task requires completion.

For high-value documents also inspect:

- section/page layout;
- headers/footers;
- images and captions;
- comments;
- revision state;
- PDF rendering if visual fidelity matters.

## Step 6 — Report accurately

Return:

- output filename/path;
- what was changed;
- validation performed;
- warnings/limitations;
- any human review still required.

Never claim "perfect formatting" without validation.

---

# 7. Editing Existing Word Documents

## Rule: Preserve before replacing

When modifying an existing document:

1. Read it.
2. Inspect its structure.
3. Locate the exact target.
4. Make the smallest safe mutation.
5. Save as a new file.
6. Re-open and verify.

### Search and replace

Before replacement:

- count or identify matches;
- distinguish exact text from similar text;
- use strict matching for legal, financial, or controlled documents;
- use dry-run/preview when available.

If a requested replacement cannot be matched cleanly, stop and report it rather than silently changing something else.

---

# 8. Tables

For tables:

- inspect existing row/column structure first;
- preserve the table style;
- preserve merged cells unless explicitly changing them;
- preserve column widths where possible;
- avoid accidental autofit changes;
- validate header and data rows after editing.

For dynamic tables:

- define the schema first;
- ensure every row has the expected columns;
- preserve number/date formatting;
- check for page-break issues in long tables.

`python-docx` supports creating rows, cells, table styles, alignment, widths, merging, and autofit controls, but complex Word tables may require deeper OOXML handling.

---

# 9. Styles and Formatting

Prefer Word styles over manually formatting every run.

Recommended hierarchy:

- Title
- Subtitle
- Heading 1
- Heading 2
- Heading 3
- Normal
- Caption
- custom corporate styles

When editing a corporate template:

- inspect available styles first;
- reuse existing styles;
- do not introduce new fonts/colors without instruction;
- preserve the template's visual language.

If a style name is unavailable, report the mismatch or map to the closest verified style rather than assuming it exists.

---

# 10. Images

When inserting images:

1. Verify the image exists.
2. Check dimensions/aspect ratio.
3. Insert using the appropriate tool.
4. Set the requested size.
5. Add a caption if required.
6. Re-open the document and confirm the image relationship is valid.

Avoid stretching logos or other identity-sensitive graphics.

---

# 11. Comments

If comments are required:

- identify the target paragraph/content;
- add the comment to the correct location;
- preserve existing comments;
- verify author/date/thread information when supported.

For threaded comment workflows, prefer tools that explicitly support:

- getting comments;
- replies;
- resolving;
- reopening;
- threaded relationships.

Do not claim a comment is attached to an exact visual location if the selected library only provides paragraph-level or XML-level anchoring.

---

# 12. Track Changes

Track Changes is a Word-native revision feature and must not be simulated with ordinary colored text or bracketed annotations.

If the selected tool explicitly supports revision marks:

1. Create a review copy.
2. Enable Track Changes if required.
3. Apply tracked replacements.
4. Verify revisions exist.
5. Keep the original unchanged.
6. Provide the review copy.

If the selected library does not support genuine revision marks:

- do not pretend it does;
- use a supported MCP/OOXML workflow;
- or clearly state that Track Changes cannot be produced with the current path.

For example, `python-office-mcp-server` documents specialized workflows for enabling Track Changes, applying tracked replacements, and accepting all changes.

---

# 13. Template Workflow

For an existing template:

```text
TEMPLATE
  ↓
INSPECT
  ↓
IDENTIFY PLACEHOLDERS / CONTENT CONTROLS / ANCHORS
  ↓
VALIDATE INPUT DATA
  ↓
RENDER / PATCH
  ↓
AUDIT
  ↓
OUTPUT NEW DOCX
```

Recommended tools:

- `docxtpl` for Python/Jinja-style template rendering.
- `docxtemplater` for Node.js/TypeScript template rendering.
- `office_template` / related MCP tooling when structural template analysis is required.
- `python-docx` for custom deterministic modifications.

Never overwrite the master template during normal generation.

---

# 14. Markdown → Word

For simple or structured reports:

```text
Markdown
  ↓
Pandoc / word_from_markdown
  ↓
Reference DOCX
  ↓
DOCX
  ↓
Audit
```

For large Markdown input, prefer a file-based input rather than putting the entire document into an MCP argument when the server supports a `markdown_file` parameter.

---

# 15. DOCX → PDF

Use:

1. `docx2pdf` when Microsoft Word automation is available.
2. LibreOffice/headless conversion when appropriate and available.
3. Other verified document conversion tools when required.

Always distinguish:

- "PDF generated successfully"
from
- "PDF visually matches Word perfectly."

Visual fidelity requires rendering/inspection.

---

# 16. Batch Processing

For batch generation:

1. Validate the template once.
2. Validate all input records.
3. Generate into a dedicated output directory.
4. Use deterministic filenames.
5. Log success/failure per record.
6. Validate each generated DOCX.
7. Produce a summary report.

Example:

```text
records.csv
templates/report.docx
        ↓
validate schema
        ↓
render N documents
        ↓
audit N documents
        ↓
output/
  report_001.docx
  report_002.docx
  report_003.docx
```

Do not let one malformed record silently corrupt the entire batch.

---

# 17. Quality Gates

Use these gates for important documents.

## Gate A — Structural

- DOCX opens successfully.
- Required sections exist.
- Required tables exist.
- No accidental section deletion.
- No broken relationships.

## Gate B — Content

- Required text is present.
- Replacements are complete.
- No unintended duplicate text.
- No unresolved placeholders when completion is required.
- Numbers, dates, names, and identifiers are preserved accurately.

## Gate C — Formatting

- Heading hierarchy is intact.
- Styles are applied correctly.
- Tables retain intended layout.
- Images are present and proportioned.
- Page breaks/sections remain sensible.

## Gate D — Review State

Where applicable:

- comments remain intact;
- requested comments exist;
- Track Changes is genuinely enabled/present;
- unresolved revisions are clearly reported.

## Gate E — Conversion

If PDF is required:

- PDF exists;
- page count is plausible;
- text is extractable;
- important visual elements are present.

---

# 18. Failure Handling

When something fails:

### Partial match

Report:

- requested targets;
- matched targets;
- unmatched targets;
- whether the file was written.

### Formatting limitation

State exactly what could not be preserved or changed.

### Unsupported feature

Do not emulate silently.

Say:

- which feature is unsupported by the selected path;
- which alternative tool/path should be used.

### Corrupt output

Do not return it as a successful document.

Attempt:

1. re-open/re-parse;
2. inspect package/XML if necessary;
3. regenerate from the original;
4. validate again.

---

# 19. Security and File Safety

Treat document manipulation as a file-system operation.

- Restrict access to approved directories.
- Reject path traversal.
- Validate file extensions.
- Apply file-size limits.
- Never expose secrets embedded in environment variables or configuration.
- Avoid overwriting source documents unless explicitly requested.
- Sanitize generated filenames.
- Do not execute macros or embedded executable content as part of normal document processing.

For MCP servers, prefer implementations with explicit path validation and file-size controls.

---

# 20. Recommended Decision Matrix

| Task | First Choice | Alternative |
|---|---|---|
| Read DOCX | `docx-mcp` / `office_read` | `python-docx` |
| Inspect structure | `office_inspect` / `word_document_map` | `python-docx` |
| Create simple DOCX | `python-docx` | `docx-mcp` |
| Create from Markdown | Pandoc | `word_from_markdown` |
| Fill Word template | `docxtpl` | `docxtemplater` |
| Node.js DOCX generation | `docx` | `docxtemplater` |
| Targeted Word patch | `office_patch` | `python-docx` |
| Comments | `office_comment` | specialized OOXML tooling |
| Track Changes | specialized MCP/OOXML path | Word automation |
| Batch documents | `docxtpl` / `docxtemplater` | custom Python/Node |
| DOCX → PDF | `docx2pdf` | LibreOffice |
| Audit | `office_audit` | custom parser/checker |

---

# 21. Operating Prompt

When using this skill, reason internally using:

```text
INPUT
→ What document(s) are involved?
→ What exact operation is requested?
→ What must be preserved?
→ What tool has the strongest verified support?

INSPECT
→ Read content
→ Inspect structure
→ Locate targets
→ Identify risks

EXECUTE
→ Use the least destructive operation
→ Preserve source
→ Write to a new output where practical

VERIFY
→ Re-open
→ Check content
→ Check structure
→ Check formatting
→ Check review state
→ Check conversion if applicable

REPORT
→ Output
→ Changes
→ Validation
→ Limitations
→ Human review required
```

---

# 22. Example Natural-Language Requests

The skill should understand requests such as:

- "Read this Word document and summarize its structure."
- "Create a professional report in Word from this Markdown."
- "Use this DOCX as a template and fill in the participant information."
- "Replace every occurrence of `<CLIENT_NAME>` with the supplied company name."
- "Add a 4×4 table after the Executive Summary."
- "Insert this logo on the title page."
- "Find all TODO and DRAFT markers."
- "Add a comment explaining this clause."
- "Make these changes using Track Changes."
- "Generate 100 certificates from this Word template."
- "Convert the final Word document to PDF."
- "Audit the final DOCX and tell me whether anything was missed."

---

# References Used

The workflow and tool-routing model are informed by these public projects/documentation:

- `aiexplorations/docx-mcp` — MCP lifecycle operations, styles, lists, images, metadata, path/file safety, and document tooling.
- `GongRzhe/Office-Word-MCP-Server` — Word MCP architecture and operations; note that the repository is archived as of 3 March 2026.
- `rcarmo/python-office-mcp-server` — core-first Office workflow, inspection, patching, templates, comments, auditing, tables, images, Track Changes, and explicit execution modes.
- `python-openxml/python-docx` — Python DOCX creation/modification, tables, styles, sections, headers/footers, comments, and document APIs.
- `jgm/pandoc` — document format conversion and DOCX generation from Markdown/HTML.
- `open-xml-templating/docxtemplater` — DOCX/PPTX templating, placeholders, loops, and conditions.

This skill intentionally treats tool capabilities as version-dependent. Before invoking a specific command or MCP operation, inspect the installed tool's current schema/documentation and verify that the requested feature is actually supported.
