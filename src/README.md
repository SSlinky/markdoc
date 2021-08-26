# Task List

<https://spec.commonmark.org/0.30/>

## Lexer

[x] Class

### Markdown Parsing

[ ] Headings
[ ] Paragraph
[ ] Line Breaks
[ ] Emphasis
[ ] Block Quotes
[ ] Ordered Lists
[ ] Unordered Lists
[ ] Inline Code
[ ] Links
[ ] Images

### Extended Markdown Parsing

[ ] Table
[ ] Code Block

### Block Section

Container for block (paragraph) text and styling information. A block will contain at least one InlineSection. This container is for logically separating out paragraphs in the document.

[x] Class

### Inline Section

Container for inline text and styling information.

[x] Class

### Table Section

[ ] Class

## DocumentWriter

This class is responsible for writing the document using the text and styling information in the BlockSection objects.

[ ] Class

## Utilities

[ ] String cleaner to remove or replace non printable characters.
