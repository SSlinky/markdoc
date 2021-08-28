# Task List

<https://spec.commonmark.org/0.30/>

## Lexer

1. ~~Class~~

### Markdown Parsing

1. Container Blocks
   1. Block Quotes
   1. List Items.
   1. Lists.
1. Leaf Blocks.
   1. Thematic Breaks.
   1. ATX Headings.
   1. Setext Headings.
   1. Indented Code Blocks.
   1. Fenced Code Blocks.
   1. HTML Blocks (implement last if at all).
   1. Link Reference Definitions.
   1. Paragraphs.
   1. Blank Lines.
1. Inlines
   1. Code Spans.
   1. Emphasis and String Emphasis.
   1. Links.
   1. Images.
   1. Autolinks.
   1. Raw HTML (implement last if at all).
   1. Hard line breaks.
   1. Soft line breaks.
   1. Textual content.
1. Word Extended
   1. Layout.
   1. Margins.
   1. Page Breaks.
   1. Columns.
   1. Header/Footer.
   1. Watermark.
   1. Page Colour.

### Block Section

Container for block (paragraph) text and styling information. A block will contain at least one InlineSection. This container is for logically separating out paragraphs in the document.

1. ~~Class~~

### Inline Section

Container for inline text and styling information.

1. ~~Class~~

### Table Section

1. Class

## DocumentWriter

This class is responsible for writing the document using the text and styling information in the BlockSection objects.

1. Class

## Utilities

1. String cleaner to remove or replace non printable characters.
1. ~~File reader.~~
1. ~~Collection wrapper for pushing and popping.~~
1. ~~Basic logger class.~~
