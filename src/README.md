---
title: Markdown Styler for MS Word
author: poifuture
homepage: https://github.com/poifuture/markdown-styler-for-word
---

<!-- prettier-ignore-start -->
<!-- markdownlint-disable -->
<!-- DO NOT FORMAT. This file is used to teach people how to use prettier in MS Word, so we keep exactly whatever it looks. -->

# Markdown Styler for MS Word

> Make MS Word a markdown friendly collaborative editor.

Welcome to markdown world!
This MS Word add-in aims to apply styles to your markdown document without changing your content.
You can easily view your document with a better look while real-time collaborating with others.
Our team is actively using it for writing our meeting notes.

<!-- INSTALL SECTION BEGIN  -->

## Install

Open MS Word Online -> Insert" -> "Office Add-ins" -> "Store -> Search "Markdown Styler" -> Add

If the add-in is successfully installed, the add-in will appear in the "Home" tab with a "Remark Selection" button.

<!-- INSTALL SECTION END -->

## Usage

1. Carefully read the Readme and Warning before using it (You can find Readme at Home Tab -> Markdown Styler Menu -> Insert Readme)
1. (Optional) (Not Implemented) Setup the document theme for a better style (e.g. monospace font)
1. Click "Remark Document" (it will style the whole document)
1. (Optional) Customize the built-in styles (Normal, Heading1, etc.) of the document theme in MS Word

## Warning

There might be unexpected changes happen. If any content is missing, try the "History" feature of MS Word. (Open folder in OneDrive Online -> Right click the file -> Version history)

## Why MS Word Online

* Good integration with MS products family and **Office Enterprise accounts**
* Acceptable by traditional companies
* Real-time collaborative editing (buggy but usable)
* Version history (extremely buggy comparing to ...)
* China friendly
* Free! (For personal use) (For our team: We paid Office 365!)
* ~~Rich functionality~~ Rich bugs

## What "Remark Document" & "Remark Selection" does

1. Clear all existing styles
1. Format your document with [Prettier](https://github.com/prettier/prettier)
    1. Prettier will format your markdown
    1. Prettier will format your front matter
    1. [Not Implemented] Prettier will format your code block
1. Parse your markdown styles with [Remark](https://github.com/remarkjs/remark)
1. [Not Implemented] ~~Apply syntax highlights to your code block with [Highlight.js](https://github.com/highlightjs/highlight.js/)~~
1. [Not Implemented] ~~Watch live changes and apply style after typing Enter~~

## What "Setup Theme" does {#theme}

**Nothing**

Due to the limitation of Word, it's not possible to modify the template through the add-in. You'll have to manually apply a markdown friendly template by the following steps.

1. Open your document in Desktop Word (Edit in Word)
1. In Desktop Work, search for "templates" and click "Add-ins" action
1. In "Templates and Add-ins" window, Attach `https://poifuture.github.io/markdown-styler-for-word/Markdown.dotx`. (Note: you can download and modify the template to fit your needs)
1. Check "Automatically update document styles" and click "OK"

* [Not Implemented] ~~Change the theme of your document~~
  - ~~Font: Courier New (A monospace font)~~
  - ~~Size: 10 (To make each line contains >=80 chars)~~

## Examples

### Long Paragraph

A long paragraph will be rewrapped at column 80 if the `prettier.proseWrap` is configured as `always`.
If there is no empty line, the two lines will be merged in markdown.
So please always remember to insert an empty line between your paragraphs.

### Headings

### A **Strong** Title

There is a **strong** word and **some phrases** in a sentence.

### List

### Table

Column 1 | Column 2 has a long head | c3 | c4
--- | --- | --- | ---
c1 | c2 | Column 3 is long | c4

### Code

```javascript
  const a=1
```

## Known Issues

### Inline Style

The inline style doesn't work at all in both Word Online and Word Desktop because of different Microsoft issues. In Word Online (Chrome), the inline style will sometimes apply to the entire paragraph after the targeted range. See [bug](https://github.com/OfficeDev/office-js/issues/586). In on-premise Word (Word Desktop), the javascript engine is IE, which lacks the functionality of choosing the exact range.

### Whitespaces

As every web UI developer knows, a normal space (0x20, ascii 32) is different from a display space (0xA0, ascii 160, also known as nbsp, non-breaking space, hard space, etc.). As a workaround, this Add-in will replace all nbsp to space before processing, and put nbsp back in the document. It works fine for most cases. However, in rare scenarios, you may accidentally get nbsp in your document. So... be careful.

### MS Word doesn't have a vim plugin

So sad...

## FAQ

> When will "not implemented" become "implemented"?

When we get 100 [github stars](https://github.com/poifuture/word-add-in-markdown-style)

> When will Google doc come true?

If we get can an average rating over 4/5.

> Alternatives?

Try [SlackEdit](https://stackedit.io/) if you prefer an independent app!

## Appreciation

This tool can't be real without the awesome work of [Remark](https://github.com/remarkjs/remark), [Prettier][Prettier] and [MSOffice][MSOffice]

## Contributing?

Warm Welcome!

[MSOffice]: https://github.com/OfficeDev/office-js
[Prettier]: https://github.com/prettier/prettier

<!-- prettier-ignore-end -->
