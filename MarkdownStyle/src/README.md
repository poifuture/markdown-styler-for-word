---
title: Markdown Style for MS Word
author: poifuture
---

<!-- prettier-ignore-start -->
<!-- markdownlint-disable -->
<!-- DO NOT FORMAT. This file is used to teach people how to use prettier in MS Word, so we keep exactly whatever it looks. -->

# Markdown Style for MS Word

> Make MS Word a markdown friendly collaborative editor.

Welcome to markdown world!
This MS Word add-in aims to apply MS Word styles to your document without changing your markdown content.
You can easily view your document with a better style while collaborating with others on the document.
Our team is actively using it for writing our meeting notes.

<!-- INSTALL SECTION BEGIN  -->

## Install

Open MS Word Online -> Insert -> Office Add-ins -> Store -> Search "Markdown Style" -> Add

<!-- INSTALL SECTION END -->

## Usage

1. Insert Readme and read the warning
1. (Optional) Setup the document theme
1. Click "Remark Document"
1. (Optional) Customize the built-in styles (Normal, Heading1, etc.) of the document theme in MS Word

## Warning

We aim to apply only styles to your document without changing your content. However, your work might be lost if there are bugs in the add-in. If it happens, please remember to use the document history feature of MS Word to retrieve your work.

## Why MS Word (Online)

* Chinese friends cannot access Google Doc easily
* Good integration with MS products family
* Free! (For personal use) (From developer: We paid Office 365)
* ~~Rich functionality~~ Buggy

## What it does

1. Clear all pre-existing styles
1. Format your document with [Prettier](https://github.com/prettier/prettier)
   1. Prettier will format your markdown
   1. [Not Implemented] Prettier will format your front matter
   1. [Not Implemented] Prettier will format your code block
1. Parse your markdown styles with [Remark](https://github.com/remarkjs/remark)
1. [Not Implemented] Apply syntax highlights to your code block with [Highlight.js](https://github.com/highlightjs/highlight.js/)
1. [Not Implemented] Watch live changes and apply style after typing Enter

## What setup does

* [Not Implemented] Change the theme font of your document
  - Face: Courier New (A monospace font)
  - Size: 10 (To make each line contains >=80 chars)

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

Sometimes the inline style suddenly apply to the entire paragraph, this is a [bug](https://github.com/OfficeDev/office-js/issues/586) in Word Online. The workaround is not to remark the end of file.

### Whitespcaces

As every web UI developer knows, a normal space (0x20) is different from a display space (0xA0, also known as &nbsp;). As a workaround, this Add-in will replace all nbsp to space before processing, and put nbsp back in document. It works fine for most cases, however, in rare scenarios, you will get nbsp in your clipboard. Be careful.

### MS Word doesn't have a vim plugin

So sad...

## FAQ

> When will "not implemented" become "implemented"?

When we get 100 [github stars](https://github.com/poifuture/word-add-in-markdown-style)

> To learn more about Markdown, see

[Daring Fireball](http://daringfireball.net/).

## Contributing?

Warm Welcome!

<!-- prettier-ignore-end -->
