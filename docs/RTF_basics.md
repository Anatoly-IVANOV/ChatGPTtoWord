# RTF basics

The Rich Text Format (RTF) is a proprietary document file format developed by Microsoft from 1987 until 2008 for cross-platform document interchange with Microsoft products.


## Basic RTF Structure

An RTF document starts and ends with curly braces `{}`, which encapsulate the entire document's content. Inside the braces, various control words define text formatting, styles, and other document features in a plain-text format that can be interpreted by various word processors.

### Example of an RTF Document

An RTF file consists of unformatted text, control words, control symbols, and groups.

```rtf
{\rtf1\ansi\deff0
   {\fonttbl{\f0 Times New Roman;}}
   {\colortbl;\red255\green0\blue0;}
   \s0 This is a normal paragraph.
   \s1\ul This is an underlined paragraph.
   \par This is another paragraph.
   \s2\cf1 This text is colored.
}
```

In this example:

- `\rtf1\ansi\deff0`: Specifies RTF version, character set, and default font.
- `\fonttbl`: Defines fonts used.
- `\colortbl`: Defines colors.

## Control words

A control word is a specially formatted command that RTF uses to mark printer control codes and information that applications use to manage documents, which follows the form:

```rtf
\LetterSequence<Delimiter>
```

Note that a backslash begins each control word.

The LetterSequence is made up of lowercase alphabetic characters between "a" and "z" inclusive. RTF is case sensitive, and all RTF control words must be lowercase.

The delimiter marks the end of an RTF control word, and can be one of the following:

* A space. In this case, the space is part of the control word.
* A digit or a hyphen (-), which indicates that a numeric parameter follows. The subsequent digital sequence is then delimited by a space or any character other than a letter or a digit. If a numeric parameter immediately follows the control word, this parameter becomes part of the control word.
* The control word is then delimited by a space or a nonalphabetic or nonnumeric character in the same manner as any other control word.
* Any character other than a letter or a digit. In this case, the delimiting character terminates the control word but is not actually part of the control word.
   * If a space delimits the control word, the space does not appear in the document. Any characters following the delimiter, including spaces, will appear in the document. For this reason, you should use spaces only where necessary; do not use spaces merely to break up RTF code.

## Common Control Words

### `\par` (Paragraph Break)

- Defines the end of a paragraph

  ```rtf
  This is a paragraph.\par
  This is another paragraph.
  ```

### `\sN` (Paragraph Style)

- Applies a specific paragraph style (defined earlier in the document).
- The number following `\s` corresponds to a style definition (usually set in a `\stylesheet` section).

  ```rtf
  \s0 This paragraph uses style 0.
  \s1 This paragraph uses style 1.
  ```

### `\ul` (Underline)

- Underlines text.

  ```rtf
  \ul This text is underlined.
  ```

### `\cf` (Color Formatting)

- Sets the color of text using a reference to the color table (defined by `\colortbl`).

  ```rtf
  \cf1 This text is red.
  ```

### `\b` (Bold)

- Makes the text bold.

  ```rtf
  \b This text is bold.
  ```

### `\i` (Italic)

- Italicizes the text.

  ```rtf
  \i This text is italicized.
  ```

## Document Components

1. **Document Header**:
   - Starts with `{\rtf1\ansi` indicating it is an RTF 1.0 document with ANSI character set.
   - **Example**:
     ```rtf
     {\rtf1\ansi
     ```

2. **Font Table** (`\fonttbl`)**:
   - Defines the fonts used in the document.
   - **Example**:
     ```rtf
     {\fonttbl{\f0 Times New Roman;}}
     ```

3. **Color Table** (`\colortbl`)**:
   - Defines the colors used in the document.
   - **Example**:
     ```rtf
     {\colortbl;\red255\green0\blue0;}
     ```

4. **Sectioning** (`\sect` / `\sectd`)**:
   - RTF also supports sections, and you can separate sections using the `\sect` control word.
   - **Example**:
     ```rtf
     \sect This is the beginning of a new section.
     ```

5. **Style Sheet** (`\stylesheet`)**:
   - Defines paragraph styles (`\s`) used throughout the document.
   - **Example**:
     ```rtf
     {\stylesheet{\s0 Normal;}{\s1 Heading;}}
     ```

### Example in Action:

```rtf
{\rtf1\ansi\deff0
  {\fonttbl{\f0 Times New Roman;}}
  {\colortbl;\red0\green0\blue255;}

  \s0 This is a paragraph with style 0.\par
  \s1\b This is a bold paragraph with style 1.\par
  \s1\i\cf1 This is an italic and colored paragraph with style 1.\par
  \s2\ul This is an underlined paragraph with style 2.\par
}
```

This example shows how fonts, colors, paragraph styles, and formatting (like bold, italic, underline) are applied in RTF. The curly braces encapsulate sections of the document, and the backslashes introduce the control words.

### Control Word Summary:

- **`\par`**: Marks a paragraph break.
- **`\sN`**: Applies paragraph style N.
- **`\b`**: Bold text.
- **`\i`**: Italic text.
- **`\ul`**: Underlined text.
- **`\cfN`**: Changes text color using a color table index.
- **`\fonttbl`**: Defines available fonts.
- **`\colortbl`**: Defines available colors.

## ChatGPT peculiarities

I suppose to retain correct DOM hierarchy, ChatGPT output uses `###` for what is actually `h1` in a Word file.

- **`h2`**: is `Chat History`

I can’t find the `h1` element in their DOM. Let me know if you do. 🙃
