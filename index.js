/*******************************************************************************
 **
 **  ANATOLY IVANOV / DESIGN
 **
 **  https://anatolyivanov.com/
 **
 **  Copyright (c) Anatoly IVANOV .com, all rights reserved - subject to and
 **  governed by French and international intellectual property law
 **
 **
 **  RTF DOCUMENTATION:   docs/RTF_basics.md
 **
 *******************************************************************************/




/*******************************************************************************
 **
 **    LOCAL CONFIG
 **
 ******************************************************************************/

// Debug (prints Markdown before conversion and RTF as copied into clipboard)
const bool_Debug = false;


// RTF header
const RTF_HEADER = '{\\rtf1\\ansi\\uc1\n';

// RTF fonts
// - bold and monospaced will show up as `Direct Formatting` in Word
const RTF_FONT_TABLE = '{\\fonttbl\n' +
  '{\\f0 Myriad;}\n' + // Main font
  '{\\f1 Myriad Pro Semibold;}\n' + // Bold sections
  '{\\f2 Consolas;}\n' + // Monospaced sections
  '}\n';

// RTF colors (RGB)
// - bold and monospaced will show up as `Direct Formatting` in Word
const RTF_COLOR_TABLE = '{\\colortbl\n' +
  '{;}\n' + // Default empty color (required to simulated "automatic" in Word)
  '{\\red49\\green132\\blue155;}\n' + // Bold character style color
  '{\\red0\\green112\\blue192;}\n' + // Monospaced character style color
  '}\n';

// RTF Style Definitions
// - MUST match Word styles (defined in advance); will not be `Direct Formatting`
const RTF_STYLESHEET = '{\\stylesheet\n' +
  '{\\s0 Normal;}\n' +
  '{\\s1 Heading 1;}\n' +
  '{\\s2 Heading 2;}\n' +
  '{\\s3 Heading 3;}\n' +
  '{\\s4 Heading 4;}\n' +
  '{\\s5 Heading 5;}\n' +
  '{\\s6 Bullets level 1;}\n' +
  '{\\s7 Bullets level 2;}\n' +
  '}\n';

// Markdown symbology to Word styles map -- headings
const obj_HeadingStyles = {
  '#': '\\s1', // Heading 1
  '##': '\\s1', // Heading 1
  '###': '\\s1', // Heading 1 (starting point)
  '####': '\\s2', // Heading 2
};

// Markdown symbology to Word styles map -- lists
const obj_ListStyles = {
  '-': '\\s6', // Bullets level 1
  '--': '\\s7', // Bullets level 2
};




/*******************************************************************************
 **
 **    IMPORT MODULES -- NPM -- GLOBAL
 **
 ******************************************************************************/

import clipboard from 'clipboardy';




/*******************************************************************************
 **
 **    UTILITY FUNCTIONS
 **
 ******************************************************************************/

/**
 * Returns a colored text string for output to the console.
 *
 * @param {string} str_Text - The text to color.
 * @param {string} str_ConsoleColorCode - The color code name ('blue', 'red'…)
 * @return {string} str_Text - Wrapped in the specified color code, or unmodified
 *                             if color is invalid.
 */
function console_color_conv ( str_Text, str_ConsoleColorCode ) {

  // ANSI escape colors
  const obj_ConsoleColorCodes = {
    black: '\x1b[30m',
    red: '\x1b[31m',
    green: '\x1b[32m',
    yellow: '\x1b[33m',
    blue: '\x1b[34m',
    magenta: '\x1b[35m',
    cyan: '\x1b[36m',
    white: '\x1b[37m'
  };

  /**
   * Check if:
   *
   * - str_ConsoleColorCode is defined
   * - if it is defined, convert the case of str_ConsoleColorCode to lowercase…
   *   just in case :p (so Blue or BLUE will both work)
   * - check whether the lower case str_ConsoleColorCode matches any key in
   *   obj_ConsoleColorCodes, prepend the correct ANSI escape chars, append the
   *   reset escape code, return the result
   * - otherwise, fallback to return str_Text unmodified
   */
  return str_ConsoleColorCode ? `${obj_ConsoleColorCodes[str_ConsoleColorCode.toLowerCase()]}${str_Text}\x1b[0m` : str_Text;

}


async function catch_message_print_exit ( obj_Error = {}, str_ErrorMessage = '', str_SolutionMessage = '', bool_Exit = false ) {

  // If a specific human-readable error message is provided
  if ( str_ErrorMessage ) {
    console.log( `- ${str_ErrorMessage }` );
  }

  // If a possible solution is provided
  if ( str_SolutionMessage ) {
    console.log( `- ${str_SolutionMessage}` );
  }

  // Print out the Node.js error
  console.log( `- ${obj_Error}` );

  // If exit is requested
  if ( bool_Exit === true ) {

    // Exit the Node.js process and its children
    console.log( `Exiting... (shutting down Node.js process)` );
    console.log(); // a spacer for legibility
    process.exit();
  }

}



/**
 * Escapes RTF special characters and handles Unicode.
 */
function RTF_special_chars_escape ( str_Text ) {
  return str_Text
    .replace( /\\/g, '\\\\' ) // Backslash
    .replace( /{/g, '\\{' ) // Left brace
    .replace( /}/g, '\\}' ) // Right brace
    .replace( /[\u0080-\uFFFF]/g, ( match ) => {
      return `\\u${match.charCodeAt(0)}?`; // Unicode characters
    } );
}



/**
 * Converts Title Case into Sentence case, except when it's an acronym
 *
 * @param {string} str_Text
 * @return {string} Sentence case of str_Text
 */
function title_case_to_sentence_case_conv ( str_Text ) {

  // Split the string into words using the space as a separator
  const arr_WordsOriginal = str_Text.split( ' ' );

  // Iterate through each word
  const arr_WordsTransformed = arr_WordsOriginal.map( ( str_Word, int_Index ) => {

    // Check if the word is fully uppercase (acronym)
    if ( str_Word === str_Word.toUpperCase() && str_Word.length > 1 ) {
      return str_Word; // Return acronyms as is, without transformation
    }

    // For the first word, apply Sentence case logic
    if ( int_Index === 0 ) {
      return str_Word.charAt( 0 ).toLocaleUpperCase() + str_Word.slice( 1 ).toLocaleLowerCase();
    }

    // For other words, convert to lowercase (if not an acronym)
    return str_Word.toLocaleLowerCase();

  } );

  // Join words back into a single string with spaces
  return arr_WordsTransformed.join( ' ' );
}



/**
 * Replaces inline Markdown formatting with RTF control words.
 *
 * @param {string} str_Text - The text to format.
 * @param {boolean} bool_IsHeading - Whether the text is inside a heading.
 * @return {string} - The formatted text.
 */
function replace_inline_formatting ( str_Text, bool_IsHeading = false ) {

  // Escape special RTF characters first
  // (otherwise wreaks havoc on output, reason unknown)
  str_Text = RTF_special_chars_escape( str_Text );

  // Handle bold '**text**' using RTF bold control words, but only if not in heading
  if ( !bool_IsHeading ) {

    str_Text = str_Text.replace( /\*\*(.+?)\*\*/g, ( match, p1 ) => {
      // Replace bold markers with color and typography
      // NB: trailing space required for correct Word import
      // (otherwise the control word is ignored -- reason unknown)
      return `\\cf1\\f1\\b1 ${p1}\\b0\\f0\\cf0 `;
    } );

  }

  else {

    // In headings with bold markers, remove them
    // (Word styles already define the proper font weight)
    str_Text = str_Text.replace( /\*\*(.+?)\*\*/g, "$1" );

    // In all cases, convert to Sentence case (because I use EN + FR + RU)
    str_Text = title_case_to_sentence_case_conv( str_Text );

    return str_Text;

  }

  // Search for inline code '`code`'
  str_Text = str_Text.replace( /`(.+?)`/g, ( match, p1 ) => {

    // If matched, apply RTF monospaced color and font, then revert to defaults
    // NB: trailing space required for correct Word import
    // (otherwise the control word is ignored -- reason unknown)
    return `\\cf2\\f2 ${p1}\\f0\\cf0 `;
  } );

  // Search for inline italics '*text*'
  str_Text = str_Text.replace( /\*(.+?)\*/g, ( match, p1 ) => {

    // If matched, apply RTF italics, then revert to defaults
    // NB: trailing space required for correct Word import
    // (otherwise the control word is ignored -- reason unknown)
    return `\\i ${p1}\\i0 `;
  } );

  return str_Text;
}




/*******************************************************************************
 **
 **    MAIN CONVERSION LOGIC
 **
 ******************************************************************************/

/**
 * Converts Markdown content to RTF format.
 * @param {string} markdown - The Markdown content.
 * @return {string} - The RTF formatted content.
 *
 * NB: An RTF paragraph starts with \sN and ends with \par
 */
function transform_markdown_to_rtf ( markdown ) {
  let str_RTFcontent = RTF_HEADER;


  /* --------------------------------------------
   * Add font table, color table, and stylesheet
   * -------------------------------------------- */
  str_RTFcontent += RTF_FONT_TABLE;
  str_RTFcontent += RTF_COLOR_TABLE;
  str_RTFcontent += RTF_STYLESHEET;

  // Start the document without additional formatting
  str_RTFcontent += '\n';

  // Parse the Markdown content
  const arr_Lines = markdown.split( '\n' );

  // Iterate through lines of text
  for ( let str_Line of arr_Lines ) {

    // Trim whitespace
    str_Line = str_Line.trim();

    // Skip empty lines
    if ( str_Line === '' ) {
      continue;
    }

    // Init the RTF line-by-line var
    let str_RTFLine = '';

    // Keep track of formatting applied -- set the flag to false by default
    let bool_FormattingApplied = false;


    /* --------------------------------------------
     * Check for headings
     * -------------------------------------------- */

    // Iterate through the Markdown symbology to Word styles map in the config
    for ( const [ str_MDprefix, str_RTFstyle ] of Object.entries( obj_HeadingStyles ) ) {

      // For each line in the config object, check whether it matches
      if ( str_Line.startsWith( str_MDprefix + ' ' ) ) {

        // Extract the contents of the heading
        let str_Content = str_Line.substring( str_MDprefix.length + 1 ).trim();

        // Remove numbering prefixes and colons
        str_Content = str_Content.replace( /^(\d+[\.\)]\s*)?(\w[\.\)]\s*)?/, '' ).trim();

        // Remove trailing colon if present (Some text:)
        str_Content = str_Content.replace( /:$/, '' ).trim();

        // Remove markdown (**Some text**)
        str_Content = replace_inline_formatting( str_Content, true );

        // Close the paragraph
        str_RTFLine += `${str_RTFstyle} ${str_Content}\\par \n`;

        // Set the flag as formatting applied
        bool_FormattingApplied = true;

        // Get out of the loop
        break;
      }
    }

    /* --------------------------------------------
     * Check for list items with bold text and colon
     * Example: - **Some Text**: description
     * -------------------------------------------- */
    if ( !bool_FormattingApplied ) {

      // Looking for `- **Some Text**` to be later split into h3 and bullets
      const arr_RegExMatch_BoldStart = str_Line.match( /^- \*\*(.+?)\*\*:\s*(.*)/ );

      // OK, we have a match
      if ( arr_RegExMatch_BoldStart ) {

        // Populate the strings for the h3 and bullets
        let str_H3text = arr_RegExMatch_BoldStart[ 1 ];
        let str_BulletText = arr_RegExMatch_BoldStart[ 2 ];

        // Remove bold markers in heading (bool_IsHeading = true)
        str_H3text = replace_inline_formatting( str_H3text, true );

        // Format bullet text (bool_IsHeading = false)
        str_BulletText = replace_inline_formatting( str_BulletText, false );

        // Append Heading 3 -- using **Some text** for content
        str_RTFLine += `\\s3 ${str_H3text}\\par \n`;

        // Append Bullets level 1 -- using text following **Some text** for content
        str_RTFLine += `\\s6 ${str_BulletText}\\par \n`;

        // Set the flag as formatting applied
        bool_FormattingApplied = true;
      }
    }

    /* --------------------------------------------
     * Check for list items
     * -------------------------------------------- */
    if ( !bool_FormattingApplied ) {

      // Iterate through the Markdown symbology to Word styles map in the config
      for ( const [ str_MDprefix, str_RTFstyle ] of Object.entries( obj_ListStyles ) ) {

        // For each line in the config object, check whether it matches
        if ( str_Line.startsWith( str_MDprefix + ' ' ) ) {

          // Populate the strings for the bullets
          let str_Content = str_Line.substring( str_MDprefix.length + 1 ).trim();

          // Format bullet text (bool_IsHeading = false)
          str_Content = replace_inline_formatting( str_Content, false ); // bool_IsHeading = false

          // Append the bullet text
          str_RTFLine += `${str_RTFstyle} ${str_Content}\\par \n`;

          // Set the flag as formatting applied
          bool_FormattingApplied = true;

          // Get out of the loop
          break;
        }
      }
    }

    /* --------------------------------------------
     * Regular paragraph
     * -------------------------------------------- */

    // The least complicated, just format and append
    if ( !bool_FormattingApplied ) {
      let str_Content = replace_inline_formatting( str_Line, false );
      str_RTFLine += `\\s0 ${str_Content}\\par \n`;
    }

    // Add the converted RTF line to... the RTF
    str_RTFcontent += str_RTFLine + '\n';
  }

  // Close the RTF document
  str_RTFcontent += '}';

  return str_RTFcontent;
}




/*******************************************************************************
 **
 **    READ AND CONVERT THE CLIPBOARD
 **
 ******************************************************************************/


/**
 * Reads the clipboard, converts Markdown to RTF, and writes back to the clipboard.
 */
async function clipboard_md_to_rtf_conv () {

  try {
    const str_OriginalClipboard = await clipboard.read();

    if ( bool_Debug ) {
      console.log( 'Original Clipboard Content:\n', str_OriginalClipboard );
    }

    const str_TransformedContent = transform_markdown_to_rtf( str_OriginalClipboard );

    if ( bool_Debug ) {
      console.log( 'Transformed Content:\n', str_TransformedContent );
    }

    await clipboard.write( str_TransformedContent );

    // Output the processing to the CLI
    console.log( `${console_color_conv('Ready to paste into Word', 'blue')}` );
  }


  catch ( obj_err_General ) {
    catch_message_print_exit( obj_err_General, 'Some major error somewhere.', 'Try to backtrace...', true );
  }

}

// Run
clipboard_md_to_rtf_conv();
