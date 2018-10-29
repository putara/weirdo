// //////////////////////////////////////////////////////////
// match.c
// Copyright (c) 2015 Stephan Brumme. All rights reserved.
// see http://create.stephan-brumme.com/disclaimer.html
//

// I improved Rob Pike's code found in the excellent book "Beautiful Code", ISBN 9780596510046
// http://www.cs.princeton.edu/courses/archive/spr09/cos333/beautiful.html
// additions (already mentioned in his extension proposals):
// - match a symbol which is repeated at least once "x+" (which is the same as "xx*")
// - match a symbol which appears at most once      "x?"
// - regex can contain escaped metasymbols => "\." searches for a dot
// - matchlines splits input into lines

#include "match.h"

// we have to define NULL because we don't include string.h directly or indirectly
#ifndef NULL
#define NULL 0
#endif

// track escaped characters
#define YES  1
#define NO   0

// Windows uses \r\n for newlines, Unix only \n
#define PROCESS_CRLF 1

#define TOWLOWER(x) (((x) >= L'A' && (x) <= L'Z') ? ((x) + ('a' - 'A')) : (x))

int matchchar(wchar_t x, wchar_t y)
{
    return TOWLOWER(x) == TOWLOWER(y);
}

/// matchhere: search for regex only at beginning of text, stop at textEnd or when a zero-byte is found
/** local helper function for match():
    @param text    - text to be matched
    @param textEnd - last byte of text to be considered, can be NULL if text is zero-terminated
    @param regex   - simplified regular expression as explained in the header file
    @param escaped - YES treats first character of regex as a literal only and ignores their meta symbol meaning, but NO is default
    @return if successful points beyond last matching literal, else NULL
**/
static const wchar_t* matchhere(const wchar_t* text, const wchar_t* textEnd, const wchar_t* regex, int escaped)
{
  // process current byte
  if (escaped == NO)
  {
    switch (*regex)
    {
    case 0: // regex empty ?
      return text;
    case L'\\': // treat next character as a literal, not a metasymbol
      return matchhere(text, textEnd, regex + 1, YES);
    case L'^': // ^ was already handled in match() / matchlines()
      return NULL;
    }
  }

  // lookahead one byte (i.e. watch out for meta-symbols)
  switch (regex[1])
  {
  case L'+': // symbol appears at least once
    while (matchchar(*regex, *text) || // keep running if next symbol matches
           (*regex == L'.' && escaped == NO && text != textEnd && *text != 0))
    {
      // match remainder
      const wchar_t* matched = matchhere(++text, textEnd, regex + 2, NO);
      if (matched != NULL)
        return matched;
    }

    // running out of text
    break; // failed

  case L'*': // symbol appears arbitrarily often
    // eat text symbol-by-symbol and try to match remaining regex
    do
    {
      // match remaining regex
      const wchar_t* matched = matchhere(text, textEnd, regex + 2, NO);
      if (matched != NULL)
        return matched;
    } while (matchchar(*regex, *text++) || // keep running if next symbol matches
             (*regex == L'.' && escaped == NO && text != textEnd && *text != 0));

    // running out of text
    break; // failed

  case L'?': // symbol appears exactly zero or one time
    // assume symbol doesn't appear
    {
      const wchar_t* matched = matchhere(text, textEnd, regex + 2, NO);
      if (matched != NULL)
        return matched;
    }

    // continue matching at next position only if symbol appears
    if (matchchar(*regex, *text) || (*regex == L'.' && escaped == NO && *text != 0))
      return matchhere(text + 1, textEnd, regex + 2, NO);

    break; // failed

  case 0: // *regex will be the last symbol
    // dollar matches only end of text
    if (*regex == L'$' && escaped == NO)
    {
      if (text == textEnd || *text == 0)
        return text;
      break; // failed
    }

    // nothing left for further matching
    if (text == textEnd || *text == 0)
      break; // failed

    // does last literal match ?
    return (matchchar(*regex, *text) || (*regex == L'.' && escaped == NO)) ? text + 1 : NULL;

  default: // no meta-symbol, just plain character
    // nothing left for further matching
    if (text == textEnd || *text == 0)
      break; // failed

    // $ must be followed by 0 (previous case)
    if (*regex == L'$' && escaped == NO)
      break; // failed

    // same character (or any character) ?
    if (matchchar(*regex, *text) || (*regex == L'.' && escaped == NO))
      return matchhere(text + 1, textEnd, regex + 1, NO);

    break; // failed
  }

  // failed
  return NULL;
}


/// points to first regex match or NULL if not found, text is treated as one line   and ^ and $ refer to its start/end
/** @param text    - text to be matched
    @param regex   - simplified regular expression as explained in the header file
    @return if successful points to the next matching literal, else NULL if no match
**/
const wchar_t* match(const wchar_t* text, const wchar_t* regex)
{
  // detect invalid input
  if (!text || !regex)
    return NULL;

  // match only start
  if (*regex == L'^')
    // matchhere looks out for zero byte if textEnd parameter is set to NULL
    return matchhere(text, NULL, regex + 1, NO) ? text : NULL;

  // try to start match at every text position
  do
  {
    // match found ?
    if (matchhere(text, NULL, regex, NO))
      return text;
  } while (*text++ != 0);

  // failed to match
  return NULL;
}


/// points to first regex match or NULL if not found, text is split into lines (\n) and ^ and $ refer to each line's start/end
/** @param text    - text to be matched
    @param regex   - simplified regular expression as explained in the header file
    @return if successful points to the next matching literal, else NULL if no match
**/
const wchar_t* matchlines(const wchar_t* text, const wchar_t* regex)
{
  const wchar_t* right;

  // detect invalid input
  if (!text || !regex)
    return NULL;

  // until the end of the world ...
  do
  {
    // use simple matching for end-of-text
    if (*text == 0)
      return match(text, regex);

    // find extents of current line
    right = text;
#ifdef PROCESS_CRLF
    while (*right != 0 && *right != L'\n' && *right != L'\r')
      right++;
#else
    while (*right != 0 && *right != L'\n')
      right++;
#endif

    // ^ refers to initial text pointer position in the first line,
    // thereafter to each start of a new line
    // => if initial text pointer refers to the middle of a line, a false ^ match may be produced

    // match only start
    if (*regex == L'^')
    {
      if (matchhere(text, right, regex + 1, NO))
        return text;
    }
    else
    {
      // try to start match at every text position
      while (text <= right)
      {
        // match found ?
        if (matchhere(text, right, regex, NO))
          return text;

        // no match yet, continue
        text++;
      }
    }

    // skip forward to next line
    text = right;

    // skip new-line
#ifdef PROCESS_CRLF
    while (*text == L'\r')
      text++;
#endif
    if    (*text == L'\n')
      text++;
  } while (*text++ != 0);

  // failed to match
  return NULL;
}
