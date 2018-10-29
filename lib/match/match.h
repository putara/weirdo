// //////////////////////////////////////////////////////////
// match.h
// Copyright (c) 2015 Stephan Brumme. All rights reserved.
// see http://create.stephan-brumme.com/disclaimer.html
//

#pragma once

#include <wchar.h>  // for wchar_t

#ifdef __cplusplus
extern "C" {
#endif

// allowed metasymbols:
// ^  -> text must start with pattern (only allowed as first symbol in regular expression)
// $  -> text must end   with pattern (only allowed as last  symbol in regular expression)
// .  -> any literal
// x? -> literal x occurs zero or one time
// x* -> literal x is repeated arbitrarily
// x+ -> literal x is repeated at least once

/// points to first regex match or NULL if not found, text is treated as one line   and ^ and $ refer to its start/end
const wchar_t* match      (const wchar_t* text, const wchar_t* regex);

/// points to first regex match or NULL if not found, text is split into lines (\n) and ^ and $ refer to each line's start/end
const wchar_t* matchlines (const wchar_t* text, const wchar_t* regex);

#ifdef __cplusplus
}
#endif
