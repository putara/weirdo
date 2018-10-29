# Introduction
**We**b **I**con **R**ea**d**er **O**mega is a fast, light-weight web icon browser.  
Currently, it supports only the [FatCow Web Icons](https://www.fatcow.com/free-icons).

## Prerequisites
- Windows Vista or later
- Any x86/x64 processor that supports the SSE2 instruction set

## Usage
1. Download _sorted icons by color_ from the [website](https://www.fatcow.com/free-icons).
1. Extract the zip archive into a folder.
1. Copy weirdo.exe into the folder.
1. Run the app.
1. Wait for the app to build the cache database. (only the first time)
1. Do whatever.

## Build
1. Type `nmake` on the Visual Studio Command Prompt.

## Screenshots
<img alt="Screenshot" src="../assets/screenshot.png?raw=true" width="320">

## Todos
- [ ] Use the thread pool API to speed up the cache building process
- [ ] Port to non-Windows platforms

## License
- The regular expression parser is based on [Stephan Brumme's code](https://create.stephan-brumme.com/tiny-regular-expression-matcher/).
- LZ4 is taken from the [official GitHub repository](https://github.com/lz4/lz4).
- Anything else is released under the [WTFPL](http://www.wtfpl.net/about/).
