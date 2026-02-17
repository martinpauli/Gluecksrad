# Glücksrad – Faire Zufallsauswahl (C++ Native)

A native Windows application for fair random selection from a name list.
No Python, no runtime dependencies — instant startup.

## Features

- Load names + counters from **CSV** files (semicolon or comma separated, UTF-8)
- Fair selection: only candidates with the lowest counter are eligible
- Animated "spin" through the list with deceleration effect
- Winner blink highlight (green)
- Multi-draw with batch duplicate avoidance
- Auto-save counters back to CSV
- Configurable spin speed, rounds, blink parameters
- Resizable window

## CSV File Format

```
Name;Counter
Alice;0
Bob;2
Charlie;0
```

- **Column 1**: Name (required)
- **Column 2**: Counter (optional, defaults to 0)
- Both `;` and `,` are auto-detected as delimiters
- UTF-8 encoding (with or without BOM)

### Converting from Excel

If you previously used `.xlsx` files, simply open them in Excel and
**Save As → CSV UTF-8 (comma delimited) (*.csv)**.

## Building

### Option A: Visual Studio (MSVC)

```
mkdir build && cd build
cmake .. -G "Visual Studio 17 2022"
cmake --build . --config Release
```

The `.exe` will be in `build\Release\gluecksrad.exe`.

### Option B: MinGW-w64

```
mkdir build && cd build
cmake .. -G "MinGW Makefiles" -DCMAKE_BUILD_TYPE=Release
cmake --build .
```

### Option C: Direct compilation (no CMake)

**MSVC Developer Command Prompt:**
```
cl /O2 /EHsc /DUNICODE /D_UNICODE gluecksrad.cpp /Fe:gluecksrad.exe user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib
```

**MinGW-w64:**
```
g++ -O2 -mwindows -DUNICODE -D_UNICODE gluecksrad.cpp -o gluecksrad.exe -lcomctl32 -lcomdlg32 -lgdi32 -luser32 -lshell32 -static
```

## Notes

- The `.exe` is fully standalone — no DLLs or runtimes needed
- Typical executable size: ~100–200 KB (vs. 50+ MB for a PyInstaller bundle)
- Startup is instant (no Python interpreter to load)
