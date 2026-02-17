cmake_minimum_required(VERSION 3.15)
project(Gluecksrad LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_executable(gluecksrad WIN32 gluecksrad.cpp)

target_link_libraries(gluecksrad PRIVATE
    comctl32
    comdlg32
    gdi32
    user32
    shell32
)

# Enable Unicode
target_compile_definitions(gluecksrad PRIVATE UNICODE _UNICODE)

# For MinGW: link statically so the .exe is fully standalone
if(MINGW)
    target_link_options(gluecksrad PRIVATE -static -static-libgcc -static-libstdc++)
endif()

# MSVC: use /utf-8 for proper Unicode source handling
if(MSVC)
    target_compile_options(gluecksrad PRIVATE /utf-8 /O2)
endif()
