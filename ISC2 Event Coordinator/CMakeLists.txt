cmake_minimum_required(VERSION 3.20)
project(ISC2EventCoordinator LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Link the C++ runtime statically so no MSVC redistributable is needed
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

include(FetchContent)

# pugixml for XML manipulation inside DOCX/DOCM files
FetchContent_Declare(
    pugixml
    GIT_REPOSITORY https://github.com/zeux/pugixml.git
    GIT_TAG        v1.15
)
FetchContent_MakeAvailable(pugixml)

# nlohmann/json for Constant Contact API JSON responses
FetchContent_Declare(
    json
    GIT_REPOSITORY https://github.com/nlohmann/json.git
    GIT_TAG        v3.11.3
)
FetchContent_MakeAvailable(json)

add_executable(EventCoordinator src/main.cpp)
target_link_libraries(EventCoordinator PRIVATE
    pugixml
    nlohmann_json::nlohmann_json
    winhttp
    crypt32
    ws2_32
    ole32
    oleaut32
    uuid
    shell32
    shlwapi
    gdiplus
)

# Copy executable to project root for easy access
add_custom_command(TARGET EventCoordinator POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:EventCoordinator>
        "${CMAKE_SOURCE_DIR}/EventCoordinator.exe"
)