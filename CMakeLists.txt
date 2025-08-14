cmake_minimum_required(VERSION 3.16)
project(FirePlayer CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_EXPORT_COMPILE_COMMANDS ON)

# Options
option(BUILD_IMGUI_DEMO "Build ImGui demo window" OFF)

include(FetchContent)

# GLFW
FetchContent_Declare(
    glfw
    GIT_REPOSITORY https://github.com/glfw/glfw.git
    GIT_TAG 3.3.9
)
set(GLFW_BUILD_DOCS OFF CACHE BOOL "" FORCE)
set(GLFW_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(GLFW_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
FetchContent_MakeAvailable(glfw)

# ImGui
FetchContent_Declare(
    imgui
    GIT_REPOSITORY https://github.com/ocornut/imgui.git
    GIT_TAG v1.91.0
)
FetchContent_MakeAvailable(imgui)

add_library(imgui
    ${imgui_SOURCE_DIR}/imgui.cpp
    ${imgui_SOURCE_DIR}/imgui_draw.cpp
    ${imgui_SOURCE_DIR}/imgui_widgets.cpp
    ${imgui_SOURCE_DIR}/imgui_tables.cpp
    ${imgui_SOURCE_DIR}/backends/imgui_impl_glfw.cpp
    ${imgui_SOURCE_DIR}/backends/imgui_impl_opengl3.cpp
    ${imgui_SOURCE_DIR}/misc/cpp/imgui_stdlib.cpp
)

target_include_directories(imgui PUBLIC
    ${imgui_SOURCE_DIR}
    ${imgui_SOURCE_DIR}/backends
    ${imgui_SOURCE_DIR}/misc/cpp
)

# Use GLEW as OpenGL loader for ImGui backend
target_compile_definitions(imgui PUBLIC IMGUI_IMPL_OPENGL_LOADER_GLEW)

target_link_libraries(imgui PUBLIC glfw)

if(UNIX AND NOT APPLE)
    target_link_libraries(imgui PUBLIC dl pthread)
endif()

# nlohmann_json
FetchContent_Declare(
    nlohmann_json
    GIT_REPOSITORY https://github.com/nlohmann/json.git
    GIT_TAG v3.11.3
)
FetchContent_MakeAvailable(nlohmann_json)

# OpenGL
find_package(OpenGL REQUIRED)
find_package(GLEW REQUIRED)

# Sources
file(GLOB FIRE_PLAYER_SOURCES
    src/*.cpp
)

add_executable(fire_player ${FIRE_PLAYER_SOURCES})

target_include_directories(fire_player PRIVATE src)

# Link order
target_link_libraries(fire_player PRIVATE
    imgui
    glfw
    GLEW::GLEW
    OpenGL::GL
    nlohmann_json::nlohmann_json
)

# On Linux, link extra system libs if necessary
if(UNIX AND NOT APPLE)
    target_link_libraries(fire_player PRIVATE X11)
endif()

# Copy assets next to the binary
add_custom_target(copy_assets ALL
    COMMAND ${CMAKE_COMMAND} -E copy_directory
            ${CMAKE_SOURCE_DIR}/assets
            ${CMAKE_BINARY_DIR}/assets
)
add_dependencies(fire_player copy_assets)