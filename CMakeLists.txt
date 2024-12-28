# Specify the minimum CMake version required
cmake_minimum_required(VERSION 3.15)

# Project information
project(MT5-Monte-Carlo-Tool VERSION 1.0.0 LANGUAGES CXX)

# Set the C++ standard to use
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Output paths for binaries and libraries
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)

# Define project directories
set(SRC_DIR ${CMAKE_SOURCE_DIR}/src)
set(INCLUDE_DIR ${CMAKE_SOURCE_DIR}/include)

# Enable vcpkg integration (ensure correct path)
if(DEFINED ENV{VCPKG_ROOT})
    set(CMAKE_TOOLCHAIN_FILE "$ENV{VCPKG_ROOT}/scripts/buildsystems/vcpkg.cmake" CACHE STRING "Vcpkg toolchain file")
else()
    message(FATAL_ERROR "VCPKG_ROOT environment variable is not set.")
endif()

# Collect source and header files
file(GLOB_RECURSE SOURCES ${SRC_DIR}/*.cpp)
file(GLOB_RECURSE HEADERS ${INCLUDE_DIR}/*.hpp)

# Add the main executable target
add_executable(MT5-Monte-Carlo-Tool ${SOURCES} ${HEADERS})

# Include directories for the project
target_include_directories(MT5-Monte-Carlo-Tool PUBLIC ${INCLUDE_DIR})

# Link external library xlnt (managed by vcpkg)
find_package(xlnt CONFIG REQUIRED)
target_link_libraries(MT5-Monte-Carlo-Tool PRIVATE xlnt::xlnt)

# Install targets for creating installable packages
install(TARGETS MT5-Monte-Carlo-Tool
    RUNTIME DESTINATION bin
    LIBRARY DESTINATION lib
    ARCHIVE DESTINATION lib
)

# Packaging support
set(CPACK_PROJECT_NAME ${PROJECT_NAME})
set(CPACK_PROJECT_VERSION ${PROJECT_VERSION})
include(CPack)
