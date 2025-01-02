project(HeadlessWindowsAutomation LANGUAGES CSharp)

# Add the source files
file(GLOB_RECURSE SOURCES "src/*.cs")

# Configure the AssemblyInfo.cs file
configure_file(${CMAKE_CURRENT_SOURCE_DIR}/AssemblyInfo.cs.in ${CMAKE_CURRENT_BINARY_DIR}/AssemblyInfo.cs @ONLY)
list(APPEND SOURCES ${CMAKE_CURRENT_BINARY_DIR}/AssemblyInfo.cs)

# Add the library target
add_library(${PROJECT_NAME} SHARED ${SOURCES})

# Set the properties for the library
set_target_properties(${PROJECT_NAME} PROPERTIES
    OUTPUT_NAME "HeadlessWindowsAutomation"
    RUNTIME_OUTPUT_DIRECTORY ${OUTPUT_PATH}
    LIBRARY_OUTPUT_DIRECTORY ${OUTPUT_PATH}
    ARCHIVE_OUTPUT_DIRECTORY ${OUTPUT_PATH}
    VS_DOTNET_TARGET_FRAMEWORK_VERSION ${DOTNET_VERSION}
)

# Add references to necessary assemblies
set_property(TARGET ${PROJECT_NAME} PROPERTY VS_DOTNET_REFERENCES
    "System"
    "System.Core"
    "System.Windows.Forms"
    "UIAutomationClient"
    "UIAutomationTypes"
    "WindowsBase"
    "System.Diagnostics"
    "System.Text.RegularExpressions"
    "System.Drawing"
)
