project(HeadlessWindowsAutomation LANGUAGES CSharp)

# Add the source files
file(GLOB_RECURSE SOURCES "src/*.cs")

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
)
