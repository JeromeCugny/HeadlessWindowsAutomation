cmake_minimum_required(VERSION 3.8)

# Describe project
project("HeadlessWindowsAutomation" 
  LANGUAGES CSharp
)

# Set the .NET Framework version
set(DOTNET_VERSION "v4.8")  # Version for Visual Studio
set(DOTNET_TARGET_FRAMEWORK "net48")  # (same) version but for nuget package

set(NUGET_PACKAGE_VERSION "1.0.5")

# Define the output path for the build
set(OUTPUT_PATH "${CMAKE_BINARY_DIR}/bin")

# Define the output path for the packaging
set(DELIVERY_OUTPUT_PATH "${CMAKE_SOURCE_DIR}/.output")

add_subdirectory("./lib")
add_subdirectory("./package")
