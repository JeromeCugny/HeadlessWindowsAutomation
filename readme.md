# Headless Windows Automation

Headless Windows Automation is a library designed for automating Windows desktop applications without the need for a visible user interface. This repository contains the source code and build scripts necessary to build and package the library.

## Requirements

To build this project, you need the following:

- Visual Studio 2022
- .NET Framework v4.8
- [NuGet CLI in your PATH](https://www.nuget.org/downloads)
- [CMake](https://cmake.org/download/)

## Building the Project

To build the project, follow these steps:

1. Generate the solution:
    ```sh
    ./cmake-vs2022-x64.bat
    ```
2. Build the project:
    ```sh
    cmake --build build --config Release --target ALL_BUILD
    ```

## Additional Information

For more details about the package, refer to the [package readme](package/readme.md).

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.