# Generate nuspec file. Call it at build time.

configure_file(${SOURCE_PATH}/template.nuspec.in ${NUSPEC_FILE} @ONLY ESCAPE_QUOTES)
