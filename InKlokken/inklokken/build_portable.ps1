$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

dotnet publish .\InKlokken.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  /p:PublishSingleFile=true `
  /p:IncludeNativeLibrariesForSelfExtract=true `
  -o .\dist\InKlokken
