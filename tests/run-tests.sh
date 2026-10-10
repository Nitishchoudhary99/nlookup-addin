#!/usr/bin/env bash
# Builds the add-in (compile check, warnings are errors) and runs the logic tests under Mono.
# Needs: dotnet SDK 8, mono-runtime, mono-mcs, libmono-system-windows-forms4.0-cil
set -euo pipefail
cd "$(dirname "$0")"
dotnet build CompileCheck.csproj -nologo -v q
OUT=bin/Debug/net48
cp ~/.nuget/packages/exceldna.integration/1.8.0/lib/net452/ExcelDna.Integration.dll "$OUT/"
mcs -nologo -r:"$OUT/ExcelDna.Integration.dll" -r:System.Windows.Forms.dll -r:System.Xml.dll -r:System.Core.dll LogicTests.cs -out:"$OUT/LogicTests.exe"
cd "$OUT" && mono LogicTests.exe
