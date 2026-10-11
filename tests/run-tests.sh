#!/usr/bin/env bash
# Builds the add-in (compile check, warnings are errors) and runs the logic tests under Mono.
# Needs: dotnet SDK 8, mono-runtime, mono-mcs, libmono-system-windows-forms4.0-cil,
#        libmono-system-net-http4.0-cil, libmono-system-web-extensions4.0-cil
set -euo pipefail
cd "$(dirname "$0")"
dotnet build CompileCheck.csproj -nologo -v q
OUT=bin/Debug/net48
cp ~/.nuget/packages/exceldna.integration/1.8.0/lib/net452/ExcelDna.Integration.dll "$OUT/"
mcs -nologo -r:"$OUT/ExcelDna.Integration.dll" -r:System.Windows.Forms.dll -r:System.Xml.dll -r:System.Core.dll -r:System.Web.Extensions.dll LogicTests.cs -out:"$OUT/LogicTests.exe"
cd "$OUT" && mono LogicTests.exe
mcs -nologo -r:System.Core.dll ../../../CalcCardTest.cs -out:CalcCardTest.exe
mono CalcCardTest.exe
# the web dashboard in a real browser, when Playwright is installed
if NODE_PATH=$(npm root -g 2>/dev/null) node -e "require('playwright')" 2>/dev/null; then
  NODE_PATH=$(npm root -g) node ../../../calc-web.mjs .
fi
