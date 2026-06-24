# Contributing

Contributions are welcome.

## Development Setup

1. Install the .NET SDK with .NET Framework 4.8 build support.
2. Clone the repository.
3. Restore and build:

```powershell
dotnet restore RKExcelReportCompare.csproj
dotnet build RKExcelReportCompare.csproj -c Release /p:Platform=x64
```

## Pull Requests

- Keep changes focused.
- Include a short explanation of the user-facing behavior changed.
- Update `docs/USER_GUIDE.html` when installation or usage changes.
- Update `CHANGELOG.md` for release-facing changes.
- Do not commit `bin/`, `obj/`, `dist/`, or generated release packages.
