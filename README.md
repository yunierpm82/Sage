# Sage

Aplicación de prueba en VB.NET (WinForms) para verificar el flujo Claude Code + GitHub.

Permite indicar servidor, base de datos y credenciales, conectarse a una base de datos SQL Server y listar sus tablas.

## Requisitos

- [.NET 8 SDK](https://dotnet.microsoft.com/download)
- [Inno Setup](https://jrsoftware.org/isinfo.php) (solo para generar el instalador)

## Ejecutar en desarrollo

```bash
dotnet run --project src/Sage/Sage.vbproj
```

## Publicar (build de release, self-contained, un solo .exe)

```bash
dotnet publish src/Sage/Sage.vbproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

## Generar el instalador

Con Inno Setup instalado (y `ISCC.exe` en el PATH):

```bash
iscc installer/Sage.iss
```

El instalador queda en `installer/Output/SageSetup.exe`.

## Estructura

```
Sage/
├── src/Sage/        Código fuente VB.NET (WinForms)
├── installer/        Script de Inno Setup
└── README.md
```
