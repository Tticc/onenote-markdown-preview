# Third-party notices

This project is distributed under the MIT License (see `LICENSE`). It depends on
third-party components that are redistributed in the release archive; their
licenses are reproduced here and their license files are shipped alongside the
binaries.

## Markdig 0.37.0

- License: BSD-2-Clause
- Copyright: Alexandre Mutel and contributors
- Shipped as: `Markdig.dll`, `Markdig.LICENSE.txt`

## System.* support libraries (MIT, .NET Foundation)

Redistributed to satisfy Markdig's dependencies on .NET Framework 4.8:

| Package | Version | License |
|---|---|---|
| System.Memory | 4.5.5 | MIT |
| System.Buffers | 4.5.1 | MIT |
| System.Numerics.Vectors | 4.5.0 | MIT |
| System.Runtime.CompilerServices.Unsafe | 4.5.3 | MIT |

Where the NuGet package provides a `LICENSE.TXT`, `build.ps1` copies it next to
the binary as `<PackageId>.LICENSE.txt`.

## Microsoft Office primary interop assemblies (not redistributed)

`Extensibility.dll`, `office.dll` and `Microsoft.Office.Interop.OneNote.dll` are
referenced from the machine's GAC at build time and are **not** included in the
release archive. They are part of the user's Office installation.
