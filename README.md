# Microsoft Office Installer

Script de PowerShell que instala Office en Windows con un comando. Hay dos versiones que hacen lo mismo: una con ventana y otra para consola. Descargan el Office Deployment Tool de los servidores de Microsoft, arman el archivo de configuracion con lo que elegiste y lanzan `setup.exe /configure`.

Soporta las ediciones LTSC Professional Plus 2024 y 2021, y las Professional Plus 2019, 2016 y 2013, en 64 o 32 bits. Project y Visio se pueden agregar como complementos. Idiomas: ingles, espanol, frances, aleman, portugues de Brasil, italiano, holandes, polaco, ruso y japones.

---

## Uso

Hay que abrir PowerShell como administrador. Si es la primera vez que corres scripts, Windows avisa; se quita con `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass`.

**Con ventana:**

```powershell
irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-gui.ps1 | iex
```

**En consola:**

```powershell
irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-cli.ps1 | iex
```

Las dos preguntan version, idioma y aplicaciones, muestran un resumen y piden confirmacion antes de tocar nada, porque una instalacion de Office no se deshace. En la ventana ademas hay interruptor de tema claro y oscuro, las apps se eligen con tarjetas con su logo y los idiomas salen con su bandera.

Tambien se puede bajar el `.ps1` y correrlo local con clic derecho, "Ejecutar con PowerShell".

---

## Requisitos

- Windows 10 o Windows 11
- PowerShell 5.0 o superior
- Permisos de administrador
- .NET Framework 4.5 o superior
- Internet (el ODT y los archivos de Office se bajan de Microsoft)

---

## Notas

Esto **instala** Office, no incluye licencia: hay que tener una propia para usarlo. Los temporales se borran solos al terminar y lo de cada ejecucion queda en `%TEMP%\OfficeInstallerGUI_Install.log` y `%TEMP%\OfficeInstallerCLI_Install.log`.