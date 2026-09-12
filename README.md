# Microsoft Office Installer

Herramienta para instalar Microsoft Office en Windows desde PowerShell con un solo comando. Disponible en dos modos: interfaz gráfica (GUI) y línea de comandos (CLI).

---

## Versiones de Office soportadas

| Versión | Edición |
|---------|---------|
| Office LTSC Professional Plus 2024 | 64 bits / 32 bits |
| Office LTSC Professional Plus 2021 | 64 bits / 32 bits |
| Office Professional Plus 2019 | 64 bits / 32 bits |
| Office Professional Plus 2016 | 64 bits / 32 bits |
| Office Professional Plus 2013 | 64 bits / 32 bits |

También soporta **Project** y **Visio** (Professional y Standard) como complementos opcionales.

---

## Uso

### Interfaz gráfica (GUI)

Descarga y ejecuta la versión con ventana gráfica:

```powershell
irm https://matthew-garay.github.io/Microsoft-Office-installer/instalar-gui.ps1 | iex
```

Enlace directo: `https://matthew-garay.github.io/Microsoft-Office-installer/instalar-gui.ps1`

### Línea de comandos (CLI)

Descarga y ejecuta la versión en consola:

```powershell
irm https://matthew-garay.github.io/Microsoft-Office-installer/instalar-cli.ps1 | iex
```

Enlace directo: `https://matthew-garay.github.io/Microsoft-Office-installer/instalar-cli.ps1`

---

## Requisitos

| Requisito | Detalle |
|-----------|---------|
| **Sistema operativo** | Windows 10 / Windows 11 |
| **Arquitectura** | 64 bits o 32 bits (se detecta automáticamente) |
| **PowerShell** | 5.0 o superior |
| **Permisos** | Administrador (obligatorio) |
| **Framework** | .NET Framework 4.5 o superior |
| **Conexión** | Internet (descarga el ODT y los archivos de Office desde los servidores de Microsoft) |

---

## Características (código)

- **PowerShell nativo** — sin dependencias externas, solo cmdlets del sistema
- **Descarga automática del ODT** — descarga el Office Deployment Tool desde los servidores oficiales de Microsoft con `Invoke-WebRequest`
- **Múltiples fuentes de ODT** — reintentos automáticos si un servidor falla (fallback entre URLs)
- **Detección de arquitectura** — identifica 32/64 bits via `Get-WmiObject` o `Environment.Is64BitOperatingSystem`
- **Configuración XML dinámica** — genera el archivo de configuración de Office en tiempo de ejecución según la versión y aplicaciones seleccionadas
- **GUI con Windows Forms** — interfaz gráfica construida con `System.Windows.Forms` y `DarkMode` personalizado
- **CLI con terminal UI** — interfaz de consola con banner tipográfico, pasos numerados y barras de progreso con bloques (`█`)
- **Activación MAS** — integración con Microsoft Activation Scripts via `irm | iex` opcional
- **Registros de instalación** — logs en `%TEMP%\OfficeInstallerGUI_Install.log` y `%TEMP%\OfficeInstallerCLI_Install.log`
- **Manejo de errores** — try/catch con captura de excepciones y escritura en logs
- **Execution Policy** — soporte para `Set-ExecutionPolicy` en primera ejecución
- **Limpieza automática** — elimina archivos temporales tras la instalación

---

## Notas

### Cómo instalar

1. Abre PowerShell como administrador (clic derecho → "Ejecutar como administrador" o `Win + X` → "Terminal (Admin)")
2. Ejecuta uno de los comandos:
   - **GUI**: `irm https://matthew-garay.github.io/Microsoft-Office-installer/instalar-gui.ps1 | iex`
   - **CLI**: `irm https://matthew-garay.github.io/Microsoft-Office-installer/instalar-cli.ps1 | iex`
3. Sigue las instrucciones en pantalla para seleccionar versión e idioma
4. Espera a que la instalación termine

Si es tu primera vez ejecutando scripts en PowerShell, puede aparecer un aviso de seguridad. Usa `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass` para permitir scripts locales.