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
irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-gui.ps1 | iex
```

Enlace directo: `https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-gui.ps1`

### Línea de comandos (CLI)

Descarga y ejecuta la versión en consola:

```powershell
irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-cli.ps1 | iex
```

Enlace directo: `https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-cli.ps1`

### Modo desatendido

La versión CLI acepta parámetros, así que también sirve desde un script sin
intervención:

```powershell
.\instalar-cli.ps1 -Version 2 -Language es-ES -Apps Word,Excel,PowerPoint,Outlook -Yes
```

| Parámetro | Valores | Por omisión |
|-----------|---------|-------------|
| `-Version` | 1 a 5 (2024, 2021, 2019, 2016, 2013) | pregunta |
| `-Language` | Código de cultura, p. ej. `es-ES`, `en-US` | pregunta |
| `-Apps` | `Word,Excel,PowerPoint,...` | pregunta (Word, Excel y PowerPoint) |
| `-Yes` | salta la confirmación final | pregunta |

Nombres válidos para `-Apps`: `Word`, `Excel`, `PowerPoint`, `Outlook`, `Access`,
`Publisher`, `OneNote`, `SkypeForBusiness`, `Project`, `Visio`. Un nombre o un
idioma que no exista aborta el script con el código de salida 1 y la lista de
valores válidos, en lugar de instalar cualquier cosa.

**Aviso:** sin `-Yes`, el script pide confirmación antes de instalar, porque una
instalación de Office no tiene vuelta atrás.

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
- **Descarga automática del ODT** — descarga el Office Deployment Tool desde los servidores oficiales de Microsoft con `System.Net.WebClient`
- **Fuentes de ODT en lista** — `$script:odtUrls` se recorre en orden y el script pasa a la siguiente si una falla; hoy hay una sola configurada, la de Microsoft
- **Detección de arquitectura** — `[Environment]::Is64BitOperatingSystem` decide entre `OfficeClientEdition="64"` y `"32"`
- **Configuración XML dinámica** — genera el archivo de configuración de Office en tiempo de ejecución según la versión y aplicaciones seleccionadas
- **Excluye lo que no se pidió** — las aplicaciones no seleccionadas y los servicios incluidos (Bing, Groove, Lync, OneDrive, Teams) van como `ExcludeApp`
- **GUI con Windows Forms** — interfaz gráfica construida con `System.Windows.Forms` y `DarkMode` personalizado
- **CLI con terminal UI** — interfaz de consola con banner tipográfico, pasos numerados y barras de progreso con bloques (`█`)
- **CLI parametrizable** — versión, idioma y aplicaciones se pueden pasar por parámetro para instalaciones desatendidas
- **Modo no interactivo seguro** — un idioma o una aplicación desconocida aborta con el código 1 y la lista de valores válidos, en vez de instalar algo que no se pidió
- **Confirmación previa** — la CLI pide confirmación antes de instalar, salvo que se pase `-Yes`
- **Registros de instalación** — logs en `%TEMP%\OfficeInstallerGUI_Install.log` y `%TEMP%\OfficeInstallerCLI_Install.log`
- **Manejo de errores** — try/catch con captura de excepciones y escritura en logs
- **Execution Policy** — soporte para `Set-ExecutionPolicy` en primera ejecución
- **Limpieza automática** — elimina archivos temporales tras la instalación

---

## Licencia

Esta herramienta **instala** Office; no lo activa. Necesitas una licencia válida
de Office para usarlo. La versión GUI incluye un intento opcional de activación
con Microsoft Activation Scripts, pero la versión CLI no lo hace.

---

## Notas

### Cómo instalar

1. Abre PowerShell como administrador (clic derecho → "Ejecutar como administrador" o `Win + X` → "Terminal (Admin)")
2. Ejecuta uno de los comandos:
   - **GUI**: `irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-gui.ps1 | iex`
   - **CLI**: `irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-cli.ps1 | iex`
3. Sigue las instrucciones en pantalla para seleccionar versión, idioma y aplicaciones
4. Espera a que la instalación termine y reinicia el equipo

Si es tu primera vez ejecutando scripts en PowerShell, puede aparecer un aviso de seguridad. Usa `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass` para permitir scripts locales.