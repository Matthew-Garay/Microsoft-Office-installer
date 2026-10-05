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

La versión CLI también es interactiva: pregunta versión, idioma y aplicaciones
en pantalla, muestra un resumen y pide confirmación antes de instalar, porque
una instalación de Office no tiene vuelta atrás.

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

## Notas

### Cómo instalar

1. Abre PowerShell como administrador (clic derecho → "Ejecutar como administrador" o `Win + X` → "Terminal (Admin)")
2. Ejecuta uno de los comandos:
   - **GUI**: `irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-gui.ps1 | iex`
   - **CLI**: `irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-cli.ps1 | iex`
3. Sigue las instrucciones en pantalla para seleccionar versión, idioma y aplicaciones
4. Espera a que la instalación termine y reinicia el equipo

Si es tu primera vez ejecutando scripts en PowerShell, puede aparecer un aviso de seguridad. Usa `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass` para permitir scripts locales.