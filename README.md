<div align="center">

# 📦 Microsoft Office Installer

**Instalador de Microsoft Office (LTSC y Professional Plus) en Windows** con un solo comando de PowerShell, con interfaz gráfica (GUI) y línea de comandos (CLI).

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1-5391FE?style=flat-square&logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?style=flat-square&logo=windows&logoColor=white)]()
[![Sin dependencias](https://img.shields.io/badge/Sin%20dependencias-externas-2ea44f?style=flat-square)]()

</div>

---

## 📖 Descripción

Scripts de **PowerShell** que instalan Microsoft Office en Windows con un comando. Descarga el **Office Deployment Tool (ODT)** desde los servidores de Microsoft, arma el archivo de configuración con las opciones elegidas y ejecuta `setup.exe /configure`.

Está dirigido a administradores y usuarios avanzados que necesitan desplegar Office de forma rápida y reproducible, sin instaladores manuales.

## ✨ Características

- **Dos modos:** interfaz gráfica (`instalar-gui.ps1`) y consola (`instalar-cli.ps1`).
- **Ediciones soportadas:** LTSC Professional Plus **2024** y **2021**; Professional Plus **2019**, **2016** y **2013**.
- **Arquitectura:** 64 o 32 bits.
- **Complementos:** Project y Visio opcionales.
- **10 idiomas:** inglés, español, francés, alemán, portugués de Brasil, italiano, holandés, polaco, ruso y japonés.
- **Flujo guiado:** cada elección se confirma antes de tocar el sistema.
- **Tema claro/oscuro** y tarjetas con logos en la versión GUI.
- **Sin dependencias externas** y sin tocar la licencia (no incluye activación).

## 🛠️ Tecnologías utilizadas

- **Lenguaje:** PowerShell 5.0+
- **Herramientas:** Office Deployment Tool (ODT) de Microsoft
- **Plataforma:** Windows 10 / 11

## 📋 Requisitos previos

- **Windows 10 u 11**
- **PowerShell 5.0 o superior**
- **Permisos de administrador**
- **.NET Framework 4.5 o superior**
- **Internet** (el ODT y los archivos de Office se descargan de Microsoft)

## 🚀 Instalación

Abre PowerShell como **administrador**. Si es la primera vez que ejecutas scripts:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass
```

**Versión con ventana (GUI):**

```powershell
irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-gui.ps1 | iex
```

**Versión en consola (CLI):**

```powershell
irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-cli.ps1 | iex
```

También puedes descargar el `.ps1` y ejecutarlo con clic derecho → *Ejecutar con PowerShell*.

## 🧪 Uso

Ambas versiones te guían por: **versión → idioma → aplicaciones → resumen → confirmación**. Nada se instala sin tu confirmación, porque una instalación de Office no se deshace con facilidad.

## 📂 Estructura del proyecto

```
Microsoft-Office-installer/
├── instalar-gui.ps1        # Instalador con interfaz gráfica
├── instalar-cli.ps1        # Instalador de línea de comandos
├── index.html              # Página de presentación del proyecto
└── README.md
```

## 📝 Notas

- Este proyecto **instala** Office, no lo activa: necesitas una licencia propia.
- Los temporales se eliminan solos al terminar.
- Registro de cada ejecución: `%TEMP%\OfficeInstallerGUI_Install.log` y `%TEMP%\OfficeInstallerCLI_Install.log`.

## 📜 Licencia

Proyecto educativo y de código abierto; puedes usarlo libremente.

## ✍️ Autor

**Matthew Garay** — [GitHub](https://github.com/Matthew-Garay)
