# MSOI · Office Installer

MSOI — Instala Office en Windows desde PowerShell con un solo comando.

```powershell
irm https://matthew-garay.github.io/Microsoft-Office-installer/instalar-gui.ps1 | iex
irm https://matthew-garay.github.io/Microsoft-Office-installer/instalar-cli.ps1 | iex
```

## Contenido

| Archivo | Qué es |
|---|---|
| `index.html` | Página web del proyecto (con selector español/inglés) |
| `instalar-gui.ps1` | Instalador con ventana gráfica |
| `instalar-cli.ps1` | Instalador en modo consola |
| `robots.txt` / `sitemap.xml` | Ayudan a que Google/Bing indexen la página |

## Requisitos

Windows 10/11, PowerShell como administrador.

## Publicar en GitHub Pages

1. Crea un repo público llamado `Microsoft-Office-installer` en tu cuenta (`Matthew-Garay`).
2. Sube TODO el contenido de esta carpeta a la raíz del repo (no dentro de otra carpeta):

```powershell
cd d:\Microsoft-Office-installer-main\office-installer
git init
git add .
git commit -m "primer commit"
git branch -M main
git remote add origin https://github.com/Matthew-Garay/Microsoft-Office-installer.git
git push -u origin main
```

3. En el repo: Settings → Pages → Branch `main` / carpeta `/ (root)` → Save.
4. Espera un par de minutos y entra a https://matthew-garay.github.io/Microsoft-Office-installer/
