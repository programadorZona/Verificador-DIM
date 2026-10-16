# Guía: levantar SmartDIM manualmente (puerto 8501)

## ¿Qué es el proyecto?

**SmartDIM** es una app web hecha en **Streamlit** que verifica Declaraciones de Importación (DIM, PDF de la DIAN) contra:

1. Un **Excel de subpartidas** (comparación DIM vs subpartida arancelaria).
2. Un **Excel de anexos FMM** (validación de anexos del Formulario de Movimiento de Mercancías).

Muestra los resultados en pantalla y permite descargarlos en Excel.

| Archivo | Rol |
|---|---|
| `app.py` | Interfaz Streamlit (carga de archivos, botones, resultados, descargas) |
| `verificacion_dim.py` | Lógica: `ExtractorDIANSimplificado` (lee PDFs con pdfplumber), `ComparadorDatos`, `ExtractorSubpartidas`, `ValidadorDeclaracionImportacionCompleto` |
| `requirements.txt` | Dependencias de Python |
| `railway.toml` | Configuración de despliegue en Railway (usa `$PORT`) |
| `.devcontainer/` | Configuración para Codespaces / Dev Containers |

## Requisitos

- Windows con **Python 3.11 o superior** (el `.venv` actual usa Python 3.14).
- Acceso a la carpeta `C:\Desarrollo\verificacion dim`.

## Pasos (PowerShell)

### 1. Ir a la carpeta del proyecto

```powershell
cd "C:\Desarrollo\verificacion dim"
```

### 2. Crear el entorno virtual (solo la primera vez)

Si la carpeta `.venv` ya existe, salta este paso.

```powershell
python -m venv .venv
```

### 3. Activar el entorno virtual

```powershell
.\.venv\Scripts\Activate.ps1
```

> Si PowerShell bloquea el script con un error de "ejecución de scripts deshabilitada", ejecuta una vez:
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`
> (o, en `cmd.exe`, usa `.venv\Scripts\activate.bat`).

Al activarlo verás `(.venv)` al inicio de la línea.

### 4. Instalar dependencias (primera vez o si cambia `requirements.txt`)

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 5. Levantar la aplicación en el puerto 8501

```powershell
streamlit run app.py --server.address 0.0.0.0 --server.port 8501 --server.headless true
```

- `--server.address 0.0.0.0` → acepta conexiones desde otros equipos de la red, no solo desde este PC.
- `--server.port 8501` → puerto de la app.
- `--server.headless true` → no abre el navegador automáticamente ni pide correo al iniciar.

La consola mostrará algo como:

```
Local URL:   http://localhost:8501
Network URL: http://10.10.2.13:8501
```

### 6. Abrir la app

- Desde el mismo equipo: <http://localhost:8501>
- Desde otro equipo de la red: `http://<IP-del-equipo>:8501` (la *Network URL* que muestra la consola).

### 7. Detener la app

En la consola donde corre, presiona **Ctrl + C**.

## Acceso desde otros equipos (firewall)

Si en este PC funciona pero desde otro equipo no carga, el Firewall de Windows está bloqueando el puerto. En una PowerShell **como administrador**:

```powershell
New-NetFirewallRule -DisplayName "SmartDIM Streamlit 8501" -Direction Inbound -Protocol TCP -LocalPort 8501 -Action Allow -Profile Domain,Private
```

## Comprobar que está arriba

```powershell
Invoke-WebRequest http://localhost:8501/_stcore/health -UseBasicParsing
```

Debe responder `ok`.

## Problemas comunes

| Síntoma | Solución |
|---|---|
| `Port 8501 is already in use` | Ya hay otra instancia corriendo. Ciérrala, o busca el proceso: `Get-NetTCPConnection -LocalPort 8501 \| Select OwningProcess` y luego `Stop-Process -Id <PID>` |
| `streamlit : no se reconoce como comando` | El entorno virtual no está activado (paso 3), o usa `python -m streamlit run app.py ...` |
| `ModuleNotFoundError` | Falta instalar dependencias (paso 4) con el `.venv` activado |
| No carga desde otro PC | Revisa el firewall (sección anterior) y que usaste `--server.address 0.0.0.0` |

## Atajo sin activar el entorno

Desde la carpeta del proyecto, en una sola línea:

```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501 --server.headless true
```
