# Herramientas GE&PE — instalación en el PC de la oficina

Guía para dejar la plataforma **Herramientas GE&PE** (revisión de servicios de ajuste y
consultas a Gemweb) funcionando en el PC de la oficina que está siempre encendido, de forma
que los 9 compañeros la usen desde el navegador, en la oficina o por VPN, con su usuario y
contraseña y sin instalar nada en sus ordenadores. Sustituye al extractor de Gemweb que hoy
funciona en ese mismo PC.

```
 Oficina (red local)            Teletrabajo (VPN)
  navegador ──┐                  navegador ──┐
              ├──  http://PC-OFICINA:8501  ────┤
         ┌────┴──────────────────────────────┴────┐
         │ PC de la oficina (siempre encendido)   │
         │  · Plataforma (Python + Streamlit)     │
         │  · Base de datos revisor_ssaa.db       │
         │    (usuarios, fichas, historial)       │
         │  · Excel de ESIOS y sus descargas      │
         │  · Credenciales de ESIOS y Gemweb      │
         └────────────────────────────────────────┘
```

Coste: ninguno. Todo es software libre (Python, Streamlit, SQLite) y corre en el servidor.

## 0. Retirar el extractor de Gemweb antiguo

La plataforma usa el mismo puerto (8501) que el extractor de Gemweb que funciona hoy en el
PC de la oficina, y lo incluye entero (sección «Gemweb»). Antes de instalar:

1. Ver qué programa usa el puerto:
   ```powershell
   Get-NetTCPConnection -LocalPort 8501 -State Listen | ForEach-Object { Get-Process -Id $_.OwningProcess }
   ```
2. Cerrarlo (la ventana de consola donde corre `streamlit run app.py`, o
   `Stop-Process -Id <PID>`).
3. Quitar su arranque automático, según cómo se arranque hoy: acceso directo en la carpeta
   de Inicio (`shell:startup`), tarea del Programador de tareas o fichero `.bat`.
4. **No borrar** la carpeta `API_Gemweb`: su `.streamlit\secrets.toml` sirve a la plataforma
   como credenciales de Gemweb de respaldo si se deja al lado de `Herramientas_GEYPE`.

El script de instalación se niega a seguir si el puerto 8501 sigue ocupado.

## 1. Requisitos

- Windows (servidor o PC) **siempre encendido**, en la red de la oficina.
- **Python 3.11 o posterior** (desarrollado con 3.13), instalado para todos los usuarios.
- Unos 2 GB libres de disco.
- Salida a internet hacia `api.esios.ree.es` (datos de REE) y `api.gemweb.es` (curvas).
- Una **cuenta de Windows para el servicio** (por ejemplo `GEYPE\svc_herramientas`), sin
  privilegios de administrador, con lectura y escritura en las carpetas del punto 2.

## 2. Carpetas

Las dos carpetas tienen que estar **una al lado de la otra**:

```
D:\GEyPE\                                  (o la ruta que se prefiera)
├── Herramientas_GEYPE\                    la plataforma (repositorio de GitHub)
└── Descarga_datos_ESIOS\                  scripts y Excel de ESIOS
```

1. App: `git clone https://github.com/iyadazig/Herramientas-GE-PE.git Herramientas_GEYPE`
2. Datos de ESIOS: copiar la carpeta `Descarga_datos_ESIOS` completa desde el ordenador
   de Noelia (scripts, Excel históricos y `esios_token.txt`).
3. Fichas de contrato ya existentes: copiar `contratos_ssaa.json` del ordenador de Noelia a
   `Herramientas_GEYPE\`. Al primer arranque se importan a la base de datos.

> `contratos_ssaa.json`, la base de datos y los Excel contienen **datos de clientes**:
> copiarlos por una carpeta compartida interna, nunca por correo ni servicios externos.

## 3. Librerías de Python

En una consola, desde `Herramientas_GEYPE`:

```bat
python -m pip install -r requirements.txt
```

## 4. Credenciales

- **ESIOS**: el fichero `Descarga_datos_ESIOS\esios_token.txt` (o la variable de entorno
  `ESIOS_TOKEN` de la cuenta del servicio).
- **Gemweb**: variables de entorno `GEMWEB_CLIENT_ID` y `GEMWEB_CLIENT_SECRET` de la cuenta
  del servicio, **o** después del primer arranque, un administrador las introduce en la app
  (barra lateral → Gemweb → Configurar credenciales); quedan cifradas en el perfil de esa
  cuenta.

Las credenciales no van nunca al repositorio de GitHub.

## 5. Instalación (cortafuegos y arranque automático)

En PowerShell **como administrador**, desde `Herramientas_GEYPE`:

```powershell
powershell -ExecutionPolicy Bypass -File servidor\instalar_servidor.ps1 `
    -Subredes "192.168.0.0/24,10.8.0.0/24" -Usuario "GEYPE\svc_herramientas"
```

- `-Subredes`: la red de la oficina **y** el rango de direcciones de la VPN. Solo esas
  direcciones podrán abrir la app; desde internet no es accesible.
- `-Usuario`: la cuenta del servicio (pedirá su contraseña).
- Si había una instalación anterior del revisor de SSAA (tareas «Revisor SSAA - …»), el
  script la sustituye.
- Opcionales: `-Puerto 8501`, `-Python "C:\ruta\python.exe"`, `-CarpetaEsios`, `-CarpetaCopias`.

El script crea:

| Tarea programada | Cuándo | Qué hace |
|---|---|---|
| Herramientas GEYPE - servidor | Al encender (y se reinicia si se cae) | La plataforma en el puerto 8501 |
| Herramientas GEYPE - descarga PVPC diaria | Todos los días, 08:30 | Total SAH del PVPC_DETALLE |
| Herramientas GEYPE - descarga componentes ESIOS | Lunes, 09:00 | Componentes, PFMHORAS_COM (y su C2), pérdidas |
| Herramientas GEYPE - copia de seguridad | Todos los días, 23:00 | Copia de la base de datos (se guardan 30) |

## 6. VPN

La VPN da acceso a las carpetas del servidor (puerto 445), pero la app usa **otro puerto**:
hay que permitir en la VPN el tráfico **TCP 8501** hacia el servidor. Comprobación desde un
equipo conectado por VPN:

```powershell
Test-NetConnection PC-OFICINA -Port 8501
```

(`TcpTestSucceeded : True`).

## 7. Primer acceso y usuarios

1. Abrir `http://PC-OFICINA:8501` (en el servidor, `http://localhost:8501`).
2. La primera vez la app pide **crear el administrador** (usuario, nombre y contraseña).
3. El administrador da de alta al resto en **Usuarios**: la app genera una contraseña
   provisional que hay que darles en persona o por teléfono; al entrar, cada uno la cambia.
4. Contraseñas: mínimo 10 caracteres, con letras y números. Tras 5 intentos fallidos el
   usuario queda bloqueado 15 minutos. Un administrador puede restablecer contraseñas y
   desactivar usuarios (por ejemplo, si alguien deja la empresa).

A cada compañero basta con darle el enlace `http://PC-OFICINA:8501` para guardarlo en
favoritos.

## 8. Comprobar que todo funciona

```bat
python lanzador.py --autoprueba
type autoprueba_resultado.txt
```

Todas las líneas deben empezar por `OK` (lectura de ESIOS, cálculo, gráfica, informe,
fichas, lectura de PDF).

## 9. Actualizar la aplicación

```bat
cd D:\GEyPE\Herramientas_GEYPE
git pull
python -m pip install -r requirements.txt
schtasks /End /TN "Herramientas GEYPE - servidor"
schtasks /Run /TN "Herramientas GEYPE - servidor"
```

Los compañeros solo tienen que recargar la página.

## 10. Copias de seguridad

- Cada noche se copia `revisor_ssaa.db` en `Herramientas_GEYPE\copias_seguridad\`
  (se conservan las 30 últimas). Conviene incluir esa carpeta en la copia de seguridad
  general de la empresa.
- Copia manual: `python almacen.py --copia D:\ruta\de\copias`
- **Restaurar**: parar la tarea del servidor, sustituir `revisor_ssaa.db` por la copia
  elegida (renombrándola) y volver a arrancar la tarea.

## 11. HTTPS (recomendado)

Por la VPN el tráfico ya va cifrado; dentro de la oficina, sin HTTPS, las contraseñas
viajan en claro por la red local. Para activarlo, con un certificado de la CA interna (o
uno autofirmado) en formato PEM:

1. Definir para la cuenta del servicio las variables `SSAA_SSL_CERT` (ruta del certificado)
   y `SSAA_SSL_KEY` (ruta de la clave privada).
2. Reiniciar la tarea «Herramientas GEYPE - servidor».
3. La dirección pasa a ser `https://PC-OFICINA:8501`.

## 12. Problemas frecuentes

| Síntoma | Causa probable |
|---|---|
| Desde la oficina no abre | Tarea parada (ver el Programador de tareas) o regla del cortafuegos con otra subred |
| Desde casa no abre y en la oficina sí | La VPN no deja pasar el puerto 8501 (punto 6) |
| «Faltan días de … en los Excel» | Las tareas de descarga de ESIOS no se están ejecutando o falta `esios_token.txt` |
| Gemweb sin credenciales | Punto 4 |
| Olvido de contraseña | Un administrador la restablece en «Usuarios» |
| Nadie es administrador | No puede pasar: la app impide quitar o desactivar al último administrador |
