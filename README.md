# 💊 FarmaAnalytics — Guía de despliegue en Railway.app

Panel profesional de análisis de ventas para farmacias.
Login por farmacia · Carga de Excel · Dashboard con 7 secciones de análisis.

---

## 📁 Archivos del proyecto

| Archivo | Para qué sirve |
|---|---|
| `app.py` | Aplicación principal (toda la lógica del dashboard) |
| `auth.py` | Sistema de login y gestión de usuarios |
| `data_processor.py` | Procesamiento del archivo Excel de ventas |
| `config.yaml` | **Usuarios y contraseñas** — aquí añades farmacias |
| `requirements.txt` | Librerías Python necesarias (Railway las instala solo) |
| `railway.toml` | Instrucciones de despliegue para Railway |
| `.streamlit/config.toml` | Tema visual oscuro de la app |

> **No necesitas tocar ningún archivo** excepto `config.yaml` para añadir farmacias.

---

## 🚀 Despliegue paso a paso (sin conocimientos técnicos)

### PASO 1 — Crear una cuenta gratuita en GitHub

GitHub es la plataforma donde guardarás los archivos del proyecto.
Railway los leerá directamente desde ahí.

1. Abre el navegador y ve a **https://github.com**
2. Haz clic en el botón **"Sign up"** (esquina superior derecha)
3. Introduce tu **dirección de email** y haz clic en **"Continue"**
4. Crea una **contraseña** (mínimo 8 caracteres) y haz clic en **"Continue"**
5. Escoge un **nombre de usuario** (por ejemplo: `mifarmacia2024`) y haz clic en **"Continue"**
6. Resuelve el **captcha** de verificación
7. GitHub te enviará un **código de 6 dígitos** al email — introdúcelo en la pantalla
8. En la siguiente pantalla, selecciona el plan **"Free"** (gratuito) y haz clic en **"Continue for free"**

✅ Ya tienes cuenta en GitHub.

---

### PASO 2 — Crear el repositorio y subir los archivos

Un "repositorio" es simplemente una carpeta en GitHub donde viven tus archivos.

**2.1 — Crear el repositorio:**

1. Una vez dentro de GitHub, haz clic en el botón verde **"New"**
   (está en la columna izquierda, o en el menú `+` de la esquina superior derecha → "New repository")
2. En el campo **"Repository name"** escribe exactamente: `farmaanalytics`
3. Deja seleccionado **"Public"** (gratuito; Railway necesita acceso al repositorio)
4. **No marques** ninguna de las opciones adicionales (Add README, .gitignore, license)
5. Haz clic en el botón verde **"Create repository"**

**2.2 — Subir los archivos:**

En la página que se abre verás instrucciones. Busca el enlace que dice
**"uploading an existing file"** y haz clic en él.

Ahora arrastra y suelta **todos** estos archivos al área central de la pantalla:

```
app.py
auth.py
data_processor.py
config.yaml
requirements.txt
railway.toml
.gitignore
```

> ⚠️ **El archivo `.streamlit/config.toml`** vive dentro de una carpeta y hay que
> subirlo de forma distinta:
>
> 1. Haz clic en **"Add file" → "Create new file"**
> 2. En el campo del nombre escribe: `.streamlit/config.toml`
>    (GitHub creará la carpeta automáticamente al escribir la barra)
> 3. Copia y pega el contenido del archivo `config.toml` en el área de texto
> 4. Haz clic en **"Commit new file"**

Una vez que tengas todos los archivos, baja al final de la página de carga,
escribe un mensaje como `"Primer despliegue"` en el campo **"Commit changes"**
y haz clic en el botón verde **"Commit changes"**.

✅ Tus archivos ya están en GitHub.

---

### PASO 3 — Crear una cuenta en Railway

Railway es la plataforma que ejecutará tu app en Internet.

1. Ve a **https://railway.app**
2. Haz clic en **"Login"** (esquina superior derecha)
3. Selecciona **"Login with GitHub"**
4. GitHub te pedirá que autorices a Railway — haz clic en **"Authorize Railway"**
5. Railway creará tu cuenta automáticamente vinculada a tu GitHub

✅ Ya tienes cuenta en Railway.

---

### PASO 4 — Crear el proyecto y desplegar la app

1. En el panel de Railway, haz clic en **"New Project"**
2. Selecciona **"Deploy from GitHub repo"**
3. Si es la primera vez, Railway te pedirá permiso para acceder a tus repositorios:
   - Haz clic en **"Configure GitHub App"**
   - Selecciona tu cuenta de GitHub
   - Selecciona **"All repositories"** (o solo `farmaanalytics`)
   - Haz clic en **"Save"**
4. Vuelve a Railway y busca el repositorio **`farmaanalytics`** en la lista
5. Haz clic en él — Railway detectará automáticamente el archivo `railway.toml`
   y comenzará a construir la app

**Durante el despliegue** verás una pantalla con logs en tiempo real.
Es normal que tarde **entre 3 y 7 minutos** la primera vez
(Railway instala todas las librerías de Python automáticamente).

Cuando el indicador cambie a verde y aparezca **"Deploy successful"**, la app está lista.

✅ Tu app está desplegada.

---

### PASO 5 — Obtener la URL pública

1. En Railway, dentro de tu proyecto, haz clic en el **bloque del servicio**
   (el recuadro con el nombre del repositorio)
2. Ve a la pestaña **"Settings"**
3. Baja hasta la sección **"Networking"**
4. Haz clic en **"Generate Domain"**
5. Railway generará una URL del tipo:
   ```
   https://farmaanalytics-production-xxxx.up.railway.app
   ```

📌 **Guarda esta URL** — es la dirección que compartirás con las farmacias.

> Puedes personalizar el subdominio haciendo clic en el lápiz junto a la URL
> y escribiendo el nombre que prefieras (ej: `analisis-mifarmacia`).

---

### PASO 6 — Primer acceso y verificación

Abre la URL en el navegador. Verás la pantalla de login de FarmaAnalytics.

Prueba con las credenciales iniciales:

| Usuario | Contraseña | Tipo de acceso |
|---|---|---|
| `admin` | `admin123` | Administrador |
| `farmacia1` | `farm1pass` | Farmacia |
| `farmacia2` | `farm2pass` | Farmacia |
| `farmacia3` | `farm3pass` | Farmacia |

> 🔒 **Seguridad:** Las contraseñas se hashean automáticamente con bcrypt
> la primera vez que arranca la app. Después de ese primer arranque
> el archivo `config.yaml` en Railway tendrá los hashes cifrados —
> nadie puede leer las contraseñas originales aunque acceda al repositorio.

---

## ✏️ Cómo añadir nuevas farmacias

Todo se gestiona editando el archivo `config.yaml` directamente en GitHub.
No necesitas instalar nada en tu ordenador.

### Añadir una farmacia nueva

1. Ve a tu repositorio en GitHub:
   `https://github.com/TU_USUARIO/farmaanalytics`
2. Haz clic en el archivo **`config.yaml`**
3. Haz clic en el icono del **lápiz** (✏️ "Edit this file") — esquina superior derecha
4. Añade el bloque de la nueva farmacia dentro de `credentials → usernames`,
   respetando exactamente la misma indentación (2 espacios) que los demás usuarios:

```yaml
credentials:
  usernames:

    admin:
      email: admin@farmacia.com
      name: Administrador
      password: admin123
      role: admin

    farmacia1:
      email: farmacia1@farmacia.com
      name: Farmacia 1
      password: farm1pass
      role: farmacia

    # ── Añade la nueva farmacia aquí ─────────────────────────────────────
    farmacia_central:                       # nombre de usuario (sin espacios ni tildes)
      email: central@mifarmacia.com         # email del responsable
      name: Farmacia Central Málaga         # nombre visible en el panel
      password: contraseñaNueva2024         # contraseña en texto plano (se cifra sola)
      role: farmacia                        # opciones: farmacia  |  admin
```

5. Baja al final, escribe `"Añadir Farmacia Central"` en el campo de commit
6. Haz clic en **"Commit changes"**
7. Railway detectará el cambio y redesplegará la app automáticamente (~2 min)

La próxima vez que la app arranque cifrará la contraseña nueva automáticamente.

### Cambiar la contraseña de una farmacia

1. Abre `config.yaml` en GitHub con el lápiz (✏️)
2. Busca el usuario y sustituye el valor de `password` por la nueva contraseña en texto plano
3. Guarda con **"Commit changes"**
4. Railway redespleará y cifrará la nueva contraseña en el siguiente arranque

### Eliminar una farmacia

1. Abre `config.yaml` en GitHub con el lápiz (✏️)
2. Borra las **5 líneas completas** del bloque del usuario
   (desde el nombre de usuario `farmaciaX:` hasta la línea `role: farmacia` inclusive)
3. Guarda con **"Commit changes"**

---

## 🔑 Cambiar la clave secreta (recomendado antes del uso real)

El archivo `config.yaml` tiene una clave secreta que protege las cookies de sesión.
**Cámbiala antes de desplegar en producción:**

```yaml
cookie:
  expiry_days: 30
  key: ESCRIBE_AQUI_UNA_FRASE_LARGA_Y_UNICA_2024   # ← cambia esto
  name: farmaanalytics_session
```

Puedes poner cualquier texto largo, por ejemplo:
`FarmaApp_ClaveSecreta_MiNombreDeFarmacia_XkZ92024`

---

## 📊 Formato del Excel de ventas

El archivo `.xlsx` que carga cada farmacia debe tener estas columnas
(el orden no importa, los nombres deben coincidir exactamente):

| Columna | Tipo | Ejemplo |
|---|---|---|
| `Fecha` | Texto o Fecha | `01/01/2024` |
| `Fecha_ES` | Fecha | `01/01/2024` |
| `Hora` | Texto `HH:MM:SS` | `10:30:00` |
| `Tipo de Operación` | Texto | `VENTA` |
| `Empresa` | Texto | `FARMACIA CENTRAL` |
| `Código` | Número | `654321` |
| `Denominación` | Texto | `PARACETAMOL 1G` |
| `Organismo` | Texto | `001 - VTA. LIBRE` |
| `Cantidad (Unidades)` | Número | `2` |
| `Pvp` | Número | `3.50` |
| `PVP Facturado` | Número | `3.50` |
| `Importe Bruto` | Número | `7.00` |
| `Descuento` | Número | `0.00` |
| `Importe Neto` | Número | `7.00` |
| `Cliente` | Texto | `CLI001` |
| `Vendedor` | Texto | `MARIA` |

> Las columnas `Existencias Anteriores` y `Existencias Posteriores` se ignoran
> aunque estén presentes en el archivo.

---

## 📦 Plan de Railway necesario

| Plan | Precio | Adecuado para |
|---|---|---|
| **Hobby Trial** | Gratis | Pruebas iniciales (500 horas/mes, puede pausarse) |
| **Hobby** | ~5 $/mes | Uso diario continuo — **recomendado para producción** |
| **Pro** | ~20 $/mes | Múltiples farmacias con uso muy intensivo |

> El plan gratuito puede pausar la app tras períodos de inactividad.
> Para uso diario en producción se recomienda el plan **Hobby** (~5 $/mes).

---

## ❓ Solución de problemas

**El despliegue falla (aparece en rojo)**
→ En Railway ve a la pestaña **"Logs"** y busca la línea que empieza con `ERROR`.
Los errores más comunes son: archivo faltante, indentación incorrecta en `config.yaml`,
o nombre de columna del Excel que no coincide exactamente.

**La app aparece en blanco o da error 502**
→ Espera 1-2 minutos más y recarga la página.
Si persiste, en Railway haz clic en **"Redeploy"**.

**No puedo hacer login**
→ Verifica que el nombre de usuario en `config.yaml` está en minúsculas
y que la contraseña no tiene espacios al principio o al final.

**Subí el Excel pero no aparecen datos**
→ Comprueba que las columnas del Excel tienen exactamente los nombres
de la tabla de formato (mayúsculas, tildes y espacios incluidos).
El error más frecuente es `Fecha_ES` escrito como `FechaES` o `Fecha ES`.

**Quiero actualizar la app después de hacer cambios en el código**
→ Sube el archivo modificado a GitHub (mismo proceso que en el Paso 2).
Railway detectará el cambio y redesplegar automáticamente en ~2 min.

---

## 🔒 Privacidad de los datos

Los datos de ventas **no se almacenan en ningún servidor**.
Se procesan únicamente en memoria durante la sesión activa de cada usuario.
Al cerrar el navegador o hacer logout, los datos desaparecen completamente.

---

*FarmaAnalytics — desarrollado con Streamlit · Plotly · Python 3.11*
