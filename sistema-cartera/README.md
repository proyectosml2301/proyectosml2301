# Sistema de gestión de vencimientos de cartera de pólizas

Envía recordatorios automáticos por correo a clientes cuya póliza está
por vencer, corriendo desatendido en GitHub Actions.

## 🚀 Cómo ponerlo en marcha

### 1. Sube este proyecto a un repositorio de GitHub

```bash
cd sistema-cartera
git init
git add .
git commit -m "Sistema de gestión de vencimientos v2"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
git push -u origin main
```

### 2. Crea una contraseña de aplicación de Gmail

1. Ve a https://myaccount.google.com/apppasswords
2. Elige "Correo" → "Dispositivo: Otro" → ponle un nombre
3. Copia la contraseña de 16 dígitos (no es la contraseña normal de tu cuenta)

### 3. Configura los secretos en GitHub

En tu repo: **Settings → Secrets and variables → Actions → New repository secret**

| Nombre            | Valor                                |
|-------------------|---------------------------------------|
| `EMAIL_EMISOR`    | tu_correo@gmail.com                  |
| `EMAIL_PASSWORD`  | la contraseña de aplicación de 16 dígitos |

Nunca pongas estos valores directamente en el código ni en el repo.

### 4. Sube tu cartera real

Reemplaza `data/cartera.csv` con tu archivo real (puede ser `.csv` o `.xlsx`,
ajustando `CARTERA_PATH` en el workflow si cambias el nombre).

Columnas que el sistema reconoce automáticamente (no importa mayúsculas,
tildes ni guiones/espacios):

| Campo del sistema | Nombres aceptados en tu archivo |
|---|---|
| cliente | cliente, nombre, nombre_cliente, asegurado |
| email | email, correo, correo_electronico, mail |
| poliza | poliza, no_poliza, numero_poliza, n_poliza |
| vencimiento | vencimiento, fecha_vencimiento, fecha_vto |
| prima | prima, prima_anual, valor_prima, monto |
| aseguradora (opcional) | aseguradora, compania, compañia |

Si tu archivo usa otros nombres, agrégalos a `ALIASES_COLUMNAS` en
`gestion_cartera.py`.

### 5. Listo — corre solo

El workflow en `.github/workflows/recordatorios.yml` corre todos los días
a las 9am (hora Colombia). También puedes dispararlo manualmente desde
la pestaña **Actions → Recordatorios de vencimiento de pólizas → Run workflow**.

## 📅 Umbrales de aviso

Por defecto se notifica a los clientes 30, 15, 5 y 1 día antes del
vencimiento (variable `DIAS_AVISO` en el workflow). Cada umbral se envía
una sola vez por póliza — el sistema lleva el control en
`logs/historial_envios.csv` para no duplicar correos aunque el workflow
corra varias veces.

## 💻 Probarlo en tu computador antes de subirlo

```bash
pip install -r requirements.txt
cp .env.example .env       # y completa tus datos reales
export $(cat .env | xargs) # carga las variables (en Windows usa otra forma)
python gestion_cartera.py --archivo data/cartera.csv
```

Si no defines `EMAIL_EMISOR`/`EMAIL_PASSWORD` y corres el script en una
terminal interactiva, te las pedirá por teclado — pero en GitHub Actions
**siempre** debe usar los secretos, nunca pide nada por consola.

## 📁 Estructura

```
sistema-cartera/
├── gestion_cartera.py              # script principal
├── requirements.txt
├── .env.example
├── data/
│   └── cartera.csv                 # tu cartera (reemplázala)
├── logs/
│   ├── sistema.log                 # log de ejecución
│   └── historial_envios.csv        # control de envíos (anti-duplicados)
└── .github/workflows/
    └── recordatorios.yml           # automatización diaria
```

## 🔧 Qué cambió respecto a la v1 (notebook)

- ✅ Ya no requiere `input()`/`getpass()` para correr — usable en CI/CD
- ✅ Detecta las columnas del archivo automáticamente (sin editar código)
- ✅ Soporta CSV y Excel
- ✅ Múltiples avisos (30/15/5/1 días) en vez de uno solo
- ✅ No reenvía un mismo aviso dos veces (control vía historial)
- ✅ Valida formato de email y descarta filas con datos inválidos
- ✅ Logging real a archivo, no solo `print()`
- ✅ Sin rutas hardcodeadas de Google Colab/Drive
