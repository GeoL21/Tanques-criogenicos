# Tanques criogénicos · Clínica Reina Catalina

Registro automático del nivel de los tanques criogénicos de Barranquilla y Baranoa.

1. El técnico sube la foto del manómetro de nivel al tema de su sede en el grupo de Telegram **Gases Medicinales**.
2. El bot **Tanque criogenico** la recoge (cada 5 minutos), Gemini lee la aguja y la lectura queda en la hoja **Tanques criogénicos** (cuenta cilindrocrc).
3. El bot responde debajo de la foto con el reporte: nivel, consumo o recarga, promedio de 7 días y cuántos días faltan para el nivel de recarga.
4. Este repositorio es el **tablero (PWA)** para consultar todo desde el celular o el PC.

| Sede | Capacidad | Escala del manómetro | Nivel de recarga |
|---|---|---|---|
| Barranquilla | 136 pulg | 0–200 pulg | 40 pulg |
| Baranoa | 80 pulg | 0–100 pulg | 24 pulg |

## Archivos

- `index.html`, `sw.js`, `manifest.webmanifest`, íconos: el tablero.
- `config.js`: dirección de la implementación de Apps Script.
- `apps-script/`: código del bot y de la hoja (se copia en el Apps Script de la hoja).

Vista de demostración con datos de ejemplo: agrega `?demo=1` a la dirección del tablero.

## Instalación (una sola vez)

### 1. Clave de Gemini
En **aistudio.google.com** (con la cuenta cilindrocrc) → *Get API key* → *Create API key*. Copiarla.

### 2. Hoja y Apps Script
1. En el Drive de cilindrocrc, crear una hoja nueva llamada **Tanques criogénicos**.
2. *Extensiones → Apps Script*. Crear dos archivos, `Code.gs` y `Calculos.gs`, y pegar el contenido de la carpeta `apps-script/`.
3. *Servicios (+)* → agregar **Google Sheets API**.
4. *Configuración del proyecto → Propiedades del script*, agregar:
   - `TELEGRAM_TOKEN`: la clave del bot que dio @BotFather.
   - `GEMINI_KEY`: la clave de Gemini.
   - `CODIGO_PWA`: el código de acceso que se escribirá en el tablero.
5. Elegir la función **configurar** y tocar *Ejecutar*. Aceptar los permisos. Crea las pestañas, las tablas, la carpeta de fotos y la revisión cada 5 minutos.
6. *Implementar → Nueva implementación → Aplicación web*. Ejecutar como: **yo**. Acceso: **cualquier usuario**. Copiar la dirección que termina en `/exec` y pegarla en `config.js`.

### 3. Prueba
Subir una foto al tema de cada sede. En máximo 5 minutos el bot responde debajo. Para no esperar: en la hoja, menú **Tanques → Revisar Telegram ahora**.

Si el bot responde que no sabe de qué tanque es el tema, escribir en ese tema `/sede Barranquilla` o `/sede Baranoa`.

## Uso diario

- `/estado` en un tema: el bot responde con el estado de ese tanque.
- Lectura **Dudosa** en la hoja: la IA no pudo leer bien la foto. Escribir el nivel correcto en *Nivel (pulg)*, cambiar *Lectura* a **Corregida** y usar **Tanques → Recalcular todo**.
- Para cambiar capacidad o nivel de recarga: pestaña **Config**.
