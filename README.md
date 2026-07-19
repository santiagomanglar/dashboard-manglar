# Dashboard Manglar

Tablero financiero para los socios, construido en Streamlit sobre el modelo financiero.

---

## Correr en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

La aplicación busca `Manglar_Modelo_Financiero.xlsx` en la misma carpeta. Si no lo
encuentra, muestra un cargador de archivos para subirlo desde el navegador.

---

## Qué muestra

- **Indicadores:** saldo en caja, ingresos netos de IVA, margen bruto y resultado operativo
- **Resultado y caja:** ingresos contra costos por mes, y evolución del saldo
- **Rentabilidad:** ingresos y neto por cliente, concentración, y detalle por proyecto
- **Composición del gasto:** salidas por cuenta
- **Alerta** cuando hay movimientos sin clasificar

Los filtros de mes y cliente afectan todo el tablero.

---

## Criterios de cálculo

Replican los del modelo, para que los números coincidan:

- El **P&G** usa el mes de causación; el **flujo de caja** usa el mes de caja
- Los ingresos y costos marcados con IVA se dividen entre 1,19 para dejarlos netos
- La cuenta de cada movimiento se recalcula desde la hoja `Data`, no desde la fórmula
  guardada, para evitar valores desactualizados
- Los movimientos sin mes de causación quedan fuera del P&G, tal como en el modelo

---

## Publicar

### Opción recomendada: Streamlit Community Cloud

Es gratis y se conecta a un repositorio de GitHub.

1. Subir esta carpeta a un repositorio
2. En share.streamlit.io, conectar el repositorio y señalar `app.py`
3. Definir quién puede entrar

**Importante sobre la confidencialidad.** En el plan gratuito el repositorio debe ser
público, así que **el modelo financiero no se debe subir al repositorio**. Dos caminos:

- **Carga manual:** no incluir el Excel. Al abrir el tablero, se sube el archivo desde
  el navegador. Simple y sin credenciales, pero hay que repetirlo en cada sesión.
- **Google Sheets en vivo:** dejar el modelo en Drive y leerlo con una cuenta de
  servicio, guardando las credenciales en los secretos de Streamlit (nunca en el
  repositorio). Los socios ven siempre la última versión.

En ambos casos conviene restringir el acceso a la aplicación a correos concretos.
Un tablero con los financieros de la empresa no debería quedar abierto en internet.

### Conectar a Google Sheets

1. Crear una cuenta de servicio en Google Cloud y habilitar la API de Sheets
2. Compartir el modelo con el correo de esa cuenta de servicio (permiso de lectura)
3. Guardar las credenciales en `.streamlit/secrets.toml` (local) y en los secretos de
   Streamlit Cloud:

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "...@....iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"

[sheets]
url = "https://docs.google.com/spreadsheets/d/ID_DEL_ARCHIVO"
```

4. Agregar `gspread` y `google-auth` a `requirements.txt`, y reemplazar la función
   `cargar` de `app.py` por una que lea las hojas con `gspread`.

### Otras alternativas

| Opción | Costo | Datos | Control de acceso |
|---|---|---|---|
| Streamlit Community Cloud | Gratis | Carga manual o Sheets | Por correo |
| Looker Studio | Gratis | En vivo desde Sheets | Permisos de Google |
| Vercel o Netlify | Gratis, clave se paga | Hay que regenerar | Público salvo pago |
| Pestaña en el propio Sheets | Gratis | En vivo | Hereda los del archivo |

---

## Mantenimiento

El tablero no requiere ajustes cuando se agregan movimientos: lee el modelo y recalcula.
Solo hay que tocarlo si cambian los nombres de las hojas o de las columnas.
