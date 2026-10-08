# Monitor de Renta Fija Argentina

Planilla de Google Sheets que se actualiza sola durante la rueda y muestra, para el mercado argentino:

- **Soberanos en USD**: matriz de retornos según la TIR de salida a los próximos dos cierres de semestre.
- **Riesgo País**: retornos de Globales (ley NY) y Bonares (ley AR) ante distintos escenarios de riesgo país.
- **Carry trade**: retorno en USD de cada letra según el tipo de cambio al vencimiento y dólar breakeven.

👉 **[Ver la planilla en vivo](https://docs.google.com/spreadsheets/d/15XMFrllATr5c0XjYJAk8p-aHXlKMM7F09t_hEUgt_VE/edit?usp=sharing)** (solo lectura)

![Carry trade: retorno en USD por LECAP según el dólar al vencimiento](carry.png)

## Cómo funciona

Todo el procesamiento está en un único script de Google Apps Script ([`Monitor_Mercado.gs`](Monitor_Mercado.gs)), organizado en tres capas:

| Capa | Qué hace |
|---|---|
| **Fetch** | Trae precios y datos de las APIs; no calcula nada. |
| **Cálculo** | Funciones puras (TIR, valor presente, retornos, curva). Se pueden portar tal cual a otro entorno. |
| **Escritura** | Arma las hojas, formatos y gráficos. Es lo único que toca Google Sheets. |

Un activador corre la actualización cada 15 minutos en días hábiles, de 11 a 18 h.

### Fuentes de datos

| Dato | Fuente |
|---|---|
| Precios de bonos, LECAPs y BONCAPs (CI y 24hs) | BYMA, datos públicos (demora aprox. 20 min) |
| Flujos de fondos de bonos en USD | A3 Mercados; para los bonos que no publica, se generan a partir de las condiciones de emisión |
| Condiciones de LECAPs (emisión, TEM) y bonos nuevos | Fichas técnicas de BYMA |
| Tasa del Tesoro de EE.UU. a 10 años | FRED (Reserva Federal de St. Louis), serie DGS10 |
| Riesgo país | ArgentinaDatos (EMBI, dato diario) |

### Metodología

- **TIR**: efectiva anual, base ACT/365, sobre precio sucio y flujos por 100 VN original. Validada contra la TIR publicada por A3 (diferencia de 0–1 pb).
- **Retorno al horizonte**: (precio teórico en la fecha de salida + flujos cobrados hasta esa fecha) / precio de hoy − 1. Sin reinversión de cupones.
- **Escenarios de riesgo país**: cada bono parte de su TIR actual y se desplaza lo mismo que cambia el riesgo país (desplazamiento paralelo). Supone que la forma de la curva, el spread de ley y la tasa de EE.UU. no cambian.
- **Valor final de LECAPs**: 100 × (1 + TEM)^n, con n = meses enteros desde la emisión + días restantes / 30.
- **Dólar breakeven**: MEP × (1 + rendimiento), con MEP = AL30 / AL30D.
- **Altas y bajas automáticas** (carry trade): las letras que vencen salen solas; las LECAPs y BONCAPs nuevas se detectan en BYMA y se dan de alta desde su ficha técnica.

## Instalación (para replicarlo)

1. Crear una planilla de Google Sheets → **Extensiones → Apps Script**.
2. Pegar `Monitor_Mercado.gs` y reemplazar el manifiesto por `appsscript.json` (activa el servicio avanzado de Sheets).
3. Ejecutar `configurarInicial()`, después `actualizar()` y por último `instalarTrigger()`.

Los parámetros (escenarios de riesgo país, grilla de TIR, grilla de tipo de cambio, tickers) se editan en la hoja `Config`.

## Aviso

Proyecto personal con fines informativos y educativos. No constituye recomendación de inversión. Los datos provienen de fuentes públicas, pueden tener demoras o errores y están sujetos a los términos de uso de cada proveedor.
