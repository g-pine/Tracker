# Bot de Gastos Personales — Multi Usuario

## Variables de entorno necesarias

| Variable | Descripción |
|----------|-------------|
| `TELEGRAM_TOKEN` | Token del bot de @BotFather |
| `SHEET_ID` | ID de tu Google Sheet |
| `GOOGLE_CREDENTIALS` | Contenido del credentials.json en una sola línea (para producción) |

## Correr localmente

1. Edita `.env` con tu token y Sheet ID
2. Pon tu `credentials.json` en esta carpeta
3. Instala dependencias:
```bash
pip install -r requirements.txt
```
4. Ejecuta:
```bash
python bot_gastos.py
```

## Desplegar en Railway

1. Sube el repositorio a GitHub (sin `.env` ni `credentials.json`)
2. En Railway agrega las variables de entorno:
   - `TELEGRAM_TOKEN`
   - `SHEET_ID`
   - `GOOGLE_CREDENTIALS` → pega el contenido completo de tu `credentials.json` aquí
3. Railway desplegará automáticamente

## Qué se arregló para producción

- ✅ Persistencia en disco (`bot_data.pkl`) — sobrevive reinicios
- ✅ Manejo de errores en todas las operaciones de Sheets
- ✅ Rate limiting (1 acción cada 2 segundos por usuario)
- ✅ Validación de inputs (montos, longitud de texto)
- ✅ Credenciales de Google desde variable de entorno (seguro para Railway)
- ✅ Error handler global con logging
- ✅ `credentials.json` en `.gitignore`
