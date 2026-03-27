# comercialespin

Cuadro de mando comercial (GitHub Pages) y actualización vía GitHub Actions.

## Un solo clon local

Este proyecto se trabaja desde **una carpeta = un `git clone`**. No hace falta mantener copias paralelas del mismo código: la referencia compartida es la rama **`main`** en GitHub.

- **GitHub Pages** y quien abra el repo ven lo que está en **`origin/main`** tras el último push.
- Tu carpeta local es el sitio donde editas; **sincronizar** = `git pull` / `git push`, no copiar archivos a mano entre carpetas.

### Rutina recomendada

1. **Antes de editar** (o al empezar el día):

   ```bash
   git pull origin main
   ```

2. **Cuando tengas cambios listos**:

   ```bash
   git add -A
   git status
   git commit -m "Describe el cambio en una frase"
   git push origin main
   ```

3. Tras el push, espera 1–3 minutos y recarga el cuadro en el navegador con **Ctrl+F5** si no ves cambios (caché).

### Si algo choca al hacer pull

Si Git avisa de cambios locales que entran en conflicto con el remoto, resuelve los archivos marcados, luego `git add` y `git commit` (o usa `git pull --rebase origin main` si preferís historial lineal).

---

Archivos como `__pycache__/` o `*.log` están en `.gitignore` para no ensuciar commits ni el historial.

## Actualización automática desde Gmail (Actions)

- El workflow **Actualizar cuadro de mando** está programado **de lunes a viernes** (no sábado ni domingo). Si el informe solo llega un fin de semana, no habrá ejecución automática hasta el siguiente día laborable salvo que lances el workflow a mano.
- Si el correo llegó **después** de la ventana ~20:30 (España) de ese día, el job ya habrá corrido con el Excel anterior; en ese caso usa **Run workflow** en GitHub (pestaña *Actions*) para procesar de inmediato.
- Revisa en *Actions* el último run: si el paso del extractor no se ejecutó o falló (IMAP, adjunto, secrets `GMAIL_*`, `ASUNTO_FILTRO`, `REMITENTE`), el `data.json` no cambiará.
- El job programado solo continúa si, al ejecutarse el paso de hora en **Europe/Madrid**, la hora cae en la ventana **20:00–21:59** (para tolerar retrasos de la cola de GitHub). Si el log muestra p. ej. `21:00` y antes la ventana era demasiado estrecha, el extractor no llegaba a lanzarse.
- El extractor elige el adjunto con la **fecha y hora de generación del informe** leídas del Excel (texto tipo `Fecha: 27/03/26 Hora: 20:21:00`, celda fecha/hora, etc.), no la hora de recepción en Gmail. El candidato **más reciente** gana; si hubiera empate exacto, se usa el **UID IMAP** más alto. La lectura se hace en la hoja **VENTAS**.

