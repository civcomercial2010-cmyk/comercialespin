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
