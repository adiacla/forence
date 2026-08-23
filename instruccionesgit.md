Como el repositorio remoto ya tiene un archivo (`README.md`), Git rechazará el subido inicial si no conectas y sincronizas ambos entornos primero.

Abre la terminal integrada en VS Code (`Ctrl + ~` en Windows/Linux o `Cmd + ~` en Mac) y ejecuta los siguientes comandos en orden:

1. **Inicializar y vincular el repositorio:** Ejecutar dentro de la carpeta local.
Inicializa Git en tu carpeta local y agrégale la dirección de tu repositorio remoto en GitHub:

```bash
git init
git remote add origin https://github.com/adiacla/forence.git

```


2. **Preparar los archivos locales:** Añadir y confirmar cambios.
Guarda el estado actual de todos tus archivos locales (incluyendo tu `README.md` completo):

```bash
git add .
git commit -m "Primer commit: subiendo archivos locales"

```


3. **Sincronizar historiales independientes:** Paso clave para evitar conflictos.
Dado que GitHub creó un `README.md` inicial en la nube y tú tienes tus propios archivos en local, debes fusionar ambos historiales permitiendo historias no relacionadas y dándole prioridad a tu versión local:

```bash
git pull origin main --allow-unrelated-histories -X ours

```

*(Nota: Si tu rama principal en GitHub se llama `master` en lugar de `main`, cambia `main` por `master` en el comando).*


4. **Subir el código a GitHub:** Enviar cambios al servidor remoto.
Finalmente, envía todos tus archivos a GitHub y establece la rama por defecto:

```bash
git push -u origin main

```





Para subir tus próximos cambios en el futuro, solo necesitas ejecutar 3 comandos básicos en la terminal de VS Code cada vez que termines de hacer modificaciones:

1. **Guardar los cambios modificados o nuevos:**
```bash
git add .

```


2. **Crear una nota explicando qué cambiaste:**
```bash
git commit -m "Descripción de lo que agregaste o corregiste"

```


3. **Enviar los cambios a GitHub:**
```bash
git push

```



---

**Tip para VS Code (Sin Comandos)**
También puedes usar el panel visual de VS Code si prefieres no usar la terminal:

* Haz clic en el ícono de **Control de código fuente** en la barra lateral izquierda (el ícono con nodos/ramas o presiona `Ctrl + Shift + G`).
* Escribe tu mensaje en el cuadro de texto que dice *Mensaje*.
* Haz clic en el botón azul **Confirmar** (Commit).
* Haz clic en el botón **Sincronizar cambios** (Sync Changes).