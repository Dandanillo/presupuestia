# Conectar PresupuestIA con Google Drive

Con esto, los presupuestos, los PDF, el preciario y los datos de la empresa se guardan en una hoja de Google que veis los dos. Es gratis y se hace **una sola vez** (unos 10 minutos).

## Qué se crea

```
Presupuestia/                  ← carpeta en tu Drive (compartida con Lou)
├── Presupuestia – Datos       ← hoja de Google con estas pestañas:
│     Presupuestos             una fila por presupuesto (obra, cliente, estado, base, IVA, total, enlace al PDF…)
│     Partidas                 todas las partidas de todos los presupuestos (con el importe calculado)
│     Preciario                vuestros precios (con columna Zona)
│     Empresa                  los datos que salen en el PDF
│     Conexión                 la clave y la dirección de conexión
└── PDFs/                      ← los PDF de los presupuestos
```

Puedes abrir la hoja en Google Sheets, o descargarla como Excel (Archivo → Descargar → Microsoft Excel). En la pestaña **Presupuestos**, la columna **Estado** se puede cambiar a mano (Borrador, Enviado, Aceptado, Rechazado).

## Pasos

### 1. La carpeta y la hoja
1. En Google Drive: **Nuevo → Carpeta** → llámala `Presupuestia`.
2. Compártela con Lou: botón derecho sobre la carpeta → **Compartir** → su correo → **Editor**.
3. Entra en la carpeta y pulsa **Nuevo → Hojas de cálculo de Google**. Llámala `Presupuestia – Datos`.

### 2. El script
1. Dentro de esa hoja: menú **Extensiones → Apps Script**. Se abre otra pestaña.
2. Borra todo el código que aparece (la función `myFunction`).
3. Abre el archivo `drive/Code.gs` de este repositorio en GitHub, pulsa el botón **Copy raw file** (los dos cuadraditos) y pégalo en Apps Script.
4. Pulsa el icono del disquete (**Guardar**).

### 3. Prepararlo (una vez)
1. Arriba, en el desplegable de funciones, elige **`configurar`** y pulsa **Ejecutar**.
2. Google pedirá permisos:
   - **Revisar permisos** → elige tu cuenta.
   - Saldrá *"Google no ha verificado esta aplicación"*. Es normal: el script es tuyo. Pulsa **Configuración avanzada** → **Ir a … (no seguro)** → **Permitir**.
3. Cuando termine, vuelve a la hoja. Habrá pestañas nuevas; en **Conexión** está tu **clave**.

### 4. Publicarlo
1. En Apps Script: **Implementar → Nueva implementación**.
2. Pulsa el engranaje ⚙ y elige **Aplicación web**.
3. Rellena:
   - **Ejecutar como:** Yo
   - **Quién tiene acceso:** Cualquier persona
4. Pulsa **Implementar** y copia la **URL de la aplicación web** (acaba en `/exec`).
5. Pega esa URL en la hoja, pestaña **Conexión**, en la celda **B2** (así Lou la tiene a mano).

> "Cualquier persona" **no** significa que los datos sean públicos: sin la clave, la conexión no responde nada.

### 5. Conectar la web
1. En PresupuestIA: **Ajustes → Drive compartido**.
2. Pega la **URL** y la **clave** (de la pestaña Conexión) → **Guardar** → **Probar conexión**.
3. Debe salir "✅ Conexión con Drive correcta".

### 6. Lou
En tu web: **Ajustes → Drive compartido → 🔗 Copiar enlace de acceso para Lou**. Mándaselo **solo a ella**, por un chat privado. Cuando lo abra en su móvil u ordenador, se conecta a Drive sin pegar nada, y solo tiene que crear su usuario y contraseña en ese dispositivo.

Después, todo esto le llega solo desde Drive: presupuestos, preciario y datos de la empresa (menos el logo, que es de cada dispositivo). Si quieres que tampoco tenga que pegar la clave de Gemini, en **Ajustes → IA** pulsa **☁️ Compartir esta clave por Drive** (necesita el script actualizado, ver abajo).

Quien tenga ese enlace entra en vuestros datos. Si se filtra, ejecuta `nuevaClave` (ver Seguridad).

## El preciario compartido
- En cada partida de un presupuesto hay un botón **📚 Guardar en preciario**, y abajo uno para guardar **todas** las partidas ya revisadas. Revisa los precios antes: si la partida ya existe, se actualiza.
- Los precios se guardan en la pestaña **Preciario** de la hoja, y la IA los usa en los presupuestos nuevos de los dos.
- Para **borrar** un precio, borra su fila en la pestaña Preciario de la hoja de Drive.
- La columna **Zona** es opcional: sirve para tener precios distintos por zona.
- No hace falta tocar el script de Drive: esta mejora solo cambia la web.

## Cuentas (entrar desde cualquier sitio)
Con la versión 2.0 del script, las cuentas viven en Drive y se entra con **usuario y contraseña** desde cualquier móvil u ordenador.
- **Primera vez:** la web te pide la clave de Drive (pestaña Conexión, celda B1) para crear la cuenta de administrador. Solo se puede hacer mientras no haya ninguna cuenta.
- **Más usuarios:** el administrador los crea en la pestaña **Administración** de la web, con nombre, correo, teléfono y dirección (opcionales). Nadie puede registrarse desde fuera.
- **Pestaña Usuarios de la hoja:** quién tiene cuenta y sus datos de contacto. **Las contraseñas no están en la hoja**: la web las convierte en una huella lenta antes de enviarlas y el script solo guarda otra huella, en sus propiedades privadas.
- **Pestaña Accesos de la hoja:** fecha y hora, usuario, resultado (Correcto, Fallo, Bloqueado, Alta, Baja, Cambio) y el dispositivo que declara el navegador. **La IP no se puede registrar**: Google no la da a los scripts.
- **Bloqueo:** 5 contraseñas malas seguidas bloquean ese usuario 15 minutos.
- **Contraseñas:** mínimo 10 caracteres.
- Al **salir**, el dispositivo olvida la clave de Drive y la clave de Gemini compartida, así que es seguro usar un PC ajeno.

## Seguridad
- **Candado:** en un dispositivo nuevo, la web no deja ni crear cuenta hasta que se pegan la URL y la clave de Drive, y Drive las acepta. Quien no las tenga solo ve la pantalla de acceso privado. Es un candado en la propia web: protege vuestros datos (que están en Drive), pero no sirve contra alguien que modifique su copia de la página en su navegador, por eso la clave de Drive es lo que de verdad importa. No la compartáis.
- Tu clave de Gemini se guarda solo en cada dispositivo. No la pases a nadie: sería la única forma de gastar tu cuota gratuita.
- Quien tenga la **URL y la clave** puede leer y escribir en la hoja. No las compartáis fuera de vosotros.
- Si la clave se filtra: en Apps Script ejecuta **`nuevaClave`**. La anterior deja de valer; copia la nueva de la pestaña Conexión y pégala en la web.
- La clave se guarda solo en el navegador de cada uno, nunca en GitHub.

## Actualizar el script
Cuando salga una versión nueva de `drive/Code.gs` (esta añade compartir la clave de Gemini):
1. En Apps Script, borra todo el código y pega el nuevo (**Copy raw file** en GitHub). Guarda con el disquete.
2. **Implementar → Gestionar implementaciones** → lápiz ✏ → **Versión: Nueva versión** → **Implementar**.

La clave de acceso y los datos no se tocan, y la URL no cambia.

## Detalle del cambio de versión
**Implementar → Gestionar implementaciones** → lápiz ✏ → **Versión: Nueva versión** → **Implementar**. La URL no cambia.
