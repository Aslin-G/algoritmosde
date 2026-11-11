# CodeSim Lab en GitHub Pages

Abre el [laboratorio publicado](https://aslin-g.github.io/algoritmosde/#agro) o descarga `index.html` y ábrelo con doble clic. El estudiante entra con nombre y correo y elige Material de estudio o Presentar evaluación. El laboratorio y la evaluación funcionan sin conexión. El navegador conserva los envíos pendientes y los reintenta cuando vuelve internet.

La misma [hoja CodeSim Lab — Evaluaciones](https://docs.google.com/spreadsheets/d/1QquvctEfxoOjl_39Ut2deKZOPY8kjlXDk_KelZEEVxA/edit) recibe los registros. **Actividad** muestra los accesos, casos abiertos, ejecuciones y logros de estudio, y el inicio, los intentos y la entrega de cada evaluación. **Evaluaciones** resume los intentos, la puntuación y las fechas de inicio y entrega. Las ejecuciones incluyen el programa en bloques. Los resultados se calculan en el navegador, con la identidad indicada por el estudiante.

## Si el receptor todavía no está activo

1. Desde esa hoja abre **Extensiones → Apps Script**. En `Código.gs` pega el contenido completo del `Code.gs` de esta carpeta y guarda el proyecto.
2. Abre **Implementar → Gestionar implementaciones**, selecciona la implementación que termina en `/exec` y pulsa el lápiz **Editar**.
3. En **Versión** elige **Nueva versión** y pulsa **Implementar**. Conserva el enlace actual, **Ejecutar como: Yo** y el acceso que ya permite recibir los envíos. Conserva las propiedades `SHEET_ID` y `SHARED_SECRET` del proyecto.
4. Vuelve al laboratorio e inicia sesión como estudiante. Al regresar la conexión, se envían los registros pendientes. Comprueba la entrada en **Actividad** y el resultado al entregar en **Evaluaciones**.

El archivo `index.html` incluye la dirección del receptor. La clave privada permanece en Apps Script y en la configuración del servidor. El alumnado no necesita acceder a la hoja.
