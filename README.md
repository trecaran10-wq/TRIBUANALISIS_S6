# Tribuanálisis · Tributación II · Semana 6 (V5)

Aplicación web Flask para diagnóstico tributario documentado de rentas de cuarta y quinta categoría y revisión de ingresos de creadores de contenido / economía de plataformas.

## Ejecutar en Windows (VS Code)

1. Instala Python 3.10 o posterior.
2. Descomprime el ZIP y abre la carpeta `tribuanalisis_semana6` en Visual Studio Code.
3. Abre Terminal y ejecuta:

   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   python app.py
   ```

4. Abre `http://127.0.0.1:5000`.

En macOS/Linux, activa el entorno con `source .venv/bin/activate`.

## Qué incluye esta versión

- Panel con tarjetas de categoría que abren el historial ya filtrado.
- Una única sección **Casos e historial** para consultar, filtrar, ver, editar, eliminar y generar informes.
- Formularios que anualizan automáticamente los ingresos mensuales y pagos a cuenta mensuales multiplicándolos por 12. Los importes anuales se mantienen sin multiplicar.
- Motor anual de rentas de trabajo para cuarta y quinta categoría: deducción de cuarta del 20% con límite de 24 UIT y exclusiones, deducción de 7 UIT, deducciones adicionales ingresadas con sustento sujetas a límite de 3 UIT, ITF, donaciones y escala progresiva acumulativa.
- Si se selecciona creador de contenido/economía de plataformas, no se impone una tasa plana de cuarta o quinta: se muestra una advertencia de clasificación y la necesidad de evaluar el tratamiento que corresponda. SUNAT ha señalado que ciertos ingresos de influencers por publicidad/monetización califican como tercera categoría (Informe N.º 044-2022-SUNAT/7T0000).
- Informe PDF técnico con resumen ejecutivo, visualización comparativa, trazabilidad de cálculo, base legal y cédula de apoyo para trasladar/verificar datos en la declaración anual.
- Guía didáctica con videos por módulo, enlaces directos a SUNAT e ilustraciones originales de la aplicación; se retiró la mini evaluación solicitada.
- Importación de Excel y plantilla de ejemplo.

## Límites profesionales importantes

El motor aplica una liquidación anual base para las rentas de trabajo con los datos ingresados. El resultado solo es tan completo como la información aportada: se debe revisar si existen otros ingresos de trabajo, rentas de fuente extranjera, ingresos exonerados/inafectos, pagos previos, créditos, deducciones con requisitos, suspensiones, excepciones, cambios normativos y reglas específicas del contribuyente. El monto de gratificaciones/bonificaciones gravadas debe incluirse en el total anual de quinta cuando corresponda.

El PDF es un **informe técnico / cédula de apoyo**, no un formulario oficial ni una declaración jurada presentada ante SUNAT. La aplicación no inicia sesión en SOL, no firma declaraciones ni las transmite. Para producir un archivo importable o una DJ oficial se necesita mapear exactamente los campos/formato de SUNAT del ejercicio correspondiente y validar sus reglas de presentación.

Los videos externos son propiedad de sus respectivos canales y requieren conexión a internet. Comprueba la fecha de publicación y contrasta siempre el contenido con las fuentes oficiales vigentes.


## Ajustes de la versión V4

- La pantalla **Nuevo análisis** muestra únicamente cuatro tarjetas en distribución 2×2; cada tarjeta abre un formulario específico.
- Los formularios de cuarta y quinta categoría permiten activar una renta de la otra categoría para analizar ambas en una sola ficha.
- Los importes mensuales de ingreso y pagos a cuenta se anualizan; los importes anuales se mantienen sin multiplicar.
- El resultado muestra primero los datos del caso y después una cédula de determinación con deducciones, impuesto por escala progresiva, pagos, retenciones y saldo.
- Los videos de la guía cambian según el módulo activo; los enlaces de búsqueda de YouTube son consultas temáticas cuando no se ha fijado un video concreto.
- Se quitó la pestaña independiente de Normativa del menú; los enlaces oficiales de SUNAT se mantienen dentro de la guía y del sustento del resultado.
- El perfil ahora presenta los créditos académicos de la Semana 6.
- La aplicación no presenta declaraciones juradas automáticamente: la cédula es un apoyo de trabajo y debe validarse con el formulario oficial del ejercicio.


## Ajustes de la versión V5

- Se eliminó de la **Guía tributaria** el bloque completo “Mini evaluación de la Semana 6”, incluidas sus preguntas, botón y lógica interactiva. El contenido de aprendizaje, videos y fuentes se conserva.
- En **Reportes PDF** se añadió la opción **Preparar 709** para casos calculados de cuarta o quinta categoría. Antes de generar el documento se solicita completar el RUC de 11 dígitos, nombres completos/razón social, periodo tributario, tipo de declaración, domicilio fiscal opcional, fecha de elaboración, pagos por categoría, retenciones por categoría, saldo anterior y renta neta extranjera si corresponde.
- El PDF presenta una cédula con el diseño tabular inspirado en la muestra proporcionada, con casillas referenciales para rentas de trabajo y/o fuente extranjera. Los importes se obtienen del caso seleccionado y de los datos adicionales ingresados.
- No se inventa número de orden SUNAT: se indica **NO ASIGNADO · GENERACIÓN LOCAL**. El PDF se marca como **modelo académico**, no es una constancia oficial, no se transmite a SUNAT y debe cotejarse con el formulario vigente del ejercicio antes de utilizar sus datos.
- Se verificaron la compilación de `app.py`, el análisis de plantillas Jinja y la generación de un PDF de prueba de una página. El entorno de desarrollo usado para esta revisión no tiene Flask instalado, por lo que no se afirma una validación completa de navegación en navegador.

## V6 · Cuentas personales y despliegue

- La primera pantalla es **Iniciar sesión** o **Crear cuenta**. Cada cuenta usa correo y contraseña; las contraseñas se almacenan como hash.
- Los análisis, importaciones Excel, reportes y eliminaciones están vinculados al usuario que inició sesión. Cada usuario solo puede consultar y modificar sus propios casos.
- El menú junto al nombre abre **Mi perfil** y **Cerrar sesión**. En Casos e historial se puede eliminar un caso individual o borrar el historial completo de la cuenta tras marcar una confirmación.
- El panel de identidad inferior usa degradados CSS borgoña–rosa–azul petróleo, sin una imagen decorativa de fondo.

### Variables de entorno recomendadas para Render

- `SECRET_KEY`: genera una clave larga y aleatoria; no uses la clave de desarrollo.
- `DATABASE_PATH`: con un disco persistente montado en `/var/data`, define `/var/data/tribuanalisis.db`.
- `SESSION_COOKIE_SECURE`: establece `true` cuando el sitio ya funciona con HTTPS.

**Persistencia:** el archivo SQLite debe guardarse en un disco persistente de Render para conservar cuentas e historial tras reinicios o nuevos despliegues. Si se utiliza almacenamiento efímero, los datos podrían perderse. Antes de publicar con usuarios reales conviene habilitar protección CSRF, límites de intentos de acceso y una política de copias de seguridad.
