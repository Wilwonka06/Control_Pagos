# � Documentación del Sistema: Control de Pagos Importaciones v1.0.0

Este documento sirve como guía técnica y de usuario para el sistema de **Proyección de Pagos (Semanal y Mensual)**. Está diseñado para ser entendido tanto por usuarios finales como por personal técnico.

---

## ¿Qué es Python?

**Python** es el lenguaje de programación en el que está construido este sistema. Imagínalo como el "motor" que impulsa todas las operaciones. Es conocido por ser:

* **Legible:** Su código se parece mucho al inglés simple.
* **Versátil:** Se usa para todo, desde sitios web hasta inteligencia artificial y automatización de Excel (como en este caso).
* **Potente:** Permite procesar miles de filas de datos en segundos, algo que manual o visualmente en Excel tomaría mucho más tiempo.

---

## �️ Librerías Utilizadas (Las "Herramientas")

Para que Python pueda hacer su magia, utiliza conjuntos de herramientas especializadas llamadas "librerías". Aquí explicamos las más importantes usadas en este proyecto:

### 1. Pandas (`pandas`)

* **¿Qué es?**: Es la herramienta de análisis de datos más potente de Python.
* **¿Para qué la usamos?**: Es el "cerebro" analítico. Una vez que sacamos los datos de Excel, Pandas los convierte en una tabla virtual (DataFrame). Nos permite:
  * Filtrar pagos por fecha.
  * Eliminar filas vacías.
  * Detectar duplicados.
  * Ordenar y transformar columnas masivamente.

### 2. PyWin32 (`win32com.client`)

* **¿Qué es?**: Un puente que permite a Python "hablar" directamente con las aplicaciones de Windows.
* **¿Para qué la usamos?**: Es nuestro "titiritero" de Excel. Pandas es genial leyendo datos, pero malo manejando archivos con Macros (`.xlsm`) o actualizando vínculos externos. Usamos esta librería para:
  * Abrir Excel en segundo plano (invisible).
  * Dar la orden de "Actualizar todo" (RefreshAll) para que las fórmulas traigan los datos más recientes.
  * Guardar una copia limpia sin macros (.xlsx) para que Pandas la pueda leer.

### 3. Tkinter (`tkinter`)

* **¿Qué es?**: La librería estándar para crear interfaces gráficas (ventanas).
* **¿Para qué la usamos?**: Para que no tengas que tocar código. Crea:
  * La ventana de bienvenida.
  * Los selectores de fechas y calendarios.
  * La barra de progreso.
  * Los mensajes de "Éxito" o "Error".

### 4. TkCalendar (`tkcalendar`)

* **¿Qué es?**: Una extensión visual para Tkinter.
* **¿Para qué la usamos?**: Provee el calendario desplegable amigable donde seleccionas la fecha de corte.

### 5. ⚙️ ConfigParser (`configparser`)

* **¿Qué es?**: Manejador de archivos de configuración.
* **¿Para qué la usamos?**: Crea y lee el archivo `config_pagos.ini`. Gracias a esto, el programa "recuerda" dónde están tus archivos y carpetas, para que no tengas que seleccionarlos cada vez que lo abres.

---

## Estructura del Proyecto y Funciones Clave

El sistema se divide en 3 archivos principales de código (`.py`). Aquí explicamos qué hace cada uno y sus funciones más importantes.

### 1. `main.py` (El Cerebro Principal)

Es el punto de entrada. Cuando haces doble clic en el ícono, este es el archivo que se ejecuta primero.

* **`main()`**: La función directora.

  * Verifica si ya configuraste las rutas (dónde están tus archivos).
  * Muestra la ventana para elegir entre "Semanal" o "Mensual".
  * Llama al proceso correspondiente según tu elección.
  * Maneja errores globales (si algo explota, aquí se captura y se guarda en el reporte de errores).
* **`ConfiguradorRutas.cargar_o_crear_config()`**:

  * Busca el archivo `config_pagos.ini`. Si no existe, lanza la ventana de "Primera Configuración" pidiéndote ubicar los archivos clave.
* **`VentanaSeleccionTipo`**:

  * Dibuja la ventana azul con los dos botones grandes: "Proyección Semanal" y "Proyección Mensual".

### 2. `proceso_semanal.py` (Lógica Semanal)

Se encarga de generar la proyección para una semana específica (usualmente el próximo miércoles).

* **`ProcesadorSemanal.copiar_archivo_base()`**: **CRÍTICA**.

  1. Abre tu archivo original con Macros (`.xlsm`) usando Excel real.
  2. Ejecuta "Actualizar Vínculos" para traer datos frescos.
  3. Guarda una copia temporal como `.xlsx` (sin macros) para que sea seguro y fácil de leer.
  4. Elimina hojas innecesarias para aligerar el archivo.
* **`ProcesadorSemanal.filtrar_por_fecha()`**:

  * Toma todos los datos y se queda SOLO con los que tienen `FECHA DE PAGO` coincidente con la fecha que elegiste en el calendario.
* **`ProcesadorSemanal.anexar_archivo_final_com()`**:

  * Esta es la función "inteligente" de guardado.
  * Abre el archivo final acumulado.
  * **Verifica duplicados:** Compara lo que vas a guardar con lo que ya existe (usando Fecha, Proveedor e Importador como clave).
  * Si encuentra registros idénticos, los **reemplaza**. Si son nuevos, los **agrega** al final.

### 3. `proceso_mensual.py` (Lógica Mensual)

Similar al semanal, pero enfocado en reportes de mes completo.

* **`InterfazMensual`**:

  * Muestra selectores de "Mes" (Enero, Febrero...) y "Año", en lugar de un calendario de días específicos.
* **`ProcesadorMensual.filtrar_por_fecha()`**:

  * En lugar de buscar un día exacto, busca todos los pagos cuyo mes y año coincidan con tu selección.

---

## Guía de Instalación Rápida (Para Técnicos)

Si necesitas reinstalar el sistema en una computadora nueva:

1. **Instalar Python**: Descarga Python (versión 3.10 o superior) desde python.org. Al instalar, marca la casilla **"Add Python to PATH"**.
2. **Instalar Librerías**: Abre la terminal (CMD) en la carpeta del proyecto y ejecuta:
   ```bash
   pip install -r requirements.txt
   ```

   *(Esto instalará automáticamente pandas, pywin32, openpyxl, tkcalendar, etc.)*
3. **Ejecutar**:
   Doble clic en `main.py` (o usa el archivo `compilar.bat` si deseas crear un ejecutable .exe).

---

## 📞 Soporte

Para cambios en la lógica de negocio (ej. cambiar columnas, nuevas reglas de filtrado) o errores técnicos, contactar al desarrollador encargado.
