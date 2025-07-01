#Manual de Usuario

*Versión 1.0*

##Indice - Tabla de Contenidos
[TOC]


- - -

* * *

##Presentación
 El siguiente manual se escribe con la finalidad de dar a conocer la estructura y funcionamiento del proyecto "Implementación de un sistema de registro para el control de inventario"

- - -

##Requerimientos de Software

###Microsoft Excel
####Descripción
Microsoft Excel es un programa que permite editar hojas de cálculo desarrollada por Microsoft para Windows, macOS, Android y iOS.

Cuenta con cálculos, gráficas, tablas calculares y un lenguaje de programación macro llamado Visual Basic para aplicaciones. Ha sido una hoja de cálculo muy aplicada para estas plataformas, especialmente desde la versión 5 en 1993, y ha reemplazado a Lotus 1-2-3 como el estándar de la industria para las hojas de cálculo. Excel forma parte de la suite de software Microsoft Office.

>Dado el uso del lenguaje de programación Visual Basic para Aplicaciones (VBA) para más del 80% del funcionamiento del sistema, éste no podrá ser ejecutado en alternativas como: Google Sheets, WPS  Office Spreadsheets, LibreOffice Calc, entre otras.

#####Requerimientos
- [ ] Hardware
 - Windows 10 o posterior, or macOS 10.14 o posterior.
 - 4 GB RAM; 10 GB de espacio en el disco duro.
 - 1280 x 768 de resolución de pantalla.

- [ ] Plataformas

 - Microsoft Excel está disponible como una aplicación para escritorio (Windows y MacOS), en dispositivos móviles y tabletas (iOS y Android™) y online través del navegador.

>Para más detalles, clic [++aquí++](https://www.microsoft.com/es-mx/microsoft-365/excel)

###Microsoft VBA (Visual Basic for Applications)
####Descripción
Es un lenguaje de macros que se emplea para crear aplicaciones que permiten ampliar la funcionalidad de programas de la suite Microsoft Office. Se puede señalar que Visual Basic para Aplicaciones es un subconjunto casi completo de Visual Basic, Microsoft VBA al estar incluido dentro del Microsoft Office, puede emplearse tanto en Word, Excel, Access así como en Powerpoint.

##Funcionalidades
- ==CRUD== de productos
- Transición de productos: `No Procesados - Procesados - Salidas`
- Asignación de ubicaciones automáticas
- Análisis ABC para determinar categorías
- Filtros por Número de Orden o Fecha
- Búsqueda Global mediante Batch

- - -

* * *

##Estructura General
 El programa se opera en su totalidad desde formularios, también se usan hojas de excel a manera de base de datos, con ellas se almacena y organiza la información y con los formularios se ve y gestiona. Adicionalmente se utilizan "Hojas Avanzadas", hojas que a diferencia de las demás no sólo contienen información usada en los formularios sino que operan sin ellos, ofrenciendo funcionalidades por sí mismas.

- - -

###Hojas de Excel
**Contenido**

|Hoja|Contenido|Datos|
|-|-|-|
|inicio|Hoja vacía, sin datos, con un sólo botón que inicia el programa.|`Botón`|
|no_procesado|Productos que recién ingresan al almacén, deben procesarse antes de poder ser despachados.|`id` `Producto` `No. Orden` `Batch` `Cantidad` `Ubicación` `Fecha de entrada`|
|procesado|Productos ya procesados, listos para despacharse.|`id` `Producto` `No. Orden` `Batch` `Cantidad` `Ubicación` `Fecha de entrada` `Fecha de ubicación` `Ubicación anterior`|
|salidas|Productos ya despachados (historial).|`id` `Producto` `No. Orden` `Batch` `Cantidad` `Fecha de entrada` `Fecha de ubicación` `Fecha de Salida` `Ubicación de entrada` `Ubicación anterior`|
|reports|Muestra tablas y gráficos con información dinámica.|Esta es una hoja avanzada|
|abc|Datos necesarios para el análisis ABC, con él se definen las categorías que se usan para discriminar ubicaciones y productos cuando se agregan a "No Procesado" o se transfieren a "Procesado". |`Productos - Categorías` `ABC Parte 1` `ABC Parte 2`|
|layout|Muestra un plano del almacén, se centra en los pasillos 7, 8 y 9. Contiene botones que permiten señalar las ubicaciones ocupadas y resaltadas con un color diferente de acuerdo a su categoría. |Esta es una hoja avanzada|
|data_inout|Lista las ubicaciones disponibles por pasillo y define que rangos corresponden a cada categoría. |`id` `Ubicación` `Categoría`|
|data_seven|Transformación de datos que modifica y ordena la lista de ubicaciones disponibles del pasillo 07. |`Lista Original` `Alfabeto` `Columna` `Orden`|
|data_en|Mismo caso que la hoja anterior, pero para los pasillos 08 y 09.|`Lista Original` `Alfabeto` `Columna` `Orden`|

- - -

###Formularios de Excel

**Contenido General**

Formularios que permiten la navegación total del sistema, permiten, entre otra cosas, las operaciones ==CRUD== para gestionar los productos desde que entran (a "No procesado") se procesan ("Procesado") hasta que se despachan ("Salidas"). Permiten la búsqueda y filtrado de los productos según ciertos requerimientos y la generación de repores tabulares y gráficas.

- - -
####Estructura General

Aquí se proporciona una relación de los formularios y las acciones y/o transiciones que tiene disponibles.

| Estado / Sección | Formulario | Acciones disponibles / Transiciones|
| - | - | - |
| **Inicio** | Menú Inicio | **Navega a:** <br> - `No Procesado - Gestionar Producto` <br> - `No Procesado - Vista Avanzada` <br> - `Procesado - Vista Avanzada` <br> - `Salidas - Vista Avanzada` <br> - `Búsqueda Global` |
| | `Búsqueda Global` | Muestra la información encontrada. |
| **No Procesado** | `- Vista Avanzada` |  Permite filtrar registros mediante Fecha, No. Orden o Batch, permite combinar los parámetros de filtrado. <br> Doble clic sobre un registro → `Vista Edición` |
| | ` - Vista Edición` | - *Editar Registro* → `No Procesado - Gestionar Producto` <br> - *Eliminar Registro* → Mensaje de confirmación antes de realizar la operación<br> - *Mover a procesados* → `No Procesado - Transición a Procesado` |
| | `- Gestionar Producto` | ***Agrega*** ó ***edita*** un producto según desde dónde se acceda al formulario → Regresa a ` - Vista Edición`|
| | `- Transición a Procesado` | Mueve un registro de producto a Procesado|
| **Procesado** | `- Vista Avanzada` | Permite filtrar registros mediante Fecha, No. Orden o Batch, permite combinar los parámetros de filtrado. <br> Doble clic sobre un registro → `Vista Edición` |
| | ` - Vista Edición` | - *Editar Registro* → `Procesado - Modificar Producto` <br> - *Eliminar Registro* → Mensaje de confirmación antes de realizar la operación <br> - *Salida de producto* mueve un registro de producto a Salidas → Mensaje de confirmación antes de realizar la operación|
| | ` - Modificar Producto` | Permite modificar la información de un producto procesado → Regresa a ` - Vista Edición`|
| **Salidas**      | `- Vista Avanzada` | Permite filtrar registros mediante Fecha, No. Orden o Batch, permite combinar los parámetros de filtrado. |

- - -

* * *

####Estructura Detallada

- - -

#####Menú de Inicio

Esta es la pantalla de inicio del sistema, diseñada para ofrecer acceso rápido a los módulos principales del formulario. Al abrir la aplicación, el usuario es recibido con un menú lateral y con el logotipo del TECNM Campus Zacapoaxtla y de la empresa CEVA Logistics.

**Objetivo de esta pantalla:**
Brindar al usuario accesos directos hacia las funciones principales.

![Menu Inicio](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/main.png?raw=true)

| Botón | Función |
|--------|--------|
|Agregar producto|Abre el formulario de ingreso de productos.** |
|Producto No Procesado|Acceso al formulario donde se visualizan y gestionan los productos que han ingresado al pasillo 7, listos para procesarse.|
|Producto Procesado|Muestra los productos que han sido ubicados en los pasillos 8 y 9.|
|Registro de salidas|Sección para consultar las salidas de productos del almacén.|
|Reportes|Cierra el formulario y permite la navegación a la hoja *reports*.|
|Excel|Cierra el formulario y muestra las hojas de datos directamente en Microsoft Excel.|

> ** Tras mostrar el formulario para agregar productos se mostrará la vista de edición de productos no procesados, independiemente de si se completa o no el ingreso del producto.

**Búsqueda de productos**
En la parte central se encuentra un cuadro de texto y un botón que permiten realizar una búsqueda global ingresando el Batch del producto, como resultado, se muestra una ventaja emergente que muestra los detalles del producto

- - -

#####Gestionar Producto
**Objetivo de esta pantalla:**
Este formulario permite tanto Agregar como Modificar un registro en No Procesado según desde dón de acceda a él. Si se accede desde el menú inicio, estará en modo "Agregar Producto", si se accede desde la "No Procesado - Vista Edición" estará en modo "Modificar Producto".

| Modo Agregar | Modo Editar |
|--------|--------|
|![Gestionar Producto - Modo Agregar](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/np_add.png?raw=true)|![Gestionar Producto - Modo Editar](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/np_edit.png?raw=true)|

> Consideraciones:
> - **Tipo** se refiere al tipo de control usado en el formulario.
> - **Origen** apunta a la hoja de donde se obtiene la información, sólo cuando sea el caso

| Elemento | Función / Restricción | Tipo | Diferencia de Modo |
|-|-|-|-|
|Producto| Permite seleccionar un producto de entre los disponibles de acuerdo a la hoja `abc` | Lista desplegable `Dropdown` |Sólo está habilitado cuando se agrega producto |
|No. Orden|No es mandatorio que sea único, si se repite agrupa varios productos en los filtros de las vistas avanzadas|Campo de Texto `TextBox` |Editable en ambos modos|
|Batch|Debe ser único|Campo de Texto `TextBox`|Editable en ambos modos|
|Cantidad|Lista desplegable que cambia de datos disponibles en función del pruducto seleccionado, permite respetar los límites físicos del almacén al tomar los valores de cantidad por tarina |Lista desplegable `Dropdown` |Sólo está habilitado cuando se agrega producto|
|Fecha de entrada|Formato **dd/mm/yyyy**|Campo de Texto `TextBox`|Editable en ambos modos|
|Ubicación|Elegida automáticamente para respetar el Análisis ABC. <br> Al seleccionar un producto elige una ubicación libre y al seleccionar la cantidad puede cambiar la ubicación en función de si la cantidad permite almacenar más producto si aún no se ha alcanzado el limite máximo por ubicación, prioriza el almacenamiento del mismo producto cuando la cantidad lo permite y pioriza una ubicación libre cuando se selecciona el máximo por tarima. <br> La ubicación no puede definirse manualmente, el control es únicamente informativo.|Etiqueta `Label`|-|
|Categoría|Muestra a qué categoría pertenece la ubicación calculada. Dado que las categorías pueden cambiar de acuerdo al Análisis ABC este dato no se asocia directamente a un producto, por lo que no forma parte de los registros. <br> La categoría no puede definirse manualmente, el control es únicamente informativo. |Etiqueta `Label`|-|
|Aceptar|Verifica que todos los cambos tengan información válida, muestra mensajes en caso de no ser así. <br> El control puede estar activado o desactivado en función de si se ha o no seleccionado un producto, para evitar registros erróneos. |Botón `Label`|-|
|Cancelar|Cancela la operación cerrando el formulario.|Botón `Label`|-|

Límite por tarima, límite por ubicación y el código de ubicación se muestran en ambos modos, pero no son editables ni afectan esta operación.

> **Importante:**
> Si estás editando un registro y necesitas cambiar la cantidad o el producto, debes eliminar el registro y volver a crear uno nuevo. Esto evita conflictos con la lógica de almacenamiento.

- - -

#####No Procesado – Vista Avanzada
**Objetivo de esta pantalla:**
Facilitar la visualización, búsqueda y filtrado de productos no procesados mediante criterios específicos.

![No Procesado – Vista Avanzada](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/np_va.png?raw=true)

> Este formulario es visualmente idéntico a Procesado – Vista Avanzada (incluso en las funciones), así que, para evitar duplicar información, esta sección sirve de referencia para ambos formularios.

######Panel de Filtros
Situada a la izquierda.
- **Fecha (mm/dd/yyyy):**
Hace uso del formato americano (mes/día/año).
Úsala cuando conozcas el día exacto en que ingresó un producto.

- **No. Orden:**
Ideal cuando deseas encontrar todos los productos relacionados con una orden.

- **Batch:**
Útil cuando se necesita encontrar un producto específico.

 -**Botón "Limpiar":**
Disponible en cada filtro. Borra el contenido del campo correspondiente para realizar una nueva búsqueda.

- **Botón "Buscar":**
Ejecuta la búsqueda con base en los filtros aplicados.

- **➡️ Botón "Siguiente":**
Avanza a la Vista Edición.

> Es posible combinar los filtros para realizar búsquedas más precisas.
> Si se limpínan todos los campos la tabla de resultados muestra todos los productos registrados en "No Procesado".

###### Tabla de resultados
Situada a la derecha.
Una vez realizada la búsqueda, se muestran los resultados en la tabla con las siguientes columnas:

- ID.

- Producto.

- No. Orden.

- Batch.

- Cantidad.

- Ubicación.

- Fecha Entrada.

> Dar doble clic a un registro en la tabla de resultados abrirá la "Vista Edición" con el registro señalado listo para operar con él.

- - -

#####No Procesado – Vista Edición
Esta pantalla permite gestionar todos los productos que se encuentran pendientes de procesarse. A diferencia de la vista avanzada que está orientada a búsquedas, esta vista está centrada en la gestión de los registros, permite editar, eliminar o mover productos a procesado.

**Objetivo de esta pantalla:**
Facilitar la administración total de productos no procesados, permitiendo ejecutar acciones sobre los registros listados.

![No Procesado – Vista Edicion](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/np_ve.png?raw=true)

######Funcionalidades disponibles
**Editar Registro:**
Permite modificar la información de un producto existente.

- Es necesario seleccionar un registro previamente.
- Al hacer clic, se abre un formulario de "Gestión de Producto" en modo de edición.

**Eliminar Registro:**
Elimina el registro de producto del sistema.

- Requiere selección previa de un registro.
- Aparece un mensaje de confirmación antes de proceder, para evitar eliminaciones accidentales.

**Mover a Procesados:**
Esta opción inicia el proceso para convertir un producto en "Procesado".

- Se muestra un nuevo formulario al hacer clic (se explica más adelante).

**Vista Avanzada:**
Regresa a la vista especializada en filtrado de registros.

> Esta vista también tiene una tabla de resultados, en la que se pueden observar todas las ubicaciones de "No Procesado", tanto si están ocupadas, como si no.
> Dar doble clic en algún registro no desencadena ninguna acción.

- - -

#####Transición a Procesado
Este formulario aparece cuando el usuario selecciona un producto desde la vista de edición y hace clic en “Mover a Procesados”. Su función es mostrar un resumen informativo del producto que está por cambiar de estado: de "No Procesado" a "Procesado".

**Objetivo de esta pantalla:**
Presentar de forma clara los datos del producto antes de realizar el traspaso al pasillo procesado, junto con la nueva ubicación asignada por el sistema.

![Transicion a Procesado](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/np_trans.png)

######Información mostrada

- Producto
- No. Orden
- Batch
- Cantidad
- Fecha de entrada

> Todos estos campos son etiquetas. Se muestran como referencia para que el usuario confirme que se trata del producto correcto.

######Ubicaciones
**Ubicación Anterior:**
Corresponde a la ubicación que tenía el producto estando en "No Procesado".

**Nueva Ubicación:**
Calculada automáticamente por el sistema con base en la categoría del producto (A, B o C), buscando espacios disponibles en las ubicaciones asignadas a dicha categoría.

> Si no hay espacio disponible, se muestra un mensaje de advertencia.

######Botones
**Aceptar:**
Confirma la transición del producto. Este proceso mueve los datos del producto de la hoja `no-procesado` a `procesado`.

**Cancelar:**
Cierra la ventana sin realizar ninguna acción.

> Recomendación:
Antes de hacer clic en "Aceptar", revisa cuidadosamente la nueva ubicación para asegurarte de que se asignó correctamente.

- - -

#####Producto Procesado – Vista Avanzada
Visualmente idéntico a la Vista Avanzada de "No Procesado" y, de hecho, realiza las mismas funciones. Revise la sección correspondiente para conocer los detalles.

![Producto Procesado – Vista Avanzada](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/p_va.png?raw=true)

- - -

#####Producto Procesado – Vista Edición
Esta pantalla permite gestionar productos que han sido ubicados en el pasillo de "Procesado". Desde aquí, el usuario puede editar la información del producto, eliminar registros o dar de alta una salida.

**Objetivo de esta pantalla:**
Brindar control sobre los productos ya procesados, permitiendo actualizaciones o mover registros de productos a la sección de "Salidas".

![Producto Procesado – Vista Edición](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/p_ve.png?raw=true)

######Funcionalidades disponibles
**Editar Registro:**
Abre un formulario específico para modificar algunos datos del producto procesado.

- Este formulario es distinto al de productos no procesados (lo explicaremos más adelante).

**Eliminar Registro:**
Permite borrar el registro de un producto procesado.

- Requiere confirmación previa para evitar eliminaciones accidentales.

**Salida de Producto:**

- Esta acción registra el movimiento en un historial permanente, sin restricciones.
- También muestra un mensaje de confirmación antes de proceder.

**Vista Avanzada:**
Regresa a la vista especializada en filtrado de registros.

- Visualmente es igual a la vista avanzada de "No Procesado" y cumple las mismas funciones.

> Esta vista también tiene una tabla de resultados, en la que se pueden observar todas las ubicaciones de "Procesado", tanto si están ocupadas, como si no.
> Dar doble clic en algún registro no desencadena ninguna acción.

- - -

#####Modificar Producto
Este formulario permite modificar ciertos datos de productos que ya han sido ubicados en "Procesado". Es similar al formulario de edición de "No Procesado", pero incluye campos adicionales relacionados con la ubicación y fechas importantes.

![Modificar Producto](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/p_edit.png?raw=true)

| Campo | Editable | Descripción |
| - | - | - |
| **Producto** | <img src="https://github.com/FortAwesome/Font-Awesome/raw/6.x/svgs/regular/circle-xmark.svg" width="20" height="20"> |-|
| **No. Orden**              | <img src="https://github.com/FortAwesome/Font-Awesome/raw/6.x/svgs/regular/circle-check.svg" width="20" height="20">| Puede ajustarse si hubo un error de registro o se reasignó.|
| **Batch** | <img src="https://github.com/FortAwesome/Font-Awesome/raw/6.x/svgs/regular/circle-check.svg" width="20" height="20"> | Puede cambiarse si no entra en conflicto con otro Batch ya existente.|
| **Cantidad** | <img src="https://github.com/FortAwesome/Font-Awesome/raw/6.x/svgs/regular/circle-xmark.svg" width="20" height="20"> | Ya fue considerada en la lógica de ubicación, por eso no puede modificarse.|
| **Fecha de entrada**|<img src="https://github.com/FortAwesome/Font-Awesome/raw/6.x/svgs/regular/circle-xmark.svg" width="20" height="20"> | Corresponde al momento en que el producto fue registrado como "No Procesado". |
| **Fecha de ubicación**| <img src="https://github.com/FortAwesome/Font-Awesome/raw/6.x/svgs/regular/circle-check.svg" width="20" height="20"> | Puede editarse para reflejar el día exacto en que el producto se ubicó.|
| **Ubicación anterior**|<img src="https://github.com/FortAwesome/Font-Awesome/raw/6.x/svgs/regular/circle-xmark.svg" width="20" height="20">| Muestra la ubicación cuando era “No Procesado”. Solo informativa.|
| **Ubicación actual**|<img src="https://github.com/FortAwesome/Font-Awesome/raw/6.x/svgs/regular/circle-xmark.svg" width="20" height="20">| Muestra la ubicación asignada en estado “Procesado”. Solo informativa.|
| **Categoría del producto**|<img src="https://github.com/FortAwesome/Font-Awesome/raw/6.x/svgs/regular/circle-xmark.svg" width="20" height="20">| Se indica si el producto pertenece a la categoría A, B o C.|

######Validaciones

- El único campo con validación estricta es Batch, que debe seguir siendo único en todo el sistema.
- La Fecha de ubicación puede modificarse libremente, pero debe mantener un formato válido.

######Botones
**Aceptar:**
Guarda los cambios, siempre que no haya conflictos en los datos.

**Cancelar:**
Cierra la ventana sin realizar cambios.

> Este formulario es útil cuando se detecta un error en el batch, número de orden o fecha de ubicación del producto. Si lo que se necesita es mover el producto, debe utilizarse el botón “Salida de producto” desde la vista de edición.

- - -

#####Salidas – Vista Avanzada
Esta pantalla permite consultar los productos que han salido del almacén, es decir, aquellos que pasaron del estado Procesado al estado Salidas mediante el botón "Salida de producto".

**Objetivo de esta pantalla:**
Brindar un historial detallado de todos los productos que han sido despachados, manteniendo así un registro inalterable.

![Salidas – Vista Avanzada](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/s_va.png?raw=true)

> Solo lectura:
Esta vista no permite editar ni eliminar registros. Los productos aquí son definitivos.

######Filtros disponibles
| Filtro        | Descripción                                            |
| ------------- | ------------------------------------------------------ |
| **Fecha**     | Filtra por la fecha en que el producto fue despachado. |
| **No. Orden** | Busca por el número de orden del producto. |
| **Batch**     | Permite ubicar productos únicos. |

> Cada filtro tiene un botón "Limpiar" para restablecer el valor y realizar una nueva búsqueda.

Al igual que las vistas avanzadas anteriormente detalladas, ésta también muestra una tabla de resultados, los campos que muestra son los siguientes:

- ID.
- Producto.
- No. Orden.
- Batch.
- Cantidad.
- Ubicación.
- Fecha de entrada.
- Fecha de ubicación.
- Ubicación anterior.
- Fecha de salida.

- - -

#####Búsqueda Global
La búsqueda global permite al usuario consultar un producto por su Batch sin necesidad de saber en qué parte del proceso se encuentra (No Procesado, Procesado o Salidas).

Este formulario presenta un resumen completo de la información del producto consultado, sin permitir modificación alguna. Su propósito es informativo.

![Busqueda Global](https://github.com/jv01-27/manual-cinveex/blob/main/screenshots/gl_sh.png?raw=true)

**Objetivo de esta pantalla:**

- Localizar rápidamente un producto con base en su Batch.
- Obtener un resumen claro del estado, ubicaciones y fechas clave del producto en su paso por el sistema.

######Campos:
| Campo| Descripción |
| - | - |
| **ID**                 | Número de identificación del registro.|
| **Producto**           | Nombre del producto.|
| **No. Orden**          | Número de orden asignado al producto.|
| **Batch**              | Código único.|
| **Cantidad**           | Cantidad registrada del producto.|
| **Fecha de entrada**   | Fecha en que ingresó a "No Procesado" *(presente en cualquier estado)*.|
| **Fecha de ubicación** | Fecha en la que fue movido a Procesado *(visible si está en Procesado o Salidas)*.|
| **Fecha de salida**    | Fecha en la que fue despachado *(visible solo si está en Salidas)*.|

######Información de ubicaciones:
Según el estado del producto, se mostrarán diferentes etiquetas de ubicación:

| Estado del producto | Ubicación mostrada                                                                                      |
| ------------------- | ------------------------------------------------------------------------------------------------------- |
| **No Procesado**    | `Ubicación` → Ubicación actual en "No Procesado".                                                         |
| **Procesado**       | `Ubicación` → Actual en "Procesado". <br> `Ubicación anterior` → Ubicación previa en "No Procesado".        |
| **Salidas**         | `Ubicación de Entrada` → Ubicación en "No Procesado". <br> `Ubicación Anterior` → Ubicación en "Procesado". |

######Hoja
Este campo indica en cuál de las tres hojas se encontró el producto:

- no-procesado
- procesado
- salidas

> Solo puede tener uno de estos tres valores.

######Consideraciones

- Si no se encuentra el batch, se muestra un mensaje de alerta.
- El formulario es completamente informativo, no editable.
- Garantiza la integridad de los datos presentados ya que extrae la información directamente de las hojas.


- - -

* * *

###Hojas Avanzadas

- - -

#####Reportes

######Funcionalidad
Al seleccionar esta opción:

- Se cierran los formularios activos del sistema.

- Se abre directamente el archivo Excel que contiene los datos maestros (no-procesado, procesado, salidas), la hoja muestra diversas tablas dinámicas, fórmulas de análisis, y gráficas automáticas.

######Contenido
1. Catálogo de Productos y Categorías

 - Lista los productos definidos en el sistema.

 - Categoría asignada según análisis ABC:

    - A = Alta rotación

    - B = Rotación media

    - C = Baja rotación

2. Ubicaciones
 - Totales por estado:

    - No procesado: ocupadas y disponibles.

    - Procesado: ocupadas y disponibles.

3. Distribución por Categoría y Estado
 - Ubicaciones ocupadas y disponibles en cada categoría

 - Para No Procesado y Procesado.

4. Conteo de Productos por Categoría
 - Muestra el número actual de productos en cada categoría para ambos estados.

5. Entradas y Salidas Registradas
 - Totales históricos de entradas (productos que han ingresado) y salidas (productos despachados).

6. Gráficas y Análisis ABC (Pareto)
 - Se genera automáticamente una gráfica de Pareto.

 - Este análisis clasifica los productos por su impacto en inventario (20/80) y determina su categoría.

######Botones
**Actualizar Datos:**
Vuelve a cargar los datos y refresca las tablas dinámicas y fórmulas si no lo hacen automáticamente.

**Volver al programa:**
Cierra la vista Excel y vuelve a abrir el formulario Menú Inicio del sistema.

> Esta sección no tiene controles VBA activos más allá de los botones.

- - -

#####Layout

######Funcionalidad
Muestra todas las ubicaciones repartidas en 3 pasillos (7, 8 y 9).

######Botones
**Mostrar/Actualizar Ubicaciones:**
Señala qué ubicaciones están ocupadas con un color diferente de acuerdo a su categoría.

**Restaurar Layout:**
Restablece el formato de las celdas que contienen ubicaciones.

- - -

* * *


##Glosario

###C
####CRUD
CRUD es un acrónimo que representa las cuatro operaciones fundamentales que se realizan sobre los datos: Crear (**C**reate), Leer (**R**ead), Actualizar (**U**pdate) y Eliminar (**D**elete). Estas operaciones permiten interactuar con datos persistentes almacenados en bases de datos u otros sistemas de almacenamiento de datos.
- - -