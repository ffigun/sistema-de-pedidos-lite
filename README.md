# Sistema de Pedidos Lite (v2.5 Release 42)
Este es un Sistema de tickets sumamente simple para realizar el seguimiento de trabajos. No se especifican requisitos mínimos por la ínfima cantidad de recursos que necesita este software para ejecutarse. Funciona con las versiones de 32 y 64 bits de Windows XP, Vista, 7, 8, 8.1, 10 y 11 así como sus variantes Windows Server.

![Sistema de Pedidos Lite](/Res/SistemaDePedidos.png)

## Modo de uso
Descargar el contenido de la carpeta /Bin/ y ejecutar el programa SistemaDePedidosLite.exe.

## Funcionamiento y primeros pasos
Para el Sistema de Pedidos Lite, un *Pedido* de trabajo es una tarea solicitada por un *Contacto* a realizar por un *Técnico* en una *Sucursal*. Cada *Informe* de trabajo es toda acción realizada sobre dicho pedido. Un pedido de trabajo puede estar en estado PEND si está pendiente u OK si ya fue finalizado.

> Tip: Las fechas del Sistema de Pedidos Lite siempre se especifican en formato ddmmaa, y los tiempso en formato hhmm.

### Carga de pedidos e informes
Para cargar un nuevo pedido se debe hacer clic en Archivo > Nuevo > Pedido, o presionar Ctrl + N. Para cargar un nuevo informe sobre un pedido existente, se debe seleccionar el pedido en la lista de pedidos, y luego hacer clic en Archivo > Nuevo > Informe o presionar Ctrl + M.

### Visualización de pedidos e informes
Para visualizar la información de un pedido basta con seleccionarlo. La información aparecerá en el panel de la derecha. También se puede hacer clic derecho y seleccionar la opción Ver pedido. Para ver los informes de trabajo se puede hacer doble clic sobre un pedido, o hacer clic derecho sobre un pedido y seleccionar la opción Ver informes.

### Eliminación de pedidos e informes
El Sistema de Pedidos no permite la eliminación de pedidos e informes desde la interfaz y no se recomienda modificar manualmente los archivos Informes.ini, Pedidos.ini ni Enlace.ini. Para quitar la información de un pedido o informe se puede realizar la edición utilizando las opciones del menú Archivo o presionando Ctrl + E.

### Cierre de pedidos
Una vez finalizado un pedido, se debe cargar un informe de trabajo y tildar la opción "Marcar como OK". Se habilitará un nuevo campo llamado F. OK donde se debe especificar la fecha de finalización, que por defecto es el día actual.

### Exportación de pedidos e informes
Para exportar un pedido con todos sus informes de trabajo en formato TXT se puede utilizar la opción Archivo > Exportar > Pedido seleccionado o bien presionar Ctrl + A. Para exportar todos los pedidos o informes en formato CSV, puede utilizar las opciones correspondientes del menú Archivo > Exportar.

### Búsqueda de pedidos e informes
Para buscar información basta con dirigirse al menú Opciones > Buscar o bien presionar Ctrl + F. Una vez en este menú, puede filtrar por pedidos o informes con el selector correspondiente. Dependiendo del filtro que elija, el marco Buscar le ofrecerá distintos campos para filtrar la información. Una vez realizada una búsqueda, puede realizar una búsqueda más específica sobre ese resultado si tilda la opción "Buscar sobre este resultado".

### Enviar pedido por correo o copiar al portapapeles
Puede enviar un pedido por correo electrónico haciendo clic en Archivo > Enviar pedido por correo. Esto abrirá una nueva ventana para crear un nuevo mail donde el cuerpo del mail contendrá la información del pedido. Se puede personalizar el nombre del asunto desde Opciones > Personalizar. Para copiar un pedido en el portapapeles basta con hacer clic en Archivo > Copiar pedido al portapapeles.

## Personalización y carga inicial de datos
La personalización de algunos campos no tiene una interfaz gráfica y se debe modificar editando el archivo Datos.ini. Para ello, abrirlo en cualquier editor de texto plano y modificar los siguientes parámetros en formato Clave = Valor:

### [SUCURSALES]
Especificar bajo este campo un índice seguido del nombre de cada sucursal, por ejemplo:
```
[SUCURSALES]
0=Central
1=San Martin
2=Recoleta
```

### [TECNICOS]
Especificar bajo este campo un índice seguido del nombre del técnico, por ejemplo:
```
[TECNICOS]
0=FGF
1=SISTEMAS
2=OTRO
```
### [FGWIKI]
Especificar bajo este campo la clave Mostrar con un valor 1 o 0. Si es 1, se mostrará la opción "Abrir Wiki" en el menú del programa.
Además, especificar la ruta del ejecutable FGWiki.exe en la clave RutaExe en caso de establecerse en 1. Por ejemplo,
```
[FGWIKI]
Mostrar=1
RutaExe=C:\Programas\FGWiki\FGWiki.exe
```

### [BACKUPS]
Especificar bajo este campo las claves Usar, CantidadDeCopias y Ruta.  Usar puede ser 1 o 0, y sirve para habilitar los backups automáticos. CantidadDeCopias especifica cuántos backups rotativos se harán por cada archivo respaldado, y Ruta es la carpeta donde se almacenarán los respaldos rotativos. Se puede utilizar el comodín &f, que se reemplazará por la ruta actual desde donde se ejecuta el Sistema de Pedidos Lite. Por ejemplo, para guardar 4 copias de cada archivo en la carpeta Backups del programa:
```
[BACKUPS]
Usar=1
CantidadDeCopias=4
Ruta=&f\Backups
```

### Campos [FLAGS], [FONT], [MAIN], etc.
El resto de los campos no es necesario modificarlos, pues se modifican dentro de la misma interfaz del Sistema de Pedidos Lite, especialmente en el menú Opciones > Personalizar.

## Funciones avanzadas
### Resetear el Sistema de Pedidos Lite
Para inicializar el Sistema de Pedidos Lite se deben borrar los archivos Pedidos.ini, Informes.ini y Enlace.ini. Luego, al iniciar el Sistema de Pedidos Lite se recibirá la advertencia de que estos archivos no existen y volverán a generarse con los valores por defecto.

### Menú ? > Funciones experimentales
Este menú contiene algunas opciones adicionales que son sólo para fines de desarrollo. Salvo la opción "Estadísticas", no hay información de interés para el usuario final.
