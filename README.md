--------------------------------------------------------
Author: Pablo Abad Aparicio  (pablo.abad96@gmail.com)

Creation date: 19/09/2024

--------------------------------------------------------

____**Referencias activadas**____
- Visual Basic For Applications
- Microsoft Outlook 16.0 Object Library
- OLE Automation
- Microsoft Office 16.0 Object Library
- Microsoft Forms 2.0 Object Library
- Microsoft Word 16.0 Object Library
- Microsoft Scripting Runtime

Programa para cargar el texto selecionado de un mensaje
de Outlook activo en un word. La ruta a los posibles ficheros
word en los que cargar la información se obtiene de un Excel
especificado en la variable `sXlsPath` de la función
M04_ExcelTableToArray.fExcelTableToList

El Excel tiene que ser de una única hoja con un objeto tabla
que contenga dos columnas: Nombre identificativo, Path.
De esta forma al carga  nuevos registros en el objeto tabla
la macro los incorporará directamente.

Al principio de la macro se muestra el formulario `UserForm1`.
Contiene una lista de las rutas disponibles para que el usuario
seleccione el fichero donde quiera cargar la información.

____ *Estructura de modulos**____
- [`UserForm`](/UserForm1.frm) Formulario con ListBox y dos botones  
(Aceptar y Cancelar) para que el usuario pueda seleccionar el fichero de 
destino de una lista de opciones.
- [`M00_Documentación`](/M00_Documentacion.bas) Misma información que en 
README.md. Está repedita  ya que en el Editor de Visual Basic no existe 
README.md.
- [`M00_GenericFunction`](/M00_GenericFunctions.bas) Modulo de funciones 
de uso general, utilizables para varios proyecto.
- [`M01_MailsToWord.bas`](/M01_MailsToWord.bas) Subrutina de entrada al 
programa, el main de la macro.
- [`M02_MailDict.bas`](/M02_MailDict.bas) Crea un diccionario de los 
elementos del email que se copiarán en el word (From, To, Date, Subject y 
SelectedText).
- [`M03_AppendToWord.bas`](/M03_AppendToWord.bas) Añade los elementos
del diccionario del modulo anterior al final de un ficheor word.
-[`M04_ExcelTableToArray.bas`](/M04_ExcelTableToArray.bas) Convierte
la tabla del Excel que contiene las rutas de los ficheros en un array
que se utilizara como fuente de la lista del formulario.

____**MEJORAS**____
- Si un word ya está abierto poder escribir en él.
- Gestion de errores
- Sugerencia de ruta en base al email del emisor.
