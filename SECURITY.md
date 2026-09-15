# Seguridad de los datos del TFG

## Estado verificado

El repositorio anterior se eliminó y este repositorio se creó de nuevo a partir de contenido revisado. La reconstrucción demostrativa añadida después utiliza únicamente datos sintéticos y no recupera archivos de la versión anterior.

La revisión de esta muestra confirma:

- La demo web funciona sin conexión a servicios externos.
- Los usuarios, servicios, categorías, protocolos y lecturas son ficticios.
- Los correos y enlaces de ejemplo utilizan `example.invalid`.
- El Excel no contiene punteros Git LFS, macros ni conexiones externas.
- No se incluyen exportaciones de Power Apps, documentos institucionales ni datos de trabajadores.
- El workflow comprueba todas las rutas alcanzables, el contenido del fixture CSV, la huella del Excel aprobado y patrones básicos de DNI, NIE y correos no ficticios.

Estas comprobaciones cubren el contenido de este repositorio. No permiten controlar copias que terceros hubieran descargado del repositorio anterior.

## Archivos aprobados

Se permiten la documentación, la demo web, el libro `sample-data/datos-demo.xlsx`, el fixture sintético y el workflow. Cualquier ruta nueva hace fallar el control automático hasta que se revise y se añada expresamente.

## Uso de datos

No añadas nombres, correos, identificadores, documentos o capturas del entorno institucional. Mantén los datos de demostración separados del entorno real. Antes de importar el Excel, crea un sitio de SharePoint exclusivo para pruebas.

En producción, configura los permisos en SharePoint. El filtrado de Power Apps mejora la experiencia de usuario, pero no sustituye el control de acceso del origen.

Si una copia antigua contenía credenciales, deben revocarse en el proveedor correspondiente. Si los registros retirados correspondían a personas reales, conserva la comunicación con el responsable de seguridad o protección de datos de la entidad.

## Cambios futuros

Antes de añadir una exportación `.msapp`, un paquete, una captura o un documento:

1. Revisa conexiones, identificadores, propiedades y metadatos.
2. Comprueba el contenido con una cuenta y un sitio de pruebas.
3. Sustituye cualquier identidad o documento por ejemplos ficticios.
4. Revisa el cambio en privado.
5. Actualiza el workflow únicamente después de aprobar todos los archivos.
