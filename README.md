# Gestión de protocolos hospitalarios — TFG

Proyecto académico de DAM para centralizar la consulta de protocolos mediante Power Apps y SharePoint. La documentación original describe filtrado por perfil y servicio, registro de lectura y control de versiones.

## Estado de esta versión

Se han retirado de la versión actual los datos, exportaciones de la aplicación, documentos, capturas y recursos institucionales pendientes de revisión. No se distribuye una aplicación ejecutable en esta versión. El 14 de septiembre de 2026 se ha reconstruido el historial de `main` desde el commit inicial revisado, conservando solo documentación y el ejemplo ficticio. El repositorio debe permanecer **privado** hasta confirmar también la retirada de objetos antiguos y cachés con GitHub Support. Consulta [SECURITY.md](SECURITY.md) para el alcance y las comprobaciones.

## Stack del proyecto original

- Power Apps (Canvas App)
- SharePoint (listas)
- Excel y Access para preparación de datos

## Ejemplo sintético

`examples/protocolos-sinteticos.csv` contiene únicamente registros inventados. No contiene datos de trabajadores ni protocolos clínicos reales.

## Diseño que se pretende demostrar

- Catálogo y búsqueda de protocolos.
- Asociación de protocolos a perfiles y servicios.
- Registro de lectura asociado a una versión.
- Separación entre visibilidad en la interfaz y permisos efectivos del origen de datos.

Estas funciones describen el proyecto original; deben verificarse con una nueva exportación saneada antes de anunciar una demo disponible.

## Próximas mejoras

1. Preparar un entorno de demostración independiente con datos sintéticos.
2. Exportar la aplicación sin conexiones, identificadores ni datos institucionales.
3. Documentar el esquema de listas y comprobar permisos en el origen.
4. Añadir capturas y una demostración verificadas antes de su publicación.
