# WEB-DE-PROTOCOLOS-HSJD
❓¿Qué es?
WEB-DE-PROTOCOLOS-HSJD es una aplicación desarrollada con Microsoft Power Apps que se integra con listas de SharePoint. Su función principal es servir como un gestor de protocolos internos: asegurar que cada empleado solo tenga acceso —y vea— los protocolos (documentos, guías, normativas, etc.) que le correspondan en función de su categoría profesional y su servicio.

😄¿Qué problema resuelve?
Evita que un empleado vea protocolos que no le conciernen, reduciendo ruido de información y mejorando la seguridad y organización interna. Facilita la gestión de documentos y protocolos dentro de una institución, haciendo más eficiente el acceso a lo necesario. Esto es útil en entornos profesionales donde hay múltiples categorías de usuarios y distintos niveles de permisos.

🛠¿Cómo está construido / estructura técnica?
- La aplicación vive en Power Apps (la carpeta App/ del repositorio contiene el archivo de la app o su exportación JSON). 
- Usa listas de SharePoint como backend (almacenamiento de datos). 
- Incluye también flujos de automatización (carpeta Flows/), presumiblemente para manejar lógica de permisos, asignaciones automáticas o notificaciones. 
- Tiene una estructura organizada: carpetas para configuración (Config/), documentación (Docs/), assets (imágenes/íconos), etc. 

## Estructura

- App/: contiene el archivo .msapp o los JSON exportados
- Assets/: imágenes e iconos
- Config/: scripts para listas de SharePoint
- Docs/: diagramas, capturas, manual
