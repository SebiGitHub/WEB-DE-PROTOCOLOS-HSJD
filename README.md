# WEB-DE-PROTOCOLOS-HSJD

## Qué es
Aplicación (TFG) para gestionar y difundir protocolos hospitalarios de forma centralizada, controlando visibilidad por perfil/servicio y registrando la lectura de cada protocolo (incluyendo el control de versiones).

> Objetivo: reducir fricción al encontrar protocolos y asegurar que las versiones nuevas vuelvan a mostrarse aunque el usuario ya haya leído una versión anterior.

## Stack
- Power Apps (Canvas App, Backend y FrontEnd de la web)
- SharePoint (listas como almacenamiento)
- Excel (Controla la correcta información de cada usuario)

## Features
- Catálogo de protocolos con filtrado por **categoría**, **servicio** y/o **perfil**
- Lectura con trazabilidad: marcar protocolo como leído por usuario
- Control de versiones: cuando se publica una versión nueva, vuelve a aparecer como “pendiente”
- Búsqueda por texto y navegación por secciones (pendientes / leídos / todos)
- Administración: alta/edición de protocolos y metadatos

## Capturas/GIF
Carpeta `/docs/screenshots/`.
<img width="1112" height="831" alt="home" src="https://github.com/user-attachments/assets/3975babf-0486-4ced-8a4a-69fdf81d0e37" />

## Cómo importar y conectar
1. Clona el repositorio.
2. Abre el proyecto en Power Apps Studio.
   - Si tienes un `.msapp`: ábrelo/importa la app.
   - Si exportas como solución: importa la solución en tu entorno.
3. Configura los conectores:
   - SharePoint Site: [`[URL]`](https://pssjd.sharepoint.com/sites/HospitalAljarafe.InformacionDocumental/SitePages/CollabHome.aspx)
   - Listas necesarias: `[ProtocolosPublicados]`, `[T_Usuarios_LDAP]`, `[T_Protocolos_Vistos]`
4. Ajusta variables/constantes del entorno (URLs, nombres de listas, etc.)
5. Publica la app y compártela con los roles adecuados.

## Qué aprendí
- Diseñar una app orientada a **procesos reales** (acceso, permisos, trazabilidad)
- Modelar datos en SharePoint para soportar filtrado, lectura y versionado
- Reducir errores de UX con reglas claras: “leído”, “no leído”, “nueva versión”
- Mejorar mantenibilidad documentando estructura, conectores y lógica clave
