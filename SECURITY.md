# Seguridad de los datos del TFG

## 1. Estado y alcance — 14 de septiembre de 2026

El repositorio está privado. La revisión previa encontró 23 commits alcanzables desde `main`; los archivos originales seguían presentes en sus antecesores. La limpieza local no se había reflejado en la rama remota.

Se ha reconstruido `main` conservando el commit inicial, que contiene únicamente un README descriptivo revisado, y una nueva instantánea de los cinco archivos revisados. Los 22 commits posteriores de la cadena anterior dejan de formar parte de la historia de `main`. Se conserva el contenido actual útil; no se recuperan exportaciones ni documentos institucionales. No se ha utilizado git-filter-repo en esta intervención remota.

Solo se encontró la referencia `refs/heads/main`, sin etiquetas. No se encontraron pull requests, issues ni releases; GitHub informa de cero forks. Las dos ejecuciones de Actions anteriores a esta intervención no tienen artefactos. Esto no acredita que nunca existieran copias externas, adjuntos o descargas. La colección global de comentarios no pudo consultarse mediante la conexión disponible.

Tras la actualización se verificó que main solo alcanzaba el commit inicial y la nueva instantánea. El control de historial pasó en GitHub Actions (ejecución 34894428026). Se comprobó también que un objeto de commit antiguo seguía siendo recuperable por su identificador mediante la conexión autenticada: la purga del servidor sigue pendiente. Esto no significa que el repositorio privado sea accesible sin autorización.

**No volver a hacer público todavía.** Desconectar la historia antigua de una rama no borra físicamente los objetos del servidor ni sus vistas en caché. La comprobación automática tampoco lo hace.

## 2. Cómo comprobar el paso 4 de la guía anterior

Este paso era verificar el resultado antes de actualizar GitHub; no era necesario crear una demo. La actualización remota de esta intervención sustituye el procedimiento local pendiente. No subas ahora una copia antigua mediante `push --mirror`, ni mezcles su historia con la nueva.

Para verificar desde tu equipo, utiliza una carpeta nueva:

```sh
git clone https://github.com/SebiGitHub/WEB-DE-PROTOCOLOS-HSJD.git TFG-verificado
cd TFG-verificado
git remote -v
git log --all --oneline
git ls-files
git log --all --format= --name-only -- App/ Config/ Docs/ Assets/ screenshots/
```

- El remoto debe apuntar únicamente a este repositorio.
- La historia inicial tras esta intervención tendrá dos commits: el inicial y el de reconstrucción. Las mejoras posteriores añadirán commits.
- Los archivos actuales permitidos son README.md, SECURITY.md, .gitignore, .github/workflows/public-data-check.yml y examples/protocolos-sinteticos.csv.
- El último comando no debe mostrar rutas. Esto solo comprueba esas carpetas; revisa también la lista completa de archivos y su contenido.
- El ejemplo debe contener únicamente DEMO-001, con servicio, perfil y documento ficticios.
- El workflow revisa rutas y modos de archivos en todos los commits descargados, y el contenido exacto del CSV sintético. Los textos Markdown requieren revisión humana.

No uses clones antiguos para continuar el trabajo. Si contienen cambios propios pendientes, revisa y traslada únicamente los cambios necesarios a una copia nueva, sin fusionar el historial anterior.

## 3. Pendiente: purga de objetos y copias residuales

GitHub Support debe evaluar la eliminación de objetos antiguos, vistas en caché y posibles referencias internas. Esta solicitud no se ha enviado. Usa el [portal de soporte de GitHub](https://support.github.com/) y proporciona la siguiente información en el ticket privado, sin adjuntar registros personales:

> Solicito revisar la eliminación de datos potencialmente personales del repositorio privado SebiGitHub/WEB-DE-PROTOCOLOS-HSJD. Se retiraron exportaciones, tablas de personal y documentos institucionales. Se ha sustituido la historia de main por una instantánea revisada basada en el commit inicial sin esos archivos. No se utilizó git-filter-repo: no dispongo de su informe First Changed Commits. El primer commit excluido de la cadena es 443f1f51e1dc2c9d2398ce1e9ce471aef53698a6; la antigua punta era 35a8e16412330b1dd7bca47b870437c7d8eabefa. En la revisión no se encontraron pull requests ni forks. Solicito comprobar referencias internas, objetos antiguos, vistas en caché y, si procede, objetos LFS. No se ha certificado la ausencia de LFS. Indiquen cualquier información adicional necesaria para evaluar y completar la purga.

Los identificadores anteriores son referencias para soporte; no son una prueba de borrado. No abras issues públicos con datos ni compartas las exportaciones. Si los registros eran reales, comunica la exposición al responsable de seguridad o protección de datos de la entidad para que evalúe las medidas necesarias. Si hubo credenciales reales, revócalas en su proveedor.

Las copias descargadas por terceros no se eliminan desde GitHub. La aprobación de un workflow no acredita su retirada.

## 4. Reapertura y demostración

Este apartado es distinto del paso 4 de comprobación. No exige publicar ahora ni recuperar los archivos retirados. La versión actual no contiene una aplicación ejecutable.

Antes de plantear la reapertura:
1. Resolver con soporte las copias que GitHub conserve y revisar las copias bajo tu control.
2. Preparar, si quieres mostrar la aplicación, un entorno de Power Apps y SharePoint independiente con identidades ficticias.
3. Revisar exportación, conexiones, identificadores, metadatos, documentos y capturas. No reutilizar archivos institucionales sin revisión y autorización.
4. Comprobar permisos en SharePoint con usuarios de prueba: ocultar elementos en Power Apps no sustituye la autorización del origen.
5. Revisar los nuevos archivos en privado antes de ampliar la lista permitida del workflow.

No se ha creado ni validado esa nueva exportación en esta intervención porque no hay acceso al entorno Power Apps/SharePoint de demostración.

[Procedimiento oficial de GitHub](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository).
