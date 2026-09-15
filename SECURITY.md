# Seguridad de los datos del TFG

## Estado verificado — 15 de septiembre de 2026

El repositorio anterior se eliminó y se creó uno nuevo con el mismo nombre a partir de una copia revisada. El repositorio nuevo está privado y su historial comenzó con un único commit raíz: `514cf2d513bb3915aeb6285c801041322d6e474c`.

La revisión remota confirmó:

- Solo existe la rama `main`.
- No hay etiquetas, pull requests, releases ni forks.
- El árbol contiene únicamente `README.md`, `SECURITY.md`, `.gitignore`, el workflow de control y el CSV sintético.
- No hay `.gitattributes` ni punteros de Git LFS en esos archivos.
- Un commit del repositorio eliminado consultado por su identificador devuelve `404 Not Found`.
- El control automático del nuevo repositorio terminó correctamente en GitHub Actions.

Estas comprobaciones acreditan el estado del repositorio nuevo y de las referencias consultables mediante GitHub. No permiten borrar ni verificar copias que terceros hubieran descargado cuando el repositorio anterior era público.

## Prevención de nuevas exposiciones

No reutilices ni mezcles clones del repositorio anterior. Continúa el trabajo desde un clon nuevo del repositorio recreado. Conserva las exportaciones institucionales fuera de Git y no amplíes la lista de archivos permitidos sin una revisión privada.

El workflow `public-data-check.yml` comprueba en todos los commits alcanzables:

- que solo existan las cinco rutas permitidas;
- que sean archivos normales;
- que el CSV coincida exactamente con el ejemplo ficticio aprobado.

El workflow no determina por sí solo si un texto Markdown contiene información sensible. Los cambios de documentación requieren revisión humana.

Si los registros retirados correspondían a personas reales, conserva la comunicación realizada con la entidad o con su responsable de seguridad o protección de datos. Si alguna copia antigua contenía credenciales, deben revocarse en el proveedor correspondiente.

## Publicación de una demostración

La versión actual no incluye una aplicación ejecutable. Antes de añadir una exportación de Power Apps, documentación o capturas:

1. Usa un entorno de Power Apps y SharePoint independiente.
2. Emplea identidades y registros completamente ficticios.
3. Revisa conexiones, identificadores, propiedades, metadatos y recursos incluidos.
4. Comprueba los permisos reales del origen de SharePoint con usuarios de prueba.
5. Revisa los archivos en privado antes de modificar la lista permitida del workflow.

[Procedimiento oficial de GitHub para retirar datos sensibles](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository).
