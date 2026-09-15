# Instalación de la demostración en Power Apps y SharePoint

## Requisitos

- Cuenta de Microsoft 365 con acceso a SharePoint y Power Apps.
- Permiso para crear listas en un sitio de pruebas.
- El archivo `sample-data/datos-demo.xlsx`.

Trabaja en un sitio independiente y utiliza únicamente los datos ficticios incluidos.

## 1. Crear las listas

Crea cinco listas en SharePoint con estos nombres e importa las filas de la hoja equivalente del Excel:

- `UsuariosDemo`
- `ServiciosDemo`
- `CategoriasDemo`
- `ProtocolosDemo`
- `LecturasDemo`

Configura `FechaVigencia` y `FechaLectura` como fecha y hora; `Version` y `VersionLeida` como número; `Activo` y `Publicado` como Sí/No. El resto puede ser texto de una línea, salvo `Descripcion`, que puede ser texto de varias líneas. SharePoint crea una columna `Title`; puedes ocultarla en los formularios o usarla como identificador técnico.

## 2. Crear la aplicación

1. En Power Apps, crea una aplicación de lienzo para tableta.
2. Añade SharePoint como origen de datos.
3. Conecta las cinco listas del sitio de pruebas.
4. Crea una pantalla de catálogo con un cuadro de búsqueda y una galería.
5. Crea una pantalla de detalle con el título, descripción, versión y botón de lectura.

## 3. Resolver el perfil

Durante una prueba puedes seleccionar un usuario ficticio:

```powerfx
Set(varUsuarioActual; "ana.demo@example.invalid");;
Set(
    varPerfil;
    LookUp(
        UsuariosDemo;
        Lower(CorreoTexto) = Lower(varUsuarioActual) And Activo = true
    )
)
```

Para un entorno real, sustituye el valor de prueba por `Lower(User().Email)` y revisa los permisos del origen.

## 4. Filtrar el catálogo

Asigna a `Items` de la galería:

```powerfx
SortByColumns(
    Filter(
        ProtocolosDemo;
        Publicado = true And
        (ServicioCodigo = varPerfil.ServicioCodigo Or ServicioCodigo = "TODOS") And
        (CategoriaCodigo = varPerfil.CategoriaCodigo Or CategoriaCodigo = "TODAS") And
        (
            IsBlank(txtBuscar.Text) Or
            StartsWith(Lower(Titulo); Lower(txtBuscar.Text))
        )
    );
    "Titulo";
    SortOrder.Ascending
)
```

## 5. Registrar una lectura

Asigna al botón de lectura:

```powerfx
Patch(
    LecturasDemo;
    Defaults(LecturasDemo);
    {
        Title: Text(GUID());
        Codigo: Text(GUID());
        UsuarioCorreo: varUsuarioActual;
        ProtocoloCodigo: ThisItem.Codigo;
        VersionLeida: ThisItem.Version;
        FechaLectura: Now()
    }
)
```

Según la configuración regional de Power Apps, quizá debas cambiar los separadores `;` por `,` y `;;` por `;`.

## 6. Validación mínima

1. Ana Demo debe ver PR-001 y PR-002.
2. Bruno Demo debe ver PR-001 y PR-003.
3. Carla Demo debe ver PR-001; PR-004 no aparece porque no está publicado.
4. Registrar una lectura debe crear una fila con usuario, protocolo, versión y fecha.
5. Un usuario sin fila activa no debe acceder al catálogo.

La demo web de `demo/` reproduce estas reglas localmente y permite validarlas sin una cuenta Microsoft.
