# Office Add-in - Gestor Documental

## Descripción

Este es un complemento (Add-in) de Microsoft Office desarrollado con React y TypeScript que permite gestionar metadatos de documentos almacenados en SharePoint. El add-in se integra con Microsoft Graph API para proporcionar una interfaz completa de gestión documental dentro de Word, Excel y PowerPoint.

## Características Principales

### 1. Gestión de Metadatos de Documentos

El add-in permite gestionar metadatos estructurados para diferentes tipos de bibliotecas de documentos en SharePoint:

- **DOCUMENTOS_CLIENTES**: Gestión de documentos relacionados con clientes
  - Cliente (con búsqueda avanzada)
  - Asunto
  - Subasunto
  - Tipo de Documento
  - Subtipo de Documento

- **DOCUMENTOS_ADMIN_RRHH**: Administración de Recursos Humanos
  - Carpeta RRHH

- **DOCUMENTOS_CONSULADO_AUSTRALIA**: Documentos del Consulado de Australia
  - Nivel 1
  - Nivel 2

- **DOCUMENTOS_CONTADURIA**: Documentos de Contaduría
  - Tema
  - Subtema
  - Tipo de Documento

- **DOCUMENTOS_DECLARACIONES_JURADAS**: Declaraciones Juradas Profesionales
  - Carpeta DJ

- **DOCUMENTOS_INTERNO**: Documentos internos con estructura jerárquica
  - Carpeta 1
  - Carpeta 2
  - Carpeta 3
  - Carpeta 4
  - Carpeta 5
  - Carpeta 6
  - Carpeta 7

### 2. Edición de Títulos de Documentos

- Edición del título del documento directamente desde el add-in
- Detección automática de documentos existentes en SharePoint
- Generación automática de DocID para nuevos documentos
- Renombrado automático de archivos con formato estandarizado

### 3. Autenticación con Microsoft 365

- Autenticación mediante MSAL.js (Microsoft Authentication Library)
- Flujo OAuth 2.0 con PKCE para aplicaciones de página única (SPA)
- Renovación automática de tokens de acceso
- Gestión de sesiones de usuario

### 4. Integración con SharePoint

- Conexión directa con SharePoint mediante Microsoft Graph API
- Carga y actualización de metadatos en listas de SharePoint
- Sistema de cola para procesamiento asíncrono de metadatos
- Detección automática de bibliotecas de documentos
- Sincronización bidireccional de metadatos

### 5. Sistema de Caché Inteligente

- Caché local para mejorar el rendimiento
- Actualización en segundo plano de datos obsoletos
- Búsqueda instantánea en grandes volúmenes de datos
- Indicadores visuales del estado del caché

### 6. Funcionalidades Adicionales

- **Búsqueda avanzada**: Búsqueda en tiempo real con filtrado de clientes
- **Carga de metadatos existentes**: Detección y carga automática de metadatos de documentos existentes
- **Validación de datos**: Validación de campos requeridos antes de guardar
- **Mensajes informativos**: Sistema de notificaciones para el usuario
- **Soporte multi-aplicación**: Compatible con Word, Excel y PowerPoint

## Arquitectura del Proyecto

```
src/
├── components/
│   ├── App.tsx                    # Componente principal de la aplicación
│   ├── DocumentTitleEditor.tsx    # Editor de títulos y metadatos
│   ├── metadata-components.tsx    # Componentes de formularios de metadatos
│   ├── Header.tsx                 # Encabezado de la aplicación
│   └── ...
├── services/
│   ├── sharepoint-data.service.ts # Servicio de datos de SharePoint
│   ├── cache-manager.ts           # Gestor de caché
│   └── cache-diagnostics.tsx      # Diagnósticos de caché
├── config/
│   └── authConfig.ts              # Configuración de autenticación
├── constants/
│   └── sharepointConfig.ts        # Configuración de SharePoint
├── interfaces/
│   └── interfaces.ts              # Definiciones de tipos TypeScript
└── utilities/
    ├── microsoft-graph-helpers.ts # Helpers para Microsoft Graph
    ├── office-apis-helpers.ts      # Helpers para Office.js
    └── sharepoint_helpers.ts      # Helpers para SharePoint
```

## Tecnologías Utilizadas

- **React**: Framework de UI
- **TypeScript**: Lenguaje de programación
- **MSAL.js**: Biblioteca de autenticación de Microsoft
- **Office.js**: API de Office Add-ins
- **Microsoft Graph API**: API para acceder a datos de Microsoft 365
- **Office UI Fabric React**: Componentes de UI de Microsoft
- **Webpack**: Bundler de módulos
- **Axios**: Cliente HTTP

## Requisitos Previos

- Node.js versión 18.20.2 o superior
- npm versión 10.5.0 o superior
- TypeScript versión 5.4.3 o superior
- Una cuenta de Microsoft 365
- Un tenant de Azure Active Directory
- Office en Windows versión 16.0.6769.2001 o superior

## Instalación y Configuración

### 1. Registro de Aplicación en Azure

1. Navegue al [Portal de Azure - Registros de aplicaciones](https://go.microsoft.com/fwlink/?linkid=2083908)
2. Inicie sesión con credenciales de administrador de su tenant de Microsoft 365
3. Seleccione **Nuevo registro**
4. Configure los siguientes valores:
   - **Nombre**: `HenkaAddin` (o el nombre que prefiera)
   - **Tipos de cuenta admitidos**: Cuentas en cualquier directorio organizativo y cuentas personales de Microsoft
   - **URI de redirección**: Seleccione "Aplicación de página única (SPA)" y establezca `https://localhost:3000/login/login.html`
5. Copie el **ID de aplicación (cliente)** para usarlo en el siguiente paso

### 2. Configuración del Proyecto

1. Abra el archivo `/login/login.ts` y reemplace `YOUR APP ID HERE` con el ID de aplicación copiado
2. Abra el archivo `/logout/logout.ts` y reemplace `YOUR APP ID HERE` con el ID de aplicación copiado
3. Abra un **Símbolo del sistema como administrador**
4. Navegue a la raíz del proyecto
5. Ejecute `npm install`
6. Ejecute `npx office-addin-dev-certs install --machine` para instalar certificados SSL

### 3. Ejecutar el Proyecto

1. En el símbolo del sistema, ejecute `start npm start` para iniciar el servidor de desarrollo
2. En otro símbolo del sistema, ejecute `npm run sideload` para cargar el add-in en Office
3. El add-in aparecerá en el panel lateral de Office con un botón **Open Add-in**

## Uso del Add-in

### Inicio de Sesión

1. Haga clic en el botón **Open Add-in** en el panel lateral de Office
2. Haga clic en **Conectarse a Office 365** para iniciar sesión
3. Complete el proceso de autenticación en la ventana emergente
4. La primera vez se le pedirá consentimiento para los permisos del add-in

### Gestión de Documentos

1. **Seleccionar Biblioteca**: Elige la biblioteca de documentos apropiada desde el menú desplegable
2. **Editar Título**: Modifica el título del documento si es necesario
3. **Completar Metadatos**: Completa los campos de metadatos según el tipo de biblioteca seleccionada
4. **Guardar**: Los metadatos se guardan automáticamente en SharePoint

### Características Especiales

- **Búsqueda de Clientes**: En la biblioteca CLIENTES, puedes buscar clientes escribiendo en el campo de búsqueda
- **Carpetas Jerárquicas**: En la biblioteca INTERNO, las carpetas se cargan de forma dependiente (Carpeta2 depende de Carpeta1, etc.)
- **Detección Automática**: El add-in detecta automáticamente si un documento ya existe en SharePoint y carga sus metadatos

## Estructura de Metadatos

Cada biblioteca tiene sus propios campos de metadatos definidos en SharePoint:

- Los metadatos se almacenan en listas de SharePoint específicas
- Los campos se validan antes de guardar
- Los metadatos existentes se cargan automáticamente para documentos ya guardados

## Sistema de Cola de Metadatos

El add-in utiliza un sistema de cola para procesar actualizaciones de metadatos:

- Las actualizaciones se envían a una lista de SharePoint especial (Metadata Queue)
- Un proceso en segundo plano procesa las solicitudes de la cola
- Esto permite operaciones asíncronas y mejora el rendimiento

## Solución de Problemas

### El add-in no se carga

- Verifique que el servidor de desarrollo esté ejecutándose (`npm start`)
- Asegúrese de que los certificados SSL estén instalados correctamente
- Revise la consola del navegador para errores

### Error de autenticación

- Verifique que el ID de aplicación esté correctamente configurado
- Asegúrese de que el URI de redirección coincida exactamente con el configurado en Azure
- Revise los permisos de la aplicación en Azure Portal

### Los metadatos no se cargan

- Verifique la conexión a Internet
- Revise que tenga permisos para acceder a las listas de SharePoint
- Compruebe que las listas de SharePoint existan y tengan los campos correctos

## Desarrollo

### Scripts Disponibles

- `npm start`: Inicia el servidor de desarrollo
- `npm run build`: Construye la versión de producción
- `npm run sideload`: Carga el add-in en Office
- `npm run lint`: Ejecuta el linter de TypeScript
- `npm run validate`: Valida el manifiesto del add-in

### Estructura de Código

El código está organizado en módulos:

- **Components**: Componentes React reutilizables
- **Services**: Lógica de negocio y servicios de datos
- **Utilities**: Funciones auxiliares y helpers
- **Config**: Archivos de configuración
- **Interfaces**: Definiciones de tipos TypeScript

## Mejoras Recientes

- ✅ Código limpio: Eliminación de comentarios innecesarios y código comentado
- ✅ Documentación mejorada: Comentarios importantes traducidos al inglés con formato JSDoc
- ✅ Optimización: Eliminación de console.logs de depuración
- ✅ Estructura: Organización mejorada del código y documentación

## Notas Importantes

### Carpetas Excluidas del Repositorio

Las siguientes carpetas están excluidas del control de versiones (ver `.gitignore`):

- `prod 02-10-20225/` - Versiones de producción antiguas
- `version de prod 07-11-2025/` - Versiones de producción
- `dist/` - Archivos compilados
- `node_modules/` - Dependencias de Node.js
- `*.zip` - Archivos comprimidos

### Sistema de Caché

El sistema de caché inteligente está optimizado para la biblioteca `DOCUMENTOS_CLIENTES`:

- Carga inicial de datos en caché local
- Actualización en segundo plano cuando los datos están obsoletos
- Búsqueda instantánea en grandes volúmenes de datos
- Indicadores visuales del estado del caché

## Soporte

Para problemas o preguntas:

- Revise la documentación de Office Add-ins
- Consulte Microsoft Q&A con la etiqueta `office-js-dev`
- Cree un issue en el repositorio del proyecto

## Licencia

Este proyecto es parte de los ejemplos de Office Add-in de Microsoft.
