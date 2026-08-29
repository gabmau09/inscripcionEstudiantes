# Sistema de Inscripción de Estudiantes

Un sistema web integral de inscripción a cursos y exámenes de suficiencia diseñado para el Instituto de Informática de la Universidad Nacional de Piura (UNP). 

Este proyecto moderniza el flujo de inscripción reemplazando la inscripción presencial por una Web App personalizada, conectada directamente a Google Sheets y Google Drive como base de datos y gestor documental, alojada mediante un iframe para su despliegue público.

## Características Principales

*   **Base de Datos Centralizada y Dinámica:** Lee la oferta académica (Cursos, Modalidades, Límites de vacantes y Enlaces) en tiempo real desde una hoja maestra en Google Sheets (`BD CURSOS`), permitiendo actualizar la información sin modificar el código.
*   **Gestión Documental Automática:** Al registrar la inscripción, el sistema crea rutas de carpetas inteligentes en Google Drive (clasificadas por curso, grupo y alumno) y guarda automáticamente los PDFs de DNI, Solicitud y Vouchers de pago.
*   **Auto-Paginación en Google Sheets:** Controla los límites de vacantes por grupo. Al llenarse un aula, el sistema clona automáticamente una pestaña de plantilla y genera un nuevo grupo para continuar el registro sin interrupciones.
*   **Lógica de Identificación (UNP vs Externos):** Interfaz adaptable que solicita datos específicos a alumnos UNP (con un buscador autocompletado y validado de Escuelas Profesionales) y genera códigos correlativos automáticos de 10 caracteres (ej. `EXT0000009`) para usuarios externos.
*   **Seguridad y "Candado Digital":** Prevención de envíos duplicados mediante el bloqueo instantáneo de la interfaz y los botones tras el primer clic de envío, protegiendo la integridad de la base de datos.
*   **Interfaz Moderna y Responsiva:** Diseño limpio basado en la paleta institucional, con ventanas modales personalizadas para notificaciones e instrucciones de pago.

## Tecnologías Utilizadas

*   **Frontend:** HTML5, CSS3, JavaScript (Vanilla).
*   **Backend:** Google Apps Script (JavaScript).
*   **Base de Datos y Almacenamiento:** Google Sheets API y Google Drive API.
*   **Despliegue (Hosting):** Google Apps Script Web App enmascarada a través de Iframe para alojamiento en cualquier servidor web.

## Estructura del Proyecto

El repositorio está dividido en dos partes principales para facilitar el despliegue del frontend y el mantenimiento del código de Apps Script:

```text
├── index.html       # Archivo de enmascarado (Iframe) utilizado para el despliegue web público.
├── README.md        # Documentación del proyecto.
└── src/             # Código fuente implementado en Google Apps Script.
    ├── Code.gs      # Backend (Lógica de base de datos, creación de carpetas en Drive y endpoints).
    └── Index.html   # Frontend de la Web App (Estructura UI, estilos y conexión con el backend).
```

## Autor

*   **Gabriela Tahis Mauriola Valdiviezo**
*   **Organización:** Instituto de Informática - Universidad Nacional de Piura (UNP)

---
*Proyecto desarrollado en el año 2026 para la optimización de procesos académicos.*
