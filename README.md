# CapturaDatos

Aplicación Web Progresiva (PWA) desarrollada como proyecto universitario para la captura, gestión y almacenamiento eficiente de datos. El sistema está diseñado con una arquitectura moderna de cliente-servidor, permitiendo operar tanto en línea como sin conexión a internet mediante IndexedDB y Service Workers, garantizando la sincronización con una base de datos centralizada.

## 🚀 Características Principales

* **Funcionamiento Offline (PWA):** Almacenamiento local temporal y persistente mediante IndexedDB para asegurar la operatividad ante fallos de conectividad.
* **Autenticación Segura con Google:** Integración de Google Identity Services en el cliente combinada con verificación de tokens JWT (JSON Web Tokens) en el servidor backend.
* **Gestión Dinámica de Registros:** Creación, edición, consulta y eliminación de registros y campos personalizados vinculados al usuario autenticado.
* **Control de Acceso Basado en Roles/Propietario:** Aislamiento de datos por usuario en el backend para garantizar la privacidad e integridad de la información.
* **Respaldo y Restauración:** Funcionalidad para exportar e importar respaldos completos de la base de datos en formato estructurado.

## 🛠️ Tecnologías Utilizadas

### Frontend
* **HTML5 / CSS3 / JavaScript (Vanilla)**
* **Service Workers & Web App Manifest** (para capacidades PWA)
* **IndexedDB** (almacenamiento local)

### Backend y Base de Datos
* **Node.js & Express.js** (API RESTful centralizada)
* **Prisma ORM** (maquetación y consultas a base de datos)
* **PostgreSQL** (base de datos relacional en la nube / local)
* **Google Auth Library & Bcrypt.js** (seguridad y autenticación)

## ⚙️ Instalación y Configuración Local

Sigue estos pasos para levantar el entorno de desarrollo en tu computadora:

1. **Clona el repositorio:**
   ```bash
   git clone [https://github.com/tovar04-cyber/CapturaData.git](https://github.com/tovar04-cyber/CapturaData.git)
   cd CapturaData
