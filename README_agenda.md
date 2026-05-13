# Mi Agenda Personal Premium

Mi Agenda Personal Premium es una aplicación de escritorio desarrollada en Python con Tkinter. Su objetivo es permitir al usuario escribir entradas personales tipo diario, agenda o bitácora, guardarlas localmente, exportarlas en distintos formatos y compartirlas mediante correo electrónico o WhatsApp.

La aplicación está diseñada para funcionar en Windows, con una interfaz gráfica moderna, historial de entradas, calendario, modo claro/oscuro, configuración de carpeta sincronizada en la nube, protección mediante PIN e inicio automático con Windows.

---

## Descripción general

Este proyecto permite registrar pensamientos, actividades diarias, reflexiones o notas importantes desde una interfaz sencilla y organizada.

La aplicación guarda las entradas en formato JSON dentro de una carpeta local del usuario y permite exportarlas en formatos más comunes como TXT, PDF y Word.

También incorpora funciones de productividad como contador de palabras, contador de caracteres, calendario mensual, configuración de tema visual, envío por WhatsApp Web y preparación de correos mediante la aplicación de correo predeterminada del sistema.

---

## Funcionalidades principales

- Crear entradas personales tipo diario o agenda.
- Seleccionar el estado del día.
- Escribir título y contenido de la entrada.
- Guardar entradas en formato JSON.
- Ver historial de entradas guardadas.
- Abrir entradas desde el historial.
- Contador automático de palabras.
- Contador automático de caracteres.
- Exportar entradas en TXT.
- Exportar entradas en PDF.
- Exportar entradas en Word/DOCX.
- Enviar contenido por correo electrónico.
- Enviar contenido por WhatsApp Web.
- Guardar archivos en una carpeta sincronizada con la nube.
- Configurar carpeta de nube local, como OneDrive, Google Drive o Dropbox.
- Cambiar entre tema oscuro y tema claro.
- Configurar PIN de acceso.
- Mostrar calendario mensual.
- Navegar entre meses del calendario.
- Activar o desactivar inicio automático con Windows.
- Guardar posición de la ventana al cerrar la aplicación.

---

## Tecnologías utilizadas

- Python 3
- Tkinter
- ttk
- JSON
- pathlib
- webbrowser
- ReportLab
- python-docx
- PyInstaller

---

## Estructura recomendada del proyecto

```text
agenda/
├── diario_agenda.py
├── diario_agenda.spec
├── README.md
├── requirements_agenda.txt
└── .gitignore