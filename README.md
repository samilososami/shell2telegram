<div align="center">

# Shell2Telegram

Agente de control remoto basado en PowerShell que utiliza la API de Telegram para la gestión de sistemas persistentes a través de una interfaz de chat.

<br>

<img src="https://img.shields.io/badge/PowerShell-5391FE?style=for-the-badge&logo=powershell&logoColor=white" alt="PowerShell" />
<img src="https://img.shields.io/badge/Telegram-26A5E4?style=for-the-badge&logo=telegram&logoColor=white" alt="Telegram" />
<img src="https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white" alt="Windows" />

</div>

---

### ✦ Características

*   **Comandos Remotos**: Ejecución directa de comandos CMD y PowerShell mediante mensajes de Telegram.
*   **Gestión de Sesiones**: Control múltiple de máquinas con visualización de estado y última conexión.
*   **Recuperación de Información**: Módulos integrados para extraer información del sistema (`/info`), llaves de producto de Windows (`/key`) y redes WiFi guardadas (`/wifi`).
*   **Scraping de Datos**: Función para buscar direcciones de correo electrónico en los logs del sistema (`/email`).
*   **Persistencia**: Diseñado para mantenerse activo y reconectarse automáticamente en caso de pérdida de red.

### ✦ Configuración

1.  Cree un bot en Telegram a través de @BotFather y obtenga su **API Token**.
2.  Actualice las variables `$token` y `$chatId` en el script `shell2telegram.ps1`.
3.  Ejecute el script en la máquina objetivo:
    ```powershell
    .\shell2telegram.ps1
    ```

### ✦ Comandos Disponibles

| Comando | Acción |
| :--- | :--- |
| `/help` | Muestra la lista de comandos y ayuda general. |
| `/info` | Genera un reporte detallado del hardware y software del sistema. |
| `/sessions` | Lista todas las sesiones activas conectadas al bot. |
| `/key` | Recupera la clave de producto de la instalación de Windows. |
| `/wifi` | Extrae las contraseñas de todos los perfiles WiFi almacenados. |

---

> [!WARNING]
> Esta herramienta debe utilizarse exclusivamente en entornos con autorización explícita y fines educativos.