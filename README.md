# excel_sendmail
VBA macro to send mails from excel using Outlook


## instalación
- crear una planilla con dos solapas (Alertas y Configuración) tomando de ejemplo los .csv
- importar el módulo de código

## Forma de usar
1- Crear un perfil de Outlook que tenga configurada la cuenta de mail para envío de mails 
   Por ejemplo, el nombre del perfil puede ser Alertas
2- Cerrar Outlook
3- Abrir SendMails v0.5, actualizar en la solapa de Configuración los datos de la cuenta de mail y el perfil
4- Completar los datos de los mails a enviar, y ejecutar la macro SendMail

Notas:
- Si Outlook está abierto, no cambiará de perfil. Si está abierto con el perfil correcto, funciona. Si no, no encontrará la cuenta de mail correspondiente.
- Si el perfil está mal escrito en la solapa de Configuración, abrirá un dialogo para seleccionar el perfil.
