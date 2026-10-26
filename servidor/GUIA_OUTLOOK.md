# Envío de correos desde la cuenta de Outlook de cada usuario

Los módulos «Enviar informes» y «Avisos de incidencias» de Seguimiento y Control envían los correos
desde la cuenta de Microsoft 365 de quien pulsa «Enviar», igual que hacían los programas de
escritorio con su Outlook. Para eso la plataforma usa Microsoft Graph, que es gratuito.

Hace falta registrar **una vez** la aplicación en Microsoft Entra (el antiguo Azure AD). Tarda unos
10 minutos y lo hace una persona con permisos de administrador de Microsoft 365 de GEYPE.

## 1. Registrar la aplicación

1. Entra en <https://entra.microsoft.com> con la cuenta de administrador.
2. Ve a **Aplicaciones → Registros de aplicaciones → Nuevo registro**.
3. Rellena:
   - **Nombre:** `Herramientas GE&PE - correo`.
   - **Tipos de cuenta admitidos:** *Solo las cuentas de este directorio organizativo (solo GEYPE,
     un solo inquilino)*.
   - **URI de redirección:** déjala vacía.
4. Pulsa **Registrar**.
5. En la página de la aplicación, copia estos dos datos:
   - **Id. de aplicación (cliente)**;
   - **Id. de directorio (inquilino)**.

## 2. Permitir el inicio de sesión con código

1. Entra en **Autenticación**.
2. Abajo, en **Configuración avanzada**, activa **Permitir flujos de clientes públicos** (Sí).
3. Pulsa **Guardar**.

La plataforma no usa contraseña ni secreto de la aplicación: cada usuario entra con su propia cuenta.

## 3. Permisos

1. Entra en **Permisos de API → Agregar un permiso → Microsoft Graph → Permisos delegados**.
2. Marca estos tres permisos:
   - `Mail.Send`;
   - `User.Read`;
   - `offline_access`.
3. Pulsa **Agregar permisos**.
4. Pulsa **Conceder consentimiento de administrador para GEYPE** y confirma. Así nadie del equipo
   tendrá que aceptar permisos al conectarse.

Son permisos *delegados*: la aplicación solo puede enviar como el usuario que se ha conectado, nunca
como otra persona. Tampoco puede leer el correo.

## 4. Guardar los datos en la plataforma

1. En la plataforma, entra en **Administración → Seguimiento y Control → Outlook**.
2. Pega el **Id. de directorio (inquilino)** y el **Id. de aplicación (cliente)**.
3. Pulsa **Guardar**.

## 5. Conectar cada cuenta (cada usuario, una vez)

1. Abre **Seguimiento y Control → Enviar informes** (o «Avisos de incidencias»).
2. Pulsa **Conectar mi Outlook**.
3. Abre el enlace `microsoft.com/devicelogin`, escribe el código que muestra la plataforma y entra con
   tu cuenta de GEYPE.
4. Vuelve a la plataforma y pulsa **Ya he entrado**.

Desde ese momento los correos salen de tu cuenta y quedan en tus *Elementos enviados*.

- **Si no se usa:** la conexión se renueva sola mientras se use; tras unos 90 días sin usarla, hay que
  volver a conectar.
- **Dónde se guarda:** el permiso, cifrado (DPAPI de Windows), en la base de datos del PC de la
  oficina. Nunca va en GitHub.
- **Cómo quitarlo:** con «Desconectar» en la misma página, o un administrador desde Entra
  (Aplicaciones empresariales → Herramientas GE&PE - correo → Usuarios).

## Antes del primer envío real

- **Modo prueba:** los envíos tienen el «Modo prueba» activado por defecto. No envía nada y genera el
  mismo Excel que los programas antiguos.
- **Primera prueba:** haz la primera prueba real con un contacto de envío que tenga **tu propio
  email** en «Para».
