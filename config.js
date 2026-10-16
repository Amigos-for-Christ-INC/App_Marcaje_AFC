// Configuración de la app. Estos datos son públicos (no son contraseñas).

// Dirección del servidor de Apps Script (termina en /exec)
self.API_URL = 'https://script.google.com/macros/s/AKfycbyWAZJYh5y4y4mfEGjy663wVpd33pAcpj_vSr4C27DDfmazERD9qwzID6JdWgxHd1Sz/exec';

// Firebase, para los recordatorios por notificación. Se copian de la consola de Firebase:
// Configuración del proyecto > General > Tus apps > Configuración del SDK.
// La clave PRIVADA de Firebase no va aquí: va solo en las propiedades del script de Apps Script.
self.FIREBASE_CONFIG = {
  apiKey: 'AIzaSyD0h1Ju5E8X35AL_umMEa3L8KN5Ty_LWe8',
  projectId: 'asistencia-afc',
  messagingSenderId: '221026166625',
  appId: '1:221026166625:web:0c53060e4cc94583b64720'
};

// Configuración del proyecto > Cloud Messaging > Certificados push web > Par de claves (clave pública)
self.VAPID_KEY = 'BDPTUHtdnW-E6Tc6YWKVAR7cD8uQfVkfxW4agXt5jxBo3NMatnBDcUEBy3fILWbQeIk4jAo4rwGfeIib7FPqrcw';

// Microsoft Entra ID (Entrar con Microsoft). Se copian del registro de la app en entra.microsoft.com:
// Id. de aplicación (cliente) e Id. de directorio (inquilino). No son contraseñas.
self.MS_CLIENT_ID = '4f9e40a5-a27d-4bca-90ef-571a2cfa1437';
self.MS_TENANT_ID = '3df804dd-19d4-414d-aea4-d693b7b73b45';
