// Configuración de la app. Estos datos son públicos (no son contraseñas).

// Dirección del servidor de Apps Script (termina en /exec)
self.API_URL = 'https://script.google.com/macros/s/AKfycbxJQppvBK5XBpo5db3rSiuNUGc4lQ-WN22t7Lr56esg6ye_F0lK41waoLZ_Y-s8p6oODA/exec';

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
