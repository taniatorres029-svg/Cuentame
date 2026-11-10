# Cuéntame cómo estás

Página sencilla para que los pacientes de la Dra. Tania Torres (Oncología médica) registren sus síntomas y se los envíen por WhatsApp.

Enlace: https://taniatorres029-svg.github.io/Cuentame/

Los datos del paciente se guardan solo en su celular.

## Base privada y panel

- `apps-script/Code.gs` va en una Hoja de cálculo de Google (Extensiones → Apps Script) e implementa `doPost` (guarda cada cuestionario) y `doGet` (devuelve las respuestas solo con la clave).
- `index.html` y `panel/index.html` tienen `const BASE_URL` con la dirección de esa implementación.
- Panel de la doctora: https://taniatorres029-svg.github.io/Cuentame/panel/
