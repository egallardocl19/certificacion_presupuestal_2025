// =============================================================
// Menú y utilidades de inicialización manual
// =============================================================

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Certificaciones')
      .addItem('Configurar sistema', 'configurarSistema')
      .addItem('Crear solo estructura', 'crearSoloEstructura')
      .addItem('Normalizar estructura', 'normalizarEstructuraSistema')
      .addSeparator()
      .addItem('Resetear sistema (borra datos)', 'resetearSistema')
      .addToUi();
  } catch (error) {
    Logger.log('No se pudo crear el menú personalizado: ' + error);
  }
}

function ejecutarConfiguracionInicial() {
  return configurarSistema();
}

function reconstruirEstructura() {
  return crearSoloEstructura();
}

function normalizarEstructura() {
  return normalizarEstructuraSistema();
}

function reiniciarSistema() {
  return resetearSistema();
}
