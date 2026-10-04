// Exécuter une fois dans la console GAS
function clearHATokenCache() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('HA_TOKEN');
  props.deleteProperty('HA_TOKEN_EXPIRY');
  Logger.log('Cache token HelloAsso vidé');
}