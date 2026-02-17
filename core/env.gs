function getScriptProperty(key, fallback = '') {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return value === null ? fallback : value;
}

function setScriptProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, String(value));
}
