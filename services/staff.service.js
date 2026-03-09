function isShiftCountValid(target, actual) {
  if (target === undefined || target === null || target === '') return false;

  const targetStr = target.toString().trim();
  const currentActual = Number(actual);

  if (targetStr.includes('-')) {
    const parts = targetStr.split('-').map(Number);
    return currentActual >= parts[0] && currentActual <= parts[1];
  }

  return Number(targetStr) === currentActual;
}
