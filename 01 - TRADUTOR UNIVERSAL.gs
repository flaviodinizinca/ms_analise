// TRADUTOR UNIVERSAL DE NÚMEROS
function parseBRNumber(valor) {
  if (typeof valor === 'number') return valor;
  if (valor === "" || valor === null || valor === undefined) return 0;
  var str = String(valor).trim();
  if (str.indexOf(',') !== -1) {
    str = str.replace(/\./g, ''); 
    str = str.replace(',', '.');
  }
  
  var num = Number(str);
  return isNaN(num) ? 0 : num;
}