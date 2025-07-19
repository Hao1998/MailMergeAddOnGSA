function convertNumberFormat222() {
  var number = "2000000"
  var parts = number.toString().split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, "'");
  console.log(parts.join('.'))
//  return parts.join('.');
}