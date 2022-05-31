/// JavaScript helper functions

function isNumber(n) {
  // https://stackoverflow.com/a/1421988/631051
  return !isNaN(parseFloat(n)) && !isNaN(n - 0)
}

function ensureValuesUnique(values) {
  const valuesSet = new Set();  // requires V8 engine
  for (var i = 0; i < values.length; i++) {
    var value = values[i].toString().trim();
    if (valuesSet.has(value)) {
      throw ('Duplicate value: ' + value);
    }
    valuesSet.add(value);
  }
}
