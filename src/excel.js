// Variable para almacenar el resultado
var resultado;

// Evento para cuando se selecciona un archivo en el input
document.getElementById('input-excel').addEventListener('change', function (e) {
  var reader = new FileReader();

  reader.onload = function (e) {
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, { type: 'array' });

    // Obtén la primera hoja del libro de trabajo
    var firstSheetName = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[firstSheetName];

    // Convierte las filas de la hoja en objetos JSON
    var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Divide el array en tres partes: el primer elemento, el último elemento y el resto
    let primerArray = jsonData[0];
    let ultimoArray = jsonData[jsonData.length - 1];
    let restoDeArrays = jsonData.slice(1, jsonData.length - 1);

    // Ordena el array 'restoDeArrays' y lo almacena en 'restoDeArraysOrdenado'
    let restoDeArraysOrdenado = restoDeArrays.sort((a, b) => {
      // Compara el tercer elemento (índice 3) de cada subarray
      // Si el tercer elemento de 'a' es menor que el de 'b', retorna -1 (coloca 'a' antes que 'b')
      if (a[3] < b[3]) return -1;
      // Si el tercer elemento de 'a' es mayor que el de 'b', retorna 1 (coloca 'a' después de 'b')
      if (a[3] > b[3]) return 1;

      // Si los terceros elementos son iguales, compara el quinto elemento (índice 4)
      // Retorna la diferencia entre el quinto elemento de 'a' y 'b' (ordena de menor a mayor)
      return a[4] - b[4];
    });

    // Une los arrays en el orden correcto
    let arrayOrdenado = [primerArray, ...restoDeArraysOrdenado, ultimoArray];

    resultado = arrayOrdenado;
  };
  reader.readAsArrayBuffer(e.target.files[0]);
});

/**
 * Función auxiliar para convertir una cadena en un array de 8 bits
 * @param {*} s 
 * @returns 
 */
function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
};

// Selecciona el formulario
const form = document.querySelector('form');

// Agrega un evento de envío al formulario
form.addEventListener('submit', e => {
  // Previene la acción por defecto del formulario
  e.preventDefault();

  // Recoge los valores de los inputs
  const teatro = "Teatro " + document.querySelector('#input-teatro').value;
  const obra = document.querySelector('#input-obra').value;
  const fecha = document.querySelector('#input-fecha').value;
  const autor = "Autor: " + document.querySelector('#input-autor').value;
  const produccion = "Produccion: " + document.querySelector('#input-produccion').value;

  // Crea un nuevo libro de trabajo
  var workbook = XLSX.utils.book_new();

  // Crea una nueva hoja de trabajo con cada variable en una fila separada
  var worksheet = XLSX.utils.aoa_to_sheet([
    [teatro],
    [obra],
    [fecha],
    [autor],
    [produccion],
    [] // Agrega una fila vacía
  ]);

  // Inicializa worksheet['!merges'] con un array vacío
  worksheet['!merges'] = [];

  // Unificar las celdas de cada fila
  for (let i = 0; i < 6; i++) { // Nota: ahora el bucle va hasta 6 para incluir la fila vacía
    worksheet['!merges'].push({ s: { r: i, c: 0 }, e: { r: i, c: 4 } }); // Unificar las celdas de la fila i
  }

  // Agrega una fila vacía al final de la hoja de trabajo
  XLSX.utils.sheet_add_aoa(worksheet, [[]], { origin: -1 });

  // Agrega cada fila de 'resultado' a la hoja de trabajo
  for (let index = 0; index < (resultado.length - 1); index++) {
    let fila = resultado[index];
    if (index == 0) {
      XLSX.utils.sheet_add_aoa(worksheet, [[fila[2], fila[3], fila[4], fila[5], fila[6], 'Importe', 'Valor con importe']], { origin: -1 });
    } else {
      XLSX.utils.sheet_add_aoa(worksheet, [[fila[2], fila[3], fila[4], fila[5], fila[6]]], { origin: -1 });
    }
  }

  // Agrega fórmulas a las celdas
  for (let index = 1; index < resultado.length - 1; index++) {
    var cell = ('G' + (index + 7));
    var formula = ('D' + (index + 7) + '*F' + (index + 7));
    worksheet[cell] = { f: formula }
    if ((index + 1) == (resultado.length - 1)) {
      var celda = ('G' + (index + 8));
      console.log(celda);

      var formulita = ('SUM(G8' + ':G' + (index + 7) + ')');
      worksheet[celda] = { f: formulita }
    }
  }

  // Inicializa un objeto para almacenar las sumas
  let sumas = {};

  // Recorremos cada array interno, excepto el primero y el último
  for (let i = 1; i < resultado.length - 1; i++) {
    let arrayInterno = resultado[i];
    let tercerString = arrayInterno[3];

    // Si el tercer string no está en el objeto sumas, lo añadimos
    if (!sumas[tercerString]) {
      sumas[tercerString] = Array(arrayInterno.length).fill(0);
    }

    // Recorremos cada elemento del array interno
    for (let j = 0; j < arrayInterno.length; j++) {
      // Si el elemento es un número, lo sumamos al elemento correspondiente en el array de sumas
      if (!isNaN(arrayInterno[j])) {
        sumas[tercerString][j] += arrayInterno[j];
      } else {
        // Si el elemento es un string, lo mantenemos
        sumas[tercerString][j] = arrayInterno[j];
      }
    }
  }

  // Convertimos el objeto sumas en un array de arrays
  let arrayDeSumas = Object.values(sumas);

  // Agrega dos filas vacías al final de la hoja de trabajo
  XLSX.utils.sheet_add_aoa(worksheet, [[]], { origin: -1 });
  XLSX.utils.sheet_add_aoa(worksheet, [[]], { origin: -1 });

  // Agrega cada fila de 'arrayDeSumas' a la hoja de trabajo
  for (let index = 0; index < arrayDeSumas.length; index++) {
    let arrayDeSumas2 = arrayDeSumas[index];
    XLSX.utils.sheet_add_aoa(worksheet, [['', arrayDeSumas2[3], arrayDeSumas2[4], arrayDeSumas2[5], '$' + arrayDeSumas2[6]]], { origin: -1 });
  }

  // Inicializa las variables para almacenar los totales
  let precioFinal = null;
  let cantidadFinal = null;
  let totalVentaFinal = null;

  // Calcula los totales
  for (let index = 0; index < arrayDeSumas.length; index++) {
    let arrayInterno = arrayDeSumas[index];

    precioFinal += arrayInterno[4];
    cantidadFinal += arrayInterno[5];
    totalVentaFinal += arrayInterno[6];

    // Si es la última iteración, agrega los totales a la hoja de trabajo
    if (index == (arrayDeSumas.length - 1)) {
      XLSX.utils.sheet_add_aoa(worksheet, [['', 'Totales', '', cantidadFinal, '$' + totalVentaFinal]], { origin: -1 });
    }
  }

  // Agrega la hoja de trabajo al libro de trabajo
  XLSX.utils.book_append_sheet(workbook, worksheet, "Hoja1");

  // Escribe el libro de trabajo en un archivo Excel
  var wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });

  // Crea un Blob con los datos del archivo Excel
  var blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });

  // Usa FileSaver.js para hacer que el archivo sea descargable
  saveAs(blob, 'Reporte ' + obra + ' ' + fecha.replace(/-/g, "") + '.xlsx');
});