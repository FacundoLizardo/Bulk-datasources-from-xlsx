// Importar módulos necesarios
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const FormData = require('form-data');
const XLSX = require('xlsx');

// Función para esperar una cantidad de milisegundos
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Ruta de la carpeta donde se encuentra el archivo XLSX
const folderPath = path.join(__dirname, 'file');

// Leer el contenido de la carpeta y filtrar el archivo .xlsx
const files = fs.readdirSync(folderPath);
const xlsxFiles = files.filter(file => path.extname(file).toLowerCase() === '.xlsx');

if (xlsxFiles.length !== 1) {
  console.error('Se espera exactamente un archivo .xlsx en la carpeta "file".');
  process.exit(1);
}

const filePath = path.join(folderPath, xlsxFiles[0]);

// Leer el libro de Excel y la primera hoja
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convertir la hoja a un arreglo de arreglos (cada fila es un array)
// Cada fila tiene: [header, valorObjeto1, valorObjeto2, ...]
const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Determinar la cantidad máxima de objetos (columnas a partir de la B)
let maxObjetos = 0;
rows.forEach(row => {
  if (row.length - 1 > maxObjetos) {
    maxObjetos = row.length - 1;
  }
});

// Función asíncrona para procesar cada objeto y enviarlo al endpoint
async function processObjects() {
  // Datos del endpoint y token
  const url = "https://dashboard.chatsappai.com/api/datasources";
  const token = "1a45061a-06a4-4c0f-b190-f6fca7903216"; // Reemplaza por tu token real
  const datastoreId = "cm8p7a4bl00msnsbhljmdat5a";

  // Iterar sobre cada columna (objeto) a partir de la columna B
  for (let col = 1; col <= maxObjetos; col++) {
    let partes = [];
    let tecnologiaValue = null; // Para almacenar el valor de la clave "tecnologia"

    // Recorrer cada fila para formar el objeto
    rows.forEach(row => {
      if (row && row.length > 0) {
        const key = row[0];       // Header en la columna A
        const value = row[col];   // Valor correspondiente en la columna actual

        // Si la clave es "tecnologia" (ignorando mayúsculas/minúsculas) se guarda su valor
        if (typeof key === 'string' && key.trim().toLowerCase() === 'tecnología') {
          if (value !== undefined && value !== null && value !== '') {
            tecnologiaValue = value;
          }
        }

        // Agregar al array el par "clave: valor" o solo la clave si no hay valor
        if (value !== undefined && value !== null && value !== '') {
          partes.push(`${key}: ${value}`);
        } else {
          partes.push(`${key}`);
        }
      }
    });

    // Construir el string del objeto a enviar
    let objectData = partes.join(', ');

    // Eliminar la palabra "Objeto" en caso de que aparezca
    objectData = objectData.replace(/Objeto\s*/gi, '');

    // Usar el valor de "tecnologia" para el nombre del archivo, o un nombre por defecto
    let fileName;
    if (tecnologiaValue) {
      fileName = `${tecnologiaValue}`;
    }
    const custom_id = `${fileName} + link`;

    // Armar el payload a enviar (simulando el envío de un archivo)
    const payload = {
      fileName: fileName,
      datastoreId: datastoreId,
      custom_id: custom_id,
    };

    // Crear el formulario y agregar los campos del payload
    const form = new FormData();
    for (const key in payload) {
      form.append(key, payload[key]);
    }

    // Agregar el contenido del objeto como "archivo" (usando Buffer)
    form.append('file', Buffer.from(objectData, 'utf-8'), {
      filename: fileName,
      contentType: 'text/csv'
    });

    // Realizar la petición POST al endpoint
    try {
      const response = await axios.post(url, form, {
        headers: {
          'Authorization': `Bearer ${token}`,
          ...form.getHeaders()
        }
      });
      console.log(`Objeto ${col} enviado. Código: ${response.status}`);
      console.log("Respuesta:", response.data);
    } catch (error) {
      console.error(`Error enviando el objeto ${col}:`, error.message);
    }

    // Esperar 3 segundos antes de procesar el siguiente objeto
    await sleep(3000);
  }
}

// Ejecutar el procesamiento y envío de objetos
processObjects();
