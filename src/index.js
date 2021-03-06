const xlsxFile = require('read-excel-file/node')
const fs = require('fs')
//CSV
const createCsvWriter = require('csv-writer').createArrayCsvWriter

const csvWriter = createCsvWriter({
  header: [
    'Matricula',
    'Tipo de Empresa',
    'Razon Social',
    'Nombre Comercial',
    'Departamento',
    'Ciudad',
    'Camara de Comercio',
    'Actividad',
    'Telefono',
    'Direccion',
    'Pagina Web',
    'Correo Electronico',
    'Nombre del Socio',
    'Identidad',
    'Nacionalidad',
    'Aporte',
  ],
  path: 'docs/empresas.csv',
  encoding: 'utf8',
})

let csvFile = ''
let records = []
xlsxFile('../EMPRESAS.xlsx').then((rows) => {
  //init
  let idCount = 0

  delete rows[0]
  //Modificar
  rows.forEach((row) => {
    let newFileRow = []
    let depa = ''
    let ciudad = ''
    let camara = ''
    let tipoEmpresa = ''
    let code = ''
    idCount++

    //Modifi Dir
    if (
      row[4] !== 'No hay información' &&
      row[4] !== 'No hay empresa' &&
      row[2] !== '' &&
      row[4] !== '' &&
      row[3] !== 'No hay socio' &&
      row[3] !== ''
    ) {
      if (row[1] === 'TGU') {
        depa = 'Francisco Morazan'
        ciudad = 'Tegucigalpa'
        camara = 'Cámara de Comercio e Industrias de Tegucigalpa'
        code = 'TGU'
      } else if (row[1] === 'SPS') {
        depa = 'San Pedro Sula'
        ciudad = 'Cortes'
        camara = 'Cámara de Comercio e Industrias de Puerto Cortés'
        code = 'SPS'
      }
      //Modifi Razon Social
      if (typeof row[4] !== 'object' && typeof row[4] !== 'number') {
        if (
          row[4].includes(' SA ') ||
          row[4].includes(' SOCIEDAD ANONIMA ') ||
          row[4].includes(' SOCIEDAD ANONIMA') ||
          row[4].includes('SOCIEDAD ANONIMA ') ||
          row[4].includes(' S.A.') ||
          row[4].includes(' S A ') ||
          row[4].includes('S A ') ||
          row[4].match(/s\.\ ?a\./i)  ||
          row[4].match(/sociedad.?an.nima/i)  ||
          row[4].match(/(sa)$/i)
        ) {
          tipoEmpresa = 'SA'
        } else if (
          row[4].includes('RESPONSABILIDAD LIMITADA') ||
          row[4].includes('S de RL') ||
          row[4].includes('S de R L') ||
          row[4].includes('RL') ||
          row[4].includes('R L') ||
          row[4].match(/r\ ?l/i) ||
          row[4].match(/S\ de\ R\ ?L/i) ||
          row[4].match(/RESPONSABILDAD.?LIMITADA/i)
        ) {
          tipoEmpresa = 'SDRL'
        } else {
          tipoEmpresa = 'CI'
        }
      }
      //Asign
      newFileRow[0] = code + '-' + row[2] //Matricula
      newFileRow[1] = tipoEmpresa //T Empresa
      //console.log(typeof(row[4]))
      typeof row[4] == 'object' || typeof row[4] == 'number'
        ? (newFileRow[2] = row[4])
        : (newFileRow[2] = row[4].replace(',', ' ')) //R Social
      typeof row[4] == 'object' || typeof row[4] == 'number'
        ? (newFileRow[3] = row[4])
        : (newFileRow[3] = row[4].replace(',', ' ')) //Nombre Empresa

      newFileRow[4] = depa //Departamento
      newFileRow[5] = ciudad //Ciudad
      newFileRow[6] = camara //Camara de comercio
      //console.log(typeof(row[7]))

      typeof row[9] == 'object' || typeof row[9] == 'number' //Actividad
        ? (newFileRow[7] = row[9])
        : (newFileRow[7] = row[9].replace(',', ''))

      newFileRow[8] = '' //Telefono
      newFileRow[9] = '' //Direccion
      newFileRow[10] = '' //SitioWeb
      newFileRow[11] = '' //Correo

      //console.log(typeof(row[3]))
      if (typeof row[3] == 'object' || typeof row[3] == 'number') { //Socio
        newFileRow[12] = row[3]
      } else {
          newFileRow[12] = row[3].replace(',', '');
      }

      newFileRow[13] = 'Temporal ' + idCount //ID
      newFileRow[14] = 'Honduras' //Nacionalidad
      newFileRow[15] = 1 //Aporte

      records.push(newFileRow)
    }
  })
  //Exportar
  //Convertir en CSV
  console.log('Espera un momento')

  csvWriter.writeRecords(records).then(() => {
    console.log('...Done')
  })
})
