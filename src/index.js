const xlsxFile = require('read-excel-file/node')
const fs = require('fs')

//data:text/csv;charset=utf-8,
let csvFile = ''

xlsxFile('../EMPRESAS.xlsx').then((rows) => {
  /*
        0-cod_registro
        1-ciudad	
        2-matricula
        3-socios	
        4-empresa	
        5-primer_nombre	
        6-segundo_nombre	
        7-primer_apellido	
        8-segundo_apellido	
        9-finalidad	
        10-denominacion	
        11-tomo	
        12-fecha	
        13-PDF
    */
  //init
  let idCount = 0

  let fileRow = [
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
  ].join(',')

  csvFile += fileRow + '\r\n'

  delete rows[0];
  //Modificar
  rows.forEach((row) => {
    let newFileRow = []
    let depa = ''
    let ciudad = ''
    let camara = ''
    let tipoEmpresa = ''
    idCount++

    //Modifi Dir
    if(row[4]!=='No hay información' && row[4]!=='No hay empresa'){
    if (row[1] === 'TGU') {
      depa = 'Francisco Morazan'
      ciudad = 'Tegucigalpa'
      camara = 'Cámara de Comercio e Industrias de Tegucigalpa'
    } else if (row[1] === 'SPS') {
      depa = 'San Pedro Sula'
      ciudad = 'Cortes'
      camara = 'Cámara de Comercio e Industrias de Puerto Cortés'
    }
    //Modifi Razon Social (SA | SA)$
    if (typeof row[4] !== 'object' && typeof row[4] !== 'number' ) {
      if (
        row[4].includes(' SA ') ||
        row[4].includes(' SOCIEDAD ANONIMA ') ||
        row[4].includes(' SOCIEDAD ANONIMA') ||
        row[4].includes('SOCIEDAD ANONIMA ') ||
        row[4].includes(' S.A.') ||
        row[4].includes(' S A ') ||
        row[4].includes('S A ')
      ) {
        tipoEmpresa = 'SA'
      } else if (
        row[4].match('RESPONSABILIDAD LIMITADA') ||
        row[4].match('S de RL') ||
        row[4].match('S de R L') ||
        row[4].match('RL') ||
        row[4].match('R L')
      ) {
        tipoEmpresa = 'SDRL'
      } else {
        tipoEmpresa = 'CI'
      }
    }
    //Asign
    newFileRow[0] = row[2]; //Matricula
    newFileRow[1] = tipoEmpresa; //T Empresa
    //console.log(typeof(row[4]))
    typeof(row[4])=='object' || typeof(row[4])=='number' ? newFileRow[2] = row[4] : newFileRow[2] = row[4].replace(',',' '); //R Social
    typeof(row[4])=='object' || typeof(row[4])=='number' ? newFileRow[3] = row[4] :  newFileRow[3] = row[4].replace(',',' ');//Nombre Empresa

    newFileRow[4] = depa //Departamento
    newFileRow[5] = ciudad //Ciudad
    newFileRow[6] = camara //Camara de comercio
    //console.log(typeof(row[7]))
    typeof(row[7])=='object' || typeof(row[7])=='number' ? newFileRow[7] = row[7] : newFileRow[7] = row[7].replace(',','')//Actividad

    newFileRow[8] = '0000-0000' //Telefono
    newFileRow[9] = 'No especificada' //Direccion
    newFileRow[10] = 'No especificado' //SitioWeb
    newFileRow[11] = 'No especificado' //Correo

    //console.log(typeof(row[3]))
    typeof(row[3])=='object'|| typeof(row[3])=='number' ? newFileRow[12] = row[3] : newFileRow[12] = row[3].replace(',','')//Socio

    newFileRow[13] = 'identidad ' + idCount //ID
    newFileRow[14] = 'Hondureña' //Nacionalidad
    newFileRow[15] = 1 //Aporte

    //Covertir en CSV
    
    let fileRow = newFileRow.join(',')
    csvFile += fileRow + '\r\n'
    }
  })
    //Exportar
    // Write data in 'newfile.txt' .
  console.log('Espera un momento');
  fs.writeFile('empresasFormated.csv', csvFile, (error) => {
    if (error) {
      throw err
    } else {
      console.log('Saved')
    }
  })
  ////console.log(rows[0])
  ////console.log(rows[1])
  ////console.log(csvFile)
})
