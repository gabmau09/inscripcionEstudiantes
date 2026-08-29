// ================= CONFIGURACIÓN GENERAL =================
const CONFIG = {
  // ID de la carpeta principal donde se crearán las subcarpetas
  ID_CARPETA_RAIZ: "id_carpeta_drive",

  // ID del Excel que contiene la pestaña "BD CURSOS"
  ID_EXCEL_BD: "id_excel_principal",
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Solicitud de Estudiantes - Instituto de Informática')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// LECTURA DINÁMICA ABSOLUTA DESDE LA HOJA "BD CURSOS"
function getInitialData() {
  let ss = SpreadsheetApp.openById(CONFIG.ID_EXCEL_BD);
  let sheet = ss.getSheetByName("BD CURSOS");
  if (!sheet) throw new Error("No se encontró la hoja 'BD CURSOS'");
  
  let data = sheet.getDataRange().getValues();
  let db = {};
  
  // Recorre desde la fila 2 (índice 1, saltando cabeceras)
  for (let i = 1; i < data.length; i++) {
    let curso = data[i][0] ? data[i][0].toString().trim().toUpperCase() : "";       // Col A: CURSOS
    let modalidad = data[i][1] ? data[i][1].toString().trim().toUpperCase() : "";   // Col B: MODALIDAD
    let limite = data[i][2] ? parseInt(data[i][2]) : 20;                            // Col C: LIMITE
    let idSheet = data[i][3] ? data[i][3].toString().trim() : "";                   // Col D: ID CURSOS
    let enlace = data[i][4] ? data[i][4].toString().trim() : "";                    // Col E: ENLACE SOLICITUD
    
    if (curso && modalidad) {
      if (!db[modalidad]) db[modalidad] = {};
      db[modalidad][curso] = {
        limite: limite,
        id: idSheet,
        enlace: enlace
      };
    }
  }
  
  return db;
}

// OBTENER VACANTES EN TIEMPO REAL DESDE EL LIBRO ASIGNADO
function getCupoDisponible(modalidad, curso) {
  if (!modalidad || !curso) return 0;
  let db = getInitialData();
  
  if (!db[modalidad] || !db[modalidad][curso]) return 0;
  
  let limite = db[modalidad][curso].limite || 20;
  let sheetId = db[modalidad][curso].id;
  
  if (!sheetId) return limite; 
  
  try {
    let ss = SpreadsheetApp.openById(sheetId);
    let numGrupo = 1;
    while (ss.getSheetByName("Grupo " + (numGrupo + 1))) {
      numGrupo++;
    }
    
    let hoja = ss.getSheetByName("Grupo " + numGrupo);
    if (hoja) {
      let inscritos = contarAlumnos(hoja, 15, 3); // Cuenta DNI (Columna C=3)
      let disponibles = limite - inscritos;
      return (disponibles <= 0) ? limite : disponibles;
    }
  } catch(e) {
    return limite;
  }
  return limite;
}

function procesarFormulario(formulario) {
  try {
    let db = getInitialData();
    let modalidad = formulario.asunto ? formulario.asunto.toString().trim().toUpperCase() : "";
    let curso = formulario.curso ? formulario.curso.toString().trim().toUpperCase() : "";
    
    let cursoData = db[modalidad] && db[modalidad][curso] ? db[modalidad][curso] : null;
    if (!cursoData) throw new Error("No se encontró configuración en BD CURSOS para: " + modalidad + " - " + curso);
    
    let sheetId = cursoData.id;
    if (!sheetId) throw new Error("Falta el ID del Excel en la Columna D de BD CURSOS para: " + curso);

    let nombreCarpetaAlumno = `${formulario.nombres} ${formulario.apellidos} - ${formulario.dni}`;
    let carpetaRaiz = DriveApp.getFolderById(CONFIG.ID_CARPETA_RAIZ);
    let carpetaDestinoAlumno;
    let hojaDestino;
    let spreadsheet = SpreadsheetApp.openById(sheetId);
    
    let filaInicioDatos;
    let isCurso = (modalidad === "CURSO");

    // ================= 1. RUTEO, HOJAS Y CARPETAS =================
    if (isCurso) {
      let limiteCurso = cursoData.limite;
      let numGrupo = determinarGrupoYDuplicarPlantilla(spreadsheet, limiteCurso);
      
      let carpetaCursos = obtenerCrearCarpeta(carpetaRaiz, "CURSOS");
      let carpetaCursoEspec = obtenerCrearCarpeta(carpetaCursos, curso);
      let carpetaGrupo = obtenerCrearCarpeta(carpetaCursoEspec, "Grupo " + numGrupo);
      carpetaDestinoAlumno = obtenerCrearCarpeta(carpetaGrupo, nombreCarpetaAlumno);
      
      hojaDestino = spreadsheet.getSheetByName("Grupo " + numGrupo);
      filaInicioDatos = 15;

    } else if (modalidad === "EXAMEN DE SUFICIENCIA") {
      hojaDestino = spreadsheet.getSheetByName("Examen - " + curso);
      if (!hojaDestino) throw new Error("No se encontró la hoja 'Examen - " + curso + "' en el archivo de este curso.");
      
      let carpetaExamenes = obtenerCrearCarpeta(carpetaRaiz, "EXAMENES DE SUFICIENCIA");
      let carpetaCursoEspec = obtenerCrearCarpeta(carpetaExamenes, curso);
      carpetaDestinoAlumno = obtenerCrearCarpeta(carpetaCursoEspec, nombreCarpetaAlumno);
      
      filaInicioDatos = 2; 
    } else {
      throw new Error("Modalidad no reconocida: " + modalidad);
    }

    // ================= 2. SUBIR PDFs AL DRIVE Y OBTENER ENLACES =================
    let urls = guardarArchivosEnDrive(formulario, carpetaDestinoAlumno);

    // ================= 3. PREPARAR DATOS Y FECHAS =================
    let hoy = new Date();
    let fechaSolicitud = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    let fCurso = formatoFecha(formulario.fechaPagoCurso);
    let fConstancia = formatoFecha(formulario.fechaPagoConstancia);

    let opCurso = formulario.numOperacionCurso || "";
    let opConstancia = formulario.numOperacionConstancia || "";
    let numOp = opCurso + "; " + opConstancia;
    let fechaOp = (fCurso === fConstancia) ? fCurso : (fCurso + "; " + fConstancia);

    // ================= 4. LÓGICA DEL CÓDIGO EXTERNO EXACTO A 10 CARACTERES =================
    let esUnp = (formulario.es_unp === "SI");
    let codUniv = "EXTERNO";
    let correoInst = "EXTERNO";
    let escuelaProf = "EXTERNO";

    if (esUnp) {
      codUniv = formulario.codigo_universitario || "EXTERNO";
      correoInst = formulario.correo_institucional || "EXTERNO";
      escuelaProf = formulario.escuela || "EXTERNO";
    } else {
      let ssBD = SpreadsheetApp.openById(CONFIG.ID_EXCEL_BD);
      let sheetBD = ssBD.getSheetByName("BD CURSOS");
      if (!sheetBD) throw new Error("No se encontró la hoja 'BD CURSOS'");
      
      let valG1 = sheetBD.getRange("G1").getValue();
      let numCorrelativoExt = parseInt(valG1) || 1;
      
      // Construye EXT + 7 dígitos con ceros = 10 caracteres
      codUniv = "EXT" + String(numCorrelativoExt).padStart(7, '0');
      
      // Incrementa G1
      sheetBD.getRange("G1").setValue(numCorrelativoExt + 1);
    }

    let alumnosActuales = contarAlumnos(hojaDestino, filaInicioDatos, 3);
    let filaAInsertar = filaInicioDatos + alumnosActuales;
    let numeroCorrelativo = alumnosActuales + 1;

    // Arreglo completo con columnas A a T
    let datosFila = [
      numeroCorrelativo,                                  // A: N°
      codUniv,                                            // B: COD. UNIV. 
      formulario.dni,                                     // C: DNI
      `${formulario.apellidos} ${formulario.nombres}`.toUpperCase(), // D: APELLIDOS Y NOMBRES
      correoInst,                                         // E: CORREO INSTITUCIONAL
      formulario.correo_personal,                         // F: CORREO PERSONAL
      numOp,                                              // G: N° DE OPERACIÓN 
      "",                                                 // H: MONTO S/. 
      fechaOp,                                            // I: FECHA DEL VOUCHER 
      "",                                                 // J: ESTADO
      "",                                                 // K: CERTIFICADO
      "",                                                 // L: CONSTANCIA
      "",                                                 // M: FECHA DE ENVIO
      formulario.celular,                                 // N: CELULAR
      escuelaProf,                                        // O: ESCUELA PROFESIONAL
      fechaSolicitud,                                     // P: FECHA DE SOLICITUD
      urls.solicitud,                                     // Q: LINK SOLICITUD
      urls.dni,                                           // R: LINK DNI
      urls.pagoCurso,                                     // S: LINK PAGO CURSO
      urls.pagoConstancia                                 // T: LINK PAGO CONSTANCIA
    ];
    
    hojaDestino.getRange(filaAInsertar, 1, 1, datosFila.length).setValues([datosFila]);

    return { success: true, message: "Envío exitoso" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ======================= FUNCIONES AUXILIARES =======================

function formatoFecha(fechaOriginal) {
  if (!fechaOriginal) return "";
  let partes = fechaOriginal.split("-");
  if (partes.length === 3) {
    return partes[2] + "/" + partes[1] + "/" + partes[0];
  }
  return fechaOriginal; 
}

function determinarGrupoYDuplicarPlantilla(ss, limite) {
  let numGrupo = 1;
  let hoja = ss.getSheetByName("Grupo " + numGrupo);
  
  while (ss.getSheetByName("Grupo " + (numGrupo + 1))) {
    numGrupo++;
    hoja = ss.getSheetByName("Grupo " + numGrupo);
  }

  if (!hoja) {
    let plantilla = ss.getSheetByName("plantilla"); 
    if (!plantilla) throw new Error("No existe la hoja 'plantilla' en este documento.");
    hoja = plantilla.copyTo(ss).setName("Grupo " + numGrupo);
  }

  let alumnosInscritos = contarAlumnos(hoja, 15, 3);
  
  if (alumnosInscritos >= limite) {
    numGrupo++;
    let plantilla = ss.getSheetByName("plantilla");
    if (!plantilla) throw new Error("No existe la hoja 'plantilla' en este documento.");
    hoja = plantilla.copyTo(ss).setName("Grupo " + numGrupo);
  }

  return numGrupo;
}

function contarAlumnos(hoja, filaInicio, columnaVerificar) {
  let maxFilas = hoja.getMaxRows();
  if (filaInicio > maxFilas) return 0;
  
  let datos = hoja.getRange(filaInicio, columnaVerificar, maxFilas - filaInicio + 1, 1).getValues();
  let cuenta = 0;
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0] !== "" && datos[i][0] !== null) cuenta++;
    else break; 
  }
  return cuenta;
}

function obtenerCrearCarpeta(carpetaPadre, nombreCarpeta) {
  let carpetas = carpetaPadre.getFoldersByName(nombreCarpeta);
  if (carpetas.hasNext()) return carpetas.next();
  else return carpetaPadre.createFolder(nombreCarpeta);
}

function guardarArchivosEnDrive(form, carpeta) {
  let urls = { solicitud: "N/A", dni: "N/A", pagoCurso: "N/A", pagoConstancia: "N/A" };
  if (form.docSolicitud) urls.solicitud = carpeta.createFile(form.docSolicitud).getUrl();
  if (form.docDni) urls.dni = carpeta.createFile(form.docDni).getUrl();
  if (form.docPagoCurso) urls.pagoCurso = carpeta.createFile(form.docPagoCurso).getUrl();
  if (form.docPagoConstancia) urls.pagoConstancia = carpeta.createFile(form.docPagoConstancia).getUrl();
  return urls;
}
