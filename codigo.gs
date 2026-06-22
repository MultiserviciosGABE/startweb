var spreadsheetName = "Consultas_Web";
var sheetName = "Respuestas";

var adminEmail = "edainballesteros8@gmail.com";


function getSpreadsheet(){

  var files = DriveApp.getFilesByName(spreadsheetName);

  if(files.hasNext()){

    var ss = SpreadsheetApp.open(files.next());
    var sheet = ss.getSheetByName(sheetName);

    if(!sheet){
      sheet = ss.insertSheet(sheetName);
      createHeaders(sheet);
    }

    return ss;

  }else{

    var ss = SpreadsheetApp.create(spreadsheetName);
    var sheet = ss.getActiveSheet();
    sheet.setName(sheetName);

    createHeaders(sheet);

    return ss;
  }
}


function createHeaders(sheet){

  sheet.appendRow([
    "ID",
    "Fecha",
    "Nombre",
    "Correo",
    "Telefono",
    "Consulta",
    "Descripcion",
    "Estado",
    "FechaCierre"
  ]);

}


/* =========================
   HOJA DE ORDENES
========================= */

function getSheetOrdenes(){

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Ordenes");

  if(!sheet){

    sheet = ss.insertSheet("Ordenes");

    sheet.appendRow([
      "Orden",
      "Fecha",
      "Cliente",
      "Telefono",
      "Vehiculo",
      "Placa",
      "Trabajo",
      "Total",
      "Estado",
      "FechaTerminada"
    ]);

  }

  return sheet;

}


/* =========================
   SIGUIENTE ORDEN
========================= */

function siguienteOrden(){

  var sheet = getSheetOrdenes();
  var lastRow = sheet.getLastRow() + 1;
  var numero = "OT-" + ("0000" + lastRow).slice(-4);

  return ContentService
  .createTextOutput(JSON.stringify({orden:numero}))
  .setMimeType(ContentService.MimeType.JSON);

}


/* =========================
   GUARDAR ORDEN
========================= */

function guardarOrden(e){

  var sheet = getSheetOrdenes();
  var lastRow = sheet.getLastRow() + 1;

  var numero = "OT-" + ("0000" + lastRow).slice(-4);

  var fecha = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );

  sheet.appendRow([
    numero,
    fecha,
    e.parameter.cliente || "",
    e.parameter.telefono || "",
    e.parameter.vehiculo || "",
    e.parameter.placa || "",
    e.parameter.trabajo || "",
    e.parameter.total || "",
    "Nueva",
    ""
  ]);

  return ContentService
  .createTextOutput(JSON.stringify({result:"ok",orden:numero}))
  .setMimeType(ContentService.MimeType.JSON);

}


/* =========================
   OBTENER ORDENES
========================= */

function obtenerOrdenes(){

  var sheet = getSheetOrdenes();
  var data = sheet.getDataRange().getValues();
  var ordenes = [];

  for(var i=1;i<data.length;i++){

    ordenes.push({

      orden:data[i][0],
      fecha:data[i][1],
      cliente:data[i][2],
      telefono:data[i][3],
      vehiculo:data[i][4],
      placa:data[i][5],
      trabajo:data[i][6],
      total:data[i][7],
      estado:data[i][8],
      fechaTerminada:data[i][9]

    });

  }

  return ordenes;

}



/* =========================
   POST
========================= */

function doPost(e){

  if(!e){
    return ContentService.createTextOutput("OK");
  }

  /* LOGIN */

  if(e.parameter.tipo == "login"){

    var ss = getSpreadsheet();
    var sheetUsers = ss.getSheetByName("Usuarios");

    var data = sheetUsers.getDataRange().getValues();

    for(var i=1;i<data.length;i++){

      if(data[i][0] == e.parameter.usuario && data[i][1] == e.parameter.password){

        return ContentService
        .createTextOutput(JSON.stringify({
          status:"ok",
          nombre:data[i][2],
          rol:data[i][3]
        }))
        .setMimeType(ContentService.MimeType.JSON);

      }

    }

    return ContentService
    .createTextOutput(JSON.stringify({status:"error"}))
    .setMimeType(ContentService.MimeType.JSON);

  }


  /* CONSULTAS PENDIENTES */

  if(e.parameter.tipo == "consultasPendientes"){

    var ss = getSpreadsheet();

    var sheet = ss.getSheetByName(sheetName);
    var data = sheet.getDataRange().getValues();

    var consultas = [];

    for(var i=1;i<data.length;i++){

      if(data[i][7] == "Pendiente"){

        consultas.push({
          id:data[i][0],
          fecha:data[i][1],
          nombre:data[i][2],
          correo:data[i][3],
          telefono:data[i][4],
          consulta:data[i][5],
          descripcion:data[i][6],
          estado:data[i][7]
        });

      }

    }

    return ContentService
    .createTextOutput(JSON.stringify(consultas))
    .setMimeType(ContentService.MimeType.JSON);

  }


  /* GUARDAR ORDEN */

  if(e.parameter.tipo == "guardarOrden"){
    return guardarOrden(e);
  }

  /* AGREGAR PRODUCTO INVENTARIO */

  if(e.parameter.tipo == "agregarProducto"){
    return agregarProducto(e);
  }

  if(e.parameter.tipo == "modificarProducto"){
    return modificarProducto(e);
  }

  
  /* MOVIMIENTO INVENTARIO (AGREGADO) */
  if(e.parameter.tipo == "movimientoInventario"){
    return registrarMovimiento(e);
  }
if(e.parameter.tipo == "eliminarProducto"){

  var r = eliminarProducto(e.parameter.id);

  return ContentService
   .createTextOutput(JSON.stringify(r))
   .setMimeType(ContentService.MimeType.JSON);
  }

  /* =========================
     REGISTRO CONSULTA WEB
  ========================= */

  if(!e.parameter.tipo || e.parameter.tipo == "consulta"){

    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    var lastRow = sheet.getLastRow() + 1;

    var id = "C" + ("0000" + (lastRow - 1)).slice(-4);

    var fechaHora = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );

    var nombre = e.parameter.nombre || "";
    var correo = e.parameter.correo || "";
    var telefono = e.parameter.telefono || "";
    var consulta = e.parameter.consulta || "";
    var descripcion = e.parameter.descripcion || "";

    sheet.appendRow([
      id,
      fechaHora,
      nombre,
      correo,
      telefono,
      consulta,
      descripcion,
      "Pendiente",
      ""
    ]);

    /* CORREO AL ADMIN */
    MailApp.sendEmail({
      to: adminEmail,
      subject: "Nueva consulta recibida",
      htmlBody:
      "<h2>Nueva consulta</h2>" +
      "<b>ID:</b> "+id+"<br>" +
      "<b>Nombre:</b> "+nombre+"<br>" +
      "<b>Correo:</b> "+correo+"<br>" +
      "<b>Teléfono:</b> "+telefono+"<br>" +
      "<b>Consulta:</b> "+consulta+"<br>" +
      "<b>Descripción:</b> "+descripcion+"<br>" +
      "<b>Fecha:</b> "+fechaHora
    });

    /* CORREO AL CLIENTE */
    if(correo != ""){

      MailApp.sendEmail({
        to: correo,
        subject: "Hemos recibido tu consulta",
        htmlBody:
        "Hola <b>"+nombre+"</b><br><br>" +
        "Hemos recibido tu consulta correctamente.<br><br>" +
        "<b>ID de seguimiento:</b> "+id+"<br><br>" +
        "Pronto te responderemos.<br><br>" +
        "Multiservicios GABE"
      });

    }

    return ContentService
    .createTextOutput(JSON.stringify({
      result:"success",
      id:id
    }))
    .setMimeType(ContentService.MimeType.JSON);

  }

}



/* =========================
   GET
========================= */

function doGet(e){

  /* TODAS CONSULTAS */

  if(e.parameter.tipo == "consultas"){

    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    var data = sheet.getDataRange().getValues();

    var resultado = [];

    for(var i=1;i<data.length;i++){

      resultado.push({
        id:data[i][0],
        fecha:data[i][1],
        nombre:data[i][2],
        correo:data[i][3],
        telefono:data[i][4],
        consulta:data[i][5],
        descripcion:data[i][6],
        estado:data[i][7]
      });

    }

    return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);

  }


  /* CERRAR CONSULTA */

  if(e.parameter.tipo == "cerrar"){

    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    var id = e.parameter.id;

    var data = sheet.getDataRange().getValues();

    for(var i=1;i<data.length;i++){

      if(data[i][0] == id){

        sheet.getRange(i+1,8).setValue("Cerrado");

        var fecha = Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd HH:mm:ss"
        );

        sheet.getRange(i+1,9).setValue(fecha);

        return ContentService
        .createTextOutput(JSON.stringify({result:"ok"}))
        .setMimeType(ContentService.MimeType.JSON);

      }

    }

  }


  /* OBTENER ORDENES */

  if(e.parameter.tipo == "ordenes"){

    return ContentService
    .createTextOutput(JSON.stringify(obtenerOrdenes()))
    .setMimeType(ContentService.MimeType.JSON);

  }


  /* CERRAR ORDEN */

  if(e.parameter.tipo == "cerrarOrden"){

    var sheet = getSheetOrdenes();
    var id = e.parameter.id;

    var data = sheet.getDataRange().getValues();

    for(var i=1;i<data.length;i++){

      if(data[i][0] == id){

        sheet.getRange(i+1,9).setValue("Terminada");

        var fecha = Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd HH:mm:ss"
        );

        sheet.getRange(i+1,10).setValue(fecha);

        return ContentService
        .createTextOutput(JSON.stringify({result:"ok"}))
        .setMimeType(ContentService.MimeType.JSON);

      }

    }

  }


/* PROCESAR ORDEN */

if(e.parameter.tipo == "procesarOrden"){

  var sheet = getSheetOrdenes();
  var id = e.parameter.id;

  var data = sheet.getDataRange().getValues();

  for(var i=1;i<data.length;i++){

    if(data[i][0] == id){

      sheet.getRange(i+1,9).setValue("Proceso");

      return ContentService
      .createTextOutput(JSON.stringify({result:"ok"}))
      .setMimeType(ContentService.MimeType.JSON);

    }

  }

  

}

/* ELIMINAR ORDEN */

if(e.parameter.tipo == "eliminarOrden"){

  var sheet = getSheetOrdenes();
  var id = e.parameter.id;

  var data = sheet.getDataRange().getValues();

  for(var i=1;i<data.length;i++){

    if(data[i][0] == id){

      sheet.deleteRow(i+1);

      return ContentService
      .createTextOutput(JSON.stringify({result:"ok"}))
      .setMimeType(ContentService.MimeType.JSON);

    }

  }

}


  
  /* OBTENER MOVIMIENTOS INVENTARIO (AGREGADO) */
  if(e.parameter.tipo == "movimientosInventario"){
    return ContentService
      .createTextOutput(JSON.stringify(obtenerMovimientos()))
      .setMimeType(ContentService.MimeType.JSON);
  }
/* INVENTARIO */

  if(e.parameter.tipo == "inventario"){

    return ContentService
    .createTextOutput(JSON.stringify(obtenerInventario()))
    .setMimeType(ContentService.MimeType.JSON);

  }

  /* ===== BUSCADOR INVENTARIO ===== */

  if(e.parameter.tipo == "buscarInventario"){

    var r = buscarInventario(
      e.parameter.campo,
      e.parameter.texto
    );

    return ContentService
    .createTextOutput(JSON.stringify(r))
    .setMimeType(ContentService.MimeType.JSON);

  }


  if(e.parameter.tipo == "modificarProducto"){
  return modificarProducto(e);
  }

   /* ELIMINAR PRODUCTO */

  /* CORRECCION: permitir eliminarProducto por POST */

  if(e.parameter.tipo == "eliminarProducto"){

  var r = eliminarProducto(e.parameter.id);

  return ContentService
  .createTextOutput(JSON.stringify(r))
  .setMimeType(ContentService.MimeType.JSON);

  }


  
  /* PROVEEDORES */

  if(e.parameter.tipo == "proveedores"){

    var sheet = getSheetInventario();
    var data = sheet.getDataRange().getValues();

    var lista = [];

    for(var i=1;i<data.length;i++){
      if(data[i][7] && lista.indexOf(data[i][7]) == -1){
        lista.push(data[i][7]);
      }
    }

    return ContentService
    .createTextOutput(JSON.stringify(lista))
    .setMimeType(ContentService.MimeType.JSON);

  }


  /* UBICACIONES */

  /* UBICACIONES */

   if(e.parameter.tipo == "ubicaciones"){

   var sheetInv = getSheetInventario();
   var ss = sheetInv.getParent();

   var hojaUb = ss.getSheetByName("Ubicaciones");

   var lista = [];

   /* PRIMERO: leer hoja Ubicaciones */

   if(hojaUb){

    var data = hojaUb.getRange(2,1,Math.max(hojaUb.getLastRow()-1,0),1).getValues();

    data.forEach(function(r){
      if(r[0]) lista.push(r[0]);
    });

  }

  /* SI ESTA VACIO usa inventario (tu logica original) */

  if(lista.length == 0){

    var dataInv = sheetInv.getDataRange().getValues();

    for(var i=1;i<dataInv.length;i++){
      if(dataInv[i][8] && lista.indexOf(dataInv[i][8]) == -1){
        lista.push(dataInv[i][8]);
      }
    }

  }

  return ContentService
  .createTextOutput(JSON.stringify(lista))
  .setMimeType(ContentService.MimeType.JSON);

}


/* SIGUIENTE ORDEN */

  if(e.parameter.tipo == "siguienteOrden"){
    return siguienteOrden();
  }

}



/* =========================
   HOJA INVENTARIO
========================= */
function getSheetInventario(){

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Inventario");

  if(!sheet){

    sheet = ss.insertSheet("Inventario");

    sheet.appendRow([
      "ID",
      "Fecha",
      "Codigo",
      "Producto",
      "Categoria",
      "Cantidad",
      "Precio",
      "Proveedor",
      "Ubicacion"
    ]);

  }

  return sheet;

}
/* =========================
   AGREGAR PRODUCTO INVENTARIO
========================= */


function agregarProducto(e){

  var sheet = getSheetInventario();
  var ss = sheet.getParent();

  var lastRow = sheet.getLastRow() + 1;

  var id = "P-" + ("0000" + lastRow).slice(-4);

  var fecha = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );

  var codigo = e.parameter.codigo || "";
  var producto = e.parameter.producto || "";
  var categoria = e.parameter.categoria || "";
  var cantidad = parseInt(e.parameter.cantidad || 0);
  var precio = e.parameter.precio || "";
  var proveedor = e.parameter.proveedor || "";
  var sucursalSeleccionada = e.parameter.ubicacion || "";
  var multiSucursal = e.parameter.multiSucursal == "true";

  /* ===== REGISTRO ORIGINAL (NO MODIFICADO) ===== */
  sheet.appendRow([
    id,
    fecha,
    codigo,
    producto,
    categoria,
    cantidad,
    precio,
    proveedor,
    sucursalSeleccionada
  ]);

  /* ===== MEJORA MULTISUCURSAL (SOLO SI CHECK ESTA ACTIVADO) ===== */

  if(multiSucursal){

    var hojaSuc = ss.getSheetByName("Ubicaciones");

    if(hojaSuc){

      var sucursales = hojaSuc
        .getRange(2,1,Math.max(hojaSuc.getLastRow()-1,0),1)
        .getValues()
        .flat()
        .filter(String);

      sucursales.forEach(function(s){

        if(s == sucursalSeleccionada) return;

        var row = sheet.getLastRow()+1;
        var newId = "P-" + ("0000" + row).slice(-4);

        sheet.appendRow([
          newId,
          fecha,
          codigo,
          producto,
          categoria,
          0,
          precio,
          proveedor,
          s
        ]);

      });

    }

  }

  return ContentService
  .createTextOutput(JSON.stringify({
    result:"ok",
    id:id
  }))
  .setMimeType(ContentService.MimeType.JSON);

}




/* =========================
   OBTENER INVENTARIO
========================= */

function obtenerInventario(){

  var sheet = getSheetInventario();
  var data = sheet.getDataRange().getValues();

  var productos = [];

  for(var i=1;i<data.length;i++){

    productos.push({

      id:data[i][0],
      fecha:data[i][1],
      codigo:data[i][2],
      producto:data[i][3],
      categoria:data[i][4],
      cantidad:data[i][5],
      precio:data[i][6],
      proveedor:data[i][7],
      ubicacion:data[i][8]

    });

  }

  return productos;

}


/* =========================
   ELIMINAR PRODUCTO
========================= */

function eliminarProducto(id){

  var sheet = getSheetInventario();
  var data = sheet.getDataRange().getValues();

  for(var i=1;i<data.length;i++){

    if(data[i][0] == id){

      sheet.deleteRow(i+1);

      return {result:"ok"};

    }

  }

  return {result:"error"};

}


/* =========================
   MODIFICAR PRODUCTO
========================= */

function modificarProducto(e){

  var sheet = getSheetInventario();
  var data = sheet.getDataRange().getValues();

  for(var i=1;i<data.length;i++){

    if(data[i][0] == e.parameter.id){

      sheet.getRange(i+1,3).setValue(e.parameter.codigo || "");
      sheet.getRange(i+1,4).setValue(e.parameter.producto || "");
      sheet.getRange(i+1,5).setValue(e.parameter.categoria || "");
      sheet.getRange(i+1,6).setValue(e.parameter.cantidad || "");
      sheet.getRange(i+1,7).setValue(e.parameter.precio || "");
      sheet.getRange(i+1,8).setValue(e.parameter.proveedor || "");
      sheet.getRange(i+1,9).setValue(e.parameter.ubicacion || "");

      return ContentService
      .createTextOutput(JSON.stringify({result:"ok"}))
      .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
  .createTextOutput(JSON.stringify({result:"error"}))
  .setMimeType(ContentService.MimeType.JSON);

}



/* =========================
   BUSCAR INVENTARIO RAPIDO
========================= */

function buscarInventario(tipo, texto){

  var sheet = getSheetInventario();
  var data = sheet.getDataRange().getValues();

  var resultado = [];

  texto = (texto || "").toLowerCase();

  for(var i=1;i<data.length;i++){

    var producto = {
      id:data[i][0],
      codigo:data[i][2],
      producto:data[i][3],
      categoria:data[i][4],
      cantidad:data[i][5],
      precio:data[i][6],
      proveedor:data[i][7],
      ubicacion:data[i][8]
    };

    if(producto[tipo] && producto[tipo].toString().toLowerCase().includes(texto)){
      resultado.push(producto);
    }

  }

  return resultado;

}


/* =========================
   HOJA MOVIMIENTOS INVENTARIO
========================= */
function getSheetMovimientos(){

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Movimientos");

  if(!sheet){
    sheet = ss.insertSheet("Movimientos");
    sheet.appendRow([
      "Fecha",
      "Codigo",
      "Producto",
      "Tipo",
      "Cantidad",
      "Descripcion",
      "Usuario",
       "sucursal"
    ]);
  }

  return sheet;
}

/* =========================
   REGISTRAR MOVIMIENTO INVENTARIO
========================= */
function registrarMovimiento(e){

  var sheetInv = getSheetInventario();
  var sheetMov = getSheetMovimientos();

  var codigo = e.parameter.codigo || "";
  var cantidad = parseInt(e.parameter.cantidad || 0);
  var tipo = e.parameter.tipoMovimiento || "Entrada";
  var descripcion = e.parameter.descripcion || "";
  var usuario = e.parameter.usuario || "";
  var sucursal = e.parameter.sucursal || "";
  var transferencia = e.parameter.transferencia == "true";
  var destino = e.parameter.destino || "";

  var data = sheetInv.getDataRange().getValues();

  var filaOrigen = -1;
  var filaDestino = -1;

  for(var i=1;i<data.length;i++){

    if(data[i][2] == codigo && data[i][8] == sucursal){
      filaOrigen = i;
    }

    if(data[i][2] == codigo && data[i][8] == destino){
      filaDestino = i;
    }

  }

  if(filaOrigen == -1){
    return ContentService
    .createTextOutput(JSON.stringify({
      result:"error",
      mensaje:"Producto no encontrado en sucursal origen"
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }

   var stockOrigen = parseInt(data[filaOrigen][5] || 0);

  /* VALIDAR STOCK SOLO PARA SALIDA O TRANSFERENCIA */

  if((tipo == "Salida" || transferencia) && stockOrigen < cantidad){
   return ContentService
   .createTextOutput(JSON.stringify({
    result:"error",
    mensaje:"Stock insuficiente"
  }))
  .setMimeType(ContentService.MimeType.JSON);
 }

  /* DESCONTAR ORIGEN */

/* MOVIMIENTO NORMAL (NO TRANSFERENCIA) */

 if(!transferencia){

  if(tipo == "Entrada"){
    stockOrigen += cantidad;
  }

  if(tipo == "Salida"){
    stockOrigen -= cantidad;
  }

  sheetInv.getRange(filaOrigen+1,6).setValue(stockOrigen);

 }

    /* TRANSFERENCIA */

 if(transferencia){

  stockOrigen -= cantidad;

  sheetInv.getRange(filaOrigen+1,6).setValue(stockOrigen);

 }

  /* SUMAR DESTINO */

  if(transferencia && filaDestino != -1){

    var stockDestino = parseInt(data[filaDestino][5] || 0);
    stockDestino += cantidad;

    sheetInv.getRange(filaDestino+1,6).setValue(stockDestino);

  }

  var fecha = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );

  sheetMov.appendRow([
    fecha,
    codigo,
    data[filaOrigen][3],
    transferencia ? "Transferencia" : tipo,
    cantidad,
    descripcion,
    usuario,
    sucursal,
    destino
  ]);

  return ContentService
  .createTextOutput(JSON.stringify({result:"ok"}))
  .setMimeType(ContentService.MimeType.JSON);

}


/* DESCONTAR INVENTARIO AUTOMATICO */
function descontarInventario(codigo,cantidad){
var sheet = getSheetInventario();
var data = sheet.getDataRange().getValues();
for(var i=1;i<data.length;i++){
if(data[i][2]==codigo){
var actual=parseInt(data[i][5]||0);
sheet.getRange(i+1,6).setValue(actual-cantidad);
}
}
}



/* ===== MEJORA: OBTENER SUCURSALES DESDE HOJA 'Ubicaciones' (SIN CAMBIAR LOGICA EXISTENTE) ===== */
function obtenerUbicaciones(){

  var sheetInv = getSheetInventario();
  var ss = sheetInv.getParent();
  var sheet = ss.getSheetByName("Ubicaciones");

  if(!sheet) return [];

  var data = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  var ubicaciones = [];

  data.forEach(function(r){
    if(r[0]){
      ubicaciones.push(r[0]);
    }
  });

  return ubicaciones;
}

/* ===== RESPUESTA PARA ?tipo=ubicaciones SIN MODIFICAR TU LOGICA ===== */
function responderUbicaciones(){

  return ContentService
  .createTextOutput(JSON.stringify(obtenerUbicaciones()))
  .setMimeType(ContentService.MimeType.JSON);

}

/* ===== HOOK SEGURO PARA doGet (NO MODIFICA TU CODIGO EXISTENTE) ===== */
function verificarUbicacionesHook(e){

  if(e && e.parameter && e.parameter.tipo == "ubicaciones"){
    return responderUbicaciones();
  }

  return null;
}
