// =====================================================================
// GOOGLE APPS SCRIPT — pegar en Extensions > Apps Script de tu Google Sheet
// Luego: Deploy > New Deployment > Web App > Anyone > Deploy
// Copiar la URL y pegarla en el formulario
// =====================================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Respuestas');
    if (!sheet) {
      sheet = ss.insertSheet('Respuestas');
      sheet.appendRow([
        'FECHA ENTREGA','MES','AÑO','CITA','CODIGO ARIUM',
        'CANTIDAD','UM RECHAZADA','MOTIVO','MONTO UM','DESCRIPCION',
        'UM MD','UXC','MONTO TOTAL','OC / OV','CLIENTE','CASA',
        'RESPONSABILIDAD','TIMESTAMP'
      ]);
      sheet.getRange(1,1,1,18).setFontWeight('bold');
    }

    sheet.appendRow([
      data.fecha || '',
      data.mes || '',
      data.anio || '',
      data.cita || '',
      data.codigo || '',
      data.cantidad || 0,
      data.um_rechazada || 'CAJA',
      data.motivo || '',
      data.monto_um || 0,
      data.descripcion || '',
      data.um_md || '',
      data.uxc || 0,
      data.monto_total || 0,
      data.oc_ov || '',
      data.cliente || '',
      data.casa || '',
      data.responsabilidad || '',
      data.timestamp || new Date().toISOString()
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({status:'ok'})
    ).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(
      JSON.stringify({status:'error', message: err.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Hallazgos Inplant API activa');
}
