const XLSX = require('xlsx');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  if (req.method !== 'POST') {
    return res.status(405).json({ 
      success: false,
      error: 'Método no permitido. Usa POST.' 
    });
  }
  
  try {
    const { file } = req.body;
    
    if (!file) {
      return res.status(400).json({ 
        success: false,
        error: 'No se recibió archivo. Envía { "file": "base64_aqui" }' 
      });
    }
    
    const buffer = Buffer.from(file, 'base64');
    const workbook = XLSX.read(buffer, { 
      type: 'buffer',
      cellDates: true 
    });
    
    if (workbook.SheetNames.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'El archivo Excel no contiene hojas'
      });
    }
    
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    // CONVERTIR TODO A ARRAY DE ARRAYS (captura TODO)
    const rawData = XLSX.utils.sheet_to_json(sheet, { 
      header: 1,  // Array de arrays
      defval: '', // Valor por defecto para celdas vacías
      raw: false  // Convertir todo a string para mejor manejo
    });
    
    if (rawData.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'La hoja Excel está completamente vacía'
      });
    }
    
    // PASO 1: DETECTAR INFORMACIÓN DEL ENCABEZADO
    const headerInfo = {};
    let productTableStartIndex = -1;
    let productHeaders = [];
    
    // Extraer información clave del encabezado (primeras filas)
    for (let i = 0; i < Math.min(20, rawData.length); i++) {
      const row = rawData[i];
      const rowText = row.join(' ').trim();
      
      // Buscar número de pedido
      if (rowText.toLowerCase().includes('pedido') && rowText.match(/\d+/)) {
        const match = rowText.match(/pedido\s*n[º°]?\s*:?\s*(\d+)/i);
        if (match) {
          headerInfo.numero_pedido = match[1];
        }
      }
      
      // Buscar fecha
      if (rowText.toLowerCase().includes('fecha') && rowText.match(/\d{1,2}\/\d{1,2}\/\d{2,4}/)) {
        const match = rowText.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
        if (match) {
          headerInfo.fecha = match[1];
        }
      }
      
      // Buscar nombre de cliente (suele estar después de la empresa propia)
      if (i >= 7 && i <= 11 && rowText.length > 10 && !rowText.toLowerCase().includes('fecha') && !rowText.toLowerCase().includes('pedido')) {
        if (!headerInfo.cliente && rowText.match(/[A-Z]/)) {
          headerInfo.cliente = rowText.trim();
        } else if (!headerInfo.direccion && rowText.length > 5) {
          headerInfo.direccion = rowText.trim();
        }
      }
      
      // Detectar inicio de tabla de productos
      // Buscar fila con "REFERENCIA" o "DESCRIPCION" o múltiples encabezados
      const potentialHeaders = row.filter(cell => 
        cell && String(cell).trim().length > 0
      );
      
      if (potentialHeaders.le
