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
      cellDates: true,
      raw: false
    });
    
    if (workbook.SheetNames.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'El archivo Excel no contiene hojas'
      });
    }
    
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    const rawData = XLSX.utils.sheet_to_json(sheet, { 
      header: 1,
      defval: '',
      raw: false
    });
    
    if (rawData.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'La hoja Excel está completamente vacía'
      });
    }
    
    // FUNCIÓN AUXILIAR: Limpiar texto
    const cleanText = (text) => {
      if (!text) return '';
      return String(text).trim().replace(/\s+/g, ' ');
    };
    
    // FUNCIÓN AUXILIAR: Verificar si una fila tiene contenido significativo
    const hasSignificantContent = (row) => {
      const cleanedCells = row.map(cell => cleanText(cell)).filter(c => c.length > 0);
      return cleanedCells.length >= 2;
    };
    
    // PASO 1: EXTRAER INFORMACIÓN DEL ENCABEZADO
    const headerInfo = {};
    let productTableStartIndex = -1;
    let productHeaders = [];
    
    for (let i = 0; i < Math.min(20, rawData.length); i++) {
      const row = rawData[i];
      const rowText = row.map(c => cleanText(c)).join(' ');
      
      // Buscar número de pedido
      if (rowText.toLowerCase().includes('pedido')) {
        const match = rowText.match(/pedido\s*n[º°]?\s*:?\s*(\d+)/i);
        if (match) {
          headerInfo.numero_pedido = match[1];
        }
      }
      
      // Buscar fecha
      if (rowText.toLowerCase().includes('fecha')) {
        const match = rowText.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
        if (match) {
          headerInfo.fecha = match[1];
        }
      }
      
      // Buscar cliente
      if (i >= 7 && i <= 11) {
        const cleaned = cleanText(rowText);
        if (cleaned.length > 10 && 
            !cleaned.toLowerCase().includes('fecha') && 
            !cleaned.toLowerCase().includes('pedido') &&
            !cleaned.toLowerCase().includes('grupauto') &&
            !cleaned.toLowerCase().includes('pág')) {
          
          if (!headerInfo.cliente) {
            headerInfo.cliente = cleaned;
          } else if (!headerInfo.direccion && cleaned.length > 5) {
            headerInfo.direccion = cleaned;
          }
        }
      }
      
      // DETECTAR TABLA DE PRODUCTOS
      const cleanedRow = row.map(c => cleanText(c).toUpperCase());
      
      const hasReferencia = cleanedRow.some(h => 
        h.includes('REFERENCIA') || h.includes('CODIGO') || h.includes('REF')
      );
      
      const hasDescripcion = cleanedRow.some(h =>
        h.includes('DESCRIPCION') || h.includes('PRODUCTO') || h.includes('ARTICULO')
      );
      
      const hasCantidad = cleanedRow.some(h =>
        h.includes('CANTIDAD')
      );
      
      if ((hasReferencia && hasDescripcion) || 
          (hasReferencia && hasCantidad) || 
          (hasDescripcion && hasCantidad)) {
        
        productHeaders = row.map(c => cleanText(c)).filter(h => h.length > 0);
        productTableStartIndex = i;
        break;
      }
    }
    
    // PASO 2: EXTRAER PRODUCTOS
    const productos = [];
    
    if (productTableStartIndex !== -1 && productHeaders.length > 0) {
      for (let i = productTableStartIndex + 1; i < rawData.length; i++) {
        const row = rawData[i];
        
        if (!hasSignificantContent(row)) continue;
        
        const firstCell = cleanText(row[0]);
        if (firstCell.length > 50 && firstCell.toLowerCase().includes('indique')) {
          break;
        }
        
        const producto = {};
        let hasProductData = false;
        
        productHeaders.forEach((header, index) => {
          const value = row[index];
          const cleanedValue = cleanText(value);
          
          if (cleanedValue.length > 0) {
            producto[header] = value;
            hasProductData = true;
          }
        });
        
        const hasRef = producto['REFERENCIA'] && cleanText(producto['REFERENCIA']).length > 0;
        const hasDesc = producto['DESCRIPCION'] && cleanText(producto['DESCRIPCION']).length > 0;
        
        if (hasProductData && (hasRef || hasDesc)) {
          productos.push(producto);
        }
      }
    }
    
    // PASO 3: EXTRAER TODO EL CONTENIDO COMO TEXTO
    let contenidoCompleto = "=== CONTENIDO COMPLETO DEL EXCEL ===\n\n";
    
    rawData.forEach((row, index) => {
      const rowText = ro
