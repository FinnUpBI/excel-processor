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
      
      if (potentialHeaders.length >= 3) {
        const hasReferencia = potentialHeaders.some(h => 
          String(h).toLowerCase().includes('referencia') ||
          String(h).toLowerCase().includes('codigo') ||
          String(h).toLowerCase().includes('ref')
        );
        
        const hasDescripcion = potentialHeaders.some(h =>
          String(h).toLowerCase().includes('descripcion') ||
          String(h).toLowerCase().includes('producto')
        );
        
        if (hasReferencia || hasDescripcion) {
          productHeaders = row.map(cell => String(cell).trim()).filter(h => h);
          productTableStartIndex = i;
          break;
        }
      }
    }
    
    // PASO 2: EXTRAER LÍNEAS DE PRODUCTOS
    const productos = [];
    
    if (productTableStartIndex !== -1 && productHeaders.length > 0) {
      // Procesar filas después de los encabezados
      for (let i = productTableStartIndex + 1; i < rawData.length; i++) {
        const row = rawData[i];
        
        // Saltar filas completamente vacías
        const nonEmptyCells = row.filter(cell => cell && String(cell).trim() !== '');
        if (nonEmptyCells.length === 0) continue;
        
        // Crear objeto de producto
        const producto = {};
        let hasRelevantData = false;
        
        productHeaders.forEach((header, index) => {
          if (header && row[index] !== undefined && String(row[index]).trim() !== '') {
            producto[header] = row[index];
            
            // Verificar si tiene datos relevantes (no solo espacios)
            if (String(row[index]).trim().length > 0) {
              hasRelevantData = true;
            }
          }
        });
        
        if (hasRelevantData) {
          productos.push(producto);
        }
      }
    }
    
    // PASO 3: EXTRAER TODO EL CONTENIDO COMO TEXTO (backup completo)
    let contenidoCompleto = "=== CONTENIDO COMPLETO DEL EXCEL ===\n\n";
    
    rawData.forEach((row, index) => {
      const rowText = row.filter(cell => cell && String(cell).trim() !== '').join(' | ');
      if (rowText) {
        contenidoCompleto += `Fila ${index + 1}: ${rowText}\n`;
      }
    });
    
    // PASO 4: FORMATEAR PARA EL LLM (estructurado y legible)
    let pedidoTexto = "=== INFORMACIÓN DEL PEDIDO ===\n\n";
    
    // Información del encabezado
    if (Object.keys(headerInfo).length > 0) {
      pedidoTexto += "--- DATOS DEL PEDIDO ---\n";
      Object.entries(headerInfo).forEach(([key, value]) => {
        const label = key.replace(/_/g, ' ').toUpperCase();
        pedidoTexto += `${label}: ${value}\n`;
      });
      pedidoTexto += "\n";
    }
    
    // Líneas de productos
    if (productos.length > 0) {
      pedidoTexto += "--- LÍNEAS DE PRODUCTOS ---\n\n";
      
      productos.forEach((producto, index) => {
        pedidoTexto += `Línea ${index + 1}:\n`;
        
        Object.entries(producto).forEach(([key, value]) => {
          // Formatear números si es precio/cantidad
          let formattedValue = value;
          if (typeof value === 'number' && 
              (key.toLowerCase().includes('precio') || 
               key.toLowerCase().includes('total') ||
               key.toLowerCase().includes('importe'))) {
            formattedValue = value.toFixed(2);
          }
          
          pedidoTexto += `  ${key}: ${formattedValue}\n`;
        });
        
        pedidoTexto += "\n";
      });
    }
    
    // PASO 5: CALCULAR ESTADÍSTICAS
    let sumaTotal = null;
    const columnasConTotal = productHeaders.filter(h => 
      h.toLowerCase().includes('total') || h.toLowerCase().includes('importe')
    );
    
    if (columnasConTotal.length > 0 && productos.length > 0) {
      const columnaTotal = columnasConTotal[0];
      sumaTotal = productos.reduce((sum, prod) => {
        const valor = parseFloat(prod[columnaTotal]) || 0;
        return sum + valor;
      }, 0);
    }
    
    // RESPUESTA COMPLETA
    return res.status(200).json({
      success: true,
      
      // Información estructurada del encabezado
      informacion_pedido: headerInfo,
      
      // Líneas de productos estructuradas
      lineas_productos: productos,
      
      // Texto formateado para LLM (lo principal)
      pedido_texto: pedidoTexto,
      
      // Contenido completo como backup
      contenido_completo: contenidoCompleto,
      
      // Metadata
      metadata: {
        total_filas_excel: rawData.length,
        total_lineas_productos: productos.length,
        columnas_productos: productHeaders,
        nombre_hoja: sheetName,
        formato_archivo: workbook.bookType,
        fila_inicio_productos: productTableStartIndex + 1,
        suma_total: sumaTotal ? sumaTotal.toFixed(2) : null,
        timestamp: new Date().toISOString()
      }
    });
    
  } catch (error) {
    console.error('Error procesando Excel:', error);
    return res.status(500).json({
      success: false,
      error: error.message,
      tipo_error: error.name,
      stack: error.stack,
      ayuda: 'Verifica que el archivo sea un Excel válido (.xls o .xlsx)'
    });
  }
};
