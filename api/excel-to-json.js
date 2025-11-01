const XLSX = require('xlsx');

module.exports = async (req, res) => {
  // Configurar CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  // Manejar preflight
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  // Solo POST
  if (req.method !== 'POST') {
    return res.status(405).json({ 
      success: false,
      error: 'Método no permitido. Usa POST.' 
    });
  }
  
  try {
    // Obtener archivo base64
    const { file } = req.body;
    
    if (!file) {
      return res.status(400).json({ 
        success: false,
        error: 'No se recibió archivo. Envía { "file": "base64_aqui" }' 
      });
    }
    
    // Convertir base64 a Buffer
    const buffer = Buffer.from(file, 'base64');
    
    // Leer Excel (autodetecta .xls y .xlsx)
    const workbook = XLSX.read(buffer, { 
      type: 'buffer',
      cellDates: true 
    });
    
    // Validar que tenga hojas
    if (workbook.SheetNames.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'El archivo Excel no contiene hojas'
      });
    }
    
    // Obtener primera hoja
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    // Convertir a JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    
    // Validar que tenga datos
    if (jsonData.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'La hoja Excel está vacía'
      });
    }
    
    // Formatear como texto para LLM
    let pedidoTexto = "=== INFORMACIÓN DEL PEDIDO ===\n\n";
    
    jsonData.forEach((row, index) => {
      pedidoTexto += `--- Línea ${index + 1} ---\n`;
      
      Object.keys(row).forEach(key => {
        const value = row[key];
        
        // Formatear números con 2 decimales si es precio/total
        const formattedValue = (typeof value === 'number' && 
                                (key.toLowerCase().includes('precio') || 
                                 key.toLowerCase().includes('total') ||
                                 key.toLowerCase().includes('price')))
                                ? value.toFixed(2)
                                : value;
        
        pedidoTexto += `  ${key}: ${formattedValue}\n`;
      });
      
      pedidoTexto += "\n";
    });
    
    // Calcular estadísticas
    const totalLineas = jsonData.length;
    const columnas = Object.keys(jsonData[0]);
    
    // Intentar calcular suma total
    let sumaTotal = null;
    const columnaTotal = columnas.find(col => 
      col.toLowerCase().includes('total')
    );
    
    if (columnaTotal) {
      sumaTotal = jsonData.reduce((sum, row) => 
        sum + (parseFloat(row[columnaTotal]) || 0), 0
      );
    }
    
    // Respuesta final
    return res.status(200).json({
      success: true,
      pedido_estructurado: jsonData,
      pedido_texto: pedidoTexto,
      metadata: {
        total_lineas: totalLineas,
        columnas: columnas,
        nombre_hoja: sheetName,
        formato_archivo: workbook.bookType,
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
      ayuda: 'Verifica que el archivo sea un Excel válido (.xls o .xlsx)'
    });
  }
};
```

