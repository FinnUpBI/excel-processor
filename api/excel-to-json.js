const XLSX = require('xlsx');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Método no permitido. Usa POST.' });
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
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });

    if (!workbook.SheetNames.length) {
      return res.status(400).json({ 
        success: false, 
        error: 'El archivo Excel no contiene hojas' 
      });
    }

    // Recorremos todas las hojas y convertimos su contenido a texto CSV
    const contenidoCompleto = workbook.SheetNames.map(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const csv = XLSX.utils.sheet_to_csv(sheet, { FS: ' | ', RS: '\n', blankrows: false });
      return `=== HOJA: ${sheetName} ===\n${csv}`;
    }).join('\n\n');

    return res.status(200).json({
      success: true,
      contenido_completo: contenidoCompleto,
      metadata: {
        hojas: workbook.SheetNames,
        total_hojas: workbook.SheetNames.length,
        timestamp: new Date().toISOString()
      }
    });

  } catch (error) {
    console.error('Error procesando Excel:', error);
    return res.status(500).json({
      success: false,
      error: error.message,
      stack: error.stack
    });
  }
};
