require('dotenv').config()

const express = require('express')
const cors = require('cors')
const { createClient } = require('@supabase/supabase-js')
const ExcelJS = require('exceljs')

const app = express()

// ===== CORS =====
app.use(cors({
  origin: [
    "https://air-quality-frontend.onrender.com"
  ]
}))

app.use(express.json())

// ===== SUPABASE =====
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_KEY
)


// =============================
// 📊 ENDPOINT DATOS
// =============================
app.get('/data', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('air_readings')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(50)

    if (error) throw error

    res.json(data)

  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Error obteniendo datos' })
  }
})


// =============================
// 📥 EXCEL PROFESIONAL COMPLETO
// =============================
app.get('/download', async (req, res) => {
  try {

    const { data, error } = await supabase
      .from('air_readings')
      .select('*')
      .order('created_at', { ascending: false })

    if (error) throw error

    const workbook = new ExcelJS.Workbook()
    const worksheet = workbook.addWorksheet('Reporte')

    // ===== TITULO =====
    worksheet.mergeCells('A1:I1')
    worksheet.getCell('A1').value = 'INFORME GENERAL DE MONITOREO'
    worksheet.getCell('A1').font = { size: 18, bold: true }
    worksheet.getCell('A1').alignment = { horizontal: 'center' }

    worksheet.mergeCells('A2:I2')
    worksheet.getCell('A2').value =
      `Fecha de generación: ${new Date().toLocaleString()}`
    worksheet.getCell('A2').alignment = { horizontal: 'center' }

    worksheet.addRow([])

    // ===== ENCABEZADOS =====
    const headerRowNumber = worksheet.rowCount + 1

    const headerRow = worksheet.addRow([
      'Fecha',
      'Temperatura',
      'Humedad',
      'PM2.5',
      'PM10',
      'CO',
      'NO2',
      'O3',
      'SO2'
    ])

    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E40AF' }
      }
      cell.alignment = { horizontal: 'center' }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
    })

    worksheet.columns = [
      { width: 22 },
      { width: 15 },
      { width: 15 },
      { width: 12 },
      { width: 12 },
      { width: 10 },
      { width: 10 },
      { width: 10 },
      { width: 10 }
    ]

    // ===== FILA DONDE INICIAN DATOS =====
    const firstDataRow = worksheet.rowCount + 1

    data.forEach(row => {
      worksheet.addRow([
        new Date(row.created_at).toLocaleString(),
        row.temperature,
        row.humidity,
        row.pm25,
        row.pm10,
        row.co,
        row.no2,
        row.o3,
        row.so2
      ])
    })

    const lastDataRow = worksheet.rowCount

    // ===== PROMEDIOS DINÁMICOS CORRECTOS =====
    const avgRow = worksheet.addRow([
      'PROMEDIO',
      { formula: `AVERAGE(B${firstDataRow}:B${lastDataRow})` },
      { formula: `AVERAGE(C${firstDataRow}:C${lastDataRow})` },
      { formula: `AVERAGE(D${firstDataRow}:D${lastDataRow})` },
      { formula: `AVERAGE(E${firstDataRow}:E${lastDataRow})` },
      { formula: `AVERAGE(F${firstDataRow}:F${lastDataRow})` },
      { formula: `AVERAGE(G${firstDataRow}:G${lastDataRow})` },
      { formula: `AVERAGE(H${firstDataRow}:H${lastDataRow})` },
      { formula: `AVERAGE(I${firstDataRow}:I${lastDataRow})` }
    ])

    avgRow.eachCell(cell => {
      cell.font = { bold: true }
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE5E7EB' }
      }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
      cell.alignment = { horizontal: 'center' }
    })

    // ===== FILTROS Y CONGELAR FILAS =====
    worksheet.autoFilter = {
      from: `A${headerRowNumber}`,
      to: `I${lastDataRow}`
    }

    worksheet.views = [
      { state: 'frozen', ySplit: headerRowNumber }
    ]

    // ===== DESCARGA =====
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    res.setHeader(
      'Content-Disposition',
      'attachment; filename=Informe_Monitoreo_Ambiental.xlsx'
    )

    await workbook.xlsx.write(res)
    res.end()

  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Error generando Excel' })
  }
})


// ===== SERVIDOR =====
const PORT = process.env.PORT || 3000

app.listen(PORT, () => {
  console.log(`Servidor corriendo en puerto ${PORT}`)
})
