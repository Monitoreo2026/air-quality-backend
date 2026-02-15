require('dotenv').config()

const express = require('express')
const cors = require('cors')
const { createClient } = require('@supabase/supabase-js')
const ExcelJS = require('exceljs')

const app = express()

app.use(cors({
  origin: [
    "https://air-quality-frontend.onrender.com"
  ]
}))

app.use(express.json())

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
// 📥 EXCEL PROFESIONAL AVANZADO
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

    worksheet.mergeCells('A1:I1')
    const titleCell = worksheet.getCell('A1')
    titleCell.value = 'INFORME GENERAL DE MONITOREO '
    titleCell.font = { size: 18, bold: true }
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' }

    worksheet.mergeCells('A2:I2')
    const subCell = worksheet.getCell('A2')
    subCell.value = `Fecha de generación: ${new Date().toLocaleString()}`
    subCell.alignment = { horizontal: 'center' }

    if (data.length > 0) {
      const latest = data[0]
      let statusText = ''
      let statusColor = ''

      if (latest.pm25 <= 50) {
        statusText = 'BUENO'
        statusColor = 'FF16A34A'
      } else if (latest.pm25 <= 150) {
        statusText = 'MODERADO'
        statusColor = 'FFEAB308'
      } else {
        statusText = 'CRITICO'
        statusColor = 'FFDC2626'
      }

      worksheet.mergeCells('A3:I3')
      const statusCell = worksheet.getCell('A3')
      statusCell.value = `Estado actual de calidad del aire: ${statusText} (PM2.5 = ${latest.pm25})`
      statusCell.alignment = { horizontal: 'center' }
      statusCell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: statusColor }
      }
    }

    worksheet.addRow([])

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
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
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

    data.forEach(row => {
      const newRow = worksheet.addRow([
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

      const pmCell = newRow.getCell(4)

      if (row.pm25 <= 50) {
        pmCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16A34A' } }
        pmCell.font = { color: { argb: 'FFFFFFFF' } }
      } else if (row.pm25 <= 150) {
        pmCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAB308' } }
      } else {
        pmCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDC2626' } }
        pmCell.font = { color: { argb: 'FFFFFFFF' } }
      }
    })

    worksheet.eachRow(row => {
      row.eachCell(cell => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
        cell.alignment = { horizontal: 'center', vertical: 'middle' }
      })
    })

    const lastRowNumber = worksheet.rowCount

    const avgRow = worksheet.addRow([
      'PROMEDIO',
      { formula: `AVERAGE(B6:B${lastRowNumber})` },
      { formula: `AVERAGE(C6:C${lastRowNumber})` },
      { formula: `AVERAGE(D6:D${lastRowNumber})` },
      { formula: `AVERAGE(E6:E${lastRowNumber})` },
      { formula: `AVERAGE(F6:F${lastRowNumber})` },
      { formula: `AVERAGE(G6:G${lastRowNumber})` },
      { formula: `AVERAGE(H6:H${lastRowNumber})` },
      { formula: `AVERAGE(I6:I${lastRowNumber})` }
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
    })

    worksheet.views = [{ state: 'frozen', ySplit: 5 }]

    worksheet.autoFilter = {
      from: 'A5',
      to: 'I5'
    }

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

const PORT = process.env.PORT || 3000

app.listen(PORT, () => {
  console.log(`Servidor corriendo en puerto ${PORT}`)
})
