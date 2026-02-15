require('dotenv').config()

const express = require('express')
const cors = require('cors')
const { createClient } = require('@supabase/supabase-js')
const ExcelJS = require('exceljs')

const app = express()

app.use(cors({
  origin: ["https://air-quality-frontend.onrender.com"]
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
// 📥 EXCEL PROFESIONAL FINAL
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

    // ===== FUNCION HORA COLOMBIA =====
    const fechaColombia = (fecha) =>
      new Date(fecha).toLocaleString('es-CO', {
        timeZone: 'America/Bogota'
      })

    // ===== TITULO =====
    worksheet.mergeCells('A1:I1')
    const title = worksheet.getCell('A1')
    title.value = 'INFORME GENERAL DE MONITOREO'
    title.font = { size: 18, bold: true }
    title.alignment = { horizontal: 'center' }

    worksheet.mergeCells('A2:I2')
    worksheet.getCell('A2').value =
      `Fecha de generación: ${fechaColombia(new Date())}`
    worksheet.getCell('A2').alignment = { horizontal: 'center' }

    // ===== ESTADO ACTUAL =====
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
      statusCell.value =
        `Estado actual: ${statusText} (PM2.5 = ${latest.pm25})`
      statusCell.alignment = { horizontal: 'center' }
      statusCell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: statusColor }
      }
    }

    worksheet.addRow([])

    // ===== ENCABEZADOS =====
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

    // ===== VARIABLES PROMEDIOS =====
    let sumTemp = 0
    let sumHum = 0
    let sumPM25 = 0
    let sumPM10 = 0
    let sumCO = 0
    let sumNO2 = 0
    let sumO3 = 0
    let sumSO2 = 0

    // ===== FILAS DE DATOS =====
    data.forEach(row => {

      const newRow = worksheet.addRow([
        fechaColombia(row.created_at),
        row.temperature,
        row.humidity,
        row.pm25,
        row.pm10,
        row.co,
        row.no2,
        row.o3,
        row.so2
      ])

      newRow.eachCell(cell => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
        cell.alignment = { horizontal: 'center' }
      })

      // Color dinámico PM2.5
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

      sumTemp += Number(row.temperature) || 0
      sumHum += Number(row.humidity) || 0
      sumPM25 += Number(row.pm25) || 0
      sumPM10 += Number(row.pm10) || 0
      sumCO += Number(row.co) || 0
      sumNO2 += Number(row.no2) || 0
      sumO3 += Number(row.o3) || 0
      sumSO2 += Number(row.so2) || 0
    })

    const total = data.length || 1

    // ===== FILA PROMEDIOS =====
    const avgRow = worksheet.addRow([
      'PROMEDIO',
      (sumTemp / total).toFixed(2),
      (sumHum / total).toFixed(2),
      (sumPM25 / total).toFixed(2),
      (sumPM10 / total).toFixed(2),
      (sumCO / total).toFixed(2),
      (sumNO2 / total).toFixed(2),
      (sumO3 / total).toFixed(2),
      (sumSO2 / total).toFixed(2)
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

    worksheet.views = [{ state: 'frozen', ySplit: 5 }]

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
