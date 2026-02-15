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
// 📥 EXCEL CON PROMEDIOS REALES
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
    worksheet.addRow([
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

    // ===== VARIABLES PARA PROMEDIOS =====
    let sumTemp = 0
    let sumHum = 0
    let sumPM25 = 0
    let sumPM10 = 0
    let sumCO = 0
    let sumNO2 = 0
    let sumO3 = 0
    let sumSO2 = 0

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

    // ===== FILA PROMEDIOS CALCULADA EN BACKEND =====
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
    })

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
