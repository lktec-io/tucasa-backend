require('dotenv').config();
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const mysql = require('mysql2/promise');
const ExcelJS = require('exceljs');


const app = express();
app.use(cors());
app.use(bodyParser.json());

// create pool
const pool = mysql.createPool({
  host: "127.0.0.1",
  user: "root",
  password: "Leonard1234#1234",
  database: "tucasa",
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
});

app.post('/api/register', async (req, res) => {
  try {
    const {
      full_name, gender, address, baptism_status,
      email, phone, course, year_of_study, leadership_position
    } = req.body;

    // 1️⃣ Check kama user tayari yupo kwa email au phone
    const [existing] = await pool.execute(
      "SELECT id FROM students WHERE email = ? OR phone = ? LIMIT 1",
      [email, phone]
    );

    if (existing.length > 0) {
      return res.status(409).json({
        success: false,
        message: "You are already registered!"
      });
    }

    // 2️⃣ Kama hayupo, then insert
    const [result] = await pool.execute(
      `INSERT INTO students 
       (full_name, gender, address, baptism_status, email, phone, course, year_of_study, leadership_position)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [full_name, gender, address, baptism_status, email, phone, course, year_of_study, leadership_position]
    );

    res.json({ success: true, id: result.insertId });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});


app.get('/api/students', async (req, res) => {
  try {
    const [rows] = await pool.query('SELECT * FROM students ORDER BY created_at ASC');
    res.json(rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/students/download', async (req, res) => {
  try {
    const [rows] = await pool.query('SELECT * FROM students ORDER BY created_at ASC');

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Tucasa-tia');

    sheet.columns = [
      { header: 'S/N', key: 'id', width: 3 },
      { header: 'Full Name', key: 'full_name', width: 28 },
      { header: 'Gender', key: 'gender', width: 8 },
      { header: 'Address', key: 'address', width: 18 },
      { header: 'Baptism Status', key: 'baptism_status', width: 14 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Phone', key: 'phone', width: 14 },
      { header: 'Course', key: 'course', width: 10 },
      { header: 'Year', key: 'year_of_study', width: 5 },
      { header: 'Leadership Position', key: 'leadership_position', width: 20 },
      { header: 'Registered on', key: 'created_at', width: 14 },
    ];

    rows.forEach(r => sheet.addRow(r));

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Tucasa-Registration.xlsx`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Server listening on ${port}`));
