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
  host: process.env.DB_HOST,
  user: process.env.DB_USER,
  password: process.env.DB_PASS,
  database: process.env.DB_NAME,
  waitForConnections: true,
  connectionLimit: 10,
});

app.post('/api/register', async (req, res) => {
  try {
    const {
      full_name, gender, address, baptism_status,
      email, phone, course, year_of_study, leadership_position
    } = req.body;

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
    const [rows] = await pool.query('SELECT * FROM students ORDER BY created_at DESC');
    res.json(rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/students/download', async (req, res) => {
  try {
    const [rows] = await pool.query('SELECT * FROM students ORDER BY created_at DESC');

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Students');

    sheet.columns = [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Full Name', key: 'full_name', width: 30 },
      { header: 'Gender', key: 'gender', width: 10 },
      { header: 'Address', key: 'address', width: 40 },
      { header: 'Baptism Status', key: 'baptism_status', width: 18 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Phone', key: 'phone', width: 18 },
      { header: 'Course', key: 'course', width: 20 },
      { header: 'Year', key: 'year_of_study', width: 12 },
      { header: 'Leadership Position', key: 'leadership_position', width: 25 },
      { header: 'Created At', key: 'created_at', width: 20 },
    ];

    rows.forEach(r => sheet.addRow(r));

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=students.xlsx`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Server listening on ${port}`));
