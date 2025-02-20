const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const sqlite3 = require('sqlite3').verbose();
const cors = require('cors');
const ExcelJS = require('exceljs'); // Add this line

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
  cors: {
    origin: 'http://localhost:5173', // Allow frontend URL
    methods: ['GET', 'POST'],
    allowedHeaders: ['Content-Type'] // Explicitly allow Content-Type
  }
});

// Enable CORS for API requests
app.use(cors({
  origin: 'http://localhost:5173', // Allow requests from frontend
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type']
}));

// Middleware for parsing JSON
app.use(express.json());

// Database setup
const db = new sqlite3.Database('./users.db', (err) => {
  if (err) {
    console.error('Error opening database:', err.message);
  } else {
    console.log('Connected to the SQLite database.');
  }
});

// Create users table
db.run(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    email TEXT,
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`, (err) => {
  if (err) {
    console.error('Error creating users table:', err.message);
  } else {
    console.log('Users table created or already exists.');
  }
});

// API endpoint to add a new user
app.post('/api/users', (req, res) => {
  const { name, email } = req.body;
  console.log('Received request to add user:', { name, email });
  db.run(
    'INSERT INTO users (name, email) VALUES (?, ?)',
    [name, email],
    (err) => {
      if (err) {
        console.error('Error inserting user into database:', err.message);
        return res.status(500).send(err.message);
      }
      console.log('User added to database:', { name, email });
      res.sendStatus(200);
    }
  );
});

// Route to convert user data to Excel file and download
app.get('/api/users/export', (req, res) => {
  console.log('Received request to export users to Excel');
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Users');

  worksheet.columns = [
    { header: 'ID', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 30 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Timestamp', key: 'timestamp', width: 30 }
  ];

  db.all('SELECT * FROM users', [], (err, rows) => {
    if (err) {
      console.error('Error fetching users from database:', err.message);
      return res.status(500).send(err.message);
    }

    console.log('Fetched users from database:', rows);
    rows.forEach((row) => {
      worksheet.addRow(row);
    });

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename=' + 'ai_bot_users.xlsx'
    );

    workbook.xlsx.write(res).then(() => {
      console.log('Excel file written and sent to client');
      res.end();
    });
  });
});

// WebSocket setup
io.on('connection', (socket) => {
  console.log(`User connected: ${socket.id}`);

  socket.on('play-video', (videoNum) => {
    console.log(`Received play-video event with videoNum: ${videoNum}`);
    io.emit('play-video', videoNum);
  });
  socket.on('stop-video', () => {
    console.log('Received stop-video event from client');
    io.emit('stop-video'); // Broadcast the stop-video event to all clients
  });
  socket.on('disconnect', () => {
    console.log(`User disconnected: ${socket.id}`);
  });
});

server.listen(5000, () => {
  console.log('Server running on port 5000');
});