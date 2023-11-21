const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const express = require("express");
const cors = require("cors")

const app = express();


app.use(express.json());
app.use(cors());

// Specify the path to your Django SQLite database file
const dbPath = path.join(__dirname, '..', 'fullstack','DjangoAPI' ,'db.sqlite3');




// Create a new database connection
const db = new sqlite3.Database(dbPath, (err) => {
  if (err) {
    console.error('Database connection error:', err.message);
  } else {
    console.log('Connected to the SQLite database');
  }
});


app.get('/get-employees', (req, res) => {
  const query = 'SELECT * FROM EmployeeApp_employees';
  const {email,password} = req.body;

  db.all(query, [], (err, rows) => {
      if (err) {
          res.status(500).json({ error: err.message });
          return;
      }
      res.json({ employees: rows });
  });
});



app.get('/get-dep', (req, res) => {
  const id = req.query.id;
  const query = `SELECT * FROM EmployeeApp_departments where DepartmentId = ${id}`;
  
  db.all(query, [], (err, rows) => {
    if (err) {
      res.status(500).json({ error: err.message });
      return;
    }
    res.json({ departments: rows });
  });
});






app.listen(5400,()=>{
  console.log("app is listening on port number 5400");
})



app.post('/login', (req, res) => {
  try {
     const { email, password } = req.body;
 
     if (!email || !password) {
       return res.status(400).json({ error: 'Email and password are required' });
     }
     
     console.log("email and paswerds datatat swrds",email,password);
     const query = "SELECT * FROM EmployeeApp_employees WHERE email = ?";
     
     console.log(query)
     db.get(query, [email], (err, row) => {
       if (err) {
         console.log(err)
         return res.status(500).json({ error: err.message });
       }
 
       if (!row) {
         return res.status(401).json({ error: 'User not found' });
       }
 
       // Compare plain text password
       if (password !== row.password) {
         return res.status(401).json({ error: 'Incorrect password' });
       }
       console.log("row",row);
 
       // Remove sensitive information like password before sending the response
       res.status(200).json(row)
     });
  } catch (error) {
    res.status(400).send({
     err : error.message,
     stack : err.stack
    }) 
  }
   
 });


app.get('/get-dep', (req, res) => {
    const id = req.query.id;
    const query = `SELECT * FROM EmployeeApp_departments where DepartmentId = ${id}`;
    
    db.all(query, [], (err, rows) => {
      if (err) {
        res.status(500).json({ error: err.message });
        return;
      }
      res.json({ departments: rows });
    });
  });








