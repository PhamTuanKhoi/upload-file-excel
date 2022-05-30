const express = require("express");
const app = express();
const bodyparser = require("body-parser");
const fs = require("fs");
const readXlsxFile = require("read-excel-file/node");
const mysql = require("mysql");
const multer = require("multer");
const path = require("path");
//use express static folder
app.use(express.static("./public"));
// body-parser middleware use
app.use(bodyparser.json());
app.use(
  bodyparser.urlencoded({
    extended: true,
  })
);
// Database connection
const db = mysql.createConnection({
  host: "localhost",
  user: "root",
  password: "",
  database: "uploadexcel",
});
db.connect(function (err) {
  if (err) {
    return console.error("error: " + err.message);
  }
  console.log("Connected to the MySQL server.");
});
// Multer Upload Storage
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null,  "./uploads/");
  },
  filename: (req, file, cb) => {
    cb(null, file.fieldname + "-" + Date.now() + "-" + file.originalname);
  },
});
const upload = multer({ storage: storage });
//! Routes start
//route for Home page
app.get("/", (req, res) => {
  res.sendFile(__dirname + "/index.html");
});
// -> Express Upload RestAPIs
app.post("/uploadfile", upload.single("uploadfile"), (req, res) => {
  importExcelData2MySQL("./uploads/" + req.file.filename);
  console.log(res);
});
// -> Import Excel Data to MySQL database
function importExcelData2MySQL(filePath) {
  // File path.
  readXlsxFile(filePath).then((rows) => {
   
    // Remove Header ROW
    rows.shift();

    //save database
    let query = "INSERT INTO file (date, content, amount) VALUES ?";
    db.query(query, [rows], (error, response) => {
        console.log(error || response);
      });
       
    db.end();
   
  });
}
// Create a Server
let server = app.listen(8080, function () {
  let host = server.address().address;
  let port = server.address().port;
  console.log("App listening at http://%s:%s", host, port);
});
