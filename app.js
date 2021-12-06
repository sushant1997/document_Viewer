const express = require('express');
var app = express();
var upload = require('express-fileupload');
var docxConverter = require('docx-pdf');
var path = require('path');
var fs = require('fs');
var imgTopdf = require('image-to-pdf')

const libre = require('libreoffice-convert');

//to move the pdf 
var fss = require('fs-extra');
const { response } = require('express');


const extend_pdf = '.pdf'
const extend_docx = '.docx'


var down_name

app.use(upload());
// app.use(express.static('./out'));
// app.use('/static', express.static(path.join(__dirname, './out')))

const MIME_TYPES = {
  'word': [
    'application/msword', 'application/vnd.ms-word.document.macroEnabled.12',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  ],
  'ppt': ['application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'application/vnd.ms-powerpoint', 'application/vnd.ms-powerpoint.presentation.macroEnabled.12'
  ],
  'pdf': ['application/pdf'
  ],
  'image': ['image/avif', 'image/gif', 'image/jpeg', 'image/png'
  ],
  'exel': ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
    'application/vnd.ms-excel',
    'application/vnd.ms-excel.sheet.macroEnabled.12']
};


app.get('/', function (req, res) {
  res.sendFile(__dirname + '/index.html');
})
app.post('/upload', function (req, res) {
  // return res.send(req.files.upfile);
  if (req.files.upfile) {
    var file = req.files.upfile,
      name = file.name,
      fileExt = file.name.split(".").pop(),
      mimetype = file.mimetype;
    console.log(mimetype)

    var uploadpath = __dirname + '/uploads/' + name;
    file.mv(uploadpath, function (err) {
      if (err) {
        console.log(err);
      } else {
        //Path of the downloaded or uploaded file
        var initialPath = path.join(__dirname, `/uploads/${name}`);

        //Path where the converted pdf will be placed/uploadednpm install ppt2pdf
        var out_path = path.join(__dirname, `/out/${Date.now()}.pdf`);


        let convertedFile = '';

        if (MIME_TYPES['word'].includes(mimetype)) {
          console.log('');
          convertedFile = convertDocToPdf(initialPath, out_path, function (response) {

            if (!response.status) {
              return res.send(response.error);
            }
            else {
              console.log(response.message);
              const pdfFile = response.message.filename.split("/").pop()
              const pdfUrl = `http://localhost:3000/out/${pdfFile}`;
              res.send(pdfUrl);

            }
          });
        }
        if (MIME_TYPES['pdf'].includes(mimetype)) {

          fss.move(initialPath, out_path, err => {
            if (err) {
              return console.error(err)
            } else {
              const pdfUrl = `http://localhost:3000/out/${name}`
              res.send(pdfUrl);
              console.log('success!')
            }
          })
        }


        if (MIME_TYPES['ppt'].includes(mimetype)) {

          const file = fs.readFileSync(initialPath);
          // Convert it to pdf format with undefined filter (see Libreoffice doc about filter)
          libre.convert(file, extend_pdf, undefined, (err, result) => {
            if (err) {
              console.log(`Error converting file: ${err}`);
            } else {

              // Here in done you have pdf file which you can save or transfer in another stream
              fs.writeFileSync(out_path, result)
              return res.send(`http://localhost:3000/out/${Date.now()}.pdf`);
            }
          });

        };

        if (MIME_TYPES['exel'].includes(mimetype)) {

          const file = fs.readFileSync(initialPath);

          libre.convert(file, extend_pdf, undefined, (err, result) => {
            if (err) {
              console.log(`Error converting file: ${err}`);
            } else {

              // Here in done you have pdf file which you can save or transfer in another stream
              fs.writeFileSync(out_path, result)
              return res.send(`http://localhost:3000/out/${Date.now()}.pdf`);
            }
          });

        };


        if (MIME_TYPES['image'].includes(mimetype)) {

          console.log(initialPath)
          var fileName = initialPath.split('/').pop()
          var ext = fileName.split('.').pop()
          const file = fs.readFileSync(initialPath);

          out_path = path.join(__dirname, `/out/${Date.now()}.${ext}`);

          fs.writeFileSync(out_path, file);
          res.send(file)


        };


      }
    });
  } else {
    res.send("No File selected !");
    res.end();
  }
});

app.get('/download', (req, res) => {
  //This will be used to download the converted file
  res.download(__dirname + `/uploads/${down_name}${extend_pdf}`, `${down_name}${extend_pdf}`, (err) => {
    if (err) {
      res.send(err);
    } else {
      //Delete the files from directory after the use
      console.log('Files deleted');
      const delete_path_doc = process.cwd() + `/uploads/${down_name}${extend_docx}`;
      const delete_path_pdf = process.cwd() + `/uploads/${down_name}${extend_pdf}`;
      try {
        fs.unlinkSync(delete_path_doc)
        fs.unlinkSync(delete_path_pdf)
        //file removed
      } catch (err) {
        console.error(err)
      }
    }
  })
})

app.get('/thankyou', (req, res) => {
  res.sendFile(__dirname + '/thankyou.html')
})


app.listen(3001, () => {
  console.log("Server Started at port 3000...");
})

const convertDocToPdf = (initialPath, out_path, cb) => {
  docxConverter(initialPath, out_path, function (err, result) {
    if (err) {
      let errorObj = {
        status: false,
        error: `${err}`
      };
      cb(errorObj);
    }
    else {
      cb({
        status: true,
        message: result
      });
    }
  });
};







