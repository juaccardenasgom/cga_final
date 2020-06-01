const express = require('express');
const Excel = require('exceljs');

var app = express();

const config = (req,res,netx) => {
    var fileName = 'output.xlsx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader("Content-Disposition", 'attachment; filename=' + fileName);
    netx();
}

// Config Middleware
app.use(config);

// Parse URL-encoded bodies (as sent by HTML forms)
app.use(express.urlencoded({ extended: true }))

// Parse JSON bodies (as sent by API clients)
app.use(express.json());

app.post("/api/create", async (req,res) => {
    var workbook = new Excel.Workbook();
    workbook.views = [
        {
          x: 0, y: 0, width: 10000, height: 20000,
          firstSheet: 0, activeTab: 1, visibility: 'visible'
        }
    ]

    var worksheet = workbook.addWorksheet('Máquina');

    worksheet.columns = [
        { header: 'Nombre', key: 'name'},
        { header: 'Valor', key: 'value'},
        { header: 'Unidad', key: 'unit'},
        { header: 'Comentarios', key: 'comment'},
    ];

    for(var pieces in req.body.data) {
        if(req.body.data.hasOwnProperty(pieces)){
            for(var piece in req.body.data[pieces]){
                if(req.body.data[pieces].hasOwnProperty(piece)){
                  worksheet.addRow(req.body.data[pieces][piece])
                  // console.log(Object.keys(req.body.data[pieces][piece]))
                }
            }
        }
    }

    try{
        await workbook.xlsx.write(res);
    }catch(err){
        console.log(err)
        res.status(500).json({error:err}) ;
    }

    res.end();
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server started on port: ${PORT}`));