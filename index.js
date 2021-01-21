var express = require('express')

var isJson = require('is-json')

var json2xls = require('json2xls')

var bodyParser = require('body-parser')

var fs = require('fs')


const app = express()

app.use(bodyParser.json())

app.use(bodyParser.urlencoded({extended:false}))

app.set('view engine','ejs')

const PORT = 4000

app.get('/',(req,res) =>{
    res.render('index',{title:'Exporting Raw Json to Excel'})
})

app.post('/jsontoexcel',(req,res) =>{

    var jsondata = req.body.json

    var exceloutput = Date.now() + "output.xlsx"

    if(isJson(jsondata)){

        var xls = json2xls(JSON.parse(jsondata));

        fs.writeFileSync(exceloutput, xls, 'binary');

        res.download(exceloutput,(err) =>{

            if(err){

                fs.unlinkSync(exceloutput)

                res.send("Unable to Download the Excel File!")
            }
            fs.unlinkSync(exceloutput)
        })

    }
    else{

        res.send("JSON Data is not Valid")
    }
})
app.listen(PORT,() =>{
    console.log("App is Listening on port 4000")
})