const express = require('express');
const app = express();
const { stringify } = require('querystring');
const mongoose = require('mongoose')
const cors = require('cors')
const bodyParser = require('body-parser')
const multer =require('multer')
const XLSX =require('xlsx')
const path = require('path');

app.use(function(req, res, next) {
    res.header("Access-Control-Allow-Origin", "*"); 
    res.header("Access-Control-Allow-Methods","GET, POST, HEAD, OPTIONS, PUT, PATCH, DELETE");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    next();
});

mongoose.Promise=global.Promise
mongoose.set('strictQuery', true);
mongoose.connect('mongodb://127.0.0.1:27017/m', {
    useNewUrlParser: true,
    
    useUnifiedTopology: true
}).then(()=>{
    console.log('database connected')
}).catch((e)=>{
    console.log('database disconnected')
    console.log(e);
});
NVR='niveau R'
var excel = mongoose.Schema({
    'Excercice':String,
    'N° LF':String,
    'niveau R':String,
    'R':String,
    'lib':String,
    'type R':String,
    'code fonctionnel':String,
    'Code Eco':String,
    'CP':String,
    'CE':String,
    'DEP':String,
})

var excelmodel = mongoose.model('exceldata',excel)

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, "uploads");
    },
    filename: function (req, file, cb) {
      cb(null, `excel.xlsx`);
    },
  });
var upload = multer({ storage: storage });


app.post('/post',upload.single('file'),async (req, res)=>{
    var workbook = XLSX.readFile(req.file.path);
    var sheet = workbook.SheetNames;
    var x = 0;
    sheet.forEach(element=>{
        var xldata = XLSX.utils.sheet_to_json(workbook.Sheets[sheet[x]]);
        excelmodel.insertMany(xldata,(err,data)=>{
            if(err){
                console.log(err);
            }else{
                //console.log(data);
            }
        })
        x++
    })
    res.redirect('/');  
});


app.get('/get/:RR',(req,res)=>{
    excelmodel.find(
        {R:req.params.RR}
    ).then((r)=>{
        res.send(r)
    })
})

app.get('/gett/:NV',(req,res)=>{
    excelmodel.find(
        {'niveau R':req.params.NV}
    ).then((read)=>{
        res.send(read)
    })
})

app.get('/get',((req,res)=>{
    excelmodel.find({}).then((re)=>{
        res.send(re)
    })
}))


app.post('/add',(req,res)=>{
    let data=req.body[0]

    let excell = new excelmodel({
        'Excercice':req.body[0],
        'N° LF':req.body[1],
        'niveau R':req.body[2],
        'R':req.body[3],
        'lib':req.body[4],
        'type R':req.body[5],
        'code fonctionnel':req.body[6],
        'Code Eco':req.body[7],
        'CP':req.body[8],
        'CE':req.body[9],
        'DEP':req.body[10],
    })
    excell.save().then(post => {
        if(post){
            res.status(201).json({
                message: "Post added successfully",
                post: {
                    ...post,
                    id: post._id
                }
            })
        }
}).catch(e => {
        console.log(e)
    })
})

app.get('/delete/:DV',(req,res)=>{
    excelmodel.findOneAndDelete(
        {R:req.params.DV}
    ).then((r)=>{
        res.send(r)
    })
})

app.post('/update',(req,res)=>{

    excelmodel.updateMany(
        {R:req.body[3]},
        {
            'Excercice':req.body[0],
            'N° LF':req.body[1],
            'niveau R':req.body[2],
            'R':req.body[3],
            'lib':req.body[4],
            'type R':req.body[5],
            'code fonctionnel':req.body[6],
            'Code Eco':req.body[7],
            'CP':req.body[8],
            'CE':req.body[9],
            'DEP':req.body[10],
        }
    ).then((r)=>{
        res.send(r)
    })
})

app.get('/search/:S',((req,res)=>{
    let text=req.params.S
    excelmodel.find( 
        {$or:[
            {'lib':{$regex:text}}
        ,
            { 'R':{$regex:text}}
        ]}
    ).then((re)=>{
        res.send(re)
    })
}))

app.get('/deletetable',((req,res)=>{
    excelmodel.remove({}).then((re)=>{
        res.send(re)
    })
}))

module.exports=app;