const express = require('express');
const app = express();
const { stringify } = require('querystring');
const cors = require('cors')
const bodyParser = require('body-parser')
const path = require('path');
const routs = require('./routes/routs')

app.listen(3000,(req,res)=>{
    console.log("server connected")
})


app.use('/apis',routs)
app.use(cors());
app.use(express.static(path.join(__dirname,'dist/excel')));
app.use('/*',function(req,res){
    res.sendFile(path.join(__dirname+'/dist/excel/index.html'))
})
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({
  extended: false
}));
app.use(function(req, res, next) {
    res.header("Access-Control-Allow-Origin", "*"); 
    res.header("Access-Control-Allow-Methods","GET, POST, HEAD, OPTIONS, PUT, PATCH, DELETE");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    next();
});



