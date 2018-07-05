
var FileName      = "НастройкаДоступаРМКМ.xlsx";                        // имя файла, который будет подвергнут изменениям (он будет перезаписан)
var WorkSheetName = "НастройкаДоступаРМКМ";                             // имя листа в НОВОМ файле XLSX
var ServCert      = '/Certificats/sbt-ouiefs-0105_sigma_sbrf_ru.pfx';   // путь к файлу с сертификатом сервера
var ServicePort   = '558';                                              // Порт по которому доступен сервис


const XLSX = require(__dirname +'/node-modules/xlsx.full.min.js');
const https = require('https');
const fs = require('fs');
const options = {
  pfx: fs.readFileSync(__dirname + ServCert)
};

//Имя файла

function run(FileName) {  const workbook = XLSX.readFile(FileName);
    const sheet_name_list = workbook.SheetNames;

    console.log('Открыт Excel файл: ' + FileName);
    console.log('Рабочий лист: ' + sheet_name_list);


    ExcelSheetJSON = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]])
    //Парсим и собираем таблицу вновь.
    //Это заголовки полей: USERID  AppRole FieldMachValue  PersonalNumber
    //Создаём содержимое HTML страницы
    //Заголовки страницы    
    var HTMLContent =   '<!doctype html>\n' +
                    '  <html lang="en">\n' +
                    '  <head>\n' +
                    '    <meta charset="utf-8">\n' +
                    '    <title>Editable table</title>\n' +
                    '    <link rel="stylesheet" href="pure-min.css">\n' +
                    '    <link rel="shortcut icon" href="img\\favicon_excel.ico" type="image/x-icon">\n' +
                    '    <style>\n' +
                    '      body{ padding: 1% 3%; color: rgb(119, 119, 119); }\n' +
                    '      h1{ color:#333 }\n' +
                    '    </style>\n' +
                    
                    '  </head>\n' +
                    '<body>\n' +
                    '<h1>НастройкаДоступаРМКМ.xlsx</h1>\n' +
                    '<button name="send" onclick="isSend()">Сохранить</button>\n' +
                    //Заголовки таблицы
                    '<table id="editable" class="pure-table pure-table-bordered">\n' +
                    '  <thead>\n' +
                    '      <tr>\n' +
                    '          <th style="text-align: left;">USERID</th>\n' +
                    '          <th style="text-align: left;">AppRole</th>\n' +
                    '          <th style="text-align: left;">FieldMatchValue</th>\n' +
                    '          <th style="text-align: left;">PersonalNumber</th>\n' +
                    '      </tr>\n' +
                    '  </thead>\n' +
                    '  <tbody>'
                    ;
    //Содержимое таблицы
    for (var i = 0; i < ExcelSheetJSON.length; i++) {
        Column_1 = ExcelSheetJSON[i].USERID
        Column_2 = ExcelSheetJSON[i].AppRole
        Column_3 = ExcelSheetJSON[i].FieldMatchValue
        Column_4 = ExcelSheetJSON[i].PersonalNumber
        var columnArray = [Column_1, Column_2, Column_3, Column_4];
        
        columnArray.forEach(function(item, i, columnArray) {
            if (columnArray[i] == 'undefined') {
                columnArray[i] = '';
            }
        });

    HTMLContent +=  '      <tr>\n' +
                    '          <td>' + columnArray[0] + '</td>\n' +
                    '          <td>' + columnArray[1] + '</td>\n' +
                    '          <td>' + columnArray[2] + '</td>\n' +
                    '          <td>' + columnArray[3] + '</td>\n' +
                    '      </tr>\n'
                    ;
    }                
    //Конец страницы
    HTMLContent +=  '  </tbody>\n' +
                    '</table>\n' +
                    //'<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.2.2/jquery.min.js"></script>\n' +
                    '   <script type="text/jscript" src="require.js"></script>\n' +
                    '<script type="text/jscript" src="editableTableWidget.js"></script>\n' +
                    '<script type="text/jscript" src="xlsx\\dist\\xlsx.extendscript.js"></script>\n' +
                    '<script type="text/jscript" src="XLSExport.js"></script>\n' +
                    
                    '<script>\n' +    
                    '  $(\'#editable\').editableTableWidget();\n' +
                    '  $(\'#editable td.uneditable\').on(\'change\', function(evt, newValue) {\n' +
                    '    return false;\n' +
                    '  });\n' +
                    '</script>\n' +

                    '</body>\n' +
                    '</html>'
                    ;
    return HTMLContent;
}
//Создаём хостинг для файлов сайта
var express = require(__dirname + '/node-modules/express/index.js');
var app = express();
// Serve up content from directory

function show(HTMLContent) {
    app.get('',
       function(Request, Response){
          //Response.writeHead(200, {"Content-Type": "text/plain; charset=utf-8"});
          Response.send(HTMLContent);
          console.log('Передано содержимое в HTML страницу');
    });
}

var HTMLContent = run(FileName);
show(HTMLContent);

app.use(express.static(__dirname + '/node-modules'));

var bodyParser = require(__dirname + '/node-modules/express/node_modules/body-parser/index.js')
// parse application/x-www-form-urlencoded 
app.use(bodyParser.urlencoded({ extended: false }))
app.post('/', function(req, res) {
    var wb = XLSX.utils.book_new();
    arrObjects = JSON.parse(req.body.arrObjects);
    XLSX.utils.book_append_sheet(wb, arrObjects, WorkSheetName);
    
    XLSX.writeFile(wb, FileName);
    console.log('Обновлен файл XLSX: ' + FileName);
    console.log('Новый лист Excel: ' + WorkSheetName);
});


https.createServer(options, app).listen(ServicePort);
console.log('Сервис запущен. Порт: ' + ServicePort);