
var FilePath      = "C:/Users/11011934/Desktop/321/rm_km_cib_sense/ExcelWebEditor/Files"; // путь к файлу               
var ServCert      = '/Certificats/sbt-ouiefs-0105_sigma_sbrf_ru.pfx';   // путь к файлу с сертификатом сервера
var ServicePort   = '558';                                              // Порт по которому доступен сервис

const XLSX = require(__dirname +'/node-modules/xlsx.full.min.js');
const https = require('https');
const fs = require('fs');
const options = {
  pfx: fs.readFileSync(__dirname + ServCert)
};

var files = fs.readdirSync(FilePath);
console.log(files);
var FileName      = files[0];     // имя файла, который будет открываться по умолчанию
FileName          = FilePath + '/' + FileName;
console.log(FileName);
const workbook = XLSX.readFile(FileName);
list_name = workbook.SheetNames[0];
//list_name = null;

function find(arr, elem){
    for (var i = 0; i < arr.length; i++){
        if (elem == arr[i]) {
            return i;
        }
    }
}

function run(FileName, sheet_name) {  
    const workbook = XLSX.readFile(FileName);
    sheet_name_list = workbook.SheetNames;
    name = FileName.split('/')[FileName.split('/').length - 1];
    file_index = find(files, name);
    list_index = find(sheet_name_list, sheet_name);

    console.log('Открыт Excel файл: ' + name);
    console.log('Рабочий лист: ' + sheet_name);

    if (sheet_name != null) {
        ExcelSheetJSON = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name]);
        var arrHead = XLSX.utils.get_fields(workbook.Sheets[sheet_name]);
    } else {
        ExcelSheetJSON = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
        var arrHead = XLSX.utils.get_fields(workbook.Sheets[sheet_name_list[0]]);
    }
    //Парсим и собираем таблицу вновь.
    //Это заголовки полей: USERID  AppRole FieldMachValue  PersonalNumber
    //Создаём содержимое HTML страницы
    //Заголовки страницы    
    var HTMLContent =   '<!doctype html>\n' +
                    '  <html lang="en">\n' +
                    '  <head>\n' +
                    '    <meta charset="utf-8">\n' +
                    '    <title>Editable table</title>\n' +

                    '    <link rel="stylesheet" href="style.css">\n' +
                    '    <link rel="stylesheet" href="pure-min.css">\n' +
                    
                    '    <link rel="shortcut icon" href="img\\favicon_excel.ico" type="image/x-icon">\n' +
                    '    <style>\n' +
                    '      body{ padding: 1% 3%; color: rgb(119, 119, 119); }\n' +
                    '      #file'+file_index+'{ background: #51CE0E}\n' +
                    '      h1{ color:#333 }\n' +
                    
                    '      #list'+list_index+'{ background: #51CE0E}\n' +
                    '      h1{ color:#333 }\n' +
                    '    </style>\n' +
                    
                    '  </head>\n' +
                    '<body>\n' +
                    '<div id="info">\n' +
                        'Веб-редактор xlsx файлов\n' +   
                        '<label id="help">\n' +
                            'Помощь\n' +   
                            '<div id="pop">\n' +
                                '<p>Чтобы редактировать ячейки, нужно на них нажать</p>\n' +
                                '<p>Чтобы добавить строку, нужно нажать на "Добавить"</p>\n' +
                                '<p>Чтобы сохранить изменения, нужно нажать на "Сохранить"</p>\n' +
                                '<p>Несохраненные изменения не фиксируются</p>\n' +
                                '<p>Число листов в документе не должно быть больше десяти</p>\n' +
                            '</div>\n' +  
                        '</label>\n' +          
                    '</div>\n' +
                    '<br>\n' +
                    '<span class="list">Файлы</span>\n'
    //добавляем список листов
    for (var i = 0; i < files.length; i++) {
        
        HTMLContent += '          <button name="file'+i+'" onclick="file'+i+'()" value="'+ files[i] +'" class="files" id="file'+i+'">'+ files[i] +'</button>\n'
    }                
    
    HTMLContent +=  '<br>\n' +
                    '<span class="list">Листы</span>\n'
    //добавляем кнопку переключения листа
    for (var i = 0; i < sheet_name_list.length; i++) {
        
        HTMLContent += '          <button name="list'+i+'" onclick="list'+i+'()" value="'+ sheet_name_list[i] +'" class="lists" id="list'+i+'">'+ sheet_name_list[i] +'</button>\n'
    }


    HTMLContent +=  '<br>\n' +
                    '<button name="send" onclick="isSend()" class="func">Сохранить</button>\n' +
                    '<button name="send" onclick="isNew()" class="func">Добавить</button>\n' +
                    //Заголовки таблицы
                    '<table id="editable" class="pure-table pure-table-bordered">\n' +
                    '  <thead>\n' +
                    '      <tr>\n'

    for (var i = 0; i < arrHead.length; i++) {
        HTMLContent += '          <th style="text-align: left;">'+ arrHead[i] +'</th>\n'
    }

    HTMLContent +=  '      </tr>\n' +
                    '  </thead>\n' +
                    '  <tbody>'
                    ;
    //Содержимое таблицы
    for (var i = 0; i < ExcelSheetJSON.length; i++) {
        var columnArray = [];
        for (var j = 0; j < arrHead.length; j++) {
            columnArray.push(ExcelSheetJSON[i][arrHead[j]]);
            columnArray.forEach(function(item, i, columnArray) {
                if ((columnArray[i] == undefined) || (columnArray[i] == 'undefined')) {                
                    columnArray[i] = '';
                }
            });
        }

    HTMLContent +=  '      <tr>\n'

    for (var j = 0; j < columnArray.length; j++) {
        HTMLContent += '   <td>' + columnArray[j] + '</td>\n'
    }

    HTMLContent +=  '      </tr>\n'
                    ;
    }      

    //Конец страницы
    HTMLContent +=  '  </tbody>\n' +
                    '</table>\n' +

                    //'<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.2.2/jquery.min.js"></script>\n' +
                    '<script type="text/jscript" src="require.js"></script>\n' +
                    '<script type="text/jscript" src="editableTableWidget.js"></script>\n' +
                    '<script type="text/jscript" src="xlsx\\dist\\xlsx.extendscript.js"></script>\n' +
                    
                    
                    '<script>\n' +    
                    '  $(\'#editable\').editableTableWidget();\n' +
                    '  $(\'#editable td.uneditable\').on(\'change\', function(evt, newValue) {\n' +
                    '    return false;\n' +
                    '  });\n' +
                    '</script>\n' +
                    '<script type="text/jscript" src="XLSExport.js"></script>\n' +
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
    app.use(express.static(__dirname + '/node-modules'));
    var bodyParser = require(__dirname + '/node-modules/express/node_modules/body-parser/index.js')
    // parse application/x-www-form-urlencoded 
    app.use(bodyParser.urlencoded({ extended: false }))
    app.post('/', function(req, res) {
        var wb = XLSX.utils.book_new();
        wb = XLSX.readFile(FileName);
        if (req.body.arrObjects != undefined) {
            arrObjects = JSON.parse(req.body.arrObjects);
            wb.Sheets[list_name] = arrObjects;
            XLSX.writeFile(wb, FileName);
            console.log('Обновлен файл XLSX: ' + FileName);
            console.log('Новый лист Excel: ' + list_name);
            console.log('');
            HTMLContent = run(FileName, list_name);
            
        } else if (req.body.list_name != undefined) {
            list_name = req.body.list_name;   
            HTMLContent = run(FileName, list_name); 
            show(HTMLContent);
        } else {
            FileName = req.body.file_name; 
            FileName = FilePath + '/' + FileName;
            const workbook = XLSX.readFile(FileName);
            
            list_name = workbook.SheetNames[0];
            console.log(FileName);
            HTMLContent = run(FileName, list_name); 
            show(HTMLContent);
        }
        res.send(HTMLContent);
    });
}

var HTMLContent = run(FileName, list_name);
show(HTMLContent);


https.createServer(options, app).listen(ServicePort);
console.log('Сервис запущен. Порт: ' + ServicePort);