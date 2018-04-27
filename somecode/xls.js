//Бот для парсинга входящего XLS-файла с примерно 20 000 строк, формирования отчета (сколько времени человек работал каждый день, сколько было перерыво, общая длина перерывов) по выбранным отделам и отправка пользователю сформированного отчета в XLS формате.

var xlsx = require("node-xlsx").default;
const fs = require('fs');

var tg = require('telegram-node-bot')('###');
var request = require("request");

var xl = require('excel4node');

var headerStyle={
	alignment: {
		horizontal: 'center'
	},
	fill: {
	    type: 'pattern',
	    patternType: 'darkUp',
	    fgColor: 'ceeeff',
	    bgColor: 'ceeeff'
    },
    font: {
     	bold: true
    }
}
var goodStyle={
	alignment: {
		horizontal: 'center'
	},
	fill: {
	    type: 'pattern',
	    patternType: 'darkUp',
	    fgColor: '8cfc8c',
	    bgColor: '8cfc8c'
    },
    font: {
     	bold: true
    }
}
var badStyle={
	alignment: {
		horizontal: 'center'
	},
	fill: {
    	type: 'pattern',
    	patternType: 'darkUp',
    	fgColor: 'fea08a',
    	bgColor: 'fea08a'
    },
    font: {
     	bold: true
    }
}
var infoStyle={
	alignment: {
		horizontal: 'center'
	},
	fill: {
	    type: 'pattern',
	    patternType: 'darkUp',
	    fgColor: 'fbb0f3',
	    bgColor: 'fbb0f3'
    },
    font: {
     	bold: true
    }
}

var u='';

tg.router.
    when(['/get'], 'Get').
    otherwise('OtherwiseController')

tg.controller('OtherwiseController', ($) => {
    //console.log($.message)
    if($.message.document){ //если пользователем отправлен документ
    	console.log('[LM bot]: Новый запрос от: '+$.message.from.username)

		u=$.message.from.id; //сохраняем id пользователя
    	var format=$.message.document.file_name.split('.') //узнаем формат файла
    	if(format[1]=='xls' || format[1]=='xlsx' || $.message.document.mime_type=='application/vnd.ms-excel'){ //если верный формат
			tg.getFile($.message.document.file_id,($) => { //получаем файл с сервера
				var url='https://api.telegram.org/file/bot474801412:AAG9HRmGBtUUQcZCqrFDXesN8Pdfxr3_rwE/'+$.result.file_path //формируем адрес для скачивания
				//console.log($.result)
				tg.sendMessage(u, 'Анализирую данные'); //уведомляем пользователя о работе над фалом
				request(url).pipe(fs.createWriteStream('ExcelIn.xls')).on("finish", function (){ //скачиваем файл
					//tg.sendMessage(u, 'Анализирую данные');
					console.log('[LM bot]: Журнал получен');
					readXSL(u); //отправляем на анализ после получения
				})
			})
		} else {
			console.log('[LM bot]: Пользователь оптравил неверный формат файла: '+$.message.document.mime_type)
			tg.sendMessage(u, 'Неверный формат файла.\nЖурнал должен быть в формате xls или xlsx'); //уведомляем пользователя о неверном формате файла
		}
    }
})

tg.controller('Get', ($) => {
	console.log('[LM bot]: проверяю наличие обработанного журнала')
	if(xlsDate) {
		console.log('[LM bot]: Доступен журнал учета '+xlsDate)
		//$.sendMessage('Доступен журнал учета '+xlsDate);
		SectionButtons($.message.from.id)
	} else {
		console.log('[LM bot]: Нет обработанных журналов')
		$.sendMessage('Нет обработанных журналов. Загрузите новый');
	}
})

tg.callbackQueries(($) => { //ждем выбор пользователя
	console.log($.data);
	
	if(xlsDate){
		tg.editMessageText('Детальный отчет по отделу "'+titlekeys[$.data]+'"',{message_id: $.message.message_id,chat_id: $.message.chat.id})
		SectionSort(titlekeys[$.data], $.message.chat.id) //отправляем запрос на создание отчета по выбранному разделу
	}
})

console.log('[LM bot]: Бот запущен')

var xlsDate; //дата дляформирования отчета
var titlekeys; //названия отделов отделов
var groupBySection; //массив отделов

function readXSL(tgUser){

	console.log('[LM bot]: Анализирую данные')
	const efile = xlsx.parse(`${__dirname}/ExcelIn.xls`); // Parse a file
	var data=efile[0].data //read data
	xlsDate=data[0][0].substring(15, data[0][0].length) //Парсим дату отчета
	xlsDate=xlsDate.replace(/\./g, '-') //форматируем дату отчета в читабельный вид
	
	var headers={};
	data[2].map(function(n,i){ //находим номера столбцов и заголовков 
		headers[n]=i //создаем массив заголовков
	})
	console.log(headers);
	var sort=[];//массив всех строк событий 
	
	data.map(function(n,i){ //фультруем нужные данные 
		if(i>2){
			let name = n[headers['Имя']]+' '+n[headers['Фамилия']]; //формирование имени
			let readDate;
			let readTime;
			if(n[headers['Дата и время']]){
				var getFullDate=''+n[headers['Дата и время']] //формирование даты события
				var parseDate=getFullDate.split('.')
				var getTime='0.'+parseDate[1]
				//console.log(parseDate);
				readDate=parseDate[0]-1;
				readTime=Math.round((Number(getTime).toPrecision(11)*100000000000).toFixed(0)/1157410) //формирование времени события
				//console.log(readTime);
			}
			if(n[headers['Дата']] && n[headers['Время']]){
				readDate=n[headers['Дата']]-1 //формирование даты события
				readTime=Math.round((n[headers['Время']].toPrecision(11)*100000000000).toFixed(0)/1157410) //формирование времени события
			}
			let normalDate=new Date(1900, 0, readDate, 0, 0, readTime) //нормальная дата
			let shortDate=normalDate.toLocaleDateString() //короткая дата
			let printDate=normalDate.toLocaleString('ru-RU', {day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit', second:'2-digit'})
			let unixTime=normalDate.getTime()/1000 //время Unix формата
			//let userId=n[headers['№ карты']].replace(/\s/g, ''); //номер карты
			let userId; //номер карты
			if(n[headers['№ карты']]){ //если есть
				userId=n[headers['№ карты']].replace(/\s/g,'') //добавляем
			} else { //если нет
				userId=''; //игнорируем
			}
			let section=n[headers['Отдел']]; //узнаем отдел
			let move=n[headers['Объект']];
			let strDate={'section': section, 'date':shortDate, 'userid':userId, 'name':name, 'move':move, 'unix':unixTime, 'print':printDate} //формируем данные для отправки в массив
			//console.log(strDate)
			sort.unshift(strDate); //добавляем в массив
		}
		
	})
	//console.log(section)
	
	groupBySection = groupBy(sort, 'section'); //группируем по отделам
	console.log('[LM bot]: Отделы сгруппированы')
	//SectionSort('скобряные изделия', tgUser) //формируем запрос на отчет напрямую
	
	titlekeys = Object.keys(groupBySection); //получаем название отделов для кнопок
	
	SectionButtons(tgUser)
}

function SectionButtons(user){
	var buttons=[]; //массив кнопок отделов
	for(var i=0;i<titlekeys.length;i++){ //создаем кнопки
		if(titlekeys[i]!='undefined'){ //игнорируем ненужные
			buttons.push([{text:titlekeys[i], callback_data:i}]) //добавляем кнопки
		}
	}
	//console.log(titlekeys)
	tg.sendMessage(user, 'Выберите отдел для детального отчета '+xlsDate, { //отправляем кнопки отделов пользователю
        reply_markup: JSON.stringify({
        	inline_keyboard: buttons
        })
    });
}

function SectionSort(inDate, usr){ //формируем отчет по выбранному отделу
	
	var wb = new xl.Workbook(); //создаем книгу
	
	console.log('[LM bot]: Формирую отчет по отделу: '+inDate)
	//console.log(inDate)
	
	var groubedById=groupBy(groupBySection[inDate], 'userid') //группируем по пользователям
	
		for(u in groubedById){
		var ws = wb.addWorksheet(groubedById[u][0].name); //создаем листы для каждого опльзователя

		var stringCount=1; //Счетчик строк для xlsx файла

		ws.column(1).setWidth(15); //указываем ширину столбца №1
		ws.column(2).setWidth(20); //указываем ширину столбца №2
		ws.column(3).setWidth(15); //указываем ширину столбца №3

		var groupByDate=groupBy(groubedById[u],'date') //групперуем события по дате у каждого пользователя

		for(x in groupByDate){ // Перебор всех дат
			//console.log('Дата: '+x)
			var start=0; // начало дня
			var end=0; // конец дня
			var rel=0; // длительность отдыха
			var total=0; // длительность всего дня с отдыхом
			var relArray=[]; // длительность каждого отдыха
			var thisrel=0; // текущая длительность отдыха
			var tArray=[]; // массив временных интервалов отдыха
			var t=[]; // массив текущего отдыха
			
			groupByDate[x].forEach(function (value, i) { //перебор конкретной даты
			    if(i==0) start=value.unix // узнаем начало рабочего дня
			    if(i==groupByDate[x].length-1) end=value.unix // узнаем конец рабочего дня
			    if(value.move=='Конец работы' && i!=groupByDate[x].length-1){ //начало перерыва
			    	rel+=value.unix;
			    	thisrel=value.unix;
			    	t.push(timeConverter(value.unix))
			    }
			    if(value.move=='Начало работы' && i!=0){ //конец перерыва
			    	rel=rel-value.unix
			    	thisrel=thisrel-value.unix
			    	relArray.push(TimeCounter(thisrel))
			    	t.push(timeConverter(value.unix))
			    	tArray.push(t) //отправляем в массив интервал отдыха
			    	t=[]; // очищаем текущий период отдыха
			    }
			});
			total=end-start; //вычисляем общее время работы с учетом отдыха

			//формируем строки в xlsx файл
			if(Math.abs(rel)>3600){
				ws.cell(stringCount, 1).string(dateConverter(x)).style(badStyle); //добавляем дату
				ws.cell(stringCount, 2, stringCount, 3, true).string('c '+timeConverter(start)+' до '+timeConverter(end)).style(badStyle); // указываем время работы
			} else if(Math.abs(rel)<2700){
				ws.cell(stringCount, 1).string(dateConverter(x)).style(infoStyle); //добавляем дату
				ws.cell(stringCount, 2, stringCount, 3, true).string('c '+timeConverter(start)+' до '+timeConverter(end)).style(infoStyle);
			} else {
				ws.cell(stringCount, 1).string(dateConverter(x)).style(goodStyle); //добавляем дату
				ws.cell(stringCount, 2, stringCount, 3, true).string('c '+timeConverter(start)+' до '+timeConverter(end)).style(goodStyle);
			}
			
			stringCount++
			ws.cell(stringCount,1,stringCount,2,true).string('Рабочий день').style({font: {bold: true},alignment: {horizontal: 'center'}});
			ws.cell(stringCount,3).string(TimeCounter(total)).style({alignment: {horizontal: 'center'}});
			stringCount++
			ws.cell(stringCount,1,stringCount,2,true).string('Перерывы').style({font: {bold: true},alignment: {horizontal: 'center'}});
			if(Math.abs(rel)>3600){
				ws.cell(stringCount,3).string(TimeCounter(rel)).style(badStyle);
			} else if(Math.abs(rel)<2700){
				ws.cell(stringCount,3).string(TimeCounter(rel)).style(infoStyle);
			} else {
				ws.cell(stringCount,3).string(TimeCounter(rel)).style(goodStyle);
			}
			stringCount++
			tArray.forEach(function(v,i){ //генерируем строки перерыва
				ws.cell(stringCount,1).string('#'+(i+1)).style({font: {bold: true},alignment: {horizontal: 'center'}});
				ws.cell(stringCount,2).string(v[0]+' - '+v[1]).style({font: {bold: true},alignment: {horizontal: 'center'}});
				ws.cell(stringCount,3).string(relArray[i]).style({alignment: {horizontal: 'center'}});
				stringCount++
			})

			stringCount++
		}
	}
	console.log('[LM bot]: Сохраняю отчет в файл')
	
	wb.write('ExcelOut.xlsx', function (err, stats) { //записываем строки в файл
		if (err) {
			console.log('[LM bot]: Ошибка сохранения файла'); // Prints out an instance of a node.js fs.Stats object
			console.error(err);
		} else {
			console.log('[LM bot]: Отправляю отчет пользователю'); // Prints out an instance of a node.js fs.Stats object
			//отправляем файл пользователю
			var doc =  {
			    value: fs.createReadStream('ExcelOut.xlsx'), //stream
			    //filename: 'Детальный отчет ('+new Date().toLocaleDateString()+').xlsx',
			    filename: 'Отчет ('+xlsDate+').xlsx',
			    contentType: 'application/vnd.ms-excel'
			}
			tg.sendDocument(usr,doc) //отправляем сформированный документ пользователю
			console.log('[LM bot]: Отчет отправлен');
		}
	});
	
}

function groupBy(xs, key) { //функция перегруппирования по ключу
	return xs.reduce(function(rv, x) {
		(rv[x[key]] = rv[x[key]] || []).push(x);
		return rv;
	}, {});
};

function TimeCounter(val){ //перевод времени из Unix в часы, минуты, секунды
	var t = Math.abs(val);
	var days = parseInt(t/86400);
	t = t-(days*86400);
	var hours = parseInt(t/3600);
	t = t-(hours*3600);
	var minutes = parseInt(t/60);
	t = t-(minutes*60);
	if(minutes<10) minutes='0'+minutes;
	if(t<10) t='0'+t;
	var content = "";
	if(days)content+=days+" дня";
	if(hours||days){ if(content)content+=", "; content+=hours+":"; }
	if(content)content+=""; content+=minutes+":"+t;
	return content;
}

function timeConverter(UNIX_timestamp){ //отображение понятного времени
  var a = new Date(UNIX_timestamp * 1000);
  var hour = a.getHours();
  if(hour<10) hour='0'+hour
  var min = a.getMinutes();
  if(min<10) min='0'+min
  var sec = a.getSeconds();
  if(sec<10) sec='0'+sec
  var time = hour + ':' + min + ':' + sec ;
  return time;
}

function dateConverter(unix){ //отображение понятной даты
  var a = unix.split('-');
  var months = ['','Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];
  var year = a[0];
  var month = months[a[1]];
  var date = a[2];
  var time = date + ' ' + month + ' ' + year;
  return time;
}