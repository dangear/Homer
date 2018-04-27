//Бот для автоматического поиска фильмов в БД сервиса Kinoplan24.ru и добавления анонса в ВК группу кинотеатра

var request = require("request");
var tg = require('telegram-node-bot')('###########################')

const fs    = require('fs');
const VKApi = require('node-vkapi');
const VK    = new VKApi({
  app: {
    id: ##########,
    secret: '#############'
  }, 
  auth: {
    login: '+##########', 
    pass: '###########'
  },
});

const g_id='##########';

var filmsArray=[];

tg.router.
    when(['/a'], 'Add').
    when(['/f'], 'Add').
    otherwise('OtherwiseController')

tg.controller('Add', ($) => {
  if ($.user.id == '#############') {
    tg.for('/a', () => {
      if($.args != '') {
        var f=$.args.split(',')
        filmsArray.push({title: f[0], format: f[1].trim(), week: f[2].trim()})
        getToken();
      } else {
        $.sendMessage('Для того, чтобы добавить фильм, используйте формат типа /a название, формат, длительность проката')
      }
    })

    tg.for('/f', () => {
      if(filmsArray.length>0){
        var str=''
        filmsArray.map(function(name){
          str+=name.title+' ('+name.format+')\n'
        })
        $.sendMessage('Список фильмов на добавление:\n'+str, { parse_mode: 'Markdown'})
      } else {
        $.sendMessage('Список пуст')
      }
    })

})

tg.controller('OtherwiseController', ($) => {    
    tg.for('ping', ($) => {
      $.sendMessage('pong')      
    })
})

// VK wall post START
function wallpost(mov){
  var mov_post={
    cover: '',
    video: '' 
  }
  
  console.log('Start add release: '+mov.title)
  request(mov.cover).pipe(fs.createWriteStream(mov.title+'.jpg')); //get cover
  request(mov.video).pipe(fs.createWriteStream(mov.title+'.mov')).on("finish", function (){
    console.log('Трейлер '+mov.title+' скачан')
    return VK.auth.user({
      scope: ['photos', 'wall', 'offline', 'video']
    }).then(token => {
      // Load Cover
      return VK.upload('photo_wall', {
        data: fs.createReadStream(mov.title+'.jpg'),
        beforeUpload: {
          group_id: g_id
        } 
      })
      .then(r => {
        mov_post.cover='photo' +r[0].owner_id+'_' + r[0].id
        var y=mov.rel_date.split('-');
        
        // Load video
        return VK.upload('video', {
          data: fs.createReadStream(mov.title+'.mov'),
          beforeUpload: {
            group_id: g_id,
            name: mov.title+' ('+y[0]+') - Трейлер HD',
            description: 'С '+dateConv(mov.rel_date, mov.end_date)+' в кинотеатре "Салют"',
            no_comments: 1
            }
        })
        .then(r => {
          console.log('Трейлер '+mov.title+'загружен на сервер VK')
          mov_post.video='video-' +g_id+'_' + r.video_id
          // Post to the wall
          return VK.call('wall.post', {
            owner_id: '-'+g_id,
            from_group: 1,
            publish_date: (Date.parse(mov.rel_date)/1000)-(86400*Math.floor(Math.random() * 3 + 4))+(600*Math.floor(Math.random() * 24 + 24)),
            attachments: mov_post.cover+','+mov_post.video,
            message: mov.message
          }).then(res => {
            // wall.post response
            fs.unlinkSync(mov.title+'.mov');
            fs.unlinkSync(mov.title+'.jpg');
            
            console.log(mov.title+': https://vk.com/wall-' + g_id + '_' + res.post_id);
            tg.sendMessage(74341139, 'Фильм '+mov.title+' опубликован: https://vk.com/wall-'+g_id + '_' + res.post_id);
            filmsArray=[];
          });
        });
      })
    }).catch(error => {
      // catching errors 
      console.log(error);
    });  
  });
}
// wall post END

// Get films info START
var baseDir = '################'
var kinoKey = '################'
var kinoToken = ''
		
function getToken(){
	var expDate = ''
	request({
			url: 'https://ts.kinoplan24.ru/api/auth/token?api_key='+kinoKey,
			json: true
		},
		function (error, response, body) {
				kinoToken = body.request_token;
				console.log(kinoToken);
				filmsFind();
		})
}

function filmsFind(){
  filmsArray.map(function(name,i){
    findRelease(name.title,name.format,name.week);
  })
}

function findRelease(name, format, week){
	request({
			url: 'http://ts.kinoplan24.ru/api/release/search?title='+encodeURIComponent(name),
			headers: {
				'REQUEST-TOKEN': kinoToken
			},
			json: true
			}, function (error, response, body) {
        //console.log(body)
        if(body<=0){
          console.log('Nothing to find');
          tg.sendMessage(74341139,'Релиз "'+name+'" не найден', { parse_mode: 'Markdown'});
          filmsArray=[];
          return false;
        } else {
      			var dt=new Date().toLocaleDateString('ru-RU', {year:'numeric', month:'2-digit', day:'2-digit'})
    				for (var i = 0; i < body.length; i++) {
      				if(body[i].date>dt) {
                fullInfo(body[i].id, format, week)
                break;
              }
    				}
        }
			})
}

function fullInfo(id,format,week){
	request({
			url: 'https://ts.kinoplan24.ru/api/release/'+id+'/full',
			headers: {
				'REQUEST-TOKEN': kinoToken
			},
			json: true
			}, function (error, response, body) {

        var enddate=new Date((Date.parse(body.date.russia.start)/1000+(604800*week-86400))*1000).toISOString().replace(/^([^T]+)T(.+)$/,'$1')
        //console.log(body)
  			if(body.duration.clean==0) body.duration.clean=100

        if (body.trailers.length <= 0) {
          console.log('Trailers is empty')
          tg.sendMessage(74341139, 'У релиза "'+body.title.ru+'" еще нет трейлера. Попробуйте позже');
          filmsArray=[];
          return false;
        } else {
          //console.log(body.trailers)
        }

        for (var i = 0; i < body.trailers.length; i++) {
          if(body.trailers[i].duration>30 && body.trailers[i].formats[0]=='2D') {
            var tr=body.trailers[i].hd720
            break;
          } else {
            var tr=body.trailers[0].hd720
          }
        }
        //console.log(tr)
  			var string = '&#127916; '+toUp(body.genres)+' "'+body.title.ru+'" '+'('+format+') \n'+
  			        '&#128197; С '+dateConv(body.date.russia.start, enddate)+' в @blagcinema (кинотеатре "Салют") \n'+
  			        '&#9203; Длительность: '+body.duration.clean+' мин \n'+
  			        '&#9888; Возрастные рекомендации: '+body.rate+' \n'+
  			        ' \n'+
  			        body.description+' \n'+
  			        ' \n'+
  			        'Подпишись на @blagcinema (нашу группу) и будь в курсе последних новинок кинопроката. \n'+
  			        'Наш адрес: г.Благовещенск, ул.Кирова,5 \n'+
                'Телефон: 8(34766)2-26-36 \n'+
                '#киновблаге #салютжив #кинотеатрсалют'       
				var mov_info = {
  				title: body.title.ru,
  				message: string,
  				cover: body.cover,
  				video: tr,
  				rel_date: body.date.russia.start,
          end_date: enddate
  		  }
  		  console.log('start to add release');
        tg.sendMessage(74341139,'Релиз "'+body.title.ru+'" найден.\nИдет добавление релиза', { parse_mode: 'Markdown'})
  		  wallpost(mov_info);
        filmsArray=[];
			})
}
// Get films info END

// Other functions
function toUp(str) {
  return str[0].substr(0, 1).toUpperCase()+str[0].substr(1);
}

function dateConv(start,end){
  var m=['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декакбря'];
  var s=start.split("-");
  var e=end.split("-");
  if(s[1]==e[1]){
    if(s[1]<10){
      s[1]=s[1].substring(1)
    }
    var str=s[2]+'-го по '+e[2]+'-ое '+m[s[1]-1]
  } else {
    if(s[1]<10){
      s[1]=s[1].substring(1)
    }
    if(e[1]<10){
      e[1]=e[1].substring(1)
    }
    var str=s[2]+'-го '+m[s[1]-1]+' по '+e[2]+'-ое '+m[e[1]-1]
  }
  return str
}
