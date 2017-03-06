var http = require('http');
var cheerio = require('cheerio');
var officegen = require('officegen');
var docx = officegen({'type':'docx'});
var fs = require('fs');
var path = require('path');
var async = require('async');
var srcReg = /original=[\'\"]?([^\'\"]*)[\'\"]?/gi;

function filterChapters(html){
	var $ = cheerio.load(html,{decodeEntities: false});
	var chapters = $('.atl-item');

	var data = [];
	chapters.each(function(item){
		var d = $(this);
		var louzhu = d.find('.atl-info strong').text();
		var time = d.find('.atl-info span').eq(1).text();
		var content = d.find('.bbs-content').html();
		if(louzhu == '楼主'){
			var everyPageData = {
				time: time,
				content: content
			}
		}else{
			return;
		}
		data.push(everyPageData);
	})
	return data;
}

//输出word的格式
function printData(data,pageNum){
	var everyPageData = [];
	data.forEach(function(item,index){
		var time = item.time;
		var content = item.content;

		// console.log(time+'\n'+content+'\n\n');
		if(content.match(/<br>/) == 'null'){

			var arr = [{
				type: 'text',
				val: 'Simple'
			},{
				type : 'text',
				val : time,
				opt:{bold:true}
			},{
				type:'text',
				val:'第'+pageNum+'页',
				lopt:{align:'right'},
				opt:{color:'999999',font_face:'Microsoft Yahei',font_size:10}
			},{
				type : 'linebreak'
			},{
				type : 'text',
				val : content
			},{
				type : 'linebreak'
			}];
		}else{
			if(content.match(/<img.*?(?:>|\/>)/gi) != null){
				var imgArr = content.match(srcReg);
				imgArr.forEach(function(item3,index3){
					imgArr[index3] = item3.replace('original=','');
				})
				var imgArrI = 0;
				content = content.replace(/<img.*?(?:>|\/>)/gi,'<img>');
				// console.log(imgArr);
			}
			var arrContent = content.split('<br>');
			
			var arr = [{
				type: 'text',
				val: 'Simple'
			},{
				type : 'text',
				val : time,
				opt:{bold:true}
			},{
				type:'text',
				val:'第'+pageNum+'页',
				lopt:{align:'right'},
				opt:{color:'999999',font_face:'Microsoft Yahei',font_size:10}
			},{
				type : 'linebreak'
			}];

			arrContent.forEach(function(item2,index2){
				if(item2.replace(/\s+/g,'') == '<img>'){
					arr.push({type:'text',val:imgArr[imgArrI]});	//,opt:{Externallink:imgArr[imgArrI]}
					arr.push({type:'linebreak'});
					imgArrI++;
				}else{
					arr.push({type:'text',val:item2});
					arr.push({type:'linebreak'});
				}
			})
		}
		everyPageData.push(arr);
	})
	everyPageData.push({
		type : 'horizontalline'
	},{
		type:'text',
		val:'第'+pageNum+'页',
		lopt:{align:'right'},
		opt:{color:'999999',font_face:'Microsoft Yahei',font_size:10}
	});
	console.log('第'+pageNum+'页完成！');
	return everyPageData;
	
}


var page = 0;
var dataArr =[];
function httpGet(pageStart,pageEnd){
	if(page == 0){
		page = pageStart;
	}else{
		page++;
	}
	if(page > pageEnd){
		//导出word
		var pobj = docx.createP();	//创建段落
		pobj.options.align = 'left'; //设置居中格式
		//创建
		docx.createByJson(dataArr);
		var out = fs.createWriteStream(pageStart+'-'+pageEnd+'.docx');
		//输出
		docx.generate(out,{
			'finalize':function(written){
				console.log('创建word完成！');
			},
			'error':function(err){
				console.log('error');
			}
		})
		console.log('导出完成');
		return;
	}
	http.get('http://bbs.tianya.cn/post-no05-381555-'+page+'.shtml',function(res){
		var html = '';

		res.on('data', function(data){
			html += data;
		})

		res.on('end',function(){
			var data = filterChapters(html);
			dataArr.push.apply(dataArr,printData(data,page));
			httpGet(pageStart,pageEnd);
		})
	}).on('error',function(){
		console.log('错误!');
	})
}
httpGet(1,10);
