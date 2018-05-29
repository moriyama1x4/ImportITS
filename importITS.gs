var sheet = SpreadsheetApp.getActive().getSheetByName('ITSリスト');;

function importIts() {
  var topUrl = 'http://www.its-kenpo.or.jp/shisetsu/golf/';
  var listPath = ['ichiran/', 'taiheiyou/', 'tokyu/', 'resol'];
  var listCategory = ['その他', '太平洋', '東急', 'リソル'];
  var row = 2;
  
  listPath.forEach(function(value1, index1){
    var listHtml = UrlFetchApp.fetch(topUrl + value1).getContentText('UTF-8');
    var detailTrs = getChildTags(listHtml,[
      ['section', '<section class="section">', '関東エリア', 0],['tr','<tr>','']
    ]);
    
    detailTrs.forEach(function(value2, index2){
      //一覧からとれるもの
      //運営会社取得
      var company = listCategory[index1];
      
      //運営会社入力
      setData(row,1,company);
      
      
      //県取得
      var prefecture = getTags(value2, 'th', '<th>', '')[0].replace(/・.*/,'');
      
      //県入力
      setData(row,2,prefecture);
      
      
      //ゴルフ場名取得
      var courseName = getTags(value2, 'a', '<a.*?>', '')[0];
      
      //ゴルフ場名入力
      setData(row,3,courseName);
      
      
      //ITS詳細Path取得
      var pathReg = new RegExp(value1.replace('/','') + '\\/.*\\.html');
      var detailPath = value2.match(pathReg)[0];
      
      //ITS詳細URL入力
      setData(row,9,topUrl + detailPath);
      
      
      //ここから詳細ページ
      var detailHtml = UrlFetchApp.fetch(topUrl + detailPath).getContentText('UTF-8');
      var feeTrs = getChildTags(detailHtml,[
	  ['section', '<section class="section">', '<h2>ご利用料金</h2>', 0],['tr', '<tr>', 'td']
	  ]);
      
	  //キャディ付き除外
	  feeTrs.forEach(function(value3, index3, array3){
		if(value3.match(/キャディ付/)){
			var rowspan;
			if(value3.match(/rowspan=.*?キャディ.*?<\/td>/)){
				rowspan = value3.match(/rowspan=.*?キャディ.*?<\/td>/)[0].match(/[0-9]/)[0];
			}else{
				rowspan = 1;
			}
			array3.splice(index3,rowspan);
		}
	  });
      
      if(feeTrs.length == 0){
		setData(row,5,'キャディ付きしかありません')
	  }else{
		var bigColumn = [];
		feeTrs.forEach(function(value3, index3){
			if(value3.match(/rowspan/)){
				bigColumn.push(index3);
			}
		});
		if(bigColumn.length > 1){
			setData(row,4,'A');
		}else{
			var holyday　= [];
			feeTrs.forEach(function(value3, index3){
				if(value3.match(/祝/)){
					holyday.push(index3);
				}
			});
			if(holyday.length == 0){
				setData(row,7,'平日しかありません');
			}else if(holyday.length == 1){
				var feeTds = getTags(feeTrs[holyday[0]],'td','<td>','円');
				setData(row,7,Utilities.formatString('%05d',feeTds[0].replace(/,|円|<(\/)?p>/g,'')));
				setData(row,8,Utilities.formatString('%05d',feeTds[1].replace(/,|円|<(\/)?p>/g,'')));
			}else{
				setData(row,4,'B');
			}
			
			var weekday　= [];
			feeTrs.forEach(function(value3, index3){
				if(value3.match(/平|月|火|水|木|金/)){
					weekday.push(index3);
				}
			});
			if(weekday.length == 0){
				setData(row,5,'休日しかありません');
			}else{
				var feeTds = getTags(feeTrs[weekday[0]],'td','<td>','円');
				setData(row,5,Utilities.formatString('%05d',feeTds[0].replace(/,|円|<(\/)?p>/g,'')));
				setData(row,6,Utilities.formatString('%05d',feeTds[1].replace(/,|円|<(\/)?p>/g,'')));
			}
		}
	  }
	  
	  
      row++;
    });
  });
}



function setData(y,x,data){
  var range = sheet.getRange(y, x);
  range.setValue(data);
}

//tagType:'div'とか, tagReg:開始タグの正規表現, elementReg:中に含まれる要素の正規表現
function getTags(xml,tagType,tagReg,elementReg){
  var indexStartTag;
  var xmls = [];
  tagReg = new RegExp(tagReg);
  elementReg = new RegExp(elementReg);
  
  for (var i = 0;true;i++){
    indexStartTag = xml.search(tagReg);
    if(indexStartTag !== -1){
      xml = xml.substring(indexStartTag + xml.match(tagReg)[0].length);
      var copyXml = xml;
      var index = 0;
      var endTagNum = 0; //開始タグに対する終了タグの数。これが1になったら親要素の終了タグとみなす
      var reg = new RegExp('<(/)?' + tagType);
      
      while(endTagNum < 1){
        index += copyXml.search(reg) + 1;
        if(copyXml.match(reg)[0] == '<' + tagType){
          endTagNum --;
        }else{
          endTagNum ++;
        }
        copyXml = xml.substring(index)
      }
      
      if(xml.substring(0,index - 1).search(elementReg) !== -1){
        xmls.push(xml.substring(0,(index - 1)));
      }
      xml = xml.substring((index - 1) + (tagType.length + 3));
    }else{
      break;
    }
  }
  return xmls;
}

function getChildTags(xml,array){ //array = [[tagType,tagReg,elementReg,num],[tagType,tagReg,elementReg]]
  array.forEach(function(value,index){
    xml = getTags(xml,value[0],value[1],value[2]);
    if(index !== array.length - 1){
      xml = xml[value[3]];
    }
  });
  return xml;
}