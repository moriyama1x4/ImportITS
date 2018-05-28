function importIts() {
  
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