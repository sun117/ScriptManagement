/*--------------------------------------------------*/
/*
/* Date: 22/Jan/2018
/* Developer: Y.S.
/* Target: A.C. web site
/* Content: Scraiping keywords from the web site.
/*
/*--------------------------------------------------*/


//今週総合ランキング
//テスト編集 v3
function getWeeklyRanking(){
  
objSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var objSheet = objSpreadsheet.getActiveSheet();
  var activeSheetName = objSheet.getName();  
  if(activeSheetName == "Ranking"){  
    var response = UrlFetchApp.fetch("http://ranking.cosme.net/products");
    var regex = /<span class="item">([\s\S]*?)<\/span>/gi;
    var items    = response.getContentText("Shift-JIS").match(regex);
    
    for (var i=0; i<items.length; i++){
      objSheet.getRange(i+1,2).setValue(items[i]);
      }
  }
}