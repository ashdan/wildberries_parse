/*================================
Парсинг главной страницы выдачи https://www.wildberries.ru/
Скрипт написан для работы в Google Sheets. Необходимо скопировать скрипт, а также переименовать диапазоны, либо поправить скрипт
==================================*/

function main(){
  const sheet = SpreadsheetApp.getActiveSheet();
  const now = new Date();
  sheet.getRange("status").setValue("Очистка данных...");
  getClean(["status","B8:K107","M8:M107", "similar_q"]);

  sheet.getRange("status").setValue("Загрузка 10 похожих запросов...");
  sheet.getRange("similar_q").setValue(getSimilarQueries());

  sheet.getRange("status").setValue("Загрузка 100 товаров...");
  sheet.getRange("B8:J107").setValues(getCatalog());

  sheet.getRange("status").setValue("Загрузка объемов продаж...");
  sheet.getRange("K8:K107").setValues(getTotalSales(sheet.getRange("B8:B107").getValues()));

  if (dialogMessage("Загружать остатки товаров на складах? Это потребует около 5 минут...")){
    sheet.getRange("status").setValue("Загрузка остатков...");
    sheet.getRange("M8:M107").setValues(getQuantities(sheet.getRange("B8:B107").getValues()));
  } 
  sheet.getRange("status").setValue("Данные актуальны на " + now.toLocaleString());
}

function getSimilarQueries(){
  //загрузка 10 похожих слов по аналогичным товарам
  const similar_queries = getJSON(`https://similar-queries.wildberries.ru/api/v2/search/query?query=${getRange("qlink")}`);
  return similar_queries.query.join(", ")
}

function getQuantities(ids){
  //Загружаем остатки на складах
  const url = "https://card.wb.ru/cards/detail?appType=1&curr=rub&dest=-1257786&regions=80,38,4,64,83,33,68,70,69,30,86,75,40,1,66,110,22,31,48,71,114&spp=0&nm=";
  const sheet = SpreadsheetApp.getActiveSheet();
  let quantities =[];

  for (let id of ids){
    try{
      sheet.getRange("status").setValue(`Загрузка остатка по ID${id}`);
      const balance = getJSON(url+id);
      quantities.push([getCalculate(balance.data.products[0].sizes[0].stocks)]);
      Utilities.sleep(2100);//перерыв менее 2,5 секунд
    }catch{
      quantities.push(["-"]);

    }
  }
  return quantities;
}

function getCatalog(){
  //Загружаем каталог товаров из первой страницы, за исключением объема продаж и складских запасов
  const raw_catalog = getJSON(`https://search.wb.ru/exactmatch/ru/common/v4/search?TestGroup=control&TestID=188&appType=1&curr=rub&dest=-1257786&query=${getRange("qlink")}&regions=80,38,4,64,83,33,68,70,69,30,86,75,40,1,66,110,22,31,48,71,114&resultset=catalog&sort=popular&spp=0&suppressSpellcheck=false`);
  
  let catalog =[];
  
  for (let i = 0; i <= raw_catalog.data.products.length - 1; i ++) {
    try{
      catalog.push(
        [raw_catalog.data.products[i].id,
        raw_catalog.data.products[i].name,
        raw_catalog.data.products[i].brand,
        raw_catalog.data.products[i].brandId,
        raw_catalog.data.products[i].priceU/100,
        raw_catalog.data.products[i].salePriceU/100,
        raw_catalog.data.products[i].sale,
        raw_catalog.data.products[i].reviewRating,
        raw_catalog.data.products[i].feedbacks]
      );
    } catch{
      break;
    }
  }
  return catalog;
}

function getTotalSales(ids){
  //Получаем объем продаж каждого товара
  const total_sale = getJSON(`https://product-order-qnt.wildberries.ru/by-nm/?nm=${ids.join(",")}`);
  let qnts=[];
  for(let qnt of total_sale){
    qnts.push([qnt.qnt]);
  }
  return qnts;
}

function dialogMessage(message){
  const answer = SpreadsheetApp.getUi().alert(message, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  return (answer===SpreadsheetApp.getUi().Button.OK ? true : false); 
}

function getClean(ranges){
  const sheet = SpreadsheetApp.getActiveSheet();
  for (let range of  ranges){
    sheet.getRange(range).clearContent();
  }
}

function getRange(range=null, isArray = false){
  const sheet = SpreadsheetApp.getActiveSheet();
  return (isArray ? sheet.getRange(range).getValues() : sheet.getRange(range).getValue())
}

function getJSON(url){
  //получаем url, возвращаем JSON
  return JSON.parse(UrlFetchApp.fetch(url).getContentText());
}

function getCalculate(data){
  //считаем количество товаров на складе (тэг qty)
  let accumulate = 0;
  for (let p of data){
    try{
      accumulate += p.qty;
    }
    catch{
      console.log("Error")
    }
  }
  return accumulate;
}

