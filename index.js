if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('file.xls');
var sheet_name_list = workbook.SheetNames;
var json = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
//console.log(json[0]['Источник']);
let geosvit = [], mizol = [], prom = [], oman = [], juta = [], biltmore = [], typar = [], roto = [], map = [], ibud = [], profil = [], izolit = [], soffit = [], agrid = [], fasiding = [], max3 = [], noname  = [];

Main();

function DoArrayBrands() {
  for (k in json) {
      let source = json[k]["Источник"];

        if ( (source.search(/prom/i) + 1) || ((source.search(/пром/i)) + 1)) { prom.push(arrayFromJson(k)); }
          else if ( source.search(/mizol/i) + 1)  {                    mizol.push(arrayFromJson(k));  }
          else if ( source.search(/agrid/i) + 1) {                agrid.push(arrayFromJson(k));  }
          else if ( (source.search(/ibud/i) + 1) || ((source.search(/айбу/i))+1)) {   ibud.push(arrayFromJson(k));  }
          else if ( source.search(/oman/i) + 1) {                oman.push(arrayFromJson(k));  }
          else if ( source.search(/juta/i) + 1) {                juta.push(arrayFromJson(k));  }
          else if ( source.search(/biltmore/i) + 1) {                biltmore.push(arrayFromJson(k));  }
          else if ( source.search(/typar/i) + 1) {                typar.push(arrayFromJson(k));  }
          else if ( (source.search(/roto/i) + 1) || (source.search(/рото/i))+1) {   roto.push(arrayFromJson(k));  }
          else if ( source.search(/google/i) + 1) {                map.push(arrayFromJson(k));  }
          else if ( source.search(/izolit/i) + 1) {               izolit.push(arrayFromJson(k));  }
          else if ( source.search(/soffit/i) + 1) {                soffit.push(arrayFromJson(k));  }
          else if ( source.search(/fasiding/i) + 1) {                fasiding.push(arrayFromJson(k));  }
          else if ( source.search(/max3/i) + 1) {                max3.push(arrayFromJson(k));  }
          else if ( source.search(/geos/i) + 1) {                geosvit.push(arrayFromJson(k));
        }
        //  Все без источника
          else {  noname.push(arrayFromJson(k));  }



    function arrayFromJson(k) {
      let tempArray = [];
        for (j in json[k]) {
          tempArray.push(json[k][j]);
        }
    return tempArray;
    }
  }


}


function SumCell (arr) {
	var b2b_oborot = 0, b2c_oborot = 0, otkaz_callcenter = 0, b2b = 0, necelevoy = 0, nedozvon = 0, oshibka = 0, povtor = 0, izn_nedozvon = 0, dymaet = 0, prodaj = 0, otkaz_diler  = 0;
	for (i = 0; i < arr.length; i++) {
   /// Проверяем колонку сумму на значание и переводим ее в число
    var tcell1 = arr[i][4];
      if (!!tcell1) {tcell1 = parseInt(tcell1, 10);}
      else tcell1 = 0;

		var stat = arr[i][1];
    b2c_oborot = b2c_oborot + tcell1;
    if (stat == 'Изначальный недозвон к клиенту') {nedozvon++;
    } else if (stat == ('Клиент отказался (B2B)' || 'Клиент отказался' || 'Клиент отказался (сп)' || 'Клиент отказался (В2С)')) {otkaz_diler++;
    } else if (stat == 'Не целевой') {necelevoy++;
    } else if (stat == ('В2В обьект (ПП)' || 'Не целевой' || 'Новый дилер (СП)')) {b2b++;
    } else if (stat == ('В2С на менеджера' || 'Дилер принял лид' || 'Сделали предложение клиент думает')) {dymaet++;
    } else if (stat == 'Закрыть' || 'Не целевой' || 'Разговор') {necelevoy++;
    } else if (stat == 'Изначальный недозвон к клиенту') {izn_nedozvon++;
		} else if (stat == 'Ошибка/тест') {oshibka++;
    } else if (stat == 'Отказ на этапе СС') {otkaz_callcenter++;
    } else if (stat == 'Отмена колл-центр') {nedozvon++;
    } else { console.log('Пустое поле статуса'); }
    
  }
  console.log(b2c_oborot);
// Заполняем массив с данными по одному бренду, на выходе получаем одну строку, которая суммирует все статусы и деньги.

}

function Main () {
  DoArrayBrands();
  let brands = [mizol, oman, roto, ibud, prom, noname];
  for (var k in brands) {
    SumCell(brands[k]);      
    }
  }
  

