<<<<<<< HEAD
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




//console.log(range);

// var  = ['geosvit.com.ua',
//             'Без источника',
//             'mizol.ua',
//             'Prom.ua_EC',
//             'oman.com.ua',
//             'juta-ukraine.com.ua',
//             'biltmore-roof.com',
//             'Typar.com.ua',
//             'Рото (мобильный)',
//             'roto.com.ua',
//             'GoogleMyBusiness_EC',
//             'Пром',
//             'iBud_EC',
//             'profil.com.ua',
//             'Перезвон по пропущенному',
//             'izolit.ua',
//             'Izolit.ua_EC',
//             'https://profil.com.ua/',
//             'soffits.com.ua',
//             'agrid.com.ua',
//             'fasiding.com.ua',
//             'max3.com.ua',
//             'Повторное обращение',
//             'Geosvit_RingoError_EC',
//             'Айбуд',
//             'profil.com.ua/'];
// Main ();
//
// function GetRowsBySource(name) {
//   // This function take range by fixide cell 48 in wich we have collume with "Source" data in excel file, and by the compare give you array id rows
//   arrayIdRow = new Array ();
//   arrayIdRow.push(name);
//     for(var R = range.s.r; R <= range.e.r; ++R) {
//       var address_of_cell = {c:cellX, r:R};
//       var cell_ref = XLSX.utils.encode_cell(address_of_cell);
//       var desired_cell = sheet[cell_ref];
//       var desired_value = (desired_cell ? desired_cell.v : undefined);
//         if (desired_value == name) {
//           var a = parseInt(cell_ref.replace(/\D+/g,""));
//           arrayIdRow.push(a);
//         }
//     }
//   // console.log(arrayIdRow); // <----Получаем номера строк по совпадению с брендом
//     return arrayIdRow;
// }
//
// function GetArrayRowsById (arr) {
//   arrayRow = [];
//   arrayRow.push(arr[0]);
//
//     for (var i = 1; i <= arr.length; i++) { // Перебираем строку бренда с id рядами
//       var id = arr[i];
//      id--; // переменная с кординами рядков по бренду для екселя
//      arrayRow.push[i];
//      arrayRow[i] = []; // ****** ТУТ ОШИБКА
//
//     //  console.log('Значение '+id);
//       for (var cell = 0; cell <= range.e.c; ++cell) { //range.e.c
//         var address_of_cell = {c:cell, r:id};
//         var cell_ref = XLSX.utils.encode_cell(address_of_cell);
//         var a = sheet[cell_ref];
//       //  console.log(address_of_cell);
//         var desired_value = (a ? a.v : undefined);
//       //  console.log(arrayRow);
//         arrayRow[i].push(desired_value);
//
//       }
//
//     }
//
//   //  console.log(arrayRow[1][1]); // Смотреть количество масивов строк по бренду.
//     // ф-ция работает верно
//   return arrayRow;
// }
//
//
//
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
  let brands = [mizol];
  for (var k in brands) {
    SumCell(brands[k]);
    }
  }
=======
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




//console.log(range);

// var  = ['geosvit.com.ua',
//             'Без источника',
//             'mizol.ua',
//             'Prom.ua_EC',
//             'oman.com.ua',
//             'juta-ukraine.com.ua',
//             'biltmore-roof.com',
//             'Typar.com.ua',
//             'Рото (мобильный)',
//             'roto.com.ua',
//             'GoogleMyBusiness_EC',
//             'Пром',
//             'iBud_EC',
//             'profil.com.ua',
//             'Перезвон по пропущенному',
//             'izolit.ua',
//             'Izolit.ua_EC',
//             'https://profil.com.ua/',
//             'soffits.com.ua',
//             'agrid.com.ua',
//             'fasiding.com.ua',
//             'max3.com.ua',
//             'Повторное обращение',
//             'Geosvit_RingoError_EC',
//             'Айбуд',
//             'profil.com.ua/'];
// Main ();
//
// function GetRowsBySource(name) {
//   // This function take range by fixide cell 48 in wich we have collume with "Source" data in excel file, and by the compare give you array id rows
//   arrayIdRow = new Array ();
//   arrayIdRow.push(name);
//     for(var R = range.s.r; R <= range.e.r; ++R) {
//       var address_of_cell = {c:cellX, r:R};
//       var cell_ref = XLSX.utils.encode_cell(address_of_cell);
//       var desired_cell = sheet[cell_ref];
//       var desired_value = (desired_cell ? desired_cell.v : undefined);
//         if (desired_value == name) {
//           var a = parseInt(cell_ref.replace(/\D+/g,""));
//           arrayIdRow.push(a);
//         }
//     }
//   // console.log(arrayIdRow); // <----Получаем номера строк по совпадению с брендом
//     return arrayIdRow;
// }
//
// function GetArrayRowsById (arr) {
//   arrayRow = [];
//   arrayRow.push(arr[0]);
//
//     for (var i = 1; i <= arr.length; i++) { // Перебираем строку бренда с id рядами
//       var id = arr[i];
//      id--; // переменная с кординами рядков по бренду для екселя
//      arrayRow.push[i];
//      arrayRow[i] = []; // ****** ТУТ ОШИБКА
//
//     //  console.log('Значение '+id);
//       for (var cell = 0; cell <= range.e.c; ++cell) { //range.e.c
//         var address_of_cell = {c:cell, r:id};
//         var cell_ref = XLSX.utils.encode_cell(address_of_cell);
//         var a = sheet[cell_ref];
//       //  console.log(address_of_cell);
//         var desired_value = (a ? a.v : undefined);
//       //  console.log(arrayRow);
//         arrayRow[i].push(desired_value);
//
//       }
//
//     }
//
//   //  console.log(arrayRow[1][1]); // Смотреть количество масивов строк по бренду.
//     // ф-ция работает верно
//   return arrayRow;
// }
//
//
//
function SumCell (arr) {
	var  indexSum = 76; indexStat = 20;
	var b2b_oborot = 0, b2c_oborot = 0, otkaz_callcenter = 0, b2b = 0, necelevoy = 0, nedozvon = 0, oshibka = 0, povtor = 0, izn_nedozvon = 0, dymaet = 0, prodaj = 0, otkaz_diler  = 0;
	for (i=0; i < arr.length; i++) {
		var tcell1 = arr[i][indexSum];
		var stat = arr[i][indexStat];
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
// Заполняем массив с данными по одному бренду, на выходе получаем одну строку, которая суммирует все статусы и деньги.
  var _line = [];
  _line.push(b2b);
  //console.log(_line);
  // console.log('*********************'+'\n');
	// console.log('Денег '+ b2c_oborot +'\n');
	// console.log('Статусов отказ '+ otkaz_callcenter+'\n');
  // console.log('Статусов не целевой '+ necelevoy+'\n');
	// console.log('Статусов недозвон '+ nedozvon+'\n');
}
function Main () {
  DoArrayBrands();
  let brands = [mizol, oman, roto, ibud, prom, noname];
  for (var k in brands) {
    SumCell(brands[k]);      
    }
  }
  }
>>>>>>> d9a564395f026e7ac0f30e538cae485e30eb4e3f
