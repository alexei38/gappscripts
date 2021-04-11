function updateMainTable(){
  // Запускает обновление всей таблички
  // У гугла ограничение работы скрипта - 6 минут
  // По этому скрипт не успевает обработать все записи, использовать можно только для отладки

  var kills_death_ss = getKillsDeathSheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ответы на форму (1)'); // Имя страницы с ответами
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  for (var i = 0; i < values.length; i++) {
    if(i == 0){
      // Пропускаем первую строчку, т.к там заголовок
      continue
    }
    updateRow(sheet, kills_death_ss, values[i])
  }
}


function updateForms(e){
  // функция работает как тригер при отправке формы
  // обновляет данные только по одной форме
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Килы/мертвые');
  var kills_death_ss = getKillsDeathSheet();
  updateRow(sheet, kills_death_ss, e.values);
}


function updateRow(orig_ss, ss, values){
  // Функция ищет и обновляет данные о игроке
  // values = массив с значениями из формы
  // values = [timestamp, login, id, honor_main, honor_farms]
  // поля считаются с 0

  var user_id = values[2];
  var honor_m = values[3];
  var honor_f = values[4];
  var orig_row_id = findRowIdByValue(orig_ss, 1, values[1])
  var row_id = findRowIdByValue(ss, 2, user_id)
  if(!row_id){
    // Если не нашли по id
    // попробуем найти по логину
    Logger.log('Find row_id by login ' + values[1]);
    var row_id = findRowIdByValue(ss, 1, values[1])
  }

  if(row_id){
    // Если нашли пользователя, обновляем данные и закрасим строку ответов в зеленый
    Logger.log('Set data row_id ' + row_id);
    setNewData(ss, row_id, honor_m, honor_f)
    setBackgroundColor(orig_ss, orig_row_id, 'green')
  } else {
    // если не нашли пользователя, форматируем строку в красный цвет
    Logger.log('Not found row_id:' + row_id);
    setBackgroundColor(orig_ss, orig_row_id, 'red')
  }
}


function getKillsDeathSheet(){
  // Тут id таблицы, в которой есть вкладка Килы/Мертвые, которую будем править
  var ss = SpreadsheetApp.openById("ВСТАВИТЬ_ID_ТАБЛИЦЫ");
  var SheetResponses = ss.getSheetByName("Килы/мертвые");
  return SheetResponses;
}


function findRowIdByValue(sheet, cell, value) {
  // Находим номер строки, в которой колонка соответствует значению
  // например находим колонку по id пользователя
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var row_id = null;
  for (var i = 0; i < values.length; i++) {
    if (trim(values[i][cell]) == trim(value)) {
      // Номера колонок начинаются с 1, по этому добавляем 1
      row_id = i + 1;
      break;
    }
  }
  return row_id
}

function trim(val){
  // удаляем пробелы в начале и конце строки
  return val.toString().trim();
}

function setNewData(ss, row_id, honor_m, honor_f){
  // Функция записывает данные о игроке в табличку с килами
  // Номера колонок в Килы/мертвые
  // Нумерация начинается с 1. A=1, B=2, C=3 и тд
  var honor_main_idx = 13;
  var honor_farms_idx = 14;
  var last_update_idx = 17;

  honor_main = ss.getRange(row_id, honor_main_idx)
  honor_main.setValue(honor_m);

  honor_farms = ss.getRange(row_id, honor_farms_idx);
  honor_farms.setValue(honor_f);

  // Сохраняем время посленей правки
  last_update = ss.getRange(row_id, last_update_idx);
  last_update.setValue(currentTime());
}

function setBackgroundColor(ss, row_id, color){
  var range = ss.getRange(row_id, 1, 1, 10);
  range.setBackground(color);
}

function currentTime() {
  var d = new Date();
  return d.toISOString()
}
