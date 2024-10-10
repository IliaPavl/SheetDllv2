using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Sheet
{
    [ComVisible(true)]
    public interface ISheetHelper
    {
        bool SetProperty(string sheetId, string sheetName, string passJsonKey, string nameProgect);
        bool SetPropertyByJson(string sheetId, string sheetName, string jsonKey, string nameProgect);
        void PrintEntries(string[,] values);
        void DeleteEntry(string start, string end);
        string ReadEntry(string point);
        string[,] ReadEntries(string start, string end);
        void CreateEntry(string start, string end, string[] listValues);
        void UpdateEntry(string point, string[] listValues);
        void PrintNotNullEntries(string[,] values);
        //string[,] ReadCommand(string programString);
        void UpdateCommand(string command, string[] listValues);
        void CreateCommand(string command, string[] listValues);
        void DeleteCommand(string command);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class SheetHelper : ISheetHelper
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName;
        static string SpreadsheetId;
        static string sheet;
        static SheetsService service;

        //Установка настроек ______________________________________________________________________
        public bool SetProperty(
            string sheetId,
            string sheetName,
            string passJsonKey,
            string nameProgect)
        {
            try
            {
                ApplicationName = nameProgect;
                SpreadsheetId = sheetId;
                sheet = sheetName;
                GoogleCredential credential;
                using (var stream = new FileStream(passJsonKey, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(Scopes);
                }

                // Create Google Sheets API service.
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return false; }
            return true;

        }

        public bool SetPropertyByJson(
            string sheetId,
            string sheetName,
            string jsonKey,
            string nameProgect)
        {
            try
            {
                ApplicationName = nameProgect;
                SpreadsheetId = sheetId;
                sheet = sheetName;
                GoogleCredential credential = GoogleCredential.FromJson(jsonKey).CreateScoped(Scopes);

                // Create Google Sheets API service.
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return false; }
            return true;
        }

        //Красивый вывод __________________________________________________________________________
        public void PrintEntries(string[,] values)
        {
            if (values != null && values.Length > 0)
            {
                for (int j = 0; j < values.GetLength(0); j++)
                    for (int i = 0; i < values.GetLength(1); i++)
                        if (i == values.GetLength(1) - 1)
                        {
                            Console.Write("{0}", values[j, i]);
                            Console.WriteLine("");
                        }
                        else
                            Console.Write("{0} | ", values[j, i]);
                Console.WriteLine("");
            }
            else
                Console.WriteLine("No data found.");
        }

        public void PrintNotNullEntries(string[,] values)
        {
            if (values != null && values.Length > 0)
            {
                for (int j = 0; j < values.GetLength(0); j++)
                {
                    for (int i = 0; i < values.GetLength(1); i++)
                        if (i == values.GetLength(1) - 1)
                            if (values[j, i] != null)
                                Console.Write("{0}", values[j, i]);
                            else if (values[j, i] != null)
                                Console.Write("{0} | ", values[j, i]);
                    Console.WriteLine("");
                }
                Console.WriteLine("");
            }
            else
                Console.WriteLine("No data found.");
        }

        //Круд операции ___________________________________________________________________________
        public string[,] ReadEntries(string start, string end)
        {
            string range = $"{sheet}!{start}:{end}";
            return ReadCommand(range);
        }

        public string ReadEntry(string point)
        {
            var range = $"{sheet}!{point}:{point}";
            string[,] list = ReadCommand(range);
            if (list != null && list.GetLength(0) > 0 && list.GetLength(1) > 0)
                return list[0, 0];
            else
                return null;
        }

        public void DeleteEntry(string start, string end)
        {
            var range = $"{sheet}!{start}:{end}";
            DeleteCommand(range);
        }

        public void CreateEntry(string start, string end, string[] listValues)
        {
            var range = $"{sheet}!{start}:{end}";
            CreateCommand(range, listValues);
        }

        public void UpdateEntry(string point, string[] listValues)
        {
            var range = $"{sheet}!{point}:{point}";
            UpdateCommand(range, listValues);
        }

        //Любой запрос сюда подставляеш ___________________________________________________________
        public string[,] ReadCommand(string programString)
        {
            try
            {
                SpreadsheetsResource.ValuesResource.GetRequest request =
                         service.Spreadsheets.Values.Get(SpreadsheetId, programString);
                IList<IList<object>> obj = request.Execute().Values;
                string[,] list = null;
                int firstColumn = obj.Count, endColumn = -1;
                if (obj != null && obj.Count > 0)
                {
                    for (int j = 0; j < obj.Count; j++)
                        if (endColumn < obj[j].Count)
                            endColumn = obj[j].Count;

                    list = new string[firstColumn, endColumn];

                    for (int j = 0; j < obj.Count; j++)
                        for (int i = 0; i < obj[j].Count; i++)
                            list[j, i] = (string)obj[j][i];
                }
                return list;
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return null; }

        }


        public void UpdateCommand(string command, string[] listValues)
        {

            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { listValues };
            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, command);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = updateRequest.Execute();

        }

        public void CreateCommand(string command, string[] listValues)
        {

            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { listValues };
            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, command);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = updateRequest.Execute();

        }

        public void DeleteCommand(string command)
        {

            var requestBody = new ClearValuesRequest();
            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, command);
            var deleteReponse = deleteRequest.Execute();

        }

        private string[,] FindWords(string[] words, string[,] array)
        {
            try
            {
                List<string> foundWords = new List<string>();
                List<string> foundCoordinates = new List<string>();
                List<int> letterValues = new List<int>(); // Список для хранения числовых значений буквенных координат

                for (int i = 0; i < array.GetLength(0); i++)
                {
                    for (int j = 0; j < array.GetLength(1); j++)
                    {
                        foreach (var word in words)
                        {
                            if (array[i, j] == word)
                            {
                                char columnLetter = (char)('a' + j); // 'a' + индекс столбца
                                int rowNumber = i + 1; // Индекс строки + 1
                                foundWords.Add(word);
                                foundCoordinates.Add($"{columnLetter}{rowNumber}");
                                letterValues.Add(columnLetter - 'a' + 1); // Сохраняем числовое значение буквы (a=1, b=2 и т.д.)
                            }
                        }
                    }
                }

                if (foundWords.Count == 0)
                {
                    return new string[3, 0]; // Возвращаем пустой массив с тремя строками
                }

                string[,] resultArray = new string[3, foundWords.Count];

                int firstLetterValue = letterValues[0]; // Числовое значение первой буквы

                for (int i = 0; i < foundWords.Count; i++)
                {
                    resultArray[0, i] = foundWords[i];         // Слова
                    resultArray[1, i] = foundCoordinates[i];   // Координаты
                    resultArray[2, i] = (letterValues[i] - firstLetterValue).ToString(); // Числовое значение текущей буквы - 1 и - значение первой буквы
                }

                return resultArray;
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return null; }
        }

        private string[] getJsonObjNames(String jsonString)
        {
            try
            {
                // Parse the JSON string into a JObject
                JObject jsonObject = JObject.Parse(jsonString);

                // Extract the keys into a list
                List<string> keysList = new List<string>();
                foreach (var property in jsonObject.Properties())
                {
                    keysList.Add(property.Name);
                }

                // Convert the list to an array
                string[] keysArray = keysList.ToArray();
                return keysArray;
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return null; }
        }

        private static (int, int) GetIndexAndCount(string json)
        {
            try
            {

                // Парсим JSON строку
                var jsonObject = JObject.Parse(json);

                // Получаем значения index и count с установкой значений по умолчанию
                int index = (int)(jsonObject["index"] ?? 0);
                int count = (int)(jsonObject["count"] ?? 50);

                return (index, count);
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return (0, 0); }
        }

        public string FindData(
            string jsonParamsNames,
            string wordToSearch,
            string dateStart,
            string dateEnd,
            string settingJson
            )
        {
            try
            {
                //ищем таблицу в диапазоне 32/32 ячейки
                string[,] array = ReadEntries("a1", "z32");

                //получаем index - отступ от первой даты и conut - число строк которые мы берём за проход.
                var settings = GetIndexAndCount(settingJson);

                //получаем список названий столбцов
                string[] wordsToFind = getJsonObjNames(jsonParamsNames);

                //находим точные координаты названий
                string[,] findedWords = FindWords(wordsToFind, array);

                string[] decimalNames = getDecimalNames(jsonParamsNames);

                DateTime.TryParse(dateStart, out DateTime startDate);
                DateTime.TryParse(dateEnd, out DateTime endDate);

                //находим координаты "квадра" таблицы с датами в промежутке от startDate до endDate
                string[] coordinates = FindDatesInRange(
                    findedWords,
                    wordToSearch,
                    startDate,
                    endDate,
                    settings.Item1,
                    settings.Item2
                    );

                //ищем данные по полученным координатам
                if (coordinates == null)
                {
                    return returnNullDataToJson();
                }

                string[,] findedData = ReadEntries(coordinates[0], coordinates[1]);

                //возврвщвем в формате json
                return returnFindDataToJson(findedWords, findedData, decimalNames);
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return null; }

        }
        private string[] getDecimalNames(string jsonString)
        {
            try
            {
                var dictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);

                // Find field names where values start with "DECIMAL"
                List<string> decimalFields = new List<string>();
                foreach (var kvp in dictionary)
                {
                    if (kvp.Value.StartsWith("DECIMAL"))
                    {
                        decimalFields.Add(kvp.Key);
                    }
                }

                // Convert to string array and print results
                return decimalFields.ToArray();
            }catch (Exception e) { Console.WriteLine(e); return null; }
        }

        private string returnNullDataToJson()
        {
            try
            {
                var results = new List<Dictionary<string, string>>();
                var finalResult = new
                {
                    next = 0,
                    result = results
                };

                return JsonConvert.SerializeObject(finalResult);
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return null; }
        }


        private string returnFindDataToJson(string[,] findedWords, string[,] findedData, string[] decimalNames)
        {
            try
            {
                int rowCount = findedData.GetLength(0);
                int columnCount = findedWords.GetLength(1);

                var results = new List<Dictionary<string, object>>();

                for (int i = 0; i < rowCount; i++)
                {
                    var obj = new Dictionary<string, object>();
                    for (int j = 0; j < columnCount; j++)
                    {
                        string header = findedWords[0, j]; // Заголовок из первой строки
                        int number2 = Convert.ToInt32(findedWords[2, j]);
                        string value = findedData[i, number2]; // Значение из findedData по индексу

                        if (isDecimal(decimalNames, header))
                        {
                            // Преобразование value в число
                            if (decimal.TryParse(value, out decimal decimalValue))
                            {
                                obj[header] = Math.Round(decimalValue, 2); // Сохранение как decimal с двумя знаками
                            }
                            else
                            {
                                obj[header] = value; // Если преобразование не удалось, сохранить оригинальное значение
                            }
                        }
                        else
                        {
                            obj[header] = value; // Сохранение как строка
                        }
                    }
                    results.Add(obj);
                }

                var finalResult = new
                {
                    next = rowCount,
                    result = results
                };

                // Сериализация без кастомного конвертера
                return JsonConvert.SerializeObject(finalResult);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e);
                return null;
            }


        }

        private bool isDecimal(string[] decimalNames, string name) {
            bool exists = false;
            if(decimalNames.Length>0 && name!=null)
            {foreach (var word in decimalNames)
                {
                    if (word == name)
                    {
                        exists = true;
                        break;
                    }
                }
            }
            return exists;
        }

        private string[] FindDatesInRange(
            string[,] result,
            string wordToSearch,
            DateTime startDate,
            DateTime endDate,
            int index,
            int count
            )
        {
            try
            {
                int wordIndex = -1;

                // Поиск индекса слова в результатах
                for (int i = 0; i < result.GetLength(1); i++)
                {
                    if (result[0, i] == wordToSearch)
                    {
                        wordIndex = i;
                        break;
                    }
                }

                if (wordIndex == -1) return null; // Если слово не найдено

                var coordinatesStr = result[1, wordIndex];
                char columnLetterStart = coordinatesStr[0];
                int rowNumberStart = int.Parse(coordinatesStr.Substring(1));




                string start = $"{columnLetterStart}";
                string end = $"{columnLetterStart}";

                string[,] currentCellValues = ReadEntries(start, end);
                int cellValueStart = -1;
                int cellValueEnd = -1;

                // Проверяем значения в ячейках текущего диапазона
                for (int rowIndex = 0, counter = -1, indexCounter = 0; rowIndex + rowNumberStart
                    < currentCellValues.GetLength(0); rowIndex++)
                {
                    var currentCellValueStart = currentCellValues[rowIndex + rowNumberStart, 0]; // Всегда первый столбец

                    if (!DateTime.TryParse(currentCellValueStart, out DateTime currentDate))
                        return null; // Если не дата - завершаем поиск

                    // Пропускаем строки на основе значения index
                    if (cellValueStart < 0 && currentDate >= startDate)
                        cellValueStart = rowIndex + 1 + rowNumberStart; // Учитываем index

                    if (currentDate >= startDate && currentDate <= endDate)
                        cellValueEnd = rowIndex + 1 + rowNumberStart;

                    if (cellValueEnd >= 0 && cellValueStart >= 0)
                        counter = cellValueEnd - cellValueStart;

                    if (counter == 0 && index != indexCounter)
                    {
                        indexCounter += 1;
                        cellValueStart += 1;
                        continue;
                    }


                    if (cellValueEnd > 0 && (counter + 1 == count || (currentDate >= endDate && counter > 0) || (rowIndex + rowNumberStart >= currentCellValues.GetLength(0) - 1)))
                    {
                        char lastColumnLetter = result[1, result.GetLength(1) - 1][0]; // Последняя буква из найденных координат
                        var finalCoordinateEnd = $"{lastColumnLetter}{cellValueEnd}";
                        return new[] { $"{result[1, 0][0]}{cellValueStart}", finalCoordinateEnd };
                    }
                }


                return null; // Если ничего не найдено
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return null; }
        }


        public string FilterDates(string[,] data, DateTime startDate, DateTime endDate, int dateColumnIndex)
        {
            try
            {


                var headers = new List<string>();
                var result = new List<Dictionary<string, string>>();

                // Получаем заголовки из первой строки
                for (int i = 0; i < data.GetLength(1); i++)
                {
                    headers.Add(data[0, i]);
                }

                // Проходим по данным начиная со второй строки
                for (int i = 1; i < data.GetLength(0); i++)
                {
                    var rowDateStr = data[i, dateColumnIndex];
                    if (!string.IsNullOrEmpty(rowDateStr) && DateTime.TryParse(rowDateStr, out DateTime rowDate))
                    {
                        // Проверка на попадание в диапазон
                        if (rowDate >= startDate && rowDate <= endDate)
                        {
                            var rowDict = new Dictionary<string, string>();
                            for (int j = 0; j < data.GetLength(1); j++)
                            {
                                rowDict[headers[j]] = data[i, j];
                            }
                            result.Add(rowDict);
                        }
                    }
                }

                // Формируем итоговый ответ
                var response = new
                {
                    count = result.Count,
                    date = DateTime.UtcNow.ToString("o"), // Формат ISO 8601
                    result
                };

                return JsonConvert.SerializeObject(response);
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return null; }

        }
    }
}
