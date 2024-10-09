using System;
using System.IO;


class Program
{
    static void Main(string[] args)
    {
        //Для работы необходимо добавить почту с возможностью редактировать exel таблицы (sheet).
        //Почта: "sheetservise@sheets-371512.iam.gserviceaccount.com"

        //idSheet берём из ссылки на нашу таблицу между d/.../edit : https://docs.google.com/spreadsheets/d/1KdBMWjLZJ-_Nhw3WHMOdEUUd-1Ln5tdf-zgCp6fc8D4/edit#gid=0
        string idSheet = "1uv3Itt3QlcYesedHIzD2PiRcxHsWLL8BzHQediX8Lb4";
       
        //nameSheet берём снизу слева (обычно "Лист 1") как на картинке https://i.imgur.com/lmJdBmC.png 
        string nameSheet = "Sheet1";

        //nameProgect уникальное имя проекта на каждую таблицу свое уникальное
        string nameProgect = "Currenew1ed3214 Legislators";

        //создаем экземпдяр SheetHelper
        Sheet.SheetHelper program = new Sheet.SheetHelper();

        //устанавливаем настройки 
        program.SetProperty(idSheet, nameSheet, "I:\\Games\\dll\\ConsoleApp2\\sheets-371512-5b0dc5434543.json", nameProgect);

        
        string json = program.FindData(
            "{\"CREATED_DATE\": \"NVARCHAR(50)\",\"ID\":\"NVARCHAR(50)\",\"NAME\":\"NVARCHAR(250)\",\"SORT\":\"NVARCHAR(250)\",}",
            "CREATED_DATE",
            "2021, 02, 25",
            "2023, 02, 28",
            "{\"index\":0, \"count\":50}");
    
        Console.WriteLine(json);    
    }
}


