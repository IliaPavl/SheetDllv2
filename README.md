# SheetDll
To start in c++
// ConsoleApplication3.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
//

#include <iostream>
#include <Windows.h>
#import "C:\Users\User\Desktop\sheets\Sheet.tlb" no_namespace

int main()
{
    //idSheet берём из ссылки на нашу таблицу между d/.../edit : https://docs.google.com/spreadsheets/d/1KdBMWjLZJ-_Nhw3WHMOdEUUd-1Ln5tdf-zgCp6fc8D4/edit#gid=0
    const char idSheet[45] = "13KXwtNkdf1Duo5bYCPrEKqb6RDYnl9JMcGE8PXR4MHE";

    //nameSheet берём снизу слева (обычно "Лист 1") как на картинке https://i.imgur.com/lmJdBmC.png 
    const char nameSheet[8] = "Sheet1";

    //nameProgect уникальное имя проекта на каждую таблицу свое уникальное
    const char nameProgect[30] = "Curre32gdsjgk Legislators";

    const char jsonProperty[70] = "C:\\Users\\User\\Desktop\\sheets\\sheetService.json";

    CoInitialize(NULL);
    ISheetHelperPtr obj;
    obj.CreateInstance(__uuidof(SheetHelper));

    //устанавливаем настройки 
    obj->SetProperty(idSheet, nameSheet, jsonProperty, nameProgect);

    //Читаем строку 
    obj->PrintEntries(obj->ReadEntries("a1", "b32"));

    CoUninitialize();

}

