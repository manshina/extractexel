using ExcelDataReader;
using extractexel;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);


string path = "C:\\Users\\Dymok\\Downloads\\Telegram Desktop\\C# Тестовое Задание\\Тестовое Задание\\ТестовоеЗадание\\ФайлСИсходнымиДанными.xls";
string xmlpath = "C:\\Users\\Dymok\\Downloads\\Telegram Desktop\\C# Тестовое Задание\\Тестовое Задание\\ТестовоеЗадание\\metanit.xml";
string resultpath = "C:\\Users\\Dymok\\Downloads\\Telegram Desktop\\C# Тестовое Задание\\Тестовое Задание\\ТестовоеЗадание\\result2.xml";


ExelReaer exelReaer = new ExelReaer(path);

XDocument xdoc = new XDocument();



XElement root = new XElement("root");

XElement date = new XElement("Period");
XAttribute datevalue = new XAttribute("date", $"{File.GetLastWriteTime(path).Date.ToString("yyyy-MM-dd")}");

date.Add(datevalue);

root.Add(date);

XElement source = new XElement("Source");
XAttribute sourceClassCode = new XAttribute("ClassCode", "ДМЦ");
XAttribute sorceCode = new XAttribute("Code", "819");

source.Add(sourceClassCode);
source.Add(sorceCode);
date.Add(source);

XElement form = new XElement("Form");
XAttribute formCode = new XAttribute("Code", "178");
XAttribute formName = new XAttribute("Name", "Счета в кредитных организациях");

XAttribute formStatus = new XAttribute("Status", "0");

form.Add(formCode);
form.Add(formName);
form.Add(formStatus);



foreach (System.Data.DataTable table in exelReaer.DataSet.Tables)
{
    List<string> codes = new List<string>();
    for (int i = 3; i < table.Rows.Count; i++)
    {

        string code = $"1{table.Rows[i].ItemArray[1]}";
        code = code.Remove(code.Length - 3);
        if (!codes.Contains(code))
        {
            codes.Add(code);
        }



        for (int j = 0; j < table.Columns.Count; j++)
        {

        }
    }
    foreach (var itemCode in codes)
    {
        XElement document = new XElement("Document");


        XAttribute plsch = new XAttribute("ПлСч11", $"{itemCode}");

        




        document.Add(plsch);
        for (int i = 3; i < table.Rows.Count; i++)
        {


            string code = $"1{table.Rows[i].ItemArray[1]}";
            code = code.Remove(code.Length - 3);
            if (code == itemCode)
            {
                XElement data = new XElement("DATA");
                XAttribute dataString = new XAttribute("СТРОКА", i - 2);
                data.Add(dataString);

                int columnNumCounter = 1;
                for (int j = 0; j < table.Columns.Count; j++)
                {



                    XElement PxElement = new XElement("Px");
                    // создаем для него атрибут name
                    XAttribute number = new XAttribute("NUM", columnNumCounter);
                    // и два элемента - company и age
                    XAttribute value = new XAttribute("VALUE", $"{table.Rows[i].ItemArray[j]}");
                    PxElement.Add(number);
                    PxElement.Add(value);
                    data.Add(PxElement);



                    
                    if ((i == 3) && (j != 1))
                    {

                        XElement column = new XElement("Column");
                        XAttribute columnNum = new XAttribute("Num", $"{columnNumCounter}");

                        if (string.IsNullOrWhiteSpace(table.Rows[0].ItemArray[j]?.ToString()))
                        {
                            XAttribute columnName = new XAttribute("Name", $"{table.Rows[1].ItemArray[j]} {table.Rows[0].ItemArray[j-1]}");
                            column.Add(columnName);
                        }
                        else
                        {
                            XAttribute columnName = new XAttribute("Name", $"{table.Rows[1].ItemArray[j]} {table.Rows[0].ItemArray[j]}");
                            column.Add(columnName);
                        }

                        columnNumCounter++;
                        column.Add(columnNum);
                        form.Add(column);


                    }



                }
                document.Add(data);


            }

        }

        form.Add(document);
    }





}
source.Add(form);


xdoc.Add(root);

//сохраняем документ
xdoc.Save(xmlpath);



