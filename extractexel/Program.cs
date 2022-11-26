using ExcelDataReader;
using extractexel;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Data;
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
string lastcode = "";
if (lastcode == "")
{

}
XElement document = new XElement("Document");
XAttribute plsch = new XAttribute("ПлСч11", $"{lastcode}");
foreach (System.Data.DataTable table in exelReaer.DataSet.Tables)
{
    for (int i = 3; i < table.Rows.Count; i++)
    {
        
        XElement data = new XElement("DATA");
        XAttribute column = new XAttribute("СТРОКА", "01");
        
        for (int j = 0; j < table.Columns.Count; j++)
            if (j == 1)
            {
                string code = $"1{table.Rows[i].ItemArray[j]}";
                lastcode = code.Remove(code.Length - 3);
            }
            else
            {

                XElement PxElement = new XElement("Px");
                // создаем для него атрибут name
                XAttribute number = new XAttribute("NUM", "1");
                // и два элемента - company и age
                XAttribute value = new XAttribute("VALUE", $"{table.Rows[i].ItemArray[j]}");
                PxElement.Add(number);
                PxElement.Add(value);
                data.Add(PxElement);
                

                
            }

        data.Add(column);
        document.Add(data);
        
        // добавляем два элемента person в корневой элемент


    }
}
xdoc.Add(document);

//сохраняем документ
xdoc.Save(xmlpath);


//XDocument xdoc = new XDocument();
//// создаем первый элемент person

//// создаем второй элемент person
//XElement bob = new XElement("Px");

//// создаем для него атрибут name
//XAttribute bobNameAttr = new XAttribute("NUM", "1");
//// и два элемента - company и age
//XAttribute bobCompanyElem = new XAttribute("VALUE", "Федеральные");

//bob.Add(bobNameAttr);

//bob.Add(bobCompanyElem);

//XElement bob2 = new XElement("Px");
//// создаем для него атрибут name
//XAttribute bobNameAttr2 = new XAttribute("NUM", "2");
//// и два элемента - company и age
//XAttribute bobCompanyElem2 = new XAttribute("VALUE", "4600");

//bob2.Add(bobNameAttr2);

//bob2.Add(bobCompanyElem2);
//XElement bob3 = new XElement("Px");
//// создаем для него атрибут name
//XAttribute bobNameAttr3 = new XAttribute("NUM", "3");
//// и два элемента - company и age
//XAttribute bobCompanyElem3 = new XAttribute("VALUE", "45");

//bob3.Add(bobNameAttr3);

//bob3.Add(bobCompanyElem3);

//// создаем корневой элемент
//XElement people = new XElement("DATA");
//XAttribute ar = new XAttribute("СТРОКА", "01");
//// добавляем два элемента person в корневой элемент
//people.Add(ar);
//people.Add(bob);
//people.Add(bob2);
//people.Add(bob3);
//// добавляем корневой элемент в документ
//xdoc.Add(people);
////сохраняем документ
//xdoc.Save(xmlpath);

//string _byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());

//XmlDocument doc = new XmlDocument();
//using (System.IO.StreamReader sr = new System.IO.StreamReader(re, System.Text.Encoding.GetEncoding(1251)))
//    doc.Load(sr);

Console.WriteLine("Data saved");
