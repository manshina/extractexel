using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using ExcelDataReader;
using System.IO;


namespace extractexel
{
    internal class ExelReaer
    {
        public DataSet DataSet { get; set; }
        public ExelReaer(string filepath)
        {
            FileStream stream = File.Open(filepath, FileMode.Open, FileAccess.Read);


            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            DataSet = excelReader.AsDataSet();
            excelReader.Close();
        }
        public void Print()
        {
            List<string> Dlist = new List<string>();

            

            //Data Reader methods
            foreach (System.Data.DataTable table in DataSet.Tables)
            {
                Console.WriteLine(table.Rows.Count);
                Console.WriteLine(table.Columns.Count);
                Console.WriteLine(table.TableName);
                //Console.WriteLine("\"" + table.Rows[6].ItemArray[4] + "\";");
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    for (int j = 0; j < table.Columns.Count; j++)
                        Console.WriteLine("\"" + table.Rows[i].ItemArray[j] + "\";");
                    Console.WriteLine(" ");
                }
                {
                    //for (int i = 0; i < table.Rows.Count; i++)
                    //{
                    //    Console.WriteLine("\"" + table.Rows[i].ItemArray[3] + "\";");
                    //    Dlist.Add(table.Rows[i].ItemArray[3].ToString());
                    //    Console.WriteLine();
                    //}
                }
            }


            //6. Free resources (IExcelDataReader is IDisposable)
            

            foreach (string item in Dlist) Console.WriteLine(item);

            Console.ReadKey();

        }
    }
}

//List<string> Dlist = new List<string>();

//FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);


//IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);



//DataSet result = excelReader.AsDataSet();

////Data Reader methods
//foreach (System.Data.DataTable table in result.Tables)
//{
//    Console.WriteLine(table.Rows.Count);
//    Console.WriteLine(table.Columns.Count);
//    Console.WriteLine(table.TableName);
//    //Console.WriteLine("\"" + table.Rows[6].ItemArray[4] + "\";");
//    for (int i = 0; i < table.Rows.Count; i++)
//    {
//        for (int j = 0; j < table.Columns.Count; j++)
//            Console.WriteLine("\"" + table.Rows[i].ItemArray[j] + "\";");
//        Console.WriteLine(" ");
//    }
//    {
//        //for (int i = 0; i < table.Rows.Count; i++)
//        //{
//        //    Console.WriteLine("\"" + table.Rows[i].ItemArray[3] + "\";");
//        //    Dlist.Add(table.Rows[i].ItemArray[3].ToString());
//        //    Console.WriteLine();
//        //}
//    }
//}


////6. Free resources (IExcelDataReader is IDisposable)
//excelReader.Close();

//foreach (string item in Dlist) Console.WriteLine(item);

//Console.ReadKey();