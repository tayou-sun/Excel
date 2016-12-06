using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Authentication.ExtendedProtection.Configuration;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleApplication1
{
    class Program
    {
        public class FCPObject
        {
            int _OID;
            String _FCPObjectName;
            int _FundingDirection;
            String _Name;
            int _Number;
            int _Order;

            public FCPObject(int OID, String FCPObjectName, int FundingDirection, String Name, int Number, int Order)
            {
                _OID = OID;
                _FCPObjectName = FCPObjectName;
                _FundingDirection = FundingDirection;
                _Name = Name;
                _Number = Number;
                _Order = Order;

            }
        }

        static void Main(string[] args)
        {
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook =
                ObjWorkExcel.Workbooks.Open(@"C:\Users\ssbye\Desktop\FCPObjectHistory_Order_Test.xlsx", Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet) ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell); //1 ячейку
            //-------------------------------------
            int lastColumn = (int) lastCell.Column; //!сохраним непосредственно требующееся в дальнейшем
            int lastRow = (int) lastCell.Row;
            //-------------------------------------
            string[,] list = new string[lastCell.Column, lastCell.Row];
                // массив значений с листа равен по размеру листу
            for (int i = 0; i < (int) lastCell.Column; i++) //по всем колонкам
                for (int j = 0; j < (int) lastCell.Row; j++) // по всем строкам
                    list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString(); //считываем текст в строку
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой -- в том числе не используемые явно объекты !
            //for (int i = 1; i < lastColumn; i++) //по всем колонкам
            //    for (int j = 1; j < lastRow; j++) // по всем строкам 

            string[] str = new string[lastRow];

            for (int i = 1; i < lastRow; i++)
            {//по всем колонкам
              //  for (int j = 1; j < lastRow; j++){ // по всем строкам 
                str[i] = "UPDATE 'FCPObjectHistory' SET  'FCPObjectName'=" + list[1, i] + ", 'FundingDirection' =" +
                         list[2, i] + ", 'Name'=" + list[3, i] + ", 'Number'=" + list[4, i] + ", 'Order' = " + list[5, i] +
                         " WHERE OID = " + list[0, i] + ";";

            } 
    foreach ( string a in str)
            {
             Console.WriteLine(a);   
            }
            Console.ReadLine();
        }

    }
}
