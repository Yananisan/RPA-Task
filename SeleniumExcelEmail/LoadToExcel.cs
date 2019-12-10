using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumExcelEmail
{
    class LoadToExcel
    {
        MicrowaveData microwave = new MicrowaveData();

        public void CheckIfExist()
        {
            string subPath = "C:/Temp";

            bool exists = System.IO.Directory.Exists(subPath);

            if (!exists)
                System.IO.Directory.CreateDirectory(subPath);
        }

        public void LoadDataToExcel()
        {
            CheckIfExist();

            Excel.Application ex = new Excel.Application();

            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            sheet.Name = "Microwaves";

            int i = 1;
            foreach (var microvawes in microwave.MicrowaveDataSearch())
            {
                sheet.Cells[i, 1] = microvawes.name;
                sheet.Cells[i, 2] = microvawes.price;
                sheet.Cells[i, 3] = microvawes.href;
                i++;
            }
            ex.Visible = true;
            ex.UserControl = true;

            workBook.SaveAs(@"D:\MicrowavesBook.xlsx");
            

            workBook.Close();
        }
    }
}
