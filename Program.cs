using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelProtection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path untuk menyimpan file Excel
            string filePath = "protected_workbook.xlsx";

            // Membuat workbook dan worksheet baru
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Data");

            // Menambahkan beberapa data ke worksheet
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Nama");
            row.CreateCell(1).SetCellValue("Umur");
            row.CreateCell(2).SetCellValue("Kota");

            row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("Alice");
            row.CreateCell(1).SetCellValue(30);
            row.CreateCell(2).SetCellValue("Jakarta");

            row = sheet.CreateRow(2);
            row.CreateCell(0).SetCellValue("Bob");
            row.CreateCell(1).SetCellValue(25);
            row.CreateCell(2).SetCellValue("Bandung");

            // Mengatur proteksi password untuk worksheet
            sheet.ProtectSheet("password123");

            // Menyimpan workbook ke file
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }
        }
    }
}
