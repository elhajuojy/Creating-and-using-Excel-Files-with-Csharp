using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace Excel_To_dataBase
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private async Task Form1_Load(object sender, EventArgs e, ExcelPackage excelPackage)
        {
           


            //the path where you want to put the file and the name file + .xlsx
            var file = new FileInfo(@"C:\SaveExalFiles\ExcelDataTest.xlsx");



            //put the data that made in getsetupdate function here in va people
            //var people = getSetupDate();


            // await SaveExcelFile(people, file);


            List<Etudiant> pepoleFromExcel = await LoadExcelFile(file);

            dataGridView1.DataSource = pepoleFromExcel;

        }
        private static async Task<List<Etudiant>> LoadExcelFile(FileInfo file)
        {
            List<Etudiant> output = new List<Etudiant>();
            var package = new ExcelPackage(file);
            

            await package.LoadAsync(file);

            var wordsheet = package.Workbook.Worksheets[0];

            int row = 2;
            int col = 5;

            while (string.IsNullOrWhiteSpace(wordsheet.Cells[row, col].Value?.ToString()) == false)
            {
                Etudiant p = new Etudiant();
                //p.Id = int.Parse(wordsheet.Cells[row, col].Value.ToString());
                p.cin = wordsheet.Cells[row, col].Value.ToString();
                p.Prenom = wordsheet.Cells[row, col + 2].Value.ToString();
                p.Nom = wordsheet.Cells[row, col + 1].Value.ToString();
                output.Add(p);

                row++;

            }

            return output;

        }

        private static async Task SaveExcelFile(List<Etudiant> people, FileInfo file)
        {
            //if the file is Exists Delete it 
            DeleteifExists(file);

           var package = new ExcelPackage(file);
             


            //Add sheet Called MainReport
            var wordsheet = package.Workbook.Worksheets.Add("MainReport");


            var range = wordsheet.Cells["A2"].LoadFromCollection(people, true);

            //auto fit the columns with data 
            range.AutoFitColumns();

            //Format The header Row 
            wordsheet.Cells["A1"].Value = "Our Cool Report";
            wordsheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            wordsheet.Cells["A1:C1"].Merge = true;
            wordsheet.Row(1).Style.Font.Size = 24;
            wordsheet.Row(1).Style.Font.Color.SetColor(Color.Blue);


            wordsheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            wordsheet.Row(2).Style.Font.Bold = true;

            wordsheet.Column(3).Width = 20;


            await package.SaveAsync();
        }

        private static void DeleteifExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static List<Etudiant> getSetupDate()
        {
            List<Etudiant> output = new List<Etudiant>();
            {
                //new() { cin = 1, Prenom = "elmahdi", Nom = "elhjuojy" },
                //new() { cin = 2, Prenom = "houssam", Nom = "jebbar" },
                //new() { cin = 3, Prenom = "zineb", Nom = "belhaid" },
                //new() { cin = 4, Prenom = "Ahmed", Nom = "bounacer" }
            };


            return output;
        }

        private async Task Form1_LoadAsync(object sender, EventArgs e)
        {

            //the path where you want to put the file and the name file + .xlsx
            var file = new FileInfo(@"C:\SaveExalFiles\ExcelDataTest.xlsx");



            //put the data that made in getsetupdate function here in va people
            //var people = getSetupDate();


            // await SaveExcelFile(people, file);


            List<Etudiant> pepoleFromExcel = await LoadExcelFile(file);

            dataGridView1.DataSource = pepoleFromExcel;
        }
    }
}
