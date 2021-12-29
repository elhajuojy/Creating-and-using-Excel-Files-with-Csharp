using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Data.SqlTypes;
using System.Threading.Tasks;
using System.Data.Common;
using System.Data.SqlClient;
using System.Data;

namespace Creating_and_using__Excel_Files_with_Csharp
{
    internal class Program
    {
       

        static async Task Main(string[] args)
        {

            //so you can use this Package 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            //the path where you want to put the file and the name file + .xlsx
            var file = new FileInfo(@"C:\SaveExalFiles\ista-ntic.xlsx");



            //put the data that made in getsetupdate function here in va people
            //var people = getSetupDate();


            //await SaveExcelFile(people, file);

           List<PersonModel> pepoleFromExcel = await LoadExcelFile(file);


            int count = 0;
            foreach (var person in pepoleFromExcel)
            {
                Console.WriteLine($"{person.cin} {person.Prenom} {person.Nom} {count} ");
                count++;
               
            }


        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            DataTable table = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter();
            SqlConnection ctn = new SqlConnection(@"Data Source=ELHAJUOJY-LAPTO\MEHDI;Initial Catalog=ista_nitc_1ere_ans_info;Integrated Security=True");
            SqlCommand cmd = new SqlCommand(" select * from  Etduaint", ctn);

            da.SelectCommand = cmd;
            da.Fill(table);
           

            List<PersonModel> output = new();
            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var wordsheet = package.Workbook.Worksheets[0];

            int row = 2;
            int filierecol = 4;
            int nomCol = 10;
            int cinCol = 8;
            int massarcol =9 ;
            int prenomCol =11;
            int gendercol = 12;
            int AdresseCol= 14;
            int dataNaisCol = 13;
            int phoneCol = 15;
            
            
            while (string.IsNullOrWhiteSpace(wordsheet.Cells[row, cinCol].Value?.ToString()) == false)
            {
                PersonModel p = new();
               
                p.cin = wordsheet.Cells[row,cinCol].Value.ToString();
                p.Prenom = wordsheet.Cells[row,prenomCol].Value.ToString();
                p.Nom= wordsheet.Cells[row,nomCol].Value.ToString();
                p.Adresse =  wordsheet.Cells[row,AdresseCol].Value.ToString();
                p.filiere= wordsheet.Cells[row,filierecol].Value.ToString();
                p.gender= wordsheet.Cells[row,gendercol].Value.ToString();
                p.massar= wordsheet.Cells[row,massarcol].Value.ToString();
                p.dateNais = DateTime.Parse(wordsheet.Cells[row, dataNaisCol].Value.ToString());


                DataRow ligne = table.NewRow();
                ligne["cin"] = wordsheet.Cells[row,cinCol].Value.ToString();
                ligne["prenom"] = wordsheet.Cells[row,prenomCol].Value.ToString();
                ligne["nom"] = wordsheet.Cells[row,nomCol].Value.ToString();
                ligne["Adresse"] = wordsheet.Cells[row, AdresseCol].Value.ToString();
                ligne["filiere"] = wordsheet.Cells[row, filierecol].Value.ToString();
                ligne["gander"] = wordsheet.Cells[row, gendercol].Value.ToString();
                ligne["massar"] = wordsheet.Cells[row, massarcol].Value.ToString();
                ligne["dateNais"] = DateTime.Parse(wordsheet.Cells[row, dataNaisCol].Value.ToString());
                ligne["Phone"] = int.Parse(wordsheet.Cells[row, phoneCol].Value.ToString());

                table.Rows.Add(ligne);

                output.Add(p);

                row++;
                
            }




            SqlCommandBuilder bldr = new SqlCommandBuilder(da);
            da.Update(table);

            return output;

        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            //if the file is Exists Delete it 
            DeleteifExists(file);

            using var package = new ExcelPackage(file) ;


            //Add sheet Called MainReport
            var wordsheet = package.Workbook.Worksheets.Add("MainReport");


            var range=wordsheet.Cells["A2"].LoadFromCollection(people,true);

            //auto fit the columns with data 
            range.AutoFitColumns();

            //Format The header Row 
            wordsheet.Cells["A1"].Value = "Etudiant info";
            wordsheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            wordsheet.Cells["A1:C1"].Merge = true;
            wordsheet.Row(1).Style.Font.Size = 24;
            wordsheet.Row(1).Style.Font.Color.SetColor(Color.Blue);


            wordsheet.Row(2).Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;
            wordsheet.Row(2).Style.Font.Bold=true;

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

        private static List<PersonModel> getSetupDate()
        {
            List<PersonModel> output = new()
            {
                new() { cin = "1", Prenom = "elmahdi", Nom = "elhjuojy" },
                new() { cin = "2", Prenom = "houssam", Nom = "jebbar" },
                new() { cin = "3", Prenom = "zineb", Nom = "belhaid" },
                new() { cin = "4", Prenom = "Ahmed", Nom = "bounacer" }
            };


            return output;
        }
    }

}
