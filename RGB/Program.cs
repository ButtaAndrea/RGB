using System;
using System.Drawing;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main()
    {
        string filePath = "C:/Documenti/RGB.xlsx";

        
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Utenti");

            //Intestazione
            worksheet.Cell(1, 1).Value = "R";
            worksheet.Cell(1, 2).Value = "G";
            worksheet.Cell(1, 3).Value = "B";


            //Formattazione Intestazione
            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true; //Grassetto
            headerRow.Style.Fill.BackgroundColor = XLColor.Black; //Colore Sfondo Cella
            headerRow.Style.Font.FontColor = XLColor.White; //Colore del Testo

            //Rosso
            int rosso;
            do
            {
                Console.WriteLine("Inserisci il numero decimale per la colorazione del rosso");
                rosso = Convert.ToInt16(Console.ReadLine());
            }while ( rosso < 0 || rosso > 255);
            worksheet.Cell(2, 1).Value = rosso;

            //Verde
            int verde;
            do
            {
                Console.WriteLine("Inserisci il numero decimale per la colorazione del verde");
                verde = Convert.ToInt16(Console.ReadLine());
            } while (verde < 0 || verde > 255);
            worksheet.Cell(2, 2).Value = verde;

            //Blu
            int blu;
            do
            {
                Console.WriteLine("Inserisci il numero decimale per la colorazione del blu");
                blu = Convert.ToInt16(Console.ReadLine());
            } while (blu < 0 || blu > 255);
            worksheet.Cell(2, 3).Value = blu;

            //Testo
            worksheet.Cell(2, 6).Value = "Colore Ottenuto";

       
            //Colore Finale
            worksheet.Cell(3, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(rosso, verde, blu);

            //Colonna
            worksheet.Column(6).Width = 17;

            //Riga
            worksheet.Row(3).Height = 25;



            workbook.SaveAs(filePath);


        }

    }
}