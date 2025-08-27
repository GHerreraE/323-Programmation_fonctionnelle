using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;

namespace exo_marché
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // On charge le fichier excel
            WorkBook workBook = WorkBook.Load(@"C:\Users\py78dal\Documents\GitHub\323-Programmation_fonctionnelle\personnel\exos\marché\Place du marché.xlsx");

            // On select la feuille
            WorkSheet workSheet = workBook.WorkSheets[1];


            // TEST AFFICHAGE
            foreach (var cell in workSheet["A10:A20"])
            {
                Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
                
            }

            // Combien y a-t-il de vendeurs de pêche dans ce marché ?

            workSheet["C1:C100"].Value = "Pêches";


            foreach (var row in workSheet.Rows)
            {
                foreach (var col in row)
                {
                    Console.WriteLine(col.Value);
                }
            }



            Console.ReadLine();
        }
    }
}
