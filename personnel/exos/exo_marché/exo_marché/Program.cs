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


            // Compteur pour les vendeurs de "Pêches"
            int peches = 0;

            for (int i = 1; i <= 80; i++)
            {
                if (workSheet[$"C{i}"].StringValue == "Pêches")
                {
                    peches++;
                }
            }

            Console.WriteLine($"Nombre de vendeurs de Pêches : {peches}");
            Console.ReadLine();

            // Solution inspirée avec chatgpt
        }
    }
}
