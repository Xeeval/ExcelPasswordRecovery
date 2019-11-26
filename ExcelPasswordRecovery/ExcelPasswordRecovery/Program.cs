using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPasswordRecovery
{
    class Program
    {
        private static List<char> CaratteriConsentiti = new List<char>();
        private static char[] PasswordDaVerificare = new char[1];

        private const string PrefissoPassword = "";
        private const string FilePath = @"E:\test.xlsx";

        static void Main(string[] args)
        {
            var WApp = new Microsoft.Office.Interop.Excel.Application();

            CaratteriConsentiti.AddRange(Enumerable.Range('A', 'Z' - 'A' + 1).Select(c => (char)c));
            CaratteriConsentiti.AddRange(Enumerable.Range('a', 'z' - 'a' + 1).Select(c => (char)c));
            CaratteriConsentiti.AddRange(Enumerable.Range('0', '9' - '0' + 1).Select(c => (char)c));
            CaratteriConsentiti.AddRange(new List<char>
            {
                '!',
                '&'
                //[...]
            });

            PasswordDaVerificare[0] = CaratteriConsentiti.First();

            bool passwordTrovata = false;

            Console.WriteLine($"Inizio: {DateTime.Now}");

            do
            {
                string tmpPwd = PrefissoPassword + new string(PasswordDaVerificare);

                try
                {
                    var WDoc = WApp.Workbooks.Open(FilePath, Password: tmpPwd, ReadOnly: true);
                    passwordTrovata = true;
                }
                catch
                {
                    Console.WriteLine(tmpPwd);
                    SetNextPassword();
                }
            }
            while (!passwordTrovata);

            Console.WriteLine("===============");
            Console.WriteLine("===============");
            Console.WriteLine(PrefissoPassword + new string(PasswordDaVerificare));
            Console.WriteLine($"Fine: {DateTime.Now}");

            Console.ReadLine();
            Console.ReadLine();
            Console.ReadLine();

        }

        private static void SetNextPassword()
        {
            bool ultimoCarattere = false;

            for (int i = PasswordDaVerificare.Length - 1; i >= 0; i--)
            {
                // Si tratta dell'ultimo carattere?
                ultimoCarattere = PasswordDaVerificare[i] == CaratteriConsentiti.Last();
                if (ultimoCarattere)
                {
                    // Sì, resetto la posizione corrente
                    PasswordDaVerificare[i] = CaratteriConsentiti.First();
                }
                else
                {
                    // No, passo al carattere successivo la posizione corrente
                    // Nessun bisogno di continuare oltre, esco
                    PasswordDaVerificare[i] = CaratteriConsentiti[CaratteriConsentiti.IndexOf(PasswordDaVerificare[i]) + 1];
                    break;
                }
            }

            // Se vero allora ogni singolo carattere della password era arrivato alla fine di quelli disponibili            
            if (ultimoCarattere)
            {
                // Allargo di uno
                Array.Resize(ref PasswordDaVerificare, PasswordDaVerificare.Length + 1);
                // L'ultimo carattere appeso deve essere uguale al primo consentito
                PasswordDaVerificare[PasswordDaVerificare.Length - 1] = CaratteriConsentiti.First();
            }
        }
    }
}
