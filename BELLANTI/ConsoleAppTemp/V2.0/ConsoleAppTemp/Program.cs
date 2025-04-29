using ConsoleAppTemp;
using System.Configuration;

Console.WriteLine("Esecuzione programma in corso...");
Run r = new Run();
string tipo= ConfigurationManager.AppSettings["Tipo"].ToString();
r.ExecProgram(tipo);