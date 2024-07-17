using CsvHelper;
using CsvHelper.Configuration;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TramitesComunicacion.Models;
using TramitesComunicacion.ViewModels;

namespace TramitesComunicacion
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            WebServiceClient webServiceClient = new WebServiceClient();

            // Verificar si se pasaron argumentos para ejecución automática
            if (args.Length > 0)
            {
                switch (args[0].ToLower())
                {
                    case "consultarportelefono":
                        await ConsultarPorTelefono(webServiceClient);
                        return; // Termina la ejecución después de la tarea automatizada
                    default:
                        Console.WriteLine($"Argumento no reconocido: {args[0]}");
                        return;
                }
            }

            // Interfaz de usuario para ejecución manual
            Console.WriteLine("Seleccione el tipo de consulta:");
            Console.WriteLine("1. Consulta por Correo");
            Console.WriteLine("2. Consulta por Teléfono");
            Console.Write("Ingrese su opción: ");
            string option = Console.ReadLine();

            try
            {
                switch (option)
                {
                    case "1":
                        await ConsultarPorCorreo(webServiceClient);
                        break;
                    case "2":
                        await ConsultarPorTelefono(webServiceClient);
                        break;
                    default:
                        Console.WriteLine("Opción no válida. Por favor, seleccione 1 o 2.");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocurrió un error: {ex.Message}");
            }

            Console.WriteLine("Presiona cualquier tecla para salir...");
            Console.ReadLine();
        }
        static async Task ConsultarPorCorreo(WebServiceClient webServiceClient)
        {
            Console.WriteLine("Ingrese los correos electrónicos separados por comas:");
            string emailInput = Console.ReadLine();
            string[] emails = emailInput.Split(',');
            string resultByEmail = await webServiceClient.ConsultarRnePorCorreoAsync(emails);
            Console.WriteLine("Resultado de la consulta por correo:");
            Console.WriteLine(resultByEmail);
        }
        static async Task ConsultarPorTelefono(WebServiceClient webServiceClient)
        {
            Console.WriteLine("Seleccione la fuente de los números de teléfono:");
            Console.WriteLine("1. Ingresar manualmente");
            Console.WriteLine("2. Leer desde archivo CSV");
            Console.Write("Ingrese su opción: ");
            string phoneOption = Console.ReadLine();

            string[] telefonos = null;

            if (phoneOption == "1")
            {
                Console.WriteLine("Ingrese los números de teléfono separados por comas:");
                string phoneInput = Console.ReadLine();
                telefonos = phoneInput.Split(',');
            }
            else if (phoneOption == "2")
            {
                string rutaArchivo = @"C:\Recamier\data.csv";
                telefonos = await LeerTelefonosDesdeCsvAsync(rutaArchivo);
            }
            else
            {
                Console.WriteLine("Opción no válida. Por favor, seleccione 1 o 2.");
                return;
            }

            string resultByPhone = await webServiceClient.ConsultarRnePorTelefonoAsync(telefonos);
            Console.WriteLine("Resultado de la consulta por teléfono:");
            Console.WriteLine(resultByPhone);

            ExportarJsonAExcel(resultByPhone, @"C:\Recamier\ResultadosConsulta.xlsx");

        }
        public static async Task<string[]> LeerTelefonosDesdeCsvAsync(string rutaArchivo)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",
                HasHeaderRecord = false
            };

            var telefonos = new ConcurrentBag<string>();

            try
            {
                using (var reader = new StreamReader(rutaArchivo))
                using (var csv = new CsvReader(reader, config))
                {
                    // Lee el archivo de manera asíncrona
                    while (await csv.ReadAsync())
                    {
                        var telefono = csv.GetField<string>(0);
                        telefonos.Add(telefono);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al leer el archivo CSV: {ex.Message}");
            }

            return telefonos.ToArray();
        }
        public static void ExportarJsonAExcel(string jsonData, string rutaArchivo)
        {
            var datos = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonData);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(rutaArchivo);

            if (file.Exists)
            {
                file.Delete();
            }

            using (var package = new ExcelPackage(file))  // Se crea un nuevo archivo
            {
                var worksheet = package.Workbook.Worksheets.Add("Datos");

                int columnIndex = 1;
                foreach (var key in datos[0].Keys)
                {
                    if (key == "opcionesContacto")
                    {
                        worksheet.Cells[1, columnIndex].Value = "Sms";
                        columnIndex++;
                        worksheet.Cells[1, columnIndex].Value = "Aplicacion";
                        columnIndex++;
                        worksheet.Cells[1, columnIndex].Value = "Llamada";
                    }
                    else
                    {
                        string header = key == "llave" ? "Telefono" : Capitalize(key);
                        worksheet.Cells[1, columnIndex].Value = header;
                    }
                    columnIndex++;
                }

                int rowIndex = 2;
                foreach (var item in datos)
                {
                    columnIndex = 1;
                    foreach (var key in item.Keys)
                    {
                        if (key == "opcionesContacto")
                        {
                            var opciones = JsonConvert.DeserializeObject<Dictionary<string, bool>>(item[key].ToString());
                            worksheet.Cells[rowIndex, columnIndex].Value = opciones["sms"];
                            columnIndex++;
                            worksheet.Cells[rowIndex, columnIndex].Value = opciones["aplicacion"];
                            columnIndex++;
                            worksheet.Cells[rowIndex, columnIndex].Value = opciones["llamada"];
                        }
                        else
                        {
                            worksheet.Cells[rowIndex, columnIndex].Value = item[key]?.ToString();
                        }
                        columnIndex++;
                    }
                    rowIndex++;
                }

                package.Save();
            }
        }
        public static string Capitalize(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            text = text.ToLower();
            string[] words = text.Split(new char[] { ' ', '_', '-' }, System.StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = words[i].Substring(0, 1).ToUpper() + words[i].Substring(1);
            }
            return string.Join(" ", words);
        }
    }
}
