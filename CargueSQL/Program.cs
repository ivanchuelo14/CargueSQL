using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;

class Program
{
    static void Main()
    {
        string filePath = @"C:\\WORK\\CompetenciasCargue.xlsx"; // Cambia esto a la ruta de tu archivo
        string connectionString = "Data Source=COBOGLTW1130007;Initial Catalog=TalentScore;User ID=sa;Password=Bogota.2024*; MultipleActiveResultSets=True; TrustServerCertificate=True"; // Ajusta tu cadena de conexión
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Obtener la primera hoja
            int rows = worksheet.Dimension.Rows;

            // Diccionario para agrupar competencias y comportamientos
            var competencias = new Dictionary<string, Competencia>();

            // Procesar las filas del Excel
            for (int row = 2; row <= rows; row++) // Saltar la fila de encabezados
            {
                var id = worksheet.Cells[row, 1].Value?.ToString();
                var nombre = worksheet.Cells[row, 2].Value?.ToString();
                var descripcion = worksheet.Cells[row, 3].Value?.ToString();
                var comportamiento = worksheet.Cells[row, 4].Value?.ToString();

                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(nombre))
                    continue;

                if (!competencias.ContainsKey(id))
                {
                    competencias[id] = new Competencia
                    {
                        Nombre = nombre,
                        Descripcion = descripcion,
                        Comportamientos = new List<Comportamiento>()
                    };
                }

                if (!string.IsNullOrEmpty(comportamiento))
                {
                    competencias[id].Comportamientos.Add(new Comportamiento { Nombre = comportamiento });
                }
            }

            // Insertar en la base de datos
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                foreach (var competencia in competencias.Values)
                {
                    string insertQuery = @"INSERT INTO CompetenciaPreDefinicionTs 
                        (CmPDNombre, CmPDDescripcion, CmPDComportamientos, CreateTime, UpdateTime, CreatedBy, UpdatedBy) 
                        VALUES (@Nombre, @Descripcion, @Comportamientos, GETDATE(), GETDATE(), 'Ivan Guerra', 'Ivan Guerra');";

                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@Nombre", competencia.Nombre ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Descripcion", competencia.Descripcion ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Comportamientos", JsonConvert.SerializeObject(competencia.Comportamientos));

                        command.ExecuteNonQuery();
                    }
                }
            }

            Console.WriteLine("Datos importados exitosamente.");
        }
    }

    public class Competencia
    {
        public string Nombre { get; set; }
        public string Descripcion { get; set; }
        public List<Comportamiento> Comportamientos { get; set; }
    }

    public class Comportamiento
    {
        public string Nombre { get; set; }
    }

    static string CleanString(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        // Eliminar tildes y otros diacríticos
        string normalized = input.Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder();

        foreach (var c in normalized)
        {
            var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
            if (unicodeCategory != UnicodeCategory.NonSpacingMark && c != ' ')
            {
                builder.Append(c);
            }
        }

        return builder.ToString().Normalize(NormalizationForm.FormC);
    }
}
