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
        string filePath = @"C:\\WORK\\VistaNomina-29-11-2024.xlsx"; // Cambia esto a la ruta de tu archivo
        string connectionString = "Data Source=COBOGLTW1130007;Initial Catalog=Correcol;User ID=sa;Password=Bogota.2024*; MultipleActiveResultSets=True; TrustServerCertificate=True"; // Ajusta tu cadena de conexión
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Obtener la primera hoja
            int rows = worksheet.Dimension.Rows;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                for (int row = 2; row <= rows; row++) // Saltar la fila de encabezados
                {
                    var ap1_emp = worksheet.Cells[row, 1].Value?.ToString() ?? "";
                    var ap2_emp = worksheet.Cells[row, 2].Value?.ToString() ?? "";
                    var nom_emp = worksheet.Cells[row, 3].Value?.ToString() ?? "";
                    var Completo = worksheet.Cells[row, 4].Value?.ToString() ?? "";
                    var cod_emp = worksheet.Cells[row, 5].Value?.ToString() ?? "";

                    // Validación y conversión de fechas (se usa la fecha mínima si es nula)
                    var fec_nac = DateTime.TryParse(worksheet.Cells[row, 6].Value?.ToString(), out DateTime parsedFechaNacimiento)
                                  ? (object)parsedFechaNacimiento
                                : (object)DBNull.Value;

                    var Genero = (worksheet.Cells[row, 7].Value?.ToString()) ?? "";
                    var tel_res = (worksheet.Cells[row, 8].Value?.ToString()) ?? "";
                    var tel_cel = (worksheet.Cells[row, 9].Value?.ToString()) ?? "";
                    var Dir = (worksheet.Cells[row, 10].Value?.ToString()) ?? "";
                    var e_mail = (worksheet.Cells[row, 11].Value?.ToString()) ?? "";
                    var e_mail_alt = (worksheet.Cells[row, 12].Value?.ToString()) ?? "";
                    var departamento = (worksheet.Cells[row, 13].Value?.ToString()) ?? "";
                    var ciu_res = (worksheet.Cells[row, 14].Value?.ToString()) ?? "";
                    var per_car = worksheet.Cells[row, 15].Value ?? 0;
                    var CIVIL = (worksheet.Cells[row, 16].Value?.ToString()) ?? "";
                    var nom_car = (worksheet.Cells[row, 17].Value?.ToString()) ?? "";
                    var ger_cod = (worksheet.Cells[row, 18].Value?.ToString()) ?? "";
                    var ger_nom = (worksheet.Cells[row, 19].Value?.ToString()) ?? "";

                    var fec_ing = DateTime.TryParse(worksheet.Cells[row, 20].Value?.ToString(), out DateTime parsedFechaIngreso)
                                  ? (object)parsedFechaIngreso
                                 : (object)DBNull.Value;

                    var fec_egr = DateTime.TryParse(worksheet.Cells[row, 21].Value?.ToString(), out DateTime parsedFechaEgreso)
                                  ? (object)parsedFechaEgreso
                                 : (object)DBNull.Value;

                    // Consulta SQL para insertar en EmpleadosNom
                    string insertQuery = @"INSERT INTO EmpleadosNom (ap1_emp, ap2_emp, nom_emp, Completo, cod_emp, fec_nac, Genero, tel_res, tel_cel, Dir, e_mail,e_mail_alt, departamento, ciu_res, per_car, CIVIL, nom_car, ger_cod, ger_nom, fec_ing, fec_egr) 
                                   VALUES (@ap1_emp, @ap2_emp, @nom_emp, @Completo, @cod_emp, @fec_nac, @Genero, @tel_res, @tel_cel, @Dir, @e_mail,@e_mail_alt, @departamento, @ciu_res, @per_car, @CIVIL, @nom_car, @ger_cod,@ger_nom, @fec_ing, @fec_egr)";

                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        // Agregar parámetros al comando
                        command.Parameters.AddWithValue("@ap1_emp", ap1_emp);
                        command.Parameters.AddWithValue("@ap2_emp", ap2_emp);
                        command.Parameters.AddWithValue("@nom_emp", nom_emp);
                        command.Parameters.AddWithValue("@Completo", Completo);
                        command.Parameters.AddWithValue("@cod_emp", cod_emp);
                        command.Parameters.AddWithValue("@fec_nac", fec_nac);
                        command.Parameters.AddWithValue("@Genero", Genero);
                        command.Parameters.AddWithValue("@tel_res", tel_res);
                        command.Parameters.AddWithValue("@tel_cel", tel_cel);
                        command.Parameters.AddWithValue("@Dir", Dir);
                        command.Parameters.AddWithValue("@e_mail", e_mail);
                        command.Parameters.AddWithValue("@e_mail_alt", e_mail_alt);
                        command.Parameters.AddWithValue("@departamento", departamento);
                        command.Parameters.AddWithValue("@ciu_res", ciu_res);
                        command.Parameters.AddWithValue("@per_car", per_car);
                        command.Parameters.AddWithValue("@CIVIL", CIVIL);
                        command.Parameters.AddWithValue("@nom_car", nom_car);
                        command.Parameters.AddWithValue("@ger_nom", ger_nom);
                        command.Parameters.AddWithValue("@ger_cod", ger_cod);
                        command.Parameters.AddWithValue("@fec_ing", fec_ing);
                        command.Parameters.AddWithValue("@fec_egr", fec_egr);

                        // Ejecutar la consulta
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
