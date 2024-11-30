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
        string filePath = @"C:\WORK\VistaNomina-29-11-2024.xlsx"; // Cambia esto a la ruta de tu archivo
        string connectionString = "Data Source=COBOGLTW1130007;Initial Catalog=TalentScore;User ID=sa;Password=Bogota.2024*; MultipleActiveResultSets=True; TrustServerCertificate=True"; // Ajusta tu cadena de conexión
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
                    var PrimerApellido = CleanString(worksheet.Cells[row, 1].Value?.ToString());
                    var SegundoApellido = CleanString(worksheet.Cells[row, 2].Value?.ToString());
                    var NombreEmpleado = CleanString(worksheet.Cells[row, 3].Value?.ToString());
                    var NombreCompleto = CleanString(worksheet.Cells[row, 4].Value?.ToString());
                    var Identificacion = CleanString(worksheet.Cells[row, 5].Value?.ToString());
                    var FechaNacimiento = worksheet.Cells[row, 6].Value == null ? DBNull.Value : worksheet.Cells[row, 6].Value;
                    var Genero = CleanString(worksheet.Cells[row, 7].Value?.ToString());
                    var Telefono = CleanString(worksheet.Cells[row, 8].Value?.ToString());
                    var Celular = CleanString(worksheet.Cells[row, 9].Value?.ToString());
                    var Direccion = CleanString(worksheet.Cells[row, 10].Value?.ToString());
                    var CorreoPersonal = CleanString(worksheet.Cells[row, 11].Value?.ToString());
                    var Departamento = worksheet.Cells[row, 12].Value == null ? DBNull.Value : worksheet.Cells[row, 12].Value;
                    var Ciudad = worksheet.Cells[row, 13].Value == null ? DBNull.Value : worksheet.Cells[row, 13].Value;
                    var Per_Car = worksheet.Cells[row, 14].Value == null ? DBNull.Value : worksheet.Cells[row, 14].Value;
                    var EstadoCivil = CleanString(worksheet.Cells[row, 15].Value?.ToString());
                    var Cargo = CleanString(worksheet.Cells[row, 16].Value?.ToString());
                    var Nombre = CleanString(worksheet.Cells[row, 17].Value?.ToString());
                    var FechaIngreso = worksheet.Cells[row, 18].Value == null ? DBNull.Value : worksheet.Cells[row, 18].Value;
                    //var FechaEgreso = worksheet.Cells[row, 19].Value == null ? DBNull.Value : worksheet.Cells[row, 19].Value;

                    // Consulta SQL para insertar en TempNomina
                    string insertQuery = @"INSERT INTO TempNomina (PrimerApellido, SegundoApellido, NombreEmpleado, NombreCompleto, Identificacion, FechaNacimiento, Genero, Telefono, Celular, Direccion, CorreoPersonal, Departamento, Ciudad, Per_Car, EstadoCivil, Cargo, Nombre, FechaIngreso) VALUES (
                                             @PrimerApellido, @SegundoApellido, @NombreEmpleado, @NombreCompleto, 
                                             @Identificacion, @FechaNacimiento, @Genero, @Telefono, @Celular, @Direccion, 
                                             @CorreoPersonal, @Departamento, @Ciudad, @Per_Car, @EstadoCivil, 
                                             @Cargo, @Nombre, @FechaIngreso
                                             )";

                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        // Agregar parámetros al comando
                        command.Parameters.AddWithValue("@PrimerApellido", string.IsNullOrEmpty(PrimerApellido) ? DBNull.Value : PrimerApellido);
                        command.Parameters.AddWithValue("@SegundoApellido", string.IsNullOrEmpty(SegundoApellido) ? DBNull.Value : SegundoApellido);
                        command.Parameters.AddWithValue("@NombreEmpleado", string.IsNullOrEmpty(NombreEmpleado) ? DBNull.Value : NombreEmpleado);
                        command.Parameters.AddWithValue("@NombreCompleto", string.IsNullOrEmpty(NombreCompleto) ? DBNull.Value : NombreCompleto);
                        command.Parameters.AddWithValue("@Identificacion", string.IsNullOrEmpty(Identificacion) ? DBNull.Value : Identificacion);
                        command.Parameters.AddWithValue("@FechaNacimiento", FechaNacimiento);
                        command.Parameters.AddWithValue("@Genero", string.IsNullOrEmpty(Genero) ? DBNull.Value : Genero);
                        command.Parameters.AddWithValue("@Telefono", string.IsNullOrEmpty(Telefono) ? DBNull.Value : Telefono);
                        command.Parameters.AddWithValue("@Celular", string.IsNullOrEmpty(Celular) ? DBNull.Value : Celular);
                        command.Parameters.AddWithValue("@Direccion", string.IsNullOrEmpty(Direccion) ? DBNull.Value : Direccion);
                        command.Parameters.AddWithValue("@CorreoPersonal", string.IsNullOrEmpty(CorreoPersonal) ? DBNull.Value : CorreoPersonal);
                        command.Parameters.AddWithValue("@Departamento", Departamento);
                        command.Parameters.AddWithValue("@Ciudad", Ciudad);
                        command.Parameters.AddWithValue("@Per_Car", Per_Car);
                        command.Parameters.AddWithValue("@EstadoCivil", string.IsNullOrEmpty(EstadoCivil) ? DBNull.Value : EstadoCivil);
                        command.Parameters.AddWithValue("@Cargo", string.IsNullOrEmpty(Cargo) ? DBNull.Value : Cargo);
                        command.Parameters.AddWithValue("@Nombre", string.IsNullOrEmpty(Nombre) ? DBNull.Value : Nombre);
                        command.Parameters.AddWithValue("@FechaIngreso", FechaIngreso);

                        // Ejecutar la consulta
                        command.ExecuteNonQuery();
                    }
                }
            }

            Console.WriteLine("Datos importados exitosamente.");
        }
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
