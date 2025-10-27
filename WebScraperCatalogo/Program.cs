using System.Globalization;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using System.IO;
using System;
using System.Threading.Tasks;

public class Producto
{
    public string Nombre { get; set; }
    // Precio de la página (Precio base)
    public decimal PrecioBase { get; set; }
    // Precio Mayorista (PrecioBase + 30%)
    public decimal PrecioMayorista { get; set; }
    // Precio Minorista (PrecioBase + 100%)
    public decimal PrecioMinorista { get; set; }
    public string UrlOrigen { get; set; }
}

public partial class Program
{
    public static async Task Main(string[] args)
    {
        // --- CONFIGURACIÓN DE LA EXTRACCIÓN ---
        string urlBase = "https://www.santerialacatedral.com.ar/products/category/aromanza-1";
        int totalPaginas = 6;

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string archivoSalida = Path.Combine(desktopPath, "ProductosConPreciosCalculados.xlsx"); // Nombre de archivo actualizado

        var todosLosProductos = new List<Producto>();

        Console.WriteLine($"Iniciando extracción de datos de {totalPaginas} páginas del catálogo...");

        for (int page = 1; page <= totalPaginas; page++)
        {
            string urlPagina = $"{urlBase}?page={page}";
            Console.WriteLine($"\n-> Extrayendo página {page} de {totalPaginas}: {urlPagina}");

            var productosPagina = ScrapearCatalogoSelenium(urlPagina);

            if (productosPagina.Any())
            {
                Console.WriteLine($"   - Se encontraron {productosPagina.Count} productos.");
                todosLosProductos.AddRange(productosPagina);
            }
            else
            {
                Console.WriteLine($"   - ¡Advertencia! No se encontraron productos en la página {page}.");
            }
        }

        // --- PROCESO FINAL ---

        if (todosLosProductos.Any())
        {
            Console.WriteLine($"\n✅ EXTRACCIÓN FINALIZADA. TOTAL DE PRODUCTOS: {todosLosProductos.Count}");
            ExportarAExcel(todosLosProductos, archivoSalida);
        }
        else
        {
            Console.WriteLine("\n❌ Fallo en la extracción de productos de todas las páginas.");
        }

        Console.WriteLine("Proceso finalizado. Presiona cualquier tecla para salir...");
        Console.ReadKey();
    }

    public static List<Producto> ScrapearCatalogoSelenium(string urlCatalogo)
    {
        var productos = new List<Producto>();

        // Configuración de WebDriver
        new DriverManager().SetUpDriver(new ChromeConfig());
        var options = new ChromeOptions();
        options.AddArgument("--headless");
        options.AddArgument("--disable-gpu");
        options.AddArgument("--window-size=1920,1080");

        using (var driver = new ChromeDriver(options))
        {
            try
            {
                driver.Navigate().GoToUrl(urlCatalogo);
                System.Threading.Thread.Sleep(12000);

                var nodosProductos = driver.FindElements(By.XPath("//div[contains(@class, 'product-default')]"));

                foreach (var nodoProducto in nodosProductos)
                {
                    IWebElement nombreElemento = null;
                    IWebElement precioElemento = null;

                    try
                    {
                        nombreElemento = nodoProducto.FindElement(By.XPath(".//a[@class='default-text-product']"));
                        precioElemento = nodoProducto.FindElement(By.XPath(".//span[@class='product-price']"));
                    }
                    catch (NoSuchElementException)
                    {
                        continue;
                    }

                    if (nombreElemento != null && precioElemento != null)
                    {
                        string nombre = nombreElemento.Text.Trim();
                        string precioTextoLimpio = precioElemento.Text
                            .Split('\n')[0]
                            .Replace("$", "")
                            .Replace(" ", "")
                            .Trim();

                        if (decimal.TryParse(
                            precioTextoLimpio,
                            NumberStyles.Currency,
                            new CultureInfo("es-AR"),
                            out decimal precioBase))
                        {
                            // 💡 CÁLCULO DE PRECIOS
                            decimal precioMayorista = precioBase * 1.30m; // 30% adicional
                            decimal precioMinorista = precioBase * 2.00m; // 100% adicional (el doble)

                            productos.Add(new Producto
                            {
                                Nombre = nombre,
                                PrecioBase = precioBase,
                                PrecioMayorista = precioMayorista,
                                PrecioMinorista = precioMinorista,
                                UrlOrigen = urlCatalogo
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocurrió un error inesperado durante el scraping de {urlCatalogo}: {ex.Message}");
            }
        }
        return productos;
    }

    public static void ExportarAExcel(List<Producto> datos, string rutaArchivo)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Catálogo Precios");

            // Configurar encabezados
            worksheet.Cell(1, 1).Value = "Nombre";
            worksheet.Cell(1, 2).Value = "Precio Base (Web)";
            worksheet.Cell(1, 3).Value = "Precio Mayorista (+30%)";
            worksheet.Cell(1, 4).Value = "Precio Minorista (+100%)";
            worksheet.Cell(1, 5).Value = "URL Origen";

            // Llenar datos
            int fila = 2;
            foreach (var producto in datos)
            {
                worksheet.Cell(fila, 1).Value = producto.Nombre;
                worksheet.Cell(fila, 2).Value = producto.PrecioBase;
                worksheet.Cell(fila, 3).Value = producto.PrecioMayorista;
                worksheet.Cell(fila, 4).Value = producto.PrecioMinorista;
                worksheet.Cell(fila, 5).Value = producto.UrlOrigen;

                fila++;
            }

            // Aplicar formato de moneda a las columnas de precios (2, 3 y 4)
            worksheet.Columns("B:D").Style.NumberFormat.Format = "$ #,##0.00";

            // Ajustar el ancho de las columnas
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(rutaArchivo);
        }
        Console.WriteLine($"✅ Exportación exitosa. Archivo guardado en: {Path.GetFullPath(rutaArchivo)}");
    }
}