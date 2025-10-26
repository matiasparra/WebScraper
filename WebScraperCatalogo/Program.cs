using System.Globalization;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using System.IO;
using System; // Agregamos System para usar Environment

public class Producto
{
    public string Nombre { get; set; }
    public decimal Precio { get; set; }
    public string UrlOrigen { get; set; }
}

public partial class Program
{
    public static async Task Main(string[] args)
    {
        // --- CONFIGURACIÓN DE LA EXTRACCIÓN ---

        // La URL base sin el parámetro de página
        string urlBase = "https://www.santerialacatedral.com.ar/products/category/aromanza-1";
        // Número total de páginas a extraer
        int totalPaginas = 6;

        // 💡 CAMBIO CLAVE: Definir la ruta de salida al Escritorio
        // 1. Obtiene la ruta al escritorio del usuario
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        // 2. Combina la ruta del escritorio con el nombre del archivo de salida
        string archivoSalida = Path.Combine(desktopPath, "ProductosExtraidos.xlsx");

        // Lista donde se guardarán los productos de TODAS las páginas
        var todosLosProductos = new List<Producto>();

        Console.WriteLine($"Iniciando extracción de datos de {totalPaginas} páginas del catálogo...");

        // Bucle para iterar de la página 1 a la 6
        for (int page = 1; page <= totalPaginas; page++)
        {
            // Construye la URL completa con el número de página
            string urlPagina = $"{urlBase}?page={page}";
            Console.WriteLine($"\n-> Extrayendo página {page} de {totalPaginas}: {urlPagina}");

            // Llama a la función de scraping para la página actual
            var productosPagina = ScrapearCatalogoSelenium(urlPagina);

            if (productosPagina.Any())
            {
                Console.WriteLine($"   - Se encontraron {productosPagina.Count} productos.");
                // Agrega los productos de esta página a la lista general
                todosLosProductos.AddRange(productosPagina);
            }
            else
            {
                Console.WriteLine("   - ¡Advertencia! No se encontraron productos en esta página. Puede ser el final del catálogo o un error.");
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

        // Configurar y descargar el driver de Chrome automáticamente
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
                // Damos 12 segundos de espera para que cargue el JavaScript
                System.Threading.Thread.Sleep(12000);

                // 1. SELECTOR DEL CONTENEDOR DE PRODUCTO
                var nodosProductos = driver.FindElements(By.XPath("//div[contains(@class, 'product-default')]"));

                foreach (var nodoProducto in nodosProductos)
                {
                    IWebElement nombreElemento = null;
                    IWebElement precioElemento = null;

                    try
                    {
                        // 2. SELECTOR DEL NOMBRE
                        nombreElemento = nodoProducto.FindElement(By.XPath(".//a[@class='default-text-product']"));

                        // 3. SELECTOR DEL PRECIO
                        precioElemento = nodoProducto.FindElement(By.XPath(".//span[@class='product-price']"));
                    }
                    catch (NoSuchElementException)
                    {
                        continue;
                    }

                    if (nombreElemento != null && precioElemento != null)
                    {
                        string nombre = nombreElemento.Text.Trim();
                        string precioTexto = precioElemento.Text;

                        // Limpieza del texto del precio
                        string precioTextoLimpio = precioTexto
                            .Split('\n')[0]
                            .Replace("$", "")
                            .Replace(" ", "")
                            .Trim();

                        if (decimal.TryParse(
                            precioTextoLimpio,
                            NumberStyles.Currency,
                            new CultureInfo("es-AR"),
                            out decimal precio))
                        {
                            productos.Add(new Producto { Nombre = nombre, Precio = precio, UrlOrigen = urlCatalogo });
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
            var worksheet = workbook.Worksheets.Add("Catálogo");
            worksheet.Cell(1, 1).InsertTable(datos);
            worksheet.Column(2).Style.NumberFormat.Format = "$ #.##0,00";
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(rutaArchivo);
        }
        Console.WriteLine($"✅ Exportación exitosa. Archivo guardado en: {Path.GetFullPath(rutaArchivo)}");
    }
}