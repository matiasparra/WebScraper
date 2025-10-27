using System.Globalization;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using System.IO;
using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

public class Producto
{
    public string Nombre { get; set; }
    public decimal PrecioBase { get; set; }
    public decimal PrecioMayorista { get; set; }
    public decimal PrecioMinorista { get; set; }
    public string UrlOrigen { get; set; }
}

public class Categoria
{
    public string Nombre { get; set; }
    public string UrlBase { get; set; }
    public List<Producto> Productos { get; set; } = new List<Producto>();
}

public partial class Program
{
    public static async Task Main(string[] args)
    {
        string urlBaseCategorias = "https://www.santerialacatedral.com.ar/products/category/";
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string archivoSalida = Path.Combine(desktopPath, "CatalogoCompleto.xlsx");

        Console.WriteLine("Iniciando descubrimiento de categorías...");

        List<Categoria> categorias = DescubrirCategorias(urlBaseCategorias);

        if (categorias.Count == 0)
        {
            Console.WriteLine("❌ No se encontraron categorías para procesar. Finalizando.");
            Console.ReadKey();
            return;
        }

        Console.WriteLine($"✅ Se encontraron {categorias.Count} categorías. Iniciando extracción total...");

        var todasLasCategoriasExtraidas = new List<Categoria>();

        foreach (var categoria in categorias)
        {
            Console.WriteLine($"\n=======================================================");
            Console.WriteLine($"== PROCESANDO CATEGORÍA: {categoria.Nombre.ToUpper()} ==");
            Console.WriteLine($"=======================================================");

            int totalPaginas = DeterminarTotalPaginas(categoria.UrlBase);
            Console.WriteLine($"-> Total de páginas detectadas para {categoria.Nombre}: {totalPaginas}");

            for (int page = 1; page <= totalPaginas; page++)
            {
                string urlPagina = $"{categoria.UrlBase}?page={page}";
                Console.WriteLine($"   -> Extrayendo página {page} de {totalPaginas}: {urlPagina}");

                var productosPagina = ScrapearCatalogoSelenium(urlPagina);

                if (productosPagina.Any())
                {
                    Console.WriteLine($"      - Encontrados {productosPagina.Count} productos.");
                    categoria.Productos.AddRange(productosPagina);
                }
                else if (page == 1)
                {
                    Console.WriteLine($"      - ¡Advertencia! La categoría {categoria.Nombre} no contiene productos. Saltando.");
                    break;
                }
            }

            if (categoria.Productos.Any())
            {
                Console.WriteLine($"   ✅ Total extraído para {categoria.Nombre}: {categoria.Productos.Count} productos.");
                todasLasCategoriasExtraidas.Add(categoria);
            }
        }

        if (todasLasCategoriasExtraidas.Any())
        {
            ExportarAExcelPorCategoria(todasLasCategoriasExtraidas, archivoSalida);
        }
        else
        {
            Console.WriteLine("\n❌ Fallo en la extracción de productos de todas las categorías.");
        }

        Console.WriteLine("Proceso finalizado. Presiona cualquier tecla para salir...");
        Console.ReadKey();
    }

    public static List<Categoria> DescubrirCategorias(string urlBaseCategorias)
    {
        var categorias = new List<Categoria>();

        new DriverManager().SetUpDriver(new ChromeConfig());
        var options = new ChromeOptions();
        options.AddArgument("--headless");
        options.AddArgument("--disable-gpu");
        options.AddArgument("--window-size=1920,1080");

        using (var driver = new ChromeDriver(options))
        {
            try
            {
                driver.Navigate().GoToUrl(urlBaseCategorias);
                System.Threading.Thread.Sleep(15000);

                var nodosCategoria = driver.FindElements(By.XPath("//a[contains(@href, '/products/category/')]"));

                foreach (var nodo in nodosCategoria)
                {
                    string href = nodo.GetAttribute("href");
                    string nombreLimpio = nodo.Text.Trim().Replace('\n', ' ').Replace('\r', ' ').Trim();

                    if (!string.IsNullOrEmpty(nombreLimpio) && href.Contains("/products/category/"))
                    {
                        string urlLimpia = href.Split('?')[0];

                        if (nombreLimpio.ToUpper() == "ACERO" || nombreLimpio.ToUpper() == "CATEGORY IMAGE") continue;

                        categorias.Add(new Categoria
                        {
                            Nombre = nombreLimpio,
                            UrlBase = urlLimpia
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al descubrir categorías: {ex.Message}");
            }
        }
        return categorias;
    }

    public static int DeterminarTotalPaginas(string urlCategoria)
    {
        int totalPaginas = 1;

        new DriverManager().SetUpDriver(new ChromeConfig());
        var options = new ChromeOptions();
        options.AddArgument("--headless");
        options.AddArgument("--disable-gpu");
        options.AddArgument("--window-size=1920,1080");

        using (var driver = new ChromeDriver(options))
        {
            try
            {
                driver.Navigate().GoToUrl(urlCategoria);
                System.Threading.Thread.Sleep(5000);

                // 💡 NUEVO SELECTOR: Busca todos los enlaces de paginación que contengan un número (el page-link)
                var paginacionNodos = driver.FindElements(By.XPath("//ul[contains(@class, 'pagination')]//li/a[@class='page-link' and string-length(normalize-space(text())) > 0]"));

                if (paginacionNodos.Any())
                {
                    // Filtramos y convertimos a enteros
                    var numerosPagina = paginacionNodos
                        .Select(n => n.Text.Trim())
                        .Where(t => int.TryParse(t, out _))
                        .Select(int.Parse)
                        .ToList();

                    if (numerosPagina.Any())
                    {
                        totalPaginas = numerosPagina.Max();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   (Error al determinar la paginación para {urlCategoria}: {ex.Message}. Asumiendo 1 página.)");
                totalPaginas = 1;
            }
        }

        return totalPaginas;
    }

    public static List<Producto> ScrapearCatalogoSelenium(string urlCatalogo)
    {
        var productos = new List<Producto>();

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
                            decimal precioMayorista = precioBase * 1.30m;
                            decimal precioMinorista = precioBase * 2.00m;

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

    public static void ExportarAExcelPorCategoria(List<Categoria> categorias, string rutaArchivo)
    {
        using (var workbook = new XLWorkbook())
        {
            foreach (var categoria in categorias)
            {
                if (!categoria.Productos.Any()) continue;

                string nombreHoja = categoria.Nombre.Replace("/", "-").Replace(":", "").Replace(" ", "_").Trim();
                if (nombreHoja.Length > 31) nombreHoja = nombreHoja.Substring(0, 31);

                var worksheet = workbook.Worksheets.Add(nombreHoja);

                worksheet.Cell(1, 1).Value = "Nombre";
                worksheet.Cell(1, 2).Value = "Precio Base (Web)";
                worksheet.Cell(1, 3).Value = "Precio Mayorista (+30%)";
                worksheet.Cell(1, 4).Value = "Precio Minorista (+100%)";
                worksheet.Cell(1, 5).Value = "URL Origen";

                int fila = 2;
                foreach (var producto in categoria.Productos)
                {
                    worksheet.Cell(fila, 1).Value = producto.Nombre;
                    worksheet.Cell(fila, 2).Value = producto.PrecioBase;
                    worksheet.Cell(fila, 3).Value = producto.PrecioMayorista;
                    worksheet.Cell(fila, 4).Value = producto.PrecioMinorista;
                    worksheet.Cell(fila, 5).Value = producto.UrlOrigen;

                    fila++;
                }

                worksheet.Columns("B:D").Style.NumberFormat.Format = "$ #,##0.00";
                worksheet.Columns().AdjustToContents();
            }

            workbook.SaveAs(rutaArchivo);
        }
        Console.WriteLine($"✅ Exportación exitosa. Archivo con múltiples solapas guardado en: {Path.GetFullPath(rutaArchivo)}");
    }
}