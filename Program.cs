using ImageMagick;
using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Threading;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using ICSharpCode.SharpZipLib.Zip;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace gaso_downloader {
    internal class Program {
        // Конфигурационные поля
        static string email;
        static string password;
        static string baseSavePath; // Вместо жестко заданного "Z:\\gaso"
        
        static IWebDriver driver;
        static ChromeOptions options;
        static WebDriverWait wait;
        static List<string> filesToSave = new List<string>();
        
        // Текущее дело (folderName) и путь для сохранения
        static string saveFolder;
        static string folderName;
        
        static void Main(string[] args){
            Console.WriteLine("ГАСО Смоленск Качалка v 1.51");
            
            // Загружаем настройки из config.txt
            LoadConfig("config.txt");
            
            new DriverManager().SetUpDriver(new ChromeConfig());
            
            List<(string fond,string opis,string delo)> tasks = LoadTasksFromFile("list.txt");
            
            foreach(var (fond,opis,delo) in tasks){
                folderName = $"{fond}-{opis}-{delo}";
                saveFolder = GetSaveFolder(fond, opis, delo);
                
                string zipFile = Path.Combine(saveFolder, folderName + ".zip");
                if(File.Exists(zipFile)){
                    Console.WriteLine($"ZIP найден ({zipFile}), пропускаем дело {fond}-{opis}-{delo}.");
                    continue;
                }
                bool success = false;
                const int delaySeconds = 5;
                
                while(!success){
                    Console.WriteLine($"Обрабатываем дело: {fond}-{opis}-{delo}");
                    StartBrowser();
                    try{
                        ProcessDelo(fond, opis, delo);
                        if(!File.Exists(zipFile)) {
                            throw new Exception("ZIP-файл не создан. Повторная попытка.");
                        }
                        success = true;
                    }
                    catch(Exception ex){
                        Console.WriteLine($"Ошибка: {ex.Message}. Повторная попытка через {delaySeconds} секунд.");
                        driver?.Quit();
                        Thread.Sleep(delaySeconds * 1000);
                    }
                }
            }
            Console.WriteLine("Все дела обработаны.");
        }
        
        /// <summary>
        /// Считываем конфигурационные данные из config.txt
        /// </summary>
        static void LoadConfig(string configFilePath){
            if(!File.Exists(configFilePath)){
                throw new FileNotFoundException($"Не найден файл конфигурации: {configFilePath}");
            }
            var lines = File.ReadAllLines(configFilePath);
            foreach(var line in lines){
                if(string.IsNullOrWhiteSpace(line)) continue;
                var parts = line.Split('=');
                if(parts.Length != 2) continue;
                
                var key = parts[0].Trim().ToLower();
                var value = parts[1].Trim();
                
                switch(key){
                    case "email":
                        email = value;
                        break;
                    case "password":
                        password = value;
                        break;
                    case "savepath":
                        baseSavePath = value;
                        break;
                }
            }
            
            // Простая проверка, что ключевые поля заполнены
            if(string.IsNullOrEmpty(email) 
               || string.IsNullOrEmpty(password) 
               || string.IsNullOrEmpty(baseSavePath)){
                throw new Exception("В config.txt не заполнены все обязательные параметры (email, password, savepath).");
            }
        }

        static List<(string fond,string opis,string delo)> LoadTasksFromFile(string filePath){
            var tasks = new List<(string fond,string opis,string delo)>();
            foreach(string line in File.ReadLines(filePath)){
                string[] parts = line.Split('\t');
                if(parts.Length!=3) continue;
                tasks.Add((parts[0].Trim(), parts[1].Trim(), parts[2].Trim()));
            }
            return tasks;
        }
        
        static void StartBrowser(){
            driver?.Quit();
            Thread.Sleep(300);
            options = new ChromeOptions();
            options.AddArgument("--ignore-certificate-errors");
            options.AddArgument("--allow-insecure-localhost");
            options.AddArgument("--headless=new");
            options.AddArgument("--disable-gpu");
            
            driver = new ChromeDriver(options);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
        }

        static void StartDelo(string fond, string opis, string delo){
            Console.WriteLine($"Запуск загрузки дела {fond}-{opis}-{delo}...");
            
            // Переходим на сайт
            driver.Navigate().GoToUrl("https://portal.gaso-smolensk.ru/page/sources/archivefund.jsf");
            
            // Логинимся, используя конфигурационные поля email и password
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//label[contains(text(), 'Электронная почта')]/preceding-sibling::input")))
                .SendKeys(email);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@type='password']")))
                .SendKeys(password);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[text()='Войти']"))).Click();
            
            // Открываем форму поиска
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[text()='Поиск по шифру/заголовку']"))).Click();
            
            // Заполняем поля поиска
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//label[text()='Фонд']/preceding-sibling::input"))).SendKeys(fond);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//label[text()='Опись']/preceding-sibling::input"))).SendKeys(opis);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//label[text()='Дело']/preceding-sibling::input"))).SendKeys(delo);
            
            // Нажимаем Найти
            IWebElement searchButton = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[text()='Найти']")));
            searchButton.Click();
            wait.Until(ExpectedConditions.StalenessOf(searchButton));
            
            Console.WriteLine($"Проверяем {folderName}...");
        }
        
        static void ClickElementByJs(IWebElement element){
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", element);
        }

        static void ProcessDelo(string fond, string opis, string delo){
            StartDelo(fond, opis, delo);
            
            string zipFile = Path.Combine(saveFolder, folderName + ".zip");
            if(File.Exists(zipFile)){
                Console.WriteLine("ZIP уже существует, дело пропускается.");
                return;
            }
            
            int totalScaned = GetTotal();
            Console.WriteLine($"Нужно скачать файлов {totalScaned}.");
            
            DeleteImgs();
            
            // Переход к просмотру сканов
            try{
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='cipherHeader']/a)[1]"))).Click();
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//span[text()='Просмотр сканов']/ancestor::button"))).Click();
                
                var button = wait.Until(ExpectedConditions.ElementExists(By.XPath("//a[contains(@id, ':0:') and contains(@onclick, 'watermarkPreviewPanel')]")));
                ClickElementByJs(button);
            }
            catch(WebDriverTimeoutException ex){
                Console.WriteLine($"Ошибка при открытии дела: {ex.Message}, прерывание дела.");
                return;
            }
            
            int slideNumber = 1;
            if(!SaveCurrentSlide(slideNumber.ToString())){
                Console.WriteLine("Ошибка при сохранении первого слайда, прерывание дела");
                return;
            }
            
            // Листаем слайды
            while(true){
                if(Directory.GetFiles(saveFolder, "*.png").Length >= totalScaned) break;
                
                if(!GoToNextSlide()){
                    Console.WriteLine("Ошибка при переходе на следующий слайд, прерывание дела");
                    return;
                }
                
                slideNumber++;
                if(!SaveCurrentSlide(slideNumber.ToString())){
                    Console.WriteLine($"Ошибка при сохранении слайда {slideNumber}, прерывание дела");
                    return;
                }
            }
            driver.Quit();
            
            int finalTotalPNGs = Directory.GetFiles(saveFolder, "*.png").Length;
            Console.WriteLine($"Всего скачано файлов {finalTotalPNGs}.");
            if(finalTotalPNGs != totalScaned){
                Console.WriteLine("Не все файлы загружены, прерывание дела");
                return;
            }
            
            // Если было сохранение без JPG, то конвертируем в JPG -> PDF
            if(filesToSave.Count == 0){
                ConvertAllPngToJpg();
                SaveToPdf();
            }
            
            if(File.Exists(Path.Combine(saveFolder, folderName + ".pdf")) && !File.Exists(zipFile)){
                Console.WriteLine("Создаём ZIP-архив...");
                SaveToZip();
            }
            Console.WriteLine("Конец программы.");
        }

        /// <summary>
        /// Ожидаем появления и нажатия кнопки "arrow-right.png" максимум 5 секунд.
        /// Каждые 10 мс проверяем, чтобы реагировать быстрее.
        /// </summary>
        static bool GoToNextSlide(){
            const int totalMs = 5000;
            const int step = 10;
            int attempts = totalMs / step;
            for(int i=0; i<attempts; i++){
                try{
                    var nextButton = driver.FindElement(By.XPath("//img[contains(@src, 'arrow-right.png')]"));
                    if(nextButton.Displayed && nextButton.Enabled){
                        ClickElementByJs(nextButton);
                        return true;
                    }
                }
                catch(Exception){}
                Thread.Sleep(step);
            }
            Console.WriteLine("Не удалось нажать кнопку «следующий слайд» за 5 секунд.");
            return false;
        }

        /// <summary>
        /// Сохраняем текущий слайд. Выводим остаток времени (если есть) из label id="form:j_idt539".
        /// </summary>
        static bool SaveCurrentSlide(string slideNumber){
            try{
                string previousStyle = driver.FindElement(By.ClassName("scanImage")).GetAttribute("style");
                if(!WaitForStyleChange(previousStyle)){
                    Console.WriteLine($"Ошибка: Слайд {slideNumber} не сменился вовремя.");
                    return false;
                }

                IWebElement imageElement = driver.FindElement(By.ClassName("scanImage"));
                string styleWithImage = imageElement.GetAttribute("style");

                var match = Regex.Match(styleWithImage, @"data:image\/[a-zA-Z]+;base64,([^""]+)");
                if(!match.Success){
                    Console.WriteLine($"Ошибка: Base64-код для слайда {slideNumber} не найден.");
                    return false;
                }

                string filePath = Path.Combine(saveFolder, $"{slideNumber}.png");
                string base64Data = match.Groups[1].Value.Trim();
                File.WriteAllBytes(filePath, Convert.FromBase64String(base64Data));
                
                string timeLeftText = "";
                try{
                    var timeLabel = driver.FindElement(By.Id("form:j_idt539"));
                    if(timeLabel != null){
                        string raw = timeLabel.Text.Trim();
                        if(raw.StartsWith("Осталось "))
                            raw = raw.Substring("Осталось ".Length);
                        timeLeftText = raw;
                    }
                }
                catch(NoSuchElementException){}
                
                Console.WriteLine(
                    string.IsNullOrEmpty(timeLeftText) 
                    ? $"Слайд {filePath} успешно сохранен." 
                    : $"Слайд {filePath} успешно сохранен. Осталось времени {timeLeftText}"
                );
                return true;
            }
            catch(Exception ex){
                Console.WriteLine($"Неожиданная ошибка на слайде {slideNumber}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Ждём до 5 секунд, проверяя каждые 10 мс, сменился ли style у .scanImage.
        /// </summary>
        static bool WaitForStyleChange(string oldStyle){
            const int totalMs = 5000;
            const int step = 10;
            int attempts = totalMs / step;
            for(int i=0; i<attempts; i++){
                try{
                    string currentStyle = driver.FindElement(By.ClassName("scanImage")).GetAttribute("style");
                    if(currentStyle != oldStyle) return true;
                }
                catch(StaleElementReferenceException){}
                Thread.Sleep(step);
            }
            return false;
        }

        /// <summary>
        /// Формируем путь для сохранения исходя из baseSavePath из config.txt.
        /// </summary>
        static string GetSaveFolder(string fond, string opis, string delo){
            string folderPath = Path.Combine(baseSavePath, folderName);
            if(!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);
            return folderPath;
        }

        static void ConvertAllPngToJpg(){
            var pngFiles = Directory.GetFiles(saveFolder, "*.png")
                .OrderBy(f => int.Parse(Path.GetFileNameWithoutExtension(f)))
                .ToArray();
            foreach(var png in pngFiles){
                string jpg = Path.ChangeExtension(png, ".jpg");
                try{
                    using(var img = new MagickImage(png)){
                        img.Format = MagickFormat.Jpg;
                        img.Quality = 80;
                        img.Write(jpg);
                    }
                    File.Delete(png);
                }
                catch(Exception ex){
                    Console.WriteLine($"Ошибка конвертации {png} -> JPG: {ex.Message}");
                }
            }
        }

        static void SaveToPdf(){
            try{
                string pdfPath = Path.Combine(saveFolder, folderName + ".pdf");
                if(File.Exists(pdfPath)){
                    Console.WriteLine("PDF уже существует, пропускаем создание.");
                    return;
                }
                string[] imageFiles = Directory.GetFiles(saveFolder, "*.jpg")
                    .OrderBy(f => int.Parse(Path.GetFileNameWithoutExtension(f)))
                    .ToArray();
                if(imageFiles.Length == 0){
                    Console.WriteLine("Ошибка: Нет файлов для создания PDF.");
                    return;
                }
                
                PdfDocument document = new PdfDocument();
                foreach(string imageFile in imageFiles){
                    XImage img = XImage.FromFile(imageFile);
                    PdfPage page = document.AddPage();
                    page.Width = img.PixelWidth;
                    page.Height = img.PixelHeight;

                    using(XGraphics gfx = XGraphics.FromPdfPage(page)){
                        gfx.DrawImage(img, 0, 0, img.PixelWidth, img.PixelHeight);
                    }
                    img.Dispose();
                }
                document.Save(pdfPath);
                Console.WriteLine($"PDF успешно сохранен: {pdfPath}");
            }
            catch(Exception ex){
                Console.WriteLine($"Ошибка при создании PDF: {ex.Message}");
            }
        }

        static void SaveToZip(){
            try{
                string zipPath = Path.Combine(saveFolder, folderName + ".zip");
                if(File.Exists(zipPath)){
                    Console.WriteLine("ZIP уже существует, создание не требуется.");
                    return;
                }
                string[] imageFiles = Directory.GetFiles(saveFolder, "*.jpg")
                    .OrderBy(f => int.Parse(Path.GetFileNameWithoutExtension(f)))
                    .ToArray();
                if(imageFiles.Length == 0){
                    Console.WriteLine("Ошибка: Нет файлов для архивации.");
                    return;
                }
                using(FileStream fs = new FileStream(zipPath, FileMode.Create))
                using(ZipOutputStream zipStream = new ZipOutputStream(fs)){
                    zipStream.SetLevel(9);
                    foreach(string file in imageFiles){
                        FileInfo fi = new FileInfo(file);
                        ZipEntry newEntry = new ZipEntry(fi.Name){
                            DateTime = fi.LastWriteTime,
                            Size = fi.Length
                        };
                        zipStream.PutNextEntry(newEntry);
                        using(FileStream fileStream = File.OpenRead(file)){
                            fileStream.CopyTo(zipStream);
                        }
                        zipStream.CloseEntry();
                    }
                    zipStream.IsStreamOwner = true;
                }
                Console.WriteLine($"ZIP успешно сохранен: {zipPath}");
                Console.WriteLine("Удаляем JPG-файлы...");
                DeleteImgs();
                Console.WriteLine("Все JPG-файлы удалены после создания ZIP.");
            }
            catch(Exception ex){
                Console.WriteLine($"Ошибка при создании ZIP: {ex.Message}");
            }
        }

        static void DeleteImgs(){
            foreach(var file in Directory.GetFiles(saveFolder, "*.png")) File.Delete(file);
            foreach(var file in Directory.GetFiles(saveFolder, "*.jpg")) File.Delete(file);
        }

        static int GetTotal(){
            var text = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[contains(@class, 'act-data-inner')]"))).Text;
            var match = Regex.Match(text, @"Отсканировано:\s*(\d+)");
            return match.Success ? int.Parse(match.Groups[1].Value) : 0;
        }
    }
}
