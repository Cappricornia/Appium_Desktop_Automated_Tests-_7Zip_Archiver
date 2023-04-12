using NUnit.Framework;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium;


namespace AppiumArchive7ZipTest
{
    public class Tests7Zip
    {
        private const string AppiumUriString = "http://127.0.0.1:4723/wd/hub";
        private const string LocationZipFile = @"C:\Program Files\7-Zip\7zFM.exe";
        private const string tempDir = @"C:\temp";
        private WindowsDriver<WindowsElement> driver;
        private WindowsDriver<WindowsElement> driverArchiveWindow;
        private AppiumOptions options;
        private AppiumOptions optionsArchiveWindow;

        [SetUp]
        public void OpenApp()
        {
            this.options = new AppiumOptions() { PlatformName = "Windows" };
            options.AddAdditionalCapability("app", LocationZipFile);
            this.driver = new WindowsDriver<WindowsElement>(new Uri(AppiumUriString), options);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);



            this.optionsArchiveWindow = new AppiumOptions() { PlatformName = "Windows" };
            optionsArchiveWindow.AddAdditionalCapability("app", "Root");
            this.driverArchiveWindow = new WindowsDriver<WindowsElement>(new Uri(AppiumUriString), optionsArchiveWindow);
            driverArchiveWindow.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);


            if (Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, true);
            }

            Directory.CreateDirectory(tempDir);
            Thread.Sleep(1000);
        }

        [TearDown]
        public void CloseApp()
        {
            driver.Quit();
        }

        [Test]
        public void Test_7Zip_Archive()
        {
            var inputFilePath = driver.FindElementByXPath("/Window/Pane/Pane/ComboBox/Edit");
            inputFilePath.SendKeys(@"C:\Program Files\7-Zip\" + Keys.Enter);
            
           
            var filesList = driver.FindElementByClassName("SysListView32");
            filesList.SendKeys(Keys.Control + "a");

            
            var addButton = driver.FindElementByName("Add");
            addButton.Click();

            var winArch = driverArchiveWindow.FindElementByName("Add to Archive");

            var inputArchPath = winArch.FindElementByXPath("/Window/ComboBox/Edit[@Name='Archive:']");
            inputArchPath.SendKeys(@"C:\temp\archive.7z");

            var dropDownArchFormat = winArch.FindElementByXPath("/Window/ComboBox[@Name='Archive format:']");
            dropDownArchFormat.SendKeys("7z");

            var dropDownCompression = winArch.FindElementByXPath("/Window/ComboBox[@Name='Compression level:']");
            dropDownCompression.SendKeys("9 - Ultra");

            var dropDownCompressionMethod = winArch.FindElementByXPath("/Window/ComboBox[@Name='Compression method:']");
            dropDownCompressionMethod.SendKeys("*");

            var okButton = winArch.FindElementByXPath("/Window/Button[@Name='OK']");
            okButton.Click();
            Thread.Sleep(1000);

            inputFilePath.SendKeys(tempDir + @"\archive.7z" + Keys.Enter);

            var extractButton = driver.FindElementByName("Extract");
            extractButton.Click();

            var inputCopyto = driver.FindElementByXPath("/Window/Window/ComboBox/Edit[@Name='Copy to:']");
            inputCopyto.SendKeys(tempDir + Keys.Enter);
            Thread.Sleep(1000);

            FileAssert.AreEqual(LocationZipFile, @"C:\temp\7zFM.exe");
        }
    }
}