namespace AbstractFactoryPattern
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
        }
    }


    /// <summary>
    /// Office 產品共同介面
    /// </summary>
    public interface IOffice
    {
        /// <summary>
        /// 產品名稱
        /// </summary>
        string ProductName { get; }

        /// <summary>
        /// 檔案名稱
        /// </summary>
        string FileName { get; }

        /// <summary>
        /// 新增檔案
        /// </summary>
        void CreateNewFile();

        /// <summary>
        /// 開啟既有檔案
        /// </summary>
        /// <param name="fileName">新檔案名稱</param>
        void OpenExistFile(string fileName);
        /// <summary>
        /// 編輯檔案
        /// </summary>
        void Edit();

        /// <summary>
        /// 存檔
        /// </summary>
        void SaveFile();

        /// <summary>
        /// 另存新檔
        /// </summary>
        /// <param name="fileName"></param>
        void SaveFileAs(string fileName);
    }

    #region   產品
    /// <summary>
    /// Word 產品介面
    /// </summary>
    public interface IWord : IOffice
    {


    }

    /// <summary>
    /// Excel 產品介面
    /// </summary>
    public interface IExcel : IOffice
    {

        /// <summary>
        /// 新增工作表
        /// </summary>
        /// <param name="sheetName">工作表名稱</param>
        void CreateWorkSheet(string sheetName);

        /// <summary>
        /// 刪除工作表
        /// </summary>
        /// <param name="sheetName">工作表名稱</param>
        void DeleteWorkSheet(string sheetName);
    }


    public class WindowsExcel : IExcel
    {
        public string ProductName { get { return "Windows Excel"; } }


        /// <summary>
        /// 檔案名稱
        /// </summary>
        public string FileName { get; set; }


        /// <summary>
        /// 新增Excel檔案
        /// </summary>
       
        public void CreateNewFile()
        {
            FileName = "新檔案";
            Console.WriteLine($"{ProductName}；開啟新檔案；檔案名稱：{FileName}");

        }

        /// <summary>
        /// 新增工作表
        /// </summary>
        /// <param name="sheetName">工作名稱</param>
        public void CreateWorkSheet(string sheetName)
        {
            Console.WriteLine($"{ProductName}-{FileName}；新增工作表；工作表名稱：{sheetName}");
        }

        /// <summary>
        /// 移除工作表
        /// </summary>
        /// <param name="sheetName">工作表名稱</param>
        public void DeleteWorkSheet(string sheetName)
        {
            Console.WriteLine($"{ProductName}-{FileName}；移除工作表；工作表名稱：{sheetName}");
        }

        /// <summary>
        /// 編輯
        /// </summary>
       public void Edit()
        {
            Console.WriteLine($"{ProductName}-{FileName}；編輯");
        }

        /// <summary>
        /// 開啟既有檔案
        /// </summary>
        /// <param name="fileName">檔案名稱</param>
        public void OpenExistFile(string fileName)
        {
            FileName = fileName;
            Console.WriteLine($"{ProductName}；開啟既有檔案；檔案名稱：{FileName}");
        }

        /// <summary>
        /// 存檔
        /// </summary>
        public void SaveFile()
        {
            Console.WriteLine($"{ProductName}；存檔；檔案名稱：{FileName}");
        }

        /// <summary>
        /// 另存新檔
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveFileAs(string fileName)
        {
            FileName= fileName;
            Console.WriteLine($"{ProductName}；另存新檔；檔案名稱：{FileName}");
        }
    }

    /// <summary>
    /// Windows Word
    /// </summary>
    public class WindowsWord : IWord
    {
        /// <summary>
        /// 產品
        /// </summary>
        public string ProductName => throw new NotImplementedException();

        public string FileName => throw new NotImplementedException();

        public void CreateNewFile()
        {
            throw new NotImplementedException();
        }

        public void CreateNewFile(string fileName)
        {
            throw new NotImplementedException();
        }

        public void Edit()
        {
            throw new NotImplementedException();
        }

        public void OpenExistFile()
        {
            throw new NotImplementedException();
        }

        public void OpenExistFile(string fileName)
        {
            throw new NotImplementedException();
        }

        public void SaveFile()
        {
            throw new NotImplementedException();
        }

        public void SaveFileAs(string fileName)
        {
            throw new NotImplementedException();
        }
    }
    #endregion 
}