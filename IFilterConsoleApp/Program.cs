using System;
using System.IO;
using System.Text;
using SDocuments = Spire.Doc.Documents;
using Spire.Doc;
using Spire.Pdf;
using Lucene.Net.Analysis.Standard;
using Version = Lucene.Net.Util.Version;
using LStore = Lucene.Net.Store;
using Lucene.Net.Index;
using LDocument = Lucene.Net.Documents.Document;
using Lucene.Net.Documents;
using Lucene.Net.QueryParsers;
using Lucene.Net.Search;
using Lucene.Net.Store;

namespace LuceneDemoConsole
{
    class Program
    {
        // Path to store index folder
        static DirectoryInfo indexDir = new DirectoryInfo(@"..\..\Index");
        static DirectoryInfo dataDir = new DirectoryInfo(@"..\..\Data");

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            /*Add references
            Lucene.Net
            ICSharpCode.SharpZipLib
            System.XML
            System.Xml.Linq*/
            // BuildIndex(indexDir, dataDir);

            /*Console.Write("Nhập keyword để tìm kiếm: ");
            string input = Console.ReadLine();

            SearchQuery(input);*/

            InitiateIndexComponent();

            Console.WriteLine("Nhấn Enter để thoát...");
            Console.ReadKey();
        }

        /*
         * Phương thức tìm kiếm văn bản
         */
        private static void SearchQuery(string keyword)
        {
            // Biến toàn cục: Nơi lưu trữ các index
            
            using(var analyzer = new StandardAnalyzer(Version.LUCENE_29))
            {
                LStore.Directory indexStore = LStore.FSDirectory.Open(indexDir);

                // Tạo truy vấn tìm kiếm
                Query query = new WildcardQuery(new Term("Content", $"*{keyword.ToLower()}*"));
                // Truyền query vào IndexSearcher
                Searcher search = new IndexSearcher(IndexReader.Open(indexStore, true));

                /*Bắt đầu tìm kiếm. Có rất nhiều cách tìm kiếm @@
                Cách 1: Tìm dựa theo số lần xuất hiện*/
                TopScoreDocCollector cllctr = TopScoreDocCollector.Create(10, true); // true: bật sắp xếp theo thứ tự
                // Lấy các kết quả đạt yêu cầu query
                var hits = search.Search(query, 100).ScoreDocs;
                // ScoreDoc[] hits = cllctr.TopDocs().ScoreDocs;

                // Vòng lặp lấy kết quả
                Console.WriteLine("Đã tìm thấy: {0} kết quả", hits.Length);
                foreach (var hit in hits)
                {
                    var foundDoc = search.Doc(hit.Doc);
                    Console.WriteLine("File found: {0}", foundDoc.Get("Filename"));
                    Console.WriteLine("File found: {0}", foundDoc.Get("Path"));
                }
            }
        }

        /*
         * Khởi tạo các đối tượng để thực hiện đánh chỉ mục
         */
        private static void InitiateIndexComponent()
        {
            using (var analyzer = new StandardAnalyzer(Version.LUCENE_29))
            {
                using (var indexDir = LStore.FSDirectory.Open(@"..\..\Index\Folders"))
                {
                    using (var indexWriter = new IndexWriter(indexDir, analyzer, IndexWriter.MaxFieldLength.UNLIMITED))
                    {
                        ScanDrivesForIndex(analyzer, indexDir, indexWriter);
                    }
                }
            }
        }

        /*
         * Được chạy khi chương trình khởi động
         * Lần thứ 2 chạy nếu thư mục Index đã có dữ liệu không cần chạy
         */
        private static void ScanDrivesForIndex(StandardAnalyzer analyzer, FSDirectory indexDir, IndexWriter indexWriter)
        {
            // TODO: Kiểm tra thư mục Index

            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                if(!drive.Name.Equals(@"C:\"))
                {
                    if (drive.DriveType == DriveType.Fixed)
                    {
                        BuildIndexFolders(drive.Name, analyzer, indexDir, indexWriter);
                    }
                }
            }
        }

        /*
         * Phương thức đánh chỉ mục các THƯ MỤC
         * Đối với thư mục con sẽ dùng đệ quy để đánh chỉ mục
         */
        private static void BuildIndexFolders(string path, StandardAnalyzer analyzer, FSDirectory indexDir, IndexWriter indexWriter)
        {
            // TODO: Thực hiện đánh chỉ mục cho path hiện tại

            try
            {
                foreach (string folder in System.IO.Directory.GetDirectories(path))
                {
                    // Console.WriteLine("Thư mục: {0}", folder);
                    BuildIndexFolders(folder, analyzer, indexDir, indexWriter);
                }
                foreach (string file in System.IO.Directory.GetFiles(path))
                {
                    // Console.WriteLine("Thư mục: {0}", file);
                    BuildIndexFiles(file, analyzer, indexDir, indexWriter);
                }

            } catch (UnauthorizedAccessException uae)
            {
                ;
            }
        }

        /*
         * Phương thức đánh chỉ mục FILE
         */
        private static void BuildIndexFiles(string file, StandardAnalyzer analyzer, FSDirectory indexDir, IndexWriter indexWriter)
        {
            StringBuilder toText = new StringBuilder();
            LDocument document;

            switch (getExtension(file))
            {
                case ".docx":
                    toText = WordToText(file);
                    break;
                case ".pdf":
                    toText = PdfToText(file);
                    break;
                case ".txt":
                    toText = TxtToText(file);
                    break;
            }

            // File Indexing
            document = new LDocument();

            document.Add(new Field("Filename", file, Field.Store.YES, Field.Index.NOT_ANALYZED));
            document.Add(new Field("Path", file, Field.Store.YES, Field.Index.NOT_ANALYZED));
            document.Add(new Field("Content", toText.ToString(), Field.Store.YES, Field.Index.ANALYZED));
            indexWriter.AddDocument(document);
            

            indexWriter.Optimize();
            indexWriter.Flush(false, false, false);
        }

        /*
         * Trả về extension của đường dẫn, file.
         */
        private static string getExtension(string file)
        {
            return Path.GetExtension(file);
        }

        /*
         * Đọc file pdf
         */
        private static StringBuilder PdfToText(string fullName)
        {
            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(fullName);

            // Initilize StringBuilder Instance
            StringBuilder toText = new StringBuilder();

            for(int page = 0; page < doc.Pages.Count; page++) {
                toText.Append(doc.Pages[page].ExtractText());
            }
            
            return toText;
        }

        private static StringBuilder TxtToText(string fullName)
        {
            StringBuilder toText = new StringBuilder();
            foreach(string line in File.ReadAllLines(fullName))
            {
                toText.AppendLine(line);
            }
            return toText;
        }

        private static StringBuilder WordToText(string fullName)
        {
            Spire.Doc.Document doc = new Spire.Doc.Document();
            doc.LoadFromFile(fullName);

            // Initilize StringBuilder Instance
            StringBuilder toText = new StringBuilder();

            //Extract Text from Word and Save to StringBuilder Instance
            foreach(Section section in doc.Sections)
            {
                foreach(SDocuments.Paragraph paragraph in section.Paragraphs)
                {
                    toText.AppendLine(paragraph.Text);
                }
            }
            return toText;
        }
    }
}
