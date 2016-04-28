using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Reflection;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office;

namespace DumpHelperLibrary
{
    public class Txt
    {
        /// <summary>
        /// Чтение xml файла
        /// </summary>
        /// <param name="pathFile"></param>
        /// <returns></returns>
        public static List<Contents> ReadXmlFile(string pathFile)
        {
            List<Contents> cont = new List<Contents>();

            StringBuilder s = new StringBuilder();            

            if (File.Exists(pathFile))
            {
                XDocument doc = XDocument.Load(pathFile);

                foreach (XElement el in doc.Root.Elements())
                {
                    if (el.Name == "content")
                    {
                        Contents contL = new Contents();

                        foreach (XElement element in el.Elements())
                        {
                            if (element.Name == "decision")
                            {
                                foreach (XAttribute cAttibutes in element.Attributes())
                                {
                                    if (cAttibutes.Name == "number")
                                    {
                                        contL.number = cAttibutes.Value;
                                    }
                                }
                            }
                            if (element.Name == "url")
                            {
                                contL.url.Add(element.Value);
                                //s += element.Value + Environment.NewLine;                                
                            }
                            if (element.Name == "domain")
                            {
                                contL.domain = element.Value;
                            }
                            if (element.Name == "ip")
                            {
                                contL.ip.Add(element.Value);
                            }
                        }
                        cont.Add(contL);
                    }
                }
            }
            return cont;
        }

        public static void SaveTxtFile(string path, string content)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            File.WriteAllText(path, content);
        }

        public static List<string> CreateTxtFile(List<Contents> contents)
        {
            List<string> pathDumps = new List<string>();

            double rowCount = 0;

            //var cc = (from con in contents select new {a = con.ip.Count(), b = con.url.Count() });
            IEnumerable<int> counts = (from con in contents select Math.Max(con.ip.Count(), con.url.Count()));
            
            foreach (int c in counts)
            {
                rowCount += c;
            }            

            int count = 0 , totalcount = 0;
            int fileCount = 0;

            double maxRows = 30000;            

            StringBuilder s = new StringBuilder();

            string format = "{0}\t{1}\t{2}\t{3}\n";            

            foreach (Contents c in contents)
            {
                if (c.ip.Count() > 1)
                {
                    foreach (string i in c.ip)
                    {
                        s.AppendFormat(format, c.number, i, c.domain, c.url.Count() == 1 ? c.url.First() : "");
                        count++;            
                    }
                }

                if (c.url.Count() > 1)
                {
                    foreach (string u in c.url)
                    {
                        s.AppendFormat(format, c.number, c.ip.Count() == 1 ? c.ip.First() : "", c.domain, u);
                        count++;                        
                    }
                }

                //if ((c.url.Count() == 1 ? true : false || c.ip.Count() == 1 ? true : false))
                    if ((c.url.Count() == 1 && c.ip.Count() == 1) || (c.url.Count() == 1 && c.ip.Count() == 0) || (c.url.Count() == 0 && c.ip.Count() == 1))
                {
                    s.AppendFormat(format, c.number, c.ip.Count() == 1 ? c.ip.First() : "", c.domain, c.url.Count() == 1 ? c.url.First() : "");
                    count++;                                           
                }

               

                if (count >= maxRows)
                {
                    fileCount++;
                    SaveTxtFile(string.Format("{0}dump{1}.txt", AppDomain.CurrentDomain.BaseDirectory, fileCount), s.ToString());
                    pathDumps.Add(string.Format("{0}dump{1}.txt", AppDomain.CurrentDomain.BaseDirectory, fileCount));
                    s = new StringBuilder();
                    totalcount += count;
                    count = 0;
                }
            }

            if (count < maxRows)
            {
                fileCount++;
                SaveTxtFile(string.Format("{0}dump{1}.txt", AppDomain.CurrentDomain.BaseDirectory, fileCount), s.ToString());
                pathDumps.Add(string.Format("{0}dump{1}.txt", AppDomain.CurrentDomain.BaseDirectory, fileCount));
                totalcount += count;
            }
            //SaveTxtFile(pathDump, s.ToString());
            return pathDumps;
        }

    /// <summary>
    /// 
    /// Создает три файла url.txt domain.txt dump.txt
    /// </summary>
    /// <param name="filePath">Полный путь с файлу xml</param>        
    public static string CreateTxtFile(string filePath)
    {
        List<Contents> cont = new List<Contents>();
        string filename = filePath;

        StringBuilder s = new StringBuilder();
        //     StringBuilder url = new StringBuilder();
        //      StringBuilder domain = new StringBuilder();

        //   string pathUrl = AppDomain.CurrentDomain.BaseDirectory + @"\url.txt";
        //  string pathDomain = AppDomain.CurrentDomain.BaseDirectory + @"\domain.txt";
        int rowCount = 0;
        string pathDump = AppDomain.CurrentDomain.BaseDirectory + @"dump.txt";
        List<string> pathDumps = new List<string>();//= AppDomain.CurrentDomain.BaseDirectory + @"dump.txt";
        pathDumps.Add(pathDump);

        if (File.Exists(filename))
        {
            XDocument doc = XDocument.Load(filename);

            foreach (XElement el in doc.Root.Elements())
            {
                if (el.Name == "content")
                {
                    Contents contL = new Contents();

                    foreach (XElement element in el.Elements())
                    {
                        if (element.Name == "decision")
                        {
                            foreach (XAttribute cAttibutes in element.Attributes())
                            {
                                if (cAttibutes.Name == "number")
                                {
                                    contL.number = cAttibutes.Value;
                                }
                            }
                        }
                        if (element.Name == "url")
                        {
                            contL.url.Add(element.Value);
                            //s += element.Value + Environment.NewLine;                                
                        }
                        if (element.Name == "domain")
                        {
                            contL.domain = element.Value;
                        }
                        if (element.Name == "ip")
                        {
                            contL.ip.Add(element.Value);
                        }
                    }
                    cont.Add(contL);
                }
            }
            if (File.Exists(pathDump))
            {
                File.Delete(pathDump);
            }
            //    File.Delete(pathUrl);
            string format = "{0}\t{1}\t{2}\t{3}\n";
            //    if (!File.Exists(pathDump) && !File.Exists(pathUrl))
            if (!File.Exists(pathDump))
            {
                // Create a file to write to.
               // int counter = 0;
                foreach (Contents c in cont)
                {
                    if (c.ip.Count() > 1)
                    {
                        foreach (string i in c.ip)
                        {
                            s.AppendFormat(format, c.number, i, c.domain, c.url.Count() == 1 ? c.url.First() : "");
                            rowCount++;
                            //url.Append(c.url.Count() == 1 ? ParseURL(c.url.First()) + Environment.NewLine : "");
                        }
                    }
                    if (c.url.Count() > 1)
                    {
                        foreach (string u in c.url)
                        {
                            s.AppendFormat(format, c.number, c.ip.Count() == 1 ? c.ip.First() : "", c.domain, u);
                            rowCount++;
                            //url.Append(ParseURL(u) + Environment.NewLine);
                        }
                    }
                    if (c.url.Count() == 1 && c.ip.Count() == 1)
                    {
                        s.AppendFormat(format, c.number, c.ip.Count() == 1 ? c.ip.First() : "", c.domain, c.url.Count() == 1 ? c.url.First() : "");
                        rowCount++;
                        //url.Append(c.url.Count() == 1 ? ParseURL(c.url.First()) + Environment.NewLine : "");                            
                    }
                    if (c.domain != null)
                    {
                        //domain.Append(ParseDomain(c.domain) + Environment.NewLine);
                    }

                    //Console.WriteLine(c.number);


                }

                File.WriteAllText(pathDump, s.ToString());

                //File.WriteAllText(pathUrl, url.ToString());

                //File.WriteAllText(pathDomain, domain.ToString());

            }
        }
        return pathDump;
    }

    public static void Parse(string xmlfile)
    {
        if (File.Exists(xmlfile))
        {
            string text = System.IO.File.ReadAllText(xmlfile);
            //string pathDomain = @"D:\domain.txt";
            string pathUrl = @"D:\url.txt";
            string path = pathUrl;


            string patternUrl = @"<url><!\[CDATA\[(?:\b[a-z\d.-]+:\/\/([^<>\s]+|\b(?:(?:(?:[^\s!@#$%^&*()_=+[\]{}\|;:,.<>\/?]+)\.)+(?:ac|ad|aero|ae|af|ag|ai|al|am|an|ao|aq|arpa|ar|asia|as|at|au|aw|ax|az|ba|bb|bd|be|bf|bg|bh|biz|bi|bj|bm|bn|bo|br|bs|bt|bv|bw|by|bz|cat|ca|cc|cd|cf|cg|ch|ci|ck|cl|cm|cn|coop|com|co|cr|cu|cv|cx|cy|cz|de|dj|dk|dm|do|dz|ec|edu|ee|eg|er|es|et|eu|fi|fj|fk|fm|fo|fr|ga|gb|gd|ge|gf|gg|gh|gi|gl|gm|gn|gov|gp|gq|gr|gs|gt|gu|gw|gy|hk|hm|hn|hr|ht|hu|id|ie|il|im|info|int|in|io|iq|ir|is|it|je|jm|jobs|jo|jp|ke|kg|kh|ki|km|kn|kp|kr|kw|ky|kz|la|lb|lc|li|lk|lr|ls|lt|lu|lv|ly|ma|mc|md|me|mg|mh|mil|mk|ml|mm|mn|mobi|mo|mp|mq|mr|ms|mt|museum|mu|mv|mw|mx|my|mz|name|na|nc|net|ne|nf|ng|ni|nl|no|np|nr|nu|nz|om|org|pa|pe|pf|pg|ph|pk|pl|pm|pn|pro|pr|ps|pt|pw|py|qa|re|ro|rs|ru|rw|sa|sb|sc|sd|se|sg|sh|si|sj|sk|sl|sm|sn|so|sr|st|su|sv|sy|sz|tc|td|tel|tf|tg|th|tj|tk|tl|tm|tn|to|tp|travel|tr|tt|tv|tw|tz|ua|ug|uk|um|us|uy|uz|va|vc|ve|vg|vi|vn|vu|wf|ws|xn--0zwm56d|xn--11b5bs3a9aj6g|xn--80akhbyknj4f|xn--9t4b11yi5a|xn--deba0ad|xn--g6w251d|xn--hgbk6aj7f53bba|xn--hlcj6aya9esc7a|xn--jxalpdlp|xn--kgbechtv|xn--zckzah|ye|yt|yu|za|zm|zw)|(?:(?:[0-9]|[1-9]\d|1\d{2}|2[0-4]\d|25[0-5])\.){3}(?:[0-9]|[1-9]\d|1\d{2}|2[0-4]\d|25[0-5]))(?:[;\/][^#?<>\s]*)?(?:\?[^#<>\s]*)?(?:#[^<>\s]*)?(?!\w)))\]\]><\/url>";
            //string patternDomain = @"<domain><!\[CDATA\[[www.]*(\S*)\]\]><\/domain>";
            string pattern = patternUrl;
            string s = @"";

            try
            {
                foreach (Match m in Regex.Matches(text, pattern))
                {
                    s += m.Groups[1].Value + Environment.NewLine;
                }
            }
            catch (Exception ex) { Console.Write(ex.ToString()); Console.ReadKey(); }
            File.Delete(path);

            if (!File.Exists(path))
            {
                File.WriteAllText(path, s);
            }
        }

    }

    public static string ParseURL(string _url)
    {
        string patternUrl = @"(?:\b[a-z\d.-]+:\/\/([^<>\s]+|\b(?:(?:(?:[^\s!@#$%^&*()_=+[\]{}\|;:,.<>\/?]+)\.)+(?:ac|ad|aero|ae|af|ag|ai|al|am|an|ao|aq|arpa|ar|asia|as|at|au|aw|ax|az|ba|bb|bd|be|bf|bg|bh|biz|bi|bj|bm|bn|bo|br|bs|bt|bv|bw|by|bz|cat|ca|cc|cd|cf|cg|ch|ci|ck|cl|cm|cn|coop|com|co|cr|cu|cv|cx|cy|cz|de|dj|dk|dm|do|dz|ec|edu|ee|eg|er|es|et|eu|fi|fj|fk|fm|fo|fr|ga|gb|gd|ge|gf|gg|gh|gi|gl|gm|gn|gov|gp|gq|gr|gs|gt|gu|gw|gy|hk|hm|hn|hr|ht|hu|id|ie|il|im|info|int|in|io|iq|ir|is|it|je|jm|jobs|jo|jp|ke|kg|kh|ki|km|kn|kp|kr|kw|ky|kz|la|lb|lc|li|lk|lr|ls|lt|lu|lv|ly|ma|mc|md|me|mg|mh|mil|mk|ml|mm|mn|mobi|mo|mp|mq|mr|ms|mt|museum|mu|mv|mw|mx|my|mz|name|na|nc|net|ne|nf|ng|ni|nl|no|np|nr|nu|nz|om|org|pa|pe|pf|pg|ph|pk|pl|pm|pn|pro|pr|ps|pt|pw|py|qa|re|ro|rs|ru|rw|sa|sb|sc|sd|se|sg|sh|si|sj|sk|sl|sm|sn|so|sr|st|su|sv|sy|sz|tc|td|tel|tf|tg|th|tj|tk|tl|tm|tn|to|tp|travel|tr|tt|tv|tw|tz|ua|ug|uk|um|us|uy|uz|va|vc|ve|vg|vi|vn|vu|wf|ws|xn--0zwm56d|xn--11b5bs3a9aj6g|xn--80akhbyknj4f|xn--9t4b11yi5a|xn--deba0ad|xn--g6w251d|xn--hgbk6aj7f53bba|xn--hlcj6aya9esc7a|xn--jxalpdlp|xn--kgbechtv|xn--zckzah|ye|yt|yu|za|zm|zw)|(?:(?:[0-9]|[1-9]\d|1\d{2}|2[0-4]\d|25[0-5])\.){3}(?:[0-9]|[1-9]\d|1\d{2}|2[0-4]\d|25[0-5]))(?:[;\/][^#?<>\s]*)?(?:\?[^#<>\s]*)?(?:#[^<>\s]*)?(?!\w)))";
        Match m = Regex.Match(_url, patternUrl);
        return m.Groups[1].Value;
    }

    public static string ParseDomain(string _url)
    {
        string patternUrl = @"[www.]*(\S*)";
        Match m;
        try
        {
            m = Regex.Match(_url, patternUrl);
            return m.Groups[1].Value;
        }
        catch (Exception ex) { }
        return "";
    }
}

public class Contents
{
    public string number { get; set; }
    public List<string> url { get; set; }
    public string domain { get; set; }
    public List<string> ip { get; set; }

    public Contents()
    {
        url = new List<string>();
        ip = new List<string>();

    }

}

/// <summary>
/// Экспорт файла txt в xls
/// </summary>
public class Excel
{

    public static void ExportTxtToExcel(string filePath)
    {
        Microsoft.Office.Interop.Excel.Application xlApp;
        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;

        xlApp = new Microsoft.Office.Interop.Excel.Application();
        //xlApp.Visible = true; //чтобы увидеть в том ли виде открылся документ                   
        object misValue = System.Reflection.Missing.Value;
        xlWorkBook = xlApp.Workbooks.Open(filePath, Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited);// 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        xlWorkBook.SaveAs(filePath.Substring(0, filePath.Length - 4) + " (" + DateTime.Now.ToShortDateString() + ").xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        xlWorkBook.Close();
    }
}
}
