using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.IO;
using System.Linq;
using System.Net;
using EXCEL = DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using util = Palma.GestorDocumental.Repository.Common.Helpers.ObjectHelper;
using System.Threading.Tasks;

namespace Palma.GestorDocumental.Repository.Common.FileUtils
{
    public class FileHelper

    {
        private static FileHelper _instance = null;

        public static FileHelper Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new FileHelper();
                return _instance;
            }
        }

        public bool ValidarExisteArchivo(string url)
        {
            try
            {
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.KeepAlive = false;
                request.Method = "HEAD";
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                response.Close();
                return (response.StatusCode == HttpStatusCode.OK);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error, " + ex.Message);
                return false;
            }
        }

        public bool ValidarExisteArchivoFTP(string url)
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(url);
                request.Method = WebRequestMethods.Ftp.DownloadFile;
                request.Credentials = new NetworkCredential("anonymous", "");
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream);
                response.Close();
                return (response.StatusCode == FtpStatusCode.CommandOK);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error, " + ex.Message);
                return false;
            }
        }

        public string _ERROR_MENSAJE = "";

        public static string GetDataReal(EXCEL.Cell cell, WorkbookPart wbPart)
        {
            string cellValue = cell.CellValue.Text;
            if (cell.DataType != null)
            {
                if (cell.DataType == EXCEL.CellValues.SharedString)
                {
                    int id = -1;
                    if (Int32.TryParse(cell.InnerText, out id))
                    {
                        EXCEL.SharedStringItem item = GetSharedStringItemById(wbPart, id);
                        if (item.Text != null)
                        {
                            cellValue = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            cellValue = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            cellValue = item.InnerXml;
                        }
                    }
                }
            }
            if (cell.StyleIndex != null)
            {
                var cellFormat = wbPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[int.Parse(cell.StyleIndex.InnerText)] as EXCEL.CellFormat;
                if (cellFormat != null)
                {
                    var dateFormat = GetDateTimeFormat(cellFormat.NumberFormatId);
                    if (!string.IsNullOrEmpty(dateFormat))
                    {
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            if (double.TryParse(cellValue, out var cellDouble))
                            {
                                var theDate = DateTime.FromOADate(cellDouble);
                                cellValue = theDate.ToString(dateFormat);
                            }
                        }
                    }
                }
            }
            return cellValue;
        }

        private static readonly Dictionary<uint, string> DateFormatDictionary = new Dictionary<uint, string>()
        {
            [14] = "dd/MM/yyyy",
            [15] = "d-MMM-yy",
            [16] = "d-MMM",
            [17] = "MMM-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "M/d/yy h:mm",
            [30] = "M/d/yy",
            [34] = "yyyy-MM-dd",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [51] = "MM-dd",
            [52] = "yyyy-MM-dd",
            [53] = "yyyy-MM-dd",
            [55] = "yyyy-MM-dd",
            [56] = "yyyy-MM-dd",
            [58] = "MM-dd",
            [165] = "M/d/yy",
            [166] = "dd MMMM yyyy",
            [167] = "dd/MM/yyyy",
            [168] = "dd/MM/yy",
            [169] = "d.M.yy",
            [170] = "yyyy-MM-dd",
            [171] = "dd MMMM yyyy",
            [172] = "d MMMM yyyy",
            [173] = "M/d",
            [174] = "M/d/yy",
            [175] = "MM/dd/yy",
            [176] = "d-MMM",
            [177] = "d-MMM-yy",
            [178] = "dd-MMM-yy",
            [179] = "MMM-yy",
            [180] = "MMMM-yy",
            [181] = "MMMM d, yyyy",
            [182] = "M/d/yy hh:mm t",
            [183] = "M/d/y HH:mm",
            [184] = "MMM",
            [185] = "MMM-dd",
            [186] = "M/d/yyyy",
            [187] = "d-MMM-yyyy"
        };

        private static string GetDateTimeFormat(UInt32Value numberFormatId)
        {
            return DateFormatDictionary.ContainsKey(numberFormatId) ? DateFormatDictionary[numberFormatId] : string.Empty;
        }

        public string GetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;
            try
            {
                // Open the spreadsheet document for read-only access.
                using (SpreadsheetDocument document =
                    SpreadsheetDocument.Open(fileName, false))
                {
                    // Retrieve a reference to the workbook part.
                    WorkbookPart wbPart = document.WorkbookPart;
                    // Find the sheet with the supplied name, and then use that 
                    // Sheet object to retrieve a reference to the first worksheet.
                    EXCEL.Sheet theSheet = wbPart.Workbook.Descendants<EXCEL.Sheet>().
                      Where(s => s.Name == sheetName).FirstOrDefault();

                    // Throw an exception if there is no sheet.
                    if (theSheet == null)
                    {
                        throw new ArgumentException("sheetName");
                    }

                    // Retrieve a reference to the worksheet part.
                    WorksheetPart wsPart =
                        (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                    // Use its Worksheet property to get a reference to the cell 
                    // whose address matches the address you supplied.
                    EXCEL.Cell theCell = wsPart.Worksheet.Descendants<EXCEL.Cell>().
                      Where(c => c.CellReference == addressName).FirstOrDefault();

                    // If the cell does not exist, return an empty string.
                    if (theCell != null)
                    {
                        value = theCell.InnerText;

                        // If the cell represents an integer number, you are done. 
                        // For dates, this code returns the serialized value that 
                        // represents the date. The code handles strings and 
                        // Booleans individually. For shared strings, the code 
                        // looks up the corresponding value in the shared string 
                        // table. For Booleans, the code converts the value into 
                        // the words TRUE or FALSE.
                        if (theCell.DataType != null)
                        {
                            switch (theCell.DataType.Value)
                            {
                                case EXCEL.CellValues.SharedString:

                                    // For shared strings, look up the value in the
                                    // shared strings table.
                                    var stringTable =
                                        wbPart.GetPartsOfType<SharedStringTablePart>()
                                        .FirstOrDefault();

                                    // If the shared string table is missing, something 
                                    // is wrong. Return the index that is in
                                    // the cell. Otherwise, look up the correct text in 
                                    // the table.
                                    if (stringTable != null)
                                    {
                                        value =
                                            stringTable.SharedStringTable
                                            .ElementAt(int.Parse(value)).InnerText;
                                    }
                                    break;

                                case EXCEL.CellValues.Boolean:
                                    switch (value)
                                    {
                                        case "0":
                                            value = "FALSE";
                                            break;
                                        default:
                                            value = "TRUE";
                                            break;
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _ERROR_MENSAJE = ex.Message;
            }
            return value;
        }

        public static string GetDataReal(EXCEL.Cell cell, WorkbookPart wbPart, string ControlPorcentaje)
        {
            string cellValue = string.Empty;
            if (ControlPorcentaje == "SI")
            {
                cellValue = (Convert.ToDouble(cell.CellValue.Text) * 100).ToString();
            }
            else
            {
                cellValue = cell.CellValue.Text;
            }
            if (cell.DataType != null)
            {
                if (cell.DataType == EXCEL.CellValues.SharedString)
                {
                    int id = -1;
                    if (Int32.TryParse(cell.InnerText, out id))
                    {
                        EXCEL.SharedStringItem item = GetSharedStringItemById(wbPart, id);
                        if (item.Text != null)
                        {
                            cellValue = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            cellValue = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            cellValue = item.InnerXml;
                        }
                    }
                }
            }
            return cellValue;
        }

        public static EXCEL.SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<EXCEL.SharedStringItem>().ElementAt(id);
        }

        public static bool ValidateDefinedNames(string etiqueta, SpreadsheetDocument document)
        {
            bool result = false;
            var wbPart = document.WorkbookPart;
            EXCEL.DefinedNames definedNames = wbPart.Workbook.DefinedNames;
            if (definedNames != null)
            {
                int contador = 0;
                foreach (EXCEL.DefinedName dn in definedNames)
                {
                    contador++;
                    if (etiqueta.Equals(dn.Name.Value))
                    {
                        result = true;
                        break;
                    }
                }
            }
            return result;
        }

        public static Stream CreateOpenXml(string htmlEncodedString)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    string altChunkId = "myId";
                    MemoryStream ms02 = new MemoryStream(new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes(htmlEncodedString)).ToArray());
                    AlternativeFormatImportPart formatImportPart = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, altChunkId);
                    formatImportPart.FeedData(ms02);
                    AltChunk altChunk = new AltChunk();
                    altChunk.Id = altChunkId;
                    mainPart.Document.Body.Append(altChunk);
                }
                Stream WordFileUpdate = ms;
                return WordFileUpdate;
            }
        }

        public static string ReplaceHiperLink(Stream file, string textToReplace, string replace)
        {
            MemoryStream ms = null;
            try
            {
                using (ms = new MemoryStream())
                {
                    file.CopyTo(ms);
                    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(ms, file.CanRead))
                    {
                        MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
                        Body body = mainPart.Document.Body;
                        #region Body
                        List<SdtBlock> taggedContentControls = mainPart.Document.Descendants<SdtBlock>().ToList();
                        List<SdtRun> taggedRunContentControls = mainPart.Document.Descendants<SdtRun>().ToList();
                        List<Hyperlink> hLinks = mainPart.Document.Descendants<Hyperlink>().ToList();
                        List<Hyperlink> hLinksBody = mainPart.Document.Body.Descendants<Hyperlink>().ToList();
                        List<FieldCode> fieldsCodes = mainPart.Document.Body.Descendants<FieldCode>().ToList();
                        
                        if (taggedRunContentControls.Count > 0)
                        {
                            foreach (SdtRun control in taggedRunContentControls)
                            {
                                var hyperLinks = control.Descendants<Hyperlink>().ToList();
                                foreach (Hyperlink hyperLink in hyperLinks)
                                {
                                    if (hyperLink != null)
                                    {
                                        string relationId = hyperLink.Id;
                                        if (relationId != string.Empty)
                                        {
                                            var hr = mainPart.HyperlinkRelationships.FirstOrDefault(a => a.Id == relationId);
                                            if (hr == null) continue;
                                            var fieldName = hr.Uri.ToString();
                                            if (fieldName.Contains(textToReplace))
                                            {
                                                mainPart.DeleteReferenceRelationship(hr);
                                                mainPart.AddHyperlinkRelationship(new System.Uri(fieldName.Replace(textToReplace, replace), System.UriKind.Absolute), true, relationId);
                                            }
                                            
                                        }
                                    }

                                }
                            }

                        }
                        if (taggedContentControls.Count > 0)
                        {
                            foreach (SdtBlock control in taggedContentControls)
                            {
                                var hyperLinks = control.Descendants<Hyperlink>().ToList();
                                foreach (Hyperlink hyperLink in hyperLinks)
                                {
                                    if (hyperLink != null)
                                    {
                                        string relationId = hyperLink.Id;
                                        if (relationId != string.Empty)
                                        {
                                            var hr = mainPart.HyperlinkRelationships.FirstOrDefault(a => a.Id == relationId);
                                            if (hr == null) continue;
                                            var fieldName = hr.Uri.ToString();
                                            if (fieldName.Contains(textToReplace))
                                            {
                                                mainPart.DeleteReferenceRelationship(hr);
                                                mainPart.AddHyperlinkRelationship(new System.Uri(fieldName.Replace(textToReplace, replace), System.UriKind.Absolute), true, relationId);
                                            }

                                        }
                                    }
                                }
                            }

                        }
                        if (hLinks.Count > 0)
                        {
                            foreach (Hyperlink hyperLink in hLinks)
                            {
                                if (hyperLink != null)
                                {
                                    string relationId = hyperLink.Id;
                                    if (relationId != string.Empty)
                                    {
                                        var hr = mainPart.HyperlinkRelationships.FirstOrDefault(a => a.Id == relationId);
                                        if (hr == null) continue;
                                        var fieldName = hr.Uri.ToString();
                                        if (fieldName.Contains(textToReplace))
                                        {
                                            mainPart.DeleteReferenceRelationship(hr);
                                            mainPart.AddHyperlinkRelationship(new System.Uri(fieldName.Replace(textToReplace, replace), System.UriKind.Absolute), true, relationId);
                                        }

                                    }
                                }
                            }
                        }
                        if (fieldsCodes.Count > 0)
                        {
                            foreach (FieldCode fieldCode in fieldsCodes)
                            {
                                if (fieldCode != null)
                                {
                                    if (fieldCode.Text.Contains("HYPERLINK "))
                                    {
                                        string textReplaced = System.Uri.UnescapeDataString(fieldCode.Text);
                                        var url = new System.Uri(textReplaced.Replace("HYPERLINK ", "").Replace(textToReplace, replace), System.UriKind.Absolute);
                                        var extension = Path.GetExtension(url.AbsoluteUri);
                                        fieldCode.Text = "HYPERLINK " + url.AbsoluteUri;
                                    }
                                }
                            }
                        }
                        #endregion
                        mainPart.Document.Save();
                        wordprocessingDocument.Close();
                    }
                    Stream WordFileUpdate = ms;
                    string resultado = ConvertToBase64(ms);
                    //SaveStreamAsFile(@"C:\Users\INTEL\Desktop\", WordFileUpdate, "test.docx");
                    return resultado;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ms?.Close();
                ms?.Dispose();
            }
        }
        public static string ReplaceHiperLinkJustFieldCode(Stream file, string textToReplace, string replace, out bool validFieldCode)
        {
            MemoryStream ms = null;
            validFieldCode = false;
            try
            {
                using (ms = new MemoryStream())
                {
                    file.CopyTo(ms);
                    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(ms, file.CanRead))
                    {
                        MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
                        Body body = mainPart.Document.Body;
                        #region Body
                        List<SdtBlock> taggedContentControls = mainPart.Document.Descendants<SdtBlock>().ToList();
                        List<SdtRun> taggedRunContentControls = mainPart.Document.Descendants<SdtRun>().ToList();
                        List<Hyperlink> hLinks = mainPart.Document.Descendants<Hyperlink>().ToList();
                        List<Hyperlink> hLinksBody = mainPart.Document.Body.Descendants<Hyperlink>().ToList();
                        List<FieldCode> fieldsCodes = mainPart.Document.Body.Descendants<FieldCode>().ToList();

                        if (taggedRunContentControls.Count > 0)
                        {
                            foreach (SdtRun control in taggedRunContentControls)
                            {
                                var hyperLinks = control.Descendants<Hyperlink>().ToList();
                                foreach (Hyperlink hyperLink in hyperLinks)
                                {
                                    if (hyperLink != null)
                                    {
                                        string relationId = hyperLink.Id;
                                        if (relationId != string.Empty)
                                        {
                                            var hr = mainPart.HyperlinkRelationships.FirstOrDefault(a => a.Id == relationId);
                                            if (hr == null) continue;
                                            var fieldName = hr.Uri.ToString();
                                            if (fieldName.Contains(textToReplace))
                                            {
                                                mainPart.DeleteReferenceRelationship(hr);
                                                mainPart.AddHyperlinkRelationship(new System.Uri(fieldName.Replace(textToReplace, replace), System.UriKind.Absolute), true, relationId);
                                            }

                                        }
                                    }

                                }
                            }

                        }
                        if (taggedContentControls.Count > 0)
                        {
                            foreach (SdtBlock control in taggedContentControls)
                            {
                                var hyperLinks = control.Descendants<Hyperlink>().ToList();
                                foreach (Hyperlink hyperLink in hyperLinks)
                                {
                                    if (hyperLink != null)
                                    {
                                        string relationId = hyperLink.Id;
                                        if (relationId != string.Empty)
                                        {
                                            var hr = mainPart.HyperlinkRelationships.FirstOrDefault(a => a.Id == relationId);
                                            if (hr == null) continue;
                                            var fieldName = hr.Uri.ToString();
                                            if (fieldName.Contains(textToReplace))
                                            {
                                                mainPart.DeleteReferenceRelationship(hr);
                                                mainPart.AddHyperlinkRelationship(new System.Uri(fieldName.Replace(textToReplace, replace), System.UriKind.Absolute), true, relationId);
                                            }

                                        }
                                    }
                                }
                            }

                        }
                        if (hLinks.Count > 0)
                        {
                            foreach (Hyperlink hyperLink in hLinks)
                            {
                                if (hyperLink != null)
                                {
                                    string relationId = hyperLink.Id;
                                    if (relationId != string.Empty)
                                    {
                                        var hr = mainPart.HyperlinkRelationships.FirstOrDefault(a => a.Id == relationId);
                                        if (hr == null) continue;
                                        var fieldName = hr.Uri.ToString();
                                        if (fieldName.Contains(textToReplace))
                                        {
                                            mainPart.DeleteReferenceRelationship(hr);
                                            mainPart.AddHyperlinkRelationship(new System.Uri(fieldName.Replace(textToReplace, replace), System.UriKind.Absolute), true, relationId);
                                        }

                                    }
                                }
                            }
                        }
                        if (fieldsCodes.Count > 0)
                        {
                            foreach (FieldCode fieldCode in fieldsCodes)
                            {
                                if (fieldCode != null)
                                {
                                    if (fieldCode.Text.Contains("HYPERLINK "))
                                    {
                                        string textReplaced = System.Uri.UnescapeDataString(fieldCode.Text);
                                        var url = new System.Uri(textReplaced.Replace("HYPERLINK ", "").Replace("\"","").Replace(textToReplace, replace), System.UriKind.Absolute);
                                        var extension = Path.GetExtension(url.AbsoluteUri);
                                        fieldCode.Text = "HYPERLINK " + url.AbsoluteUri;
                                        validFieldCode = true;
                                    }
                                }
                            }
                        }
                        #endregion
                        mainPart.Document.Save();
                        wordprocessingDocument.Close();
                    }
                    Stream WordFileUpdate = ms;
                    string resultado = ConvertToBase64(ms);
                    //SaveStreamAsFile(@"C:\Users\INTEL\Desktop\", WordFileUpdate, "test.docx");
                    return resultado;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ms?.Close();
                ms?.Dispose();
            }
        }
        public static bool GetHiperLinkWithoutExtension(Stream file)
        {
            MemoryStream ms = null;
            try
            {
                using (ms = new MemoryStream())
                {
                    file.CopyTo(ms);
                    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(ms, file.CanRead))
                    {
                        MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
                        Body body = mainPart.Document.Body;
                        #region Body
                        List<SdtBlock> taggedContentControls = mainPart.Document.Descendants<SdtBlock>().ToList();
                        List<SdtRun> taggedRunContentControls = mainPart.Document.Descendants<SdtRun>().ToList();
                        List<Hyperlink> hLinks = mainPart.Document.Descendants<Hyperlink>().ToList();
                        List<Hyperlink> hLinksBody = mainPart.Document.Body.Descendants<Hyperlink>().ToList();
                        List<FieldCode> fieldsCodes = mainPart.Document.Body.Descendants<FieldCode>().ToList();

                        if (taggedRunContentControls.Count > 0)
                        {
                            foreach (SdtRun control in taggedRunContentControls)
                            {
                                var hyperLinks = control.Descendants<Hyperlink>().ToList();
                                foreach (Hyperlink hyperLink in hyperLinks)
                                {
                                    if (hyperLink != null)
                                    {
                                        string relationId = hyperLink.Id;
                                        if (relationId != string.Empty)
                                        {
                                            var hr = mainPart.HyperlinkRelationships.FirstOrDefault(a => a.Id == relationId);
                                            if (hr == null) continue;
                                            var fieldName = hr.Uri.ToString();
                                            var extension = Path.GetExtension(fieldName);
                                            if (string.IsNullOrEmpty(extension)) return true;

                                        }
                                    }

                                }
                            }

                        }
                        if (taggedContentControls.Count > 0)
                        {
                            foreach (SdtBlock control in taggedContentControls)
                            {
                                var hyperLinks = control.Descendants<Hyperlink>().ToList();
                                foreach (Hyperlink hyperLink in hyperLinks)
                                {
                                    if (hyperLink != null)
                                    {
                                        string relationId = hyperLink.Id;
                                        if (relationId != string.Empty)
                                        {
                                            var hr = mainPart.HyperlinkRelationships.FirstOrDefault(a => a.Id == relationId);
                                            if (hr == null) continue;
                                            var fieldName = hr.Uri.ToString();
                                            var extension = Path.GetExtension(fieldName);
                                            if (string.IsNullOrEmpty(extension)) return true;

                                        }
                                    }
                                }
                            }

                        }
                        if (hLinks.Count > 0)
                        {
                            foreach (Hyperlink hyperLink in hLinks)
                            {
                                if (hyperLink != null)
                                {
                                    string relationId = hyperLink.Id;
                                    if (relationId != string.Empty)
                                    {
                                        var hr = mainPart.HyperlinkRelationships.FirstOrDefault(a => a.Id == relationId);
                                        if (hr == null) continue;
                                        var fieldName = hr.Uri.ToString();
                                        var extension = Path.GetExtension(fieldName);
                                        if (string.IsNullOrEmpty(extension)) return true;

                                    }
                                }
                            }
                        }
                        if (fieldsCodes.Count > 0)
                        {
                            foreach (FieldCode fieldCode in fieldsCodes)
                            {
                                if (fieldCode != null)
                                {
                                    if(fieldCode.Text.Contains("HYPERLINK ")) return true;

                                }
                            }
                        }
                        #endregion
                        mainPart.Document.Save();
                        wordprocessingDocument.Close();
                    }
                    //Stream WordFileUpdate = ms;
                    //string resultado = ConvertToBase64(ms);
                    //SaveStreamAsFile(@"C:\Users\INTEL\Desktop\", WordFileUpdate, "test.docx");
                    return false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ms?.Close();
                ms?.Dispose();
            }
        }
        public static bool GetHiperLinkFieldCode(Stream file)
        {
            MemoryStream ms = null;
            try
            {
                using (ms = new MemoryStream())
                {
                    file.CopyTo(ms);
                    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(ms, file.CanRead))
                    {
                        MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
                        Body body = mainPart.Document.Body;
                        #region Body
                        List<FieldCode> fieldsCodes = mainPart.Document.Body.Descendants<FieldCode>().ToList();
                        if (fieldsCodes.Count > 0)
                        {
                            foreach (FieldCode fieldCode in fieldsCodes)
                            {
                                if (fieldCode != null)
                                {
                                    if(fieldCode.Text.Contains("HYPERLINK ")) return true;

                                }
                            }
                        }
                        #endregion
                        mainPart.Document.Save();
                        wordprocessingDocument.Close();
                    }
                    //Stream WordFileUpdate = ms;
                    //string resultado = ConvertToBase64(ms);
                    //SaveStreamAsFile(@"C:\Users\INTEL\Desktop\", WordFileUpdate, "test.docx");
                    return false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ms?.Close();
                ms?.Dispose();
            }
        }
        public static string CreateTableHeaderOpenXml(MemoryStream file, List<string> columns, List<Dictionary<string, object>> datos, Dictionary<string, object> totales)
        {
            MemoryStream ms = null;
            try
            {
                using (ms = new MemoryStream())
                {
                    file.CopyTo(ms);
                    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(ms, file.CanRead))
                    {
                        MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
                        Body body = mainPart.Document.Body;
                        #region Body
                        List<SdtBlock> taggedContentControls = mainPart.Document.Descendants<SdtBlock>().ToList();
                        List<SdtRun> taggedRunContentControls = mainPart.Document.Descendants<SdtRun>().ToList();
                        bool founded = false;
                        if (taggedRunContentControls.Count > 0 && !founded)
                        {
                            var ob = taggedRunContentControls.Where(x => x.Descendants<SdtAlias>().FirstOrDefault().Val.Value.Contains("Tabla"));
                            if (ob.Count() > 0)
                            {
                                var lst = ob.FirstOrDefault();
                                lst.SdtContentRun.Remove();
                                ob.FirstOrDefault().AppendChild(CreateTableFooter(columns, datos, totales));
                                founded = true;
                            }

                        }
                        if (taggedContentControls.Count > 0 && !founded)
                        {
                            var ob = taggedContentControls.Where(x => x.Descendants<SdtAlias>().FirstOrDefault().Val.Value.Contains("Tabla"));
                            if (ob.Count() > 0)
                            {

                                var lst = ob.FirstOrDefault();
                                lst.SdtContentBlock.Remove();
                                ob.FirstOrDefault().AppendChild(CreateTableFooter(columns, datos, totales));
                                founded = true;
                            }

                        }
                        #endregion
                        mainPart.Document.Save();
                        wordprocessingDocument.Close();
                    }
                    Stream WordFileUpdate = ms;
                    string resultado = ConvertToBase64(ms);
                    //OpenXml.SaveStreamAsFile(@"C:\Users\INTEL\Desktop\", WordFileUpdate, "test.docx");
                    return resultado;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ms?.Close();
                ms?.Dispose();
            }
        }

        private static Table CreateTableFooter(List<string> columns, List<Dictionary<string, object>> datos, Dictionary<string, object> totales)
        {
            Table table = new Table();
            TableProperties tableProperties = new TableProperties(
                new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = Convert.ToUInt32(columns.Count.ToString()), Color = "#000000" },
                new BottomBorder() { Val = BorderValues.Single, Size = Convert.ToUInt32(columns.Count.ToString()), Color = "#000000" },
                new LeftBorder() { Val = BorderValues.Single, Size = Convert.ToUInt32(columns.Count.ToString()), Color = "#000000" },
                new RightBorder() { Val = BorderValues.Single, Size = Convert.ToUInt32(columns.Count.ToString()), Color = "#000000" },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = Convert.ToUInt32(columns.Count.ToString()), Color = "#000000" },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = Convert.ToUInt32(columns.Count.ToString()), Color = "#000000" },
                new TableJustification() { Val = TableRowAlignmentValues.Center }),
                new TableWidth() { Width = "4500", Type = TableWidthUnitValues.Pct },
                new TableStyle() { Val = "TableGrid" },
                new TableLook { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true },
                new Justification { Val = JustificationValues.Center },
                new TableJustification { Val = TableRowAlignmentValues.Center },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
            );
            table.AppendChild<TableProperties>(tableProperties);
            TableGrid tg = new TableGrid(new GridColumn(), new GridColumn(), new GridColumn(), new GridColumn());
            table.AppendChild(tg);
            TableRow tr = new TableRow();
            foreach (string item in columns)
            {
                TableCell tc = new TableCell(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Left }, new SpacingBetweenLines() { Line = "280", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0" }), new Run(new RunProperties(new Bold(), new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = new StringValue("12") }), new Text(util.toString(item)))), new TableCellProperties(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }));
                tr.Append(tc);
            }
            table.AppendChild(tr);
            foreach (Dictionary<string, object> dato in datos)
            {
                TableRow tr1 = new TableRow();
                foreach (KeyValuePair<string, object> item in dato)
                {
                    TableCell tc1 = new TableCell(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Left }, new SpacingBetweenLines() { Line = "280", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0" }), new Run(new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = new StringValue("12") }), new Text(util.toString(item.Value)))), new TableCellProperties(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }));
                    tr1.Append(tc1);
                }
                table.AppendChild(tr1);
            }
            foreach (KeyValuePair<string, object> item in totales)
            {
                TableRow tr2 = new TableRow();
                if (columns.Count > 2)
                {
                    int count = columns.Count - 2;
                    TableCell sub1 = new TableCell(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Left }, new SpacingBetweenLines() { Line = "280", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0" }), new Run(new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = new StringValue("12") }), new Text(""))), new TableCellProperties(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                                    new GridSpan() { Val = count }, new VerticalMerge() { Val = MergedCellValues.Continue }));
                    tr2.Append(sub1);
                }
                TableCell sub3 = new TableCell(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Left }, new SpacingBetweenLines() { Line = "280", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0" }), new Run(new RunProperties(new Bold(), new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = new StringValue("12") }), new Text(util.toString(item.Key)))), new TableCellProperties(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }));
                TableCell sub4 = new TableCell(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Left }, new SpacingBetweenLines() { Line = "280", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0" }), new Run(new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = new StringValue("12") }), new Text(util.toString(item.Value)))), new TableCellProperties(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }));

                tr2.Append(sub3);
                tr2.Append(sub4);
                table.AppendChild(tr2);
            }

            return table;
        }


        public static byte[] ConvertBaseToByte(string base64)
        {
            return Convert.FromBase64String(base64);
        }

        public static MemoryStream ConvertBaseToMemoryStream(string base64)
        {
            var file = Convert.FromBase64String(base64);
            MemoryStream stream = new MemoryStream(file);
            return stream;
        }

        public static string ConvertToBase64(Stream stream)
        {
            byte[] bytes;
            using (var memoryStream = new MemoryStream())
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(memoryStream);
                bytes = memoryStream.ToArray();
            }

            string base64 = Convert.ToBase64String(bytes);
            return base64;
        }


        public static void SaveStreamAsFile(string filePath, Stream inputStream, string fileName)
        {
            DirectoryInfo info = new DirectoryInfo(filePath);
            if (!info.Exists)
            {
                info.Create();
            }
            string path = Path.Combine(filePath, fileName);
            //using (FileStream outputFileStream = new FileStream(path, FileMode.Create))
            //{
            //    inputStream.CopyTo(outputFileStream);
            //}

            using (var fileStream = File.Create(path))
            {
                inputStream.Seek(0, SeekOrigin.Begin);
                inputStream.CopyTo(fileStream);
            }
        }
    }
}
