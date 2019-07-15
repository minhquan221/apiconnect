using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml.Serialization;

namespace NewPMS.Helpers
{
    public class DataSourceTable {
        public bool IsAddHeader { get; set; }
        public DataTable Source { get; set; }
    }
    public class ExcelHelper {
        Dictionary<string, object> dic = new Dictionary<string, object>();
        Dictionary<string, DataSourceTable> dic2 = new Dictionary<string, DataSourceTable>();
        private List<string> errors = new List<string>();
        private Regex reg = new Regex(@"^([A-Z])(\d+)$"); //only accept column from A to Z
        public ExcelHelper AddCell(string offset, object value) {
            if (!string.IsNullOrWhiteSpace(offset)) {
                dic[offset.Replace(" ", "").ToUpper()] = value;
            }
            return this;
        }
        public ExcelHelper AddDataSource(string offset, DataSourceTable dt)
        {
            if (!string.IsNullOrWhiteSpace(offset))
            {
                dic2[offset.Replace(" ", "").ToUpper()] = dt;
            }
            return this;
        }
        private bool Translate(string offset, ref int col, ref int row) {
            Match m = reg.Match(offset);
            
            if (m.Success) {
                col = m.Groups[1].ToString()[0] - 'A';
                row = int.Parse(m.Groups[2].ToString()) - 1;
                return true;
            }
            return false;   
        }
        public byte[] Fill(string template)
        {
            using (var workBook = new Workbook())
            {
                if (workBook.LoadDocument(template))
                {
                    var workSheet = workBook.Worksheets[0];
                    int col = 0;
                    int row = 0;
                    foreach (var pair in dic2) {
                        if (Translate(pair.Key, ref col, ref row)) {
                            workSheet.Import(pair.Value.Source, pair.Value.IsAddHeader, row, col);
                        }
                    }
                    foreach (var pair in dic) {
                        if (Translate(pair.Key, ref col, ref row))
                        {
                            workSheet.Cells[row, col].SetValue(pair.Value);
                        }
                    }
                    return workBook.SaveDocument(DocumentFormat.Xlsx);
                }
            }
            return null;
        }
        public static byte[] ExportExcel(string templateFile, string offset, DataTable dt, bool addHeader = false) {
            ExcelHelper ex = new ExcelHelper();
            return ex.AddDataSource(offset, new DataSourceTable { IsAddHeader = addHeader, Source = dt }).Fill(templateFile);
        }
        public static bool ImportExcel<T>(IEnumerable<HttpPostedFileBase> files, IteratorOption option, Func<Worksheet, int, IteratorOption, T> func,Func<Worksheet, bool> funcSetup, out List<T> list) {
            HttpPostedFileBase file = null;
            foreach (HttpPostedFileBase f in files)
            {
                file = f;
                break;
            }
            return ImportExcel_Implement(file, option, func,funcSetup, out list);
        }
        public static bool ImportExcel<T>(HttpFileCollectionBase files, IteratorOption option, Func<Worksheet, int, IteratorOption, T> func,Func<Worksheet, bool> funcSetup, out List<T> list)
        {
            return ImportExcel_Implement(files[0], option, func, funcSetup, out list);
        }
        public static bool ImportExcel_Implement<T>(HttpPostedFileBase file, IteratorOption option, Func<Worksheet, int, IteratorOption, T> func, Func<Worksheet, bool> funcSetup , out List<T> list)
        {
            list = new List<T>();
            if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
            {
                try
                {
                    using (var workBook = new Workbook())
                    {
                        
                        workBook.LoadDocument(file.InputStream);
                        Worksheet sheet = workBook.Worksheets[0];
                        var range = sheet.GetDataRange();
                        int row = range.BottomRowIndex;
                        string cp = string.Empty;

                        funcSetup(sheet);

                        for (int i = 1; i <= row; i++)
                        {
                            var ret = func(sheet, i, option);
                            if (option.StopWhenError && (option.IsError || ret == null))
                            {
                                return false;
                            }
                            else
                            {
                                list.Add(ret);
                            }
                        }
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    option.AddError(ex.Message);
                    return false;
                }
            }
            else {
                option.AddError("Chưa chọn file hoặc file bị lỗi");
                return false;
            }
        }
    }
    public class IteratorOption {
        private readonly bool _StopWhenError = false;
            public IteratorOption(bool stopWhenError) {
                _StopWhenError = stopWhenError;
            }
        public bool StopWhenError {
            get {
                return _StopWhenError;
            }
        }
        public  bool IsError { get; set; }
        private List<string> errorList = new List<string>();
        public void AddError(string msg) {
            IsError = true;
            errorList.Add(msg);
        }
        public List<string> GetErrors()
        {
            return errorList;
        }
        public string GetErrors(string sep)
        {
            return string.Join(sep, errorList);
        }
    }
    public static class StaticUtils {
        public static int GetInt(this object ob, int _default = 0) {
            int val = _default;
            if (ob != null) {
                if (!int.TryParse(ob.ToString(), out val)) {
                    val = _default;
                }
            }
            return val;
        }
    }
}