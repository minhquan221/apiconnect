using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.IO;
using System.Text.RegularExpressions;
using DevExpress.Spreadsheet;

namespace NewPMS.Models
{
    public class sys
    {
        public static Info current
        {
            get
            {
                return HttpContext.Current.Session["INFO"] as Info;
            }
        }
    }
    public class result
    {
        public string Code { get; set; }
        public string Message { get; set; }
    }
    public class XmlUtils
    {
        public static string Serialize(object ob)
        {
            XmlSerializer x = new XmlSerializer(ob.GetType());
            using (StringWriter sw = new StringWriter())
            {
                x.Serialize(sw, ob);
                return sw.ToString();
            }
        }
    }
    public static class Helper
    {
        #region ImportCompliance

        public static List<Kpi> KpisForIpCompliance(string userId, decimal periodId)
        {
            string errMsg = string.Empty;
            List<Kpi> lst = new List<Kpi>();
            DataTable tb = DbHelper.GetKpisForIpCompliance(userId, periodId, out errMsg);
            if (tb == null || tb.Rows.Count == 0)
            {
                return lst;
            }
            else
            {
                foreach (DataRow row in tb.Rows)
                {
                    Kpi item = new Kpi()
                    {
                        Id = row["ID"] != DBNull.Value ? row["ID"].ToString() : "",
                        Name = row["NAME"] != DBNull.Value ? row["NAME"].ToString() : string.Empty,
                    };
                    lst.Add(item);
                }
            }
            return lst;
        }

        public static List<object> OuListForIpCompliance(string userId, decimal periodId, string kpiId)
        {
            string errMsg = string.Empty;
            List<object> lst = new List<object>();
            DataTable tb = DbHelper.GetOuListForIpCompliance(userId, periodId, kpiId, out errMsg);
            if (tb == null || tb.Rows.Count == 0)
            {
                return lst;
            }
            else
            {
                foreach (DataRow row in tb.Rows)
                {
                    object item = new
                    {
                        Id = row["CODE"] != DBNull.Value ? row["CODE"].ToString() : "",
                        Name = row["NAME"] != DBNull.Value ? row["NAME"].ToString() : string.Empty,
                    };
                    lst.Add(item);
                }
            }
            return lst;
        }

        public static object ReadImportCompliance(string userId, int periodId, Stream fileStream, ref string error)
        {
            string errMsg = string.Empty;
            try
            {
                List<object> list = new List<object>();
                var cps = DbHelper.GetKpisForIpCompliance(userId, periodId, out errMsg).AsEnumerable().Select(q => new { Id = q["ID"].ToString(), Name = q["NAME"] }).ToList();
                using (var workBook = new Workbook())
                {
                    workBook.LoadDocument(fileStream);
                    Worksheet sheet = workBook.Worksheets[0];
                    var range = sheet.GetDataRange();
                    int row = range.BottomRowIndex;
                    string cp = string.Empty;
                    for (int i = 1; i <= row; i++)
                    {
                        var userid = sheet.GetCellValue(0, i).TextValue;
                        var username = sheet.GetCellValue(1, i).TextValue;
                        var kpiid = sheet.GetCellValue(2, i).TextValue;
                        var result = sheet.GetCellValue(3, i).NumericValue;
                        var note = sheet.GetCellValue(4, i).TextValue;
                        //var formular = sheet.GetCellValue(5, i).TextValue;
                        //var format = sheet.GetCellValue(6, i).TextValue;
                        if (string.IsNullOrEmpty(userid))
                        {
                            error += string.Format("Dòng thứ {0} trên file, mã nhân sự phải là bắt buộc <br/>", i);
                            break;
                        }
                        if (string.IsNullOrEmpty(kpiid))
                        {
                            error += string.Format("Dòng thứ {0} trên file, mã KPI phải là bắt buộc <br/>", i);
                            break;
                        }
                        else
                        {
                            if (cps.Find(x => x.Id == kpiid.Trim()) == null)
                            {
                                error += string.Format("Dòng thứ {0} trên file, bạn không được phân quyền KPI này <br/>", i);
                                break;
                            }
                            else
                            {
                                var Kpidetail = GetKpi(userId, kpiid.Trim());
                                if (Kpidetail == null)
                                {
                                    error += string.Format("Dòng thứ {0} trên file, không tìm thấy KPI {1} P_ID {2} <br/>", i, kpiid.Trim(), userid);
                                    continue;
                                }
                            }
                        }
                        string err = string.Empty;
                        var bolExistsInPeriod = DbHelper.CheckExistsInPeriod(userId, userid, kpiid.Trim(), periodId, ref err);
                        if (!bolExistsInPeriod)
                        {
                            error += string.Format("Dòng thứ {0} trên file, không tìm thấy trong kỳ đánh giá đang Active nhân sự: Mã nhân sự {1} và Mã KPI {2}", i, userid, kpiid.Trim());
                            continue;
                        }

                        list.Add(
                            new
                            {
                                Id = userid,
                                Name = username,
                                KpiId = kpiid,
                                Result = result,
                                Note = note,
                            }
                        );
                    }
                }
                return list;
            }
            catch (Exception ex)
            {
                error = ex.ToString();
                return null;
            }
        }

        public static byte[] ExportInternalIpCompliance(string userId, int periodId, string kpiId, string ouId, int check = 0)
        {
            string errMsg = string.Empty;
            DataTable list = DbHelper.GetDataForIpCompliance(userId, periodId, kpiId, ouId, out errMsg, check);

            MemoryStream stream = new MemoryStream();
            using (var workBook = new Workbook())
            {
                workBook.LoadDocument(HttpContext.Current.Server.MapPath("~/Template/IMPORT/IMPORT_COMPLIANCE.xlsx"));
                DevExpress.Spreadsheet.Worksheet sheet = workBook.Worksheets[0];
                DevExpress.Spreadsheet.Style lStyle = workBook.Styles.Add("leftStyle");
                lStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Alignment.WrapText = true;
                DevExpress.Spreadsheet.Style cStyle = workBook.Styles.Add("centerStyle");
                cStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                cStyle.Alignment.WrapText = true;
                DevExpress.Spreadsheet.Style rStyle = workBook.Styles.Add("rightStyle");
                rStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right;
                rStyle.Alignment.WrapText = true;
                int index = 1;
                foreach (DataRow item in list.Rows)
                {
                    Cell c1 = sheet.Cells[index, 0];
                    c1.SetValue(item["EMPLOYEE_NO"]);
                    c1.Style = lStyle;

                    Cell c2 = sheet.Cells[index, 1];
                    c2.SetValue(item["NAME"] != DBNull.Value ? item["NAME"].ToString() : "");
                    c2.Style = cStyle;

                    Cell c3 = sheet.Cells[index, 2];
                    c3.SetValue(item["KPI_ID"].ToString());
                    c3.Style = rStyle;
                    index++;
                }
                workBook.SaveDocument(stream, DocumentFormat.Xlsx);
                return stream.ToArray();
            }

        }

        public static bool SaveImportIpCompliance(string userId, decimal periodId, List<PostedImportIpCompliance> list, ref string error)
        {
            return DbHelper.SaveImportIpCompliance(userId, periodId, list, ref error);
        }

        #endregion

        public static string NonUnicode(this string text)
        {
            string[] arr1 = new string[] { "á", "à", "ả", "ã", "ạ", "â", "ấ", "ầ", "ẩ", "ẫ", "ậ", "ă", "ắ", "ằ", "ẳ", "ẵ", "ặ",
        "đ",
        "é","è","ẻ","ẽ","ẹ","ê","ế","ề","ể","ễ","ệ",
        "í","ì","ỉ","ĩ","ị",
        "ó","ò","ỏ","õ","ọ","ô","ố","ồ","ổ","ỗ","ộ","ơ","ớ","ờ","ở","ỡ","ợ",
        "ú","ù","ủ","ũ","ụ","ư","ứ","ừ","ử","ữ","ự",
        "ý","ỳ","ỷ","ỹ","ỵ",};
            string[] arr2 = new string[] { "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a",
        "d",
        "e","e","e","e","e","e","e","e","e","e","e",
        "i","i","i","i","i",
        "o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o",
        "u","u","u","u","u","u","u","u","u","u","u",
        "y","y","y","y","y",};
            for (int i = 0; i < arr1.Length; i++)
            {
                text = text.Replace(arr1[i], arr2[i]);
                text = text.Replace(arr1[i].ToUpper(), arr2[i].ToUpper());
            }
            return text;
        }


        public static string GetSafeFilename(string filename)
        {
            string str = NonUnicode(filename);
            return Regex.Replace(Regex.Replace(str, @"[^a-zA-Z0-9 ]+", "", RegexOptions.Compiled), @"\s+", "_");
        }

        //menu
        private static string GetAction(string str)
        {
            string action = string.Empty;
            string[] s = str.Split(';');
            foreach (var item in s)
            {
                if (item.ToLower().Contains("action="))
                    return item.Replace("ACTION=", string.Empty);
            }
            return string.Empty;
        }

        private static string GetController(string str)
        {
            string action = string.Empty;
            string[] s = str.Split(';');
            foreach (var item in s)
            {
                if (item.ToLower().Contains("controller="))
                    return item.Replace("CONTROLLER=", string.Empty);
            }
            return string.Empty;
        }

        public static string GetMenu(List<Module> modules)
        {
            StringBuilder builder = new StringBuilder();
            string action = string.Empty;
            string controller = string.Empty;
            string url = "#";
            foreach (var item in modules.Where(q => string.IsNullOrEmpty(q.ParentCode)).OrderBy(q => q.DisplayOrder))
            {
                var subMenus = modules.Where(q => q.ParentCode == item.Code).OrderBy(q => q.DisplayOrder);
                if (subMenus.Count() > 0)
                    builder.Append("<li class=\"has-submenu\">");
                else
                    builder.Append("<li>");

                action = GetAction(item.Action);
                controller = GetController(item.Action);
                if (!string.IsNullOrEmpty(item.Action) && !string.IsNullOrEmpty(action) && !string.IsNullOrEmpty(controller))
                    url = new UrlHelper(HttpContext.Current.Request.RequestContext).Action(action, controller);
                builder.Append(string.Format("<a href=\"{0}\"><i class=\"{1}\" {3}></i>{2}</a>", url, item.Icon, item.Name, item.IsOpenNewTab ? "target=\"_blank\"" : ""));
                if (subMenus.Count() > 0)
                {
                    builder.Append("<ul class=\"submenu\">");
                    foreach (var item1 in subMenus)
                    {
                        builder.Append(CreateSubMenu(item1, modules));
                    }
                    builder.Append("</ul>");
                }
                builder.Append("</li>");
            }
            return builder.ToString();
        }

        private static string CreateSubMenu(Module module, List<Module> list)
        {
            string action = string.Empty;
            string controller = string.Empty;
            string url = "#";
            StringBuilder builder = new StringBuilder();
            var subMenus = list.Where(q => q.ParentCode == module.Code).OrderBy(q => q.DisplayOrder);
            if (subMenus.Count() > 0)
                builder.Append("<li class=\"has-submenu\">");
            else
                builder.Append("<li>");
            action = GetAction(module.Action);
            controller = GetController(module.Action);
            if (!string.IsNullOrEmpty(module.Action) && !string.IsNullOrEmpty(action) && !string.IsNullOrEmpty(controller))
                url = new UrlHelper(HttpContext.Current.Request.RequestContext).Action(action, controller);
            builder.Append(string.Format("<a href=\"{0}\" {3}><i class=\"{1}\"></i>{2}</a>", url, module.Icon, module.Name, module.IsOpenNewTab ? "target=\"_blank\"" : ""));
            if (subMenus.Count() > 0)
            {
                builder.Append("<ul class=\"submenu\">");
                foreach (var item in subMenus)
                {
                    builder.Append(CreateSubMenu(item, list));
                }
                builder.Append("</ul>");
            }
            builder.Append("</li>");
            return builder.ToString();
        }
        //end menu
        public static UserInfo GetUserInfo(string userName)
        {
            DataTable tb = DbHelper.GetUserInfo(userName);
            if (tb == null || tb.Rows.Count == 0)
                return null;
            DataRow row = tb.Rows[0];
            return new UserInfo()
            {
                EmployeeNo = row["EMPLOYEE_NO"].ToString(),
                Name = row["NAME"].ToString(),
                UserDomain = row["USER_DOMAIN"].ToString(),
                UserDomainCheckPass = row["USER_DOMAIN_CHECK_PASS"].ToString(),
                UserId = row["USER_DOMAIN"].ToString(),
                Locked = row["LOCKED"] != DBNull.Value ? int.Parse(row["LOCKED"].ToString()) : 0,
                OuId = row["OU_ID"].ToString(),
            };
        }

        public static DataTable GetAppConfig()
        {
            //dữ liệu test
            DataTable tb = new DataTable();
            tb.Columns.Add("KEY", typeof(string));
            tb.Columns.Add("VALUE", typeof(string));

            DataRow row = tb.NewRow();
            DataRow row2 = tb.NewRow();
            DataRow row3 = tb.NewRow();

            row["KEY"] = "APP_LOCKED";
            row["VALUE"] = "0";

            row["KEY"] = "APP_EXCEPT_USER";
            row["VALUE"] = "";

            row["KEY"] = "APP_LOCKED_USER";
            row["VALUE"] = "";

            tb.Rows.Add(row);
            tb.Rows.Add(row2);
            tb.Rows.Add(row3);

            return tb;
        }

        public static DataTable GetPeriodList(string userId) => DbHelper.GetPeriodList(userId);

        public static DataTable GetCompanyStructByType(string userId, int? type) => DbHelper.GetCompanyStructByType(userId, type);

        public static DataTable GetCompanyStructByClasscification(string userId, string classId) => DbHelper.GetCompanyStructByClasscification(userId, classId);

        public static DataTable GetEvaluationByOuId(string userId, string ouId, int periodId) => DbHelper.GetEvaluationByOuId(userId, ouId, periodId);

        public static DataTable GetEvaluationOuList(string userId, int periodId)
        {
            return DbHelper.GetEvaluationOuList(userId, periodId);
        }

        public static DataTable GetEmployeesByOuId(string userId, string ouId, int isAllChildren)
        {
            return DbHelper.GetEmployeesByOuId(userId, ouId, isAllChildren);
        }

        public static DataTable GetEmployeesByUser(string userId, string employeeNo, int isAllChildren)
        {
            return DbHelper.GetEmployeesByUser(userId, employeeNo, isAllChildren);
        }
        public static DataTable GetEvaluatedOuList(string userId, string ouId, int periodId)
        {
            return DbHelper.GetEvaluatedOuList(userId, ouId, periodId);
        }

        public static DataSet GetEvaluatonInfo(string userId, string ouId, string evaluationOuId, int periodId)
        {
            return DbHelper.GetEvaluatonInfo(userId, ouId, evaluationOuId, periodId);
        }

        public static DataTable GetAllEmployeeForSync(string userId) => DbHelper.GetAllEmployeeForSync(userId);

        public static DataTable GetSLAOu(string userId, string ouId, int periodId) => DbHelper.GetSLAOu(userId, ouId, periodId);

        public static DataTable GetEmployeeType(string userId) => DbHelper.GetEmployeeType(userId);

        public static DataTable GetContractType(string userId) => DbHelper.GetContractType(userId);

        public static DataTable GetEmployeePeriodList(string userId, int periodId) => DbHelper.GetEmployeePeriodList(userId, periodId);

        public static DataTable GetDataForSetupEmployeePeriod(string userId, string comparison, string startDate, int? status, string employeeType, string contractType)
        {
            DateTime start = DateTime.MinValue;
            DateTime.TryParseExact(startDate, new string[] { "dd/MM/yyyy", "d/M/yyyy" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out start);
            return DbHelper.GetDataForSetupEmployeePeriod(userId, comparison, start == DateTime.MinValue ? null : (DateTime?)start, status, employeeType, contractType);
        }

        public static List<Module> GetModuleByUser(string userId)
        {
            ////Dữ liệu test
            //List<Module> modules = new List<Module>();

            //modules.Add(new Module() { Code = "1", Name = "Apps", Icon = "icon-layers" });
            //modules.Add(new Module() { Code = "2", Name = "UI Elements", Icon = "icon-layers" });
            //modules.Add(new Module() { Code = "3", Name = "Components", Icon = "icon-layers" });
            //modules.Add(new Module() { Code = "4", Name = "Pages", Icon = "icon-layers" });

            //modules.Add(new Module() { Code = "5", Name = "Calendar", ParentCode = "1" });
            //modules.Add(new Module() { Code = "6", Name = "Ticket", ParentCode = "1" });
            //modules.Add(new Module() { Code = "7", Name = "TaskBoard", ParentCode = "1" });
            //modules.Add(new Module() { Code = "8", Name = "Task detail", ParentCode = "1" });
            //modules.Add(new Module() { Code = "9", Name = "Contract", ParentCode = "1" });

            //modules.Add(new Module() { Code = "10", Name = "Email", ParentCode = "2" });
            //modules.Add(new Module() { Code = "11", Name = "Wigets", ParentCode = "2" });
            //modules.Add(new Module() { Code = "12", Name = "Chart", ParentCode = "2" });
            //modules.Add(new Module() { Code = "13", Name = "Form", ParentCode = "2" });
            //modules.Add(new Module() { Code = "14", Name = "Table", ParentCode = "2" });
            //modules.Add(new Module() { Code = "15", Name = "Map", ParentCode = "2" });

            //modules.Add(new Module() { Code = "16", Name = "Inbox", ParentCode = "10" });
            //modules.Add(new Module() { Code = "17", Name = "Read Mail", ParentCode = "10" });
            //modules.Add(new Module() { Code = "18", Name = "Compose Mail", ParentCode = "10" });

            //modules.Add(new Module() { Code = "19", Name = "Flot chart", ParentCode = "12" });
            //modules.Add(new Module() { Code = "20", Name = "Google chart", ParentCode = "12" });
            //modules.Add(new Module() { Code = "21", Name = "JQuery knob", ParentCode = "12" });

            //return modules;
            List<Module> modules = new List<Module>();
            foreach (DataRow item in DbHelper.GetModuleByUser(userId).Rows)
            {
                Module m = new Module()
                {
                    Code = item["CODE"] != DBNull.Value ? item["CODE"].ToString() : string.Empty,
                    Action = item["ACTION"] != DBNull.Value ? item["ACTION"].ToString() : string.Empty,
                    Icon = item["ICON"] != DBNull.Value ? item["ICON"].ToString() : string.Empty,
                    IsOpenNewTab = (item["OPEN_NEW_TAB"] != DBNull.Value ? int.Parse(item["OPEN_NEW_TAB"].ToString()) : 0) == 1,
                    Name = item["NAME"] != DBNull.Value ? item["NAME"].ToString() : string.Empty,
                    ParentCode = item["PARENT_CODE"] != DBNull.Value ? item["PARENT_CODE"].ToString() : string.Empty,
                    DisplayOrder = item["DISPLAY_ORDER"] != DBNull.Value ? int.Parse(item["DISPLAY_ORDER"].ToString()) : 1,
                };
                modules.Add(m);
            }
            return modules.Distinct().ToList();
        }

        public static List<Role> GetRoleByUser(string userId)
        {
            List<Role> roles = new List<Role>();
            foreach (DataRow item in DbHelper.GetRoleByUser(userId).Rows)
            {
                Role m = new Role()
                {
                    Id = item["ID"] != DBNull.Value ? int.Parse(item["ID"].ToString()) : -1,
                    Name = item["NAME"] != DBNull.Value ? item["NAME"].ToString() : string.Empty,
                };
                roles.Add(m);
            }
            return roles;
        }

        public static Period GetActivePeriod(string userId)
        {
            DataTable tb = DbHelper.GetActivePeriod(userId);
            if (tb == null || tb.Rows.Count == 0)
                return null;
            else
            {
                DataRow row = tb.Rows[0];
                return new Period()
                {
                    Id = row["ID"] != DBNull.Value ? int.Parse(row["ID"].ToString()) : -1,
                    Name = row["NAME"] != DBNull.Value ? row["NAME"].ToString() : string.Empty,
                    Year = row["YEAR"] != DBNull.Value ? int.Parse(row["YEAR"].ToString()) : -1
                };
            }
        }

        public static Period GetPeriodInfo(string userId, int periodId)
        {
            DataTable tb = DbHelper.GetPeriodInfo(userId, periodId);
            if (tb == null || tb.Rows.Count == 0)
                return null;
            else
            {
                DataRow row = tb.Rows[0];
                return new Period()
                {
                    Id = row["ID"] != DBNull.Value ? int.Parse(row["ID"].ToString()) : -1,
                    Name = row["NAME"] != DBNull.Value ? row["NAME"].ToString() : string.Empty,
                    Year = row["YEAR"] != DBNull.Value ? int.Parse(row["YEAR"].ToString()) : -1
                };
            }
        }

        public static Employee GetEvaluationUser(string userId, string ouId, int periodId)
        {
            DataTable tb = DbHelper.GetEvaluationUser(userId, ouId, periodId);
            if (tb == null || tb.Rows.Count == 0)
                return null;
            else
            {
                DataRow row = tb.Rows[0];
                return new Employee()
                {
                    EmployeeNo = row["EMPLOYEE_NO"] != DBNull.Value ? row["EMPLOYEE_NO"].ToString() : string.Empty,
                    Name = row["NAME"] != DBNull.Value ? row["NAME"].ToString() : string.Empty,
                    Dob = row["DOB"] != DBNull.Value ? (DateTime?)DateTime.Parse(row["DOB"].ToString()) : null,
                    UserDomain = row["USER_DOMAIN"] != DBNull.Value ? row["USER_DOMAIN"].ToString() : string.Empty,
                    UserDomainCheckPass = row["USER_DOMAIN_CHECK_PASS"] != DBNull.Value ? row["USER_DOMAIN_CHECK_PASS"].ToString() : string.Empty,
                    OuId = row["OU_ID"] != DBNull.Value ? row["OU_ID"].ToString() : string.Empty,
                    StartDate = row["START_DATE"] != DBNull.Value ? (DateTime?)DateTime.Parse(row["START_DATE"].ToString()) : null,
                    StarCompanyDate = row["START_COMPANY_DATE"] != DBNull.Value ? (DateTime?)DateTime.Parse(row["START_COMPANY_DATE"].ToString()) : null,
                    LeaveOffDate = row["LEAVE_OFF_WORK_DATE"] != DBNull.Value ? (DateTime?)DateTime.Parse(row["LEAVE_OFF_WORK_DATE"].ToString()) : null,
                    ContractType = row["CONTRACT_TYPE"] != DBNull.Value ? row["CONTRACT_TYPE"].ToString() : string.Empty,
                    EmployeeType = row["EMPLOYEE_TYPE"] != DBNull.Value ? row["EMPLOYEE_TYPE"].ToString() : string.Empty,
                    SyncUser = row["SYNC_USER"] != DBNull.Value ? row["SYNC_USER"].ToString() : string.Empty,
                    JobTitleId = row["JOB_TITLE_ID"] != DBNull.Value ? row["JOB_TITLE_ID"].ToString() : string.Empty,
                    OuName = row["OU_NAME"] != DBNull.Value ? row["OU_NAME"].ToString() : string.Empty,
                };
            }
        }

        public static Employee GetEmployeePeriod(string employeeNo, int periodId)
        {
            DataTable tb = DbHelper.GetEmployeePeriod(employeeNo, periodId);
            if (tb == null || tb.Rows.Count == 0)
                return null;
            else
            {
                DataRow row = tb.Rows[0];
                return new Employee()
                {
                    EmployeeNo = row["EMPLOYEE_NO"] != DBNull.Value ? row["EMPLOYEE_NO"].ToString() : string.Empty,
                    Name = row["NAME"] != DBNull.Value ? row["NAME"].ToString() : string.Empty,
                    Dob = row["DOB"] != DBNull.Value ? (DateTime?)DateTime.Parse(row["DOB"].ToString()) : null,
                    UserDomain = row["USER_DOMAIN"] != DBNull.Value ? row["USER_DOMAIN"].ToString() : string.Empty,
                    UserDomainCheckPass = row["USER_DOMAIN_CHECK_PASS"] != DBNull.Value ? row["USER_DOMAIN_CHECK_PASS"].ToString() : string.Empty,
                    OuId = row["OU_ID"] != DBNull.Value ? row["OU_ID"].ToString() : string.Empty,
                    StartDate = row["START_DATE"] != DBNull.Value ? (DateTime?)DateTime.Parse(row["START_DATE"].ToString()) : null,
                    StarCompanyDate = row["START_COMPANY_DATE"] != DBNull.Value ? (DateTime?)DateTime.Parse(row["START_COMPANY_DATE"].ToString()) : null,
                    LeaveOffDate = row["LEAVE_OFF_WORK_DATE"] != DBNull.Value ? (DateTime?)DateTime.Parse(row["LEAVE_OFF_WORK_DATE"].ToString()) : null,
                    ContractType = row["CONTRACT_TYPE"] != DBNull.Value ? row["CONTRACT_TYPE"].ToString() : string.Empty,
                    EmployeeType = row["EMPLOYEE_TYPE"] != DBNull.Value ? row["EMPLOYEE_TYPE"].ToString() : string.Empty,
                    SyncUser = row["SYNC_USER"] != DBNull.Value ? row["SYNC_USER"].ToString() : string.Empty,
                    JobTitleId = row["JOB_TITLE_ID"] != DBNull.Value ? row["JOB_TITLE_ID"].ToString() : string.Empty,
                    JobTitle = row["JOB_TITLE"] != DBNull.Value ? row["JOB_TITLE"].ToString() : string.Empty,
                    OuName = row["OU_NAME"] != DBNull.Value ? row["OU_NAME"].ToString() : string.Empty,
                };
            }
        }

        public static DataSet GetKotEmpForCreateEvaluation(string employeeNo, int periodId) => DbHelper.GetKotEmpForCreateEvaluation(employeeNo, periodId);

        public static CompanyStructure GetCompanyInfo(string userId, string ouId)
        {
            DataTable tb = DbHelper.GetCompanyInfo(userId, ouId);
            if (tb == null || tb.Rows.Count == 0) return null;
            return new CompanyStructure()
            {
                Id = tb.Rows[0]["ID"] != DBNull.Value ? int.Parse(tb.Rows[0]["ID"].ToString()) : -1,
                Code = tb.Rows[0]["CODE"].ToString(),
                Name = tb.Rows[0]["NAME"].ToString(),
                ManagedEmployee = tb.Rows[0]["MANAGED_EMPLOYEE"] != DBNull.Value ? tb.Rows[0]["MANAGED_EMPLOYEE"].ToString() : string.Empty,
                MixedCode = tb.Rows[0]["MIXED_CODE"] != DBNull.Value ? tb.Rows[0]["MIXED_CODE"].ToString() : string.Empty,
                Level = tb.Rows[0]["CLEVEL"] != DBNull.Value ? (int?)int.Parse(tb.Rows[0]["CLEVEL"].ToString()) : null,
                ParentId = tb.Rows[0]["PARENT_ID"] != DBNull.Value ? int.Parse(tb.Rows[0]["PARENT_ID"].ToString()) : -1,
                HO = tb.Rows[0]["HO"] != DBNull.Value ? int.Parse(tb.Rows[0]["HO"].ToString()) : -1,
                Used = tb.Rows[0]["USED"] != DBNull.Value ? int.Parse(tb.Rows[0]["USED"].ToString()) : -1,
                SyncUser = tb.Rows[0]["SYNC_USER"] != DBNull.Value ? tb.Rows[0]["SYNC_USER"].ToString() : string.Empty,
                SyncDate = tb.Rows[0]["SYNC_DATE"] != DBNull.Value ? (DateTime?)DateTime.Parse(tb.Rows[0]["SYNC_DATE"].ToString()) : null,
                HeadOfUnit = tb.Rows[0]["HEAD_OF_UNIT"] != DBNull.Value ? tb.Rows[0]["HEAD_OF_UNIT"].ToString() : string.Empty,
            };
        }

        public static Kpi GetKpi(string userId, string Id)
        {
            DataTable tb = DbHelper.GetKpi(userId, Id);
            if (tb == null || tb.Rows.Count == 0)
                return null;
            DataRow row = tb.Rows[0];
            return new Kpi()
            {
                Id = row["ID"] != DBNull.Value ? row["ID"].ToString() : string.Empty,
                Name = row["NAME"] != DBNull.Value ? row["NAME"].ToString() : string.Empty,
                TypeId = row["TYPE_ID"] != DBNull.Value ? row["TYPE_ID"].ToString() : string.Empty,
                Description = row["DESCRIPTION"] != DBNull.Value ? row["DESCRIPTION"].ToString() : string.Empty,
                GroupId = row["GROUP_ID"] != DBNull.Value ? row["GROUP_ID"].ToString() : string.Empty,
                Method = row["METHOD"] != DBNull.Value ? row["METHOD"].ToString() : string.Empty,
                Format = row["FORMAT"] != DBNull.Value ? row["FORMAT"].ToString() : string.Empty,
                Formular = row["FORMULAR"] != DBNull.Value ? row["FORMULAR"].ToString() : string.Empty,
                OuId = row["OU_ID"] != DBNull.Value ? row["OU_ID"].ToString() : string.Empty,
                Used = row["USED"] != DBNull.Value ? int.Parse(row["USED"].ToString()) : 0,
                UserCreated = row["USER_CREATED"] != DBNull.Value ? row["USER_CREATED"].ToString() : string.Empty,
                DateCreated = row["DATE_CREATED"] != DBNull.Value ? (DateTime?)DateTime.Parse(row["DATE_CREATED"].ToString()) : null,
            };
        }

        public static DataSet GetInternalSatisfactionResult(string userId, int periodId) => DbHelper.GetInternalSatisfactionResult(userId, periodId);

        public static DataSet GetInternalSatisfactionReport(string userId, string ouId, int periodId) => DbHelper.GetInternalSatisfactionReport(userId, ouId, periodId);

        public static DataTable GetKpiType(string userId) => DbHelper.GetKpiType(userId);

        public static DataTable GetKpiGroup(string userId) => DbHelper.GetKpiGroup(userId);

        public static DataTable GetKpiFormular(string userId) => DbHelper.GetKpiFormular(userId);

        public static DataTable GetKpiFormat(string userId) => DbHelper.GetKpiFormat(userId);

        public static DataTable GetUserOu(string userId, string employeeNo, int? isAllChilren, int? roleId) => DbHelper.GetUserOu(userId, employeeNo, isAllChilren, roleId);

        public static DataTable GetKotByRole(string userId, string employeeNo) => DbHelper.GetKotByRole(userId, employeeNo);

        public static DataTable SearchKpi(string userId, int? status, string kpiId, string ouId, string kpiType, string groupId, string kpiName) => DbHelper.SearchKpi(userId, status, kpiId, ouId, kpiType, groupId, kpiName);

        public static DataTable SearchKotList(string userId, int? status, string kotId, string kotName, string ouId) => DbHelper.SearchKotList(userId, status, kotId, kotName, ouId);

        public static DataTable SearchKotEmployees(string userId, int? status, string ouId, int periodId) => DbHelper.SearchKotEmployees(userId, status, ouId, periodId);

        public static DataSet GetKotDetail(string userId, int kotId) => DbHelper.GetKotDetail(userId, kotId);

        public static DataTable GetEmployeesWaitApprove(string userId, int periodId, int? status) => DbHelper.GetEmployeesWaitApprove(userId, periodId, status);

        public static DataSet GetPersonalEvaluatioinReport(string userId, string employeeNo, int periodId) => DbHelper.GetPersonalEvaluatioinReport(userId, employeeNo, periodId);

        public static DataTable GetModulesByRole(string userId, int roleId) => DbHelper.GetModulesByRole(userId, roleId);

        public static DataTable GetUsersByRole(string userId, int roleId) => DbHelper.GetUsersByRole(userId, roleId);

        public static bool DoCreatePeriod(int? id, string name, int? year, int active, string userId, ref string error)
        {
            return DbHelper.DoCreatePeriod(id, name, year, active, userId, ref error);
        }

        public static bool DoSyncCompanyStructure(string userId, int type, ref string error)
        {
            bool kq = false;
            result rs;
            switch (type)
            {
                case 1:
                    rs = SyncCompanyData(userId, type);
                    error = rs.Message;
                    kq = "0".Equals(rs.Code);
                    break;
                case 2:
                    rs = SyncOuGroupData(userId, type);
                    error = rs.Message;
                    kq = "0".Equals(rs.Code);
                    break;
                case 3:
                    rs = SyncOuData(userId, type);
                    error = rs.Message;
                    kq = "0".Equals(rs.Code);
                    break;
                case 4:
                    rs = SyncDepartmentData(userId, type);
                    error = rs.Message;
                    kq = "0".Equals(rs.Code);
                    break;
                case 5:
                    rs = SyncTeamData(userId, type);
                    error = rs.Message;
                    kq = "0".Equals(rs.Code);
                    break;
                case 6:
                    rs = SyncGroupData(userId, type);
                    error = rs.Message;
                    kq = "0".Equals(rs.Code);
                    break;
                case 7:
                    rs = SyncJobTitleData(userId, type);
                    error = rs.Message;
                    kq = "0".Equals(rs.Code);
                    break;
                default:
                    break;
            }
            return kq;
        }
        private static result SyncCompanyData(string userId, int type)
        {
            try
            {
                LMS.ResponseOCB lmsResponse;
                EOFFICE.ResponseOCB eResponse;
                try
                {
                    string lmsAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.LMS"];
                    LMS.OCBBankClient oCBBankClient = new LMS.OCBBankClient();
                    lmsResponse = oCBBankClient.iHRP_LMS_LayDMCongTy(lmsAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "LMS request failed: " + ex.ToString() };
                }

                try
                {
                    string eofficeAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.EOFFICE"];
                    EOFFICE.OCB1Client oCB1Client = new EOFFICE.OCB1Client();
                    eResponse = oCB1Client.iHRP_LMS_LayTTCoCau(eofficeAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "EOFFICE request failed: " + ex.ToString() };
                }

                if (lmsResponse != null && eResponse != null)
                {
                    if (!"0".Equals(lmsResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from LMS: " + lmsResponse.Message };
                    }
                    if (!"0".Equals(eResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from EOFFICE: " + eResponse.Message };
                    }

                    List<PostedCompany> list = new List<PostedCompany>();
                    foreach (XElement company in lmsResponse.Result.Elements())
                    {
                        PostedCompany cp = new PostedCompany()
                        {
                            ID = company.Element("LSCompanyID").Value,
                            CODE = company.Element("LSCompanyCode").Value,
                            NAME = company.Element("VNName").Value,
                            USED = "true".Equals(company.Element("Used").Value) ? "1" : "0",

                        };

                        var eElements = eResponse.Result.Elements();
                        var e = eElements.Where(q => q.Element("MaDonVi").Value == cp.CODE).FirstOrDefault();
                        if (e != null)
                            cp.MIXED_CODE = e.Element("MaTongHop").Value;
                        list.Add(cp);
                    }

                    if (list.Count == 0)
                        return new result() { Code = "1", Message = "Không có dòng nào" };
                    else
                    {
                        string error = string.Empty;
                        return new result() { Code = DbHelper.CreateCompanyStructure(userId, type, list, ref error) ? "0" : "1", Message = error };
                    }
                }

                return new result() { Code = "1", Message = "Request NULL" };
            }
            catch (Exception ex)
            {
                return new result() { Code = "1", Message = ex.Message };
            }

        }
        private static result SyncOuGroupData(string userId, int type)
        {
            try
            {
                LMS.ResponseOCB lmsResponse;
                EOFFICE.ResponseOCB eResponse;
                try
                {
                    string lmsAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.LMS"];
                    LMS.OCBBankClient oCBBankClient = new LMS.OCBBankClient();
                    lmsResponse = oCBBankClient.iHRP_LMS_LayDMKhoi(lmsAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "LMS request failed: " + ex.ToString() };
                }

                try
                {
                    string eofficeAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.EOFFICE"];
                    EOFFICE.OCB1Client oCB1Client = new EOFFICE.OCB1Client();
                    eResponse = oCB1Client.iHRP_LMS_LayTTCoCau(eofficeAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "EOFFICE request failed: " + ex.ToString() };
                }

                if (lmsResponse != null && eResponse != null)
                {
                    if (!"0".Equals(lmsResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from LMS: " + lmsResponse.Message };
                    }
                    if (!"0".Equals(eResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from EOFFICE: " + eResponse.Message };
                    }

                    List<PostedCompany> list = new List<PostedCompany>();
                    foreach (XElement company in lmsResponse.Result.Elements())
                    {
                        PostedCompany cp = new PostedCompany()
                        {
                            ID = company.Element("LSLevel1ID").Value,
                            CODE = company.Element("LSLevel1Code").Value,
                            NAME = company.Element("VNName").Value,
                            USED = "true".Equals(company.Element("Used").Value) ? "1" : "0",
                            PARENT = company.Element("LSCompanyID").Value

                        };

                        var eElements = eResponse.Result.Elements();
                        var e = eElements.Where(q => q.Element("MaDonVi").Value == cp.CODE).FirstOrDefault();
                        if (e != null)
                            cp.MIXED_CODE = e.Element("MaTongHop").Value;
                        list.Add(cp);
                    }

                    if (list.Count == 0)
                        return new result() { Code = "1", Message = "Không có dòng nào" };
                    else
                    {
                        string error = string.Empty;
                        return new result() { Code = DbHelper.CreateCompanyStructure(userId, type, list, ref error) ? "0" : "1", Message = error };
                    }
                }

                return new result() { Code = "1", Message = "Request NULL" };
            }
            catch (Exception ex)
            {
                return new result() { Code = "1", Message = ex.Message };
            }
        }
        private static result SyncOuData(string userId, int type)
        {
            try
            {
                LMS.ResponseOCB lmsResponse;
                EOFFICE.ResponseOCB eResponse;
                try
                {
                    string lmsAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.LMS"];
                    LMS.OCBBankClient oCBBankClient = new LMS.OCBBankClient();
                    lmsResponse = oCBBankClient.iHRP_LMS_LayDMDonVi(lmsAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "LMS request failed: " + ex.ToString() };
                }

                try
                {
                    string eofficeAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.EOFFICE"];
                    EOFFICE.OCB1Client oCB1Client = new EOFFICE.OCB1Client();
                    eResponse = oCB1Client.iHRP_LMS_LayTTCoCau(eofficeAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "EOFFICE request failed: " + ex.ToString() };
                }

                if (lmsResponse != null && eResponse != null)
                {
                    if (!"0".Equals(lmsResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from LMS: " + lmsResponse.Message };
                    }
                    if (!"0".Equals(eResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from EOFFICE: " + eResponse.Message };
                    }

                    List<PostedCompany> list = new List<PostedCompany>();
                    foreach (XElement company in lmsResponse.Result.Elements())
                    {
                        PostedCompany cp = new PostedCompany()
                        {
                            ID = company.Element("LSLevel2ID").Value,
                            CODE = company.Element("LSLevel2Code").Value,
                            NAME = company.Element("VNName").Value,
                            USED = "true".Equals(company.Element("Used").Value) ? "1" : "0",
                            PARENT = company.Element("LSLevel1ID").Value

                        };

                        var eElements = eResponse.Result.Elements();
                        var e = eElements.Where(q => q.Element("MaDonVi").Value == cp.CODE).FirstOrDefault();
                        if (e != null)
                            cp.MIXED_CODE = e.Element("MaTongHop").Value;
                        list.Add(cp);
                    }

                    if (list.Count == 0)
                        return new result() { Code = "1", Message = "Không có dòng nào" };
                    else
                    {
                        string error = string.Empty;
                        return new result() { Code = DbHelper.CreateCompanyStructure(userId, type, list, ref error) ? "0" : "1", Message = error };
                    }
                }

                return new result() { Code = "1", Message = "Request NULL" };
            }
            catch (Exception ex)
            {
                return new result() { Code = "1", Message = ex.Message };
            }
        }
        private static result SyncDepartmentData(string userId, int type)
        {
            try
            {
                LMS.ResponseOCB lmsResponse;
                EOFFICE.ResponseOCB eResponse;
                try
                {
                    string lmsAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.LMS"];
                    LMS.OCBBankClient oCBBankClient = new LMS.OCBBankClient();
                    lmsResponse = oCBBankClient.iHRP_LMS_LayDMPhongBan(lmsAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "LMS request failed: " + ex.ToString() };
                }

                try
                {
                    string eofficeAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.EOFFICE"];
                    EOFFICE.OCB1Client oCB1Client = new EOFFICE.OCB1Client();
                    eResponse = oCB1Client.iHRP_LMS_LayTTCoCau(eofficeAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "EOFFICE request failed: " + ex.ToString() };
                }

                if (lmsResponse != null && eResponse != null)
                {
                    if (!"0".Equals(lmsResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from LMS: " + lmsResponse.Message };
                    }
                    if (!"0".Equals(eResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from EOFFICE: " + eResponse.Message };
                    }

                    List<PostedCompany> list = new List<PostedCompany>();
                    foreach (XElement company in lmsResponse.Result.Elements())
                    {
                        PostedCompany cp = new PostedCompany()
                        {
                            ID = company.Element("LSLevel3ID").Value,
                            CODE = company.Element("LSLevel3Code").Value,
                            NAME = company.Element("VNName").Value,
                            USED = "true".Equals(company.Element("Used").Value) ? "1" : "0",
                            PARENT = company.Element("LSLevel2ID").Value

                        };

                        var eElements = eResponse.Result.Elements();
                        var e = eElements.Where(q => q.Element("MaDonVi").Value == cp.CODE).FirstOrDefault();
                        if (e != null)
                            cp.MIXED_CODE = e.Element("MaTongHop").Value;
                        list.Add(cp);
                    }

                    if (list.Count == 0)
                        return new result() { Code = "1", Message = "Không có dòng nào" };
                    else
                    {
                        string error = string.Empty;
                        return new result() { Code = DbHelper.CreateCompanyStructure(userId, type, list, ref error) ? "0" : "1", Message = error };
                    }
                }

                return new result() { Code = "1", Message = "Request NULL" };
            }
            catch (Exception ex)
            {
                return new result() { Code = "1", Message = ex.Message };
            }
        }
        private static result SyncTeamData(string userId, int type)
        {
            try
            {
                LMS.ResponseOCB lmsResponse;
                EOFFICE.ResponseOCB eResponse;
                try
                {
                    string lmsAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.LMS"];
                    LMS.OCBBankClient oCBBankClient = new LMS.OCBBankClient();
                    lmsResponse = oCBBankClient.iHRP_LMS_LayDMBoPhan(lmsAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "LMS request failed: " + ex.ToString() };
                }

                try
                {
                    string eofficeAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.EOFFICE"];
                    EOFFICE.OCB1Client oCB1Client = new EOFFICE.OCB1Client();
                    eResponse = oCB1Client.iHRP_LMS_LayTTCoCau(eofficeAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "EOFFICE request failed: " + ex.ToString() };
                }

                if (lmsResponse != null && eResponse != null)
                {
                    if (!"0".Equals(lmsResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from LMS: " + lmsResponse.Message };
                    }
                    if (!"0".Equals(eResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from EOFFICE: " + eResponse.Message };
                    }

                    List<PostedCompany> list = new List<PostedCompany>();
                    foreach (XElement company in lmsResponse.Result.Elements())
                    {
                        PostedCompany cp = new PostedCompany()
                        {
                            ID = company.Element("LSLevel4ID").Value,
                            CODE = company.Element("LSLevel4Code").Value,
                            NAME = company.Element("VNName").Value,
                            USED = "true".Equals(company.Element("Used").Value) ? "1" : "0",
                            PARENT = company.Element("LSLevel3ID").Value

                        };

                        var eElements = eResponse.Result.Elements();
                        var e = eElements.Where(q => q.Element("MaDonVi").Value == cp.CODE).FirstOrDefault();
                        if (e != null)
                            cp.MIXED_CODE = e.Element("MaTongHop").Value;
                        list.Add(cp);
                    }

                    if (list.Count == 0)
                        return new result() { Code = "1", Message = "Không có dòng nào" };
                    else
                    {
                        string error = string.Empty;
                        return new result() { Code = DbHelper.CreateCompanyStructure(userId, type, list, ref error) ? "0" : "1", Message = error };
                    }
                }

                return new result() { Code = "1", Message = "Request NULL" };
            }
            catch (Exception ex)
            {
                return new result() { Code = "1", Message = ex.Message };
            }
        }
        private static result SyncGroupData(string userId, int type)
        {
            try
            {
                LMS.ResponseOCB lmsResponse;
                EOFFICE.ResponseOCB eResponse;
                try
                {
                    string lmsAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.LMS"];
                    LMS.OCBBankClient oCBBankClient = new LMS.OCBBankClient();
                    lmsResponse = oCBBankClient.iHRP_LMS_LayDMToNhom(lmsAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "LMS request failed: " + ex.ToString() };
                }

                try
                {
                    string eofficeAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.EOFFICE"];
                    EOFFICE.OCB1Client oCB1Client = new EOFFICE.OCB1Client();
                    eResponse = oCB1Client.iHRP_LMS_LayTTCoCau(eofficeAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "EOFFICE request failed: " + ex.ToString() };
                }

                if (lmsResponse != null && eResponse != null)
                {
                    if (!"0".Equals(lmsResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from LMS: " + lmsResponse.Message };
                    }
                    if (!"0".Equals(eResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from EOFFICE: " + eResponse.Message };
                    }

                    List<PostedCompany> list = new List<PostedCompany>();
                    foreach (XElement company in lmsResponse.Result.Elements())
                    {
                        PostedCompany cp = new PostedCompany()
                        {
                            ID = company.Element("LSLevel5ID").Value,
                            CODE = company.Element("LSLevel5Code").Value,
                            NAME = company.Element("VNName").Value,
                            USED = "true".Equals(company.Element("Used").Value) ? "1" : "0",
                            PARENT = company.Element("LSLevel4ID").Value

                        };

                        var eElements = eResponse.Result.Elements();
                        var e = eElements.Where(q => q.Element("MaDonVi").Value == cp.CODE).FirstOrDefault();
                        if (e != null)
                            cp.MIXED_CODE = e.Element("MaTongHop").Value;
                        list.Add(cp);
                    }

                    if (list.Count == 0)
                        return new result() { Code = "1", Message = "Không có dòng nào" };
                    else
                    {
                        string error = string.Empty;
                        return new result() { Code = DbHelper.CreateCompanyStructure(userId, type, list, ref error) ? "0" : "1", Message = error };
                    }
                }

                return new result() { Code = "1", Message = "Request NULL" };
            }
            catch (Exception ex)
            {
                return new result() { Code = "1", Message = ex.Message };
            }
        }
        private static result SyncJobTitleData(string userId, int type)
        {
            try
            {
                LMS.ResponseOCB lmsResponse;
                try
                {
                    string lmsAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.LMS"];
                    LMS.OCBBankClient oCBBankClient = new LMS.OCBBankClient();
                    lmsResponse = oCBBankClient.iHRP_LMS_LayDMChucDanh(lmsAuthenKey);
                }
                catch (Exception ex)
                {
                    return new result() { Code = "1", Message = "LMS request failed: " + ex.ToString() };
                }

                if (lmsResponse != null)
                {
                    if (!"0".Equals(lmsResponse.Code))
                    {
                        return new result() { Code = lmsResponse.Code, Message = "Error from LMS: " + lmsResponse.Message };
                    }

                    List<PostedJobTitle> list = new List<PostedJobTitle>();
                    foreach (XElement jobTitle in lmsResponse.Result.Elements())
                    {
                        PostedJobTitle cp = new PostedJobTitle()
                        {
                            ID = jobTitle.Element("LSJobTitleID").Value,
                            CODE = jobTitle.Element("LSJobTitleCode").Value,
                            NAME = jobTitle.Element("VNName").Value,
                            USED = "true".Equals(jobTitle.Element("Used").Value) ? "1" : "0"
                        };

                        list.Add(cp);
                    }

                    if (list.Count == 0)
                        return new result() { Code = "1", Message = "Không có dòng nào" };
                    else
                    {
                        string error = string.Empty;
                        return new result() { Code = DbHelper.CreateJobTitle(userId, type, list, ref error) ? "0" : "1", Message = error };
                    }
                }

                return new result() { Code = "1", Message = "Request NULL" };
            }
            catch (Exception ex)
            {
                return new result() { Code = "1", Message = ex.Message };
            }
        }

        public static bool DoSyncEmployee(string userId, string from, string to, ref string error)
        {
            try
            {
                DateTime fromDate = DateTime.ParseExact(from, new string[] { "d/M/yyyy", "dd/MM/yyyy" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None);
                DateTime toDate = DateTime.ParseExact(to, new string[] { "d/M/yyyy", "dd/MM/yyyy" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None);





                int fromYear = fromDate.Year;
                int toYear = toDate.Year;

                string dataDate = string.Format("{0}-{1}", fromDate.ToString("ddMMyyyy"), toDate.ToString("ddMMyyyy"));

                try
                {
                    String eoffAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.EOFFICE"];
                    EOFFICE.OCB1Client client = new EOFFICE.OCB1Client();
                    EOFFICE.ResponseOCB response = client.iHRP_LMS_LayTTTongHop(fromDate.ToString("yyyy-dd-MM 00:00:00"), toDate.ToString("yyyy-dd-MM 00:00:00"), eoffAuthenKey);

                    List<result> l = new List<result>();

                    foreach (XElement employee in response.Result.Elements())
                    {
                        var empNo = employee.Element("MaNV");
                        var mixedCode = employee.Element("MaTongHop");
                        l.Add(new result()
                        {
                            Code = empNo != null ? empNo.Value : string.Empty,
                            Message = mixedCode != null ? mixedCode.Value : string.Empty
                        });
                    }



                    string lmsAuthenKey = ConfigurationManager.AppSettings["AUT.IHRP.LMS"];
                    LMS.OCBBankClient oCBBankClient = new LMS.OCBBankClient();
                    List<Task<LMS.iHRP_LMS_LayThongTinNhanVienResponse>> list = new List<Task<LMS.iHRP_LMS_LayThongTinNhanVienResponse>>();
                    for (int i = fromYear; i <= toYear; i++)
                    {
                        //if (i != fromYear)
                        //    fromDate = new DateTime(i, 1, 1);
                        //if (i != toYear)
                        //    toDate = new DateTime(i, 12, 31);
                        list.Add(ProcessURLAsync(lmsAuthenKey, i, fromDate, toDate, oCBBankClient));
                    }

                    Task.WaitAll(list.ToArray());
                    List<PostedEmployee> ls = new List<PostedEmployee>();
                    foreach (var item in list)
                    {
                        LMS.iHRP_LMS_LayThongTinNhanVienResponse lmsResponse = item.Result;
                        if (lmsResponse != null && lmsResponse.Body != null && lmsResponse.Body.iHRP_LMS_LayThongTinNhanVienResult != null)
                        {
                            if (!"0".Equals(lmsResponse.Body.iHRP_LMS_LayThongTinNhanVienResult.Code) && !"Response data is empty: 0 rows".Equals(lmsResponse.Body.iHRP_LMS_LayThongTinNhanVienResult.Message))
                            {
                                error = "LMS request failed: " + lmsResponse.Body.iHRP_LMS_LayThongTinNhanVienResult.Message;
                                return false;
                            }


                            foreach (XElement employee in lmsResponse.Body.iHRP_LMS_LayThongTinNhanVienResult.Result.Elements())
                            {
                                XElement empCode = employee.Element("EmpCode");
                                XElement VLastName = employee.Element("VLastName");
                                XElement VFirstName = employee.Element("VFirstName");
                                XElement Email = employee.Element("Email");
                                XElement DOB = employee.Element("DOB");
                                XElement StartDate = employee.Element("StartDate");
                                XElement StartDateCompany = employee.Element("StartDateCompany");
                                XElement LastWorkingDate = employee.Element("LastWorkingDate");
                                XElement EmployeeType = employee.Element("EmployeeType");
                                XElement ContractType = employee.Element("ContractType");
                                XElement MixedCode = employee.Element("MixedCode");
                                XElement JobTitle = employee.Element("JobTitle");

                                DateTime dob = new DateTime();
                                DateTime owd = new DateTime();
                                DateTime scd = new DateTime();
                                DateTime sd = new DateTime();
                                string userDomain = Email != null ? Email.Value.Split('@')[0] : string.Empty;
                                if (DOB != null)
                                    DateTime.TryParseExact(DOB.Value, new string[] { "dd/MM/yyyy", "d/M/yyyy" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dob);
                                if (LastWorkingDate != null)
                                    DateTime.TryParseExact(LastWorkingDate.Value, new string[] { "dd/MM/yyyy", "d/M/yyyy" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out owd);
                                if (StartDateCompany != null)
                                    DateTime.TryParseExact(StartDateCompany.Value, new string[] { "dd/MM/yyyy", "d/M/yyyy" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out scd);
                                if (StartDate != null)
                                    DateTime.TryParseExact(StartDate.Value, new string[] { "dd/MM/yyyy", "d/M/yyyy" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out sd);

                                PostedEmployee e = new PostedEmployee()
                                {
                                    EMPLOYEE_NO = empCode != null ? empCode.Value : string.Empty,
                                    NAME = VLastName != null && VFirstName != null ? string.Format("{0} {1}", VLastName.Value, VFirstName.Value) : string.Empty,
                                    DOB = dob != DateTime.MinValue ? dob.ToString("dd/MM/yyyy") : null,
                                    LEAVE_OFF_WORK_DATE = owd != DateTime.MinValue ? owd.ToString("dd/MM/yyyy") : null,
                                    START_COMPANY_DATE = scd != DateTime.MinValue ? scd.ToString("dd/MM/yyyy") : null,
                                    START_DATE = sd != DateTime.MinValue ? sd.ToString("dd/MM/yyyy") : null,
                                    USER_DOMAIN = userDomain,
                                    USER_DOMAIN_CHECK_PASS = userDomain,
                                    CONTRACT_TYPE = ContractType != null ? ContractType.Value : string.Empty,
                                    EMPLOYEE_TYPE = EmployeeType != null ? EmployeeType.Value : string.Empty,
                                    //OU_ID =  ;//MixedCode != null ? MixedCode.Value : string.Empty,
                                    JOB_TITLE_ID = JobTitle != null ? JobTitle.Value : string.Empty,
                                    DATA_DATE = dataDate
                                };
                                var el = l.Where(q => q.Code == e.EMPLOYEE_NO).FirstOrDefault();
                                string[] ou = el != null ? el.Message.Split('-') : new string[0];
                                e.OU_ID = ou.Length > 0 ? ou[ou.Length - 1] : string.Empty;
                                ls.Add(e);
                            }
                        }
                    }
                    if (ls.Count == 0)
                    {
                        error = "Không có dòng nào";
                        return false;
                    }
                    return DbHelper.CreateEmployee(userId, ls, fromDate, toDate, ref error);

                }
                catch (Exception ex)
                {
                    error = "LMS request failed: " + ex.ToString();
                    return false;
                }
            }
            catch (Exception ex)
            {
                error = "LMS request failed: " + ex.ToString();
                return false;
            }
        }

        public static async Task<LMS.iHRP_LMS_LayThongTinNhanVienResponse> ProcessURLAsync(string authenKey, int year, DateTime fromDate, DateTime toDate, LMS.OCBBankClient client)
        {
            LMS.iHRP_LMS_LayThongTinNhanVienResponse response = await client.iHRP_LMS_LayThongTinNhanVienAsync(string.Empty, 0, 90000, year, 1, authenKey, fromDate.ToString("yyyy-dd-MM 00:00:00"), toDate.ToString("yyyy-dd-MM 00:00:00")).ConfigureAwait(false);
            return response;
        }

        public static bool SaveDeclareEvaluationOu(string userId, List<PostedDeclareEvalOu> list, int? periodId, int hasSLA, ref string error)
        {
            return DbHelper.SaveDeclareEvaluationOu(userId, list, periodId, hasSLA, ref error);
        }

        public static bool SaveEvaluationUser(string userId, string ouId, string employeeNo, int periodId, ref string error)
        {
            return DbHelper.SaveEvaluationUser(userId, ouId, employeeNo, periodId, ref error);
        }

        public static bool SaveHeadOfUnit(string userId, string ouId, string employeeNo, ref string error)
        {
            return DbHelper.SaveHeadOfUnit(userId, ouId, employeeNo, ref error);
        }

        public static bool SaveInternalSatisfaction(string userId, string ouId, string evaluationOuId, int periodId, int completed, List<PostedInternalSatisfactionQuestion> postedList, List<PostedInternalSatisfactionSubQuestion> subPostedList, ref string error)
        {
            return DbHelper.SaveInternalSatisfaction(userId, ouId, evaluationOuId, periodId, completed, postedList, subPostedList, ref error);
        }

        public static byte[] ExportInternalSatisfactionResult(string userId, int periodId)
        {
            DataSet ds = GetInternalSatisfactionResult(userId, periodId);

            DataTable data = GetEvaluationByOuId(userId, string.Empty, periodId);

            Period period = GetPeriodInfo(userId, periodId);

            DataTable header = ds.Tables[0];
            DataTable list = ds.Tables[1];
            MemoryStream stream = new MemoryStream();
            using (var workBook = new DevExpress.Spreadsheet.Workbook())
            {
                workBook.LoadDocument(HttpContext.Current.Server.MapPath("~/Template/HLNB/HLNB_RESULT.xlsx"));

                DevExpress.Spreadsheet.Worksheet sheet = workBook.Worksheets[0];
                DevExpress.Spreadsheet.Style lStyle = workBook.Styles.Add("leftStyle");
                lStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Alignment.WrapText = true;

                DevExpress.Spreadsheet.Style cStyle = workBook.Styles.Add("centerStyle");
                cStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                cStyle.Alignment.WrapText = true;

                DevExpress.Spreadsheet.Style rStyle = workBook.Styles.Add("rightStyle");
                rStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right;
                rStyle.Alignment.WrapText = true;

                DevExpress.Spreadsheet.Style rsStyle = workBook.Styles.Add("rightMoneyStyle");
                rsStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rsStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rsStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rsStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rsStyle.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right;
                rsStyle.NumberFormat = "#,#";

                DevExpress.Spreadsheet.Cell cTitle = sheet.Cells[1, 0];
                cTitle.SetValue(period != null ? period.Name : string.Empty);

                var rheader = header != null && header.Rows.Count > 0 ? header.Rows[0] : null;
                if (rheader != null)
                {
                    sheet.Cells[2, 2].SetValue(rheader["COUNT"]);
                    sheet.Cells[3, 2].SetValue(rheader["MAX"] != DBNull.Value ? (double?)double.Parse(rheader["MAX"].ToString()) / 100 : null);
                    sheet.Cells[4, 2].SetValue(rheader["MIN"] != DBNull.Value ? (double?)double.Parse(rheader["MIN"].ToString()) / 100 : null);
                    sheet.Cells[5, 2].SetValue(rheader["AVG"] != DBNull.Value ? (double?)double.Parse(rheader["AVG"].ToString()) / 100 : null);
                }
                int i = 1;
                int index = 9;
                foreach (DataRow item in list.Rows)
                {
                    DevExpress.Spreadsheet.Cell c0 = sheet.Cells[index, 0];
                    c0.SetValue(i++);
                    c0.Style = cStyle;

                    DevExpress.Spreadsheet.Cell c1 = sheet.Cells[index, 1];
                    c1.SetValue(item["NAME"]);
                    c1.Style = lStyle;

                    DevExpress.Spreadsheet.Cell c2 = sheet.Cells[index, 2];
                    c2.SetValue(item["COUNT"]);
                    c2.Style = cStyle;

                    DevExpress.Spreadsheet.Cell c2_plus = sheet.Cells[index, 3];
                    c2_plus.SetValue(item["COUNT_C"]);
                    c2_plus.Style = cStyle;

                    DevExpress.Spreadsheet.Cell c3 = sheet.Cells[index, 4];
                    c3.SetValue(item["RESULT"] != DBNull.Value ? (double?)double.Parse(item["RESULT"].ToString()) / 100 : null);
                    c3.Style = rStyle;
                    c3.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c4 = sheet.Cells[index, 5];
                    c4.SetValue(item["MAX"] != DBNull.Value ? (double?)double.Parse(item["MAX"].ToString()) / 100 : null);
                    c4.Style = cStyle;
                    c4.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c5 = sheet.Cells[index, 6];
                    c5.SetValue(item["MIN"] != DBNull.Value ? (double?)double.Parse(item["MIN"].ToString()) / 100 : null);
                    c5.Style = cStyle;
                    c5.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c6 = sheet.Cells[index++, 7];
                    c6.SetValue(item["RANKING"]);
                    c6.Style = cStyle;
                }
                index = 1;
                var sheet2 = workBook.Worksheets[1];
                foreach (DataRow item in data.Rows)
                {
                    DevExpress.Spreadsheet.Cell ci0 = sheet2.Cells[index, 0];
                    ci0.SetValue(item["NAME_X"]);
                    ci0.Style = lStyle;

                    DevExpress.Spreadsheet.Cell ci1 = sheet2.Cells[index, 1];
                    ci1.SetValue(item["NAME"]);
                    ci1.Style = lStyle;

                    string str = string.Empty;
                    DevExpress.Spreadsheet.Cell ci2 = sheet2.Cells[index++, 2];
                    if (item["COMPLETED"] == DBNull.Value)
                    {
                        str = "CHƯA ĐÁNH GIÁ";
                        if (item["USER_DOMAIN"] == DBNull.Value && item["USER_DOMAIN_2"] == DBNull.Value)
                        {
                            str = "KHÔNG CÓ THÔNG TIN NGƯỜI ĐẠI ĐIỆN ĐÁNH GIÁ HAY TRƯỞNG ĐƠN VỊ";
                        }
                        else if (item["USER_DOMAIN"] != DBNull.Value)
                        {
                            str = item["USER_DOMAIN"].ToString() + " - " + item["EMPLOYEE_NO"].ToString() + " - " + item["EMPLOYEE_NAME"].ToString();
                        }
                        else if (item["USER_DOMAIN_2"] != DBNull.Value)
                        {
                            str = item["USER_DOMAIN_2"].ToString() + " - " + item["EMPLOYEE_NO_2"].ToString() + " - " + item["EMPLOYEE_NAME_2"].ToString();
                        }

                    }
                    else if (item["COMPLETED"].ToString() == "1")
                    {
                        str = "ĐÃ HOÀN THÀNH";
                        if (item["VALUE_X"] != DBNull.Value)
                        {
                            int kq = int.Parse(item["VALUE_X"].ToString());
                            if (kq == 1 || kq == 2)
                            {
                                str += ". (Có check 1 hoặc 2 điểm)";
                            }
                        }
                    }
                    else
                    {
                        str = "CHƯA NỘP KẾT QUẢ." + item["USER_DOMAIN"].ToString() + " - " + item["EMPLOYEE_NO"].ToString() + " - " + item["EMPLOYEE_NAME"].ToString();
                    }

                    ci2.SetValue(str);
                    ci2.Style = lStyle;
                }

                workBook.SaveDocument(stream, DevExpress.Spreadsheet.DocumentFormat.Xlsx);
                return stream.ToArray();
            }
        }

        public static byte[] ExportInternalSatisfactionReport(string userId, string ouId, int periodId)
        {



            DataSet ds = GetInternalSatisfactionReport(userId, ouId, periodId);
            Period period = GetPeriodInfo(userId, periodId);

            DataTable header = ds.Tables[0];
            DataTable list = ds.Tables[1];
            MemoryStream stream = new MemoryStream();
            using (var workBook = new DevExpress.Spreadsheet.Workbook())
            {
                workBook.LoadDocument(HttpContext.Current.Server.MapPath("~/Template/HLNB/HLNB_REPORT.xlsx"));

                DevExpress.Spreadsheet.Worksheet sheet = workBook.Worksheets[0];
                DevExpress.Spreadsheet.Style lStyle = workBook.Styles.Add("leftStyle");
                lStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                lStyle.Alignment.WrapText = true;

                DevExpress.Spreadsheet.Style cStyle = workBook.Styles.Add("centerStyle");
                cStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                cStyle.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                cStyle.Alignment.WrapText = true;

                DevExpress.Spreadsheet.Style rStyle = workBook.Styles.Add("rightStyle");
                rStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rStyle.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right;
                rStyle.Alignment.WrapText = true;

                DevExpress.Spreadsheet.Style rsStyle = workBook.Styles.Add("rightMoneyStyle");
                rsStyle.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rsStyle.Borders.LeftBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rsStyle.Borders.TopBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rsStyle.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin;
                rsStyle.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right;
                rsStyle.NumberFormat = "#,#";

                DevExpress.Spreadsheet.Cell cTitle = sheet.Cells[1, 0];
                cTitle.SetValue(period != null ? period.Name : string.Empty);

                var rheader = header != null && header.Rows.Count > 0 ? header.Rows[0] : null;
                if (rheader != null)
                {
                    sheet.Cells[2, 2].SetValue(rheader["COUNT"]);
                    double? totalI = null, totalII = null;
                    if (list != null && list.Rows.Count > 0)
                    {
                        var aE = list.AsEnumerable();
                        totalI = aE.Select(q => double.Parse(q["RS_1_123"].ToString())).Sum() / list.Rows.Count;
                        totalII = aE.Select(q => double.Parse(q["RS_2_123"].ToString())).Sum() / list.Rows.Count;
                    }

                    sheet.Cells[3, 2].SetValue(totalI.HasValue ? totalI / 100 : null);
                    sheet.Cells[4, 2].SetValue(totalII.HasValue ? totalII / 100 : null);
                    if (list.Rows.Count > 0)
                        sheet.Cells[5, 2].SetValue(rheader["RESULT"] != DBNull.Value ? (double?)double.Parse(rheader["RESULT"].ToString()) / 100 : null);
                }
                int index = 8;
                foreach (DataRow item in list.Rows)
                {
                    //DevExpress.Spreadsheet.Cell c0 = sheet.Cells[index, 0];
                    //c0.SetValue(i++);
                    //c0.Style = cStyle;

                    DevExpress.Spreadsheet.Cell c1 = sheet.Cells[index, 0];
                    c1.SetValue(item["NAME"]);
                    c1.Style = lStyle;

                    DevExpress.Spreadsheet.Cell c2 = sheet.Cells[index, 1];
                    c2.SetValue(item["RS_1_123"] != DBNull.Value ? (double?)double.Parse(item["RS_1_123"].ToString()) / 100 : null);
                    c2.Style = cStyle;
                    c2.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c3 = sheet.Cells[index, 2];
                    c3.SetValue(item["RS_1_1"] != DBNull.Value ? (double?)double.Parse(item["RS_1_1"].ToString()) / 100 : null);
                    c3.Style = rStyle;
                    c3.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c4 = sheet.Cells[index, 3];
                    c4.SetValue(item["RS_1_2"] != DBNull.Value ? (double?)double.Parse(item["RS_1_2"].ToString()) / 100 : null);
                    c4.Style = rStyle;
                    c4.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c5 = sheet.Cells[index, 4];
                    c5.SetValue(item["RS_1_3"] != DBNull.Value ? (double?)double.Parse(item["RS_1_3"].ToString()) / 100 : null);
                    c5.Style = rStyle;
                    c5.Style.NumberFormat = "0.00%";

                    /////----------///////
                    DevExpress.Spreadsheet.Cell c6 = sheet.Cells[index, 5];
                    c6.SetValue(item["RS_2_123"] != DBNull.Value ? (double?)double.Parse(item["RS_2_123"].ToString()) / 100 : null);
                    c6.Style = cStyle;
                    c6.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c7 = sheet.Cells[index, 6];
                    c7.SetValue(item["RS_2_1"] != DBNull.Value ? (double?)double.Parse(item["RS_2_1"].ToString()) / 100 : null);
                    c7.Style = rStyle;
                    c7.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c8 = sheet.Cells[index, 7];
                    c8.SetValue(item["RS_2_2"] != DBNull.Value ? (double?)double.Parse(item["RS_2_2"].ToString()) / 100 : null);
                    c8.Style = rStyle;
                    c8.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c9 = sheet.Cells[index, 8];
                    c9.SetValue(item["RS_2_3"] != DBNull.Value ? (double?)double.Parse(item["RS_2_3"].ToString()) / 100 : null);
                    c9.Style = rStyle;
                    c9.Style.NumberFormat = "0.00%";

                    DevExpress.Spreadsheet.Cell c10 = sheet.Cells[index++, 9];
                    var sum = (double?)double.Parse(item["RS_1_123"].ToString()) + (double?)double.Parse(item["RS_2_123"].ToString());
                    c10.SetValue(sum / 2 / 100);
                    c10.Style = rStyle;
                    c10.Style.NumberFormat = "0.00%";


                }

                workBook.SaveDocument(stream, DevExpress.Spreadsheet.DocumentFormat.Xlsx);
                return stream.ToArray();
            }

        }

        public static byte[] GetImportHeadOfUnitTemplate(string userId)
        {

            string url = HttpContext.Current.Server.MapPath("~/Template/IMPORT/IMPORT_HEAD_OF_UNIT.xlsx");

            return File.ReadAllBytes(url);
            DataTable tb = GetCompanyStructByType(sys.current.UserInfo.UserId, null);

            var list = tb.AsEnumerable().Select(q => new { Id = q["ID"].ToString(), Code = q["CODE"].ToString(), Name = q["NAME"].ToString(), ParentId = q["PARENT_ID"] != DBNull.Value ? q["PARENT_ID"].ToString() : string.Empty }).ToList();



            MemoryStream stream = new MemoryStream();
            using (var workBook = new DevExpress.Spreadsheet.Workbook())
            {
                workBook.LoadDocument(url);

                DevExpress.Spreadsheet.Worksheet sheet = workBook.Worksheets[0];
                DevExpress.Spreadsheet.Worksheet sheetData = workBook.Worksheets[1];
                int i = 0;
                int x = 0;
                var companyList = list.Where(q => string.IsNullOrEmpty(q.ParentId));
                //create range for company dropdownlist
                foreach (var item in companyList)
                {
                    sheetData.Cells[i++, x].SetValue(string.Format("_{0}_{1}", GetSafeFilename(item.Code), GetSafeFilename(item.Name)));
                }
                var rangeCompany = sheetData.Range[string.Format("$A$1:$A${0}", companyList.Count())];
                rangeCompany.Name = "Range_Company";
                sheet.DataValidations.Add(sheet.Range["A2:A1000"], DevExpress.Spreadsheet.DataValidationType.List, ValueObject.FromRange(rangeCompany));


                i = 0;
                x++;
                foreach (var item in companyList)
                {
                    var unitGroupList = list.Where(q => q.ParentId == item.Id);
                    if (unitGroupList.Count() > 0)
                    {
                        i = 0;
                        foreach (var item2 in unitGroupList)
                        {
                            sheetData.Cells[i++, x].SetValue(string.Format("_{0}_{1}", GetSafeFilename(item2.Code), GetSafeFilename(item2.Name)));
                        }
                        var range = sheetData.Range.FromLTRB(x, 0, x++, unitGroupList.Count() - 1);
                        workBook.DefinedNames.Add(string.Format("_{0}_{1}", GetSafeFilename(item.Code), GetSafeFilename(item.Name)), "DATA!" + range.GetReferenceA1());
                    }
                }
                i = 0;
                foreach (var item in companyList)
                {
                    var unitGroupList = list.Where(q => q.ParentId == item.Id);
                    if (unitGroupList.Count() > 0)
                    {
                        foreach (var item2 in unitGroupList)
                        {
                            var unitList = list.Where(q => q.ParentId == item2.Id);
                            i = 0;
                            if (unitList.Count() > 0)
                            {
                                foreach (var item3 in unitList)
                                {
                                    sheetData.Cells[i++, x].SetValue(string.Format("_{0}_{1}", GetSafeFilename(item3.Code), GetSafeFilename(item3.Name)));
                                }
                                var range = sheetData.Range.FromLTRB(x, 0, x++, unitList.Count() - 1);
                                workBook.DefinedNames.Add(string.Format("_{0}_{1}", GetSafeFilename(item2.Code), GetSafeFilename(item2.Name)), "DATA!" + range.GetReferenceA1());
                            }
                        }
                    }
                }
                i = 0;
                foreach (var item in companyList)
                {
                    var unitGroupList = list.Where(q => q.ParentId == item.Id);
                    if (unitGroupList.Count() > 0)
                    {
                        foreach (var item2 in unitGroupList)
                        {
                            var unitList = list.Where(q => q.ParentId == item2.Id);
                            foreach (var item3 in unitList)
                            {
                                var departmenttList = list.Where(q => q.ParentId == item3.Id);
                                i = 0;
                                if (departmenttList.Count() > 0)
                                {
                                    foreach (var item4 in departmenttList)
                                    {
                                        sheetData.Cells[i++, x].SetValue(string.Format("_{0}_{1}", GetSafeFilename(item4.Code), GetSafeFilename(item4.Name)));
                                    }
                                    var range = sheetData.Range.FromLTRB(x, 0, x++, departmenttList.Count() - 1);
                                    workBook.DefinedNames.Add(string.Format("_{0}_{1}", GetSafeFilename(item3.Code), GetSafeFilename(item3.Name)), "DATA!" + range.GetReferenceA1());
                                }

                            }

                        }
                    }
                }
                i = 0;
                foreach (var item in companyList)
                {
                    var unitGroupList = list.Where(q => q.ParentId == item.Id);
                    if (unitGroupList.Count() > 0)
                    {
                        foreach (var item2 in unitGroupList)
                        {
                            var unitList = list.Where(q => q.ParentId == item2.Id);
                            foreach (var item3 in unitList)
                            {
                                var departmenttList = list.Where(q => q.ParentId == item3.Id);
                                foreach (var item4 in departmenttList)
                                {
                                    var teamList = list.Where(q => q.ParentId == item4.Id);
                                    i = 0;
                                    if (teamList.Count() > 0)
                                    {
                                        foreach (var item5 in teamList)
                                        {
                                            sheetData.Cells[i++, x].SetValue(string.Format("_{0}_{1}", GetSafeFilename(item5.Code), GetSafeFilename(item5.Name)));
                                        }
                                        var range = sheetData.Range.FromLTRB(x, 0, x++, teamList.Count() - 1);
                                        workBook.DefinedNames.Add(string.Format("_{0}_{1}", GetSafeFilename(item4.Code), GetSafeFilename(item4.Name)), "DATA!" + range.GetReferenceA1());
                                    }
                                }

                            }

                        }
                    }
                }
                i = 0;
                foreach (var item in companyList)
                {
                    var unitGroupList = list.Where(q => q.ParentId == item.Id);
                    if (unitGroupList.Count() > 0)
                    {
                        foreach (var item2 in unitGroupList)
                        {
                            var unitList = list.Where(q => q.ParentId == item2.Id);
                            foreach (var item3 in unitList)
                            {
                                var departmenttList = list.Where(q => q.ParentId == item3.Id);
                                foreach (var item4 in departmenttList)
                                {
                                    var teamList = list.Where(q => q.ParentId == item4.Id);
                                    foreach (var item5 in teamList)
                                    {
                                        var groupList = list.Where(q => q.ParentId == item5.Id);
                                        i = 0;
                                        if (groupList.Count() > 0)
                                        {
                                            foreach (var item6 in groupList)
                                            {
                                                sheetData.Cells[i++, x].SetValue(string.Format("_{0}_{1}", GetSafeFilename(item6.Code), GetSafeFilename(item6.Name)));
                                            }
                                            var range = sheetData.Range.FromLTRB(x, 0, x++, groupList.Count() - 1);
                                            workBook.DefinedNames.Add(string.Format("_{0}_{1}", GetSafeFilename(item5.Code), GetSafeFilename(item5.Name)), "DATA!" + range.GetReferenceA1());
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                sheet.DataValidations.Add(sheet.Range["B2:B1000"], DevExpress.Spreadsheet.DataValidationType.List, ValueObject.FromFormula("=INDIRECT(A2)", false));
                sheet.DataValidations.Add(sheet.Range["C2:C1000"], DevExpress.Spreadsheet.DataValidationType.List, ValueObject.FromFormula("=INDIRECT(B2)", false));
                sheet.DataValidations.Add(sheet.Range["D2:D1000"], DevExpress.Spreadsheet.DataValidationType.List, ValueObject.FromFormula("=INDIRECT(C2)", false));
                sheet.DataValidations.Add(sheet.Range["E2:E1000"], DevExpress.Spreadsheet.DataValidationType.List, ValueObject.FromFormula("=INDIRECT(D2)", false));
                sheet.DataValidations.Add(sheet.Range["F2:F1000"], DevExpress.Spreadsheet.DataValidationType.List, ValueObject.FromFormula("=INDIRECT(E2)", false));

                workBook.SaveDocument(stream, DevExpress.Spreadsheet.DocumentFormat.Xlsx);
                return stream.ToArray();
            }

        }

        public static object ReadImportHeadOfUnit(string userId, Stream fileStream, ref string error)
        {
            try
            {
                List<object> list = new List<object>();
                var cps = GetCompanyStructByType(userId, null).AsEnumerable().Select(q => new { Id = q["ID"].ToString(), Code = q["CODE"].ToString(), Name = q["NAME"], ParentId = q["PARENT_ID"] != DBNull.Value ? q["PARENT_ID"].ToString() : string.Empty }).ToList();
                using (var workBook = new Workbook())
                {
                    workBook.LoadDocument(fileStream);
                    Worksheet sheet = workBook.Worksheets[0];
                    var range = sheet.GetDataRange();
                    int row = range.BottomRowIndex;
                    string cp = string.Empty;
                    for (int i = 1; i <= row; i++)
                    {
                        //var company = sheet.GetCellValue(0, i).TextValue;
                        //var unitGroup = sheet.GetCellValue(1, i).TextValue;
                        //var unit = sheet.GetCellValue(2, i).TextValue;
                        //var department = sheet.GetCellValue(3, i).TextValue;
                        //var team = sheet.GetCellValue(4, i).TextValue;
                        //var group = sheet.GetCellValue(5, i).TextValue;
                        var employeeNo = sheet.GetCellValue(2, i).TextValue;

                        //company = string.IsNullOrEmpty(company) ? company : company.Trim();
                        //unitGroup = string.IsNullOrEmpty(unitGroup) ? unitGroup : unitGroup.Trim();
                        //unit = string.IsNullOrEmpty(unit) ? unit : unit.Trim();
                        //department = string.IsNullOrEmpty(department) ? department : department.Trim();
                        //team = string.IsNullOrEmpty(team) ? team : team.Trim();
                        //group = string.IsNullOrEmpty(group) ? group : group.Trim();
                        employeeNo = string.IsNullOrEmpty(employeeNo) ? employeeNo : employeeNo.Trim();

                        //if (!string.IsNullOrEmpty(group))
                        //    cp = group;
                        //else if (!string.IsNullOrEmpty(team))
                        //    cp = team;
                        //else if (!string.IsNullOrEmpty(department))
                        //    cp = department;
                        //else if (!string.IsNullOrEmpty(unit))
                        //    cp = unit;
                        //else if (!string.IsNullOrEmpty(unitGroup))
                        //    cp = unitGroup;
                        //else if (!string.IsNullOrEmpty(company))
                        //    cp = company;

                        cp = sheet.GetCellValue(0, i).TextValue;

                        if (string.IsNullOrEmpty(cp))
                        {
                            error += string.Format("Dòng thứ {0} trên file empty<br/>", i);
                            continue;
                        }
                        if (string.IsNullOrEmpty(employeeNo))
                        {
                            error += string.Format("Dòng thứ {0} trên file, không tách được mã nhân sự ({1})", i, cp);
                            continue;
                        }
                        var o = cps.Where(q => cp.Equals(q.Code)).FirstOrDefault();
                        if (o == null)
                            error += string.Format("Dòng thứ {0} trên file, tách mã cơ cấu nhưng không tìm thấy trên hệ thống <br/> ({1})", i, cp);
                        else
                        {
                            list.Add(new
                            {
                                OuId = o.Code,
                                OuName = o.Name,
                                EmployeeNo = employeeNo
                            });
                        }

                        //string[] sp = cp.Split('_');
                        //if (sp.Length > 1 && !string.IsNullOrEmpty(sp[1].Trim()))
                        //{
                        //    if (!cps.Any(q => sp[1].Equals(q.Code)))
                        //        error += string.Format("Dòng thứ {0} trên file, tách mã cơ cấu nhưng không tìm thấy trên hệ thống <br/> ({1})", i, cp);
                        //    else
                        //    {
                        //        list.Add(new
                        //        {
                        //            Company = company,
                        //            UnitGroup = unitGroup,
                        //            Unit = unit,
                        //            Department = department,
                        //            Team = team,
                        //            Group = group,
                        //            EmployeeNo = employeeNo
                        //        });
                        //    }
                        //}
                        //else
                        //{
                        //    error += string.Format("Dòng thứ {0} trên file không tìm thấy mã cơ cấu<br/> ({1})", i, cp);
                        //}
                    }
                }
                return list;
            }
            catch (Exception ex)
            {
                error = ex.ToString();
                return null;
            }
        }

        public static bool SaveImportHeadOfUnit(string userId, List<PostedImportHeadOfUnit> list, ref string error)
        {
            return DbHelper.SaveImportHeadOfUnit(userId, list, ref error);
        }

        public static bool SaveSetupEmployeePeriod(string userId, int periodId, string[] employees, ref string error)
        {
            return DbHelper.SaveSetupEmployeePeriod(userId, periodId, employees, ref error);
        }

        public static bool RemoveSetupEmployeePeriod(string userId, int periodId, string[] employees, ref string error)
        {
            return DbHelper.RemoveSetupEmployeePeriod(userId, periodId, employees, ref error);
        }

        public static bool SaveKpi(string userId, string kpiType, string kpiGroup, string kpiId, string kpiName, string kpiDesc, string kpiMethod, string kpiFormat, string kpiFormular, string ouId, string active, string mode, ref string error)
        {
            switch (mode)
            {
                case "ADD":
                    {
                        return CreateKpi(userId, kpiType, kpiGroup, kpiId, kpiName, kpiDesc, kpiMethod, kpiFormat, kpiFormular, ouId, active, ref error);
                    }
                case "EDIT":
                    {
                        return UpdateKpi(userId, kpiType, kpiGroup, kpiId, kpiName, kpiDesc, kpiMethod, kpiFormat, kpiFormular, ouId, active, ref error);
                    }
                default:
                    break;
            }
            error = "MODE IS EMPTY";
            return false;
        }

        public static bool SaveKot(List<PostedKotDetail> postedKotDetail, string userId, int? kotId, string kotName, int active, string ouId, string ouUpdate, ref string error) => DbHelper.SaveKot(postedKotDetail, userId, kotId, kotName, active, ouId, ouUpdate, ref error);

        public static bool CreateKpi(string userId, string kpiType, string kpiGroup, string kpiId, string kpiName, string kpiDesc, string kpiMethod, string kpiFormat, string kpiFormular, string ouId, string active, ref string error)
        {
            return DbHelper.CreateKpi(userId, kpiType, kpiGroup, kpiId, kpiName, kpiDesc, kpiMethod, kpiFormat, kpiFormular, ouId, active, ref error);
        }

        public static bool UpdateKpi(string userId, string kpiType, string kpiGroup, string kpiId, string kpiName, string kpiDesc, string kpiMethod, string kpiFormat, string kpiFormular, string ouId, string active, ref string error)
        {
            return DbHelper.UpdateKpi(userId, kpiType, kpiGroup, kpiId, kpiName, kpiDesc, kpiMethod, kpiFormat, kpiFormular, ouId, active, ref error);
        }

        public static bool SaveAssignKot(List<PostedAssignKot> postedList, string userId, int periodId, ref string error) => DbHelper.SaveAssignKot(postedList, userId, periodId, ref error);

        public static bool SaveEvaluation(List<PostedSaveEvaluation> postedList, int periodId, int isCompleted, string userId, int headerId, string note1, string note2, string note3, string note4, string note5, string note6, ref string error) => DbHelper.SaveEvaluation(postedList, periodId, isCompleted, userId, headerId, note1, note2, note3, note4, note5, note6, ref error);

        public static bool DoApproveEffect(string userId, int status, int periodId, int headerId, string note, List<PostedApproveEffect> postedList, string feedBack1, string feedBack2, string feedBack3, string feedBack4, string feedBack5, string feedBack6, ref string error) => DbHelper.DoApproveEffect(status, userId, periodId, headerId, note, postedList, feedBack1, feedBack2, feedBack3, feedBack4, feedBack5, feedBack6, ref error);

        public static DataTable GetKPI4ExImport(out string errMsg, int period_id = 0, string ouId = null) => DbHelper.GetKPI4ExImport(sys.current.UserInfo.UserDomain, period_id, ouId, out errMsg);

        public static DataTable GetOU4KPIExImport(out string errMsg, string pki_id, int period_id = 0) => DbHelper.GetOU4KPIExImport(sys.current.UserInfo.UserDomain, period_id, pki_id, out errMsg);

        public static DataTable GetDataKPI4ExImport(out string errMsg, out bool isLevel3, string pki_id, bool forChecking = false, string outId = null, int period_id = 0) => DbHelper.GetDataKPI4ExImport(sys.current.UserInfo.UserDomain, period_id, forChecking, pki_id, outId, out errMsg, out isLevel3);

        public static DataTable GetDataKPI4ExImportEffect(out string errMsg, string pki_id, bool forChecking = false, string outId = null, int period_id = 0) => DbHelper.GetDataKPI4ExImportEffect(sys.current.UserInfo.UserDomain, period_id, forChecking, pki_id, outId, out errMsg);

        public static DataTable GetAllRoles(string userId) => DbHelper.GetAllRoles(userId);

        public static bool AddUserToRole(string userId, int roleId, string addedUser, ref string error) => DbHelper.AddUserToRole(userId, roleId, addedUser, ref error);

        public static bool AddOuUser(string userId, int roleId, string addedUser, string ouId, int isDVKD, ref string error) => DbHelper.AddOuUser(userId, roleId, addedUser, ouId, isDVKD, ref error);

        public static bool RemoveOuUser(string userId, int roleId, string userRole, string ouId, ref string error) => DbHelper.RemoveOuUser(userId, roleId, userRole, ouId, ref error);

        public static bool RemoveRoleUser(string userId, int roleId, string userRole, ref string error) => DbHelper.RemoveRoleUser(userId, roleId, userRole, ref error);

        public static object ReadImportOuKpiResult(string userId, Stream fileStream, ref string error)
        {
            try
            {
                List<object> list = new List<object>();
                using (var workBook = new Workbook())
                {
                    workBook.LoadDocument(fileStream);
                    Worksheet sheet = workBook.Worksheets[0];
                    var range = sheet.GetDataRange();
                    int row = range.BottomRowIndex;
                    for (int i = 1; i <= row; i++)
                    {
                        var ouId = sheet.GetCellValue(0, i).TextValue;
                        ouId = string.IsNullOrEmpty(ouId) ? ouId : ouId.Trim();
                        var ouName = sheet.GetCellValue(1, i).TextValue;
                        var result = sheet.GetCellValue(2, i).NumericValue;
                        list.Add(new
                        {
                            OU_ID = ouId,
                            OU_NAME = ouName,
                            RESULT = result
                        });
                    }
                }
                return list;
            }
            catch (Exception ex)
            {
                error = ex.ToString();
                return null;
            }
        }

        public static bool SaveImportOuKpiResult(string userId, int periodId, List<PostedImportOuKpiResult> postedList, ref string error) => DbHelper.SaveImportOuKpiResult(userId, periodId, postedList, ref error);

        public static bool SaveImportOuProportion(string userId, int periodId, List<PostedImportOuProportion> postedList, ref string error) => DbHelper.SaveImportOuProportion(userId, periodId, postedList, ref error);

        public static DataTable ImportAndCheckKPIData(out string errMsg, string xml, string reason, int period_id = 0) => DbHelper.ImportAndCheckKPIData(sys.current.UserInfo.UserDomain, period_id, xml, reason, out errMsg);

        public static bool SaveImportPlan(string userId, int periodId, List<PostedImportPlan> postedList, int isImport3L, ref string error) => DbHelper.SaveImportPlan(userId, periodId, postedList, isImport3L, ref error);

        public static bool SaveImportEffect(string userId, int periodId, List<PostedImportEffect> postedList, ref string error) => DbHelper.SaveImportEffect(userId, periodId, postedList, ref error);

        public static DataTable GetDataProposeRanking(string userId, int periodId) => DbHelper.GetDataProposeRanking(userId, periodId);

        public static bool SaveUpdateRanking(string userId, List<PostedUpdateRanking> postedList, int periodId, ref string error) => DbHelper.SaveUpdateRanking(userId, postedList, periodId, ref error);

        public static DataTable GetJobTitleList(string userId) => DbHelper.GetJobTitleList(userId);

        public static bool SaveUpdatEmpPeriod(string userId, string  employeeNo, int periodId, int jobTitleId, string ouId, string impersonator, ref string error) => DbHelper.SaveUpdatEmpPeriod(userId, employeeNo, periodId, jobTitleId, ouId,impersonator ,ref error);
    }
}