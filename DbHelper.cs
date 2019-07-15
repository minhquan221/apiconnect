using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using NewPMS.Helpers;

namespace NewPMS.Models
{
    public static class DbHelper
    {
        public static DataTable GetUserInfo(string userName)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_USER_INFO").VarcharIN("P_USER_ID", userName)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetCompanyInfo(string userId, string ouId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_OU_INFO").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_OU_ID", ouId)
                .RefCursorOut("P_INFO").ExecuteDataTable();
        }

        public static DataTable GetPeriodList(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_PERIOD_LIST").VarcharIN("P_USER_ID", userId)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetModuleByUser(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_MODULE_BY_USER").VarcharIN("P_USER_ID", userId)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetRoleByUser(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_ROLE_BY_USER").VarcharIN("P_USER_ID", userId)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetActivePeriod(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_ACTIVE_PERIOD").VarcharIN("P_USER_ID", userId)
                .RefCursorOut("P_INFO").ExecuteDataTable();
        }

        public static DataTable GetPeriodInfo(string userId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_PERIOD_INFO").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .RefCursorOut("P_INFO").ExecuteDataTable();
        }

        public static DataTable GetCompanyStructByType(string userId, int? type)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_COMPANY_STRUCTURE_BY_TYPE").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_TYPE", type)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetCompanyStructByClasscification(string userId, string classId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_OU_BY_CLASSCIFICATION").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_CLASS_ID", classId)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetEvaluationByOuId(string userId, string ouId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_EVAL_OU_LIST_BY_OU_ID").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_OU_ID", ouId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetEvaluationOuList(string userId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_EVALUATION_OU_LIST").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetEmployeesByOuId(string userId, string ouId, int isAllChildren)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_EMPLOYEES_BY_OU_ID").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_OU_ID", ouId)
                .DecimalIN("P_ALL_CHILDREN", isAllChildren)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetEmployeesByUser(string userId, string employeeNo, int isAllChildren)
        {
            string sql = "PKG_NEW_PMS.GET_EMPLOYEES_BY_USER";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.IsNullable = false;
            P_USER_ID.Value = userId;

            OracleParameter P_EMPLOYEE_NO = new OracleParameter("P_EMPLOYEE_NO", OracleDbType.Varchar2);
            P_EMPLOYEE_NO.Direction = ParameterDirection.Input;
            P_EMPLOYEE_NO.IsNullable = false;
            P_EMPLOYEE_NO.Value = employeeNo;

            OracleParameter P_ALL_CHILDREN = new OracleParameter("P_ALL_CHILDREN", OracleDbType.Decimal);
            P_ALL_CHILDREN.Direction = ParameterDirection.Input;
            P_ALL_CHILDREN.IsNullable = false;
            P_ALL_CHILDREN.Value = isAllChildren;

            OracleParameter P_LIST = new OracleParameter("P_LIST", OracleDbType.RefCursor);
            P_LIST.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_EMPLOYEE_NO);
            cm.Parameters.Add(P_ALL_CHILDREN);
            cm.Parameters.Add(P_LIST);

            OracleDataAdapter ap = new OracleDataAdapter(cm);
            DataSet ds = new DataSet();
            ap.Fill(ds);
            cn.Close();

            return ds.Tables[0];
        }

        public static DataTable GetEvaluationUser(string userId, string ouId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_EVALUATION_USER").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_OU_ID", ouId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .RefCursorOut("P_INFO").ExecuteDataTable();
        }

        public static DataTable GetEvaluatedOuList(string userId, string ouId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_EVALUATED_OU_LIST").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_OU_ID", ouId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataSet GetEvaluatonInfo(string userId, string ouId, string evaluationOuId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_EVALUATION_INFO").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_OU_ID", ouId)
                .VarcharIN("P_EVALUATION_OU_ID", evaluationOuId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .RefCursorOut("P_HEADER", "P_RESULT", "P_COMMENT")
                .ExecuteDataSet();
        }

        public static DataTable GetAllEmployeeForSync(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_ALL_EMPLOYEES_FOR_SYNC")
                .VarcharIN("P_USER_ID", userId)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetSLAOu(string userId, string ouId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_SLA_OU").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .VarcharIN("P_OU_ID", ouId)
                .RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataSet GetInternalSatisfactionResult(string userId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_REPORT.INTERNAL_SATISFACTION_RESULT").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .RefCursorOut("P_HEADER", "P_LIST")
                .ExecuteDataSet();
        }

        public static DataSet GetInternalSatisfactionReport(string userId, string ouId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_REPORT.INTERNAL_SATISFACTION_REPORT").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .VarcharIN("P_OU_ID", ouId)
                .RefCursorOut("P_HEADER", "P_LIST")
                .ExecuteDataSet();
        }

        public static DataTable GetEmployeeType(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_EMPLOYEE_TYPE").VarcharIN("P_USER_ID", userId)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable();
        }

        public static DataTable GetContractType(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_CONTRACT_TYPE").VarcharIN("P_USER_ID", userId)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable();
        }

        public static DataTable GetEmployeePeriodList(string userId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_EMPLOYEE_PERIOD_LIST").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable();
        }

        public static DataTable GetDataForSetupEmployeePeriod(string userId, string comparison, DateTime? startDate, int? status, string employeeType, string contractType)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_DATA_FOR_SETUP_EMP_PERIOD").VarcharIN("P_USER_ID", userId)
               .DateIN("P_START_DATE", startDate)
               .VarcharIN("P_START_DATE_COMPARISON", comparison)
               .VarcharIN("P_EMPLOYEE_TYPE", employeeType)
               .VarcharIN("P_CONTRACT_TYPE", contractType)
               .DecimalIN("P_STATUS", status)
               .RefCursorOut("P_LIST")
               .ExecuteDataTable();
        }

        public static DataTable GetKpiType(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_KPI_TYPE").VarcharIN("P_USER_ID", userId)
               .RefCursorOut("P_LIST")
               .ExecuteDataTable();
        }

        public static DataTable GetKpiGroup(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_KPI_GROUP").VarcharIN("P_USER_ID", userId)
               .RefCursorOut("P_LIST")
               .ExecuteDataTable();
        }

        public static DataTable GetKpiFormular(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_KPI_FORMULAR").VarcharIN("P_USER_ID", userId)
               .RefCursorOut("P_LIST")
               .ExecuteDataTable();
        }

        public static DataTable GetKpiFormat(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_KPI_FORMAT").VarcharIN("P_USER_ID", userId)
               .RefCursorOut("P_LIST")
               .ExecuteDataTable();
        }

        public static DataTable GetUserOu(string userId, string employeeNo, int? isAllChilren, int? roleId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_OU_LIST_BY_ROLE").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_EMPLOYEE_NO", employeeNo)
                .DecimalIN("P_IS_ALL_CHILREN", isAllChilren)
                .DecimalIN("P_ROLE_ID", roleId)
               .RefCursorOut("P_LIST")
               .ExecuteDataTable();
        }

        public static DataTable GetKotByRole(string userId, string employeeNo, int isAllChilren = 0)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_KOT_BY_ROLE").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_EMPLOYEE_NO", employeeNo)
                .DecimalIN("P_IS_ALL_CHILREN", isAllChilren)
               .RefCursorOut("P_LIST")
               .ExecuteDataTable();
        }

        public static DataTable SearchKpi(string userId, int? status, string kpiId, string ouId, string kpiType, string groupId, string kpiName)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.SEARCH_KPI_LIST").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_STATUS", status)
                .VarcharIN("P_KPI_ID", kpiId)
                .VarcharIN("P_KPI_NAME", kpiName)
                .VarcharIN("P_OU_ID", ouId)
                .VarcharIN("P_TYPE_ID", kpiType)
                .VarcharIN("P_GROUP_ID", groupId)
               .RefCursorOut("P_LIST")
               .ExecuteDataTable();
        }

        public static DataTable GetKpi(string userId, string Id)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_KPI_DETAIL").VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_ID", Id)
               .RefCursorOut("P_INFO")
               .ExecuteDataTable();
        }

        public static DataSet GetKotDetail(string userId, int kotId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_KOT_DETAIL").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_KOT_ID", kotId)
                .RefCursorOut("P_KOT", "P_KOT_DETAIL")
               .ExecuteDataSet();
        }

        public static DataTable SearchKotList(string userId, int? status, string kotId, string kotName, string ouId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.SEARCH_KOT_LIST").VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_STATUS", status)
                .VarcharIN("P_KOT_ID", kotId)
                .VarcharIN("P_KOT_NAME", kotName)
                .VarcharIN("P_OU_ID", ouId)
               .RefCursorOut("P_LIST")
               .ExecuteDataTable();
        }

        public static DataTable SearchKotEmployees(string userId, int? status, string ouId, int periodId)
        {
            string sql = "PKG_NEW_PMS.GET_KOT_EMPLOYEES";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.IsNullable = false;
            P_USER_ID.Value = userId;

            OracleParameter P_STATUS = new OracleParameter("P_STATUS ", OracleDbType.Decimal);
            P_STATUS.Direction = ParameterDirection.Input;
            P_STATUS.IsNullable = false;
            P_STATUS.Value = status;

            OracleParameter P_PERIOD_ID = new OracleParameter("P_PERIOD_ID ", OracleDbType.Decimal);
            P_PERIOD_ID.Direction = ParameterDirection.Input;
            P_PERIOD_ID.IsNullable = false;
            P_PERIOD_ID.Value = periodId;

            OracleParameter P_OU_ID = new OracleParameter("P_OU_ID ", OracleDbType.Varchar2);
            P_OU_ID.Direction = ParameterDirection.Input;
            P_OU_ID.IsNullable = false;
            P_OU_ID.Value = ouId;

            OracleParameter P_LIST = new OracleParameter("P_LIST", OracleDbType.RefCursor);
            P_LIST.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_STATUS);
            cm.Parameters.Add(P_PERIOD_ID);
            cm.Parameters.Add(P_OU_ID);
            cm.Parameters.Add(P_LIST);

            OracleDataAdapter ap = new OracleDataAdapter(cm);
            DataSet ds = new DataSet();
            ap.Fill(ds);
            cn.Close();

            return ds.Tables[0];
        }

        public static DataTable GetEmployeePeriod(string employeeNo, int periodId)
        {
            string sql = "PKG_NEW_PMS.GET_EMPLOYEE_PERIOD";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_EMPLOYEE_NO = new OracleParameter("P_EMPLOYEE_NO", OracleDbType.Varchar2);
            P_EMPLOYEE_NO.Direction = ParameterDirection.Input;
            P_EMPLOYEE_NO.IsNullable = false;
            P_EMPLOYEE_NO.Value = employeeNo;

            OracleParameter P_PERIOD_ID = new OracleParameter("P_PERIOD_ID", OracleDbType.Decimal);
            P_PERIOD_ID.Direction = ParameterDirection.Input;
            P_PERIOD_ID.IsNullable = false;
            P_PERIOD_ID.Value = periodId;

            OracleParameter P_LIST = new OracleParameter("P_INFO", OracleDbType.RefCursor);
            P_LIST.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_EMPLOYEE_NO);
            cm.Parameters.Add(P_PERIOD_ID);
            cm.Parameters.Add(P_LIST);

            OracleDataAdapter ap = new OracleDataAdapter(cm);
            DataSet ds = new DataSet();
            ap.Fill(ds);
            cn.Close();

            return ds.Tables[0];
        }

        public static DataSet GetKotEmpForCreateEvaluation(string employeeNo, int periodId)
        {
            string sql = "PKG_NEW_PMS.GET_KOT_EMP_EVALUATION";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_EMPLOYEE_NO = new OracleParameter("P_EMPLOYEE_NO", OracleDbType.Varchar2);
            P_EMPLOYEE_NO.Direction = ParameterDirection.Input;
            P_EMPLOYEE_NO.IsNullable = false;
            P_EMPLOYEE_NO.Value = employeeNo;

            OracleParameter P_PERIOD_ID = new OracleParameter("P_PERIOD_ID", OracleDbType.Decimal);
            P_PERIOD_ID.Direction = ParameterDirection.Input;
            P_PERIOD_ID.IsNullable = false;
            P_PERIOD_ID.Value = periodId;

            OracleParameter P_HEADER = new OracleParameter("P_HEADER", OracleDbType.RefCursor);
            P_HEADER.Direction = ParameterDirection.Output;

            OracleParameter P_DETAIL = new OracleParameter("P_DETAIL", OracleDbType.RefCursor);
            P_DETAIL.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_EMPLOYEE_NO);
            cm.Parameters.Add(P_PERIOD_ID);
            cm.Parameters.Add(P_HEADER);
            cm.Parameters.Add(P_DETAIL);

            OracleDataAdapter ap = new OracleDataAdapter(cm);
            DataSet ds = new DataSet();
            ap.Fill(ds);
            cn.Close();

            return ds;
        }

        public static DataTable GetEmployeesWaitApprove(string userId, int periodId, int? status)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_EMPS_WAIT_APPROVE")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .DecimalIN("P_STATUS", status)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable();
        }

        public static DataSet GetPersonalEvaluatioinReport(string userId, string employeeNo, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_REPORT.GET_PERSONAL_EVALUATION_RP")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .VarcharIN("P_EMPLOYEE_NO", employeeNo)
                .RefCursorOut("P_HEADER")
                .RefCursorOut("P_DETAIL")
                .RefCursorOut("P_RESULT")
                .ExecuteDataSet();
        }

        public static DataTable GetAllRoles(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_ALL_ROLES").VarcharIN("P_USER_ID", userId).RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static DataTable GetModulesByRole(string userId, int roleId) { return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_MODULE_BY_ROLE").VarcharIN("P_USER_ID", userId).DecimalIN("P_ROLE_ID", roleId).RefCursorOut("P_LIST").ExecuteDataTable(); }

        public static DataTable GetUsersByRole(string userId, int roleId) { return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_USER_BY_ROLE").VarcharIN("P_USER_ID", userId).DecimalIN("P_ROLE_ID", roleId).RefCursorOut("P_LIST").ExecuteDataTable(); }

        public static bool CreateEmployee(string userId, List<PostedEmployee> list, DateTime from, DateTime to, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_SYNC_PMS.CREATE_EMPLOYEE")
                .VarcharIN("P_USER_ID", userId)
                .DateIN("P_FROM", from)
                .DateIN("P_TO", to)
                .XmlIN("P_XML", XmlUtils.Serialize(new EmployeeSender() { PostedList = list }))
                .VarcharOUT("P_ERR_MSG", 2000).ExecuteGetValue("P_ERR_MSG");
            var strPara = (Oracle.ManagedDataAccess.Types.OracleString)ret;

            error = !strPara.IsNull ? strPara.Value : string.Empty;
            if ("null".Equals(error)) error = string.Empty;
            return string.IsNullOrEmpty(error);
        }

        private static bool HandlerResult(List<object> list, ref string error)
        {
            var kq = 0;
            bool b = int.TryParse(list[0].ToString(), out kq);
            if (b)
            {
                if (kq != -1)
                {
                    var ep = (Oracle.ManagedDataAccess.Types.OracleString)list[1];
                    error = !ep.IsNull ? ep.Value : string.Empty;
                    return false;
                }
                return true;
            }
            return false;
        }

        public static bool DoCreatePeriod(int? id, string name, int? year, int active, string userId, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_SYNC_PMS.CREATE_PERIOD")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_ID", id)
                .VarcharIN("P_NAME", name)
                .DecimalIN("P_ACTIVE", active)
                .DecimalIN("P_YEAR", year)
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000).
                ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");

            return HandlerResult(ret, ref error);
        }

        public static bool CreateCompanyStructure(string userId, int level, List<PostedCompany> postedList, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_SYNC_PMS.CREATE_COMPANY_STRUCTURE")
               .VarcharIN("P_USER_ID", userId)
               .DecimalIN("P_LEVEL", level)
               .XmlIN("P_XML", XmlUtils.Serialize(new CompanySender() { PostedList = postedList }))
               .VarcharOUT("P_ERR_MSG", 2000).ExecuteGetValue("P_ERR_MSG");
            var strPara = (Oracle.ManagedDataAccess.Types.OracleString)ret;
            error = strPara.IsNull ? string.Empty : strPara.Value;
            if ("null".Equals(error))
                error = string.Empty;
            return string.IsNullOrEmpty(error);
        }

        public static bool CreateJobTitle(string userId, int level, List<PostedJobTitle> postedList, ref string error)
        {
            string sql = "PKG_SYNC_PMS.CREATE_JOB_TITLE";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_XML = new OracleParameter("P_XML", OracleDbType.XmlType);
            P_XML.Direction = ParameterDirection.Input;
            P_XML.Value = XmlUtils.Serialize(new JobTitleSender() { PostedList = postedList });

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_XML);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
            if ("null".Equals(error))
                error = string.Empty;
            return string.IsNullOrEmpty(error) || "null".Equals(error);
        }

        public static bool SaveDeclareEvaluationOu(string userId, List<PostedDeclareEvalOu> list, int? periodId, int hasSLA, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.SAVE_EVALUATION_OU_LIST")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .DecimalIN("P_HAS_SLA", hasSLA)
                .XmlIN("P_XML", XmlUtils.Serialize(new DeclareEvalOuSender() { PostedList = list }))
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            int kq = 0;
            if (int.TryParse(ret[0].ToString(), out kq))
            {
                if (kq != -1)
                {
                    var ep = (Oracle.ManagedDataAccess.Types.OracleString)ret[1];
                    error = ep.IsNull ? string.Empty : ep.Value;
                    return false;
                }
                return true;
            }
            return false;
        }

        public static bool SaveEvaluationUser(string userId, string ouId, string employeeNo, int periodId, ref string error)
        {
            string sql = "PKG_NEW_PMS.SAVE_EVALUATION_USER";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_OU_ID = new OracleParameter("P_OU_ID", OracleDbType.Varchar2);
            P_OU_ID.Direction = ParameterDirection.Input;
            P_OU_ID.Value = ouId;

            OracleParameter P_EMPLOYEE_NO = new OracleParameter("P_EMPLOYEE_NO", OracleDbType.Varchar2);
            P_EMPLOYEE_NO.Direction = ParameterDirection.Input;
            P_EMPLOYEE_NO.Value = employeeNo;

            OracleParameter P_PERIOD_ID = new OracleParameter("P_PERIOD_ID", OracleDbType.Decimal);
            P_PERIOD_ID.Direction = ParameterDirection.Input;
            P_PERIOD_ID.Value = periodId;

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_OU_ID);
            cm.Parameters.Add(P_EMPLOYEE_NO);
            cm.Parameters.Add(P_PERIOD_ID);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool SaveHeadOfUnit(string userId, string ouId, string employeeNo, ref string error)
        {
            string sql = "PKG_NEW_PMS.SAVE_HEAD_OF_UNIT";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_OU_ID = new OracleParameter("P_OU_ID", OracleDbType.Varchar2);
            P_OU_ID.Direction = ParameterDirection.Input;
            P_OU_ID.Value = ouId;

            OracleParameter P_EMPLOYEE_NO = new OracleParameter("P_EMPLOYEE_NO", OracleDbType.Varchar2);
            P_EMPLOYEE_NO.Direction = ParameterDirection.Input;
            P_EMPLOYEE_NO.Value = employeeNo;

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_OU_ID);
            cm.Parameters.Add(P_EMPLOYEE_NO);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool SaveInternalSatisfaction(string userId, string ouId, string evaluationOuId, int periodId, int completed, List<PostedInternalSatisfactionQuestion> postedList, List<PostedInternalSatisfactionSubQuestion> subPostedList, ref string error)
        {
            string sql = "PKG_NEW_PMS.SAVE_EVALUATION";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_OU_ID = new OracleParameter("P_OU_ID", OracleDbType.Varchar2);
            P_OU_ID.Direction = ParameterDirection.Input;
            P_OU_ID.Value = ouId;

            OracleParameter P_EVALUATION_OU_ID = new OracleParameter("P_EVALUATION_OU_ID", OracleDbType.Varchar2);
            P_EVALUATION_OU_ID.Direction = ParameterDirection.Input;
            P_EVALUATION_OU_ID.Value = evaluationOuId;

            OracleParameter P_PERIOD_ID = new OracleParameter("P_PERIOD_ID", OracleDbType.Decimal);
            P_PERIOD_ID.Direction = ParameterDirection.Input;
            P_PERIOD_ID.Value = periodId;

            OracleParameter P_COMPLETED = new OracleParameter("P_COMPLETED", OracleDbType.Decimal);
            P_COMPLETED.Direction = ParameterDirection.Input;
            P_COMPLETED.Value = completed;

            OracleParameter P_INFO = new OracleParameter("P_INFO", OracleDbType.XmlType);
            P_INFO.Direction = ParameterDirection.Input;
            P_INFO.Value = XmlUtils.Serialize(new InfoSender() { PostedList = postedList });

            OracleParameter P_SUB_INFO = new OracleParameter("P_SUB_INFO", OracleDbType.XmlType);
            P_SUB_INFO.Direction = ParameterDirection.Input;
            P_SUB_INFO.Value = XmlUtils.Serialize(new SubInfoSender() { PostedList = subPostedList });

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_OU_ID);
            cm.Parameters.Add(P_EVALUATION_OU_ID);
            cm.Parameters.Add(P_PERIOD_ID);
            cm.Parameters.Add(P_COMPLETED);
            cm.Parameters.Add(P_INFO);
            cm.Parameters.Add(P_SUB_INFO);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool SaveImportHeadOfUnit(string userId, List<PostedImportHeadOfUnit> list, ref string error)
        {
            string sql = "PKG_NEW_PMS.SAVE_IMPORT_HEAD_OF_UNIT";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_INFO = new OracleParameter("P_INFO", OracleDbType.XmlType);
            P_INFO.Direction = ParameterDirection.Input;
            P_INFO.Value = XmlUtils.Serialize(new PostedImportHeadOfUnitSender() { PostedList = list });

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_INFO);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool SaveSetupEmployeePeriod(string userId, int periodId, string[] employees, ref string error)
        {
            string sql = "PKG_NEW_PMS.UPDATE_EMPLOYEE_PERIOD";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_PERIOD_ID = new OracleParameter("P_PERIOD_ID", OracleDbType.Decimal);
            P_PERIOD_ID.Direction = ParameterDirection.Input;
            P_PERIOD_ID.Value = periodId;

            OracleParameter P_EMPLOYEES = new OracleParameter("P_EMPLOYEES", OracleDbType.XmlType);
            P_EMPLOYEES.Direction = ParameterDirection.Input;
            P_EMPLOYEES.Value = XmlUtils.Serialize(new SetupEmployeesSender() { PostedList = employees.Select(q => new PostedEmployeeCode() { EMPLOYEE_NO = q }).ToList() });

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_PERIOD_ID);
            cm.Parameters.Add(P_EMPLOYEES);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool SaveKot(List<PostedKotDetail> postedKotDetail, string userId, int? kotId, string kotName, int active, string ouId, string ouUpdate, ref string error)
        {
            string sql = "PKG_NEW_PMS.UPDATE_KOT_DETAIL";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_KOT_ID = new OracleParameter("P_KOT_ID", OracleDbType.Decimal);
            P_KOT_ID.Direction = ParameterDirection.Input;
            P_KOT_ID.Value = kotId;

            OracleParameter P_OU_ID = new OracleParameter("P_OU_ID", OracleDbType.Varchar2);
            P_OU_ID.Direction = ParameterDirection.Input;
            P_OU_ID.Value = ouId;

            OracleParameter P_KOT_NAME = new OracleParameter("P_KOT_NAME", OracleDbType.Varchar2);
            P_KOT_NAME.Direction = ParameterDirection.Input;
            P_KOT_NAME.Value = kotName;

            OracleParameter P_USED = new OracleParameter("P_USED", OracleDbType.Decimal);
            P_USED.Direction = ParameterDirection.Input;
            P_USED.Value = active;

            OracleParameter P_OU_UPDATE = new OracleParameter("P_OU_UPDATE", OracleDbType.Varchar2);
            P_OU_UPDATE.Direction = ParameterDirection.Input;
            P_OU_UPDATE.Value = ouUpdate;

            OracleParameter P_XML = new OracleParameter("P_XML", OracleDbType.XmlType);
            P_XML.Direction = ParameterDirection.Input;
            P_XML.Value = XmlUtils.Serialize(new KotDetailSender() { PostedList = postedKotDetail });

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_XML);
            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_KOT_ID);
            cm.Parameters.Add(P_KOT_NAME);
            cm.Parameters.Add(P_OU_ID);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_USED);
            cm.Parameters.Add(P_OU_UPDATE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool RemoveSetupEmployeePeriod(string userId, int periodId, string[] employees, ref string error)
        {
            string sql = "PKG_NEW_PMS.REMOVE_EMPLOYEE_PERIOD";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_PERIOD_ID = new OracleParameter("P_PERIOD_ID", OracleDbType.Decimal);
            P_PERIOD_ID.Direction = ParameterDirection.Input;
            P_PERIOD_ID.Value = periodId;

            OracleParameter P_EMPLOYEES = new OracleParameter("P_EMPLOYEES", OracleDbType.XmlType);
            P_EMPLOYEES.Direction = ParameterDirection.Input;
            P_EMPLOYEES.Value = XmlUtils.Serialize(new SetupEmployeesSender() { PostedList = employees.Select(q => new PostedEmployeeCode() { EMPLOYEE_NO = q }).ToList() });

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_PERIOD_ID);
            cm.Parameters.Add(P_EMPLOYEES);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool CreateKpi(string userId, string kpiType, string kpiGroup, string kpiId, string kpiName, string kpiDesc, string kpiMethod, string kpiFormat, string kpiFormular, string ouId, string active, ref string error)
        {
            string sql = "PKG_NEW_PMS.CREATE_KPI";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_ID = new OracleParameter("P_ID", OracleDbType.Varchar2);
            P_ID.Direction = ParameterDirection.Input;
            P_ID.Value = kpiId;

            OracleParameter P_NAME = new OracleParameter("P_NAME", OracleDbType.Varchar2);
            P_NAME.Direction = ParameterDirection.Input;
            P_NAME.Value = kpiName;

            OracleParameter P_GROUP_ID = new OracleParameter("P_GROUP_ID", OracleDbType.Varchar2);
            P_GROUP_ID.Direction = ParameterDirection.Input;
            P_GROUP_ID.Value = kpiGroup;

            OracleParameter P_TYPE_ID = new OracleParameter("P_TYPE_ID", OracleDbType.Varchar2);
            P_TYPE_ID.Direction = ParameterDirection.Input;
            P_TYPE_ID.Value = kpiType;

            OracleParameter P_DESCRIPTION = new OracleParameter("P_DESCRIPTION", OracleDbType.Varchar2);
            P_DESCRIPTION.Direction = ParameterDirection.Input;
            P_DESCRIPTION.Value = kpiDesc;

            OracleParameter P_METHOD = new OracleParameter("P_METHOD", OracleDbType.Varchar2);
            P_METHOD.Direction = ParameterDirection.Input;
            P_METHOD.Value = kpiMethod;

            OracleParameter P_FORMULAR = new OracleParameter("P_FORMULAR", OracleDbType.Varchar2);
            P_FORMULAR.Direction = ParameterDirection.Input;
            P_FORMULAR.Value = kpiFormular;

            OracleParameter P_FORMAT = new OracleParameter("P_FORMAT", OracleDbType.Varchar2);
            P_FORMAT.Direction = ParameterDirection.Input;
            P_FORMAT.Value = kpiFormat;

            OracleParameter P_USED = new OracleParameter("P_USED", OracleDbType.Varchar2);
            P_USED.Direction = ParameterDirection.Input;
            P_USED.Value = active;

            OracleParameter P_OU_ID = new OracleParameter("P_OU_ID", OracleDbType.Varchar2);
            P_OU_ID.Direction = ParameterDirection.Input;
            P_OU_ID.Value = ouId;

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_ID);
            cm.Parameters.Add(P_NAME);
            cm.Parameters.Add(P_GROUP_ID);
            cm.Parameters.Add(P_TYPE_ID);
            cm.Parameters.Add(P_DESCRIPTION);
            cm.Parameters.Add(P_METHOD);
            cm.Parameters.Add(P_FORMULAR);
            cm.Parameters.Add(P_FORMAT);
            cm.Parameters.Add(P_USED);
            cm.Parameters.Add(P_OU_ID);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool UpdateKpi(string userId, string kpiType, string kpiGroup, string kpiId, string kpiName, string kpiDesc, string kpiMethod, string kpiFormat, string kpiFormular, string ouId, string active, ref string error)
        {
            string sql = "PKG_NEW_PMS.UPDATE_KPI";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_ID = new OracleParameter("P_ID", OracleDbType.Varchar2);
            P_ID.Direction = ParameterDirection.Input;
            P_ID.Value = kpiId;

            OracleParameter P_NAME = new OracleParameter("P_NAME", OracleDbType.Varchar2);
            P_NAME.Direction = ParameterDirection.Input;
            P_NAME.Value = kpiName;

            OracleParameter P_GROUP_ID = new OracleParameter("P_GROUP_ID", OracleDbType.Varchar2);
            P_GROUP_ID.Direction = ParameterDirection.Input;
            P_GROUP_ID.Value = kpiGroup;

            OracleParameter P_TYPE_ID = new OracleParameter("P_TYPE_ID", OracleDbType.Varchar2);
            P_TYPE_ID.Direction = ParameterDirection.Input;
            P_TYPE_ID.Value = kpiType;

            OracleParameter P_DESCRIPTION = new OracleParameter("P_DESCRIPTION", OracleDbType.Varchar2);
            P_DESCRIPTION.Direction = ParameterDirection.Input;
            P_DESCRIPTION.Value = kpiDesc;

            OracleParameter P_METHOD = new OracleParameter("P_METHOD", OracleDbType.Varchar2);
            P_METHOD.Direction = ParameterDirection.Input;
            P_METHOD.Value = kpiMethod;

            OracleParameter P_FORMULAR = new OracleParameter("P_FORMULAR", OracleDbType.Varchar2);
            P_FORMULAR.Direction = ParameterDirection.Input;
            P_FORMULAR.Value = kpiFormular;

            OracleParameter P_FORMAT = new OracleParameter("P_FORMAT", OracleDbType.Varchar2);
            P_FORMAT.Direction = ParameterDirection.Input;
            P_FORMAT.Value = kpiFormat;

            OracleParameter P_USED = new OracleParameter("P_USED", OracleDbType.Varchar2);
            P_USED.Direction = ParameterDirection.Input;
            P_USED.Value = active;

            OracleParameter P_OU_ID = new OracleParameter("P_OU_ID", OracleDbType.Varchar2);
            P_OU_ID.Direction = ParameterDirection.Input;
            P_OU_ID.Value = ouId;

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_ID);
            cm.Parameters.Add(P_NAME);
            cm.Parameters.Add(P_GROUP_ID);
            cm.Parameters.Add(P_TYPE_ID);
            cm.Parameters.Add(P_DESCRIPTION);
            cm.Parameters.Add(P_METHOD);
            cm.Parameters.Add(P_FORMULAR);
            cm.Parameters.Add(P_FORMAT);
            cm.Parameters.Add(P_USED);
            cm.Parameters.Add(P_OU_ID);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool SaveAssignKot(List<PostedAssignKot> postedList, string userId, int periodId, ref string error)
        {
            string sql = "PKG_NEW_PMS.SAVE_ASSIGN_KOT";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;

            OracleParameter P_PERIOD_ID = new OracleParameter("P_PERIOD_ID", OracleDbType.Decimal);
            P_PERIOD_ID.Direction = ParameterDirection.Input;
            P_PERIOD_ID.Value = periodId;

            OracleParameter P_XML = new OracleParameter("P_XML", OracleDbType.XmlType);
            P_XML.Direction = ParameterDirection.Input;
            P_XML.Value = XmlUtils.Serialize(new AssignKotSender() { PostedList = postedList });

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_PERIOD_ID);
            cm.Parameters.Add(P_XML);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }

        public static bool SaveEvaluation(List<PostedSaveEvaluation> postedList, int periodId, int isCompleted, string userId, int headerId, string note1, string note2, string note3, string note4, string note5, string note6, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.SAVE_KPI_EVALUATION")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .DecimalIN("P_IS_COMPLETED", isCompleted)
                .DecimalIN("P_HEADER_ID", headerId)
                .XmlIN("P_XML", XmlUtils.Serialize(new EvaluationSender() { PostedList = postedList }))
                .VarcharIN("P_NOTE_1", note1)
                .VarcharIN("P_NOTE_2", note2)
                .VarcharIN("P_NOTE_3", note3)
                .VarcharIN("P_NOTE_4", note4)
                .VarcharIN("P_NOTE_5", note5)
                .VarcharIN("P_NOTE_6", note6)
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");

            return HandlerResult(ret, ref error);
        }

        public static bool DoApproveEffect(int status, string userId, int periodId, int headerId, string note, List<PostedApproveEffect> postedList, string feedBack1, string feedBack2, string feedBack3, string feedBack4, string feedBack5, string feedBack6, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.SAVE_APPROVE_EFFECT")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_STATUS_ID", status)
                .DecimalIN("P_PERIOD_ID", periodId)
                .DecimalIN("P_HEADER_ID", headerId)
                .VarcharIN("P_NOTE", note)
                .XmlIN("P_XML", XmlUtils.Serialize(new ApproveEffectSender() { PostedList = postedList }))
                .VarcharIN("P_FEED_BACK_1", feedBack1)
                .VarcharIN("P_FEED_BACK_2", feedBack2)
                .VarcharIN("P_FEED_BACK_3", feedBack3)
                .VarcharIN("P_FEED_BACK_4", feedBack4)
                .VarcharIN("P_FEED_BACK_5", feedBack5)
                .VarcharIN("P_FEED_BACK_6", feedBack6)
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }

        public static bool AddUserToRole(string userId, int roleId, string addedUser, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.ADD_USER_TO_ROLE")
                .VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_ADDED_USER", addedUser)
                .DecimalIN("P_ROLE_ID", roleId)
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }
        
        public static bool AddOuUser(string userId, int roleId, string addedUser, string ouId, int isDVKD, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.ADD_OU_USER")
                .VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_ADDED_USER", addedUser)
                .DecimalIN("P_ROLE_ID", roleId)
                .VarcharIN("P_OU_ID", ouId)
                .DecimalIN("P_IS_DVKD", isDVKD)
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }

        public static bool RemoveOuUser(string userId, int roleId, string userRole, string ouId, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.REMOVE_OU_USER")
                .VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_USER_ROLE", userRole)
                .DecimalIN("P_ROLE_ID", roleId)
                .VarcharIN("P_OU_ID", ouId)
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }

        public static bool RemoveRoleUser(string userId, int roleId, string userRole, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.REMOVE_ROLE_USER")
                .VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_USER_ROLE", userRole)
                .DecimalIN("P_ROLE_ID", roleId)
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }
        #region Import KPI KH / TH
        public static DataTable GetKPI4ExImport(string userId, int period_id, string ouId, out string errMsg)
        {
            List<object> ret = null;
            errMsg = "ROLE_GRANTED";
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_T.GET_KPI_LIST_FOR_IMPORT")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", period_id)
                .VarcharIN("P_OU_ID", ouId)
                .VarcharOUT("P_ERR_MSG", 100)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable(new string[] { "P_ERR_MSG" }, out ret);
            errMsg = ret[0].ToString();
            return dt;
        }
        public static DataTable GetOU4KPIExImport(string userId, int period_id, string pki_id, out string errMsg)
        {
            List<object> ret = null;
            errMsg = "ROLE_GRANTED";
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_T.GET_OU_LIST_FOR_IMPORT")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", period_id)
                .VarcharIN("P_KPI_ID", pki_id)
                .VarcharOUT("P_ERR_MSG", 100)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable(new string[] { "P_ERR_MSG" }, out ret);
            errMsg = ret[0].ToString();
            return dt;
        }
        public static DataTable GetDataKPI4ExImport(string userId, int period_id, bool forChecking, string pki_id, string ouId, out string errMsg, out bool isLevel3)
        {
            List<object> ret = null;
            errMsg = "ROLE_GRANTED";
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_T.GET_DATA_FOR_IMPORT_PLAN")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", period_id)
                .DecimalIN("P_CHECK", forChecking ? 1 : 0)
                .VarcharIN("P_KPI_ID", pki_id)
                .VarcharIN("P_OU_ID", ouId)
                .VarcharOUT("P_ERR_MSG", 100)
                .DecimalOUT("P_IS_3_LEVEL")
                .RefCursorOut("P_LIST")
                .ExecuteDataTable(new string[] { "P_ERR_MSG", "P_IS_3_LEVEL" }, out ret);
            errMsg = ret[0].ToString();
            int iLevel3 = 0;
            isLevel3 = false;
            String r1 = ret[1].ToString();
            if (!int.TryParse(r1, out iLevel3))
            {
                errMsg = "ROLE_GRANTED_BUT_IS_LEVEL_3_" + r1;
            }
            else
            {
                isLevel3 = iLevel3 == 1;
            }
            return dt;
        }
        public static DataTable GetDataKPI4ExImportEffect(string userId, int period_id, bool forChecking, string pki_id, string ouId, out string errMsg)
        {
            List<object> ret = null;
            errMsg = "ROLE_GRANTED";
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_T.GET_DATA_FOR_IMPORT_EFFECT")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", period_id)
                .DecimalIN("P_CHECK", forChecking ? 1 : 0)
                .VarcharIN("P_KPI_ID", pki_id)
                .VarcharIN("P_OU_ID", ouId)
                .VarcharOUT("P_ERR_MSG", 100)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable(new string[] { "P_ERR_MSG" }, out ret);
            errMsg = ret[0].ToString();
            return dt;
        }
        public static DataTable GetOneKPI4ExImport(string userId, int period_id, string pki_id, out string errMsg) {
            errMsg = "ROLE_GRANTED";
            List<object> ret = null;
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_T.GET_DATA_FOR_IMPORT_PLAN")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", period_id)
                .VarcharIN("P_KPI_ID", pki_id)
                .VarcharOUT("P_ERR_MSG", 100)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable(new string[] { "P_ERR_MSG" }, out ret);
            errMsg = ret[0].ToString();
            return dt;
        }

        public static DataTable ImportAndCheckKPIData(string userId, int period_id, string xml, string reason, out string errMsg)
        {
            errMsg = "ROLE_GRANTED";
            List<object> ret = null;
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_T.PKI_DATA_IMPORT_BUFFER")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", period_id)
                .XmlIN("P_XML", xml)
                .VarcharIN("P_REASON_IMPORT", reason)
                .VarcharOUT("P_ERR_MSG", 100)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable(new string[] { "P_ERR_MSG" }, out ret);
            errMsg = ret[0].ToString();
            return dt;
        }

        public static List<T> ReadFromDataTable<T>(DataTable dt, IteratorOption options, Func<DataRow, int, IteratorOption, T> func) {
            List<T> ret = new List<T>();
            var rows = dt?.Rows;
            int cnt = rows != null ? rows.Count : 0;
            if (cnt > 0)
            {
                for (int i = 0; i < cnt; i++) {
                    T it = func(rows[i], i, options);
                    if (options.IsError)
                    {
                        if (options.StopWhenError) {
                            return ret;
                        }
                    }
                    else {
                        ret.Add(it);
                    }
                }
            }
            else {
                options.AddError("DataTable Empty!");
            }
            return ret;
        }

        public static bool WalkThroughDataTable(DataTable dt, IteratorOption options, Func<DataRow, int, IteratorOption, bool> walk_func)
        {
            var rows = dt?.Rows;
            int cnt = rows != null ? rows.Count : 0;
            if (cnt > 0)
            {
                for (int i = 0; i < cnt; i++)
                {
                    bool it = walk_func(rows[i], i, options);
                    if (options.IsError)
                    {
                        if (options.StopWhenError)
                        {
                            return false;
                        }
                    }
                }
            }
            else
            {
                options.AddError("DataTable Empty!");
                return false;
            }
            return true;
        }
        //GET_KPI_FOR_IMPORT
        public static bool SaveImportPlan(string userId, int periodId, List<PostedImportPlan> postedList, int isImport3L, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.SAVE_IMPORT_PLAN")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .DecimalIN("P_IS_3_L", isImport3L)
                .XmlIN("P_XML", XmlUtils.Serialize(new ImportPlanSender() { PostedList = postedList }))
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }

        public static bool SaveImportEffect(string userId, int periodId, List<PostedImportEffect> postedList, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.SAVE_IMPORT_EFFECT")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .XmlIN("P_XML", XmlUtils.Serialize(new ImportEffectSender() { PostedList = postedList }))
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }
        #endregion
        #region ImportCompliance

        public static DataTable GetKpisForIpCompliance(string userId, decimal periodId, out string errMsg)
        {
            List<object> ret = null;
            errMsg = "ROLE_GRANTED";
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_Q.GET_KPIS_FOR_IP_COMPLIANCE")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
               .VarcharOUT("P_ERR_MSG", 100)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable(new string[] { "P_ERR_MSG" }, out ret);
            errMsg = ret[0].ToString();
            return dt;
        }

        public static DataTable GetOuListForIpCompliance(string userId, decimal periodId, string kpiId, out string errMsg)
        {
            List<object> ret = null;
            errMsg = "ROLE_GRANTED";
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_Q.GET_OU_LIST_FOR_IP_COMPLIANCE")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
               .VarcharOUT("P_ERR_MSG", 100)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable(new string[] { "P_ERR_MSG" }, out ret);
            errMsg = ret[0].ToString();
            return dt;
        }

        public static DataTable GetDataForIpCompliance(string userId, decimal periodId, string kpiId, string outId, out string errMsg, int check = 0)
        {
            List<object> ret = null;
            errMsg = "ROLE_GRANTED";
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_Q.GET_DATA_FOR_IP_COMPLIANCE")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .DecimalIN("P_CHECK", check)
                .VarcharIN("P_KPI_ID", kpiId)
                .VarcharIN("P_OU_ID", outId)
               .VarcharOUT("P_ERR_MSG", 100)
                .RefCursorOut("P_LIST")
                .ExecuteDataTable(new string[] { "P_ERR_MSG" }, out ret);
            errMsg = ret[0].ToString();
            return dt;
        }

        public static bool CheckExistsInPeriod(string userId, string employeeNo, string kpiId, int periodId, ref string err)
        {
            List<object> ret = null;
            var dt = OraExecuter.Instance.ProcName("PKG_NEW_PMS_Q.CHECK_EXISTS_IN_PERIOD")
                .VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_EMPLOYEE_NO", employeeNo)
                .VarcharIN("P_KPI_ID", kpiId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .DecimalOUT("V_VAL")
                .ExecuteDataTable(new string[] { "V_VAL" }, out ret);
            int checkResult = 0;
            if (int.TryParse(ret[0].ToString(), out checkResult))
            {
                if (checkResult == 1)
                    return true;
                else if (checkResult == 0)
                    return false;
            }
            return false;
        }

        public static bool SaveImportIpCompliance(string userId, decimal periodId, List<PostedImportIpCompliance> list, ref string error)
        {
            string sql = "PKG_NEW_PMS_Q.SAVE_IMPORT_COMPLIANCE";
            OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString);
            OracleCommand cm = new OracleCommand(sql, cn);
            cm.CommandType = CommandType.StoredProcedure;
            cm.BindByName = true;

            OracleParameter P_USER_ID = new OracleParameter("P_USER_ID", OracleDbType.Varchar2);
            P_USER_ID.Direction = ParameterDirection.Input;
            P_USER_ID.Value = userId;
            OracleParameter P_PERIOD_ID = new OracleParameter("P_PERIOD_ID", OracleDbType.Decimal);
            P_PERIOD_ID.Direction = ParameterDirection.Input;
            P_PERIOD_ID.Value = periodId;

            OracleParameter P_XML = new OracleParameter("P_XML", OracleDbType.XmlType);
            P_XML.Direction = ParameterDirection.Input;
            P_XML.Value = XmlUtils.Serialize(new PostedImportIpComplianceOfUnitSender() { PostedList = list });

            OracleParameter P_ERR_CODE = new OracleParameter("P_ERR_CODE", OracleDbType.Decimal);
            P_ERR_CODE.Direction = ParameterDirection.Output;

            OracleParameter P_ERR_MSG = new OracleParameter("P_ERR_MSG", OracleDbType.Varchar2, 2000);
            P_ERR_MSG.Direction = ParameterDirection.Output;

            cm.Parameters.Add(P_USER_ID);
            cm.Parameters.Add(P_PERIOD_ID);
            cm.Parameters.Add(P_XML);
            cm.Parameters.Add(P_ERR_CODE);
            cm.Parameters.Add(P_ERR_MSG);
            cn.Open();
            cm.ExecuteNonQuery();

            cn.Close();

            var kq = int.Parse(P_ERR_CODE.Value.ToString());
            if (kq != -1)
            {
                error = P_ERR_MSG.Value != null ? P_ERR_MSG.Value.ToString() : string.Empty;
                return false;
            }
            return true;
        }
        #endregion

        public static bool SaveImportOuKpiResult(string userId, int periodId, List<PostedImportOuKpiResult> postedList, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.SAVE_IMPORT_OU_KPI_RESULT")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .XmlIN("P_XML", XmlUtils.Serialize(new ImportOuKpiResultSender() { PostedList = postedList }))
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }

        public static bool SaveImportOuProportion(string userId, int periodId, List<PostedImportOuProportion> postedList, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.SAVE_IMPORT_OU_PROPORTION")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .XmlIN("P_XML", XmlUtils.Serialize(new ImportOuProportionSender() { PostedList = postedList }))
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }

        public static DataTable GetDataProposeRanking(string userId, int periodId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_DATA_PROPOSE_RANKING").VarcharIN("P_USER_ID", userId).DecimalIN("P_PERIOD_ID", periodId).RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static bool SaveUpdateRanking(string userId, List<PostedUpdateRanking> postedList, int periodId, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.SAVE_UPDATE_PROPSE_RANKING")
                .VarcharIN("P_USER_ID", userId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .XmlIN("P_XML", XmlUtils.Serialize(new UpdateRankingSender() { PostedList = postedList }))
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error);
        }

        public static DataTable GetJobTitleList(string userId)
        {
            return OraExecuter.Instance.ProcName("PKG_NEW_PMS.GET_JOB_TITLE").VarcharIN("P_USER_ID", userId).RefCursorOut("P_LIST").ExecuteDataTable();
        }

        public static bool SaveUpdatEmpPeriod(string userId, string  employeeNo, int periodId, int jobTitleId, string ouId, string impersonator, ref string error)
        {
            var ret = OraExecuter.Instance.ProcName("PKG_NEW_PMS.SAVE_EMP_PERIOD")
                .VarcharIN("P_USER_ID", userId)
                .VarcharIN("P_EMPLOYEE_NO", employeeNo)
                .VarcharIN("P_OU_ID", ouId)
                .DecimalIN("P_PERIOD_ID", periodId)
                .DecimalIN("P_JOB_TITLE_ID", jobTitleId)
                .DecimalIN("P_IMPERSONATOR", impersonator)
                .DecimalOUT("P_ERR_CODE")
                .VarcharOUT("P_ERR_MSG", 2000)
                .ExecuteGetOutValues("P_ERR_CODE", "P_ERR_MSG");
            return HandlerResult(ret, ref error); 
        }
    }
    public class OraExecuter
    {
        private string procName;
        private List<OracleParameter> paramList;
        public OraExecuter ProcName(string procName)
        {
            this.procName = procName;
            return this;
        }
        private OraExecuter()
        {
            paramList = new List<OracleParameter>();
        }
        public static OraExecuter Instance
        {
            get
            {
                return new OraExecuter();
            }
        }
        public OraExecuter VarcharIN(string pName, string val)
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.Varchar2, ParameterDirection.Input) { Value = val });
            return this;
        }
        public OraExecuter VarcharOUT(string pName, int size)
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.Varchar2, ParameterDirection.Output) { Size = size });
            return this;
        }
        public OraExecuter VarcharIN_OUT(string pName, string val, int size)
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.Varchar2, ParameterDirection.InputOutput) { Value = val, Size = size });
            return this;
        }
        public OraExecuter RefCursorOut(string pName = "p_RET")
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.RefCursor, ParameterDirection.Output));
            return this;
        }
        public OraExecuter RefCursorOut(params string[] pNames)
        {
            foreach (string pName in pNames)
            {
                paramList.Add(new OracleParameter(pName, OracleDbType.RefCursor, ParameterDirection.Output));
            }
            return this;
        }
        public OraExecuter RefCursorReturn(string pName = "P_LIST")
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.RefCursor, ParameterDirection.ReturnValue));
            return this;
        }

        public OraExecuter DecimalIN(string pName, object val)
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.Decimal, ParameterDirection.Input) { Value = val });
            return this;
        }
        public OraExecuter DecimalOUT(string pName)
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.Decimal, ParameterDirection.Output));
            return this;
        }
        public OraExecuter DecimalOUT(params string[] pNames)
        {
            foreach (string pName in pNames)
            {
                paramList.Add(new OracleParameter(pName, OracleDbType.Decimal, ParameterDirection.Output));
            }
            return this;
        }
        public OraExecuter DecimalIN_OUT(string pName, object val)
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.Decimal, ParameterDirection.Input) { Value = val });
            return this;
        }
        public OraExecuter DateIN(string pName, DateTime? val)
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.Date, ParameterDirection.Input) { Value = val });
            return this;
        }
        public OraExecuter XmlIN(string pName, string xml)
        {
            paramList.Add(new OracleParameter(pName, OracleDbType.XmlType, ParameterDirection.Input) { Value = xml });
            return this;
        }
        public object ExecuteGetValue(string outParamName)
        {
            using (OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString))
            {
                cn.Open();
                OracleCommand cm = new OracleCommand(procName, cn);
                cm.CommandType = CommandType.StoredProcedure;
                cm.BindByName = true;
                cm.Parameters.AddRange(paramList.ToArray());
                cm.ExecuteNonQuery();
                var para = paramList.FirstOrDefault(q => q.ParameterName.Equals(outParamName));
                return para.Value;
            }
        }
        public List<object> ExecuteGetOutValues(params string[] outParamNames)
        {
            using (OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString))
            {
                cn.Open();
                OracleCommand cm = new OracleCommand(procName, cn);
                cm.CommandType = CommandType.StoredProcedure;
                cm.BindByName = true;
                cm.Parameters.AddRange(paramList.ToArray());
                cm.ExecuteNonQuery();
                var ret = new List<object>();
                foreach (string outParamName in outParamNames)
                {
                    var para = paramList.FirstOrDefault(q => q.ParameterName.Equals(outParamName));
                    ret.Add(para.Value);
                }
                return ret;
            }
        }
        public string ExecuteGetCLOB(string outParamName)
        {
            using (OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString))
            {
                cn.Open();
                OracleCommand cm = new OracleCommand(procName, cn);
                cm.CommandType = CommandType.StoredProcedure;
                cm.BindByName = true;
                cm.Parameters.AddRange(paramList.ToArray());
                cm.ExecuteNonQuery();
                var para = paramList.FirstOrDefault(q => q.ParameterName.Equals(outParamName));
                var ret = (Oracle.ManagedDataAccess.Types.OracleClob)para.Value;
                return ret?.Value;
            }
        }
        public DataSet ExecuteDataSet()
        {
            using (OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString))
            {
                cn.Open();
                OracleCommand cm = new OracleCommand(procName, cn);
                cm.CommandType = CommandType.StoredProcedure;
                cm.BindByName = true;
                cm.Parameters.AddRange(paramList.ToArray());
                OracleDataAdapter ap = new OracleDataAdapter(cm);
                DataSet ds = new DataSet();
                ap.Fill(ds);
                return ds;
            }
        }
        public DataSet ExecuteDataSet(string[] outParamNames, out List<object> outValues)
        {
            using (OracleConnection cn = new OracleConnection(
                System.Configuration.ConfigurationManager.ConnectionStrings["pms"].ConnectionString))
            {
                cn.Open();
                OracleCommand cm = new OracleCommand(procName, cn);
                cm.CommandType = CommandType.StoredProcedure;
                cm.BindByName = true;
                cm.Parameters.AddRange(paramList.ToArray());
                OracleDataAdapter ap = new OracleDataAdapter(cm);
                DataSet ds = new DataSet();
                ap.Fill(ds);
                outValues = new List<object>();
                foreach (string outParamName in outParamNames)
                {
                    var para = paramList.FirstOrDefault(q => q.ParameterName.Equals(outParamName));
                    outValues.Add(para.Value);
                }
                return ds;
            }
        }
        public DataTable ExecuteDataTable()
        {
            DataSet ds = ExecuteDataSet();
            return ds == null || ds.Tables.Count == 0 ? null : ds.Tables[0];
        }

        public DataTable ExecuteDataTable(string[] outParamNames, out List<object> outValues)
        {
            DataSet ds = ExecuteDataSet(outParamNames, out outValues);
            return ds == null || ds.Tables.Count == 0 ? null : ds.Tables[0];
        }

    }
}
