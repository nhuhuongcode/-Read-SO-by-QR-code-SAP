using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddlinePromotion
{
    public static class Globals
    {
        private static string path;

        private static SAPbouiCOM.Application SBO_Application;
        private static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Application SapApplication { get { return SBO_Application; } }
        public static SAPbobsCOM.Company SapCompany { get { return oCompany; } }


        #region Avariable config  b1
        private static int roundAmount = 6;
        private static int rountPercent = 6;
        private static int rountRate = 6;
        public static int RoundAmount { get { return roundAmount; } }
        public static int RountPercent { get { return rountPercent; } }
        public static int RountRate { get { return rountRate; } }
        public static void GetConfigB1()
        {
            try
            {
                SAPbouiCOM.DataTable oDataTable = Globals.GetSapDataTable("exec ESS_GetConfig");
                if (oDataTable.IsEmpty) return;
                roundAmount = int.Parse(oDataTable.GetValue("SumDec", 0).ToString().Trim());
                rountPercent = int.Parse(oDataTable.GetValue("PercentDec", 0).ToString().Trim());
                rountRate = int.Parse(oDataTable.GetValue("RateDec", 0).ToString().Trim());
            }
            catch { }
        }
        #endregion

        #region Variable Global
        public static int pvalrow = 1;
        public static System.Data.DataTable Table_Obj { get; set; }
        public static System.Data.DataTable Promos_Line { get; set; }
        public static System.Data.DataTable DataTable_Return { get; set; }

        public static int G_Round { get { return 0; } }
        public static string DocAR { get; set; }
        public static string DocNumSO { get; set; }
        public static string DocEntrySO { get; set; }
        public static bool ChooseCopyStatusChoose { get; set; }
        public static string ChooseCopyFormNameReturn { get; set; }
        #endregion

        #region Connect to Application
        public static bool SetApplicationUI()
        {
            string sConnectionString;
            SAPbouiCOM.SboGuiApi SboGuiApi;
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            if (Environment.GetCommandLineArgs().Length > 1)
                sConnectionString = Environment.GetCommandLineArgs()[1];
            else
                sConnectionString = Environment.GetCommandLineArgs()[0];
            try
            {
                SboGuiApi.Connect(sConnectionString);
                SBO_Application = SboGuiApi.GetApplication();
                return true;
            }
            catch { return false; }
        }
        //public static void connectToHana()
        //{
        //    string connectionString;
        //    connectionString = "DRIVER={HDBODBC32};UID=B1HADMIN;PWD=Hana@12345;SERVERNODE=183.91.11.140:30015;DATABASE=KVPC_GOLIVE";
        //    OdbcConnection conn = new OdbcConnection(connectionString);
        //    try
        //    {
        //        conn.Open();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //    MessageBox.Show("OK");
        //    conn.Close();
        //}
        public static bool SetApplication()
        {
            string sConnectionString;
            string sCookie = null;
            string sConnectionContext = null;
            SAPbouiCOM.SboGuiApi SboGuiApi;
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            if (Environment.GetCommandLineArgs().Length > 1)
                sConnectionString = Environment.GetCommandLineArgs()[1];
            else
                sConnectionString = Environment.GetCommandLineArgs()[0];
            try
            {
                SboGuiApi.Connect(sConnectionString);
                SBO_Application = SboGuiApi.GetApplication();
                oCompany = new SAPbobsCOM.Company();
                sCookie = oCompany.GetContextCookie();
                sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
                if (oCompany.Connected == true)
                    oCompany.Disconnect();
                if (oCompany.SetSboLoginContext(sConnectionContext) != 0)
                    return false;
                if (oCompany.Connect() != 0)
                    return false;
                SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SetPath();
                if (Directory.Exists(path))
                    Directory.Delete(path, true);
                return true;
            }
            catch { return false; }

        }
        public static int lRowNeedToMerge;


        // '''''''''''''''''''''''''''''''''
        //  Connect with connection string '
        // '''''''''''''''''''''''''''''''''
        public static int SetConnectionContext()
        {
            int setConnectionContextReturn = 0;

            string sCookie = null;
            string sConnectionContext = null;
            int lRetCode = 0;

            oCompany = new SAPbobsCOM.Company();

            sCookie = oCompany.GetContextCookie();
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);

            if (oCompany.Connected == true)
            {
                oCompany.Disconnect();
            }
            setConnectionContextReturn = oCompany.SetSboLoginContext(sConnectionContext);

            return setConnectionContextReturn;
        }


        // '''''''''''''''''
        //  Connect to SBO '
        // '''''''''''''''''
        public static int ConnectToCompany()
        {
            int connectToCompanyReturn = 0;

            // Establish the connection to the company database.
            connectToCompanyReturn = oCompany.Connect();

            return connectToCompanyReturn;
        }
        //public static bool SetApplication()
        //{
        //    string sConnectionString;
        //    string sCookie = null;
        //    string sConnectionContext = null;
        //    SAPbouiCOM.SboGuiApi SboGuiApi;
        //    SboGuiApi = new SAPbouiCOM.SboGuiApi();
        //    if (Environment.GetCommandLineArgs().Length > 1)
        //         sConnectionString = Environment.GetCommandLineArgs()[1];
        //     else
        //         sConnectionString = Environment.GetCommandLineArgs()[0];
        //    try
        //    {
        //        SboGuiApi.Connect(sConnectionString);
        //        SBO_Application = SboGuiApi.GetApplication();
        //        oCompany = new SAPbobsCOM.Company();
        //        sCookie = oCompany.GetContextCookie();
        //        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
        //        if (oCompany.Connected == true)
        //            oCompany.Disconnect();
        //        if (oCompany.SetSboLoginContext(sConnectionContext) != 0)
        //            return false;
        //        if (oCompany.Connect() != 0)
        //            return false;
        //        SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
        //        SetPath();
        //        if (Directory.Exists(path))
        //            Directory.Delete(path, true);
        //        return true;
        //    }
        //    catch { return false; }
        //}
        private static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //System.Environment.Exit(0);
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    //System.Environment.Exit(0);
                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }
        #endregion

        private static void SetPath()
        {
            path = Application.StartupPath + @"\Category\" + oCompany.CompanyName;
        }

        #region Lay thong tin cac truong danh muc
        public static System.Data.DataTable DatatableShop()
        {
            System.Data.DataTable DatatableShop = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\DatatableShop.xml"))
                {
                    string SQL = null;
                    DataRow Row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SQL = "select Code, Name,U_Address, U_AccCash, U_AccTran, U_REGION, U_COUNTRY from [@FPTSHOP] with(nolock) ";
                    DatatableShop.TableName = "FPTSHOP";
                    DatatableShop.Columns.Add("Code", Type.GetType("System.String"));
                    DatatableShop.Columns.Add("Name", Type.GetType("System.String"));
                    DatatableShop.Columns.Add("U_Address", Type.GetType("System.String"));
                    DatatableShop.Columns.Add("U_AccCash", Type.GetType("System.String"));
                    DatatableShop.Columns.Add("U_AccTran", Type.GetType("System.String"));
                    DatatableShop.Columns.Add("U_REGION", Type.GetType("System.String"));
                    DatatableShop.Columns.Add("U_COUNTRY", Type.GetType("System.String"));

                    SAPbouiCOM.DataTable oDataTable = Globals.GetSapDataTable(SQL);
                    if (!oDataTable.IsEmpty)
                    {
                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            Row = DatatableShop.NewRow();
                            Row["Code"] = oDataTable.GetValue("Code", i).ToString().Trim();
                            Row["Name"] = oDataTable.GetValue("Name", i).ToString().Trim();
                            Row["U_Address"] = oDataTable.GetValue("U_Address", i).ToString().Trim();
                            Row["U_AccCash"] = oDataTable.GetValue("U_AccCash", i).ToString().Trim();
                            Row["U_AccTran"] = oDataTable.GetValue("U_AccTran", i).ToString().Trim();
                            Row["U_REGION"] = oDataTable.GetValue("U_REGION", i).ToString().Trim();
                            Row["U_COUNTRY"] = oDataTable.GetValue("U_COUNTRY", i).ToString().Trim();
                            DatatableShop.Rows.Add(Row);
                        }
                        SaveDataTableToXML(path + @"\DatatableShop.xml", DatatableShop);
                    }
                }
                else
                    DatatableShop = ConvertXmlToDataTable(path + @"\DatatableShop.xml");

            }
            catch { }
            return DatatableShop;
        }
        public static System.Data.DataTable SaleEmployee()
        {
            System.Data.DataTable SaleEmployee = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\SaleEmployee.xml"))
                {
                    string SQL = null;
                    DataRow Row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);

                    SQL = "Select SlpCode, SlpName, U_Code_SH from OSLP with(nolock) WHERE Locked = 'N'";
                    SaleEmployee.TableName = "OSLP";
                    SaleEmployee.Columns.Add("SlpCode", Type.GetType("System.String"));
                    SaleEmployee.Columns.Add("SlpName", Type.GetType("System.String"));
                    SaleEmployee.Columns.Add("U_Code_SH", Type.GetType("System.String"));
                    SAPbouiCOM.DataTable oDataTable = Globals.GetSapDataTable(SQL);
                    if (!oDataTable.IsEmpty)
                    {

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            Row = SaleEmployee.NewRow();
                            Row["SlpCode"] = oDataTable.GetValue("SlpCode", i).ToString().Trim();
                            Row["SlpName"] = oDataTable.GetValue("SlpName", i).ToString().Trim();
                            Row["U_Code_SH"] = oDataTable.GetValue("U_Code_SH", i).ToString().Trim();
                            SaleEmployee.Rows.Add(Row);
                        }
                        SaveDataTableToXML(path + @"\SaleEmployee.xml", SaleEmployee);
                    }
                }
                else
                    SaleEmployee = ConvertXmlToDataTable(path + @"\SaleEmployee.xml");

            }
            catch { }
            return SaleEmployee;
        }
        public static void GetInfoLocal(out string g_GetHostName, out string g_IP_Address)
        {
            g_GetHostName = System.Net.Dns.GetHostName();
            g_IP_Address = "";
            try
            {
                g_IP_Address = System.Net.Dns.GetHostAddresses(g_GetHostName).GetValue(0).ToString();
                if (g_IP_Address.LastIndexOf(".") <= 0)
                {
                    g_IP_Address = System.Net.Dns.GetHostAddresses(g_GetHostName).GetValue(1).ToString();
                    if (g_IP_Address.LastIndexOf(".") <= 0)
                        g_IP_Address = System.Net.Dns.GetHostAddresses(g_GetHostName).GetValue(2).ToString();
                }

            }
            catch { }
        }
        public static DateTime GetSystemDate()
        {
            SAPbouiCOM.DataTable oDataTable = GetSapDataTable("select getdate() as SystemDate");
            if (!oDataTable.IsEmpty)
                return Convert.ToDateTime(oDataTable.GetValue("SystemDate", 0).ToString().Trim());
            return DateTime.Now;
        }
        public static System.Data.DataTable GetPticeListHO()
        {
            System.Data.DataTable g_PriceListHO = null;
            if (!File.Exists(path + @"\g_PriceListHO.xml"))
            {
                string SQL = "";
                string ConnectStr = "";
                System.Data.OleDb.OleDbConnection Olecon = new System.Data.OleDb.OleDbConnection();
                System.Data.OleDb.OleDbCommand OlemyCommand;
                System.Data.OleDb.OleDbDataAdapter OleAdapter;
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                SQL = "select U_SYSNAME ,U_VALUES , U_ServIP, U_DBLink , U_Port  from [@FPTSYS] where U_SYSCODE ='COMHO'";
                SAPbouiCOM.DataTable oDataTable = GetSapDataTable(SQL);
                oDataTable.ExecuteQuery(SQL);
                if (!oDataTable.IsEmpty)
                {
                    if (!string.IsNullOrEmpty(oDataTable.Columns.Item("U_DBLink").Cells.Item(0).ToString().Trim()))
                    {
                        if (string.IsNullOrEmpty(oDataTable.Columns.Item("U_Port").Cells.Item(0).ToString().Trim()))
                        {
                            ConnectStr = "Provider=SQLOLEDB;Data Source=" + oDataTable.GetValue("U_ServIP", 0).ToString().Trim();
                            ConnectStr += ";Persist Security Info=True;User ID=" + oDataTable.GetValue("U_SYSNAME", 0).ToString().Trim();
                            ConnectStr += ";Password=" + oDataTable.GetValue("U_VALUES", 0).ToString().Trim();
                            ConnectStr += ";Initial Catalog=" + oDataTable.GetValue("U_DBLink", 0).ToString().Trim() + ";Connect Timeout=300";

                        }
                        else
                        {
                            ConnectStr = "Provider=SQLOLEDB;Server=" + oDataTable.GetValue("U_ServIP", 0).ToString().Trim();
                            ConnectStr += "," + oDataTable.GetValue("U_Port", 0).ToString().Trim();
                            ConnectStr += ";Database=" + oDataTable.GetValue("U_DBLink", 0).ToString().Trim();
                            ConnectStr += ";User ID=" + oDataTable.GetValue("U_SYSNAME", 0).ToString().Trim() + ";Password=";
                            ConnectStr += oDataTable.GetValue("U_VALUES", 0).ToString().Trim();
                            ConnectStr += ";Trusted_Connection=False;Connect Timeout=300";
                        }
                    }
                }
                SQL = "select 9000000+ isnull(oo.DocNum,0) as DocNum,oo.U_ListName,U_BLine , U_ParList,isnull(oo.DocNum,0) as DocNumOld, oo1.U_ShpCod   from [@FPTOPLN1]  oo with(nolock), [@FPTITM12] oo1 with(nolock)  where " +
                                               "ISNULL(U_BLine ,'')<>'' and oo.DocEntry=oo1.docentry and ISNULL(oo1.U_ShpCod,'')<>'' and GETDATE() >= dbo.FPT_UnionDateTime(isnull(oo.U_FromDate, getdate()), isnull(oo.U_FromHour, 0)) " +
                                               "and GETDATE()<= dbo.FPT_UnionDateTime(isnull(oo.U_ToDate,getdate()) ,isnull(oo.U_ToHour ,2359)) ";
                Olecon.ConnectionString = ConnectStr;
                Olecon.Open();
                if (Olecon.State.ToString().Equals("Open"))
                {
                    OlemyCommand = new System.Data.OleDb.OleDbCommand(SQL, Olecon);
                    OlemyCommand.CommandTimeout = 300;
                    OleAdapter = new System.Data.OleDb.OleDbDataAdapter(OlemyCommand);
                    OleAdapter.Fill(g_PriceListHO);
                }
                Olecon.Close();
                Olecon = null;
                g_PriceListHO.TableName = "FPTOPLN1";
                SaveDataTableToXML(path + @"\g_PriceListHO.xml", g_PriceListHO);
            }
            else
            {
                g_PriceListHO = ConvertXmlToDataTable(path + @"\g_PriceListHO.xml");
            }

            return g_PriceListHO;
        }
        public static int GetRoundCurrency()
        {
            System.Data.DataTable GetRoundCurrency = new System.Data.DataTable();
            int RoundCurrency = 0;
            try
            {
                if (!File.Exists(path + @"\GetRoundCurrency.xml"))
                {
                    System.Data.DataRow row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable("select U_VALUES from [@FPTSYS] with(nolock) where U_SYSCODE='ROUND'");
                    GetRoundCurrency.TableName = "ROUNDCUR";
                    GetRoundCurrency.Columns.Add("RoundCurrency", Type.GetType("System.Int16"));
                    row = GetRoundCurrency.NewRow();
                    if (!oDataTable.IsEmpty)
                        RoundCurrency = Convert.ToInt16(oDataTable.GetValue(0, 0).ToString().Trim());
                    row["RoundCurrency"] = RoundCurrency;
                    GetRoundCurrency.Rows.Add(row);
                    SaveDataTableToXML(path + @"\GetRoundCurrency.xml", GetRoundCurrency);
                }
                else
                {
                    GetRoundCurrency = ConvertXmlToDataTable(path + @"\GetRoundCurrency.xml");
                    if (GetRoundCurrency != null)
                        if (GetRoundCurrency.Rows.Count > 0)
                            RoundCurrency = Convert.ToInt16(GetRoundCurrency.Rows[0][0]);
                }
            }
            catch { }
            return RoundCurrency;
        }
        public static string SetCurrency()
        {
            System.Data.DataTable SetCurrency = new System.Data.DataTable();
            string currency = "";
            try
            {
                if (!File.Exists(path + @"\SetCurrency.xml"))
                {
                    System.Data.DataRow row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable("select U_VALUES from [@FPTSYS] with(nolock) where U_SYSCODE='CURRENCY'");
                    SetCurrency.TableName = "CURRENCY";
                    SetCurrency.Columns.Add("Currency", Type.GetType("System.String"));
                    row = SetCurrency.NewRow();
                    if (!oDataTable.IsEmpty)
                        currency = oDataTable.GetValue(0, 0).ToString().Trim();
                    if (string.IsNullOrEmpty(currency))
                    {
                        oDataTable.Clear();
                        oDataTable = GetSapDataTable("select SysCurrncy  from OADM  with(nolock) where CompnyName ='" + oCompany.CompanyName.Trim() + "'");
                        if (!oDataTable.IsEmpty)
                            currency = oDataTable.GetValue(0, 0).ToString().Trim();
                    }
                    row["RoundCurrency"] = currency;
                    SetCurrency.Rows.Add(row);
                    SaveDataTableToXML(path + @"\SetCurrency.xml", SetCurrency);
                }
                else
                {
                    SetCurrency = ConvertXmlToDataTable(path + @"\SetCurrency.xml");
                    if (SetCurrency != null)
                        if (SetCurrency.Rows.Count > 0)
                            currency = SetCurrency.Rows[0][0].ToString().Trim();
                }
            }
            catch { }
            return currency;
        }
        public static System.Data.DataTable GetPticeListCompany()
        {
            System.Data.DataTable g_PriceListCompany = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_PriceListCompany.xml"))
                {
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    string sql = "select oo.DocNum,oo.U_ListName,U_BLine , U_ParList, U_ShopCod from [@FPTOPLN1]  oo where ";
                    sql += "GETDATE() >= dbo.FPT_UnionDateTime(isnull(oo.U_FromDate, getdate()), isnull(oo.U_FromHour, 0)) ";
                    sql += "and GETDATE()<= dbo.FPT_UnionDateTime(isnull(oo.U_ToDate,getdate()) ,isnull(oo.U_ToHour ,2359)) ";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_PriceListCompany.TableName = "FPTOPLN1";
                        g_PriceListCompany.Columns.Add("DocNum", Type.GetType("System.Double"));
                        g_PriceListCompany.Columns.Add("U_ListName", Type.GetType("System.String"));
                        g_PriceListCompany.Columns.Add("U_BLine", Type.GetType("System.String"));
                        g_PriceListCompany.Columns.Add("U_ParList", Type.GetType("System.Double"));
                        g_PriceListCompany.Columns.Add("U_ShopCod", Type.GetType("System.String"));
                        System.Data.DataRow row;
                        for (int i = 0; i <= oDataTable.Rows.Count - 1; i++)
                        {
                            row = g_PriceListCompany.NewRow();
                            row["DocNum"] = oDataTable.GetValue("DocNum", i).ToString().Trim();
                            row["U_ListName"] = oDataTable.GetValue("U_ListName", i).ToString().Trim();
                            row["U_BLine"] = oDataTable.GetValue("U_BLine", i).ToString().Trim();
                            row["U_ParList"] = oDataTable.GetValue("U_ParList", i).ToString().Trim();
                            row["U_ShopCod"] = oDataTable.GetValue("U_ShopCod", i).ToString().Trim();
                            g_PriceListCompany.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_PriceListCompany.xml", g_PriceListCompany);
                    }
                }
                else
                {
                    g_PriceListCompany = ConvertXmlToDataTable(path + @"\g_PriceListCompany.xml");
                }
            }
            catch { }
            return g_PriceListCompany;
        }
        public static System.Data.DataTable GetVAT()
        {
            System.Data.DataTable g_VatList = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_VatList.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "SELECT itemcode, VatGourpSa,isnull(Rate,1) as Rate FROM OITM A1 with(nolock), OVTG B with(nolock)  WHERE  B.Code= A1.VatGourpSa";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_VatList.TableName = "VAT";
                        g_VatList.Columns.Add("itemcode", Type.GetType("System.String"));
                        g_VatList.Columns.Add("VatGourpSa", Type.GetType("System.String"));
                        g_VatList.Columns.Add("Rate", Type.GetType("System.Double"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_VatList.NewRow();
                            row["itemcode"] = oDataTable.GetValue("itemcode", i).ToString().Trim();
                            row["VatGourpSa"] = oDataTable.GetValue("VatGourpSa", i).ToString().Trim();
                            row["Rate"] = oDataTable.GetValue("Rate", i).ToString().Trim();
                            g_VatList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_VatList.xml", g_VatList);

                    }
                }
                else
                    g_VatList = ConvertXmlToDataTable(path + @"\g_VatList.xml");
            }
            catch
            { }
            return g_VatList;
        }
        public static System.Data.DataTable GetBank()
        {
            System.Data.DataTable g_BankList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_BankList.xml"))
                {
                    System.Data.DataRow row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable("select BankCode,BankName from odsc with(nolock)");
                    if (!oDataTable.IsEmpty)
                    {
                        g_BankList.TableName = "ODSC";
                        g_BankList.Columns.Add("BankCode", Type.GetType("System.String"));
                        g_BankList.Columns.Add("BankName", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_BankList.NewRow();
                            row["BankCode"] = oDataTable.GetValue("BankCode", i).ToString().Trim();
                            row["BankName"] = oDataTable.GetValue("BankName", i).ToString().Trim();
                            g_BankList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_BankList.xml", g_BankList);
                    }

                }
                else
                {
                    g_BankList = ConvertXmlToDataTable(path + @"\g_BankList.xml");
                }
            }
            catch
            { }
            return g_BankList;

        }
        public static System.Data.DataTable g_OCRD_NVC()
        {
            System.Data.DataTable g_OCRD_NVC = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_OCRD_NVC.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "SELECT CardCode, CardName FROM OCRD with(nolock) where CardType = 'S' And U_NhVC = 'Y'";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_OCRD_NVC.TableName = "OCRD_NVC";
                        g_OCRD_NVC.Columns.Add("CardCode", Type.GetType("System.String"));
                        g_OCRD_NVC.Columns.Add("CardName", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_OCRD_NVC.NewRow();
                            row["CardCode"] = oDataTable.GetValue("CardCode", i).ToString().Trim();
                            row["CardName"] = oDataTable.GetValue("CardName", i).ToString().Trim();
                            g_OCRD_NVC.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_OCRD_NVC.xml", g_OCRD_NVC);
                    }

                }
                else
                {
                    g_OCRD_NVC = ConvertXmlToDataTable(path + @"\g_OCRD_NVC.xml");
                }
            }
            catch
            { }
            return g_OCRD_NVC;
        }
        public static System.Data.DataTable GetWhs()
        {
            System.Data.DataTable g_WhsCode = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_WhsCode.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "select U_Code_SH, WhsCode,WhsName,U_Whs_Type, U_Dep , U_CogDep,RevenuesAc, SaleCostAc   from OWHS with(nolock)";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_WhsCode.TableName = "OWHS";
                        g_WhsCode.Columns.Add("U_Code_SH", Type.GetType("System.String"));
                        g_WhsCode.Columns.Add("WhsCode", Type.GetType("System.String"));
                        g_WhsCode.Columns.Add("WhsName", Type.GetType("System.String"));
                        g_WhsCode.Columns.Add("U_Whs_Type", Type.GetType("System.String"));

                        g_WhsCode.Columns.Add("U_Dep", Type.GetType("System.String"));
                        g_WhsCode.Columns.Add("U_CogDep", Type.GetType("System.String"));

                        g_WhsCode.Columns.Add("RevenuesAc", Type.GetType("System.String"));
                        g_WhsCode.Columns.Add("SaleCostAc", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_WhsCode.NewRow();
                            row["U_Code_SH"] = oDataTable.GetValue("U_Code_SH", i).ToString().Trim();
                            row["WhsCode"] = oDataTable.GetValue("WhsCode", i).ToString().Trim();
                            row["WhsName"] = oDataTable.GetValue("WhsName", i).ToString().Trim();
                            row["U_Whs_Type"] = oDataTable.GetValue("U_Whs_Type", i).ToString().Trim();
                            row["U_Dep"] = oDataTable.GetValue("U_Dep", i).ToString().Trim();
                            row["U_CogDep"] = oDataTable.GetValue("U_CogDep", i).ToString().Trim();
                            row["RevenuesAc"] = oDataTable.GetValue("RevenuesAc", i).ToString().Trim();
                            row["SaleCostAc"] = oDataTable.GetValue("SaleCostAc", i).ToString().Trim();

                            g_WhsCode.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_WhsCode.xml", g_WhsCode);
                    }
                }
                else
                {
                    g_WhsCode = ConvertXmlToDataTable(path + @"\g_WhsCode.xml");
                }
            }
            catch
            { }
            return g_WhsCode;
        }
        public static System.Data.DataTable GetOSHP()
        {
            System.Data.DataTable g_OSHPList = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_OSHPList.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "select abc.TrnspCode , abc.TrnspName  from  OSHP abc with(nolock) order by TrnspCode";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_OSHPList.TableName = "OSHP";
                        g_OSHPList.Columns.Add("TrnspCode", Type.GetType("System.String"));
                        g_OSHPList.Columns.Add("OcrName", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_OSHPList.NewRow();
                            row["TrnspCode"] = oDataTable.GetValue("TrnspCode", i).ToString().Trim();
                            row["TrnspName"] = oDataTable.GetValue("TrnspName", i).ToString().Trim();

                            g_OSHPList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_OSHPList.xml", g_OSHPList);
                    }
                }
                else
                {
                    g_OSHPList = ConvertXmlToDataTable(path + @"\g_OSHPList.xml");
                }
            }
            catch
            { }
            return g_OSHPList;
        }
        public static System.Data.DataTable GetOOCR()
        {
            System.Data.DataTable g_OOCRList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_OOCRList.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "select OcrCode,OcrName from OOCR with(nolock) where (OcrCode='B01' or OcrCode='B02' or OcrCode='B05') and DimCode=4 and Active='Y'";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_OOCRList.TableName = "OSHP";
                        g_OOCRList.Columns.Add("OcrCode", Type.GetType("System.String"));
                        g_OOCRList.Columns.Add("OcrName", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_OOCRList.NewRow();
                            row["OcrCode"] = oDataTable.GetValue("OcrCode", i).ToString().Trim();
                            row["OcrName"] = oDataTable.GetValue("OcrName", i).ToString().Trim();

                            g_OOCRList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_OOCRList.xml", g_OOCRList);
                    }
                }
                else
                {
                    g_OOCRList = ConvertXmlToDataTable(path + @"\g_OOCRList.xml");
                }
            }
            catch
            { }
            return g_OOCRList;
        }
        public static System.Data.DataTable GetFPTSAL_TYPE()
        {
            System.Data.DataTable g_FPTSAL_TYPEList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_FPTSAL_TYPEList.xml"))
                {
                    System.Data.DataRow row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable("SELECT Code,U_NAME FROM [@FPTSAL_TYPE] with(nolock) Order by Code");
                    if (!oDataTable.IsEmpty)
                    {
                        g_FPTSAL_TYPEList.TableName = "FPTSAL_TYPE";
                        g_FPTSAL_TYPEList.Columns.Add("Code", Type.GetType("System.String"));
                        g_FPTSAL_TYPEList.Columns.Add("U_NAME", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_FPTSAL_TYPEList.NewRow();
                            row["Code"] = oDataTable.GetValue("Code", i).ToString().Trim();
                            row["U_NAME"] = oDataTable.GetValue("U_NAME", i).ToString().Trim();

                            g_FPTSAL_TYPEList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_FPTSAL_TYPEList.xml", g_FPTSAL_TYPEList);
                    }
                }
                else
                {
                    g_FPTSAL_TYPEList = ConvertXmlToDataTable(path + @"\g_FPTSAL_TYPEList.xml");
                }
            }
            catch
            { }
            return g_FPTSAL_TYPEList;
        }
        public static System.Data.DataTable GetOCRN()
        {
            System.Data.DataTable g_OCRNList = new System.Data.DataTable();

            try
            {

                if (!File.Exists(path + @"\g_OCRNList.xml"))
                {
                    System.Data.DataRow row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable("SELECT  CurrCode, CurrName FROM OCRN with(nolock) ORDER BY CurrCode");
                    if (!oDataTable.IsEmpty)
                    {
                        g_OCRNList.TableName = "OCRN";
                        g_OCRNList.Columns.Add("CurrCode", Type.GetType("System.String"));
                        g_OCRNList.Columns.Add("CurrName", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_OCRNList.NewRow();
                            row["CurrCode"] = oDataTable.GetValue("CurrCode", i).ToString().Trim();
                            row["CurrName"] = oDataTable.GetValue("CurrName", i).ToString().Trim();

                            g_OCRNList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_OCRNList.xml", g_OCRNList);
                    }
                }
                else
                {
                    g_OCRNList = ConvertXmlToDataTable(path + @"\g_OCRNList.xml");
                }
            }
            catch
            { }
            return g_OCRNList;
        }
        public static System.Data.DataTable GetOPYM()
        {
            System.Data.DataTable g_OPYMList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_OPYMList.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "Select  a.PayMethCod, a.Descript  from OPYM a with(nolock) where a.Active='Y' and TYPE='I'";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_OPYMList.TableName = "OPYM";
                        g_OPYMList.Columns.Add("Code", Type.GetType("System.String"));
                        g_OPYMList.Columns.Add("Name", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_OPYMList.NewRow();
                            row["Code"] = oDataTable.GetValue("Code", i).ToString().Trim();
                            row["Name"] = oDataTable.GetValue("Name", i).ToString().Trim();

                            g_OPYMList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_OPYMList.xml", g_OPYMList);
                    }
                }
                else
                {
                    g_OPYMList = ConvertXmlToDataTable(path + @"\g_OPYMList.xml");
                }
            }
            catch
            { }
            return g_OPYMList;
        }
        public static System.Data.DataTable GetFPTSAL_PAY()
        {
            System.Data.DataTable g_FPTSAL_PAYList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_FPTSAL_PAYList.xml"))
                {
                    System.Data.DataRow row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable("SELECT  Code,U_NAME FROM [@FPTSAL_PAY] with(nolock)");
                    if (!oDataTable.IsEmpty)
                    {
                        g_FPTSAL_PAYList.TableName = "FPTSAL_PAY";
                        g_FPTSAL_PAYList.Columns.Add("Code", Type.GetType("System.String"));
                        g_FPTSAL_PAYList.Columns.Add("U_NAME", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_FPTSAL_PAYList.NewRow();
                            row["Code"] = oDataTable.GetValue("Code", i).ToString().Trim();
                            row["U_NAME"] = oDataTable.GetValue("U_NAME", i).ToString().Trim();

                            g_FPTSAL_PAYList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_FPTSAL_PAYList.xml", g_FPTSAL_PAYList);
                    }
                }
                else
                {
                    g_FPTSAL_PAYList = ConvertXmlToDataTable(path + @"\g_FPTSAL_PAYList.xml");
                }
            }
            catch
            { }
            return g_FPTSAL_PAYList;
        }
        public static System.Data.DataTable GetOCTG()
        {
            System.Data.DataTable g_OCTGList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_OCTGList.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "SELECT GroupNum, PymntGroup FROM OCTG with(nolock) WHERE GroupNum>0 ORDER BY GroupNum";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_OCTGList.TableName = "OCTG";
                        g_OCTGList.Columns.Add("GroupNum", Type.GetType("System.Double"));
                        g_OCTGList.Columns.Add("PymntGroup", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_OCTGList.NewRow();
                            row["GroupNum"] = oDataTable.GetValue("GroupNum", i).ToString().Trim();
                            row["PymntGroup"] = oDataTable.GetValue("PymntGroup", i).ToString().Trim();

                            g_OCTGList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_OCTGList.xml", g_OCTGList);
                    }
                }
                else
                {
                    g_OCTGList = ConvertXmlToDataTable(path + @"\g_OCTGList.xml");
                }
            }
            catch
            { }
            return g_OCTGList;
        }
        public static System.Data.DataTable GetSHOP()
        {
            System.Data.DataTable g_ShopList = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_ShopList.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "select distinct  a.Code, a.Name from [@FPTSHOP] as a with(nolock),[@FPTRO_SH] as b with(nolock) where U_Code_US ='" + Globals.SapCompany.UserName.Trim() + "' and (U_Code_SH='ALL' or a.Code = b.U_Code_SH) and a.U_Status='Y'";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_ShopList.TableName = "FPTSHOP";
                        g_ShopList.Columns.Add("Code", Type.GetType("System.Double"));
                        g_ShopList.Columns.Add("NAME", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_ShopList.NewRow();
                            row["Code"] = oDataTable.GetValue("Code", i).ToString().Trim();
                            row["NAME"] = oDataTable.GetValue("NAME", i).ToString().Trim();

                            g_ShopList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_ShopList.xml", g_ShopList);
                    }
                }
                else
                {
                    g_ShopList = ConvertXmlToDataTable(path + @"\g_ShopList.xml");
                }
            }
            catch
            { }
            return g_ShopList;
        }
        public static System.Data.DataTable GetBill()
        {
            System.Data.DataTable g_BillList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_BillList.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "Select U_Symbol, U_ShpCod from [@FPTBILL] with(nolock) Where  U_Status = '0'";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_BillList.TableName = "FPTBILL";
                        g_BillList.Columns.Add("U_Symbol", Type.GetType("System.String"));
                        g_BillList.Columns.Add("U_ShpCod", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_BillList.NewRow();
                            row["U_Symbol"] = oDataTable.GetValue("U_Symbol", i).ToString().Trim();
                            row["U_ShpCod"] = oDataTable.GetValue("U_ShpCod", i).ToString().Trim();

                            g_BillList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_BillList.xml", g_BillList);
                    }
                }
                else
                {
                    g_BillList = ConvertXmlToDataTable(path + @"\g_BillList.xml");
                }
            }
            catch
            { }
            return g_BillList;
        }
        public static System.Data.DataTable GetOSLP()
        {
            System.Data.DataTable g_OSLPList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_OSLPList.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "SELECT SlpCode,SlpName, U_Code_SH FROM OSLP with(nolock) WHERE Locked = 'N'";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_OSLPList.TableName = "OSLP";
                        g_OSLPList.Columns.Add("SlpCode", Type.GetType("System.Double"));
                        g_OSLPList.Columns.Add("SlpName", Type.GetType("System.String"));
                        g_OSLPList.Columns.Add("U_Code_SH", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_OSLPList.NewRow();
                            row["SlpCode"] = oDataTable.GetValue("SlpCode", i).ToString().Trim();
                            row["SlpName"] = oDataTable.GetValue("SlpName", i).ToString().Trim();
                            row["U_Code_SH"] = oDataTable.GetValue("U_Code_SH", i).ToString().Trim();
                            g_OSLPList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_OSLPList.xml", g_OSLPList);
                    }
                }
                else
                {
                    g_OSLPList = ConvertXmlToDataTable(path + @"\g_OSLPList.xml");
                }
            }
            catch
            { }
            return g_OSLPList;
        }
        public static System.Data.DataTable GetCRD1()
        {
            System.Data.DataTable g_CRD1List = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_CRD1List.xml"))
                {
                    System.Data.DataRow row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable("SELECT ADDRESS, CardCode FROM CRD1 with(nolock)");
                    if (!oDataTable.IsEmpty)
                    {
                        g_CRD1List.TableName = "CRD1";
                        g_CRD1List.Columns.Add("ADDRESS", Type.GetType("System.String"));
                        g_CRD1List.Columns.Add("CardCode", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_CRD1List.NewRow();
                            row["ADDRESS"] = oDataTable.GetValue("ADDRESS", i).ToString().Trim();
                            row["CardCode"] = oDataTable.GetValue("CardCode", i).ToString().Trim();
                            g_CRD1List.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_CRD1List.xml", g_CRD1List);
                    }
                }
                else
                {
                    g_CRD1List = ConvertXmlToDataTable(path + @"\g_CRD1List.xml");
                }
            }
            catch
            { }
            return g_CRD1List;
        }
        public static System.Data.DataTable GetCaptionItem()
        {
            System.Data.DataTable g_CaptionItem = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_CaptionItem.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "SELECT [U_FormUID],[U_FormType],[U_Item],[U_Column] ,[U_Text], ButtonIndex from [@FPTLANG] with(nolock) ";
                    sql += " WHERE U_Lang=CONVERT(nvarchar(10), isnull((select U_Values from [@FPTSYS] with(nolock) where U_SysCode='LNS'),'EN'))  ORDER BY U_FormUID";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_CaptionItem.TableName = "FPTLANG";
                        g_CaptionItem.Columns.Add("U_FormUID", Type.GetType("System.String"));
                        g_CaptionItem.Columns.Add("U_FormType", Type.GetType("System.String"));

                        g_CaptionItem.Columns.Add("U_Item", Type.GetType("System.String"));
                        g_CaptionItem.Columns.Add("U_Column", Type.GetType("System.String"));
                        g_CaptionItem.Columns.Add("U_Text", Type.GetType("System.String"));
                        g_CaptionItem.Columns.Add("ButtonIndex", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_CaptionItem.NewRow();
                            row["U_FormUID"] = oDataTable.GetValue("U_FormUID", i).ToString().Trim();
                            row["U_FormType"] = oDataTable.GetValue("U_FormType", i).ToString().Trim();
                            row["U_Item"] = oDataTable.GetValue("U_Item", i).ToString().Trim();
                            row["U_Column"] = oDataTable.GetValue("U_Column", i).ToString().Trim();
                            row["U_Text"] = oDataTable.GetValue("U_Text", i).ToString().Trim();
                            row["ButtonIndex"] = oDataTable.GetValue("ButtonIndex", i).ToString().Trim();
                            g_CaptionItem.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_CaptionItem.xml", g_CaptionItem);

                    }
                }
                else
                {
                    g_CaptionItem = ConvertXmlToDataTable(path + @"\g_CaptionItem.xml");
                }
            }
            catch
            { }
            return g_CaptionItem;

        }
        public static System.Data.DataTable GetOCPR()
        {
            System.Data.DataTable g_OCPRList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_OCPRList.xml"))
                {
                    System.Data.DataRow row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable("SELECT CntctCode, Name, CardCode from OCPR with(nolock)");
                    if (!oDataTable.IsEmpty)
                    {
                        g_OCPRList.TableName = "OCPR";
                        g_OCPRList.Columns.Add("CntctCode", Type.GetType("System.String"));
                        g_OCPRList.Columns.Add("Name", Type.GetType("System.String"));
                        g_OCPRList.Columns.Add("CardCode", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_OCPRList.NewRow();
                            row["CntctCode"] = oDataTable.GetValue("CntctCode", i).ToString().Trim();
                            row["Name"] = oDataTable.GetValue("Name", i).ToString().Trim();
                            row["CardCode"] = oDataTable.GetValue("CardCode", i).ToString().Trim();
                            g_OCPRList.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_OCPRList.xml", g_OCPRList);
                    }
                }
                else
                {
                    g_OCPRList = ConvertXmlToDataTable(path + @"\g_OCPRList.xml");
                }
            }
            catch
            { }
            return g_OCPRList;
        }
        //public static System.Data.DataTable GetOITM()
        //{
        //    System.Data.DataTable g_ItemList = new System.Data.DataTable();

        //    try
        //    {
        //        if (!File.Exists(path + @"\g_ItemList.xml"))
        //        {
        //            System.Data.DataRow row = null;
        //            string sql = "";
        //            if (!Directory.Exists(path))
        //                Directory.CreateDirectory(path);

        //            sql = "select o.U_OcrCode2, o.U_OcrCode3, o.U_CogsO2, o.U_CogsO3, ItemCode, ItemName,o.CodeBars ,";
        //            sql += " o.ItemType ,o.ItmsGrpCod , o.U_NG_CODE,U_NHOM, U_DongHH , o.InvntItem , o.SellItem , o.U_TGBH, ";
        //            sql += "o.FirmCode , o.ManSerNum, o.NumInSale , o.VatGourpSa    from OITM o with(nolock) where o.SellItem='Y'";
        //            SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
        //            if (!oDataTable.IsEmpty)
        //            {
        //                g_ItemList.TableName = "OITM";
        //                g_ItemList.Columns.Add("ItemCode", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("ItemName", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("CodeBars", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("ItemType", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("ItmsGrpCod", Type.GetType("System.Double"));
        //                g_ItemList.Columns.Add("U_NG_CODE", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("U_NHOM", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("U_DongHH", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("InvntItem", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("SellItem", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("U_TGBH", Type.GetType("System.Double"));
        //                g_ItemList.Columns.Add("FirmCode", Type.GetType("System.Double"));
        //                g_ItemList.Columns.Add("ManSerNum", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("NumInSale", Type.GetType("System.Double"));
        //                g_ItemList.Columns.Add("VatGourpSa", Type.GetType("System.String"));

        //                g_ItemList.Columns.Add("U_OcrCode2", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("U_OcrCode3", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("U_CogsO2", Type.GetType("System.String"));
        //                g_ItemList.Columns.Add("U_CogsO3", Type.GetType("System.String"));

        //                for (int i = 0; i < oDataTable.Rows.Count; i++)
        //                {
        //                    row = g_ItemList.NewRow();
        //                    row["U_OcrCode2"] = oDataTable.GetValue("U_OcrCode2", i).ToString().Trim();
        //                    row["U_OcrCode3"] = oDataTable.GetValue("U_OcrCode3", i).ToString().Trim();
        //                    row["U_CogsO2"] = oDataTable.GetValue("U_CogsO2", i).ToString().Trim();
        //                    row["U_CogsO3"] = oDataTable.GetValue("U_CogsO3", i).ToString().Trim();
        //                    row["ItemCode"] = oDataTable.GetValue("ItemCode", i).ToString().Trim();
        //                    row["ItemName"] = oDataTable.GetValue("ItemName", i).ToString().Trim();
        //                    row["CodeBars"] = oDataTable.GetValue("CodeBars", i).ToString().Trim();
        //                    row["ItemType"] = oDataTable.GetValue("ItemType", i).ToString().Trim();
        //                    if (!string.IsNullOrEmpty(oDataTable.GetValue("ItmsGrpCod", i).ToString().Trim()))
        //                    {
        //                        row["ItmsGrpCod"] = oDataTable.GetValue("ItmsGrpCod", i).ToString().Trim();
        //                    }
        //                    row["U_NG_CODE"] = oDataTable.GetValue("U_NG_CODE", i).ToString().Trim();
        //                    row["U_NHOM"] = oDataTable.GetValue("U_NHOM", i).ToString().Trim();
        //                    row["U_DongHH"] = oDataTable.GetValue("U_DongHH", i).ToString().Trim();
        //                    row["InvntItem"] = oDataTable.GetValue("InvntItem", i).ToString().Trim();
        //                    row["SellItem"] = oDataTable.GetValue("SellItem", i).ToString().Trim();
        //                    if (!string.IsNullOrEmpty(oDataTable.GetValue("U_TGBH", i).ToString().Trim()))
        //                    {
        //                        row["U_TGBH"] = oDataTable.GetValue("U_TGBH", i).ToString().Trim();
        //                    }
        //                    if (!string.IsNullOrEmpty(oDataTable.GetValue("FirmCode", i).ToString().Trim()))
        //                    {
        //                        row["FirmCode"] = oDataTable.GetValue("FirmCode", i).ToString().Trim();
        //                    }
        //                    row["ManSerNum"] = oDataTable.GetValue("ManSerNum", i).ToString().Trim();
        //                    if (!string.IsNullOrEmpty(oDataTable.GetValue("NumInSale", i).ToString().Trim()))
        //                    {
        //                        row["NumInSale"] = oDataTable.GetValue("NumInSale", i).ToString().Trim();
        //                    }
        //                    row["VatGourpSa"] = oDataTable.GetValue("VatGourpSa", i).ToString().Trim();
        //                    g_ItemList.Rows.Add(row);
        //                }
        //                SaveDataTableToXML(path + @"\g_ItemList.xml", g_ItemList);
        //            }
        //        }
        //        else
        //        {
        //            g_ItemList = ConvertXmlToDataTable(path + @"\g_ItemList.xml");
        //        }
        //    }
        //    catch
        //    { }
        //    return g_ItemList;

        //}

        public static System.Data.DataTable GetODPI()
        {
            System.Data.DataTable g_ODPIList = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_ODPIList.xml"))
                {
                    System.Data.DataTable g_Combobox_Value = null;
                    System.Data.DataRow row = null;
                    System.Data.DataRow[] rowArr;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    g_ODPIList.TableName = "ODPI";
                    g_ODPIList.Columns.Add("Code", Type.GetType("System.String"));
                    g_ODPIList.Columns.Add("NAME", Type.GetType("System.String"));

                    g_Combobox_Value = GetComboBoxValue();
                    if (g_Combobox_Value == null) return new System.Data.DataTable();
                    if (g_Combobox_Value.Rows.Count <= 0) return new System.Data.DataTable();
                    rowArr = g_Combobox_Value.Select("U_FormUID='FPTSO' and [U_Item]='U_DMoney'");
                    if (rowArr.Length > 0)
                    {
                        for (int i = 0; i < rowArr.Length; i++)
                        {
                            row = g_ODPIList.NewRow();
                            row["Code"] = rowArr[i]["U_Code"].ToString().Trim();
                            row["NAME"] = rowArr[i]["U_Name"].ToString().Trim();
                            g_ODPIList.Rows.Add(row);
                        }
                    }
                    sql = "SELECT 'N' AS Code, N'Không đặt cọc' AS NAME UNION ALL SELECT 'Y' AS Code,"
                           + "N'Có đặt cọc' AS NAME UNION ALL SELECT 'P' AS Code, N'Đã trả lại đặt cọc' AS NAME ";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_ODPIList.NewRow();
                            row["Code"] = oDataTable.GetValue("Code", i).ToString().Trim();
                            row["NAME"] = oDataTable.GetValue("NAME", i).ToString().Trim();

                            g_ODPIList.Rows.Add(row);
                        }
                    }
                    SaveDataTableToXML(path + @"\g_ODPIList.xml", g_ODPIList);
                }
                else
                {
                    g_ODPIList = ConvertXmlToDataTable(path + @"\g_ODPIList.xml");
                }
            }
            catch
            { }
            return g_ODPIList;
        }
        public static System.Data.DataTable GetMoney()
        {
            System.Data.DataTable g_MoneyList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_MoneyList.xml"))
                {
                    System.Data.DataTable g_Combobox_Value = null;
                    System.Data.DataRow row = null;
                    System.Data.DataRow[] rowArr;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    g_MoneyList.TableName = "MoneyList";
                    g_MoneyList.Columns.Add("Code", Type.GetType("System.String"));
                    g_MoneyList.Columns.Add("NAME", Type.GetType("System.String"));

                    g_Combobox_Value = GetComboBoxValue();
                    if (g_Combobox_Value == null) return new System.Data.DataTable();
                    if (g_Combobox_Value.Rows.Count <= 0) return new System.Data.DataTable();
                    rowArr = g_Combobox_Value.Select("U_FormUID='FPTSO' and [U_Item]='U_Cmoney'");
                    if (rowArr.Length > 0)
                    {
                        for (int i = 0; i < rowArr.Length; i++)
                        {
                            row = g_MoneyList.NewRow();
                            row["Code"] = rowArr[i]["U_Code"].ToString().Trim();
                            row["NAME"] = rowArr[i]["U_Name"].ToString().Trim();
                            g_MoneyList.Rows.Add(row);
                        }
                    }

                    row = g_MoneyList.NewRow();
                    row["Code"] = "N";
                    row["NAME"] = "Chưa thu tiền";
                    g_MoneyList.Rows.Add(row);

                    row = g_MoneyList.NewRow();
                    row["Code"] = "Y";
                    row["NAME"] = "Đã thu tiền";
                    g_MoneyList.Rows.Add(row);
                    SaveDataTableToXML(path + @"\g_MoneyList.xml", g_MoneyList);
                }
                else
                {
                    g_MoneyList = ConvertXmlToDataTable(path + @"\g_MoneyList.xml");
                }
            }
            catch
            { }
            return g_MoneyList;
        }
        public static System.Data.DataTable GetStatus()
        {
            System.Data.DataTable g_StatusList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_StatusList.xml"))
                {
                    System.Data.DataTable g_Combobox_Value = null;
                    System.Data.DataRow row = null;
                    System.Data.DataRow[] rowArr;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    g_StatusList.TableName = "StatusList";
                    g_StatusList.Columns.Add("Code", Type.GetType("System.String"));
                    g_StatusList.Columns.Add("NAME", Type.GetType("System.String"));

                    g_Combobox_Value = GetComboBoxValue();
                    if (g_Combobox_Value == null) return new System.Data.DataTable();
                    if (g_Combobox_Value.Rows.Count <= 0) return new System.Data.DataTable();
                    rowArr = g_Combobox_Value.Select("U_FormUID='FPTSO' and [U_Item]='U_Status'");
                    if (rowArr.Length > 0)
                    {
                        for (int i = 0; i < rowArr.Length; i++)
                        {
                            row = g_StatusList.NewRow();
                            row["Code"] = rowArr[i]["U_Code"].ToString().Trim();
                            row["NAME"] = rowArr[i]["U_Name"].ToString().Trim();
                            g_StatusList.Rows.Add(row);
                        }
                    }

                    row = g_StatusList.NewRow();
                    row["Code"] = "1";
                    row["NAME"] = "Mở";
                    g_StatusList.Rows.Add(row);

                    row = g_StatusList.NewRow();
                    row["Code"] = "2";
                    row["NAME"] = "Hủy";
                    g_StatusList.Rows.Add(row);

                    row = g_StatusList.NewRow();
                    row["Code"] = "3";
                    row["NAME"] = "Đã đóng";
                    g_StatusList.Rows.Add(row);

                    row = g_StatusList.NewRow();
                    row["Code"] = "4";
                    row["NAME"] = "Đã đẩy API";
                    g_StatusList.Rows.Add(row);

                    row = g_StatusList.NewRow();
                    row["Code"] = "5";
                    row["NAME"] = "Đang xử lý API";
                    g_StatusList.Rows.Add(row);

                    SaveDataTableToXML(path + @"\g_StatusList.xml", g_StatusList);
                }
                else
                {
                    g_StatusList = ConvertXmlToDataTable(path + @"\g_StatusList.xml");
                }
            }
            catch
            { }
            return g_StatusList;
        }
        public static System.Data.DataTable GetItemOut()
        {
            System.Data.DataTable g_ItemOutList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_ItemOutList.xml"))
                {
                    System.Data.DataTable g_Combobox_Value = null;
                    System.Data.DataRow row = null;
                    System.Data.DataRow[] rowArr;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    g_ItemOutList.TableName = "ItemOutList";
                    g_ItemOutList.Columns.Add("Code", Type.GetType("System.String"));
                    g_ItemOutList.Columns.Add("NAME", Type.GetType("System.String"));
                    g_Combobox_Value = GetComboBoxValue();
                    if (g_Combobox_Value == null) return new System.Data.DataTable();
                    if (g_Combobox_Value.Rows.Count <= 0) return new System.Data.DataTable();
                    rowArr = g_Combobox_Value.Select("U_FormUID='FPTSO' and [U_Item]='U_ItmOut1'");
                    if (rowArr.Length > 0)
                    {
                        for (int i = 0; i < rowArr.Length; i++)
                        {
                            row = g_ItemOutList.NewRow();
                            row["Code"] = rowArr[i]["U_Code"].ToString().Trim();
                            row["NAME"] = rowArr[i]["U_Name"].ToString().Trim();
                            g_ItemOutList.Rows.Add(row);
                        }
                    }

                    sql = "SELECT 'Y' AS Code, N'Xuất hàng' AS NAME UNION ALL SELECT 'N' AS Code, N'Chưa xuất hàng' AS NAME ";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_ItemOutList.NewRow();
                            row["Code"] = oDataTable.GetValue("Code", i).ToString().Trim();
                            row["NAME"] = oDataTable.GetValue("NAME", i).ToString().Trim();
                            g_ItemOutList.Rows.Add(row);
                        }
                    }
                    SaveDataTableToXML(path + @"\g_ItemOutList.xml", g_ItemOutList);
                }
                else
                {
                    g_ItemOutList = ConvertXmlToDataTable(path + @"\g_ItemOutList.xml");
                }
            }
            catch
            { }
            return g_ItemOutList;
        }
        public static System.Data.DataTable GetSOTYPE()
        {
            System.Data.DataTable g_SOTYPEList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_SOTYPEList.xml"))
                {
                    System.Data.DataTable g_Combobox_Value = null;
                    System.Data.DataRow row = null;
                    System.Data.DataRow[] rowArr;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    g_SOTYPEList.TableName = "SOTYPE";
                    g_SOTYPEList.Columns.Add("Code", Type.GetType("System.String"));
                    g_SOTYPEList.Columns.Add("NAME", Type.GetType("System.String"));

                    g_Combobox_Value = GetComboBoxValue();
                    if (g_Combobox_Value == null) return new System.Data.DataTable();
                    if (g_Combobox_Value.Rows.Count <= 0) return new System.Data.DataTable();
                    rowArr = g_Combobox_Value.Select("U_FormUID='FPTSO' and [U_Item]='U_ItmOut1'");
                    if (rowArr.Length > 0)
                    {
                        for (int i = 0; i < rowArr.Length; i++)
                        {
                            row = g_SOTYPEList.NewRow();
                            row["Code"] = rowArr[i]["U_Code"].ToString().Trim();
                            row["NAME"] = rowArr[i]["U_Name"].ToString().Trim();
                            g_SOTYPEList.Rows.Add(row);
                        }
                    }

                    row = g_SOTYPEList.NewRow();
                    row["Code"] = "1";
                    row["NAME"] = "Thu tiền ngay";
                    g_SOTYPEList.Rows.Add(row);

                    row = g_SOTYPEList.NewRow();
                    row["Code"] = "2";
                    row["NAME"] = "Bán nợ";
                    g_SOTYPEList.Rows.Add(row);

                    row = g_SOTYPEList.NewRow();
                    row["Code"] = "3";
                    row["NAME"] = "Đặt cọc";
                    g_SOTYPEList.Rows.Add(row);

                    row = g_SOTYPEList.NewRow();
                    row["Code"] = "4";
                    row["NAME"] = "Đặt hàng";
                    g_SOTYPEList.Rows.Add(row);
                    SaveDataTableToXML(path + @"\g_SOTYPEList.xml", g_SOTYPEList);
                }
                else
                {
                    g_SOTYPEList = ConvertXmlToDataTable(path + @"\g_SOTYPEList.xml");
                }
            }
            catch
            { }
            return g_SOTYPEList;
        }
        public static System.Data.DataTable GetNoiBo()
        {
            System.Data.DataTable g_NoiBoList = new System.Data.DataTable();

            try
            {
                if (!File.Exists(path + @"\g_NoiBoList.xml"))
                {
                    System.Data.DataTable g_Combobox_Value = null;
                    System.Data.DataRow row = null;
                    System.Data.DataRow[] rowArr;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    g_NoiBoList.TableName = "NoiBo";
                    g_NoiBoList.Columns.Add("Code", Type.GetType("System.String"));
                    g_NoiBoList.Columns.Add("NAME", Type.GetType("System.String"));

                    g_Combobox_Value = GetComboBoxValue();
                    if (g_Combobox_Value == null) return new System.Data.DataTable();
                    if (g_Combobox_Value.Rows.Count <= 0) return new System.Data.DataTable();
                    rowArr = g_Combobox_Value.Select("U_FormUID='FPTSO' and [U_Item]='U_ImpType'");
                    if (rowArr.Length > 0)
                    {
                        for (int i = 0; i < rowArr.Length; i++)
                        {
                            row = g_NoiBoList.NewRow();
                            row["Code"] = rowArr[i]["U_Code"].ToString().Trim();
                            row["NAME"] = rowArr[i]["U_Name"].ToString().Trim();
                            g_NoiBoList.Rows.Add(row);
                        }
                    }
                    row = g_NoiBoList.NewRow();
                    row["Code"] = "LCNB";
                    row["NAME"] = "Có";
                    g_NoiBoList.Rows.Add(row);

                    row = g_NoiBoList.NewRow();
                    row["Code"] = "";
                    row["NAME"] = "Không";
                    g_NoiBoList.Rows.Add(row);

                    SaveDataTableToXML(path + @"\g_NoiBoList.xml", g_NoiBoList);
                }
                else
                {
                    g_NoiBoList = ConvertXmlToDataTable(path + @"\g_NoiBoList.xml");
                }
            }
            catch
            { }
            return g_NoiBoList;
        }

        public static System.Data.DataTable GetComboBoxValue()
        {
            System.Data.DataTable g_Combobox_Value = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_Combobox_Value.xml"))
                {
                    System.Data.DataRow row = null;
                    string sql = "";
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "SELECT [U_Lang],[U_FormUID],[U_Item],[U_Code] ,[U_Name] FROM [@FPTLANGCOMBO] with(nolock) ";
                    sql += " WHERE U_Lang=CONVERT(nvarchar(10), isnull((select U_Values from [@FPTSYS] with(nolock) where U_SysCode='LNS'),'EN')) ORDER BY U_FormUID";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    g_Combobox_Value.TableName = "FPTLANGCOMBO";
                    g_Combobox_Value.Columns.Add("U_FormUID", Type.GetType("System.String"));
                    g_Combobox_Value.Columns.Add("U_Item", Type.GetType("System.String"));
                    g_Combobox_Value.Columns.Add("U_Code", Type.GetType("System.String"));
                    g_Combobox_Value.Columns.Add("U_Name", Type.GetType("System.String"));

                    if (!oDataTable.IsEmpty)
                    {
                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_Combobox_Value.NewRow();
                            row["U_FormUID"] = oDataTable.GetValue("U_FormUID", i).ToString().Trim();
                            row["U_Item"] = oDataTable.GetValue("U_Item", i).ToString().Trim();
                            row["U_Code"] = oDataTable.GetValue("U_Code", i).ToString().Trim();
                            row["U_Name"] = oDataTable.GetValue("U_Name", i).ToString().Trim();
                            g_Combobox_Value.Rows.Add(row);
                        }
                    }
                    SaveDataTableToXML(path + @"\g_Combobox_Value.xml", g_Combobox_Value);
                }
                else
                {
                    g_Combobox_Value = ConvertXmlToDataTable(path + @"\g_Combobox_Value.xml");
                }
            }
            catch
            { }
            return g_Combobox_Value;
        }

        public static DataTable Get_FormItem()
        {
            DataTable Get_FormItem = new DataTable();
            try
            {
                if (!File.Exists(path + @"\Get_FormItem.xml"))
                {
                    string SQL = null;
                    System.Data.DataRow row = null;
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    SQL = "SELECT A.U_FORMUID, A.U_MNAME, U_ITEM_CODE FROM [@FPTITEMS]  A with(nolock)";
                    SAPbouiCOM.DataTable oDataTale = Globals.GetSapDataTable(SQL);
                    if (oDataTale.IsEmpty) return null;
                    Get_FormItem.TableName = "FPTITEMS";

                    Get_FormItem.Columns.Add("U_FORMUID", Type.GetType("System.String"));
                    Get_FormItem.Columns.Add("U_MNAME", Type.GetType("System.String"));
                    Get_FormItem.Columns.Add("U_ITEM_CODE", Type.GetType("System.String"));

                    for (int i = 0; i < oDataTale.Rows.Count; i++)
                    {
                        row = Get_FormItem.NewRow();
                        row["U_FORMUID"] = oDataTale.GetValue("U_FORMUID", i).ToString().Trim();
                        row["U_MNAME"] = oDataTale.GetValue("U_MNAME", i).ToString().Trim();
                        row["U_ITEM_CODE"] = oDataTale.GetValue("U_ITEM_CODE", i).ToString().Trim();
                        Get_FormItem.Rows.Add(row);
                    }
                    SaveDataTableToXML(path + @"\Get_FormItem.xml", Get_FormItem);

                }
                else
                    Get_FormItem = ConvertXmlToDataTable(path + @"\Get_FormItem.xml");
            }
            catch { }
            return Get_FormItem;
        }

        #endregion

        #region Chuyen mot so phuong thuc tu class B1Customize sang Globals

        public static System.Data.DataTable GetLocation()
        {
            System.Data.DataTable g_Location = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_Location.xml"))
                {
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    string sql = "SELECT A.ItemCode, B.WhsCode, B.U_LOCATION  FROM  OITM A, OITW B WHERE  A.ItemCode=b.ItemCode ";
                    sql += " and  B.U_LOCATION  is not null and  B.U_LOCATION <>''";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_Location.Columns.Add("ItemCode", Type.GetType("System.String"));
                        g_Location.Columns.Add("WhsCode", Type.GetType("System.String"));
                        g_Location.Columns.Add("U_LOCATION", Type.GetType("System.String"));
                        System.Data.DataRow row;
                        for (int i = 0; i <= oDataTable.Rows.Count - 1; i++)
                        {
                            row = g_Location.NewRow();
                            row["ItemCode"] = oDataTable.GetValue("ItemCode", i).ToString().Trim();
                            row["WhsCode"] = oDataTable.GetValue("WhsCode", i).ToString().Trim();
                            row["U_LOCATION"] = oDataTable.GetValue("U_LOCATION", i).ToString().Trim();
                            g_Location.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_Location.xml", g_Location);
                    }
                }
                else
                {
                    g_Location = ConvertXmlToDataTable(path + @"\g_Location.xml");
                }
            }
            catch
            {
                g_Location = null;
            }
            return g_Location;
        }
        public static System.Data.DataTable Get_CFL_Customize()
        {
            System.Data.DataTable g_OBJ_CFL = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_OBJ_CFL.xml"))
                {
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    string sql = "SELECT [U_FormUID],[U_Item_Type] ,[U_UniqueID],[U_Matrix_Name],[U_CFL],[U_OBJTYPE],[U_TABTYPE],[U_OBJALIAS] ";
                    sql += " FROM [@FPTOBJ] WHERE U_CFL='Y' and U_OBJTYPE is not null ";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_OBJ_CFL.Columns.Add("U_FormUID", Type.GetType("System.String"));
                        g_OBJ_CFL.Columns.Add("U_Item_Type", Type.GetType("System.String"));
                        g_OBJ_CFL.Columns.Add("U_UniqueID", Type.GetType("System.String"));
                        g_OBJ_CFL.Columns.Add("U_Matrix_Name", Type.GetType("System.String"));
                        g_OBJ_CFL.Columns.Add("U_CFL", Type.GetType("System.String"));
                        g_OBJ_CFL.Columns.Add("U_OBJTYPE", Type.GetType("System.String"));
                        g_OBJ_CFL.Columns.Add("U_TABTYPE", Type.GetType("System.String"));
                        g_OBJ_CFL.Columns.Add("U_OBJALIAS", Type.GetType("System.String"));
                        System.Data.DataRow row;
                        for (int i = 0; i <= oDataTable.Rows.Count - 1; i++)
                        {
                            row = g_OBJ_CFL.NewRow();
                            string s = oDataTable.GetValue("U_FormUID", i).ToString().Trim();
                            row["U_FormUID"] = oDataTable.GetValue("U_FormUID", i).ToString().Trim();
                            row["U_Item_Type"] = oDataTable.GetValue("U_Item_Type", i).ToString().Trim();
                            row["U_UniqueID"] = oDataTable.GetValue("U_UniqueID", i).ToString().Trim();
                            row["U_Matrix_Name"] = oDataTable.GetValue("U_Matrix_Name", i).ToString().Trim();
                            row["U_CFL"] = oDataTable.GetValue("U_CFL", i).ToString().Trim();
                            row["U_OBJTYPE"] = oDataTable.GetValue("U_OBJTYPE", i).ToString().Trim();
                            row["U_TABTYPE"] = oDataTable.GetValue("U_TABTYPE", i).ToString().Trim();
                            row["U_OBJALIAS"] = oDataTable.GetValue("U_OBJALIAS", i).ToString().Trim();
                            g_OBJ_CFL.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_OBJ_CFL.xml", g_OBJ_CFL);
                    }
                }
                else
                {
                    g_OBJ_CFL = ConvertXmlToDataTable(path + @"\g_OBJ_CFL.xml");
                }
            }
            catch
            {
                g_OBJ_CFL = null;
            }
            return g_OBJ_CFL;
        }
        public static System.Data.DataTable Get_FormCFL()
        {
            System.Data.DataTable g_FormCFL = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_FormCFL.xml"))
                {
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    string sql = "SELECT DISTINCT  U_FORMCFL, U_FORM  FROM [@FPTCFL] ORDER BY U_FORM";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_FormCFL.Columns.Add("U_FORMCFL", Type.GetType("System.String"));
                        g_FormCFL.Columns.Add("U_FORM", Type.GetType("System.String"));
                        System.Data.DataRow row;
                        for (int i = 0; i <= oDataTable.Rows.Count - 1; i++)
                        {
                            row = g_FormCFL.NewRow();
                            row["U_FORMCFL"] = oDataTable.GetValue("U_FORMCFL", i).ToString().Trim();
                            row["U_FORM"] = oDataTable.GetValue("U_FORM", i).ToString().Trim();
                            g_FormCFL.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_FormCFL.xml", g_FormCFL);
                    }
                }
                else
                {
                    g_FormCFL = ConvertXmlToDataTable(path + @"\g_FormCFL.xml");
                }
            }
            catch
            {
                g_FormCFL = null;
            }
            return g_FormCFL;
        }
        public static System.Data.DataTable Get_CFL()
        {
            System.Data.DataTable g_CFL = new System.Data.DataTable();
            try
            {
                if (!File.Exists(path + @"\g_CFL.xml"))
                {
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    string sql = "SELECT U_FORM,U_CODE ,U_ITYPE,U_ITEM,U_PITEM,U_IVALUE,U_MNAME, U_FORMCFL FROM [@FPTCFL] ORDER BY U_FORM";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_CFL.Columns.Add("U_FORM", Type.GetType("System.String"));
                        g_CFL.Columns.Add("U_CODE", Type.GetType("System.String"));
                        g_CFL.Columns.Add("U_ITYPE", Type.GetType("System.String"));
                        g_CFL.Columns.Add("U_PITEM", Type.GetType("System.String"));
                        g_CFL.Columns.Add("U_IVALUE", Type.GetType("System.String"));
                        g_CFL.Columns.Add("U_MNAME", Type.GetType("System.String"));
                        g_CFL.Columns.Add("U_FORMCFL", Type.GetType("System.String"));
                        g_CFL.Columns.Add("U_ITEM", Type.GetType("System.String"));

                        System.Data.DataRow row;
                        for (int i = 0; i <= oDataTable.Rows.Count - 1; i++)
                        {
                            row = g_CFL.NewRow();
                            row["U_FORM"] = oDataTable.GetValue("U_FORM", i).ToString().Trim();
                            row["U_CODE"] = oDataTable.GetValue("U_CODE", i).ToString().Trim();
                            row["U_ITYPE"] = oDataTable.GetValue("U_ITYPE", i).ToString().Trim();
                            row["U_PITEM"] = oDataTable.GetValue("U_PITEM", i).ToString().Trim();
                            row["U_IVALUE"] = oDataTable.GetValue("U_IVALUE", i).ToString().Trim();
                            row["U_MNAME"] = oDataTable.GetValue("U_MNAME", i).ToString().Trim();
                            row["U_FORMCFL"] = oDataTable.GetValue("U_FORMCFL", i).ToString().Trim();
                            row["U_ITEM"] = oDataTable.GetValue("U_ITEM", i).ToString().Trim();
                            g_CFL.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_CFL.xml", g_CFL);
                    }
                }
                else
                {
                    g_CFL = ConvertXmlToDataTable(path + @"\g_CFL.xml");
                }
            }
            catch
            {
                g_CFL = null;
            }
            return g_CFL;
        }

        public static void Folder_iTem(int From_Panel, int To_Panel, SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Item_Caption, int Panel, string U_GROUP_WITH, bool Resize = false)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);
            SAPbouiCOM.Folder Folder = default(SAPbouiCOM.Folder);


            try
            {
                try
                {
                    Item = Form.Items.Item(UniqueID);
                    Resize = true;
                }
                catch
                {
                    Resize = false;
                }


                if (Resize == false)
                {
                    Item = Form.Items.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                }
                else
                {
                    Item = Form.Items.Item(UniqueID);
                }


                if (!string.IsNullOrEmpty(From_Panel.ToString()))
                {
                    Item.FromPane = int.Parse(From_Panel.ToString());

                }
                if (!string.IsNullOrEmpty(To_Panel.ToString()))
                {
                    Item.ToPane = int.Parse(To_Panel.ToString());

                }
                Item.Visible = true;
                Item.Width = 80;
                Item.Height = 20;
                Item.Top = 10;
                Item.Left = 10;
                //  Item.
                Folder = (SAPbouiCOM.Folder)Item.Specific;
                if (!string.IsNullOrEmpty(U_GROUP_WITH.ToString()))
                {
                    Folder.GroupWith(U_GROUP_WITH);
                }
                if (Resize == true)
                    return;
                Folder.Pane = Panel;

                // Form.PaneLevel = Panel
                Folder.Caption = Item_Caption;
                Folder.AutoPaneSelection = true;

            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }

        }

        public static void Button_iTem(int[] Get_Item, int From_Panel, int To_Panel, SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Item_Caption, string Get_Item_Name = "", string Get_Item_Width = "N",
                string Get_Item_Height = "N", int Position = 1, int Space = 0, int Left = 0, int Top = 0, int Width = 100, int Height = 20, bool Resize = false)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);
            SAPbouiCOM.Button obtn = default(SAPbouiCOM.Button);
            int Top1 = 0;
            int Left1 = 0;
            int Width1 = 0;
            int Height1 = 0;

            try
            {

                try
                {
                    Item = Form.Items.Item(UniqueID);
                    Resize = true;
                }
                catch
                {
                    Resize = false;
                }

                if (Resize == false)
                {
                    Item = Form.Items.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                }
                else
                {
                    Item = Form.Items.Item(UniqueID);
                }
                Top1 = Top;
                Left1 = Left;
                Width1 = Width;
                Height1 = Height;

                if (!string.IsNullOrEmpty(Get_Item_Name))
                {
                    if (Get_Item[0] == 0)
                    {
                        Top1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "TOP"));
                    }
                    else
                    {
                        Top1 = Get_Item[0];
                    }
                    if (Get_Item[1] == 0)
                    {
                        Left1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "LEFT"));
                    }
                    else
                    {
                        Left1 = Get_Item[1];
                    }

                    if (Get_Item_Width == "Y")
                    {
                        if (Get_Item[3] == 0)
                        {
                            Width1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH"));
                        }
                        else
                        {
                            Width1 = Get_Item[3];

                        }
                    }
                    if (Get_Item_Height == "Y")
                    {
                        if (Get_Item[2] == 0)
                        {
                            Height1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT"));
                        }
                        else
                        {
                            Height1 = Get_Item[2];
                        }
                    }
                    switch (Position)
                    {
                        case 1:
                            if (Get_Item[2] == 0)
                            {
                                Top1 = Top1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT")) + Space;
                            }
                            else
                            {
                                Top1 = Top1 + Get_Item[2] + Space;
                            }
                            break;
                        case 2:
                            Top1 = Top1 - Height1 - Space;
                            break;
                        case 3:
                            Left1 = Left1 - Width1 - Space;
                            break;
                        case 4:
                            if (Get_Item[3] == 0)
                            {
                                Left1 = Left1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH")) + Space;
                            }
                            else
                            {
                                Left1 = Left1 + Get_Item[3] + Space;
                            }
                            break;
                    }

                }

                Item.Left = Left1;
                Item.Top = Top1;
                Item.Width = Width1;
                Item.Height = Height1;

                if (!string.IsNullOrEmpty(From_Panel.ToString()))
                {
                    Item.FromPane = From_Panel;

                }
                if (!string.IsNullOrEmpty(To_Panel.ToString()))
                {
                    Item.ToPane = To_Panel;

                }

                obtn = (SAPbouiCOM.Button)Item.Specific;


                if (Resize == true)
                    return;

                obtn.Caption = Item_Caption;


            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }

        }

        public static string Get_Item_Property(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, string Item_Name, string Type)
        {
            //Dim Form As SAPbouiCOM.Form
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);
            string Return = null;

            SAPbouiCOM.Items Items = default(SAPbouiCOM.Items);
            int Count = 0;


            Return = "0";

            // Form = SBO_Application.Forms.Item(FormUID)
            Items = Form.Items;
            //Item = Form.Items.Item(Item_Name)

            for (Count = 0; Count <= Items.Count - 1; Count++)
            {
                Item = Items.Item(Count);
                if (Item_Name.ToUpper() == Item.UniqueID.ToUpper())
                {
                    Item = Form.Items.Item(Item_Name);
                    switch (Type.ToUpper())
                    {
                        case "TOP":
                            Return = Item.Top.ToString();
                            return Return;
                        case "HEIGHT":
                            Return = Item.Height.ToString();
                            return Return;
                        case "WIDTH":
                            Return = Item.Width.ToString();
                            return Return;
                        case "LEFT":
                            Return = Item.Left.ToString();
                            return Return;
                        case "TYPE":
                            Return = Item.Type.ToString();
                            return Return;
                        case "UNIQUEID":
                            Return = Item.UniqueID.ToString();
                            return Return;
                        case "VISIBLE":
                            if (Item.Visible == true)
                            {
                                Return = "TRUE";
                            }
                            else
                            {
                                Return = "FALSE";
                            }
                            return Return;
                        case "ENABLE":
                            if (Item.Enabled == true)
                            {
                                Return = "TRUE";
                            }
                            else
                            {
                                Return = "FALSE";
                            }

                            return Return;
                    }
                }
            }
            return Return;
        }

        public static void LinkButton_iTem(int[] Get_Item, int From_Panel, int To_Panel, SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Link_Item_To, string LinkedObjectType = "", int LinkedObject = 0,
            int Left = 0, int Top = 0, string Get_Item_Name = "", string Get_Item_Width = "N", string Get_Item_Height = "N", int Position = 1, int Space = 0, int Width = 30, int Height = 80)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);

            SAPbouiCOM.LinkedButton LinkedButton = default(SAPbouiCOM.LinkedButton);

            int Top1 = 0;
            int Left1 = 0;
            int Width1 = 0;
            int Height1 = 0;
            //SAPbouiCOM.Item Item_Value = null;
            try
            {
                if (Check_Item(SBO_Application, Form, pVal, UniqueID) == true)
                {
                    return;
                }
                Top1 = Top;
                Left1 = Left;
                Width1 = Width;
                Height1 = Height;
                Item = Form.Items.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                if (!string.IsNullOrEmpty(Get_Item_Name))
                {

                    if (Get_Item[0] == 0)
                    {
                        Top1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "TOP"));
                    }
                    else
                    {
                        Top1 = Get_Item[0];
                    }
                    if (Get_Item[1] == 0)
                    {
                        Left1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "LEFT"));
                    }
                    else
                    {
                        Left1 = Get_Item[1];
                    }

                    if (Get_Item_Width == "Y")
                    {
                        if (Get_Item[3] == 0)
                        {
                            Width1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH"));
                        }
                        else
                        {
                            Width1 = Get_Item[3];

                        }
                    }
                    if (Get_Item_Height == "Y")
                    {
                        if (Get_Item[2] == 0)
                        {
                            Height1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT"));
                        }
                        else
                        {
                            Height1 = Get_Item[2];
                        }
                    }
                    switch (Position)
                    {
                        case 1:
                            if (Get_Item[2] == 0)
                            {
                                Top1 = Top1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT")) + Space;
                            }
                            else
                            {
                                Top1 = Top1 + Get_Item[2] + Space;
                            }
                            break;
                        case 2:
                            Top1 = Top1 - Height1 - Space;
                            break;
                        case 3:
                            Left1 = Left1 - Width1 - Space;
                            break;
                        case 4:
                            if (Get_Item[3] == 0)
                            {
                                Left1 = Left1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH")) + Space;
                            }
                            else
                            {
                                Left1 = Left1 + Get_Item[3] + Space;
                            }
                            break;
                    }
                }
                Item.Left = Left1;
                Item.Top = Top1;
                Item.Width = Width1;
                Item.Height = Height1;
                if (!string.IsNullOrEmpty(From_Panel.ToString()))
                {
                    Item.FromPane = From_Panel;

                }
                if (!string.IsNullOrEmpty(To_Panel.ToString()))
                {
                    Item.ToPane = To_Panel;

                }
                Item.LinkTo = Link_Item_To;
                LinkedButton = (SAPbouiCOM.LinkedButton)Item.Specific;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }



        }

        private static bool Check_Item(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string Item_Name)
        {
            //Dim Form As SAPbouiCOM.Form
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);
            SAPbouiCOM.Items Items = default(SAPbouiCOM.Items);
            int Count = 0;
            bool Return = false;
            Return = false;
            if (string.IsNullOrEmpty(Item_Name)) return true;
            // Form = SBO_Application.Forms.Item(FormUID)
            Items = Form.Items;
            for (Count = 0; Count <= Items.Count - 1; Count++)
            {
                Item = Items.Item(Count);
                if (Item_Name.ToUpper() == Item.UniqueID.ToUpper())
                {
                    //SBO_Application.MessageBox("Tong so item la : " & Item.UniqueID)
                    Return = true;
                }
            }
            return Return;
        }

        public static void Label_iTem(int[] Get_Item, int From_Panel, int To_Panel, SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Item_Caption, int Left = 0, int Top = 0,
            string Get_Item_Name = "", string Get_Item_Width = "N", string Get_Item_Height = "N", int Position = 1, int Space = 0, int Width = 100, int Height = 20, bool Resize = false)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);
            SAPbouiCOM.StaticText Static = default(SAPbouiCOM.StaticText);
            int Top1 = 0;
            int Left1 = 0;
            int Width1 = 0;
            int Height1 = 0;

            try
            {
                Top1 = Top;
                Left1 = Left;
                Width1 = Width;
                Height1 = Height;

                try
                {
                    Item = Form.Items.Item(UniqueID);
                    Resize = true;
                }
                catch
                {
                    Resize = false;
                }


                if (Resize == false)
                {
                    Item = Form.Items.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_STATIC);
                }
                else
                {
                    Item = Form.Items.Item(UniqueID);
                }
                if (!string.IsNullOrEmpty(Get_Item_Name))
                {
                    if (Get_Item[0] == 0)
                    {
                        Top1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "TOP"));
                    }
                    else
                    {
                        Top1 = Get_Item[0];
                    }
                    if (Get_Item[1] == 0)
                    {
                        Left1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "LEFT"));
                    }
                    else
                    {
                        Left1 = Get_Item[1];
                    }

                    if (Get_Item_Width == "Y")
                    {
                        if (Get_Item[3] == 0)
                        {
                            Width1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH"));
                        }
                        else
                        {
                            Width1 = Get_Item[3];

                        }
                    }
                    if (Get_Item_Height == "Y")
                    {
                        if (Get_Item[2] == 0)
                        {
                            Height1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT"));
                        }
                        else
                        {
                            Height1 = Get_Item[2];
                        }
                    }
                    switch (Position)
                    {
                        case 1:
                            if (Get_Item[2] == 0)
                            {
                                Top1 = Top1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT")) + Space;
                            }
                            else
                            {
                                Top1 = Top1 + Get_Item[2] + Space;
                            }
                            break;
                        case 2:
                            Top1 = Top1 - Height1 - Space;
                            break;
                        case 3:
                            Left1 = Left1 - Width1 - Space;
                            break;
                        case 4:
                            if (Get_Item[3] == 0)
                            {
                                Left1 = Left1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH")) + Space;
                            }
                            else
                            {
                                Left1 = Left1 + Get_Item[3] + Space;
                            }
                            break;
                    }
                }



                Item.Left = Left1;
                Item.Top = Top1;
                Item.Width = Width1;
                Item.Height = Height1;
                if (!string.IsNullOrEmpty(From_Panel.ToString()))
                {
                    Item.FromPane = From_Panel;

                }
                if (!string.IsNullOrEmpty(To_Panel.ToString()))
                {
                    Item.ToPane = To_Panel;

                }
                Static = (SAPbouiCOM.StaticText)Item.Specific;
                if (Resize == true)
                    return;
                Static.Caption = Item_Caption;

            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }

        }

        public static void ComboBox_iTem(int[] Get_Item, int From_Panel, int To_Panel, SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Item_Caption, string Table_Name, string Field_Name,
            string SQL_Source, int Left = 0, int Top = 0, string Get_Item_Name = "", string Get_Item_Width = "N", string Get_Item_Height = "N", int Position = 1, int Space = 0, string Data_Bind = "N", int Width = 100,
            int Height = 20, bool Resize = false)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);

            SAPbouiCOM.ComboBox ComboBox = default(SAPbouiCOM.ComboBox);

            int Top1 = 0;
            int Left1 = 0;
            int Width1 = 0;
            int Height1 = 0;

            try
            {
                Top1 = Top;
                Left1 = Left;
                Width1 = Width;
                Height1 = Height;

                try
                {
                    Item = Form.Items.Item(UniqueID);
                    Resize = true;
                }
                catch
                {
                    Resize = false;
                }


                if (Resize == true)
                {
                    Item = Form.Items.Item(UniqueID);
                }
                else
                {
                    Item = Form.Items.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                }


                if (!string.IsNullOrEmpty(Get_Item_Name))
                {
                    if (Get_Item[0] == 0)
                    {
                        Top1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "TOP"));
                    }
                    else
                    {
                        Top1 = Get_Item[0];
                    }
                    if (Get_Item[1] == 0)
                    {
                        Left1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "LEFT"));
                    }
                    else
                    {
                        Left1 = Get_Item[1];
                    }

                    if (Get_Item_Width == "Y")
                    {
                        if (Get_Item[3] == 0)
                        {
                            Width1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH"));
                        }
                        else
                        {
                            Width1 = Get_Item[3];

                        }
                    }
                    if (Get_Item_Height == "Y")
                    {
                        if (Get_Item[2] == 0)
                        {
                            Height1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT"));
                        }
                        else
                        {
                            Height1 = Get_Item[2];
                        }
                    }
                    switch (Position)
                    {
                        case 1:
                            if (Get_Item[2] == 0)
                            {
                                Top1 = Top1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT")) + Space;
                            }
                            else
                            {
                                Top1 = Top1 + Get_Item[2] + Space;
                            }
                            break;
                        case 2:
                            Top1 = Top1 - Height1 - Space;
                            break;
                        case 3:
                            Left1 = Left1 - Width1 - Space;
                            break;
                        case 4:
                            if (Get_Item[3] == 0)
                            {
                                Left1 = Left1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH")) + Space;
                            }
                            else
                            {
                                Left1 = Left1 + Get_Item[3] + Space;
                            }
                            break;
                    }
                }
                Item.Left = Left1;
                Item.Top = Top1;
                Item.Width = Width1;
                Item.Height = Height1;

                if (!string.IsNullOrEmpty(From_Panel.ToString()))
                {
                    Item.FromPane = From_Panel;

                }
                if (!string.IsNullOrEmpty(To_Panel.ToString()))
                {
                    Item.ToPane = To_Panel;

                }
                Item.DisplayDesc = true;
                ComboBox = (SAPbouiCOM.ComboBox)Item.Specific;

                if (Resize == true)
                    return;

                if (Data_Bind == "Y")
                {
                    ComboBox.DataBind.SetBound(true, Table_Name, Field_Name);
                }
                else
                {
                    ComboBox.DataBind.SetBound(false, Table_Name, Field_Name);
                }
                ComboBox_ValidValues_Add(ComboBox, ref SQL_Source);
            }
            catch
            {

            }
        }

        public static void ComboBox_ValidValues_Add(SAPbouiCOM.ComboBox p_ComboBox, ref string p_SQL)
        {
            SAPbouiCOM.DataTable oDataTable = GetSapDataTable(p_SQL);
            int p_Count;
            if (!oDataTable.IsEmpty)
            {
                for (p_Count = 0; p_Count <= oDataTable.Rows.Count - 1; p_Count++)
                {
                    if (oDataTable.Columns.Count > 1)
                        p_ComboBox.ValidValues.Add(oDataTable.GetValue(0, p_Count).ToString().Trim(), oDataTable.GetValue(1, p_Count).ToString().Trim());
                    else
                        p_ComboBox.ValidValues.Add(oDataTable.GetValue(0, p_Count).ToString().Trim(), oDataTable.GetValue(0, p_Count).ToString().Trim());
                }
            }

        }

        public static void CheckBox_iTem(int[] Get_Item, int From_Panel, int To_Panel, SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Item_Caption, string Table_Name, string Field_Name,
            int Left = 0, int Top = 0, string Get_Item_Name = "", string Get_Item_Width = "N", string Get_Item_Height = "N", int Position = 1, int Space = 0, string Data_Bind = "N", int Width = 100, int Height = 20,
            bool Resize = false)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);

            SAPbouiCOM.CheckBox Option = default(SAPbouiCOM.CheckBox);

            int Top1 = 0;
            int Left1 = 0;
            int Width1 = 0;
            int Height1 = 0;

            try
            {
                Top1 = Top;
                Left1 = Left;
                Width1 = Width;
                Height1 = Height;

                try
                {
                    Item = Form.Items.Item(UniqueID);
                    Resize = true;
                }
                catch
                {
                    Resize = false;
                }

                if (Resize == false)
                {
                    Item = Form.Items.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                }
                else
                {
                    Item = Form.Items.Item(UniqueID);
                }

                if (!string.IsNullOrEmpty(Get_Item_Name))
                {
                    if (Get_Item[0] == 0)
                    {
                        Top1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "TOP"));
                    }
                    else
                    {
                        Top1 = Get_Item[0];
                    }
                    if (Get_Item[1] == 0)
                    {
                        Left1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "LEFT"));
                    }
                    else
                    {
                        Left1 = Get_Item[1];
                    }

                    if (Get_Item_Width == "Y")
                    {
                        if (Get_Item[3] == 0)
                        {
                            Width1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH"));
                        }
                        else
                        {
                            Width1 = Get_Item[3];

                        }
                    }
                    if (Get_Item_Height == "Y")
                    {
                        if (Get_Item[2] == 0)
                        {
                            Height1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT"));
                        }
                        else
                        {
                            Height1 = Get_Item[2];
                        }
                    }
                    switch (Position)
                    {
                        case 1:
                            if (Get_Item[2] == 0)
                            {
                                Top1 = Top1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT")) + Space;
                            }
                            else
                            {
                                Top1 = Top1 + Get_Item[2] + Space;
                            }
                            break;
                        case 2:
                            Top1 = Top1 - Height1 - Space;
                            break;
                        case 3:
                            Left1 = Left1 - Width1 - Space;
                            break;
                        case 4:
                            if (Get_Item[3] == 0)
                            {
                                Left1 = Left1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH")) + Space;
                            }
                            else
                            {
                                Left1 = Left1 + Get_Item[3] + Space;
                            }
                            break;
                    }
                }
                Item.Left = Left1;
                Item.Top = Top1;
                Item.Width = Width1;
                Item.Height = Height1;
                if (!string.IsNullOrEmpty(From_Panel.ToString()))
                {
                    Item.FromPane = From_Panel;

                }
                if (!string.IsNullOrEmpty(To_Panel.ToString()))
                {
                    Item.ToPane = To_Panel;

                }
                Item.DisplayDesc = true;
                Option = (SAPbouiCOM.CheckBox)Item.Specific;
                if (Resize == true)
                    return;
                Option.Caption = Item_Caption;
                if (Data_Bind == "Y")
                {
                    Option.DataBind.SetBound(true, Table_Name, Field_Name);
                }
                else
                {
                    Option.DataBind.SetBound(false, Table_Name, Field_Name);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public static void Option_iTem(int[] Get_Item, int From_Panel, int To_Panel, SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Item_Caption, string Table_Name, string Field_Name,
            int Left = 0, int Top = 0, string Get_Item_Name = "", string Get_Item_Width = "N", string Get_Item_Height = "N", int Position = 1, int Space = 0, string Data_Bind = "N", int Width = 100, int Height = 20,
            bool Resize = false)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);

            SAPbouiCOM.OptionBtn Option = default(SAPbouiCOM.OptionBtn);

            int Top1 = 0;
            int Left1 = 0;
            int Width1 = 0;
            int Height1 = 0;

            try
            {
                Top1 = Top;
                Left1 = Left;
                Width1 = Width;
                Height1 = Height;

                try
                {
                    Item = Form.Items.Item(UniqueID);
                    Resize = true;
                }
                catch
                {
                    Resize = false;
                }


                if (Resize == false)
                {
                    Item = Form.Items.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                }
                else
                {
                    Item = Form.Items.Item(UniqueID);
                }

                if (!string.IsNullOrEmpty(Get_Item_Name))
                {
                    if (Get_Item[0] == 0)
                    {
                        Top1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "TOP"));
                    }
                    else
                    {
                        Top1 = Get_Item[0];
                    }
                    if (Get_Item[1] == 0)
                    {
                        Left1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "LEFT"));
                    }
                    else
                    {
                        Left1 = Get_Item[1];
                    }

                    if (Get_Item_Width == "Y")
                    {
                        if (Get_Item[3] == 0)
                        {
                            Width1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH"));
                        }
                        else
                        {
                            Width1 = Get_Item[3];

                        }
                    }
                    if (Get_Item_Height == "Y")
                    {
                        if (Get_Item[2] == 0)
                        {
                            Height1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT"));
                        }
                        else
                        {
                            Height1 = Get_Item[2];
                        }
                    }
                    switch (Position)
                    {
                        case 1:
                            if (Get_Item[2] == 0)
                            {
                                Top1 = Top1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT")) + Space;
                            }
                            else
                            {
                                Top1 = Top1 + Get_Item[2] + Space;
                            }
                            break;
                        case 2:
                            Top1 = Top1 - Height1 - Space;
                            break;
                        case 3:
                            Left1 = Left1 - Width1 - Space;
                            break;
                        case 4:
                            if (Get_Item[3] == 0)
                            {
                                Left1 = Left1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH")) + Space;
                            }
                            else
                            {
                                Left1 = Left1 + Get_Item[3] + Space;
                            }
                            break;
                    }
                }



                Item.Left = Left1;
                Item.Top = Top1;
                Item.Width = Width1;
                Item.Height = Height1;
                if (!string.IsNullOrEmpty(From_Panel.ToString()))
                {
                    Item.FromPane = From_Panel;

                }
                if (!string.IsNullOrEmpty(To_Panel.ToString()))
                {
                    Item.ToPane = To_Panel;

                }

                Item.DisplayDesc = true;
                Option = (SAPbouiCOM.OptionBtn)Item.Specific;
                if (Resize == true)
                    return;


                Option.Caption = Item_Caption;
                //Option.DataBind.SetBound(Data_Bind, Table_Name, Field_Name)
                if (Data_Bind == "Y")
                {
                    Option.DataBind.SetBound(true, Table_Name, Field_Name);
                }
                else
                {
                    Option.DataBind.SetBound(false, Table_Name, Field_Name);
                }


            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }

        }

        public static void Text_iTem(int[] Get_Item, int From_Panel, int To_Panel, SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Item_Caption, string Table_Name, string Field_Name,
            int Left = 0, int Top = 0, string Get_Item_Name = "", string Get_Item_Width = "N", string Get_Item_Height = "N", int Position = 1, int Space = 0, string Data_Bind = "N", int Width = 100, int Height = 20,
            bool Resize = false)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);

            SAPbouiCOM.EditText TextBox = default(SAPbouiCOM.EditText);

            int Top1 = 0;
            int Left1 = 0;
            int Width1 = 0;
            int Height1 = 0;

            try
            {
                Top1 = Top;
                Left1 = Left;
                Width1 = Width;
                Height1 = Height;
                try
                {
                    Item = Form.Items.Item(UniqueID);
                    Resize = true;
                }
                catch
                {
                    Resize = false;
                }
                if (Resize == false)
                {
                    Item = Form.Items.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                }
                else
                {
                    Item = Form.Items.Item(UniqueID);
                }
                if (!string.IsNullOrEmpty(Get_Item_Name))
                {
                    if (Get_Item[0] == 0)
                    {
                        Top1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "TOP"));
                    }
                    else
                    {
                        Top1 = Get_Item[0];
                    }
                    if (Get_Item[1] == 0)
                    {
                        Left1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "LEFT"));
                    }
                    else
                    {
                        Left1 = Get_Item[1];
                    }

                    if (Get_Item_Width == "Y")
                    {
                        if (Get_Item[3] == 0)
                        {
                            Width1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH"));
                        }
                        else
                        {
                            Width1 = Get_Item[3];

                        }
                    }
                    if (Get_Item_Height == "Y")
                    {
                        if (Get_Item[2] == 0)
                        {
                            Height1 = int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT"));
                        }
                        else
                        {
                            Height1 = Get_Item[2];
                        }
                    }
                    switch (Position)
                    {
                        case 1:
                            if (Get_Item[2] == 0)
                            {
                                Top1 = Top1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "HEIGHT")) + Space;
                            }
                            else
                            {
                                Top1 = Top1 + Get_Item[2] + Space;
                            }
                            break;
                        case 2:
                            Top1 = Top1 - Height1 - Space;
                            break;
                        case 3:
                            Left1 = Left1 - Width1 - Space;
                            break;
                        case 4:
                            if (Get_Item[3] == 0)
                            {
                                Left1 = Left1 + int.Parse(Get_Item_Property(SBO_Application, Form, Get_Item_Name, "WIDTH")) + Space;
                            }
                            else
                            {
                                Left1 = Left1 + Get_Item[3] + Space;
                            }
                            break;
                    }
                }
                Item.Left = Left1;
                Item.Top = Top1;
                Item.Width = Width1;
                Item.Height = Height1;
                if (!string.IsNullOrEmpty(From_Panel.ToString()))
                {
                    Item.FromPane = From_Panel;

                }
                if (!string.IsNullOrEmpty(To_Panel.ToString()))
                {
                    Item.ToPane = To_Panel;

                }

                //Item.DisplayDesc = True
                TextBox = (SAPbouiCOM.EditText)Item.Specific;

                if (Resize == true)
                    return;

                if (Data_Bind == "Y" & !string.IsNullOrEmpty(Table_Name.ToString()) & !string.IsNullOrEmpty(Field_Name.ToString()))
                {
                    TextBox.DataBind.SetBound(true, Table_Name, Field_Name);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public static void Matrix_Col(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Item_Type, string Item_Caption, string Matrix_Name, string Table_Name, string Field_Name, string Data_Bind = "N",
             string Combo_Sql = "", string Link_Item_To = "", string LinkedObjectType = "", int LinkedObject = 0, int Width = 100)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);

            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.Column Column = default(SAPbouiCOM.Column);
            SAPbouiCOM.Columns Columns = default(SAPbouiCOM.Columns);
            try
            {
                Item = Form.Items.Item(Matrix_Name);
                Matrix = (SAPbouiCOM.Matrix)Item.Specific;
                Columns = Matrix.Columns;
                switch (Item_Type)
                {
                    case "16":
                        Column = Columns.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        Column.TitleObject.Caption = Item_Caption;
                        Column.Width = Width;
                        if (Data_Bind == "Y")
                        {
                            Column.DataBind.SetBound(true, Table_Name, Field_Name);
                        }
                        else
                        {
                            Column.DataBind.SetBound(false, Table_Name, Field_Name);
                        }
                        break;
                    case "121":
                        Column = Columns.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                        Column.TitleObject.Caption = Item_Caption;
                        Column.Width = Width;
                        if (Data_Bind == "Y")
                        {
                            Column.DataBind.SetBound(true, Table_Name, Field_Name);
                        }
                        else
                        {
                            Column.DataBind.SetBound(false, Table_Name, Field_Name);
                        }
                        break;
                    case "113":
                        Column = Columns.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                        Column.TitleObject.Caption = Item_Caption;
                        Column.Width = Width;
                        if (Data_Bind == "Y")
                        {
                            Column.DataBind.SetBound(true, Table_Name, Field_Name);
                        }
                        else
                        {
                            Column.DataBind.SetBound(false, Table_Name, Field_Name);
                        }
                        Matrix_ComboBox_ValidValues_Add(ref Column, Combo_Sql);
                        break;
                    case "114":
                        Column = Columns.Add(UniqueID, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        Column.Width = Width;
                        if (Data_Bind == "Y")
                        {
                            Column.DataBind.SetBound(true, Table_Name, Field_Name);
                        }
                        else
                        {
                            Column.DataBind.SetBound(false, Table_Name, Field_Name);
                        }
                        Matrix_ComboBox_ValidValues_Add(ref Column, Combo_Sql);
                        break;
                    case "116":
                        break;
                }
            }
            catch
            { }

        }

        public static void Matrix_B1_Col(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string UniqueID, string Item_Type, string Item_Caption, string Matrix_Name, string Table_Name, string Field_Name, string Data_Bind = "N",
             string Combo_Sql = "", string Link_Item_To = "", string LinkedObjectType = "", int LinkedObject = 0, int Width = 100)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);

            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.Column Column = default(SAPbouiCOM.Column);
            SAPbouiCOM.Columns Columns = default(SAPbouiCOM.Columns);
            try
            {
                Item = Form.Items.Item(Matrix_Name);
                Matrix = (SAPbouiCOM.Matrix)Item.Specific;
                if (Check_Column(Matrix, UniqueID) == false)
                {
                    return;
                }
                Columns = Matrix.Columns;
                Column = Matrix.Columns.Item(UniqueID);
                switch (Column.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                        Column.TitleObject.Caption = Item_Caption;
                        Column.Width = Width;
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                        Column.TitleObject.Caption = Item_Caption;
                        Column.Width = Width;
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        Column.TitleObject.Caption = Item_Caption;
                        Column.Width = Width;
                        Matrix_ComboBox_ValidValues_Add(ref Column, Combo_Sql);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }

        }

        public static bool Check_Column(SAPbouiCOM.Matrix Matrix, string Col_Name)
        {
            bool functionReturnValue = false;
            int Count = 0;
            functionReturnValue = false;
            for (Count = 0; Count <= Matrix.Columns.Count - 1; Count++)
            {
                if (Matrix.Columns.Item(Count).UniqueID.ToUpper() == Col_Name.ToUpper())
                {
                    functionReturnValue = true;
                }
            }
            return functionReturnValue;
        }
        public static void Loc_Line_Matrix_AddRow(SAPbouiCOM.Form Form, ref SAPbouiCOM.Matrix Matrix01)
        {
            SAPbouiCOM.EditText EditText = default(SAPbouiCOM.EditText);
            SAPbouiCOM.ComboBox Combobox = default(SAPbouiCOM.ComboBox);
            SAPbouiCOM.DBDataSource DBSource = default(SAPbouiCOM.DBDataSource);
            try
            {
                Matrix01.FlushToDataSource();
                DBSource = Form.DataSources.DBDataSources.Item("@TCLOC");
                DBSource.InsertRecord(DBSource.Size);
                Matrix01.LoadFromDataSource();

                EditText = (SAPbouiCOM.EditText)Matrix01.Columns.Item("V_0").Cells.Item(Matrix01.RowCount).Specific;
                EditText.Value = "I";

                Combobox = (SAPbouiCOM.ComboBox)Matrix01.Columns.Item("V_2").Cells.Item(Matrix01.RowCount).Specific;
                Combobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);


                Matrix01.Columns.Item("V_5").Cells.Item(Matrix01.RowCount).Click();

            }
            catch
            {
            }

        }
        public static void LocationFilDataMatrix(bool Requery, SAPbouiCOM.Application SBO_Application, ref SAPbouiCOM.Form Form, ref SAPbouiCOM.Matrix Matrix01)
        {
            SAPbouiCOM.Column Col = default(SAPbouiCOM.Column);
            SAPbouiCOM.DBDataSource DBDataSource = default(SAPbouiCOM.DBDataSource);
            SAPbouiCOM.ComboBox ComboBox = default(SAPbouiCOM.ComboBox);
            SAPbouiCOM.EditText pEditText = default(SAPbouiCOM.EditText);

            try
            {
                if (Requery == false)
                {
                    DBDataSource = Form.DataSources.DBDataSources.Add("@TCLOC");
                }
                else
                {
                    DBDataSource = Form.DataSources.DBDataSources.Item("@TCLOC");
                }

                Col = Matrix01.Columns.Item("V_6");
                Col.DataBind.SetBound(true, "@TCLOC", "DocEntry");
                Col.Description = "DocEntry";
                Col.Visible = false;

                Col = Matrix01.Columns.Item("V_5");
                Col.DataBind.SetBound(true, "@TCLOC", "U_WHSCODE");
                Col.Description = "U_WHSCODE";
                Get_Whs(ref Col);


                Col = Matrix01.Columns.Item("V_4");
                Col.DataBind.SetBound(true, "@TCLOC", "U_CODE");
                Col.Description = "U_CODE";
                //Col.Visible = False
                Col = Matrix01.Columns.Item("V_3");
                Col.DataBind.SetBound(true, "@TCLOC", "U_NAME");
                Col.Description = "U_NAME";
                //Col.Editable = True

                Col = Matrix01.Columns.Item("V_2");
                Col.DisplayDesc = true;
                if (Col.ValidValues.Count > 0)
                {
                    //ComboBox = Matrix01.Columns.Item("V_5").ce

                }
                else
                {
                    Col.ValidValues.Add("Y", "Sử dụng");
                    Col.ValidValues.Add("N", "Không sử dụng");
                    Col.DataBind.SetBound(true, "@TCLOC", "U_ENABLE");
                }
                Col.Description = "U_ENABLE";

                Col = Matrix01.Columns.Item("V_1");
                Col.DataBind.SetBound(true, "@TCLOC", "U_DESC");
                Col.Description = "U_DESC";

                Col = Matrix01.Columns.Item("V_0");
                Col.DataBind.SetBound(true, "@TCLOC", "U_RTYPE");
                Col.Description = "U_RTYPE";
                Col.Visible = false;

                DBDataSource.Query();
                DBDataSource.InsertRecord(DBDataSource.Size);
                //DBDataSource.SetValue("U_RTYPE", DBDataSource.Size, "I")

                //DBDataSource.SetValue("STYPE", DBDataSource.Size, "I")
                // DBDataSource.SetValue("U_ENABLE", DBDataSource.Size, "Y")

                Matrix01.LoadFromDataSource();
                Matrix01.AutoResizeColumns();
                if (Matrix01.RowCount == 0)
                {
                    Matrix01.AddRow();
                    //Maker_Line_Matrix_AddRow(Matrix01)
                }
                pEditText = (SAPbouiCOM.EditText)Matrix01.Columns.Item("V_0").Cells.Item(Matrix01.RowCount).Specific;
                pEditText.Value = "I";
                ComboBox = (SAPbouiCOM.ComboBox)Matrix01.Columns.Item("V_2").Cells.Item(Matrix01.RowCount).Specific;
                ComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                Matrix01.SetCellFocus(Matrix01.RowCount, 2);

                //Matrix01.Columns.Item("V_4").Cells(Matrix01.GetCellFocus.rowIndex).

                Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                //End If
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }
        private static void Get_Whs(ref SAPbouiCOM.Column Column)
        {
            string SQL = null;
            SAPbouiCOM.DataTable oDataTable = null;
            if (Column.ValidValues.Count > 0)
            {
                return;
            }

            SQL = "select WhsCode,WhsName from OWHS  order by WhsCode";
            oDataTable = Globals.GetSapDataTable(SQL);
            if (!oDataTable.IsEmpty)
            {
                for (int i = 0; i < oDataTable.Rows.Count; i++)
                {
                    Column.ValidValues.Add(oDataTable.GetValue(0, 0).ToString().Trim(), oDataTable.GetValue(1, 0).ToString().Trim());
                }
            }

        }
        public static void LocMatrixToDatabase(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, ref SAPbouiCOM.Matrix Matrix02)
        {
            int Row = 0;
            //int Col = 0;
            int NextEntry = 0;
            SAPbouiCOM.EditText EditText = default(SAPbouiCOM.EditText);
            SAPbouiCOM.DBDataSource DBDataSource = default(SAPbouiCOM.DBDataSource);
            string SQL = null;


            try
            {
                Matrix02.FlushToDataSource();
                DBDataSource = Form.DataSources.DBDataSources.Item("@TCLOC");
                for (Row = 1; Row <= Matrix02.RowCount; Row++)
                {
                    EditText = (SAPbouiCOM.EditText)Matrix02.Columns.Item("V_0").Cells.Item(Row).Specific;

                    //INSERT 
                    if (EditText.Value.ToString() == "I" & !string.IsNullOrEmpty(DBDataSource.GetValue("U_CODE", Row - 1).ToString().Trim()))
                    {
                        NextEntry = LocNextVal();
                        SQL = "INSERT INTO [@TCLOC] (DocEntry,U_CODE,U_WHSCODE, U_NAME,U_ENABLE,U_DESC) " + " VALUES(" + NextEntry + ", N'" + DBDataSource.GetValue("U_CODE", Row - 1).ToString().Trim() + "',N'" + DBDataSource.GetValue("U_WHSCODE", Row - 1).ToString().Trim() + "', N'" + DBDataSource.GetValue("U_NAME", Row - 1).ToString().Trim() + "', '" + DBDataSource.GetValue("U_ENABLE", Row - 1).ToString().Trim() + "',N'" + DBDataSource.GetValue("U_DESC", Row - 1).ToString().Trim() + "')";
                        // ()
                        if (Globals.ExcuteQuery(SQL) == false)
                        {
                            SBO_Application.StatusBar.SetText("Lỗi khi Insert dữ liệu");
                            return;
                        }
                        // Exit Sub
                    }
                    // "U_WHSCODE='" & Trim(DBDataSource.GetValue("U_WHSCODE", Row - 1).ToString) & "'," & _
                    //UPDATE 
                    if (EditText.Value.ToString() == "U" & !string.IsNullOrEmpty(DBDataSource.GetValue("U_CODE", Row - 1).ToString()))
                    {
                        SQL = "UPDATE [@TCLOC] SET U_NAME=N'" + DBDataSource.GetValue("U_NAME", Row - 1).ToString().Trim() + "', " + "U_ENABLE='" + DBDataSource.GetValue("U_ENABLE", Row - 1).ToString().Trim() + "'," + " U_DESC =N'" + DBDataSource.GetValue("U_DESC", Row - 1).ToString().Trim() + "' " + " WHERE DocEntry=" + DBDataSource.GetValue("DocEntry", Row - 1).Trim();
                        if (Globals.ExcuteQuery(SQL) == false)
                        {
                            SBO_Application.StatusBar.SetText("Lỗi khi Update dữ liệu");
                            return;
                        }
                        // Exit Sub
                    }

                }

                //DBDataSource.Query()
                //Matrix02.LoadFromDataSource()
                //Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                LocationFilDataMatrix(true, SBO_Application, ref Form, ref Matrix02);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Message:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }


        }
        private static int LocNextVal()
        {
            int functionReturnValue = 0;
            SAPbouiCOM.DataTable oDataTable = null;
            string SQL = "select ISNULL( MAX(docEntry) ,0) +1 as DocEntry from [@TCLOC]";
            try
            {
                functionReturnValue = 1;
                oDataTable = Globals.GetSapDataTable(SQL);
                if (!oDataTable.IsEmpty)
                {
                    functionReturnValue = int.Parse(oDataTable.GetValue(0, 0).ToString().Trim());
                }

            }
            catch
            {
            }
            return functionReturnValue;


        }
        public static void Matrix_AddRow(ref SAPbouiCOM.Matrix Matrix01, string ColName)
        {
            SAPbouiCOM.EditText EditText = default(SAPbouiCOM.EditText);
            SAPbouiCOM.ComboBox Combobox = default(SAPbouiCOM.ComboBox);
            try
            {
                Matrix01.AddRow();
                EditText = (SAPbouiCOM.EditText)Matrix01.Columns.Item("STYPE").Cells.Item(Matrix01.RowCount).Specific;
                EditText.Value = "I";
                EditText = (SAPbouiCOM.EditText)Matrix01.Columns.Item("V_4").Cells.Item(Matrix01.RowCount).Specific;
                EditText.Value = "";
                EditText = (SAPbouiCOM.EditText)Matrix01.Columns.Item("V_0").Cells.Item(Matrix01.RowCount).Specific;
                EditText.Value = "";
                Combobox = (SAPbouiCOM.ComboBox)Matrix01.Columns.Item("V_1").Cells.Item(Matrix01.RowCount).Specific;
                Combobox.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                EditText = (SAPbouiCOM.EditText)Matrix01.Columns.Item("V_2").Cells.Item(Matrix01.RowCount).Specific;
                EditText.Value = "";
                EditText = (SAPbouiCOM.EditText)Matrix01.Columns.Item("V_3").Cells.Item(Matrix01.RowCount).Specific;
                EditText.Value = "";
            }
            catch
            { }

        }
        #endregion

        #region Create
        public static SAPbouiCOM.DataTable GetSapDataTable(string sql)
        {
            try
            {
                SAPbouiCOM.DataTable oDataTable = null;
                SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("0", 0);
                try
                {
                    oDataTable = oForm.DataSources.DataTables.Item("TableQuery");
                }
                catch
                {
                    oDataTable = oForm.DataSources.DataTables.Add("TableQuery");
                }
                oDataTable.ExecuteQuery(sql);
                return oDataTable;
            }
            catch { }
            return new SAPbouiCOM.DataTable();
        }
        public static bool ExcuteQuery(string sql)
        {
            try
            {
                SAPbouiCOM.DataTable oDataTable = null;
                SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("0", 0);
                try
                {
                    oDataTable = oForm.DataSources.DataTables.Add("TableQuery");
                }
                catch
                {
                    oDataTable = oForm.DataSources.DataTables.Item("TableQuery");
                }
                oDataTable.ExecuteQuery(sql);
                return true;
            }
            catch { }
            return false;
        }
        /// <summary>
        /// Method write string to file 
        /// </summary>
        /// <param name="content">String content</param>
        /// <param name="path">Path save file</param>
        public static void WriteFile(String content, String path)
        {
            //System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + @"\tmp.txt"
            FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine(content);
            sw.Flush();
            sw.Close();
        }
        /// <summary>
        /// Save DataTable to file xml
        /// </summary>
        /// <param name="filePath">Full file path</param>
        /// <param name="dt">DataTable contents</param>
        public static void SaveDataTableToXML(string filePath, System.Data.DataTable dt)
        {
            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
                System.IO.StringWriter sw = new System.IO.StringWriter();
                dt.WriteXml(sw, System.Data.XmlWriteMode.IgnoreSchema, true);
                WriteFile(sw.ToString(), filePath);
            }
            catch { }

        }

        /// <summary>
        /// Method convert from file xml to datatable
        /// </summary>
        /// <param name="path">Path file xml</param>
        /// <returns>DataTable</returns>
        public static System.Data.DataTable ConvertXmlToDataTable(string path)
        {
            try
            {
                if (!File.Exists(path)) return null;
                System.Data.DataSet ds = new System.Data.DataSet();
                ds.ReadXml(path);
                return ds.Tables[0];
            }
            catch { }
            return new System.Data.DataTable();
        }

        public static void ColumnValidValuesAdd(SAPbouiCOM.Form form, SAPbouiCOM.Column column, string sql)
        {
            SAPbouiCOM.DataTable oDataTable = null;
            try
            {
                oDataTable = form.DataSources.DataTables.Add("TableQuery");
            }
            catch
            {
                oDataTable = form.DataSources.DataTables.Item("TableQuery");
            }
            oDataTable.ExecuteQuery(sql);
            if (!oDataTable.IsEmpty)
                for (int i = 0; i <= oDataTable.Rows.Count - 1; i++)
                    column.ValidValues.Add(oDataTable.GetValue(0, i).ToString(), oDataTable.GetValue(1, i).ToString());
        }

        /// <summary>
        /// Get value in to column map data in matrix
        /// Create by khanhvv
        /// </summary>
        /// <param name="columnVal">Column ID on matrix exist value</param>
        /// <param name="columnMap">Column ID on matrix map data</param>
        /// <param name="fieldVal">Field name exist data on table</param>
        /// <param name="fieldMap">Field name map data on table</param>
        /// <param name="oDBDataSource">Object data source exist table</param>
        /// <param name="oMatrix">Matrix map data</param>
        /// <param name="SBO_Application">SAPbouiCOM Application</param>
        /// <remarks></remarks>
        public static void LoadColumnMapData(string columnVal, string columnMap, string fieldVal, string fieldMap, SAPbouiCOM.DBDataSource oDBDataSource, ref SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.Application SBO_Application)
        {
            int i = 0;
            SAPbouiCOM.EditText oEdit = default(SAPbouiCOM.EditText);
            SAPbouiCOM.Conditions oCondts = default(SAPbouiCOM.Conditions);
            try
            {
                for (i = 1; i <= oMatrix.RowCount; i++)
                {
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(columnVal).Cells.Item(i).Specific;
                    oCondts = (SAPbouiCOM.Conditions)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    Get_Conditional(SBO_Application, ref oCondts, fieldVal, oEdit.Value.Trim(), SAPbouiCOM.BoConditionOperation.co_EQUAL, "", "");
                    oDBDataSource.Query(oCondts);
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(columnMap).Cells.Item(i).Specific;
                    if (oDBDataSource.Size > 0)
                        oEdit.Value = oDBDataSource.GetValue(fieldMap, 0).ToString();
                    else
                        oEdit.Value = "";
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void ChooFormlistEventOnMatrix_New(SAPbouiCOM.Application Aplication, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.DBDataSource oDBSource, string matrixUID, string aliasName, string aliasSource = "", string aliasSMap = "", string columnMap = "", string aliasCMap = "")
        {
            SAPbouiCOM.ChooseFromListEvent oCFLEvento = default(SAPbouiCOM.ChooseFromListEvent);
            SAPbouiCOM.DataTable oDataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            SAPbouiCOM.EditText oEdit = default(SAPbouiCOM.EditText);
            oDataTable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly);
            try
            {
                oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                oDataTable = oCFLEvento.SelectedObjects;
                if ((oDataTable != null))
                {
                    oForm = Aplication.Forms.Item(pVal.FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                    oMatrix.FlushToDataSource();
                    try
                    {
                        oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, oDataTable.GetValue(aliasName, 0).ToString().Trim());
                    }
                    catch
                    {
                    }
                    if (!string.IsNullOrEmpty(aliasSource) & !string.IsNullOrEmpty(aliasSMap))
                    {
                        oMatrix.FlushToDataSource();
                        oDBSource.SetValue(aliasSource, pVal.Row - 1, oDataTable.GetValue(aliasSMap, 0).ToString().Trim());
                        oMatrix.LoadFromDataSource();
                    }
                    if (!string.IsNullOrEmpty(columnMap) & !string.IsNullOrEmpty(aliasCMap))
                    {
                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(columnMap).Cells.Item(pVal.Row).Specific;
                        oEdit.Value = oDataTable.GetValue(aliasCMap, 0).ToString().Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                Aplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void ChooFormlistEventOnMatrixRefTable(SAPbouiCOM.Application Aplication, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.DBDataSource oDBSource, string matrixUID, string aliasName, string aliasSource = "", string aliasSMap = "", string aliasSource2 = "", string aliasSMap2 = "", string columnMapNS = "",
        string aliasMapNS = "", string columnMapNS2 = "", string aliasMapNS2 = "", string tableMap = "", string aliasWhere = "", string aliasCondition = "", string columnMap = "", string aliasCMap = "", string columnMap2 = "", string aliasCMap2 = "")
        {
            SAPbouiCOM.ChooseFromListEvent oCFLEvento = default(SAPbouiCOM.ChooseFromListEvent);
            SAPbouiCOM.DataTable oDataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            SAPbouiCOM.EditText oEdit = default(SAPbouiCOM.EditText);
            try
            {
                oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                oDataTable = oCFLEvento.SelectedObjects;
                if ((oDataTable != null))
                {
                    oForm = Aplication.Forms.Item(pVal.FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                    try
                    {
                        oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, oDataTable.GetValue(aliasName, 0).ToString().Trim());
                    }
                    catch
                    {
                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific;
                        oEdit.Value = oDataTable.GetValue(aliasName, 0).ToString().Trim();
                    }
                    if (!string.IsNullOrEmpty(aliasSource) & !string.IsNullOrEmpty(aliasSMap))
                    {
                        try
                        {
                            oMatrix.FlushToDataSource();
                            oDBSource.SetValue(aliasSource, pVal.Row - 1, oDataTable.GetValue(aliasSMap, 0).ToString().Trim());
                            oMatrix.LoadFromDataSource();
                        }
                        catch
                        {
                        }
                    }
                    if (!string.IsNullOrEmpty(aliasSource2) & !string.IsNullOrEmpty(aliasSMap2))
                    {
                        try
                        {
                            oMatrix.FlushToDataSource();
                            oDBSource.SetValue(aliasSource2, pVal.Row - 1, oDataTable.GetValue(aliasSMap2, 0).ToString().Trim());
                            oMatrix.LoadFromDataSource();
                        }
                        catch
                        {
                        }
                    }
                    if (!string.IsNullOrEmpty(columnMapNS) & !string.IsNullOrEmpty(aliasMapNS))
                    {
                        try
                        {
                            oMatrix.SetCellWithoutValidation(pVal.Row, columnMapNS, oDataTable.GetValue(aliasMapNS, 0).ToString().Trim());
                        }
                        catch
                        {
                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(columnMapNS).Cells.Item(pVal.Row).Specific;
                            oEdit.Value = oDataTable.GetValue(aliasMapNS, 0).ToString().Trim();
                        }
                    }
                    if (!string.IsNullOrEmpty(columnMapNS2) & !string.IsNullOrEmpty(aliasMapNS2))
                    {
                        try
                        {
                            oMatrix.SetCellWithoutValidation(pVal.Row, columnMapNS2, oDataTable.GetValue(aliasMapNS2, 0).ToString().Trim());
                        }
                        catch
                        {
                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(columnMapNS2).Cells.Item(pVal.Row).Specific;
                            oEdit.Value = oDataTable.GetValue(aliasMapNS2, 0).ToString().Trim();
                        }
                    }
                    if (!string.IsNullOrEmpty(columnMap) & !string.IsNullOrEmpty(aliasCMap) & !string.IsNullOrEmpty(tableMap) & !string.IsNullOrEmpty(aliasWhere) & !string.IsNullOrEmpty(aliasCondition))
                    {
                        SAPbouiCOM.DataTable DataTable = null;
                        string sql = null;
                        sql = "Select * from " + tableMap + " where " + aliasWhere + "='" + oDataTable.GetValue(aliasCondition, 0).ToString().Trim() + "'";
                        try
                        {
                            DataTable = oForm.DataSources.DataTables.Add("TableQuery");
                        }
                        catch
                        {
                            DataTable = oForm.DataSources.DataTables.Item("TableQuery");
                        }
                        DataTable.ExecuteQuery(sql);
                        if (!DataTable.IsEmpty)
                        {
                            try
                            {
                                oMatrix.SetCellWithoutValidation(pVal.Row, columnMap, DataTable.GetValue(aliasCMap, 0).ToString().Trim());
                            }
                            catch
                            {
                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(columnMap).Cells.Item(pVal.Row).Specific;
                                oEdit.Value = DataTable.GetValue(aliasCMap, 0).ToString().Trim();
                            }
                            if (!string.IsNullOrEmpty(columnMap2) & !string.IsNullOrEmpty(aliasCMap2))
                            {
                                try
                                {
                                    oMatrix.SetCellWithoutValidation(pVal.Row, columnMap2, DataTable.GetValue(aliasCMap2, 0).ToString().Trim());
                                }
                                catch
                                {
                                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(columnMap2).Cells.Item(pVal.Row).Specific;
                                    oEdit.Value = DataTable.GetValue(aliasCMap2, 0).ToString().Trim();
                                }

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void ValidateValueBarCodeImeiItemCode(SAPbouiCOM.Application Aplication, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, string oDBSourceMatrixName, SAPbouiCOM.ItemEvent pVal, string colBarCode, string aliasBarCode, string colItemCode, string aliasItemCode, string colImei,
        string colItemName)
        {
            SAPbouiCOM.EditText oEdit = default(SAPbouiCOM.EditText);
            SAPbouiCOM.DataTable DataTable = null;
            SAPbouiCOM.DBDataSource oDBSourceMatrix = default(SAPbouiCOM.DBDataSource);
            string sql = null;
            try
            {
                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific;
                try
                {
                    oDBSourceMatrix = oForm.DataSources.DBDataSources.Item(oDBSourceMatrixName);
                }
                catch
                {
                    oDBSourceMatrix = oForm.DataSources.DBDataSources.Add(oDBSourceMatrixName);
                }
                if (pVal.ColUID.Equals(colBarCode))
                {
                    if (!string.IsNullOrEmpty(oEdit.Value.Trim()))
                    {
                        sql = "Select * from OITM where CodeBars ='" + oEdit.Value.Trim() + "'";
                        try
                        {
                            DataTable = oForm.DataSources.DataTables.Add("TableQuery");
                        }
                        catch
                        {
                            DataTable = oForm.DataSources.DataTables.Item("TableQuery");
                        }
                        DataTable.Clear();
                        DataTable.ExecuteQuery(sql);
                        if (!DataTable.IsEmpty)
                        {
                            try
                            {
                                oMatrix.FlushToDataSource();
                                //Change 07/08/2014
                                oDBSourceMatrix.SetValue(aliasItemCode, pVal.Row - 1, DataTable.Columns.Item("ItemCode").Cells.Item(0).Value.ToString().Trim());
                                //oDBSourceMatrix.SetValue(aliasItemCode, pVal.Row - 1, DataTable.Rows.Item(0).Item("ItemCode").ToString().Trim());
                                oMatrix.LoadFromDataSource();
                            }
                            catch
                            {
                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colItemCode).Cells.Item(pVal.Row).Specific;

                                //Change 07/08/2014
                                oEdit.Value = DataTable.Columns.Item("ItemCode").Cells.Item(0).ToString().Trim();
                                //oEdit.Value = DataTable.Rows.Item(0).Item("ItemCode").ToString().Trim();
                            }
                            try
                            {
                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colItemName).Cells.Item(pVal.Row).Specific;

                                //Changed 07/08/2014
                                oEdit.Value = DataTable.Columns.Item("ItemName").Cells.Item(0).ToString().Trim();
                                //oEdit.Value = DataTable.Rows.Item(0).Item("ItemName").ToString().Trim();
                            }
                            catch
                            {
                                //Changed 07/08/2014
                                oMatrix.SetCellWithoutValidation(pVal.Row, colItemName, DataTable.Columns.Item("ItemName").Cells.Item(0).ToString().Trim());
                                //oMatrix.SetCellWithoutValidation(pVal.Row, colItemName, DataTable.Rows.Item(0).Item("ItemName").ToString().Trim());
                            }
                        }

                        DataTable = null;
                    }

                }
                else if (pVal.ColUID.Equals(colImei))
                {
                    if (!string.IsNullOrEmpty(oEdit.Value.Trim()))
                    {
                        sql = "Select i.* from OITM as i,OSRI as s where i.ItemCode = s.ItemCode and s.IntrSerial ='" + oEdit.Value.Trim() + "'";
                        DataTable.Clear();
                        DataTable.ExecuteQuery(sql);
                        if (!DataTable.IsEmpty)
                        {
                            try
                            {
                                oMatrix.FlushToDataSource();
                                oDBSourceMatrix.SetValue(aliasItemCode, pVal.Row - 1, DataTable.Columns.Item("ItemCode").Cells.Item(0).ToString().Trim());
                                oDBSourceMatrix.SetValue(aliasItemCode, pVal.Row - 1, DataTable.Columns.Item("CodeBars").Cells.Item(0).ToString().Trim());
                                oMatrix.LoadFromDataSource();
                            }
                            catch
                            {
                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colItemCode).Cells.Item(pVal.Row).Specific;
                                oEdit.Value = DataTable.Columns.Item("ItemCode").Cells.Item(0).ToString().Trim();
                            }
                            try
                            {
                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colItemName).Cells.Item(pVal.Row).Specific;
                                oEdit.Value = DataTable.Columns.Item("ItemName").Cells.Item(0).ToString().Trim();
                            }
                            catch
                            {
                                oMatrix.SetCellWithoutValidation(pVal.Row, colItemName, DataTable.Columns.Item("ItemName").Cells.Item(0).ToString().Trim());
                            }

                        }
                        DataTable = null;
                    }
                }
                else if (pVal.ColUID.Equals(colItemCode))
                {
                    if (!string.IsNullOrEmpty(oEdit.Value.Trim()))
                    {
                        sql = "Select * from OITM where ItemCode ='" + oEdit.Value.Trim() + "'";
                        DataTable.Clear();
                        DataTable.ExecuteQuery(sql);
                        if (!DataTable.IsEmpty)
                        {
                            try
                            {
                                oMatrix.FlushToDataSource();
                                oDBSourceMatrix.SetValue(aliasBarCode, pVal.Row - 1, DataTable.Columns.Item("CodeBars").Cells.Item(0).ToString().Trim());
                                oMatrix.LoadFromDataSource();
                            }
                            catch
                            {
                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colBarCode).Cells.Item(pVal.Row).Specific;
                                oEdit.Value = DataTable.Columns.Item("CodeBars").Cells.Item(0).ToString().Trim();
                            }
                            try
                            {
                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colItemName).Cells.Item(pVal.Row).Specific;
                                oEdit.Value = DataTable.Columns.Item("ItemName").Cells.Item(0).ToString().Trim();
                            }
                            catch
                            {
                                oMatrix.SetCellWithoutValidation(pVal.Row, colItemName, DataTable.Columns.Item("ItemName").Cells.Item(0).ToString().Trim());
                            }

                        }

                        DataTable = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Aplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private static OleDbConnection GetConnection(string DBServer, string DBName, string DBUser, string DBPass, string DBPort)
        {
            string connectstr = null;
            OleDbConnection connection = null;
            try
            {
                if (string.IsNullOrEmpty(DBPort.Trim()))
                {
                    connectstr = "Provider=SQLOLEDB;Data Source=" + DBServer + ";Persist Security Info=True;User ID=" + DBUser + ";Password=" + DBPass + ";Initial Catalog=" + DBName;
                }
                else
                {
                    connectstr = "Provider=SQLOLEDB;Server=" + DBServer + "," + DBPort + ";" + "Database=" + DBName + ";User ID=" + DBUser + ";Password=" + DBPass + ";" + "Trusted_Connection=False;";
                }
                try
                {
                    connection = new OleDbConnection(connectstr);
                    connection.Open();
                    return connection;
                }
                catch
                {
                }
            }
            catch
            {
                return null;
            }
            return null;
        }

        private static System.Data.DataTable GetData(string sql, SAPbobsCOM.Company oCompany)
        {
            SAPbouiCOM.DataTable oRs = null;
            OleDbConnection connection = null;
            DataTable dt = new DataTable();
            try
            {
                oRs = Globals.GetSapDataTable("Select * from [@FPTSYS] where U_SYSCODE='COMHO'");
                if (!oRs.IsEmpty)
                {
                    connection = GetConnection(oRs.GetValue("U_ServIP", 0).ToString().Trim(), oRs.GetValue("U_DBLink", 0).ToString().Trim(),
                        oRs.GetValue("U_SYSNAME", 0).ToString().Trim(), oRs.GetValue("U_VALUES", 0).ToString().Trim(),
                        oRs.GetValue("U_Port", 0).ToString().Trim());
                    if (connection.State == ConnectionState.Open)
                    {
                        dt = GetDataTable(connection, sql);
                        connection.Close();
                        connection = null;
                        GC.Collect();
                        return dt;
                    }
                }
            }
            catch { }
            return dt;
        }

        private static System.Data.DataTable GetDataTable(OleDbConnection connection, string sql)
        {
            OleDbCommand cm = null;
            OleDbDataAdapter da = null;
            DataSet ds = new DataSet();
            DataTable dt = null;

            try
            {
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }
                cm = new OleDbCommand(sql, connection);
                da = new OleDbDataAdapter(cm);
                try
                {
                    ds.Tables.Remove("Store");
                }
                catch
                {
                }
                da.Fill(ds, "Store");
                dt = ds.Tables["Store"];
                cm.Dispose();
                da.Dispose();
                return dt;
            }
            catch
            {
                cm.Dispose();
                da.Dispose();
                return null;
            }

        }
        //public Form LoadForm(string formPath) 
        //{ 
        //    var oXmlDoc = new XmlDocument(); 
        //    var oCreationPackage = (FormCreationParams)oApp.CreateObject(BoCreatableObjectType.cot_FormCreationParams); 
        //    oCreationPackage.UniqueID = string.Format("{0}{1}", oCreationPackage.UniqueID, Guid.NewGuid().ToString().Substring(2, 10)).Replace("-", string.Empty); 
        //    oXmlDoc.Load(formPath); oCreationPackage.XmlData = oXmlDoc.InnerXml; 
        //    return oApp.Forms.AddEx(oCreationPackage); 
        //}
        public static SAPbouiCOM.Form GetFormUDO(SAPbouiCOM.Application SBO_Application, string UniqueID, string FileSrf, string ObjectType = "", bool checkHasForm = false)
        {
            SAPbouiCOM.Form functionReturnValue = default(SAPbouiCOM.Form);
            SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);
            try
            {
                fcp = (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.FormType = UniqueID;
                //fcp.UniqueID = UniqueID;
                fcp.UniqueID = string.Format("{0}{1}", fcp.UniqueID, Guid.NewGuid().ToString().Substring(2, 10)).Replace("-", string.Empty);
                if (!string.IsNullOrEmpty(ObjectType))
                    fcp.ObjectType = ObjectType;
                fcp.XmlData = Globals.LoadFromXML(FileSrf);
                try
                {
                    functionReturnValue = SBO_Application.Forms.AddEx(fcp);
                    Globals.SetLangueForItem(functionReturnValue, functionReturnValue.UniqueID, Globals.GetCaptionItem());
                    checkHasForm = false;
                    return functionReturnValue;
                }
                catch (Exception ex)
                {
                    functionReturnValue = SBO_Application.Forms.Item(UniqueID);
                    functionReturnValue.Select();
                    checkHasForm = true;
                    return functionReturnValue;
                }
            }
            catch (Exception ex)
            {
                functionReturnValue = null;
                return functionReturnValue;
            }
        }

        public static void CopyDataFromSOToMatrixReturn(SAPbouiCOM.Application Aplication, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, string dBMatrixName, SAPbouiCOM.DBDataSource oDBSourceLineSO)
        {
            int i = 0;
            int j = 0;
            SAPbouiCOM.DBDataSource oDBMatrix = default(SAPbouiCOM.DBDataSource);
            try
            {
                try
                {
                    oDBMatrix = oForm.DataSources.DBDataSources.Add(dBMatrixName);
                }
                catch
                {
                    oDBMatrix = oForm.DataSources.DBDataSources.Item(dBMatrixName);
                }
                oDBMatrix.Clear();
                for (i = 0; i <= oDBSourceLineSO.Size - 1; i++)
                {
                    if (!string.IsNullOrEmpty(oDBSourceLineSO.GetValue("U_ItmCod", i).Trim()))
                    {
                        oDBMatrix.InsertRecord(i);
                        for (j = 0; j <= oDBMatrix.Fields.Count - 1; j++)
                        {
                            if (oDBMatrix.Fields.Item(j).Name.StartsWith("U_"))
                            {
                                if (oDBMatrix.Fields.Item(j).Name.Equals("U_PriceSO"))
                                {
                                    oDBMatrix.SetValue(oDBMatrix.Fields.Item(j).Name, i, oDBSourceLineSO.GetValue("U_Price", i).Trim());
                                }
                                else if (oDBMatrix.Fields.Item(j).Name.Equals("U_SO_Entry"))
                                {
                                    oDBMatrix.SetValue(oDBMatrix.Fields.Item(j).Name, i, oDBSourceLineSO.GetValue("DocEntry", i).Trim());
                                }
                                else if (oDBMatrix.Fields.Item(j).Name.Equals("U_SO_Line"))
                                {
                                    oDBMatrix.SetValue(oDBMatrix.Fields.Item(j).Name, i, oDBSourceLineSO.GetValue("LineId", i).Trim());
                                }
                                else
                                {
                                    try
                                    {
                                        oDBMatrix.SetValue(oDBMatrix.Fields.Item(j).Name, i, oDBSourceLineSO.GetValue(oDBMatrix.Fields.Item(j).Name, i).Trim());
                                    }
                                    catch
                                    {
                                    }
                                }

                            }
                        }
                    }
                }
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Aplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void CopyDataFromSOToHeaderKhanhVV(SAPbouiCOM.Application Aplication, SAPbouiCOM.Form oForm, SAPbouiCOM.DBDataSource oDBDataSource)
        {
            //SAPbouiCOM.Item oItem = default(SAPbouiCOM.Item);
            SAPbouiCOM.EditText oEdit = default(SAPbouiCOM.EditText);
            SAPbouiCOM.ComboBox oComboBox = default(SAPbouiCOM.ComboBox);
            SAPbouiCOM.CheckBox oCheckBox = default(SAPbouiCOM.CheckBox);
            SAPbouiCOM.Item oItem = null;
            int i;
            try
            {

                for (i = 0; i < oForm.Items.Count; i++)
                {
                    oItem = oForm.Items.Item(i);
                    switch (oItem.Type)
                    {
                        case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                            if (!oItem.UniqueID.ToUpper().Equals("DOCENTRY") & !oItem.UniqueID.ToUpper().Equals("DOCNUM"))
                            {
                                oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                                try
                                {
                                    if (oItem.UniqueID.ToUpper().Equals("U_SRBILL") & oDBDataSource.TableName.ToString() == "@FPTORDR")
                                    {
                                        oEdit.Value = oDBDataSource.GetValue("U_SyBill", 0).ToString().Trim();
                                    }
                                    else
                                    {
                                        oEdit.Value = oDBDataSource.GetValue(oItem.UniqueID, 0).ToString().Trim();
                                    }
                                }
                                catch
                                {

                                }
                            }
                            break;
                        case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                            oCheckBox = (SAPbouiCOM.CheckBox)oItem.Specific;
                            try
                            {
                                if (oDBDataSource.GetValue(oItem.UniqueID, 0).ToString().Trim().Equals("Y"))
                                {
                                    oCheckBox.Checked = true;
                                }
                                else
                                {
                                    oCheckBox.Checked = false;
                                }
                            }
                            catch
                            {
                            }
                            break;
                        case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                            oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
                            try
                            {
                                oComboBox.Select(oDBDataSource.GetValue(oItem.UniqueID, 0).ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }
                            catch
                            {
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Aplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static void CopyDataFromSOToMatrixKhanhVV(SAPbouiCOM.Application Aplication, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, string dBMatrixName, SAPbouiCOM.DBDataSource oDBSourceLineSO)
        {
            int i = 0;
            int j = 0;
            SAPbouiCOM.DBDataSource oDBMatrix = default(SAPbouiCOM.DBDataSource);
            try
            {
                try
                {
                    oDBMatrix = oForm.DataSources.DBDataSources.Add(dBMatrixName);
                }
                catch
                {
                    oDBMatrix = oForm.DataSources.DBDataSources.Item(dBMatrixName);
                }
                oDBMatrix.Clear();
                for (i = 0; i <= oDBSourceLineSO.Size - 1; i++)
                {
                    oDBMatrix.InsertRecord(i);
                    for (j = 0; j <= oDBMatrix.Fields.Count - 1; j++)
                    {
                        if (oDBMatrix.Fields.Item(j).Name.StartsWith("U_"))
                        {
                            if (oDBMatrix.Fields.Item(j).Name.Equals("U_DocSO"))
                            {
                                oDBMatrix.SetValue(oDBMatrix.Fields.Item(j).Name, i, oDBSourceLineSO.GetValue("DocEnTry", i).Trim());
                            }
                            else if (oDBMatrix.Fields.Item(j).Name.Equals("U_LineSO"))
                            {
                                oDBMatrix.SetValue(oDBMatrix.Fields.Item(j).Name, i, oDBSourceLineSO.GetValue("LineId", i).Trim());
                            }
                            else
                            {
                                try
                                {
                                    oDBMatrix.SetValue(oDBMatrix.Fields.Item(j).Name, i, oDBSourceLineSO.GetValue(oDBMatrix.Fields.Item(j).Name, i).Trim());
                                }
                                catch
                                {
                                }
                            }

                        }

                    }
                }
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Aplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static bool GetCompanyConnect(ref SAPbobsCOM.Company Company, string LicenseServer, string ComIP, string ComUS, string ComPass, SAPbobsCOM.BoSuppLangs ComLanguage, SAPbobsCOM.BoDataServerTypes ComServerType, string DBName, string DBUser, string DBPass)
        {
            Company = new SAPbobsCOM.Company();
            Company.Server = ComIP;
            //"10.15.106.105" '' change to your company server
            Company.LicenseServer = LicenseServer;
            //
            Company.language = ComLanguage;
            //SAPbobsCOM.BoSuppLangs.ln_English ' change to your language
            Company.DbServerType = ComServerType;
            Company.UseTrusted = false;
            Company.DbUserName = DBUser;
            // DbUserName 'APIUser '' "sa"
            Company.DbPassword = DBPass;
            // DbPassword ' "123456"

            Company.CompanyDB = DBName;
            // "ASC_TEST"
            Company.UserName = ComUS;
            // APIUser
            Company.Password = ComPass;
            // APIPass
            if (Company.Connect() != 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Method add new a chooseformlist
        /// Create by khanhvv
        /// </summary>
        /// <param name="CFLsUID">chooseformlist UID</param>
        /// <param name="ObjectName">Object reference to table in database</param>
        /// <param name="oCFLs">Object chooseformlist collection store chooseformlist</param>
        /// <param name="oForm">Object SAPbouiCOM Form</param>
        /// <param name="SBO_Application">Object SAPbouiCOM Application</param>
        /// <remarks></remarks>
        public static void AddChooseFromList(string CFLsUID, string ObjectName, ref SAPbouiCOM.ChooseFromListCollection oCFLs, SAPbouiCOM.Form oForm, SAPbouiCOM.Application SBO_Application)
        {
            try
            {
                oCFLs = oForm.ChooseFromLists;
                SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = default(SAPbouiCOM.ChooseFromListCreationParams);
                oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = ObjectName;
                oCFLCreationParams.UniqueID = CFLsUID;
                oCFL = oCFLs.Add(oCFLCreationParams);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Method add new a chooseformlist with conditions
        /// Create by khanhvv
        /// </summary>
        /// <param name="CFLsUID">chooseformlist UID</param>
        /// <param name="ObjectName">Object reference to table in database</param>
        /// <param name="oCFLs">Object chooseformlist collection store chooseformlist</param>
        /// <param name="oForm">Object SAPbouiCOM Form</param>
        /// <param name="SBO_Application">Object SAPbouiCOM Application</param>
        /// <param name="oConditions">SAPbouiCOM Conditions store condition</param>
        /// <remarks></remarks>
        public static void AddChooseFromList(string CFLsUID, string ObjectName, ref SAPbouiCOM.ChooseFromListCollection oCFLs, SAPbouiCOM.Form oForm, SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Conditions oConditions)
        {
            try
            {
                oCFLs = oForm.ChooseFromLists;
                SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = default(SAPbouiCOM.ChooseFromListCreationParams);
                oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = ObjectName;
                oCFLCreationParams.UniqueID = CFLsUID;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oCFL.SetConditions(oConditions);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void GetConditions(ref SAPbouiCOM.Conditions oConditions, SAPbouiCOM.Condition oCondition, string field, string val, SAPbouiCOM.BoConditionOperation operation, SAPbouiCOM.BoConditionRelationship relationship)
        {
            oCondition = oConditions.Add();
            //oCondition.BracketOpenNum = backetNumber
            oCondition.Alias = field;
            oCondition.Operation = operation;
            oCondition.CondVal = val;
            //oCondition.BracketCloseNum = backetNumber
            oCondition.Relationship = relationship;
        }
        public static void GetConditions(ref SAPbouiCOM.Conditions oConditions, string field, string val, SAPbouiCOM.BoConditionOperation operation, SAPbouiCOM.BoConditionRelationship relationship)
        {
            SAPbouiCOM.Condition oCondition = default(SAPbouiCOM.Condition);
            oCondition = oConditions.Add();
            //oCondition.BracketOpenNum = backetNumber
            oCondition.Alias = field;
            oCondition.Operation = operation;
            oCondition.CondVal = val;
            //oCondition.BracketCloseNum = backetNumber
            oCondition.Relationship = relationship;
        }
        /// <summary>
        /// Method create menu user
        /// </summary>
        /// <param name="menuId">Identity menu</param>
        /// <param name="menuName">Menu name (display)</param>
        /// <param name="menuType">Menu type</param>
        /// <param name="parentMenuId">Identity menu parent</param>
        /// <param name="SBO_Application">SAPbouiCOM Application</param>
        /// <remarks></remarks>
        public static void CreateMenu(string menuId, string menuName, SAPbouiCOM.BoMenuType menuType, string parentMenuId,
            SAPbouiCOM.Application SBO_Application)
        {
            if (SBO_Application.Menus.Exists(menuId)) return;
            try
            {
                SAPbouiCOM.MenuCreationParams oMenuCreationParams = default(SAPbouiCOM.MenuCreationParams);
                SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);
                SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);
                oMenuCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuCreationParams.Type = menuType;
                oMenuCreationParams.UniqueID = menuId;
                oMenuCreationParams.String = menuName;
                oMenuCreationParams.Enabled = true;
                if (SBO_Application.Menus.Exists(parentMenuId))
                {
                    oMenuItem = SBO_Application.Menus.Item(parentMenuId);
                    oMenus = oMenuItem.SubMenus;
                    oMenuCreationParams.Position = oMenus.Count + 1;
                    oMenus.AddEx(oMenuCreationParams);
                }
                oMenuCreationParams = null;
                oMenuItem = null;
                oMenus = null;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        #endregion

        #region Method sap b1
        public static void CFLFillDataToItemNew(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal,
            ref SAPbouiCOM.EditText EditText, string Col_Alias, ref SAPbouiCOM.EditText EditText1, string Col_Alias1 = "")
        {
            SAPbouiCOM.DataTable DataTable = default(SAPbouiCOM.DataTable);
            string pStr = null;
            string Item_Name = null;
            SAPbouiCOM.IChooseFromListEvent MakerCFLEvnt = default(SAPbouiCOM.IChooseFromListEvent);
            try
            {
                MakerCFLEvnt = (SAPbouiCOM.ChooseFromListEvent)pVal;
                if (MakerCFLEvnt.SelectedObjects == null)
                    return;
                Item_Name = pVal.ItemUID;
                DataTable = MakerCFLEvnt.SelectedObjects;
                if ((DataTable != null))
                {
                    if ((EditText1 != null) & !string.IsNullOrEmpty(Col_Alias1))
                    {
                        try
                        {
                            pStr = DataTable.GetValue(Col_Alias1, 0).ToString();
                            if (EditText1.Value.ToString().Trim() != pStr.ToString().Trim())
                            {
                                EditText1.Value = pStr;
                            }

                        }
                        catch { }
                    }
                    try
                    {
                        pStr = DataTable.GetValue(Col_Alias, 0).ToString();
                        if (EditText.Value.ToString().Trim() != pStr.ToString().Trim())
                        {
                            EditText.Value = pStr;
                        }

                    }
                    catch { }
                    if (Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                }
                DataTable = null;
                Form = null;
            }
            catch
            {
                //SBO_Application.StatusBar.SetText("Loi SBO_Application_AppEvent ! Loi : " & ex.ToString, _
                //            SAPbouiCOM.BoMessageTime.bmt_Short, _
                //            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            }
        }

        public static System.Data.DataTable GetReport()
        {
            System.Data.DataTable g_ReportParameter = new System.Data.DataTable();
            System.Data.DataRow row = null;

            string sql = "";
            try
            {
                if (!File.Exists(path + @"\g_ReportParameter.xml"))
                {
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    sql = "SELECT a.Code,[U_Order],[U_ParCode],[U_ItmCode],[U_ItmDes] ,[U_ItmType],[U_CFL],[U_CFLAlias],[U_SQL],[U_ReQuired] ";
                    sql += " FROM [@FPTRPT1] A, [@FPTORPT] b where A.code=b.code and b.U_Status='Y' and a.[U_Status]='Y'  Order by a.Code, a.U_Order";
                    SAPbouiCOM.DataTable oDataTable = GetSapDataTable(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_ReportParameter.Columns.Add("U_Order", Type.GetType("System.Double"));
                        g_ReportParameter.Columns.Add("Code", Type.GetType("System.String"));
                        g_ReportParameter.Columns.Add("U_ParCode", Type.GetType("System.String"));
                        g_ReportParameter.Columns.Add("U_ItmCode", Type.GetType("System.String"));
                        g_ReportParameter.Columns.Add("U_ItmDes", Type.GetType("System.String"));
                        g_ReportParameter.Columns.Add("U_ItmType", Type.GetType("System.String"));
                        g_ReportParameter.Columns.Add("U_CFL", Type.GetType("System.String"));
                        g_ReportParameter.Columns.Add("U_CFLAlias", Type.GetType("System.String"));
                        g_ReportParameter.Columns.Add("U_SQL", Type.GetType("System.String"));
                        g_ReportParameter.Columns.Add("U_ReQuired", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_ReportParameter.NewRow();
                            row["Code"] = oDataTable.GetValue("Code", i).ToString().Trim();
                            row["U_Order"] = oDataTable.GetValue("U_Order", i).ToString().Trim();
                            row["U_ParCode"] = oDataTable.GetValue("U_ParCode", i).ToString().Trim();
                            row["U_ItmCode"] = oDataTable.GetValue("U_ItmCode", i).ToString().Trim();
                            row["U_ItmDes"] = oDataTable.GetValue("U_ItmDes", i).ToString().Trim();
                            row["U_ItmType"] = oDataTable.GetValue("U_ItmType", i).ToString().Trim();
                            row["U_CFL"] = oDataTable.GetValue("U_CFL", i).ToString().Trim();
                            row["U_CFLAlias"] = oDataTable.GetValue("U_CFLAlias", i).ToString().Trim();
                            row["U_SQL"] = oDataTable.GetValue("U_SQL", i).ToString().Trim();
                            row["U_ReQuired"] = oDataTable.GetValue("U_ReQuired", i).ToString().Trim();
                            g_ReportParameter.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_ReportParameter.xml", g_ReportParameter);
                    }
                }
                else
                {
                    g_ReportParameter = ConvertXmlToDataTable(path + @"\g_ReportParameter.xml");
                }
            }
            catch
            { }
            return g_ReportParameter;
        }

        public static System.Data.DataTable[] GetSeries()
        {
            string SQL = null;
            SAPbouiCOM.DataTable Rs = default(SAPbouiCOM.DataTable);
            int Count = 0;
            System.Data.DataRow Row = null;
            System.Data.DataTable g_BP_Custom_Series = null;
            System.Data.DataTable g_BP_Supply_Series = null;
            System.Data.DataTable[] g_Series_Array = new System.Data.DataTable[2];
            if (g_BP_Custom_Series == null)
            {
                g_BP_Custom_Series = new System.Data.DataTable("Table01");
                g_BP_Custom_Series.Columns.Add("Series", Type.GetType("System.Double"));
                g_BP_Custom_Series.Columns.Add("SeriesName", Type.GetType("System.String"));

            }
            else
            {
                if (g_BP_Custom_Series.Rows.Count > 0)
                {
                    g_BP_Custom_Series.Clear();
                }

            }

            if (g_BP_Supply_Series == null)
            {
                g_BP_Supply_Series = new System.Data.DataTable("Table01");
                g_BP_Supply_Series.Columns.Add("Series", Type.GetType("System.Double"));
                g_BP_Supply_Series.Columns.Add("SeriesName", Type.GetType("System.String"));

            }
            else
            {
                if (g_BP_Supply_Series.Rows.Count > 0)
                {
                    g_BP_Supply_Series.Clear();
                }

            }

            try
            {
                SQL = "select Series, SeriesName  from NNM1 where  Locked='N' and ObjectCode='2' and DocSubType='C' and IsManual = 'N' Order by Series";
                Rs = GetSapDataTable(SQL);
                for (Count = 0; Count <= Rs.Rows.Count - 1; Count++)
                {
                    Row = g_BP_Custom_Series.NewRow();
                    Row["Series"] = Rs.Columns.Item("Series").Cells.Item(Count).ToString().Trim();
                    Row["SeriesName"] = Rs.Columns.Item("SeriesName").Cells.Item(Count).ToString().Trim();
                    g_BP_Custom_Series.Rows.Add(Row);
                }

            }
            catch
            {
            }
            Rs = null;


            try
            {
                SQL = "select Series, SeriesName  from NNM1 where  Locked='N' and ObjectCode='2' and DocSubType='S' and IsManual = 'N' Order by Series";

                Rs = GetSapDataTable(SQL);
                for (Count = 0; Count <= Rs.Rows.Count - 1; Count++)
                {
                    Row = g_BP_Supply_Series.NewRow();
                    Row["Series"] = Rs.Columns.Item("Series").Cells.Item(Count).ToString().Trim();
                    Row["SeriesName"] = Rs.Columns.Item("SeriesName").Cells.Item(Count).ToString().Trim();
                    g_BP_Supply_Series.Rows.Add(Row);
                }

            }
            catch
            {
            }
            Rs = null;
            g_Series_Array[0] = g_BP_Custom_Series;
            g_Series_Array[1] = g_BP_Supply_Series;
            return g_Series_Array;

            //Supply
        }

        public static System.Data.DataTable GetBPCustomSeries()
        {
            System.Data.DataTable g_BP_Custom_Series = new System.Data.DataTable();
            System.Data.DataRow row = null;
            SAPbouiCOM.DataTable oDataTable = null;
            SAPbouiCOM.Form oForm = null;
            string sql = "";
            try
            {
                if (!File.Exists(path + @"\g_BP_Custom_Series.xml"))
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    oForm = SBO_Application.Forms.GetForm("0", 0);
                    try
                    {
                        oDataTable = oForm.DataSources.DataTables.Add("TableQuery");
                    }
                    catch
                    {
                        oDataTable = oForm.DataSources.DataTables.Item("TableQuery");
                    }
                    sql = "select Series, SeriesName  from NNM1 where  Locked='N' and ObjectCode='2' and DocSubType='C' and IsManual = 'N' Order by Series";
                    oDataTable.ExecuteQuery(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_BP_Custom_Series.Columns.Add("Series", Type.GetType("System.Double"));
                        g_BP_Custom_Series.Columns.Add("SeriesName", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_BP_Custom_Series.NewRow();
                            row["Series"] = oDataTable.GetValue("Series", i).ToString().Trim();
                            row["SeriesName"] = oDataTable.GetValue("SeriesName", i).ToString().Trim();
                            g_BP_Custom_Series.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_BP_Custom_Series.xml", g_BP_Custom_Series);
                    }
                }
                else
                {
                    g_BP_Custom_Series = ConvertXmlToDataTable(path + @"\g_BP_Custom_Series.xml");
                }
            }
            catch
            { }
            return g_BP_Custom_Series;
        }

        public static System.Data.DataTable GetBPSupplySeries()
        {
            System.Data.DataTable g_BP_Supply_Series = new System.Data.DataTable();
            System.Data.DataRow row = null;
            SAPbouiCOM.DataTable oDataTable = null;
            SAPbouiCOM.Form oForm = null;
            string sql = "";
            try
            {
                if (!File.Exists(path + @"\g_BP_Supply_Series.xml"))
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    oForm = SBO_Application.Forms.GetForm("0", 0);
                    try
                    {
                        oDataTable = oForm.DataSources.DataTables.Add("TableQuery");
                    }
                    catch
                    {
                        oDataTable = oForm.DataSources.DataTables.Item("TableQuery");
                    }
                    sql = "select Series, SeriesName  from NNM1 where  Locked='N' and ObjectCode='2' and DocSubType='S' and IsManual = 'N' Order by Series";
                    oDataTable.ExecuteQuery(sql);
                    if (!oDataTable.IsEmpty)
                    {
                        g_BP_Supply_Series.Columns.Add("Series", Type.GetType("System.Double"));
                        g_BP_Supply_Series.Columns.Add("SeriesName", Type.GetType("System.String"));

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            row = g_BP_Supply_Series.NewRow();
                            row["Series"] = oDataTable.GetValue("Series", i).ToString().Trim();
                            row["SeriesName"] = oDataTable.GetValue("SeriesName", i).ToString().Trim();
                            g_BP_Supply_Series.Rows.Add(row);
                        }
                        SaveDataTableToXML(path + @"\g_BP_Supply_Series.xml", g_BP_Supply_Series);
                    }
                }
                else
                {
                    g_BP_Supply_Series = ConvertXmlToDataTable(path + @"\g_BP_Supply_Series.xml");
                }
            }
            catch
            { }
            return g_BP_Supply_Series;
        }

        public static string GetItemFromSerial(string DistNumber, ref string BarCode, ref string ItemCode, ref string ItemName, ref string WhsCode, bool Check = false, string ShpCode = "")
        {
            string functionReturnValue = null;
            SAPbouiCOM.DataTable Rs = default(SAPbouiCOM.DataTable);
            string SQL = null;
            try
            {
                functionReturnValue = "";
                if (string.IsNullOrEmpty(DistNumber))
                    return functionReturnValue;
                if (Check == true)
                {
                    SQL = "select bbb.DistNumber,bbb.ItemCode,  aaa.ItemName, bbb.CodeBars,aaa.WhsCode  FROM  ";
                    SQL += "(select WhsCode,SysNumber ,ItemCode, (select ItemName from OITM where ItemCode=aa.ItemCode ) as ItemName ";
                    SQL += " from  OSRQ aa where Quantity >0  and WhsCode='" + WhsCode + "' ) aaa, ";
                    SQL += "(select DistNumber, ItemCode, SysNumber, (select CodeBars from OITM where ItemCode=bb.ItemCode ) as CodeBars ";
                    SQL += " from OSRN bb where DistNumber='" + DistNumber + "') bbb ";
                    SQL += "where  aaa.SysNumber=bbb.SysNumber and  aaa.ItemCode=bbb.ItemCode";
                }
                else
                {
                    SQL = "select bbb.DistNumber,bbb.ItemCode,  aaa.ItemName, bbb.CodeBars,aaa.WhsCode  FROM  ";
                    SQL += "(select WhsCode,SysNumber ,ItemCode, (select ItemName from OITM where ItemCode=aa.ItemCode ) as ItemName ";
                    SQL += "  from  OSRQ aa where Quantity >0 ) aaa, (select DistNumber, ItemCode, SysNumber, ";
                    SQL += "(select CodeBars from OITM where ItemCode=bb.ItemCode ) as CodeBars  from OSRN bb where DistNumber='";
                    SQL += DistNumber + "') bbb " + "where  aaa.SysNumber=bbb.SysNumber and  aaa.ItemCode=bbb.ItemCode ";
                    SQL += " and aaa.WhsCode  in (select WhsCode from OWHS where U_Code_SH='" + ShpCode + "') ";
                }

                Rs = GetSapDataTable(SQL);
                if (!Rs.IsEmpty)
                {
                    functionReturnValue = Rs.Columns.Item(0).Cells.Item(0).ToString().Trim();
                    BarCode = Rs.GetValue(3, 0).ToString().Trim();
                    ItemCode = Rs.GetValue(1, 0).ToString().Trim();
                    ItemName = Rs.GetValue(2, 0).ToString().Trim();
                    WhsCode = Rs.GetValue(4, 0).ToString().Trim();
                }
            }
            catch { }
            return functionReturnValue;
        }

        public static string GetTimeServerClient()
        {
            SAPbouiCOM.DataTable Rs = default(SAPbouiCOM.DataTable);
            Rs = GetSapDataTable("select convert(varchar,GETDATE(),121)");
            return (Rs.GetValue(0, 0).ToString().Trim() + "-" + System.DateTime.Now);
        }

        public static void HeaderDataBind(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.DBDataSource Db, string DocNum)
        {
            int Count = 0;
            string TableName = "";
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);
            SAPbouiCOM.ComboBox ComboBox = default(SAPbouiCOM.ComboBox);
            SAPbouiCOM.EditText EditText = default(SAPbouiCOM.EditText);
            SAPbouiCOM.CheckBox CheckBox = default(SAPbouiCOM.CheckBox);

            try
            {
                TableName = Db.TableName;
                for (Count = 0; Count <= Form.Items.Count - 1; Count++)
                {
                    try
                    {
                        //If Count = 59 Then
                        //    MsgBox("sds")
                        //End If
                        Item = Form.Items.Item(Count);

                    }
                    catch
                    {
                    }


                    //If Item.UniqueID = "U_TMonBI1" Then
                    //    MsgBox("dfgdg")
                    //End If
                    switch (Item.Type)
                    {
                        case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                            try
                            {
                                Item.DisplayDesc = true;
                                ComboBox = (SAPbouiCOM.ComboBox)Item.Specific;
                                ComboBox.DataBind.SetBound(true, TableName, Item.UniqueID);

                            }
                            catch
                            {
                            }

                            break;
                        case (SAPbouiCOM.BoFormItemTypes.it_EDIT):
                            try
                            {
                                EditText = (SAPbouiCOM.EditText)Item.Specific;
                                EditText.DataBind.SetBound(true, TableName, Item.UniqueID);

                            }
                            catch
                            {
                            }
                            break;
                        case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                            try
                            {
                                CheckBox = (SAPbouiCOM.CheckBox)Item.Specific;
                                CheckBox.DataBind.SetBound(true, TableName, Item.UniqueID);

                            }
                            catch
                            {
                            }
                            break;
                    }
                }

                Form.DataBrowser.BrowseBy = DocNum;

                Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                Form.PaneLevel = 1;
                //  Db.Query()
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
            }
        }

        public static string LoadFromXML(string FileName)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            //sPath = Application.StartupPath & "\"
            oXmlDoc.Load(FileName);
            //string sXML = oXmlDoc.InnerXml.ToString();
            //SBO_Application.LoadBatchActions(ref sXML);
            return (oXmlDoc.InnerXml);
            //return (sXML);
        }
        public static void SetLangueForItem(SAPbouiCOM.Form Form, string FormUID, System.Data.DataTable g_CaptionItem)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);
            SAPbouiCOM.StaticText Static = default(SAPbouiCOM.StaticText);
            SAPbouiCOM.Folder Folder = default(SAPbouiCOM.Folder);
            SAPbouiCOM.Column Column = default(SAPbouiCOM.Column);
            SAPbouiCOM.Button Button = default(SAPbouiCOM.Button);
            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.CheckBox CheckBox = default(SAPbouiCOM.CheckBox);
            System.Data.DataRow[] RowArray = null;
            int Count = 0;
            if (g_CaptionItem == null)
                return;
            if (g_CaptionItem.Rows.Count <= 0) return;
            RowArray = g_CaptionItem.Select("U_FormUID='" + FormUID.ToUpper() + "'");
            if (RowArray.Length > 0)
            {
                for (Count = 0; Count <= RowArray.Length - 1; Count++)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(RowArray[Count]["U_Item"].ToString().Trim()))
                        {
                            if (!string.IsNullOrEmpty(RowArray[Count]["U_Text"].ToString().Trim()))
                            {
                                Item = Form.Items.Item(RowArray[Count]["U_Item"].ToString().Trim());
                                switch (Item.Type)
                                {
                                    case SAPbouiCOM.BoFormItemTypes.it_BUTTON:
                                        Button = (SAPbouiCOM.Button)Item.Specific;
                                        Button.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                    case SAPbouiCOM.BoFormItemTypes.it_FOLDER:
                                        Folder = (SAPbouiCOM.Folder)Item.Specific;
                                        Folder.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                    case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                                        Matrix = (SAPbouiCOM.Matrix)Item.Specific;
                                        Column = Matrix.Columns.Item(RowArray[Count]["U_Column"].ToString().Trim());
                                        Column.TitleObject.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                                        break;
                                    //Grid = Item.Specific
                                    //Grid.Columns.Item(RowArray(Count).Item("U_Column").ToString.Trim()).TitleObject.Caption = RowArray(Count).Item("U_Text").ToString.Trim
                                    case SAPbouiCOM.BoFormItemTypes.it_STATIC:
                                        Static = (SAPbouiCOM.StaticText)Item.Specific;
                                        Static.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                                        CheckBox = (SAPbouiCOM.CheckBox)Item.Specific;
                                        CheckBox.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                }
                            }
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(RowArray[Count]["U_Text"].ToString().Trim()))
                            {
                                Form.Title = RowArray[Count]["U_Text"].ToString().Trim();
                            }
                        }

                    }
                    catch
                    {
                    }
                }
            }
        }

        public static bool SetLangueForItemGrid(SAPbouiCOM.Form Form, string FormUID)
        {
            bool functionReturnValue = false;
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);
            SAPbouiCOM.Grid Grid = default(SAPbouiCOM.Grid);


            System.Data.DataRow[] RowArray = null;
            int Count = 0;
            functionReturnValue = false;
            System.Data.DataTable dt = Globals.GetCaptionItem();
            if (dt == null) return false;
            if (dt.Rows.Count <= 0) return false;
            RowArray = dt.Select("U_FormUID='" + FormUID.ToUpper() + "'");
            if (RowArray.Length > 0)
            {
                for (Count = 0; Count <= RowArray.Length - 1; Count++)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(RowArray[Count]["U_Item"].ToString().Trim()))
                        {
                            if (!string.IsNullOrEmpty(RowArray[Count]["U_Text"].ToString().Trim()))
                            {
                                Item = Form.Items.Item(RowArray[Count]["U_Item"].ToString().Trim());
                                if (Item.Type == SAPbouiCOM.BoFormItemTypes.it_GRID)
                                {
                                    Grid = (SAPbouiCOM.Grid)Item.Specific;
                                    Grid.Columns.Item(RowArray[Count]["U_Column"].ToString().Trim()).TitleObject.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                    functionReturnValue = true;
                                }
                                //Select Case Item.Type
                                //    Case SAPbouiCOM.BoFormItemTypes.it_BUTTON
                                //        Button = Item.Specific
                                //        Button.Caption = RowArray(Count).Item("U_Text").ToString.Trim
                                //    Case SAPbouiCOM.BoFormItemTypes.it_FOLDER
                                //        Folder = Item.Specific
                                //        Folder.Caption = RowArray(Count).Item("U_Text").ToString.Trim
                                //    Case SAPbouiCOM.BoFormItemTypes.it_MATRIX
                                //        Matrix = Item.Specific
                                //        Column = Matrix.Columns.Item(RowArray(Count).Item("U_Column").ToString.Trim())
                                //        Column.TitleObject.Caption = RowArray(Count).Item("U_Text").ToString.Trim
                                //    Case SAPbouiCOM.BoFormItemTypes.it_GRID
                                //        Grid = Item.Specific
                                //        Grid.Columns.Item(RowArray(Count).Item("U_Column").ToString.Trim()).TitleObject.Caption = RowArray(Count).Item("U_Text").ToString.Trim
                                //    Case SAPbouiCOM.BoFormItemTypes.it_STATIC
                                //        Static = Item.Specific
                                //        Static.Caption = RowArray(Count).Item("U_Text").ToString.Trim
                                //End Select
                            }
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(RowArray[Count]["U_Text"].ToString().Trim()))
                            {
                                Form.Title = RowArray[Count]["U_Text"].ToString().Trim();
                            }
                        }

                    }
                    catch
                    {
                    }
                }
            }
            return functionReturnValue;
        }


        private static string GetChooseFromListsName(SAPbouiCOM.ChooseFromListCollection oCFLs1, string ObjectType)
        {
            string functionReturnValue = null;
            try
            {
                functionReturnValue = oCFLs1.Item(ObjectType).UniqueID;
                functionReturnValue = "CFL" + ObjectType + oCFLs1.Count;
            }
            catch
            {
                functionReturnValue = ObjectType;
            }
            return functionReturnValue;


        }

        private static string Get_ChooseFromLists_Name(SAPbouiCOM.ChooseFromListCollection p_oCFLs1, string p_ObjectType)
        {
            string functionReturnValue = null;
            try
            {
                functionReturnValue = p_oCFLs1.Item(p_ObjectType).UniqueID;
                functionReturnValue = "CFL" + p_ObjectType + p_oCFLs1.Count;
            }
            catch
            {
                functionReturnValue = p_ObjectType;
            }
            return functionReturnValue;
        }

        public static void ChooFormlistEventOnMatrix(SAPbouiCOM.Application Aplication, SAPbouiCOM.ItemEvent pVal, string matrixUID,
            string aliasName, string columnMap = "", string alliasMap = "", string columnMap1 = "", string alliasMap1 = "",
            string columnMap2 = "", string alliasMap2 = "", string columnMap3 = "", string alliasMap3 = "")
        {
            SAPbouiCOM.ChooseFromListEvent oCFLEvento = default(SAPbouiCOM.ChooseFromListEvent);
            SAPbouiCOM.DataTable oDataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            SAPbouiCOM.EditText oEdit = default(SAPbouiCOM.EditText);
            int row = 0;
            try
            {
                oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                oDataTable = oCFLEvento.SelectedObjects;
                if ((oDataTable != null))
                {
                    oForm = Aplication.Forms.Item(pVal.FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                    // Try
                    //oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, oDataTable.GetValue(aliasName, 0).ToString.Trim())
                    //If pVal.Row = 0 And oMatrix.RowCount = 1 Then
                    //    oEdit = oMatrix.GetCellSpecific(pVal.ColUID, 1)
                    //Else
                    row = oMatrix.GetCellFocus().rowIndex;


                    //        Catch ex As Exception
                    //    'MsgBox(ex.Message)
                    //End Try
                    if (!string.IsNullOrEmpty(columnMap) & !string.IsNullOrEmpty(alliasMap))
                    {
                        //oEdit = oMatrix.GetCellSpecific(columnMap, pVal.Row)
                        //If pVal.Row = 0 Then
                        //    oEdit = oMatrix.GetCellSpecific(columnMap, 1)
                        //Else
                        oEdit = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(columnMap, row);
                        // oMatrix.GetCellFocus.rowIndex) 'pVal.Row)


                        oEdit.Value = oDataTable.GetValue(alliasMap, 0).ToString().Trim();
                        // oMatrix.SetCellWithoutValidation(pVal.Row, columnMap, oDataTable.GetValue(alliasMap, 0).ToString.Trim())
                    }
                    if (!string.IsNullOrEmpty(columnMap1) & !string.IsNullOrEmpty(alliasMap1))
                    {
                        //oEdit = oMatrix.GetCellSpecific(columnMap, pVal.Row)
                        //If pVal.Row = 0 Then
                        //    oEdit = oMatrix.GetCellSpecific(columnMap, 1)
                        //Else
                        oEdit = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(columnMap1, row);
                        // oMatrix.GetCellFocus.rowIndex) 'pVal.Row)
                        // End If

                        oEdit.Value = oDataTable.GetValue(alliasMap1, 0).ToString().Trim();
                        // oMatrix.SetCellWithoutValidation(pVal.Row, columnMap, oDataTable.GetValue(alliasMap, 0).ToString.Trim())
                    }

                    if (!string.IsNullOrEmpty(columnMap2) & !string.IsNullOrEmpty(alliasMap2))
                    {
                        //oEdit = oMatrix.GetCellSpecific(columnMap, pVal.Row)
                        //If pVal.Row = 0 Then
                        //    oEdit = oMatrix.GetCellSpecific(columnMap, 1)
                        //Else
                        oEdit = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(columnMap2, row);
                        // oMatrix.GetCellFocus.rowIndex) 'pVal.Row)
                        // End If

                        oEdit.Value = oDataTable.GetValue(alliasMap2, 0).ToString().Trim();
                        // oMatrix.SetCellWithoutValidation(pVal.Row, columnMap, oDataTable.GetValue(alliasMap, 0).ToString.Trim())
                    }

                    if (!string.IsNullOrEmpty(columnMap3) & !string.IsNullOrEmpty(alliasMap3))
                    {
                        //oEdit = oMatrix.GetCellSpecific(columnMap, pVal.Row)
                        //If pVal.Row = 0 Then
                        //    oEdit = oMatrix.GetCellSpecific(columnMap, 1)
                        //Else
                        oEdit = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(columnMap3, row);
                        // oMatrix.GetCellFocus.rowIndex) 'pVal.Row)
                        // End If

                        oEdit.Value = oDataTable.GetValue(alliasMap3, 0).ToString().Trim();
                        // oMatrix.SetCellWithoutValidation(pVal.Row, columnMap, oDataTable.GetValue(alliasMap, 0).ToString.Trim())
                    }
                    try
                    {
                        oEdit = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(pVal.ColUID, row);
                        //pVal.Row)
                        //End If

                        oEdit.Value = oDataTable.GetValue(aliasName, 0).ToString().Trim();
                        //oMatrix.SetCellFocus(row, 5)
                        //oEdit = oMatrix.GetCellSpecific("V_7", row)
                        //oEdit.Active = True
                        //oMatrix.SetCellFocus(row, 5)

                    }
                    catch
                    {
                    }
                    //If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    //End If
                }
            }
            catch
            {
                //Aplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            }
        }

        public static void ChooFormlistEventOnMatrixWithoutValidation(SAPbouiCOM.Application Aplication, SAPbouiCOM.ItemEvent pVal,
            string matrixUID, string aliasName, string columnMap = "", string alliasMap = "", string columnMap1 = "", string alliasMap1 = "",
            string columnMap2 = "", string alliasMap2 = "", string columnMap3 = "", string alliasMap3 = "")
        {
            SAPbouiCOM.ChooseFromListEvent oCFLEvento = default(SAPbouiCOM.ChooseFromListEvent);
            SAPbouiCOM.DataTable oDataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            int row = 0;
            try
            {
                oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                oDataTable = oCFLEvento.SelectedObjects;
                if ((oDataTable != null))
                {
                    oForm = Aplication.Forms.Item(pVal.FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                    // Try
                    //oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, oDataTable.GetValue(aliasName, 0).ToString.Trim())
                    //If pVal.Row = 0 And oMatrix.RowCount = 1 Then
                    //    oEdit = oMatrix.GetCellSpecific(pVal.ColUID, 1)
                    //Else
                    row = oMatrix.GetCellFocus().rowIndex;


                    //        Catch ex As Exception
                    //    'MsgBox(ex.Message)
                    //End Try
                    if (!string.IsNullOrEmpty(columnMap) & !string.IsNullOrEmpty(alliasMap))
                    {
                        //'oEdit = oMatrix.GetCellSpecific(columnMap, pVal.Row)
                        //'If pVal.Row = 0 Then
                        //'    oEdit = oMatrix.GetCellSpecific(columnMap, 1)
                        //'Else
                        //oEdit = oMatrix.GetCellSpecific(columnMap, row) ' oMatrix.GetCellFocus.rowIndex) 'pVal.Row)


                        //oEdit.Value = oDataTable.GetValue(alliasMap, 0).ToString.Trim
                        oMatrix.SetCellWithoutValidation(pVal.Row, columnMap, oDataTable.GetValue(alliasMap, 0).ToString().Trim());
                    }
                    if (!string.IsNullOrEmpty(columnMap1) & !string.IsNullOrEmpty(alliasMap1))
                    {
                        //'oEdit = oMatrix.GetCellSpecific(columnMap, pVal.Row)
                        //'If pVal.Row = 0 Then
                        //'    oEdit = oMatrix.GetCellSpecific(columnMap, 1)
                        //'Else
                        //oEdit = oMatrix.GetCellSpecific(columnMap1, row) ' oMatrix.GetCellFocus.rowIndex) 'pVal.Row)
                        //' End If

                        //oEdit.Value = oDataTable.GetValue(alliasMap1, 0).ToString.Trim
                        oMatrix.SetCellWithoutValidation(pVal.Row, columnMap, oDataTable.GetValue(alliasMap, 0).ToString().Trim());
                    }

                    if (!string.IsNullOrEmpty(columnMap2) & !string.IsNullOrEmpty(alliasMap2))
                    {
                        //'oEdit = oMatrix.GetCellSpecific(columnMap, pVal.Row)
                        //'If pVal.Row = 0 Then
                        //'    oEdit = oMatrix.GetCellSpecific(columnMap, 1)
                        //'Else
                        //oEdit = oMatrix.GetCellSpecific(columnMap2, row) ' oMatrix.GetCellFocus.rowIndex) 'pVal.Row)
                        //' End If

                        //oEdit.Value = oDataTable.GetValue(alliasMap2, 0).ToString.Trim
                        oMatrix.SetCellWithoutValidation(pVal.Row, columnMap, oDataTable.GetValue(alliasMap, 0).ToString().Trim());
                    }

                    if (!string.IsNullOrEmpty(columnMap3) & !string.IsNullOrEmpty(alliasMap3))
                    {
                        //'oEdit = oMatrix.GetCellSpecific(columnMap, pVal.Row)
                        //'If pVal.Row = 0 Then
                        //'    oEdit = oMatrix.GetCellSpecific(columnMap, 1)
                        //'Else
                        //'oEdit = oMatrix.GetCellSpecific(columnMap3, row) ' oMatrix.GetCellFocus.rowIndex) 'pVal.Row)
                        //' End If

                        //oEdit.Value = oDataTable.GetValue(alliasMap3, 0).ToString.Trim
                        oMatrix.SetCellWithoutValidation(pVal.Row, columnMap, oDataTable.GetValue(alliasMap, 0).ToString().Trim());
                    }
                    try
                    {
                        //oEdit = oMatrix.GetCellSpecific(pVal.ColUID, row) 'pVal.Row)
                        //'End If

                        //oEdit.Value = oDataTable.GetValue(aliasName, 0).ToString.Trim


                        oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, oDataTable.GetValue(aliasName, 0).ToString().Trim());


                        //oMatrix.SetCellFocus(row, 5)
                        //oEdit = oMatrix.GetCellSpecific("V_7", row)
                        //oEdit.Active = True
                        //oMatrix.SetCellFocus(row, 5)

                    }
                    catch
                    {
                    }
                    //If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    //End If
                }
            }
            catch
            {
                //Aplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            }
        }

        public static void MessageRemove(string NemuId, SAPbouiCOM.Application SBO_Application, bool BubbleEvent)
        {
            if (NemuId == "1293")
            {
                if (SBO_Application.MessageBox("Bạn có chắc chắn xóa dòng chứng từ không", 2, "Có", "Không") == 2)
                {
                    BubbleEvent = false;
                }
            }
            if (NemuId == "1283")
            {
                if (SBO_Application.MessageBox("Bạn có chắc chắn xóa chứng từ không", 2, "Có", "Không") == 2)
                {
                    BubbleEvent = false;
                }
            }
        }
        public static void FormItem_Enabled(ref SAPbouiCOM.Form Form, bool Enable)
        {
            int Count = 0;
            System.Data.DataTable FormItem = Get_FormItem();
            if (FormItem == null)
                return;
            if (FormItem.Rows.Count <= 0)
                return;
            System.Data.DataRow[] Row = FormItem.Select("U_FORMUID='" + Form.UniqueID + "'");
            if (Row.Length <= 0) return;
            for (Count = 0; Count <= Row.Length - 1; Count++)
            {
                try
                {
                    Form.Items.Item(Row[Count]["U_ITEM_CODE"].ToString().Trim()).Enabled = Enable;
                }
                catch { }
            }
        }
        public static void FormItem_Enabled_New(ref SAPbouiCOM.Form Form, bool Enable)
        {
            System.Data.DataTable FormItem = Get_FormItem();
            if (FormItem == null)
                return;
            if (FormItem.Rows.Count <= 0)
                return;
            SAPbouiCOM.BoModeVisualBehavior Check = default(SAPbouiCOM.BoModeVisualBehavior);
            if (Enable == false)
                Check = SAPbouiCOM.BoModeVisualBehavior.mvb_False;
            else
                Check = SAPbouiCOM.BoModeVisualBehavior.mvb_True;
            int Count = 0;
            System.Data.DataRow[] Row = FormItem.Select("U_FORMUID='" + Form.UniqueID + "'");
            if (Row.Length <= 0) return;
            for (Count = 0; Count <= Row.Length - 1; Count++)
            {
                try
                {
                    Form.Items.Item(Row[Count]["U_ITEM_CODE"].ToString().Trim()).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, Check);
                }
                catch { }
            }
        }
        public static void cmdGetBPList(SAPbouiCOM.Application SBO)
        {
            SAPbobsCOM.Company oCompany11 = default(SAPbobsCOM.Company);
            SAPbouiCOM.Grid Grid = default(SAPbouiCOM.Grid);
            SAPbouiCOM.Form Form = default(SAPbouiCOM.Form);
            string SQL = null;
            int lRetCode = 0;
            int lErrCode = 0;
            string sErrMsg = null;
            SAPbouiCOM.EditText Edit = default(SAPbouiCOM.EditText);

            SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);

            SAPbouiCOM.DBDataSource Db = default(SAPbouiCOM.DBDataSource);

            oCompany11 = new SAPbobsCOM.Company();



            // Init Connection Properties
            oCompany11.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
            // = cmbDBType.SelectedIndex + 1
            oCompany11.Server = "10.15.106.105";
            //' change to your company server
            oCompany11.language = SAPbobsCOM.BoSuppLangs.ln_English;
            // change to your language
            oCompany11.UseTrusted = false;
            oCompany11.DbUserName = "sa";
            oCompany11.DbPassword = "123456";


            oCompany11.CompanyDB = "ASC_TEST";
            oCompany11.UserName = "montv";
            oCompany11.Password = "123456";



            //Try to connect
            lRetCode = oCompany11.Connect();

            // if the connection failed
            if (lRetCode != 0)
            {
                oCompany11.GetLastError(out lErrCode, out sErrMsg);
                //MsgBox("Connection Failed - ");
                //Interaction.MsgBox("Connection Failed - " + sErrMsg, MsgBoxStyle.Exclamation, "Default Connection Failed");
            }
            // if connected
            if (oCompany11.Connected)
            {
                //Me.Text = Me.Text & " - Connected to " & oCompany.CompanyDB
                //grpConn.Enabled = false
                //grpOrder.Enabled = True
                //LoadGui() ' Load data for UI elements like combo boxes

            }
            // SBO = oCompany11.get




            //oCompany = SBO_Application.Company.GetDICompany()
            // MsgBox(SBO.Company.UserName)
            fcp = (SAPbouiCOM.FormCreationParams)SBO.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            fcp.FormType = "anhqh11";
            fcp.UniqueID = "anhqh11";
            fcp.XmlData = LoadFromXML("FrmBP v1.0.srf");

            Form = SBO.Forms.AddEx(fcp);


            Db = Form.DataSources.DBDataSources.Add("@FPTPRO_ITEM");
            Grid = (SAPbouiCOM.Grid)Form.Items.Item("338").Specific;
            Edit = (SAPbouiCOM.EditText)Form.Items.Item("22").Specific;
            Edit.DataBind.SetBound(true, "@FPTPRO_ITEM", "U_ITEMCODE");
            SetChooseFormListToItem(SBO, Form, "22", "CFL_Customer", "2", "CardCode");

            SQL = "select CardCode, CardName from OCRD where cardname is not null ";

            Form.DataSources.DataTables.Add("MyDataTable");
            Form.DataSources.DataTables.Item(0).ExecuteQuery(SQL);
            Grid.DataTable = Form.DataSources.DataTables.Item("MyDataTable");
            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.DBDataSource psadd = default(SAPbouiCOM.DBDataSource);
            psadd = Form.DataSources.DBDataSources.Add("@FPTMACHINE_OLD");
            Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("38").Specific;
            Form.DataSources.UserDataSources.Add("X1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            Matrix.Columns.Add("X", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);

            Matrix.Columns.Item("X").DataBind.SetBound(true, "", "X1");
            psadd.Query();

            Matrix.Columns.Item("V_0").DataBind.SetBound(true, "@FPTMACHINE_OLD", "U_CRITERIA");
            Matrix.LoadFromDataSource();

            try
            {
            }
            catch
            {
                return;
            }

            oCompany.GetLastError(out lErrCode, out sErrMsg);

            if (lErrCode != 0)
            {
                //Interaction.MsgBox(sErrMsg);
            }
            else
            {

            }


        }

        public static void Mod_IPAddress()
        {
            //g_GetHostName = System.Net.Dns.GetHostName();
            ////Try
            ////    g_IP_Address = System.Net.Dns.GetHostAddresses(g_GetHostName).GetValue(1).ToString
            ////Catch ex As Exception
            ////    g_IP_Address = System.Net.Dns.GetHostAddresses(g_GetHostName).GetValue(0).ToString
            ////End Try
            //try
            //{
            //    g_IP_Address = System.Net.Dns.GetHostAddresses(g_GetHostName).GetValue(0).ToString();
            //    // .GetHostByName(strHostName).ToString() ' AddressList(0).ToString()                  

            //    if (g_IP_Address.LastIndexOf(".") <= 0)
            //    {
            //        g_IP_Address = System.Net.Dns.GetHostAddresses(g_GetHostName).GetValue(1).ToString();
            //        if (g_IP_Address.LastIndexOf(".") <= 0)
            //        {
            //            g_IP_Address = System.Net.Dns.GetHostAddresses(g_GetHostName).GetValue(2).ToString();
            //        }
            //    }

            //}
            //catch (Exception ex)
            //{
            //}

        }
        public static System.DateTime ConvertDatetime(string DateString, string sperate)
        {
            string strDate = DateString;
            strDate = string.Format("{0}{1}{2}{3}{4}", strDate.Substring(0, 4), sperate, strDate.Substring(4, 2), sperate, strDate.Substring(6, 2));
            return System.DateTime.Parse(strDate);
        }
        public static void SetCondition_For_CFL(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form11, string CFL_Name, SAPbouiCOM.Conditions oCons)
        {
            SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
            SAPbouiCOM.ChooseFromListCollection oCFLs = Form11.ChooseFromLists;

            try
            {


                oCFL = oCFLs.Item(CFL_Name);
                oCFL.SetConditions(oCons);

                //Form11 = Nothing
            }
            catch
            {
                //SBO_Application.StatusBar.SetText("Loi SBO_Application_AppEvent ! Loi : " & ex.ToString, _
                //            SAPbouiCOM.BoMessageTime.bmt_Short, _
                //            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            }

        }
        public static SAPbouiCOM.Conditions GetMultiCondition(ref SAPbouiCOM.Conditions Conditions, string[] Alias, string[] Value, SAPbouiCOM.BoConditionOperation[] Operation, SAPbouiCOM.BoConditionRelationship[] Relationship, string[] ConEndVal)
        {

            SAPbouiCOM.Condition Condition = default(SAPbouiCOM.Condition);
            Conditions = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
            for (int i = 0; i <= Alias.Length - 1; i++)
            {
                Condition = Conditions.Add();

                Condition.Operation = Operation[i];
                Condition.Alias = Alias[i];
                Condition.Operation = Operation[i];
                Condition.CondVal = Value[i];
                if (Operation[i] == SAPbouiCOM.BoConditionOperation.co_BETWEEN)
                {
                    Condition.CondEndVal = ConEndVal[i];
                }

                if (i != Alias.Length - 1)
                {
                    Condition.Relationship = Relationship[i];
                }
            }
            return Conditions;

        }
        public static void SetComboItem(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, string ItemName, string SQL)
        {
            SAPbouiCOM.ComboBox ComboBox = default(SAPbouiCOM.ComboBox);
            int Count = 0;
            try
            {
                Form.Items.Item(ItemName).DisplayDesc = true;
                ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item(ItemName).Specific;
                for (Count = ComboBox.ValidValues.Count - 1; Count >= 0; Count += -1)
                {
                    try
                    {
                        ComboBox.ValidValues.Remove(Count, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        //SBO_Application.MessageBox(ex.Message);
                    }
                }
                ComboBox_ValidValues_Add(ComboBox, ref SQL);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                //SBO_Application.MessageBox(ex.Message);
            }
        }
        public static void SetComboItemMatrix(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, SAPbouiCOM.Matrix Matrix, string ColName, string SQL)
        {
            SAPbouiCOM.Column ComboBox = default(SAPbouiCOM.Column);
            int Count = 0;
            try
            {
                ComboBox = Matrix.Columns.Item(ColName);
                for (Count = ComboBox.ValidValues.Count - 1; Count >= 0; Count--)
                {
                    try
                    {
                        ComboBox.ValidValues.Remove(Count, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    catch { }
                }
                Matrix_ComboBox_ValidValues_Add(ref ComboBox, SQL);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
            }
        }
        public static void CreateField(string name, string description, SAPbobsCOM.BoFieldTypes fieldDataType, SAPbobsCOM.BoFldSubTypes fieldSubType, int size, bool isMandatory, string tableName, SAPbobsCOM.Company oCompany, SAPbouiCOM.Application SBO_Application)
        {
            try
            {
                GC.Collect();
                SAPbobsCOM.UserFieldsMD oUserFieldsMD = default(SAPbobsCOM.UserFieldsMD);
                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                //Create Fields to table.
                oUserFieldsMD.TableName = tableName;
                oUserFieldsMD.Name = name;
                oUserFieldsMD.Description = description;
                oUserFieldsMD.Type = fieldDataType;
                oUserFieldsMD.SubType = fieldSubType;
                if (isMandatory == true)
                    oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
                else
                    oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tNO;
                int ErrCode = 0;
                string ErrMsg = "";
                if ((size > 0))
                {
                    oUserFieldsMD.EditSize = size;
                }
                if ((oUserFieldsMD.Add() != 0))
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    //SBO_Application.StatusBar.SetText(ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    SBO_Application.StatusBar.SetText("Fields was added successfully to " + tableName + " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                oUserFieldsMD = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.ToString() + "on filed :" + name, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static void CreateField(string name, string description, SAPbobsCOM.BoFieldTypes fieldDataType, SAPbobsCOM.BoFldSubTypes fieldSubType, int size, bool isMandatory, string tableName, string defaultValue, List<string[]> validValues, SAPbobsCOM.Company oCompany, SAPbouiCOM.Application SBO_Application)
        {
            try
            {
                GC.Collect();
                SAPbobsCOM.UserFieldsMD oUserFieldsMD = default(SAPbobsCOM.UserFieldsMD);
                //Create Fields to table.
                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = tableName;
                oUserFieldsMD.Name = name;
                oUserFieldsMD.Description = description;
                oUserFieldsMD.Type = fieldDataType;
                if (isMandatory == true)
                    oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
                else
                    oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tNO;
                if ((validValues != null))
                {
                    int i = 0;
                    for (i = 0; i <= validValues.Count - 1; i++)
                    {
                        oUserFieldsMD.ValidValues.Value = validValues[i].GetValue(0).ToString().Trim();
                        oUserFieldsMD.ValidValues.Description = validValues[i].GetValue(1).ToString().Trim();
                        if (i < validValues.Count - 1)
                            oUserFieldsMD.ValidValues.Add();
                    }
                }
                if (defaultValue != null)
                    oUserFieldsMD.DefaultValue = defaultValue;
                oUserFieldsMD.SubType = fieldSubType;
                if ((size > 0))
                    oUserFieldsMD.EditSize = size;
                int ErrCode = 0;
                string ErrMsg = "";
                if ((oUserFieldsMD.Add() != 0))
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    if (ErrCode != -2035) { }
                    //SBO_Application.StatusBar.SetText("Field: '" + name + "' " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                    SBO_Application.StatusBar.SetText("Fields '" + name + "' was added successfully to " + tableName + " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                oUserFieldsMD = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                //SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static void CreateTable(string name, string description, SAPbobsCOM.BoUTBTableType tableType, SAPbobsCOM.Company oCompany, SAPbouiCOM.Application SBO_Application)
        {
            try
            {
                GC.Collect();
                SAPbobsCOM.UserTablesMD oTable = default(SAPbobsCOM.UserTablesMD);
                oTable = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (oTable.GetByKey(name))
                {
                    oTable = null;
                    GC.Collect();
                    return;
                }

                oTable.TableName = name;
                oTable.TableDescription = description;
                oTable.TableType = tableType;
                int ErrCode = 0;
                string ErrMsg = "";
                if (oTable.Add() != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    //SBO_Application.StatusBar.SetText(ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                    SBO_Application.StatusBar.SetText("Table: " + name + " was added successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                oTable = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                //SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static void CreateObject(string objUID, string objName, string tableName, SAPbobsCOM.BoUDOObjType udoTypes, SAPbobsCOM.BoYesNoEnum canCancel, SAPbobsCOM.BoYesNoEnum canClose, SAPbobsCOM.BoYesNoEnum canCreateDefaultForm, SAPbobsCOM.BoYesNoEnum canDelete, SAPbobsCOM.BoYesNoEnum canFind, SAPbobsCOM.BoYesNoEnum canLog,
        SAPbobsCOM.BoYesNoEnum canYearTransfer, SAPbobsCOM.BoYesNoEnum manageSeries, SAPbobsCOM.Company oCompany, SAPbouiCOM.Application SBO_Application, List<string> listChild = null)
        {
            try
            {
                GC.Collect();
                SAPbobsCOM.UserObjectsMD oUserObjectMD = (SAPbobsCOM.UserObjectsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                if (oUserObjectMD.GetByKey(objUID))
                {
                    return;
                }
                oUserObjectMD.CanCancel = canCancel;
                oUserObjectMD.CanClose = canClose;
                oUserObjectMD.CanCreateDefaultForm = canCreateDefaultForm;
                oUserObjectMD.CanDelete = canDelete;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanLog = canLog;
                oUserObjectMD.CanYearTransfer = canYearTransfer;

                if ((listChild != null))
                {
                    string Child = null;
                    foreach (string Child_loopVariable in listChild)
                    {
                        Child = Child_loopVariable;
                        oUserObjectMD.ChildTables.Add();
                        oUserObjectMD.ChildTables.TableName = Child;
                    }
                }

                oUserObjectMD.Code = objUID;
                oUserObjectMD.ManageSeries = manageSeries;
                oUserObjectMD.Name = objName;
                oUserObjectMD.ObjectType = udoTypes;
                oUserObjectMD.TableName = tableName;
                int ErrCode = 0;
                string ErrMsg = "";
                if (oUserObjectMD.Add() != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    //SBO_Application.StatusBar.SetText(ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                    GC.Collect();
            }
            catch (Exception ex)
            {
                //SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static void CreateObject(string objUID, string objName, string tableName, SAPbobsCOM.BoUDOObjType udoTypes,
            string menuCaption, string menuID, int fatherMenuID, SAPbobsCOM.BoYesNoEnum menuItem, SAPbobsCOM.BoYesNoEnum enableEnhancedForm,
            SAPbobsCOM.BoYesNoEnum canCancel, SAPbobsCOM.BoYesNoEnum canClose, SAPbobsCOM.BoYesNoEnum canCreateDefaultForm, SAPbobsCOM.BoYesNoEnum canDelete, SAPbobsCOM.BoYesNoEnum canFind, SAPbobsCOM.BoYesNoEnum canLog,
            SAPbobsCOM.BoYesNoEnum canYearTransfer, SAPbobsCOM.BoYesNoEnum manageSeries, SAPbobsCOM.Company oCompany, SAPbouiCOM.Application SBO_Application, List<string> listChild = null)
        {
            try
            {
                GC.Collect();
                SAPbobsCOM.UserObjectsMD oUserObjectMD = (SAPbobsCOM.UserObjectsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                if (oUserObjectMD.GetByKey(objUID))
                    return;
                oUserObjectMD.CanCancel = canCancel;
                oUserObjectMD.CanClose = canClose;
                oUserObjectMD.CanCreateDefaultForm = canCreateDefaultForm;
                oUserObjectMD.CanDelete = canDelete;
                oUserObjectMD.CanFind = canFind;
                oUserObjectMD.CanLog = canLog;
                oUserObjectMD.CanYearTransfer = canYearTransfer;

                if (!string.IsNullOrEmpty(menuID))
                {
                    oUserObjectMD.MenuCaption = menuCaption;
                    oUserObjectMD.MenuItem = menuItem;
                    oUserObjectMD.MenuUID = menuID;
                    oUserObjectMD.FatherMenuID = fatherMenuID;
                    oUserObjectMD.EnableEnhancedForm = enableEnhancedForm;
                }
                if ((listChild != null))
                {
                    string Child = null;
                    foreach (string Child_loopVariable in listChild)
                    {
                        Child = Child_loopVariable;
                        oUserObjectMD.ChildTables.Add();
                        oUserObjectMD.ChildTables.TableName = Child;
                    }
                }

                oUserObjectMD.Code = objUID;
                oUserObjectMD.ManageSeries = manageSeries;
                oUserObjectMD.Name = objName;
                oUserObjectMD.ObjectType = udoTypes;
                oUserObjectMD.TableName = tableName;

                int ErrCode = 0;
                string ErrMsg = "";
                if (oUserObjectMD.Add() != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    SBO_Application.StatusBar.SetText(ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                    GC.Collect();
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void Get_Conditional(SAPbouiCOM.Application p_SBO_Application, ref SAPbouiCOM.Conditions p_Conditions, string p_Alias, string p_Value, SAPbouiCOM.BoConditionOperation p_Operation, string p_relation = "", string p_Value1 = "")
        {
            SAPbouiCOM.Condition p_Condition = default(SAPbouiCOM.Condition);
            try
            {
                p_Conditions = (SAPbouiCOM.Conditions)p_SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                p_Condition = p_Conditions.Add();
                p_Condition.Alias = p_Alias;
                p_Condition.CondVal = p_Value;
                p_Condition.Operation = p_Operation;
                if (p_Operation == SAPbouiCOM.BoConditionOperation.co_BETWEEN | p_Operation == SAPbouiCOM.BoConditionOperation.co_NOT_BETWEEN)
                    p_Condition.CondEndVal = p_Value1;
            }
            catch (Exception ex)
            {
                p_SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void Matrix_ComboBox_ValidValues_Add(ref SAPbouiCOM.Column p_Col, string p_SQL)
        {
            int p_Count;
            GC.Collect();
            SAPbouiCOM.DataTable oDataTable = GetSapDataTable(p_SQL);

            if (!oDataTable.IsEmpty)
                for (p_Count = 0; p_Count < oDataTable.Rows.Count - 1; p_Count++)
                    p_Col.ValidValues.Add(oDataTable.GetValue(0, p_Count).ToString().Trim(), oDataTable.GetValue(1, p_Count).ToString().Trim());
            p_Col.DisplayDesc = true;
        }

        public static string DateConvertSO(string Date, bool AddKey = true)
        {
            string Value;
            try
            {
                if (AddKey == true)
                {
                    Value = "#" + DateTime.Parse(Date).ToString("yyyy/MMM/dd") + "#";
                }
                else
                {
                    Value = DateTime.Parse(Date).ToString("yyyy/MMM/dd");
                }



            }
            catch
            {
                Value = "";

            }

            return Value;

        }

        public static string DateConvert(string Date, bool AddKey = true)
        {
            string functionReturnValue = null;
            string Value = null;
            Value = Date;
            if (AddKey)
            {
                if (Value.Length > 8)
                {
                    functionReturnValue = "#" + Value + "#";
                    return functionReturnValue;
                    //Value = Replace(Value, "/", "", 1)
                }
                if (Value.Length == 8)
                {
                    functionReturnValue = "#" + Value.Substring(0, 4) + "/" + Value.Substring(5, 2) + "/" + Value.Substring(Value.Length - 2) + "#";
                    return functionReturnValue;
                    //Value = Replace(Value, "/", "", 1)
                }
            }
            else
            {
                if (Value.Length > 8)
                {
                    functionReturnValue = Value;
                    return functionReturnValue;
                    //Value = Replace(Value, "/", "", 1)
                }
                if (Value.Length == 8)
                {
                    functionReturnValue = Value.Substring(0, 4) + "/" + Value.Substring(5, 2) + "/" + Value.Substring(Value.Length - 2);
                    return functionReturnValue;
                    //Value = Replace(Value, "/", "", 1)
                }
            }
            return functionReturnValue;
        }
        #endregion

        #region  New Method
        public static SAPbouiCOM.UserDataSource AddUserDataSource(ref SAPbouiCOM.Form form, string usSource, SAPbouiCOM.BoDataType type, int size)
        {
            SAPbouiCOM.UserDataSource functionReturnValue = default(SAPbouiCOM.UserDataSource);
            //Dim us As SAPbouiCOM.DataTable
            try
            {
                functionReturnValue = form.DataSources.UserDataSources.Add(usSource, type, size);
            }
            catch (Exception ex)
            {
                functionReturnValue = (SAPbouiCOM.UserDataSource)form.DataSources.DataTables.Item(usSource);
            }
            return functionReturnValue;
        }
        //public static void FillDataChooseFormLisToColumn(SAPbouiCOM.Application cApplication, SAPbouiCOM.ItemEvent pVal, string colAlias, string colMap = "", string colAliasMap = "")
        //{
        //    SAPbouiCOM.DataTable DataTable = default(SAPbouiCOM.DataTable);
        //    SAPbouiCOM.EditText oEdit = default(SAPbouiCOM.EditText);
        //    SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
        //    SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
        //    SAPbouiCOM.IChooseFromListEvent MakerCFLEvnt = default(SAPbouiCOM.IChooseFromListEvent);
        //    try
        //    {
        //        oForm = cApplication.Forms.Item(pVal.FormUID);
        //        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
        //        MakerCFLEvnt = (SAPbouiCOM.ChooseFromListEvent)pVal;
        //        if (MakerCFLEvnt.SelectedObjects == null)
        //            return;
        //        DataTable = MakerCFLEvnt.SelectedObjects;
        //        if ((DataTable != null))
        //        {
        //            if (!string.IsNullOrEmpty(colMap) & !string.IsNullOrEmpty(colAliasMap))
        //            {
        //                try
        //                {
        //                    oEdit = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(colMap, pVal.Row);
        //                    oEdit.Value = DataTable.GetValue(colAliasMap, 0).ToString().Trim();
        //                }
        //                catch { }
        //            }
        //            try
        //            {
        //                oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, DataTable.GetValue(colAlias, 0).ToString().Trim());
        //            }
        //            catch { }
        //            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        //        }
        //        DataTable = null;
        //        oForm = null;
        //    }
        //    catch { }
        //}
        public static void FillDataChooseFormLisToColumn(SAPbouiCOM.Application cApplication, SAPbouiCOM.ItemEvent pVal, string colAlias, params string[] colMap_ColVal)
        {
            SAPbouiCOM.DataTable DataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.IChooseFromListEvent MakerCFLEvnt = default(SAPbouiCOM.IChooseFromListEvent);
            try
            {
                oForm = cApplication.Forms.Item(pVal.FormUID);
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                MakerCFLEvnt = (SAPbouiCOM.ChooseFromListEvent)pVal;
                if (MakerCFLEvnt.SelectedObjects == null)
                    return;
                DataTable = MakerCFLEvnt.SelectedObjects;
                if ((DataTable != null))
                {
                    if (colMap_ColVal != null)
                        if (colMap_ColVal.Length > 1)
                            for (int i = 0; i < colMap_ColVal.Length; i += 2)
                            {
                                try
                                {
                                    oMatrix.SetCellWithoutValidation(pVal.Row, colMap_ColVal[i], DataTable.GetValue(colMap_ColVal[i + 1], 0).ToString().Trim());
                                }
                                catch { }
                            }
                    try
                    {
                        oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, DataTable.GetValue(colAlias, 0).ToString().Trim());
                        //oMatrix.AddRow();
                    }
                    catch { }
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                DataTable = null;
                oForm = null;
            }
            catch { }
        }
        public static void FillDataChooseFormLisToColumn1(SAPbouiCOM.Application cApplication, SAPbouiCOM.ItemEvent pVal, string colAlias, params string[] colMap_ColVal)
        {
            SAPbouiCOM.DataTable DataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.IChooseFromListEvent MakerCFLEvnt = default(SAPbouiCOM.IChooseFromListEvent);
            try
            {
                oForm = cApplication.Forms.Item(pVal.FormUID);
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                MakerCFLEvnt = (SAPbouiCOM.ChooseFromListEvent)pVal;
                if (MakerCFLEvnt.SelectedObjects == null)
                    return;
                DataTable = MakerCFLEvnt.SelectedObjects;
                if ((DataTable != null))
                {
                    if (colMap_ColVal != null)
                        if (colMap_ColVal.Length > 1)
                            //for (int i = 0; i < colMap_ColVal.Length; i += 2)
                            //{
                            try
                            {
                                oMatrix.SetCellWithoutValidation(pVal.Row, colMap_ColVal[0], DataTable.GetValue(colMap_ColVal[1], 0).ToString().Trim()
                                    + " " + DataTable.GetValue(colMap_ColVal[2], 0).ToString().Trim() + " " + DataTable.GetValue(colMap_ColVal[3], 0).ToString().Trim());
                            }
                            catch { }
                    //}
                    try
                    {
                        oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, DataTable.GetValue(colAlias, 0).ToString().Trim());
                        oMatrix.AddRow();
                    }
                    catch { }
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                DataTable = null;
                oForm = null;
            }
            catch { }
        }
        public static void FillDataChooseFormLisToItem(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.ItemEvent pVal, string itemMap = "", string col_AliasMap = "", bool formUpd = true)
        {
            SAPbouiCOM.DataTable DataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.IChooseFromListEvent MakerCFLEvnt = default(SAPbouiCOM.IChooseFromListEvent);
            try
            {
                MakerCFLEvnt = (SAPbouiCOM.ChooseFromListEvent)pVal;
                if (MakerCFLEvnt.SelectedObjects == null)
                    return;
                DataTable = MakerCFLEvnt.SelectedObjects;
                if (!DataTable.IsEmpty)
                {
                    SAPbouiCOM.Form oForm = SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.EditText oEdit;
                    if (!string.IsNullOrWhiteSpace(itemMap) & !string.IsNullOrWhiteSpace(col_AliasMap))
                    {
                        try
                        {
                            oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(itemMap).Specific;
                            oEdit.Value = DataTable.GetValue(col_AliasMap, 0).ToString();
                        }
                        catch { }
                    }
                    try
                    {
                        oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific;
                        oEdit.Value = DataTable.GetValue(oEdit.ChooseFromListAlias, 0).ToString();
                    }
                    catch { }
                    //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE & formUpd)
                    //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("FillDataChooseFormLisToItem: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static void MatrixDataBind(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, string tableName, string Matrix_Name, bool AutoResize = false)
        {
            if (string.IsNullOrWhiteSpace(tableName)) return;
            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            try
            {
                Matrix = (SAPbouiCOM.Matrix)Form.Items.Item(Matrix_Name).Specific;
                for (int i = 1; i <= Matrix.Columns.Count - 1; i++)
                    if (!string.IsNullOrEmpty(Matrix.Columns.Item(i).Description))
                    {
                        try
                        {
                            Matrix.Columns.Item(i).DataBind.SetBound(true, tableName, Matrix.Columns.Item(i).Description);
                            if (Matrix.Columns.Item(i).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                                Matrix.Columns.Item(i).DisplayDesc = true;
                            else
                                Matrix.Columns.Item(i).DisplayDesc = false;
                        }
                        catch { }
                    }
                if (AutoResize == true)
                    Matrix.AutoResizeColumns();
            }
            catch (Exception ex) { SBO_Application.StatusBar.SetText("MatrixDataBind: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
        }
        public static void MatrixDataBindTable(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, string tableName, string Matrix_Name, bool AutoResize = false)
        {
            if (string.IsNullOrWhiteSpace(tableName)) return;
            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            try
            {
                Matrix = (SAPbouiCOM.Matrix)Form.Items.Item(Matrix_Name).Specific;
                for (int i = 1; i <= Matrix.Columns.Count - 1; i++)
                    if (!string.IsNullOrEmpty(Matrix.Columns.Item(i).Description))
                    {
                        try
                        {
                            Matrix.Columns.Item(i).DataBind.Bind(tableName, Matrix.Columns.Item(i).Description);
                            if (Matrix.Columns.Item(i).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                                Matrix.Columns.Item(i).DisplayDesc = true;
                            else
                                Matrix.Columns.Item(i).DisplayDesc = false;
                        }
                        catch { }
                    }
                if (AutoResize == true)
                    Matrix.AutoResizeColumns();
            }
            catch (Exception ex) { SBO_Application.StatusBar.SetText("MatrixDataBind: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
        }
        public static void SetChooseFormListColumn(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form form, string Matrix_Name, string Column_Name, string UniqueID, string ObjectType, string Alias_Field_Name)
        {
            SAPbouiCOM.ChooseFromList oCFL1 = default(SAPbouiCOM.ChooseFromList);
            SAPbouiCOM.ChooseFromListCollection oCFLs1 = default(SAPbouiCOM.ChooseFromListCollection);
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = default(SAPbouiCOM.ChooseFromListCreationParams);
            SAPbouiCOM.Column Column = default(SAPbouiCOM.Column);
            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            try
            {
                oCFLs1 = form.ChooseFromLists;
                Matrix = (SAPbouiCOM.Matrix)form.Items.Item(Matrix_Name).Specific;
                Column = Matrix.Columns.Item(Column_Name);
                if (Column.ChooseFromListUID.ToString() != UniqueID)
                {
                    oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = ObjectType;
                    oCFLCreationParams.UniqueID = UniqueID;
                    oCFL1 = oCFLs1.Add(oCFLCreationParams);
                    Column.ChooseFromListUID = UniqueID;
                    if (!string.IsNullOrEmpty(Alias_Field_Name.ToString()))
                        Column.ChooseFromListAlias = Alias_Field_Name;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("SetChooseFormListColumn: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        public static void SetChooseFormListToItem(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form form, string Item_Name, string UniqueID, string ObjectType, string Alias_Field_Name)
        {
            SAPbouiCOM.ChooseFromList oCFL1 = default(SAPbouiCOM.ChooseFromList);
            SAPbouiCOM.ChooseFromListCollection oCFLs1 = default(SAPbouiCOM.ChooseFromListCollection);
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = default(SAPbouiCOM.ChooseFromListCreationParams);
            SAPbouiCOM.EditText EditText = default(SAPbouiCOM.EditText);
            try
            {
                oCFLs1 = form.ChooseFromLists;
                EditText = (SAPbouiCOM.EditText)form.Items.Item(Item_Name).Specific;
                if (EditText.ChooseFromListUID.ToString() != UniqueID)
                {
                    oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = ObjectType;
                    oCFLCreationParams.UniqueID = UniqueID;
                    oCFL1 = oCFLs1.Add(oCFLCreationParams);
                    EditText.ChooseFromListUID = UniqueID;
                    if (!string.IsNullOrEmpty(Alias_Field_Name.ToString()))
                        EditText.ChooseFromListAlias = Alias_Field_Name;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("SetChooseFormListToItem: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        public static void SetChooseFormListToItem1(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form form, string Item_Name, string UniqueID, string ObjectType, string Alias_Field_Name)
        {
            SAPbouiCOM.ChooseFromList oCFL1 = default(SAPbouiCOM.ChooseFromList);
            SAPbouiCOM.ChooseFromListCollection oCFLs1 = default(SAPbouiCOM.ChooseFromListCollection);
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = default(SAPbouiCOM.ChooseFromListCreationParams);
            SAPbouiCOM.EditText EditText = default(SAPbouiCOM.EditText);
            try
            {
                oCFLs1 = form.ChooseFromLists;
                EditText = (SAPbouiCOM.EditText)form.Items.Item(Item_Name).Specific;
                if (EditText.ChooseFromListUID.ToString() != UniqueID)
                {
                    oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = ObjectType;
                    oCFLCreationParams.UniqueID = UniqueID;
                    oCFL1 = oCFLs1.Add(oCFLCreationParams);
                    EditText.ChooseFromListUID = UniqueID;
                    if (!string.IsNullOrEmpty(Alias_Field_Name.ToString()))
                        EditText.ChooseFromListAlias = Alias_Field_Name;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("SetChooseFormListToItem: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        public static void AddCFL_DI(string objectType, string FieldName)
        {
            SAPbobsCOM.ChooseFromList oCFL_DI = oCompany.GetBusinessObject(BoObjectTypes.oChooseFromList);
            try
            {
                if (oCFL_DI.GetByKey(objectType) == false)
                    throw new Exception("CFL not defined");
                oCFL_DI.ChooseFromList_Lines.Add();
                oCFL_DI.ChooseFromList_Lines.SetCurrentLine(oCFL_DI.ChooseFromList_Lines.Count - 1);
                oCFL_DI.ChooseFromList_Lines.FieldNo = FieldName;
                try
                {
                    oCFL_DI.Update();
                }
                catch (Exception ex)
                {
                    SBO_Application.StatusBar.SetText(ex.Message);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message);
            }
        }

        public static void SetLinkButtonToItem(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form form, string item_Link, string item_Val, SAPbouiCOM.BoLinkedObject objType, string user_Object)
        {
            try
            {
                SAPbouiCOM.Item oItem = default(SAPbouiCOM.Item);
                try
                {
                    oItem = form.Items.Add(item_Link, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                }
                catch
                {
                    oItem = form.Items.Item(item_Link);
                }
                oItem.Top = form.Items.Item(item_Val).Top - 1;
                oItem.Left = form.Items.Item(item_Val).Left - 20;
                oItem.LinkTo = item_Val;
                if (!string.IsNullOrWhiteSpace(user_Object))
                    ((SAPbouiCOM.LinkedButton)form.Items.Item(item_Link).Specific).LinkedObjectType = user_Object;
                else
                    ((SAPbouiCOM.LinkedButton)form.Items.Item(item_Link).Specific).LinkedObject = objType;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("SetLinkButtonToItem: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        public static void SetLinkButtonToColumn(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form form, string item, string col, SAPbouiCOM.BoLinkedObject objType, string user_Object)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item(item).Specific;
                if (!string.IsNullOrWhiteSpace(user_Object))
                    ((SAPbouiCOM.LinkedButton)oMatrix.Columns.Item(col).ExtendedObject).LinkedObjectType = user_Object;
                else
                    ((SAPbouiCOM.LinkedButton)oMatrix.Columns.Item(col).ExtendedObject).LinkedObject = objType;

            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("SetLinkButtonToColumn: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        public static void AddRowMatrix(SAPbouiCOM.Application Aplication, string formUID, string matrixUID, string oDBDataSourceName, params string[] col_val)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.DBDataSource oDBDataSource = default(SAPbouiCOM.DBDataSource);
            try
            {
                oForm = Aplication.Forms.Item(formUID);
                oForm.Freeze(true);
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                oDBDataSource = oForm.DataSources.DBDataSources.Item(oDBDataSourceName);
                if (oMatrix.RowCount > 0)
                {
                    oMatrix.FlushToDataSource();
                    oDBDataSource.InsertRecord(oDBDataSource.Size);
                }
                if (oDBDataSource.Size == 0)
                    oDBDataSource.InsertRecord(oDBDataSource.Size);
                if (col_val != null)
                    if (col_val.Length > 1)
                        for (int i = 0; i < col_val.Length; i += 2)
                            oDBDataSource.SetValue(col_val[i], oDBDataSource.Size - 1, col_val[i + 1]);
                oMatrix.LoadFromDataSource();
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE & oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                if ((oForm != null))
                    oForm.Freeze(false);
                Aplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static void DeleteRowMatrix(SAPbouiCOM.Application Aplication, string formUID, string matrixUID, string oDBDataSourceName, params string[] col_val)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.DBDataSource oDBDataSource = default(SAPbouiCOM.DBDataSource);
            try
            {
                oForm = Aplication.Forms.Item(formUID);
                oForm.Freeze(true);
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                oDBDataSource = oForm.DataSources.DBDataSources.Item(oDBDataSourceName);
                if (oMatrix.RowCount > 0)
                {
                    oMatrix.FlushToDataSource();
                    oMatrix.LoadFromDataSource();
                    oDBDataSource.RemoveRecord(oMatrix.RowCount - 1);
                }
                oMatrix.DeleteRow(oMatrix.RowCount);
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE & oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                if ((oForm != null))
                    oForm.Freeze(false);
                Aplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static void HeaderDataBind(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, string TableName, string DocNum = "")
        {
            SAPbouiCOM.ComboBox ComboBox = default(SAPbouiCOM.ComboBox);
            SAPbouiCOM.EditText EditText = default(SAPbouiCOM.EditText);
            SAPbouiCOM.CheckBox CheckBox = default(SAPbouiCOM.CheckBox);
            try
            {
                for (int Count = 0; Count <= Form.Items.Count - 1; Count++)
                {
                    switch (Form.Items.Item(Count).Type)
                    {
                        case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                            try
                            {
                                Form.Items.Item(Count).DisplayDesc = true;
                                ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item(Count).Specific;
                                ComboBox.DataBind.SetBound(true, TableName, Form.Items.Item(Count).UniqueID);
                            }
                            catch { }
                            break;
                        case (SAPbouiCOM.BoFormItemTypes.it_EDIT):
                            try
                            {
                                EditText = (SAPbouiCOM.EditText)Form.Items.Item(Count).Specific;
                                EditText.DataBind.SetBound(true, TableName, Form.Items.Item(Count).UniqueID);
                            }
                            catch { }
                            break;
                        case (SAPbouiCOM.BoFormItemTypes.it_EXTEDIT):
                            try
                            {
                                EditText = (SAPbouiCOM.EditText)Form.Items.Item(Count).Specific;
                                EditText.DataBind.SetBound(true, TableName, Form.Items.Item(Count).UniqueID);
                            }
                            catch { }
                            break;
                        case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                            try
                            {
                                CheckBox = (SAPbouiCOM.CheckBox)Form.Items.Item(Count).Specific;
                                CheckBox.DataBind.SetBound(true, TableName, Form.Items.Item(Count).UniqueID);
                            }
                            catch { }
                            break;
                    }
                }

                if (!string.IsNullOrWhiteSpace(DocNum))
                    Form.DataBrowser.BrowseBy = DocNum;
                Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                Form.PaneLevel = 1;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Method save form to file (*.srf or *.xml)
        /// Create by khanhvv
        /// </summary>
        /// <param name="oForm">Object SAPbouiCOM Form</param>
        /// <param name="fullFileName">Full file name (*.srf)</param>
        /// <remarks></remarks>
        /// 
        public struct MyStruct
        {
            public string StringData { get; set; }
            public string StringData1 { get; set; }
            public string StringData2 { get; set; }
        };

        public static void GetItemId(SAPbouiCOM.Form oForm, ref string uid)
        {
            //string uid = "";
            //if (!System.IO.Directory.Exists(path)) return;
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            string sXmlString = null;
            sXmlString = oForm.GetAsXML();
            oXmlDoc.LoadXml(sXmlString);

            System.Xml.XmlNodeList elemList = oXmlDoc.GetElementsByTagName("item");
            for (int i = 0; i < elemList.Count; i++)
            {
                try
                {
                    if (elemList[i].Attributes != null && elemList[i].Attributes["description"].Value != null)
                        if (elemList[i].Attributes["description"].Value == "Thêm dòng")
                            if (elemList[i].Attributes["type"].Value == "16")
                            {
                                uid = elemList[i].Attributes["uid"].Value;
                                return;
                            }
                }
                catch
                {
                    continue;
                }
            }
        }

        public static void GetMatrixId(SAPbouiCOM.Form oForm, ref List<MyStruct> listColumnId)
        {
            //string uid = "";
            //if (!System.IO.Directory.Exists(path)) return;
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            string sXmlString = null;
            sXmlString = oForm.GetAsXML();
            oXmlDoc.LoadXml(sXmlString);
            MyStruct t = new MyStruct();
            System.Xml.XmlNodeList elemList = oXmlDoc.GetElementsByTagName("item");
            System.Xml.XmlNodeList elemList1 = oXmlDoc.GetElementsByTagName("databind");
            for (int i = 0; i < elemList.Count; i++)
            {
                try
                {
                    if (elemList[i].Attributes != null && elemList[i].Attributes["type"].Value != null)
                        if (elemList[i].Attributes["type"].Value == "127")
                            for (int j = 0; j < elemList1.Count; j++)
                            {
                                try
                                {
                                    if (elemList1[j].Attributes != null && elemList1[j].Attributes["alias"].Value != null
                                        && elemList1[j].Attributes["table"].Value != null)
                                    {
                                        t.StringData = elemList[i].Attributes["uid"].Value;
                                        t.StringData1 = elemList1[j].Attributes["alias"].Value;
                                        t.StringData2 = elemList1[j].Attributes["table"].Value;
                                        listColumnId.Add(t);
                                    }
                                }
                                catch { continue; }
                            }
                }
                catch { continue; }
            }
        }
        public static string GetFormUnique(SAPbouiCOM.Form oForm)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            string sXmlString = null;
            sXmlString = oForm.GetAsXML();
            oXmlDoc.LoadXml(sXmlString);
            string FormID = "";
            try
            {
                FormID = oXmlDoc.SelectSingleNode("Application/forms/action/form/@FormType").Value;
            }
            catch { }
            return FormID;
        }
        public static void GetTable(SAPbouiCOM.Form oForm, ref List<string> table)
        {
            //string uid = "";
            //if (!System.IO.Directory.Exists(path)) return;
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            string sXmlString = null;
            sXmlString = oForm.GetAsXML();
            oXmlDoc.LoadXml(sXmlString);
            System.Xml.XmlNodeList elemList = oXmlDoc.GetElementsByTagName("item");
            System.Xml.XmlNodeList elemList1 = oXmlDoc.GetElementsByTagName("databind");
            for (int i = 0; i < elemList.Count; i++)
            {
                try
                {
                    if (elemList[i].Attributes != null && elemList[i].Attributes["type"].Value != null)
                        if (elemList[i].Attributes["type"].Value == "127")
                            for (int j = 0; j < elemList1.Count; j++)
                            {
                                try
                                {
                                    if (elemList1[j].Attributes != null
                                        && elemList1[j].Attributes["table"].Value != null && elemList1[j].Attributes["table"].Value != "")
                                    {
                                        table.Add(elemList1[j].Attributes["table"].Value);
                                    }
                                }
                                catch { continue; }
                            }
                }
                catch { continue; }
            }
        }
        public static void AddLineUDT(SAPbouiCOM.ItemEvent pVal, bool BubbleEvent)
        {

            //SAPbouiCOM.Form oForm = null;
            //int count ;
            //string ItemId = null;
            //List<string> listMatrixId = new List<string>();
            //List<Table> listColumnId = new List<Table>();
            //SAPbouiCOM.EditText oEdit = null;
            //try
            //{
            //    oForm = SBO_Application.Forms.Item(pVal.FormUID);
            //    if (checkHasForm) return;
            //    if (oForm == null) return;
            //    GetItemId(oForm, ref ItemId);
            //    GetMatrixId(oForm, ref listColumnId);
            //    oForm.Freeze(true);
            //    oEdit = oForm.Items.Item(ItemId).Specific;
            //    if (oEdit.Value == "" || pVal.Row != 1)
            //    {
            //        oForm.Freeze(false);
            //        return;
            //    }
            //    count = int.Parse(oEdit.Value.Replace(".000000", ""));
            //    if (pVal.ItemUID == "0_U_G")
            //        for (int i = 0; i <= count; i++)
            //            AddRowMatrix(SBO_Application, oForm.UniqueID, "0_U_G", "@MANGNHANH1", "U_Gio", "");
            //    if (pVal.ItemUID == "1_U_G")
            //        for (int i = 0; i <= count; i++)
            //            AddRowMatrix(SBO_Application, oForm.UniqueID, "1_U_G", "@MANGNHANH2", "U_Gio", "");
            //    oForm.Items.Item(ItemId).Click();
            //    oEdit.Value = "";
            //    oForm.Freeze(false);
            //}
            //catch (Exception ex)
            //{
            //    oForm.Freeze(false);
            //    SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //}
        }
        public static void CreateMenuR(string UniqueID, string Name)
        {
            SAPbouiCOM.MenuItem oMenuItem;
            SAPbouiCOM.Menus oMenus;
            SAPbouiCOM.MenuCreationParams oCreationPackage;
            oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);


            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            oCreationPackage.UniqueID = UniqueID;
            oCreationPackage.String = Name;
            oCreationPackage.Enabled = true;

            oMenuItem = SBO_Application.Menus.Item("1280");
            oMenus = oMenuItem.SubMenus;
            if (!oMenuItem.SubMenus.Exists(Name))
            {
                try
                {

                    oMenus.AddEx(oCreationPackage);
                }
                catch
                {

                }
            }

        }
        public static void RemoveMenuR(string UniqueID)
        {
            SAPbouiCOM.MenuItem oMenuItem;
            SAPbouiCOM.Menus oMenus;
            oMenuItem = SBO_Application.Menus.Item("1280");
            oMenus = oMenuItem.SubMenus;
            if (oMenuItem.SubMenus.Exists(UniqueID))
            {
                try
                {

                    oMenus.RemoveEx(UniqueID);
                }
                catch
                {

                }
            }

        }
        public static void SaveFormAsXML(SAPbouiCOM.Form oForm, string path, string fullFileName)
        {
            if (!System.IO.Directory.Exists(path)) return;
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            string sXmlString = null;
            sXmlString = oForm.GetAsXML();
            oXmlDoc.LoadXml(sXmlString);
            oXmlDoc.Save((path + "\\" + fullFileName));
        }
        #endregion

        #region New functions
        public static bool MenuExist(string MenuUID)
        {
            bool flag = false;
            try
            {
                SAPbouiCOM.Menus oMenus = SBO_Application.Menus;
                flag = oMenus.Exists(MenuUID);
                oMenus = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                //p_cApplication.MessageBox("Lỗi MenuExist : " & ex.ToString)
                SBO_Application.StatusBar.SetText("Error menu exist : " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return flag;
        }
        public static void AddMenuItems(string menu_Type, string menu_Items, string menu_UID, string menu_Name, int position, string patch)
        {
            try
            {
                // The menus collection
                SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);
                // The new menu item
                SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);
                if (MenuExist(menu_UID))
                {
                    return;
                }
                // Get the menus collection from the application
                oMenus = SBO_Application.Menus;

                SAPbouiCOM.MenuCreationParams oCreationPackage = default(SAPbouiCOM.MenuCreationParams);

                oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                oMenuItem = SBO_Application.Menus.Item(menu_Items);
                // oMenuItem = SBO_Application.Menus.Item("43531");
                oMenus = oMenuItem.SubMenus;

                // New menu parameters
                switch (menu_Type)
                {
                    case "String":
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        break;
                    case "Popup":
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                        break;
                }
                oCreationPackage.UniqueID = menu_UID;
                oCreationPackage.String = menu_Name;
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = position;
                oCreationPackage.Image = patch;

                // If the manu already exists this code will fail
                try
                {
                    oMenus.AddEx(oCreationPackage);
                }
                catch { }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        public static void RemoveMenu(string menu)
        {
            try
            {
                SBO_Application.Menus.RemoveEx(menu);
            }
            catch { }
        }
        public static long DisplayDate(SAPbouiCOM.ItemEvent pVal, ref string sErrDesc)
        {
            long functionReturnValue = 0;
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            SAPbouiCOM.Form oFrmParent = default(SAPbouiCOM.Form);
            SAPbouiCOM.Item oItem = default(SAPbouiCOM.Item);
            SAPbouiCOM.EditText oEdit = default(SAPbouiCOM.EditText);
            SAPbouiCOM.StaticText oTxt = default(SAPbouiCOM.StaticText);
            SAPbouiCOM.Button oButton = default(SAPbouiCOM.Button);
            SAPbouiCOM.FormCreationParams creationPackage = default(SAPbouiCOM.FormCreationParams);
            int iCount = 0;
            string sFuncName = string.Empty;

            try
            {
                oFrmParent = SBO_Application.Forms.Item(pVal.FormUID);
                //Check whether the form exists.If exists then close the form
                for (iCount = 0; iCount <= SBO_Application.Forms.Count - 1; iCount++)
                {
                    oForm = SBO_Application.Forms.Item(iCount);
                    if (oForm.UniqueID == "dDate")
                    {
                        oForm.Close();
                        break; // TODO: might not be correct. Was : Exit For
                    }
                }
                //Add Form
                creationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "dDate";
                creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_FixedNoTitle;
                creationPackage.FormType = "OBT_dDate";
                oForm = SBO_Application.Forms.AddEx(creationPackage);
                var _with1 = oForm;
                oForm.Width = 300;
                oForm.Height = 100;
                //if (oFrmParent == null)
                //{
                //    _with1.Left = (Screen.PrimaryScreen.WorkingArea.Width - oForm.Width) / 2;
                //    _with1.Top = (Screen.PrimaryScreen.WorkingArea.Height - oForm.Height) / 3;
                //}
                //else
                //{
                //    _with1.Left = ((oFrmParent.Left * 2) + oFrmParent.Width - oForm.Width) / 2;
                //    _with1.Top = ((oFrmParent.Top * 2) + oFrmParent.Height - oForm.Height) / 2;
                //}

                //Add Label
                oItem = oForm.Items.Add("3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = 10;
                oItem.Left = 9;
                oItem.Width = 65;
                oTxt = oItem.Specific;
                oTxt.Caption = "FromDate";
                oItem = oForm.Items.Add("4", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = 30;
                oItem.Left = 9;
                oItem.Width = 65;
                oTxt = oItem.Specific;
                oTxt.Caption = "ToDate";
                //Add Field
                oItem = oForm.Items.Add("FromDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Top = 10;
                oItem.Left = 89;
                oItem.Width = 150;
                oItem = oForm.Items.Add("ToDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Top = 30;
                oItem.Left = 89;
                oItem.Width = 150;
                //Add Button
                oItem = oForm.Items.Add("btOK", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Top = 60;
                oItem.Left = 9;
                oItem.Width = 65;
                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Ok";

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Top = 60;
                oItem.Left = 89;
                oItem.Width = 65;
                try
                {
                    oForm.DataSources.UserDataSources.Add("FromDate", SAPbouiCOM.BoDataType.dt_DATE);
                    oForm.DataSources.UserDataSources.Add("ToDate", SAPbouiCOM.BoDataType.dt_DATE);
                }
                catch
                {
                    oForm.DataSources.UserDataSources.Item("FromDate");
                    oForm.DataSources.UserDataSources.Item("ToDate");
                }
                oEdit = oForm.Items.Item("FromDate").Specific;
                oEdit.DataBind.SetBound(true, "", "FromDate");
                oEdit = oForm.Items.Item("ToDate").Specific;
                oEdit.DataBind.SetBound(true, "", "ToDate");
                oForm.Items.Item("FromDate").AffectsFormMode = false;
                oForm.Items.Item("ToDate").AffectsFormMode = false;
                oForm.Visible = true;
                functionReturnValue = 1;
            }
            catch (Exception exc)
            {
                functionReturnValue = 0;
                sErrDesc = exc.Message;
            }
            finally
            {
                creationPackage = null;
                oForm = null;
                oItem = null;
                oTxt = null;
            }
            return functionReturnValue;

        }
        public static long EndDate(ref string sErrDesc)
        {
            long functionReturnValue = 0;
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            int iCount = 0;
            string sFuncName = string.Empty;

            try
            {
                //Check whether the form is exist. If exist then close the form
                for (iCount = 0; iCount <= SBO_Application.Forms.Count - 1; iCount++)
                {
                    oForm = SBO_Application.Forms.Item(iCount);
                    if (oForm.UniqueID == "dDate")
                    {
                        oForm.Close();
                        break; // TODO: might not be correct. Was : Exit For
                    }
                }
                functionReturnValue = 1;
            }
            catch (Exception exc)
            {
                functionReturnValue = 0;
                sErrDesc = exc.Message;
            }
            finally
            {
                oForm = null;
            }
            return functionReturnValue;

        }
        public static long DisplayStatus(SAPbouiCOM.Form oFrmParent, string sMsg, ref string sErrDesc)
        {
            long functionReturnValue = 0;
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            SAPbouiCOM.Item oItem = default(SAPbouiCOM.Item);
            SAPbouiCOM.StaticText oTxt = default(SAPbouiCOM.StaticText);
            SAPbouiCOM.FormCreationParams creationPackage = default(SAPbouiCOM.FormCreationParams);
            int iCount = 0;
            string sFuncName = string.Empty;

            try
            {
                //Check whether the form exists.If exists then close the form
                for (iCount = 0; iCount <= SBO_Application.Forms.Count - 1; iCount++)
                {
                    oForm = SBO_Application.Forms.Item(iCount);
                    if (oForm.UniqueID == "dStatus")
                    {
                        oForm.Close();
                        break; // TODO: might not be correct. Was : Exit For
                    }
                }
                //Add Form
                creationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "dStatus";
                creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_FixedNoTitle;
                creationPackage.FormType = "OBT_dStatus";
                oForm = SBO_Application.Forms.AddEx(creationPackage);
                var _with1 = oForm;
                oForm.AutoManaged = false;
                oForm.Width = 300;
                oForm.Height = 100;
                //if (oFrmParent == null)
                //{
                //    _with1.Left = (Screen.PrimaryScreen.WorkingArea.Width - oForm.Width) / 2;
                //    _with1.Top = (Screen.PrimaryScreen.WorkingArea.Height - oForm.Height) / 3;
                //}
                //else
                //{
                //    _with1.Left = ((oFrmParent.Left * 2) + oFrmParent.Width - oForm.Width) / 2;
                //    _with1.Top = ((oFrmParent.Top * 2) + oFrmParent.Height - oForm.Height) / 2;
                //}

                //Add Label
                oItem = oForm.Items.Add("3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = 40;
                oItem.Left = 40;
                oItem.Width = 250;
                oTxt = oItem.Specific;
                oTxt.Caption = sMsg;
                oForm.Visible = true;
                //Set Form visible at last, if not sMsg will not show

                functionReturnValue = 1;
            }
            catch (Exception exc)
            {
                functionReturnValue = 0;
                sErrDesc = exc.Message;
            }
            finally
            {
                creationPackage = null;
                oForm = null;
                oItem = null;
                oTxt = null;
            }
            return functionReturnValue;

        }
        public static long EndStatus(ref string sErrDesc)
        {
            long functionReturnValue = 0;
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            int iCount = 0;
            string sFuncName = string.Empty;

            try
            {
                //Check whether the form is exist. If exist then close the form
                for (iCount = 0; iCount <= SBO_Application.Forms.Count - 1; iCount++)
                {
                    oForm = SBO_Application.Forms.Item(iCount);
                    if (oForm.UniqueID == "dStatus")
                    {
                        oForm.Close();
                        break; // TODO: might not be correct. Was : Exit For
                    }
                }
                functionReturnValue = 1;
            }
            catch (Exception exc)
            {
                functionReturnValue = 0;
                sErrDesc = exc.Message;
            }
            finally
            {
                oForm = null;
            }
            return functionReturnValue;

        }
        #endregion
    }
}
