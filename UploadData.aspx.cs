using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DataAccess;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Configuration;
using System.Data.OleDb;

public partial class UploadData : System.Web.UI.Page
{
    DBAccess objDBAccss = new DBAccess(ConfigurationManager.ConnectionStrings["saurabhDbcon"].ConnectionString);
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            InitializeComponent();
        }
    }

    private void InitializeComponent()
    {

    }

    List<String> lstFilesFound = new List<String>();
    public IEnumerable<string> GetFiles(string path)
    {
        Queue<string> queue = new Queue<string>();
        queue.Enqueue(path);
        String[] allowedExt = { ".xls", ".xlsx" };
        while (queue.Count > 0)
        {
            path = queue.Dequeue();
            try
            {
                foreach (string subDir in Directory.GetDirectories(path))
                {
                    queue.Enqueue(subDir);
                }
            }
            catch (Exception ex)
            {
                //Console.Error.WriteLine(ex);
            }
            string[] files = null;
            try
            {
                files = Directory.GetFiles(path).Where(f => allowedExt.Contains(Path.GetExtension(f))).ToArray();

            }
            catch (Exception ex)
            {
                //Console.Error.WriteLine(ex);
            }
            if (files != null)
            {
                for (int i = 0; i < files.Length; i++)
                {
                    yield return files[i];
                }
            }
        }
    }

    public List<String> GetSheetsOfFile(String filePath)
    {

        List<String> SheetName = new List<string>();

        //Creating connection string.
        string conString = string.Empty;
        string extension = Path.GetExtension(filePath);
        switch (extension)
        {
            case ".xls": //Excel 97-03
                conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                break;
            case ".xlsx": //Excel 07 or higher
                conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
                break;

        }
        conString = string.Format(conString, filePath);
        //Creating connection string ends.


        using (OleDbConnection excel_con = new OleDbConnection(conString))
        {
            excel_con.Open();
            System.Data.DataTable dtSheetname = new System.Data.DataTable();
            System.Data.DataTable dtExcelData = new System.Data.DataTable();
            dtSheetname = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            OleDbCommand cmd;
            foreach (System.Data.DataRow row in dtSheetname.Rows)
            {
                if (!row["TABLE_NAME"].ToString().Contains("FilterDatabase"))
                {
                    string query1 = "SELECT count(*) FROM [" + row["TABLE_NAME"].ToString() + "]";
                    cmd = new OleDbCommand(query1, excel_con);
                    if (Convert.ToInt32(cmd.ExecuteScalar()) > 0)
                    {
                        SheetName.Add(row["TABLE_NAME"].ToString());
                    }
                }
            }
            excel_con.Close();
        }
        return SheetName;
    }

    public DataTable GetColumnsOFSheet(String filePath, String Sheetname)
    {
        System.Data.DataTable ExcelColumDt = new System.Data.DataTable();
        System.Data.DataTable dtExcelData = new System.Data.DataTable();
        //Creating connection string.
        string conString = string.Empty;
        string extension = Path.GetExtension(filePath);
        switch (extension)
        {
            case ".xls": //Excel 97-03
                conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                break;
            case ".xlsx": //Excel 07 or higher
                conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
                break;

        }
        conString = string.Format(conString, filePath);
        //Creating connection string ends.


        using (OleDbConnection excel_con = new OleDbConnection(conString))
        {
            excel_con.Open();
            string query1 = "SELECT * FROM [" + Sheetname + "] where 1 = 2";
            using (OleDbDataAdapter oda = new OleDbDataAdapter(query1, excel_con))
            {
                oda.Fill(dtExcelData);
            }
            excel_con.Close();
        }

        ExcelColumDt.Columns.Add("ExcelColumn");
        DataRow dr;
        foreach (DataColumn clmn in dtExcelData.Columns)
        {
            dr = ExcelColumDt.NewRow();
            dr["ExcelColumn"] = clmn.ColumnName;
            ExcelColumDt.Rows.Add(dr);
        }

        return ExcelColumDt;
    }

    public void bindFileColumGrid()
    {

        //Get file list for uploading.
        DataSet Ds = new DataSet();
        Ds = objDBAccss.getDataSet("GetFileList", null,CommandType.StoredProcedure);
        grdFileQuea.DataSource = String.Empty;
        grdColumnMapping.DataSource = String.Empty;
        hdnFilePath.Value = "";
        hdnProcessFieldId.Value = "";
        hdnSheetName.Value = "";
        lblWarning.Text = "";
        if (Ds != null && Ds.Tables.Count > 0)
        {
            grdFileQuea.DataSource = Ds.Tables[0];
            //grdFileQuea.DataBind();

            if (Ds.Tables.Contains("Table1") && Ds.Tables.Contains("Table2"))
            {

                if (Ds.Tables[1].Rows.Count > 0)
                {
                    //Bind column mapping grid.
                    DataTable dt = Ds.Tables[2].Copy();

                    DataColumn DatClmn = new DataColumn("FilePath", typeof(String));
                    DatClmn.DefaultValue = Convert.ToString(Ds.Tables[1].Rows[0]["FilePath"]);
                    dt.Columns.Add(DatClmn);

                    DatClmn = new DataColumn("SheetName", typeof(String));
                    DatClmn.DefaultValue = Convert.ToString(Ds.Tables[1].Rows[0]["SheetName"]);
                    dt.Columns.Add(DatClmn);



                    //DataRow[] drArry = dt.Select("isBase = 1");
                    DataView dataView = dt.DefaultView;
                    dataView.RowFilter = "isBase = 1 and masterColumnName <> 'BatchId'";
                    grdColumnMapping.DataSource = dataView;
                    //grdColumnMapping.DataBind();

                    //Catch mapping columns values
                    Cache.Insert("columnMapping", Ds.Tables[2], null, DateTime.Now.AddMinutes(15), TimeSpan.Zero);
                    if (Convert.ToString(Ds.Tables[1].Rows[0]["DataStatus"]) == "Processed")
                    {
                        lblWarning.Text = "This data is Already been uploaded.";
                    }
                    hdnFilePath.Value = Convert.ToString(Ds.Tables[1].Rows[0]["FilePath"]);
                    hdnProcessFieldId.Value = Convert.ToString(Ds.Tables[1].Rows[0]["FileId"]);
                    hdnSheetName.Value = Convert.ToString(Ds.Tables[1].Rows[0]["SheetName"]);
                }//if sheet is under process.
                else
                {
                    lblWarning.Text = "All Record processed in this folder.";
                }
            }
        }
        grdFileQuea.DataBind();
        grdColumnMapping.DataBind();
        Cache.Remove("ExcelColumnsDt");

    }

    protected void btnGetFiles_Click(object sender, EventArgs e)
    {
        if (!String.IsNullOrWhiteSpace(txtfolderPath.Text))
        {
            lstFilesFound = new List<String>();
            lstFilesFound = GetFiles(txtfolderPath.Text).ToList();

            String StatusChk = "";
            objDBAccss.ExecNonQuery("truncate table Temp_fileList", null);
            StatusChk = Convert.ToString(objDBAccss.SelectScalar("select count(1) from Temp_fileList"));

            //Bulk insert for the files.
            if (!String.IsNullOrWhiteSpace(StatusChk) && StatusChk.Trim() == "0")
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("FileId", typeof(Int64));
                dt.Columns.Add("FilePath", typeof(String));
                dt.Columns.Add("SheetName", typeof(String));
                //dt.Columns.Add("IsRead", typeof(Boolean));
                //dt.Columns.Add("IsSkipped", typeof(Boolean));

                DataRow dr;
                foreach (String item in lstFilesFound)
                {
                    foreach (String WkSheet in GetSheetsOfFile(item))
                    {
                        dr = dt.NewRow();
                        //dr["FileId"] = 
                        dr["FilePath"] = item;
                        dr["SheetName"] = WkSheet;
                        //dr["IsSkipped"]
                        dt.Rows.Add(dr);
                    }
                }

                SqlBulkCopy bluckCop = new SqlBulkCopy(objDBAccss.conObj);
                bluckCop.DestinationTableName = "Temp_fileList";

                objDBAccss.openCon(objDBAccss.conObj);
                try
                {
                    bluckCop.WriteToServer(dt);
                }
                catch
                {
                }
                objDBAccss.CloseCon(objDBAccss.conObj);
            }

            bindFileColumGrid();
        }
    }
    protected void grdFileQuea_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            String FileStatus = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "FileStatus"));

            if (!String.IsNullOrWhiteSpace(FileStatus))
            {
                e.Row.Attributes.Add("class", FileStatus);
                //e.Row.BackColor = System.Drawing.Color.LightPink;
            }
        }
    }

    protected void grdColumnMapping_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            String ProcessFileId = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "ProcessFileId"));
            String FilePath = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "FilePath"));
            String SheetName = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "SheetName"));
            String masterColumnName = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "masterColumnName"));
            DataTable dtExcelColumn = new DataTable();
            if (Cache["ExcelColumnsDt"] == null)
            {
                dtExcelColumn = GetColumnsOFSheet(FilePath, SheetName);
                Cache.Insert("ExcelColumnsDt", dtExcelColumn, null, DateTime.Now.AddMinutes(15), TimeSpan.Zero);
            }
            dtExcelColumn = Cache["ExcelColumnsDt"] as DataTable;

            DropDownList drdwExcelColumn = e.Row.FindControl("drdwExcelColumn") as DropDownList;
            drdwExcelColumn.DataSource = dtExcelColumn;
            drdwExcelColumn.DataValueField = "ExcelColumn";
            drdwExcelColumn.DataTextField = "ExcelColumn";
            drdwExcelColumn.DataBind();
            drdwExcelColumn.Items.Insert(0, new ListItem("--Select--", "-1"));

            //Set selected value logic.
            DataTable DtMappingDetails = new DataTable();
            if (Cache["columnMapping"] == null)
            {

                String SqlQuery = "select distinct @InprogressId as ProcessFileId, masterColumnName,MatchedExcelColoum,isBase from tbl_columMapping";
                SqlParameter[] param = new SqlParameter[]{
                new SqlParameter("@InprogressId",ProcessFileId)
                };

                DataSet ds = new DataSet();
                ds = objDBAccss.getDataSet(SqlQuery, param);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DtMappingDetails = ds.Tables[0];
                    Cache.Insert("columnMapping", ds.Tables[0], null, DateTime.Now.AddMinutes(15), TimeSpan.Zero);
                }
            }

            DtMappingDetails = Cache["columnMapping"] as DataTable;
            DataRow[] drArr = DtMappingDetails.Select("masterColumnName = '" + masterColumnName + "'");
            DataRow[] DrArrExcelMatch;
            foreach (DataRow dr in drArr)
            {
                DrArrExcelMatch = dtExcelColumn.Select("ExcelColumn = '" + Convert.ToString(dr["MatchedExcelColoum"]) + "'");
                if (DrArrExcelMatch.Count() > 0)
                {
                    drdwExcelColumn.SelectedValue = Convert.ToString(DrArrExcelMatch[0]["ExcelColumn"]);
                    break;
                }
            }

        }
    }
    protected void btnRejectExcel_Click(object sender, EventArgs e)
    {
        String processFieldId = hdnProcessFieldId.Value;
        String FilePath = hdnFilePath.Value;
        String SheetName = hdnSheetName.Value;
        if (!String.IsNullOrWhiteSpace(processFieldId))
        {
            SqlParameter[] param = new SqlParameter[]{
                new SqlParameter("@InprogressId",processFieldId)
            };
            objDBAccss.ExecNonQuery("update Temp_fileList set IsRead = 0, IsSkipped = 1 where FileId = @InprogressId", param);
            bindFileColumGrid();
        }
    }
    protected void btnUploadData_Click(object sender, EventArgs e)
    {
        if (ValidateUpload())
        {
            String processFieldId = hdnProcessFieldId.Value;
            String FilePath = hdnFilePath.Value;
            String SheetName = hdnSheetName.Value;
            String BatchId = "", ExcelColmStr = "", chkBulkCopy = "";
            Dictionary<String, String> DestinationSouce = new Dictionary<string, string>();
            DataTable ColumMapping = new DataTable();
            ColumMapping.Columns.Add("masterColumnName", typeof(String));
            ColumMapping.Columns.Add("MatchedExcelColoum", typeof(String));
            ColumMapping.Columns.Add("BatchId", typeof(Int64));

            SqlParameter[] param = new SqlParameter[]{
                new SqlParameter("@FilePath",FilePath),
                new SqlParameter("@SheetName",SheetName),
                new SqlParameter("@overrideData",((chkOverride.SelectedValue == "True")?"1":"0"))
            };
            BatchId = Convert.ToString(objDBAccss.SelectScalar(@"insert into tbl_processedFile(FilePath,SheetName,overrideData) values(@FilePath,@SheetName,@overrideData); Select @@IDENTITY", param));
            if (!String.IsNullOrWhiteSpace(BatchId))
            {
                DataRow DrMatchedColum;
                //Bulk insert data
                foreach (GridViewRow grvRow in grdColumnMapping.Rows)
                {
                    DrMatchedColum = ColumMapping.NewRow();
                    if (grvRow.RowType == DataControlRowType.DataRow)
                    {
                        DropDownList drpDwnExcelColum = grvRow.FindControl("drdwExcelColumn") as DropDownList;
                        TextBox txtDefaultData = grvRow.FindControl("txtDefaultData") as TextBox;
                        Label lblMasterColum = grvRow.FindControl("lblMasterColum") as Label;
                        if (drpDwnExcelColum.SelectedValue != "-1")
                        {
                            ExcelColmStr = ExcelColmStr + "[" + drpDwnExcelColum.SelectedValue + "], ";
                            DestinationSouce.Add(lblMasterColum.Text, drpDwnExcelColum.SelectedValue);
                            DrMatchedColum["masterColumnName"] = lblMasterColum.Text;
                            DrMatchedColum["MatchedExcelColoum"] = drpDwnExcelColum.SelectedValue;
                            DrMatchedColum["BatchId"] = Convert.ToInt64(BatchId);
                            ColumMapping.Rows.Add(DrMatchedColum);
                        }
                        else if (!String.IsNullOrWhiteSpace(txtDefaultData.Text))
                        {
                            ExcelColmStr = ExcelColmStr + "'" + txtDefaultData.Text.Trim() + "' as [DFT" + lblMasterColum.Text + "], ";
                            DestinationSouce.Add(lblMasterColum.Text, "DFT" + lblMasterColum.Text);
                        }

                    }
                }
                ExcelColmStr = ExcelColmStr.Trim().TrimEnd(',');
                if (!String.IsNullOrWhiteSpace(ExcelColmStr))
                {
                    //Get Data from excel to datatable
                    DataTable ExcelData = new DataTable();
                    string conString = string.Empty;
                    string extension = Path.GetExtension(FilePath);
                    switch (extension)
                    {
                        case ".xls": //Excel 97-03
                            conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                            break;
                        case ".xlsx": //Excel 07 or higher
                            conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
                            break;

                    }
                    conString = string.Format(conString, FilePath);
                    //Creating connection string ends.


                    using (OleDbConnection excel_con = new OleDbConnection(conString))
                    {
                        excel_con.Open();
                        string query1 = "SELECT '" + BatchId + "' as [BatchId], " + ExcelColmStr + "FROM [" + SheetName + "]";
                        using (OleDbDataAdapter oda = new OleDbDataAdapter(query1, excel_con))
                        {
                            oda.Fill(ExcelData);
                        }
                        excel_con.Close();
                    }
                    if (ExcelData != null && ExcelData.Rows.Count > 0)
                    {
                        //Sql bulk upload of data
                        SqlBulkCopy bluckCop = new SqlBulkCopy(objDBAccss.conObj);
                        bluckCop.DestinationTableName = "tbl_MasterData";
                        bluckCop.ColumnMappings.Add("BatchId", "BatchId");
                        //Column mapping
                        foreach (var item in DestinationSouce)
                        {
                            bluckCop.ColumnMappings.Add(item.Value, item.Key);
                        }
                        objDBAccss.openCon(objDBAccss.conObj);
                        try
                        {
                            bluckCop.WriteToServer(ExcelData);
                            chkBulkCopy = "Success";
                        }
                        catch
                        {
                            chkBulkCopy = "Fail";
                        }
                        objDBAccss.CloseCon(objDBAccss.conObj);

                        //Bulk upload of Mapping.
                        bluckCop = new SqlBulkCopy(objDBAccss.conObj);
                        bluckCop.DestinationTableName = "Temp_columMapping";
                        objDBAccss.openCon(objDBAccss.conObj);
                        try
                        {
                            bluckCop.WriteToServer(ColumMapping);
                        }
                        catch (Exception)
                        {
                            
                        }
                        objDBAccss.CloseCon(objDBAccss.conObj);

                        param = new SqlParameter[]{
                                new SqlParameter("@TempFileLstFieldId",processFieldId),
                                new SqlParameter("@BatchId",BatchId),
                                new SqlParameter("@isOverride",((chkOverride.SelectedValue == "True")?"1":"0")),
                                new SqlParameter("@IsBulkSuccess",chkBulkCopy == "Success"?"1":"0")
                            };
                        objDBAccss.ExecNonQuery("AfterBulkDataUpload", param, CommandType.StoredProcedure);
                    }

                    
                    bindFileColumGrid();
                }
            }
        }
    }

    public Boolean ValidateUpload()
    {
        String Errmsg = "";
        if (grdColumnMapping.Rows.Count <= 0)
        {
            Errmsg = Errmsg + "No Column Mapping Found or No excel is In Process.<br/>";
        }
        if (grdColumnMapping.Rows.Count > 0)
        {
            foreach (GridViewRow grvRow in grdColumnMapping.Rows)
            {
                if (grvRow.RowType == DataControlRowType.DataRow)
                {
                    DropDownList drpDwnExcelColum = grvRow.FindControl("drdwExcelColumn") as DropDownList;
                    TextBox txtDefaultData = grvRow.FindControl("txtDefaultData") as TextBox;
                    Label lblErrMsg = grvRow.FindControl("lblErrMsg") as Label;
                    lblErrMsg.Text = "";
                    if (drpDwnExcelColum.SelectedValue == "-1" && String.IsNullOrWhiteSpace(txtDefaultData.Text))
                    {
                        lblErrMsg.Text = "Eighter select Excel colum or enter default data";
                        if (String.IsNullOrWhiteSpace(Errmsg))
                            Errmsg = Errmsg + "Error Found in Mapping Grid";
                        //Errmsg = Errmsg + "No Mapping found for row no " + (grvRow.RowIndex + 1).ToString() + " Eighter select Excel colum or enter default data .<br/>";
                    }

                }
            }
        }
        lblValidationError.Text = HttpUtility.HtmlDecode(Errmsg);
        return String.IsNullOrWhiteSpace(Errmsg);
    }
}