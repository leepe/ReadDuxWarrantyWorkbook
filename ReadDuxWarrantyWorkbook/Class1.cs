using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace ReadDuxWarrantyWorkbook
{
    public class Class1
    {
        // class constructor for project body details
        //
        public class duxWarranty
        {
            public string warrantyItem { get; set; }
            public string warrantyWhse { get; set; }
            public string warrantyState { get; set; }
            public int warrantyDate { get; set; }
            public int warrantyQty { get; set; }
        }
        //
        //
        private static void Main()
        {
            string f0 = @"C:\Users\peterlee\Downloads\KW -  SE Warranty FC 250315.xls";
            List<duxWarranty> final = readUploadedExcelFiles(f0);
            SqlConnection cnxn = setUpEXPRconn();
            //
            insertIntoTable(cnxn, final);
            //
        }
        //
        // insert into STAGING table dbo.[duxWarrantyForecasts
        //
        public static void insertIntoTable(SqlConnection xn, List<duxWarranty>final)
        {
            // truncate SQL table before uploading data back into it
            xn.Open();
            SqlCommand cmd = new SqlCommand("truncate table dbo.[DUX_Warranty_Forecast_Manual]", xn);
            cmd.CommandType = CommandType.Text;
            cmd.CommandTimeout = 100000;
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            foreach (duxWarranty dw in final)
            {
                // set up INSERT string for commandtext
                string strval = string.Format("insert into dbo.[DUX_Warranty_Forecast_Manual] (warrantyItem,warrantyWhse,warrantyState,warrantyMonth,warrantyQty) " +
                                              "VALUES ('{0}','{1}','{2}',{3},{4})",
                                              dw.warrantyItem,dw.warrantyWhse,dw.warrantyState,dw.warrantyDate,dw.warrantyQty);
                SqlCommand upl = new SqlCommand(strval, xn);
                upl.CommandType = CommandType.Text;
                upl.ExecuteNonQuery();
                upl.Dispose();
            }
            //
            xn.Close();

        }
        //..........................................................................
        // pass fileInfo to spreadsheet reading routine
        //..........................................................................
        //
        public static List<duxWarranty> readUploadedExcelFiles(string fx)
        {

            using (FileStream file = new FileStream(fx, FileMode.Open, FileAccess.Read))
            {
                //.................................................................
                // do not use XSSF version - expects xlsx HSSF version expects xls
                // will generate WEB API INTERNAL SERVER ERROR otherwise
                //.................................................................
                IWorkbook xfn = new HSSFWorkbook(file);
                //.................................................................
                List<duxWarranty> dlist = new List<duxWarranty>();
                //
                foreach (HSSFSheet xf in xfn)
                {
                    //-- read header information for project ---
                    if (xf.SheetName.Equals("warrantyData"))
                    {
                        dlist = loadLineData(xf);
                    }
                }

                return dlist;
            }
        }
        //--------------------------------------------------------------------
        // load warranty workbook line data into warrantyDetails LIST
        //--------------------------------------------------------------------
        public static List<duxWarranty> loadLineData(HSSFSheet xl)
        {
            List<duxWarranty> preLimLines = new List<duxWarranty>();
            // now read project data off body of sheet
            for (int w = 1; w <= xl.LastRowNum; w++)
            {
                for (int z = 3; z < 15; z++)
                {
                    duxWarranty duxw = new duxWarranty();
                    if (xl.GetRow(w).GetCell(z) != null)
                    {
                        duxw.warrantyItem = xl.GetRow(w).GetCell(0).ToString().Trim().ToUpper();        // item codes - trim and capitalise item codes
                        duxw.warrantyWhse = xl.GetRow(w).GetCell(2).ToString().Trim();                  // item state
                        duxw.warrantyState = xl.GetRow(w).GetCell(1).ToString().Trim();                 // item whse
                        duxw.warrantyDate = Convert.ToInt32(xl.GetRow(0).GetCell(z).DateCellValue.ToString("yyyyMMdd"))/100;
                        duxw.warrantyQty = Convert.ToInt32(xl.GetRow(w).GetCell(z).ToString());
                        preLimLines.Add(duxw);
                    }
                }
            }
            return preLimLines;
        }
        //-------------------------------------------------------------------
        // set up connection to SQL-EXPRESS - desktop at bella-vista
        //-------------------------------------------------------------------
        public static SqlConnection setUpEXPRconn()
        {
            // create production connection object
            string strSQLsrvr = "SERVER=WETNT260;USER ID=#DSXdbadmin;PASSWORD=F0res!3R;" +
                    "DATABASE=STAGING;CONNECTION TIMEOUT=30;";
            // create connection object
            //string strSQLsrvr = @"Data Source=PL-X230\SQLEXPRESS;Initial Catalog=STAGING;Integrated Security=True";
            SqlConnection SqlConn = new SqlConnection(strSQLsrvr);
            return SqlConn;
        }
    }
}
