
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices;
using System.Globalization;

namespace VEE2
{
    class Program
    {
        static void Main(string[] args)
        {
            #region prerequisite
            SqlConnection cnn;
            string connectionString = null;
            string sql = null;
            string data = null;

            int i = 0;
            int j = 0;

            var shamsiDate = new Shamsi_to_Miladi_convertorDate();

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\ExcellprojectTemplate.xls");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            connectionString = "Data Source=.;Initial Catalog=SabaCandH;User ID=sa;Password=88102351-7";

            cnn = new SqlConnection(connectionString);
            cnn.Open();
            sql = "select distinct ReadDate,TransferDate,obis,Value,ObisFarsiDesc from Meter inner join OBISValueHeader on meter.MeterID = OBISValueHeader.MeterID inner join OBISValueDetail on OBISValueDetail.OBISValueHeaderID = OBISValueHeader.OBISValueHeaderID inner join OBISS on obiss.OBISID = OBISValueDetail.OBISID where Meter.MeterNumber = '1949400024728' and Value != '0'and(Obiss.OBISID = 83 or ObisFarsiDesc like '%آب مصرفي کل%' or Obiss.OBISID = 88 or ObisFarsiDesc like '%ساعت%') order by ReadDate,OBISS.Obis";

            SqlDataAdapter dscmd = new SqlDataAdapter(sql, cnn);
            DataSet ds = new DataSet();
            dscmd.Fill(ds);

            List<object> ReadOutList = new List<object>();
            #endregion

            #region Get ReadOut Data from SQL into a List
            for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
            {
                object[] dataEntryArray = new object[6];

                for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();

                    dataEntryArray[j] = data;
                }
                ReadOutList.Add(dataEntryArray);
            }
            #endregion

            #region Create FinalReadOutList and Change Data into  DataMOdel 
            List<AnalysisDataModel> FinalReadOutList = new List<AnalysisDataModel>();

            DateTime tempTime1 = new DateTime();
            DateTime temptime2 = new DateTime();

            var ctt = new CreateTempTime();
            DateTime lastReadOutDate = new DateTime();

            foreach (object[] element in ReadOutList)
            {
                if (element[0] != null)
                {
                    DateTime convertedDate = shamsiDate.DateConvertor(element);
                    AnalysisDataModel analysisData = new AnalysisDataModel();

                    tempTime1 = ctt.CreateTime(convertedDate);

                    if (tempTime1 == temptime2 && lastReadOutDate == convertedDate)
                    {
                        element[0] = convertedDate;
                        analysisData = new AnalysisDataModel((DateTime)element[0],
                            element[1].ToString(),
                            element[2].ToString(),
                            element[3].ToString(),
                            element[4].ToString(),
                            "");
                        FinalReadOutList.Add(analysisData);
                    }

                    if (temptime2 != tempTime1)
                    {
                        temptime2 = tempTime1;
                        element[0] = convertedDate;
                        analysisData = new AnalysisDataModel((DateTime)element[0],
                            element[1].ToString(),
                            element[2].ToString(),
                            element[3].ToString(),
                            element[4].ToString(),
                            "");
                        FinalReadOutList.Add(analysisData);
                        lastReadOutDate = convertedDate;
                    }
                }
            }

            #endregion

            #region CreateFinalOrderedReadOutList--Add Date(2016.8)
            List<AnalysisDataModel> FinalOrderedReadOutList = new List<AnalysisDataModel>();
            int m = 0;

            DateTime cmdate = new DateTime();

            AnalysisDataModel asb = FinalReadOutList.FirstOrDefault();
            AnalysisDataModel lastItemInThisList = new AnalysisDataModel(new DateTime(), "", "", "", "", "");
            cmdate = asb.ReadOutDate;

            int tm = cmdate.Month;
            int ym = cmdate.Year;
            int dm = cmdate.Day;

            foreach (AnalysisDataModel element in FinalReadOutList)
            {
                if (cmdate != element.ReadOutDate)
                {
                    DateTime d = element.ReadOutDate;

                    tm = d.Month;
                    ym = d.Year;
                    dm = d.Day;
                    cmdate = element.ReadOutDate;
                    m = 0;
                }

                if (element.Obis != null && element.Obis.Contains("802010000") && element.Obis != ("0802010000FF"))
                {

                    if (dm >= 20 && dm <= 31)
                    {
                        if (tm < 1)
                        {
                            tm = 12;
                            ym = ym - 1;
                            element.Date = ym.ToString() + "." + tm.ToString();
                            tm = tm - 1;
                        }
                        else
                        {
                            element.Date = ym.ToString() + "." + tm.ToString();
                            tm = tm - 1;
                        }
                    }

                    else
                    {
                        tm = tm - 1;
                        if (tm < 1)
                        {
                            tm = 12;
                            ym = ym - 1;
                            element.Date = ym.ToString() + "." + (tm).ToString();

                        }
                        else
                        {
                            element.Date = ym.ToString() + "." + (tm).ToString();
                        }
                    }

                    m++;
                    if (lastItemInThisList.ObisFarciDesc != element.ObisFarciDesc)
                    {
                        FinalOrderedReadOutList.Add(element);
                        lastItemInThisList = FinalOrderedReadOutList.LastOrDefault();
                    }
                    else
                    {
                        tm = tm + 1;
                    }
                }

                //if (element[2] != null && element[2].ToString().Contains("802606202"))
                //{
                //    FinalOrderedReadOutList.Add(element);
                //}
            }
            #endregion

            #region  Create FinalSortedList
            List<AnalysisDataModel> ReadOutListOld = new List<AnalysisDataModel>();
            List<AnalysisDataModel> ReadOutListNew = new List<AnalysisDataModel>();
            List<AnalysisDataModel> ReversReadOutListNew = new List<AnalysisDataModel>();

            AnalysisDataModel OldSampleData = FinalOrderedReadOutList.FirstOrDefault();

            int index = 1;

            AnalysisDataModel lastitem1 = new AnalysisDataModel();
            AnalysisDataModel lastitem2 = new AnalysisDataModel();
            List<AnalysisDataModel> SampleLastItem = new List<AnalysisDataModel>();
            List<AnalysisDataModel> FinalConcatList = new List<AnalysisDataModel>();

            foreach (AnalysisDataModel item in FinalOrderedReadOutList)
            {
                if (item.ReadOutDate != OldSampleData.ReadOutDate)
                {
                    if (item.Date == OldSampleData.Date)
                    {
                        if (item.Value == OldSampleData.Value)
                        {
                            foreach (AnalysisDataModel itemm in FinalOrderedReadOutList)
                            {
                                if (item.ReadOutDate == itemm.ReadOutDate)
                                {
                                    if (index == 1)
                                    {
                                        OldSampleData = itemm;
                                        index = 0;
                                        ReadOutListOld = ReadOutListNew.ToList();
                                        ReadOutListNew.Clear();
                                    }
                                    ReadOutListNew.Add(itemm);
                                }
                            }

                            ReadOutListNew.Reverse();
                            lastitem1 = ReadOutListNew.FirstOrDefault();
                            ReadOutListOld.Reverse();

                            foreach (AnalysisDataModel u in ReadOutListOld)
                            {
                                if (lastitem1.Date == u.Date)
                                {
                                    break;
                                }
                                if (lastitem1.Date != u.Date)
                                {
                                    SampleLastItem.Add(u);
                                }
                            }

                            ReadOutListNew = SampleLastItem.Concat(ReadOutListNew).ToList();
                            SampleLastItem.Clear();
                            ReadOutListNew.Reverse();
                            index = 1;
                        }
                    }
                }
            }
            #endregion

            #region Add List to the excell file
            int row = 2;
            int column = 1;



            foreach (AnalysisDataModel item in ReadOutListNew)
            {
                column = 1;
                xlWorkSheet.Cells[row, column++].Value = item.ReadOutDate;
                xlWorkSheet.Cells[row, column++].Value = item.Value;
                xlWorkSheet.Cells[row++, column].Value = item.Date;
            }

            xlWorkSheet.Cells[row++, column].Value = "";

            xlWorkSheet.Cells[row, 1].Value = "ReadDate";
            xlWorkSheet.Cells[row, 2].Value = "TransferDate";
            xlWorkSheet.Cells[row, 3].Value = "Obis";
            xlWorkSheet.Cells[row, 4].Value = "Value";
            xlWorkSheet.Cells[row, 5].Value = "ObisFarciDesc";
            xlWorkSheet.Cells[row++, 6].Value = "Date";
            foreach (AnalysisDataModel item in FinalOrderedReadOutList)
            {
                column = 1;
                xlWorkSheet.Cells[row, column++].Value = item.ReadOutDate;
                xlWorkSheet.Cells[row, column++].Value = item.TransferDate;
                xlWorkSheet.Cells[row, column++].Value = item.Obis;
                xlWorkSheet.Cells[row, column++].Value = item.Value;
                xlWorkSheet.Cells[row, column++].Value = item.ObisFarciDesc;
                xlWorkSheet.Cells[row++, column].Value = item.Date;

            }
            #endregion

            #region Save Excell and close
            xlWorkBook.SaveAs(@"D:\Excellproject.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            var a = new aaa();
            a.releaseObject(xlWorkSheet);
            a.releaseObject(xlWorkBook);
            a.releaseObject(xlApp);
            #endregion
        }
    }
}


