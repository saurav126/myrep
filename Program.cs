/*****************************************************************************
                            Modification Log
 ***************************************************************************
  Initials  Date      Log #               Description
  -------- -------- ---------- ----------------------------------------------
  Basha	   10/28/14            Created
 **************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.ComponentModel;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace CreateServerList
{
    public class CreateServerList
    {
        static void Main(string[] args)
        {
            string sRDGFile = string.Empty;
            string sOutputFile = string.Empty;
            

            CreateServerList obj = new CreateServerList();

            
            if (args != null && args.Length == 1)
            {
                sRDGFile = args[0].ToString();
                //sOutputFile = args[1].ToString();
                obj.Createreport(sRDGFile, sOutputFile);
            }
            else
            {
                Console.WriteLine("USAGE : CreateServerList [RDGFILE] [OUTPUTFILE]");
            }

                  
        }

        public enum ServerType
        {
            Web,
            Proc,
            MP,
            MPMax,
            DXI,
            DXM,
            DXIR,
            ES,
            REDIR,
            EventRouter,
            WFNAll,
            TLMAll,
            v18All,
            UnKnown
        }


        public void Createreport(string sRDGFile, string sOutputFile)
        {
            StringBuilder sbDetails = new StringBuilder();
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            int row=2, col=1;
            DataSet ds = new DataSet();
            ds.ReadXml(sRDGFile);

            System.Data.DataTable Groups = ds.Tables["properties"];
            System.Data.DataTable Servers = ds.Tables["server"];

            var query =
            from groups in Groups.AsEnumerable()
            join servers in Servers.AsEnumerable()
            on groups.Field<int?>("group_Id") equals
                servers.Field<int?>("group_Id")
            select new
            {
                GroupName =
                    groups.Field<string>("Name"),
                ServerName =
                    servers.Field<string>("Name"),
                ServerDisplayName =
                   servers.Field<string>("displayName")
            };

            
            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp.Visible = true;

            xlWorkBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets[1];

            if (xlWorkSheet == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            //xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "Server Name";
            xlWorkSheet.Cells[1, 2] = "Environment";
            xlWorkSheet.Cells[1, 3] = "Release";
            xlWorkSheet.Cells[1, 4] = "Server Type";
            xlWorkSheet.Cells[1, 5] = "Legacy Status";
            xlWorkSheet.Cells[1, 6] = "NextGen Status";
            xlWorkSheet.Cells[1, 7] = "DXI Status";
            xlWorkSheet.Cells[1, 8] = "Mobile Status";
            xlWorkSheet.Cells[1, 9] = "TLMServices Status";
            xlWorkSheet.Cells[1, 10] = "Dispatcher";

            Range firstRow = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Rows[1];
            firstRow.Activate();
            firstRow.Select();
            firstRow.Font.Bold = true;
            firstRow.AutoFilter(1, Type.Missing,
                Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

            //xlWorkBook.Worksheets.Add(xlWorkSheet);

            //xlWorkBook.SaveAs("C:\\Work\\SVN_Source\\Utilities\\RemoteDesktopConnectionManager\\CreateServerList\\CreateServerList\\bin\\Debug\\ServerList.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //xlWorkBook.Save();
            

            foreach (var server in query)
            {
                col = 1;
                xlWorkSheet.Cells[row, col++] = server.ServerName;
                xlWorkSheet.Cells[row, col++] = server.GroupName.Contains("-") ? server.GroupName.Substring(0, server.GroupName.IndexOf("-")) : server.GroupName;
                xlWorkSheet.Cells[row, col++] = server.GroupName.Contains("-") ? server.GroupName.Substring(server.GroupName.IndexOf("-") ,
                    server.GroupName.Length - server.GroupName.IndexOf("-")).Replace("_", ".").Replace(",", " ").Replace("-","").Trim() : string.Empty;

                ServerType eServerType = GetServerType(server.ServerDisplayName, server.ServerName);

                xlWorkSheet.Cells[row, col++] = eServerType == ServerType.UnKnown ? "" : eServerType.ToString();

                switch (eServerType)
                {
                    case ServerType.Web:
                        col = 5;
                        break;
                    case ServerType.Proc:
                        break;
                    case ServerType.MP:
                        col = 7;
                        break;
                    case ServerType.MPMax:
                        col = 7;
                        break;
                    case ServerType.DXI:
                        col = 7;
                        break;
                    case ServerType.DXM:
                        col = 8;
                        break;
                    case ServerType.DXIR:
                        col = 7;
                        break;
                    case ServerType.ES:
                        col = 10;
                        break;
                    case ServerType.EventRouter:
                        col = 10;
                        break;
                    case ServerType.REDIR:
                        col = 5;
                        break;
                    case ServerType.v18All:
                        col = 5;
                        break;
                    default:
                        break;
                }

                string[] strArry = CreateStatusLinks(eServerType, server.ServerName);
                if(strArry !=null && strArry.Length>0)
                {
                    foreach (string str in strArry)
                    {
                        if (str != "")
                        {
                            string[] strLinks = str.ToString().Split(',');
                            //string strLink = "=HYPERLINK(\"" + strLinks[1] + "\",\"\"Infragistics\"\")";    
                            if (eServerType == ServerType.v18All)
                            {
                                if (col == 6)
                                    col = 7;
                                else if (col == 9)
                                    col = 10;
                            }
                            xlWorkSheet.Hyperlinks.Add(xlWorkSheet.Cells[row, col++], strLinks[1], misValue, misValue, strLinks[0]);
                        }
                        
                    }
                }
                row++;
            }

           // xlWorkBook.SaveAs("C:\\Work\\SVN_Source\\Utilities\\RemoteDesktopConnectionManager\\CreateServerList\\CreateServerList\\bin\\Debug\\ServerList.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);


            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(sOutputFile, true))
            //{
            //    file.WriteLine(sbDetails.ToString());
            //}

            ds.Clear();     

        }

        public string[] CreateStatusLinks(ServerType eServerType, string strServerName)
        {
            StringBuilder sbLink = new StringBuilder();
            switch (eServerType)
            {
                case ServerType.Web:
                    sbLink.Append("LegacyWeb,")
                   .Append("http://" + strServerName + "/ezLaborManagerNet/TLMMonitorStatus.aspx").Append(";")
                   .Append("NextGenWeb,")
                   .Append("http://" + strServerName + ":81/TLMWeb/TLMMonitorStatus.aspx").Append(";");
                    break;
                case ServerType.Proc:
                    break;
                case ServerType.MP:
                    sbLink.Append("ezLMServicesInternal,")
                    .Append("http://" + strServerName + ":83/ezLMServicesInternal/TLMMonitorStatus.aspx").Append(";")
                    .Append("Mobile,")
                    .Append("http://" + strServerName + ":85/ezLMWebServicesMobile/TLMMonitorStatus.aspx").Append(";")
                    .Append("TLMServices,")
                    .Append("http://" + strServerName + ":84/TLM/api/v1/TLMMonitorStatus").Append(";")
                    .Append("Dispatcher,")
                    .Append("http://" + strServerName + ":8002").Append(";");
                    break;
                case ServerType.MPMax:
                    sbLink.Append("ezLMServicesInternal,")
                    .Append("http://" + strServerName + ":83/ezLMServicesInternal/TLMMonitorStatus.aspx").Append(";")
                    .Append("Mobile,")
                    .Append("http://" + strServerName + ":85/ezLMWebServicesMobile/TLMMonitorStatus.aspx").Append(";")
                    .Append("TLMServices,")
                .Append("http://" + strServerName + ":84/TLM/api/v1/TLMMonitorStatus").Append(",");
                    break;
                case ServerType.DXI:
                    sbLink.Append("ezLMServicesInternal,").Append("http://" + strServerName + "/ezLMServicesInternal/TLMMonitorStatus.aspx").Append(";");
                    break;
                case ServerType.DXM:
                    sbLink.Append("Mobile,").Append("http://" + strServerName + "/ezLMWebServicesMobile/TLMMonitorStatus.aspx").Append(";");
                    break;
                case ServerType.DXIR:
                    sbLink.Append("DXIR,").Append("http://" + strServerName + "/EndpointRoutingServices/TLMMonitorStatus.aspx").Append(";");            
                    break;
                case ServerType.ES:
                    sbLink.Append("Dispatcher,").Append("http://" + strServerName + ":8002").Append(";");
                    break;
                case ServerType.EventRouter:
                    sbLink.Append("EventRouter,").Append("http://" + strServerName + ":8100").Append(";");
                    break;
                case ServerType.REDIR:
                    sbLink.Append("Redirector,").Append("http://" + strServerName + "/ezLaborManagerNetRedirect/TLMMonitorStatus.aspx").Append(";");
                    break;
                case ServerType.WFNAll:
                case ServerType.TLMAll:
                    sbLink.Append("LegacyWeb,")
                   .Append("http://" + strServerName + "/ezLaborManagerNet/TLMMonitorStatus.aspx").Append(";")
                   .Append("NextGenWeb,")
                   .Append("http://" + strServerName + ":81/TLMWeb/TLMMonitorStatus.aspx").Append(";")
                   .Append("ezLMServicesInternal,")
                    .Append("http://" + strServerName + ":83/ezLMServicesInternal/TLMMonitorStatus.aspx").Append(";")
                    .Append("Mobile,")
                    .Append("http://" + strServerName + ":85/ezLMWebServicesMobile/TLMMonitorStatus.aspx").Append(";")
                    .Append("TLMServices,")
                    .Append("http://" + strServerName + ":84/TLM/api/v1/TLMMonitorStatus").Append(";")
                    .Append("Dispatcher,")
                    .Append("http://" + strServerName + ":8002").Append(";");
                    break;
                case ServerType.v18All:
                    sbLink.Append("LegacyWeb,")
                   .Append("http://" + strServerName + "/ezLaborManagerNet/TLMMonitorStatus.aspx").Append(";")                   
                   .Append("ezLMServicesInternal,")
                    .Append("http://" + strServerName + "/ezLMServicesInternal/TLMMonitorStatus.aspx").Append(";")
                    .Append("Mobile,")
                    .Append("http://" + strServerName + "/ezLMWebServicesMobile/TLMMonitorStatus.aspx").Append(";")                    
                    .Append("Dispatcher,")
                    .Append("http://" + strServerName + ":8002").Append(";");
                    break;
                case ServerType.UnKnown:
                    break;
                default:
                    break;
            }
            
            return sbLink.ToString().Split(';');
        }

        public ServerType GetServerType(string strServerDisplayName, string strServerName)
        {
            strServerDisplayName = strServerDisplayName.Replace(strServerName, "");
            string strServerType = string.Empty;
            
            if (strServerDisplayName.ToLower().Contains("web"))
            {
                return ServerType.Web;
            }
            else if (strServerDisplayName.ToLower().Contains("mpmax"))
            {
                return ServerType.MPMax;
            }
            else if (strServerDisplayName.ToLower().Contains("mp"))
            {
                return ServerType.MP;
            }
            else if (strServerDisplayName.ToLower().Contains("dxir") ||
                strServerDisplayName.ToLower().Contains("dxi router") ||
                strServerDisplayName.ToLower().Contains("endpt routing"))
            {
                return ServerType.DXIR;
            }
            else if (strServerDisplayName.ToLower().Contains("dxi"))
            {
                return ServerType.DXI;
            }
            else if (strServerDisplayName.ToLower().Contains("dxm"))
            {
                return ServerType.DXM;
            }
            else if (strServerDisplayName.ToLower().Contains("proc") || strServerDisplayName.ToLower().Contains("ps"))
            {
                return ServerType.Proc;
            }
            else if (strServerDisplayName.ToLower().Contains("eventrouter") ||
                strServerDisplayName.ToLower().Contains("event router"))
            {
                return ServerType.EventRouter;
            }
            else if (strServerDisplayName.ToLower().Contains("es"))
            {
                return ServerType.ES;
            }
            else if (strServerDisplayName.ToLower().Contains("redir"))
            {
                return ServerType.REDIR;
            }
            else if (strServerDisplayName.ToLower().Contains("wfnall"))
            {
                return ServerType.WFNAll;
            }
            else if (strServerDisplayName.ToLower().Contains("tlmall"))
            {
                return ServerType.TLMAll;
            }
            else if (strServerDisplayName.ToLower().Contains("v18all"))
            {
                return ServerType.v18All;
            }
            return ServerType.UnKnown;
        }

        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
