using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWIP
{
    public class Connection
    {
        //connection string
        private static string ConnectionString = "Data Source=reportdb;Initial Catalog=EasyWIPDB;Integrated Security=True";
        //instance variable - query
        private static string QueryAll = @"SELECT R.RefNumber,DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as Aging,DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as TATByDateIn,DATEDIFF(day, R.DateDockIn, convert(date, GETDATE())) as TATByDockIn,D.Name as DealerName,R.DealerID,R.GSPNTicketNo,R.DateIn,R.DateComplete,R.DateDockIn,R.DateDockOut,R.DateEstimation,
                            R.DateApproved,R.DateReject,R.DateBackorder,R.Warranty,R.Program,R.FuturetelLocation,R.ESN,R.MCN,R.SVP,R.ModelNumber,
                            R.DelayReason, CONCAT(U.FirstName,' ',U.LastName) as Technician,R.DealerRefNumber,R.Manufacturer,R.Status, R.DateFinish
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            LEFT JOIN EasyWIPDB.dbo.tblDealer D
                            ON R.DealerID = D.DealerID
							LEFT JOIN EasyWIPDB.dbo.tblUser U
                            ON R.LastTechnician = U.UserName
                            WHERE (R.Status != 'C' and R.Status != 'S' and R.Status != 'M' and R.Status != 'X')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.FuturetelLocation != 'STOCK')
                            AND R.Manufacturer = 'SAMSUNG'
							union
							SELECT R.RefNumber,DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as Aging,DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as TATByDateIn,DATEDIFF(day, R.DateDockIn, convert(date, GETDATE())) as TATByDockIn,D.Name as DealerName,R.DealerID,R.GSPNTicketNo,R.DateIn,R.DateComplete,R.DateDockIn,R.DateDockOut,R.DateEstimation,
                            R.DateApproved,R.DateReject,R.DateBackorder,R.Warranty,R.Program,R.FuturetelLocation,R.ESN,R.MCN,R.SVP,R.ModelNumber,
                            R.DelayReason, CONCAT(U.FirstName,' ',U.LastName) as Technician,R.DealerRefNumber,R.Manufacturer,R.Status, R.DateFinish
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            LEFT JOIN EasyWIPDB.dbo.tblDealer D
                            ON R.DealerID = D.DealerID
							LEFT JOIN EasyWIPDB.dbo.tblUser U
                            ON R.LastTechnician = U.UserName
                            WHERE (R.Status != 'S' and R.Status != 'M' and R.Status != 'X')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.FuturetelLocation != 'STOCK')
                            AND R.Manufacturer = 'SAMSUNG'
							AND (R.DateIn > (Format(GetDate(), N'yyyy-MM-dd')+' 08:00:00')
							or R.DateComplete > (Format(GetDate(), N'yyyy-MM-dd')+' 07:00:00'))
                            ORDER BY Aging ASC, R.RefNumber ASC";

        private static string QueryWIP = @"SELECT R.RefNumber,DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as Aging,D.Name as DealerName,R.DealerID,R.GSPNTicketNo,R.DateIn,R.DateComplete,R.DateDockIn,R.DateDockOut,R.DateEstimation,
                            R.DateApproved,R.DateReject,R.DateBackorder,R.Warranty,R.Program,R.FuturetelLocation,R.ESN,R.MCN,R.SVP,R.ModelNumber,
                            R.DelayReason,R.LastTechnician,R.DealerRefNumber
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            LEFT JOIN EasyWIPDB.dbo.tblDealer D
                            ON R.DealerID = D.DealerID
                            WHERE (R.Status != 'C' and R.Status != 'S' and R.Status != 'M' and R.Status != 'X')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.FuturetelLocation != 'STOCK')
                            AND R.Manufacturer = 'SAMSUNG'
                            ORDER BY Aging Desc, R.RefNumber ASC";
        private static string QueryIn = @"SELECT R.RefNumber,R.ModelNumber,D.Name as DealerName,R.DealerID,R.DateIn,R.DateDockIn,R.Warranty
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            LEFT JOIN EasyWIPDB.dbo.tblDealer D
                            ON R.DealerID = D.DealerID
                            WHERE (R.Status != 'C' and R.Status != 'S' and R.Status != 'M' and R.Status != 'X')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.FuturetelLocation != 'STOCK')
                            AND R.Manufacturer = 'SAMSUNG'
                            AND (R.DateIn > (Format(GetDate(), N'yyyy-MM-dd')+' 07:00:00') and R.DateIn < GETDATE())
                            ORDER BY R.DateIn Asc, R.RefNumber ASC";
        private static string QueryOut = @"SELECT R.RefNumber,R.ModelNumber,D.Name as DealerName,R.DealerID,R.DateIn,R.DateDockIn,R.DateComplete
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            LEFT JOIN EasyWIPDB.dbo.tblDealer D
                            ON R.DealerID = D.DealerID
                            WHERE (R.Status != 'S' and R.Status != 'M' and R.Status != 'X')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.FuturetelLocation != 'STOCK')
                            AND R.Manufacturer = 'SAMSUNG'
                            AND (R.DateComplete > (Format(GetDate(), N'yyyy-MM-dd')+' 07:00:00') and R.DateComplete < GETDATE())
                            ORDER BY R.DateComplete Asc, R.RefNumber ASC";
        private static string QueryTAT = @"SELECT DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as TATByDateIn,DATEDIFF(day, R.DateDockIn, convert(date, GETDATE())) as TATByDockIn,R.RefNumber,D.Name as DealerName,R.DateIn,R.DateDockIn,R.DateComplete
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            LEFT JOIN EasyWIPDB.dbo.tblDealer D
                            ON R.DealerID = D.DealerID
                            WHERE (R.Status != 'S' and R.Status != 'M' and R.Status != 'X')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.FuturetelLocation != 'STOCK')
                            AND R.Manufacturer = 'SAMSUNG'
                            AND (R.DateComplete > (Format(GetDate(), N'yyyy-MM-dd')+' 07:00:00') and R.DateComplete < GETDATE())
                            AND R.DelayReason =''
                            ORDER BY R.DateComplete Asc, R.RefNumber ASC";
        private static string QueryBackOrder = @"SELECT R.RefNumber,DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as Aging,R.ModelNumber,R.DateIn,R.DateDockIn,
                            R.DelayReason,R.FuturetelLocation,R.LastTechnician,R.Status
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            WHERE (R.Status != 'S' and R.Status != 'M' and R.Status != 'X')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.FuturetelLocation = 'B')
                            AND R.Manufacturer = 'SAMSUNG'
                            ORDER BY Aging Desc, R.RefNumber ASC";
        private static string QueryBulk = @"SELECT R.RefNumber,DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as Aging,R.FuturetelLocation,R.Program
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            WHERE (R.Status != 'S' and R.Status != 'M' and R.Status != 'X' and R.Status != 'C')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.Program = 'BULK')
                            AND R.Manufacturer = 'SAMSUNG'
                            ORDER BY Aging ASC, R.RefNumber ASC";
        private static string QueryRetail = @"SELECT R.RefNumber,DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as Aging,R.FuturetelLocation,R.Program
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            WHERE (R.Status != 'S' and R.Status != 'M' and R.Status != 'X' and R.Status != 'C')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.DealerID = 517)
                            AND R.Manufacturer = 'SAMSUNG'
                            ORDER BY Aging ASC, R.RefNumber ASC";
        private static string QueryRogers = @"SELECT R.RefNumber,DATEDIFF(day, R.DateIn, convert(date, GETDATE())) as Aging,R.FuturetelLocation,R.Program
                            FROM EasyWIPDB.dbo.tblReportWIP R
                            WHERE (R.Status != 'S' and R.Status != 'M' and R.Status != 'X' and R.Status != 'C')
                            AND (R.DealerID!= 7430 and R.DealerID!= 7432 and R.DealerID!= 7481 and R.DealerID!= 7482 and R.DealerID!= 7498 and R.DealerID!= 7550 and R.DealerID!= 7551 and R.DealerID!= 7552 and R.DealerID!= 7595)
                            AND (R.SVP != 'TCHURN' and R.SVP != 'TCC' and R.SVP != 'KCC' and R.SVP != 'TXREPAIR' and R.SVP != 'KXREPAIR' and R.SVP != 'KCHURN' and R.SVP != 'TEXPRESS' and R.SVP != 'KEXPRESS')
                            AND (R.DealerID = '517') and (R.SVP = 'ROGERS' or R.SVP = 'FIDO')
                            AND R.Manufacturer = 'SAMSUNG'
                            ORDER BY Aging ASC, R.RefNumber ASC";

        //file name
        private static string FileName;

        //constructor
        public Connection()
        {
            //current location
            string currentLocation = System.AppDomain.CurrentDomain.BaseDirectory;
            string currentDate = DateTime.Now.ToString("yyyyMMdd-HHmm");
            //file name
            FileName = @"\\server\shop\Manufacturer Files\Samsung\WIP Report\Samsung WIP " + currentDate + ".xlsx";
        }

        //method
        public string GetFileName()
        {
            return FileName;
        }
        public string GetConnectionString()
        {
            return ConnectionString;
        }
        public string GetQueryAll()
        {
            return QueryAll;
        }
        public string GetQueryWIP()
        {
            return QueryWIP;
        }
        public string GetQueryIn()
        {
            return QueryIn;
        }
        public string GetQueryOut()
        {
            return QueryOut;
        }
        public string GetQueryTAT()
        {
            return QueryTAT;
        }
        public string GetQueryBackOrder()
        {
            return QueryBackOrder;
        }
        public string GetQueryBulk()
        {
            return QueryBulk;
        }
        public string GetQueryRetail()
        {
            return QueryRetail;
        }
        public string GetQueryRogers()
        {
            return QueryRogers;
        }
    }
}
