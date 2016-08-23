using ExcelWIP.Models;
using SQL = System.Data;
using System.Collections.ObjectModel;
using ExcelWIP.Controllers;
using System;
using ExcelWIP.VersionTwo.Models;
using ExcelWIP.VersionTwo.Controllers;

namespace ExcelWIP.VersionTwo
{
    public class ExcelManager:DataTableManager
    {
        private ObservableCollection<Repair> AllRepairData;
        private SQL.DataTable WIPDataTable;
        private SQL.DataTable DateInDataTable;
        private SQL.DataTable DateOutDataTable;
        private SQL.DataTable TATDataTable;
        private SQL.DataTable BackOrderDataTable;
        private SQL.DataTable BulkDataTable;
        private SQL.DataTable RetailDataTable;
        private SQL.DataTable RogersDataTable;
        private SQL.DataTable NonRogersDataTable;
        private SQL.DataTable PriorityDataTable;
        private SQL.DataTable TechDataTable;

        //constructor
        public ExcelManager()
        {
            AllRepairData = new ObservableCollection<Repair>();
            //get all data by DataController
            DataController dtController = new DataController();
            AllRepairData = dtController.GetAllRepairData();

            //get WIP data by WIPController
            WIPController wipController = new WIPController();
            WIPDataTable = wipController.GetWIPDataTable(AllRepairData);

            //get dataIn data by DateInController
            DateInController dtInController = new DateInController();
            DateInDataTable = dtInController.GetDateInDataTable(AllRepairData);

            //get dataOut data by DateOutController
            DateOutController dtOutController = new DateOutController();
            DateOutDataTable = dtOutController.GetDateOutDataTable(AllRepairData);

            //get TAT data by TATController
            TATController dtTATController = new TATController();
            TATDataTable = dtTATController.GetTATDataTable(AllRepairData);

            //get BackOrder data by BackOrderController
            BackOrderController dtBackOrderController = new BackOrderController();
            BackOrderDataTable = dtBackOrderController.GetBackOrderDataTable(AllRepairData);

            //get Bulk data by BulkController
            BulkController dtBulkController = new BulkController();
            BulkDataTable = dtBulkController.GetBulkDataTable(AllRepairData);

            //get Retail data by RetailController
            RetailController dtRetailController = new RetailController();
            RetailDataTable = dtRetailController.GetRetailDataTable(AllRepairData);

            //get Rogers data by RogersController
            RogersController dtRogersController = new RogersController();
            RogersDataTable = dtRogersController.GetRogersDataTable(AllRepairData);

            //get Non Rogers data by NonRogersController
            NonRogersController dtNonRogersController = new NonRogersController();
            NonRogersDataTable = dtNonRogersController.GetNonRogersDataTable(AllRepairData);

            //get Priority data by PriorityController
            PriorityController dtPriorityController = new PriorityController();
            PriorityDataTable = dtPriorityController.GetPriorityDataTable(AllRepairData);

            //get Technician data by TechnicianController
            TechnicianController dtTechnicianController = new TechnicianController();
            TechDataTable = dtTechnicianController.GetTechDataTable(AllRepairData);
        }

        public void GetExcelSheet(ref string fileName)
        {
            //set work sheet name
            #region Local variables for sheet name
            string sheetWIP = "WIP";
            string sheetDateIn = "In";
            string sheetDateOut = "Out";
            string sheetTAT = "TAT";
            string sheetBackOrder = "BackOrder";
            string sheetBulk = "Bulk";
            string sheetRetail = "Retail";
            string sheetRogers = "Rogers";
            string sheetNonRogers = "NonRogers";
            string sheetProrityList = "ProrityList";
            string sheetTechOutput = "TechOutput";
            #endregion

            //create excel woork
            #region Create Excel
            fileName = "";
            CreateExcel(ref fileName);
            Console.WriteLine("File Name: " + fileName);
            #endregion

            //add work sheet by name
            #region Create WoorkSheet
            //call CreatSheet
            CreatSheet(TechDataTable, sheetTechOutput);
            CreatSheet(PriorityDataTable, sheetProrityList);
            CreatSheet(NonRogersDataTable, sheetNonRogers);
            CreatSheet(RogersDataTable, sheetRogers);
            CreatSheet(RetailDataTable, sheetRetail);
            CreatSheet(BulkDataTable, sheetBulk);
            CreatSheet(BackOrderDataTable, sheetBackOrder);
            CreatSheet(TATDataTable, sheetTAT);
            CreatSheet(DateOutDataTable, sheetDateOut);
            CreatSheet(DateInDataTable, sheetDateIn);
            CreatSheet(WIPDataTable, sheetWIP);

            #endregion

            //save excel sheet
            #region Save Excel
            SaveExcel();
            #endregion
            
        }

        //methods overried
        public override void CreateExcel(ref string fileName)
        {
            base.CreateExcel(ref fileName);
        }

        public override void CreatSheet(SQL.DataTable dataTable, string sheetName)
        {
            base.CreatSheet(dataTable, sheetName);
        }

        public override void SaveExcel()
        {
            base.SaveExcel();
        }
    }
}
