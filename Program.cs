using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {

            ExportExcelWithoutHeader();


            ExportExcelWithHeader();
        }

        private static void ExportExcelWithoutHeader()
        {
            var list = new List<string>();
            list.Add("H|20190503|500882019000004|");
            list.Add("D|1|20190502|PDB01|999999999999999|00000001|70838155100000011|Sales|500.00|FuelSales|A01922381|transaction-01|Y|");
            list.Add("D|1|20190502|PDB01|999999999999999|00000001|70838155100000011|Sales|500.00|FuelSales|A01922381|transaction-02|Y|");
            list.Add("D|1|20190502|PDB01|999999999999999|00000001|70838155100000011|Sales|500.00|FuelSales|A01922381|transaction-03|Y|");
            list.Add("D|1|20190502|PDB01|999999999999999|00000001|70838155100000011|Sales|500.00|FuelSales|A01922381|transaction-04|Y|");
            list.Add("D|1|20190502|PDB01|999999999999999|00000001|70838155100000011|Sales|500.00|FuelSales|A01922381|transaction-05|N|");
            list.Add("D|1|20190502|PDB01|999999999999999|00000001|70838155100000011|Sales|500.00|FuelSales|A01922381|transaction-06|Y|");
            list.Add("D|1|20190502|PDB01|999999999999999|00000001|70838155100000011|Sales|500.00|FuelSales|A01922381|transaction-07|Y|");
            list.Add("D|1|20190502|PDB01|999999999999999|00000001|70838155100000011|Sales|500.00|FuelSales|A01922381|transaction-08|Y|");
            list.Add("T|8|2450.00|4|4|0|4|1550.00|3|1|");

            var toExport = new List<string>();
            foreach (var item in list)
            {
                toExport.Add(item);
            }
            var ExcelPkg = CreateExcel(list);
            FileInfo file = new FileInfo(Directory.GetCurrentDirectory() + @"\" + DateTime.Now.ToString("yyyyMMddHHmmssffftt") + ".xlsx");
            ExcelPkg.SaveAs(file);
        }

        private static void ExportExcelWithHeader()
        {
            var list = new List<VehiclesListModel>();
            list.Add(new VehiclesListModel { AppcId = "111", CardExpiry = "201912", CardNo = "Card No", VehRegtNo = "VehRegtNo" });
            list.Add(new VehiclesListModel { AppcId = "111", CardExpiry = "201912", CardNo = "Card No", VehRegtNo = "VehRegtNo" });
            list.Add(new VehiclesListModel { AppcId = "111", CardExpiry = "201912", CardNo = "Card No", VehRegtNo = "VehRegtNo" });
            list.Add(new VehiclesListModel { AppcId = "111", CardExpiry = "201912", CardNo = "Card No", VehRegtNo = "VehRegtNo" });
            list.Add(new VehiclesListModel { AppcId = "111", CardExpiry = "201912", CardNo = "Card No", VehRegtNo = "VehRegtNo" });
            list.Add(new VehiclesListModel { AppcId = "111", CardExpiry = "201912", CardNo = "Card No", VehRegtNo = "VehRegtNo" });

            var title = "Vehicle List";
            var toExport = new List<string[]>();
            var Header = list.First().CsvHeader();
            foreach (var item in list)
            {
                toExport.Add(item.ToCsv());
            }
            var ExcelPkg = CreateExcel(Header, toExport, title);
            //at controller
            // return File(ExcelPkg.GetAsByteArray(), "application/vnd.ms-excel", title + ".xlsx");
            //at win
            FileInfo file = new FileInfo(Directory.GetCurrentDirectory() + @"\" + DateTime.Now.ToString("yyyyMMddHHmmssffftt") + "_WithHeader.xlsx");
            ExcelPkg.SaveAs(file);
        }

        public static ExcelPackage CreateExcel(List<string> listValue)
        {
            int colIndex = 1, rowIndex = 1;
            var pkg = PrepareExcelHeader();
            var ws = pkg.Workbook.Worksheets[1];
            var cell = ws.Cells[rowIndex, colIndex];
            foreach (var value in listValue)
            {
                cell = ws.Cells[rowIndex, colIndex];
                cell.Value = value;
                colIndex++;
                colIndex = 1;
                rowIndex++;
            }
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            return pkg;
        }

        public static ExcelPackage CreateExcel(string[] Headers, List<string[]> Rows, string Title = "Report")
        {
            int colIndex = 1, rowIndex = 4;
            var pkg = PrepareExcelHeader(Title, Headers);
            var ws = pkg.Workbook.Worksheets[1];
            var cell = ws.Cells[rowIndex, colIndex];

            var border = ws.Cells[rowIndex, colIndex, Rows.Count + rowIndex - 1, Headers.Length].Style.Border;
            border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;


            foreach (var rowVal in Rows)
            {
                foreach (var CellValue in rowVal)
                {
                    cell = ws.Cells[rowIndex, colIndex];
                    cell.Value = CellValue;
                    cell.Merge = true;
                    colIndex++;
                }
                colIndex = 1;
                rowIndex++;
            }
            

            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            return pkg;
        }
        public static ExcelPackage PrepareExcelHeader(string heading, string[] colnames)
        {
            var ExcelPkg = new ExcelPackage();
            ExcelPkg.Workbook.Worksheets.Add("Account Info");
            ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets[1];
            ws.Name = "Account Info";
            ws.Cells.Style.Font.Size = 11;
            ws.Cells.Style.Font.Name = "Calibri";
            ws.Cells[1, 1].Value = heading;
            ws.Cells[1, 1, 1, colnames.Length].Merge = true;
            ws.Cells[1, 1, 1, colnames.Length].Style.Font.Bold = true;
            ws.Cells[1, 1, 1, colnames.Length].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[1, 1, 1, colnames.Length].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            ws.Cells[1, 1, 1, colnames.Length].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            
            int colIndex = 1, rowIndex = 3;
            var cell = ws.Cells[rowIndex, colIndex];
            foreach (var col in colnames)
            {
                cell = ws.Cells[rowIndex, colIndex];
                cell.Value = col;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor.SetColor(Color.Black);
                cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                colIndex++;
            }




            return ExcelPkg;
        }

        public static ExcelPackage PrepareExcelHeader()
        {
            var ExcelPkg = new ExcelPackage();
            ExcelPkg.Workbook.Worksheets.Add("Reconciliation File");
            ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets[1];
            ws.Name = "Reconciliation File";
            ws.Cells.Style.Font.Size = 11;
            ws.Cells.Style.Font.Name = "Calibri";
            return ExcelPkg;
        }
    }

    public class VehiclesListModel
    {
        // x.CardNo, x.VehRegtNo, x.SelectedVehModel, x.VehRegDate, x.SelectedVehType, x.OdoMeterReading, x.OdoMeterUpdate, x.SelectedSts, x.PolicyExpDate, x.XrefCardNo, x.SelectedCardType, x.SkdsQuota
        public string[] ToCsv()
        {
            return new string[] { CardNo, VehRegtNo, SelectedVehModel, VehRegDate, SelectedVehType, OdoMeterReading, OdoMeterUpdate, SelectedSts, PolicyExpDate, XrefCardNo, SelectedCardType, SkdsQuota };
        }
        public string[] CsvHeader()
        {
            return new string[] { "Card No", "Vehicle Regn.No", "Vehicle Model", "Vehicle Regn.Date", "Vehicle Type", "Odometer Reading", "Odometer Update", "Status", "Policy Expiry Date", "XRef Card No", "Card Type", "SKDS Quota" };
        }

        [DisplayName("Card No")]
        //[RegularExpression(@"^[0-9]{16,19}$", ErrorMessage = "Card No Range = 16 to 19 digit")]
        public string CardNo { get; set; }
        [DisplayName("Card Type")]
        public string SelectedCardType { get; set; }
        //public IEnumerable<SelectListItem> CardType { get; set; }
        [DisplayName("Card Terminated")]
        public string CardTerminated { get; set; }
        [DisplayName("Card Expiry")]
        public string CardExpiry { get; set; }
        [DisplayName("Vehicle Registration No")]
        public string VehRegtNo { get; set; }
        [DisplayName("Vehicle Maker")]
        public string SelectedVehMaker { get; set; }
        //public IEnumerable<SelectListItem> VehMaker { get; set; }
        [DisplayName("Vehicle Model")]
        public string SelectedVehModel { get; set; }
        //public IEnumerable<SelectListItem> VehModel { get; set; }
        [DisplayName("Vehicle Registration Date")]
        public string VehRegDate { get; set; }
        [DisplayName("Vehicle Year")]
        public string SelectedVehYr { get; set; }
        //public IEnumerable<SelectListItem> VehYr { get; set; }
        [DisplayName("Vehicle Type")]
        public string SelectedVehType { get; set; }
        //public IEnumerable<SelectListItem> VehType { get; set; }
        [DisplayName("Vehicle Color")]
        public string SelectedVehColor { get; set; }
        //public IEnumerable<SelectListItem> VehColor { get; set; }
        [DisplayName("Odometer Reading")]
        public string OdoMeterReading { get; set; }
        [DisplayName("Odometer Update")]
        public string OdoMeterUpdate { get; set; }
        [DisplayName("Road Tax Expiry Date")]
        public string RoadTaxExpDate { get; set; }
        [DisplayName("Road Tax Amount")]
        //decimalvalidationbug
        public string RoadTaxAmt { get; set; }
        [DisplayName("Renewal Period")]
        //[RegularExpression(@"[-+]?[0-9]*\.?[0-9]?[0-9]", ErrorMessage = "Numbers only")]
        public int? RenewalPeriod { get; set; }
        [DisplayName("Insurer Cd")]
        public string InsurerCd { get; set; }
        [DisplayName("Premium Amount")]
        public string PremiumAmt { get; set; }
        [DisplayName("PIN")]
        public string pin { get; set; }
        [DisplayName("Policy No")]
        public string PolicyNo { get; set; }
        [DisplayName("Policy Expiry Date")]
        public string PolicyExpDate { get; set; }
        [DisplayName("Policy Amount")]
        //decimalvalidationbug
        public string PolicyAmt { get; set; }
        [DisplayName("Status")]
        public string SelectedSts { get; set; }
        //public IEnumerable<SelectListItem> Sts { get; set; }
        [DisplayName("SKDS Indicator")]
        public bool SkdsInd { get; set; }
        public string RawSKDSInd { get; set; }
        [DisplayName("SKDS Quota")]
        //[RegularExpression(@"[-+]?[0-9]*\.?[0-9]?[0-9]", ErrorMessage = "Numbers only")]
        public string SkdsQuota { get; set; }
        [DisplayName("Vehicle Service Date")]
        public string VehicleServiceDate { get; set; }
        //[DisplayName("VRN")]
        //public string VRN { get; set; }
        [DisplayName("Xref CardNo")]
        //[RegularExpression(@"^[0-9]{16,19}$", ErrorMessage = "XrefCardNo Range = 16 to 19 digit")]
        public string XrefCardNo { get; set; }
        [DisplayName("Descpription")]
        public string Descp { get; set; }

        public string AppcId { get; set; }


    }
}
