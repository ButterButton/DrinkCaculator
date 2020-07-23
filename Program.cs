using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using DrinksCaculator;
using System.Linq;
using DrinksCaculator.DrinksObject.Drink;
using DrinksCaculator.DrinksObject.Member;
using DrinksCaculator.DrinksObject.Order;

namespace SheetsQuickstart
{
    class Program
    {
        static void Main(string[] args)
        {
            var InititalFile = new FileStream("../../InitialData.txt", FileMode.Open, FileAccess.Read);
            var Reader = new StreamReader(InititalFile);
            var GoogleCredentialLocation = Reader.ReadLine().Split('=')[1];
            var MemberFileLocation = Reader.ReadLine().Split('=')[1];
            var MemberFile = new StreamReader(new FileStream(MemberFileLocation, FileMode.Open, FileAccess.Read)).ReadToEnd();
            var ApexStaffs = JsonConvert.DeserializeObject<ApexMembers>(MemberFile).Staffs;

            // Define request parameters.
            //range "sheet的名稱!格子位置(位置類似座標的概念)"
            String SpreadSheetID = "your Sheets ID";
            String Range = "工作表1!B4:G33";

            // Create Google Sheets API service.
            var SheetsService = new SheetsAPIService(SpreadSheetID, Range, CrednetialLoaction: GoogleCredentialLocation);

            //Get
            #region Get
            IList<IList<Object>> SheetsValues = SheetsService.GetSheet().Values;
            var DrinkList = new List<Drink>();

            if (SheetsValues != null && SheetsValues.Count > 0)
            {
                Console.WriteLine("名字, 飲料, 冰量, 甜度, 價格");
                foreach (var row in SheetsValues)
                {
                    if (row.Count != 0)
                    {
                        var TempDrink = new Drink();

                        TempDrink.StaffName = row[0].ToString();
                        TempDrink.DrinkName = row[1].ToString();
                        TempDrink.Sugar = row[2].ToString();
                        TempDrink.Ice = row[3].ToString();
                        TempDrink.Price = Convert.ToInt32(row[4]);
                        DrinkList.Add(TempDrink);

                        Console.WriteLine("{0}, {1}, {2}, {3}, {4}", row[0], row[1], row[2], row[3], row[4]);
                    }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
                Console.Read();
            }
            #endregion

            //Process
            #region Process
            DrinkList = DrinkList.OrderBy(o => o.DrinkName).ThenBy(o => o.Sugar).ThenBy(o => o.Ice).ThenBy(o => o.StaffName).ToList();
            var DrinkListAfterGroup = DrinkList.GroupBy(g => new { g.DrinkName, g.Sugar, g.Ice });
            var OrderList = new List<Order>();

            DrinkListAfterGroup.ToList().ForEach(f =>
            {
                var NewOrder = new Order();

                NewOrder.DrinkName = f.Key.DrinkName;
                NewOrder.Sugar = f.Key.Sugar;
                NewOrder.Ice = f.Key.Ice;
                NewOrder.Price = f.Sum(s => s.Price);
                NewOrder.SumOfCups = f.Count();
                NewOrder.PoepleOfOrdered = DrinkList.Where(w => w.DrinkName == f.Key.DrinkName && w.Sugar == f.Key.Sugar && w.Ice == f.Key.Ice).Select(s => s.StaffName).ToArray();

                OrderList.Add(NewOrder);
            });

            //未訂購者檢查

            var NotInOrder = new List<string>();
            var StaffsOfOrdered = new List<string>();

            foreach (var D in DrinkList)
            {
                StaffsOfOrdered.Add(D.StaffName.ToLower());
            }

            NotInOrder.AddRange(ApexStaffs.Where(w => !w.NickName.Where(ww => StaffsOfOrdered.Contains(ww.ToLower())).Any()).Select(s => s.StaffName));
            NotInOrder = NotInOrder.Select(s => s).Distinct().ToList();
            #endregion

            //Delete
            #region Delete

            //Request RequestBody = new Request()
            //{
            //    DeleteDimension = new DeleteDimensionRequest()
            //    {
            //        Range = new DimensionRange()
            //        {
            //            SheetId = 739704159,
            //            Dimension = "COLUMNS",
            //            StartIndex = 1,
            //            EndIndex = 7
            //        }
            //    }
            //};
            //List<Request> RequestContainer = new List<Request>();
            //RequestContainer.Add(RequestBody);
            //BatchUpdateSpreadsheetRequest DeleteRequest = new BatchUpdateSpreadsheetRequest();
            //DeleteRequest.Requests = RequestContainer;

            //SpreadsheetsResource.BatchUpdateRequest Deletion = new SpreadsheetsResource.BatchUpdateRequest(ShService, DeleteRequest, SpreadSheetID);
            //Deletion.Execute();
            #endregion

            //Update
            #region Update

            IList<IList<Object>> UpdateTableValue = new List<IList<object>>();
            foreach (var Item in OrderList)
            {
                var OrderedName = "";
                foreach (var TempOrderedName in Item.PoepleOfOrdered)
                {
                    OrderedName = OrderedName + TempOrderedName + "_";
                }
                var StringList = new List<object>();

                StringList.Add(Item.DrinkName);
                StringList.Add(Item.Sugar);
                StringList.Add(Item.Ice);
                StringList.Add(OrderedName.TrimEnd('_'));
                StringList.Add(Item.SumOfCups);
                StringList.Add(Item.Price);

                UpdateTableValue.Add(StringList);
            }

            UpdateTableValue.Add(new List<object> { "", "", "", "", "共" + OrderList.Sum(s => s.SumOfCups) + "杯","共" + OrderList.Sum(s => s.Price) + "元"});
            UpdateTableValue.Add(new List<object> { });
            UpdateTableValue.Add(new List<object> { "尚未訂購者" });

            var NoOrderManList = new List<object>();

            foreach (var NoOrderStaff in NotInOrder)
            {
                NoOrderManList.Add(NoOrderStaff);
                if (NoOrderManList.Count % 5 == 0)
                {
                   UpdateTableValue.Add(NoOrderManList);
                   NoOrderManList = new List<object>();
                }
            }
            if (NoOrderManList.Count > 0)
            {
                UpdateTableValue.Add(NoOrderManList);
            }

            //空的格子
            IList<IList<Object>> UpdateTableValueSpace = new List<IList<object>>();
            for (int n = 0; n <= (UpdateTableValue.Count)*3; n++)
            {
                UpdateTableValueSpace.Add(new List<object> { "", "", "", "", "", "", "" });
            }

            //清空格子
            SheetsService.UpdateSheet("工作表3!B4", UpdateTableValueSpace);

            //更新格子
            SheetsService.UpdateSheet("工作表3!B4", UpdateTableValue);

            //BackgroudColor顏色更改
            #region BackgroundColorChange
            var DrinkListAfterGroupCount = DrinkListAfterGroup.Count();

            //清空背景色
            SheetsService.BackGroundColorChange(DrinkListAfterGroupCount, true);

            //填入背景色
            SheetsService.BackGroundColorChange(DrinkListAfterGroupCount);
            #endregion

            #endregion

            Console.Read();
        }
    }
}
