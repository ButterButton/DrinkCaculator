using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;

namespace DrinksCaculator
{
    interface ISheetsAPI
    {
        SheetsService ShService { get; }
    }

    class SheetsAPIService : GoogleService, ISheetsAPI
    {
        public SheetsService ShService { get; set; }

        public string SpreadSheetID { get; set; }

        public string Range { get; set; }

        public SheetsAPIService(string ShID, string Rg, string CrednetialLoaction) : base(CrednetialLoaction)
        {
            SpreadSheetID = ShID;
            Range = Rg;
            Scopes = new string[] { SheetsService.Scope.Spreadsheets };
            ShService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = Credential,
                ApplicationName = ApplicationName,
            });
        }

        public ValueRange GetSheet()
        {
            SpreadsheetsResource.ValuesResource.GetRequest GetRequest = ShService.Spreadsheets.Values.Get(SpreadSheetID, Range);
            return GetRequest.Execute();
        }

        public void UpdateSheet(string Range, IList<IList<Object>> UpdateValue)
        {
            this.Range = Range;
            SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum valueInputOption = (SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum)2;
            ValueRange requestBody = new ValueRange();
            requestBody.Values = UpdateValue;
            SpreadsheetsResource.ValuesResource.UpdateRequest requestUpdate = ShService.Spreadsheets.Values.Update(requestBody, SpreadSheetID, Range);
            requestUpdate.ValueInputOption = valueInputOption;
            requestUpdate.Execute();
        }

        public void BackGroundColorChange(int GroupCount, bool IsSpace = false)
        {
            BatchUpdateSpreadsheetRequest Bussr = new BatchUpdateSpreadsheetRequest();
            int Sr = 4, Er = 5, Grade = 2;
            Bussr.Requests = new List<Request>();
            var userEnteredFormat = new CellFormat();

            if (IsSpace)
            {
                userEnteredFormat.BackgroundColor = null;
                Sr = Sr - 1;
                GroupCount = GroupCount * 3;
                Grade = 1;
            }
            else
            {
                userEnteredFormat.BackgroundColor = new Color() { Red = 100, Green = 80, Blue = 80 };
                GroupCount = (GroupCount / 2) - 1;
            }

            for (int n = 0; n <= GroupCount; n++)
            {
                var updateCellsRequest = new Request()
                {
                    RepeatCell = new RepeatCellRequest()
                    {
                        Range = new GridRange()
                        {
                            //工作表的ID
                            SheetId = 739704159,
                            StartColumnIndex = 1,
                            StartRowIndex = Sr,
                            EndColumnIndex = 7,
                            EndRowIndex = Er
                        },
                        Cell = new CellData()
                        {
                            UserEnteredFormat = userEnteredFormat
                        },
                        Fields = "UserEnteredFormat(BackgroundColor)"
                    }
                };

                Bussr.Requests.Add(updateCellsRequest);
                Sr = Sr + Grade;
                Er = Er + Grade;
            }

            ShService.Spreadsheets.BatchUpdate(Bussr, SpreadSheetID).Execute();
        }
    }
}
