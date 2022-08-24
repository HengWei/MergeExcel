using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Excel;

namespace MergeExcel
{
    internal class Service
    {
        public void ReadExcel(string _Path, int _TitleRow)
        {
            using (XLWorkbook workbook = new XLWorkbook(_Path))
            {
                var newWorkSheet = workbook.AddWorksheet();

                int sheetNo = 1;
                int newSheetRow = 1;

                foreach (var worksheet in workbook.Worksheets)
                {
                    var firstCell = worksheet.FirstCellUsed();
                    var lastCell = worksheet.LastCellUsed();

                    int lastCellNum = lastCell.Address.ColumnNumber;

                    // 使用資料起始、結束 Cell，來定義出一個資料範圍
                    var data = worksheet.Range(firstCell.Address, lastCell.Address);

                    // 將資料範圍轉型
                    var table = data.AsTable();
                    int i = 0;

                    foreach (var row in table.Rows())
                    {
                        i++;
                        if (i < _TitleRow)
                        {                            
                            continue;
                        }

                        if ((sheetNo == 1) && (i == _TitleRow))
                        {
                            for (int j = 1; j <= lastCellNum; j++)
                            {
                                newWorkSheet.Row(newSheetRow).Cell(j).SetValue(row.Cell(j).Value);
                            }

                            newSheetRow++;
                            sheetNo++;
                            continue;
                        } 
                        else if(i == _TitleRow)
                        {
                            continue;
                        }

                        if(row.Cell(1).Value.ToString().IndexOf("合計")>-1)
                        {
                            break;
                        }
                        else
                        {
                            for (int j = 1; j <= lastCellNum; j++)
                            {
                                newWorkSheet.Row(newSheetRow).Cell(j).SetValue(row.Cell(j).Value);
                            }
                            newSheetRow++;
                        }
                    }

                    sheetNo++;
                }


                workbook.Save();

                
            }
        }

    }
}
