using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoadComparison
{
    class ExcelOperation
    {
        List<BladedData> bladedDatas;
        BladedDataOperation bladedDataOperation;
        Excel.Application app;
        Excel.Workbooks wbs;
        Excel.Workbook wb;

        public ExcelOperation()
        {
            bladedDatas = new List<BladedData>();
            bladedDataOperation = new BladedDataOperation();
            app = new Excel.ApplicationClass();
            this.InitialExcelSetting();
            this.GetDataFromBladedResults();
        }
        
        public void CreateMainUltimateLoadsSheet()
        {
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets["MainUltimateLoads"];
            if (bladedDatas.Count < 1)
            {
                Console.WriteLine("CreateMainUltimateLoadsSheet-bladedDatas.Count < 1 error!");
            }
            else
            {
                int col = 1; 
                foreach (BladedData b in bladedDatas)
                {
                    int row = 1;
                    //机组名称
                    Excel.Range rb = ws.get_Range(ws.Cells[row, col], ws.Cells[row++, col+4]);
                    rb.Merge();
                    rb.Value = b.turbineMainCompenontResult.turbineID;
                    //主要部件名称和数据
                    var comBase = bladedDatas[0].turbineMainCompenontResult.results.ultmateData.component;
                    var comNum = 0;
                    foreach (var com in b.turbineMainCompenontResult.results.ultmateData.component)
                    {
                        Excel.Range rHeader = ws.get_Range(ws.Cells[row, col], ws.Cells[row++, col+4]);
                        rHeader.Merge();
                        rHeader.Value = com.name;
                        Excel.Range rData = ws.get_Range(ws.Cells[row, col], ws.Cells[(row = row+8), col+4]);
                        //计算对比值
                        for (int i = 0; i < 8; i++)
                        {
                            var baseValue = Convert.ToSingle(comBase[comNum].resultMatrix[i, 2]);
                            var compValue = Convert.ToSingle(com.resultMatrix[i, 2]);
                            var divValue = (compValue / baseValue).ToString(".000");
                            com.resultMatrix[i, 4] = divValue;
                        }
                        //数据放入excel表中
                        rData.Value = com.resultMatrix;

                        //转化为数字，将>1.05的数值标红
                        Excel.Range numData = ws.get_Range(ws.Cells[row-8, col + 4], ws.Cells[row, col + 4]);
                        numData.FormulaR1C1 = "1";
                        Excel.FormatCondition condition1 = (Excel.FormatCondition)numData.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "1.05", Type.Missing);
                        condition1.Interior.Color = 13551615;

                        row++;
                        comNum++;
                    }
                    col = col + 6;
                }
            }
        }

        public void CreateMainEequivalentFatigueLoadsSheet()
        {
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets["MainEequivalentFatigueLoads"];
            if (bladedDatas.Count < 1)
            {
                Console.WriteLine("CreateMainEequivalentFatigueLoadsSheet-bladedDatas.Count < 1 error!");
            }
            else
            {
                int colStart = 1;
                int colCount = 4;
                foreach (BladedData b in bladedDatas)
                {
                    int rowStart = 1;
                    int rowCount = 6;
                    //机组名称
                    Excel.Range rb = ws.get_Range(ws.Cells[rowStart, colStart], ws.Cells[rowStart++, colStart + colCount]);
                    rb.Merge();
                    rb.Value = b.turbineMainCompenontResult.turbineID;
                    //主要部件名称和数据
                    var comBase = bladedDatas[0].turbineMainCompenontResult.results.equivalentFatigueData.component;
                    var comNum = 0;
                    foreach (var com in b.turbineMainCompenontResult.results.equivalentFatigueData.component)
                    {
                        Excel.Range rHeader = ws.get_Range(ws.Cells[rowStart, colStart], ws.Cells[rowStart++, colStart + colCount]);
                        rHeader.Merge();
                        rHeader.Value = com.name;
                        
                        for (int i = 0; i < 10; i++)
                        {
                            for(int j= 0; j< 6; j++ )
                            {
                                var baseValue = Convert.ToSingle(comBase[comNum].resultMatrix[i+1, j+1]);
                                var compValue = Convert.ToSingle(com.resultMatrix[i + 1, j + 1]);
                                var divValue = (compValue / baseValue).ToString("G3");
                                com.resultMatrix[i + 1,  j + 7] = divValue;
                            }
                        }
                        //只输出m=4&&m=10
                        string[,] tempMatrix = new string[7, 5];
                        for(int i=0; i<7; i++)
                        {
                            tempMatrix[i, 0] = com.resultMatrix[0, i];
                            tempMatrix[i , 1] = com.resultMatrix[2, i];
                            tempMatrix[i, 2] = com.resultMatrix[8, i];
                            if(i == 0)
                            {
                                tempMatrix[i, 3] = com.resultMatrix[2, 0];
                                tempMatrix[i, 4] = com.resultMatrix[8, 0];
                            }
                            else
                            {
                                tempMatrix[i, 3] = com.resultMatrix[2, i + 6];
                                tempMatrix[i, 4] = com.resultMatrix[8, i + 6];
                            }
                        }
                        Excel.Range rData = ws.get_Range(ws.Cells[rowStart, colStart], ws.Cells[(rowStart = rowStart + rowCount), colStart + colCount]);
                        rData.Value = tempMatrix;

                        //转化为数字，将>1.05的数值标红
                        Excel.Range numData = ws.get_Range(ws.Cells[rowStart + 1 - rowCount, colStart + 3], ws.Cells[rowStart, colStart + colCount]);
                        numData.FormulaR1C1 = "1";
                        Excel.FormatCondition condition1 = (Excel.FormatCondition)numData.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "1.05", Type.Missing);
                        condition1.Interior.Color = 13551615;

                        rowStart = rowStart + 2;
                        comNum++;
                    }
                    colStart = colStart + colCount + 2;
                }
            }
        }

            void CreateMainEquivalentFatigueLoadsSheet()
        {
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.Add();
        }

        public void CreateMainUltimateLoadsThermodynamicChartSheet()    //创建热力图
        {
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets["UltimateThermodynamicChart"];
//             Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, wsTemp, Type.Missing, Type.Missing);
//             ws.Name = "UltimateThermodynamicChart";
            if (bladedDatas.Count < 1)
            {
                Console.WriteLine("CreateMainUltimateLoadsSheet-bladedDatas.Count < 1 error!");
            }
            else
            {
                int col = 1;
                int colCount = 0;
                foreach (BladedData b in bladedDatas)
                {
                    int row = 1;
                    //机组名称
                    Excel.Range rb = ws.get_Range(ws.Cells[row, col], ws.Cells[row, col + 4]);
                    rb.Merge();
                    rb.Value = b.turbineMainCompenontResult.turbineID;
                    //主要部件名称和数据
 //                   var comBase = bladedDatas[0].turbineMainCompenontResult.results.ultmateData.component;
                    foreach (var com in b.turbineMainCompenontResult.results.ultmateData.component)
                    {
                        //部件名称
                        Excel.Range rHeader = ws.get_Range(ws.Cells[row + 1, col], ws.Cells[row+1, col + 4]);
                        rHeader.Merge();
                        rHeader.Value = com.name;
                        colCount = com.mainDlcData.Count ;  //数据 
                        int rowCount = com.variableHeader.Length;
                        //列表头
                        Excel.Range rHeadercol = ws.get_Range(ws.Cells[row+2, col + 1], ws.Cells[row + 2, col + colCount]);
                        rHeadercol.Value = com.dlcNameList.ToArray();
                        //行表头
                        Excel.Range rHeaderrow = ws.get_Range(ws.Cells[(row + 3), col], ws.Cells[(row+3+ rowCount-1), col]);
                        rHeaderrow.Value = com.variableHeader;
                        //数据放入excel表中
                        Excel.Range rData = ws.get_Range(ws.Cells[(row + 3), col + 1], ws.Cells[(row + 3 + rowCount-1), col+colCount]);
                        rData.Value = com.dlcMaxValueList;
                        ////色阶设置
                        Excel.Range numData = ws.get_Range(ws.Cells[(row + 3), col + 1], ws.Cells[(row + 3 + rowCount - 1), col + colCount]);
                        //numData.FormulaR1C1 = "1";
                        Excel.ColorScale condition1 = (Excel.ColorScale)numData.FormatConditions.AddColorScale(3);
                        condition1.SetFirstPriority();
                        Excel.ColorScaleCriterion colorScaleCriterion = condition1.ColorScaleCriteria.Item[1];
                        condition1.ColorScaleCriteria.Item[1].Type = Microsoft.Office.Interop.Excel.XlConditionValueTypes.xlConditionValueLowestValue;
                        condition1.ColorScaleCriteria.Item[1].FormatColor.Color = 8109667;
                        condition1.ColorScaleCriteria.Item[2].Type = Excel.XlConditionValueTypes.xlConditionValuePercentile;
                        condition1.ColorScaleCriteria.Item[2].Value = 50;
                        condition1.ColorScaleCriteria.Item[2].FormatColor.Color = 8711167;
                        condition1.ColorScaleCriteria.Item[3].Type = Excel.XlConditionValueTypes.xlConditionValueHighestValue;
                        condition1.ColorScaleCriteria.Item[3].FormatColor.Color = 7039480;
                        row = row + 3 + rowCount +1;
                    }
                    col = col + 18;
                }
            }
            SaveAsExcelFile(bladedDatas[0]);
        }

        void GetDataFromBladedResults()
        {
            this.GetPostPathFromInfoSheet();
            bladedDatas = bladedDataOperation.GetMainCompinentLoadsResult(bladedDatas);
        }
        void InitialExcelSetting()
        {
            if (app == null)
            {
                Console.WriteLine("Excel无法启动");
                return;
            }
            app.Visible = true;
            wbs = app.Workbooks;
            //wb = wbs.Add(Missing.Value);
            string templatePath = Directory.GetCurrentDirectory() + "\\LoadComparisonTemplate.xlsx";
            wb = wbs.Open(templatePath);
        }

        void GetPostPathFromInfoSheet()
        {
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets["Info"];
            for(int i=0; ; i++)
            {
                BladedData bladedDataTemp = new BladedData();
                if (ws.get_Range(ws.Cells[i*2 + 2, 1], ws.Cells[i*2 + 2, 1]).Value == (object)null)
                {
                    break;
                }
                else
                {
                    bladedDataTemp.turbineMainCompenontResult.turbineID = ws.get_Range(ws.Cells[i*2 + 2, 1], ws.Cells[i*2 + 2, 1]).Value.ToString();
                    bladedDataTemp.ultimateLoads.path = ws.get_Range(ws.Cells[i * 2 + 2, 2], ws.Cells[i * 2 + 2, 2]).Value.ToString();
                    bladedDataTemp.equivalentFatigueLoads.path = ws.get_Range(ws.Cells[i*2 + 3, 2], ws.Cells[i*2 + 3, 2]).Value.ToString();
                    bladedDatas.Add(bladedDataTemp);
                }
            }
            GetMainComPathFromPostPath();
        }

        void GetMainComPathFromPostPath() // 从excel表中获取post路径
        {
            foreach(BladedData dd in bladedDatas)
            {
                string[] comPaths = Directory.GetDirectories(dd.ultimateLoads.path);
                foreach(string s in comPaths)
                {
                    BladedData.TurbineMainCompenontResult.Results.MainComponentDataStruct com
                   = new BladedData.TurbineMainCompenontResult.Results.MainComponentDataStruct();
                    com.path = s;
                    string[]temp = s.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                    com.name = temp.LastOrDefault();
                    dd.turbineMainCompenontResult.results.ultmateData.component.Add(com);
                }

                string[] comPaths2 = Directory.GetDirectories(dd.equivalentFatigueLoads.path);
                foreach (string s in comPaths2)
                {
                    BladedData.TurbineMainCompenontResult.Results.MainComponentDataStruct com
                    = new BladedData.TurbineMainCompenontResult.Results.MainComponentDataStruct();
                    com.path = s;
                    string[] temp = s.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                    com.name = temp.LastOrDefault();
                    dd.turbineMainCompenontResult.results.equivalentFatigueData.component.Add(com);
                }
            }
            bladedDataOperation.GetMainCompinentLoadsResult(bladedDatas);
        }

        void SaveAsExcelFile(BladedData b)
        {
            var dir = Directory.GetCurrentDirectory();
            var filePath = dir +"\\"+ b.turbineMainCompenontResult.turbineID + "-" + "Comparison" + DateTime.Today.ToString("yyyyMMdd") + ".xlsx";
            wb.SaveAs(filePath);
//            app.Quit();
        }

        public void QuitExcel()
        {
            app.Quit();
        }
    }
}
