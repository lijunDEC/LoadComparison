using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace LoadComparison
{
    public class BladedDataOperation
    {
        public BladedData br, hr, hs, tt, tb;
        private string[] headerU1 = new string[] { "Mx", "My", "Mxy", "Mz", "Fx", "Fy", "Fxy", "Fz" };
        private string[] headerU2 = new string[] { "Mx", "My", "Mz", "Myz", "Fx", "Fy", "Fz", "Fyz" };
        private string[] headerF = new string[] { "m", "Mx", "My", "Mz", "Fx", "Fy", "Fz" };
        public BladedDataOperation()
        {
            InitialBladedDataOperation();
        }

        private void InitialBladedDataOperation()
        {
            br = new BladedData();
            hr = new BladedData();
            hs = new BladedData();
            tt = new BladedData();
            tb = new BladedData();
        }

        private BladedData GetBladedResults(string filePath, string resultType)//打开文件获取数据结果
        {
            BladedData result = new BladedData();
            if (filePath == String.Empty)
                Console.WriteLine("File path is null!");
            var mode = CheckBladedResultType(resultType);
            switch (mode)
            {
                case 0:
                    result.ultimateLoads = GetUltimateLoadsResultsFromFile(filePath).ultimateLoads;
                    break;
                case 1:
                    result.equivalentFatigueLoads = GetEquivalentFatigueLoadsResultsFromFile(filePath).equivalentFatigueLoads;
                    break;
                default:
                    Console.WriteLine("CheckBladedResultType error!");
                    break;
            }
            return result;
        }

        private int CheckBladedResultType(string resultType)
        {
            if (resultType.Equals("UltimateLoads")) return 0;
            else if (resultType.Equals("EquivalentFatigueLoads")) return 1;
            else
                return -1;
        }

        private string GetBladedRunPath(string filePath, string dlcName)
        {
            var pjFilePath = Directory.GetFiles(filePath, "*.$PJ")[0];  //获取.$pj文件路径
            var buff = File.ReadAllText(pjFilePath);
            var endPos = buff.IndexOf("\\" + dlcName) + dlcName.Length;
            string dlcPathTemp = null;
            for (int i = endPos; i > 0; i--)
            {
                if (buff[i] == '\t')
                {
                    dlcPathTemp = buff.Substring(i + 1, endPos - i);
                    break;
                }
            }
            return dlcPathTemp;
        }

        private BladedData GetUltimateLoadsResultsFromFile(string filePath)
        {
            BladedData ultimateLoadsResults = new BladedData();
            var pjFilePath = Directory.GetFiles(filePath, "*.$PJ")[0];  //获取.$pj文件路径
            var fileNameTemp = Path.GetFileNameWithoutExtension(pjFilePath);
            var filePathTemp = filePath + "\\" + fileNameTemp + ".$MX";
            var buff = File.ReadAllText(filePathTemp);
            string[,] dataResults = new string[16, 10];
            string temp = buff.Replace(" ", "");
            string[] data = temp.Split('\n');

            int j = 0;
            int i = 0;

            foreach (string ii in data)
            {
                i = 0;
                string[] temp1 = ii.Split('\t');

                foreach (string jj in temp1)
                {
                    if (j > 1 && i > 1 && j < 18 && i < 12)
                    {
                        dataResults[j - 2, i - 2] = jj;
                    }
                    i++;
                }
                j++;
            }
            if (j > 19)
            {
                Console.WriteLine("GetUltimateLoadsResults error!");
                return (null);
            }
            ultimateLoadsResults.ultimateLoads.arrayResults = dataResults;
            return ultimateLoadsResults;
        }

        private float[,] GetUltimateLoadsDlcMaxValueFromFile(List<string> filePath)  //获取每一个工况的极值矩阵；
        {
            float[,] dataResults = new float[16, 8]; //16行8列数值矩阵
            float[,] maxResults = new float[16, 8]; //16行8列数值矩阵            
            int j = 0;
            int i = 0;
            foreach (string ss in filePath)
            {
                var buff = File.ReadAllText(ss);
                string temp = buff.Replace(" ", "");
                string[] data = buff.Split('\n');
                j = 0;
                foreach (string ii in data)
                {
                    i = 0;
                    string[] temp1 = ii.Split('\t');

                    foreach (string jj in temp1)
                    {
                        if (j > 1 && i > 2 && j < 18 && i < 11)
                        {
                            dataResults[j - 2, i - 3] = Convert.ToSingle(jj);
                        }
                        i++;
                    }
                    j++;
                }
                maxResults = GetMatrixMaxValue(maxResults, dataResults);
            }
            return maxResults;
        }

        private float[,] GetMatrixMaxValue(float[,] maxResults, float[,] dataResults)       //返回矩阵中每个值取两者最大
        {
            var row = maxResults.GetLength(0);
            var col = maxResults.GetLength(1);
            float[,] maxTemp = new float[16, 8];
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    maxTemp[i, j] = Math.Max(Math.Abs(maxResults[i, j]), Math.Abs(dataResults[i, j]));
                }
            }
            return maxTemp;
        }

        private BladedData.TurbineMainCompenontResult.Results.MainComponentDataStruct GetMatrixMaxValue(BladedData.TurbineMainCompenontResult.Results.MainComponentDataStruct mainCom)       //格式化矩阵输出单个工况各变量的最大值
        {
            float[] maxValue = new float[8];
            List<string> dlcNameList = new List<string>();
            List<float[]> dlcMaxValueList = new List<float[]>();
            float[,] templist = new float[8, mainCom.mainDlcData.Count];
            int n = 0;
            foreach (BladedData.TurbineMainCompenontResult.Results.MainDlcDataStruct dd in mainCom.mainDlcData)
            {
                float[,] dataResults = dd.resultMatrix; //数据
                dlcNameList.Add(dd.dlcName);  //列表头
                var row = dataResults.GetLength(0);
                var col = dataResults.GetLength(1);
                float[] maxTemp = new float[8];
                for (int j = 0; j < col; j++)
                {
                    templist[j, n] = Math.Max(Math.Abs(dataResults[2 * j, j]), Math.Abs(dataResults[2 * j + 1, j]));
                    maxValue[j] = Math.Max(templist[j, n], maxValue[j]);
                }
                n++;
            }
            for (int i = 0; i < templist.GetLength(0); i++) //转化为与最大值的百分比
            {
                for (int j = 0; j < templist.GetLength(1); j++)
                {
                    templist[i, j] = templist[i, j] / maxValue[i];
                }
            }

            mainCom.dlcNameList = dlcNameList;
            mainCom.dlcMaxValueList = templist;
            return mainCom;
        }


        private BladedData GetEquivalentFatigueLoadsResultsFromFile(string filePath)
        {
            BladedData results = new BladedData();
            string[,] dataResults = new string[10, 6];
            var fileNameTemp = Path.GetFileNameWithoutExtension(Directory.GetFiles(filePath, "*.$PJ")[0]);
            var filePathTemp = filePath + "\\" + fileNameTemp + ".$EQ";
            if (filePathTemp.Length <= 1 || !File.Exists(filePathTemp))
            {
                Console.WriteLine("GetEquivalentFatigueLoadsResults filePath error!");
                return (null);
            }

            var buff = File.ReadAllText(filePath + "\\" + fileNameTemp + ".$EQ");
            string[] temp = buff.Split('\n');

            int j = 0;
            int i = 0;
            foreach (string ii in temp)
            {
                if (ii.Contains("Case   0") && ((ii.Contains("FXT")) || (ii.Contains("FYT")) || (ii.Contains("FZT")) || (ii.Contains("MXT")) || (ii.Contains("MYT")) || (ii.Contains("MZT"))
                     || (ii.Contains("Fx")) || (ii.Contains("Fy")) || (ii.Contains("Fz")) || (ii.Contains("Mx")) || (ii.Contains("My")) || (ii.Contains("Mz"))))
                {
                    var tempii = ii.Remove(0, ii.IndexOf("Case   0") + "Case   0".Length);
                    string[] data = tempii.Split(new string[] { " ", "\r", "\t" }, StringSplitOptions.RemoveEmptyEntries);
                    for (j = 0; j < 10; j++)
                    {
                        dataResults[j, i] = data[j];
                    }
                    i++;
                }
            }
            if (i > 6)
            {
                Console.WriteLine("GetEquivalentFatigueLoadsResults error");
                return null;
            }
            results.equivalentFatigueLoads.arrayResults = dataResults;
            return results;
        }

        private string[,] GetFormatResult(string filePath, string[] header, string resultsType)//格式化excel输出数据
        {
            if (resultsType == "UltimateLoads")
            {
                BladedData brResults = new BladedData();
                string[,] brArrary = new string[9, 5];
                brResults = GetBladedResults(filePath, resultsType);//"EquivalentFatigueLoads"
                string[,] temp = brResults.ultimateLoads.arrayResults;

                for (int i = 1; i < 9; i++)//提取工况和载荷
                {
                    string maxmum = temp[2 * (i - 1), i];
                    string dlc1 = temp[2 * (i - 1), 0];
                    string minmum = temp[2 * (i - 1) + 1, i];
                    string dlc2 = temp[2 * (i - 1) + 1, 0];
                    string[] row = ComparisionStringNumValue(filePath, maxmum, minmum, dlc1, dlc2, header[i - 1]);
                    for (int j = 0; j < 4; j++)
                    {
                        brArrary[i - 1, j] = row[j];
                    }
                }
                return brArrary;
            }
            if (resultsType == "EquivalentFatigueLoads")
            {
                BladedData brResults = new BladedData();
                string[,] formatToExcelResult = new string[12, 13];
                brResults = GetBladedResults(filePath, resultsType);//"EquivalentFatigueLoads"
                string[,] temp = brResults.equivalentFatigueLoads.arrayResults;

                for (int j = 0; j < 7; j++)//添加表头
                {
                    formatToExcelResult[0, j] = header[j];
                }
                for (int i = 0; i < 11; i++)
                {
                    if (i < 1) formatToExcelResult[i, 0] = "m";
                    else
                    {
                        formatToExcelResult[i, 0] = (i + 2).ToString();
                    }
                }
                for (int i = 0; i < 6; i++)//添加内容
                {
                    for (int j = 0; j < 10; j++)
                    {
                        formatToExcelResult[j + 1, i + 1] = (Convert.ToSingle(temp[j, i]) / 1000.0f).ToString("G5");
                    }
                }
                return formatToExcelResult;
            }
            else
            {
                return null;
            }
        }

        private string[] ComparisionStringNumValue(string filePath, string maxmum, string minmum, string dlc1, string dlc2, string variable = "Mx")
        {
            string[] col = new string[4];
            col[0] = variable;
            var x = Math.Abs(Convert.ToSingle(maxmum)) / 1000.0f;
            var y = Math.Abs(Convert.ToSingle(minmum)) / 1000.0f;
            if (x >= y)
            {
                col[1] = dlc1;
                col[2] = x.ToString("G5");
                col[3] = GetBladedRunPath(filePath, dlc1);
                return col;
            }
            else
            {
                col[1] = dlc2;
                col[2] = y.ToString("G5");
                col[3] = GetBladedRunPath(filePath, dlc2);
                return col;
            }
        }

        //public string[,] GetFormatBladeRootUltimateLoadsResult(string filePath)
        //{
        //    return GetFormatUltimateLoadsResult(filePath, headerU1);
        //}

        //public string[,] GetFormatStationHubUltimateLoadsResult(string filePath)
        //{
        //    return GetFormatUltimateLoadsResult(filePath, headerU2);
        //}

        //public string[,] GetFormatRotorHubUltimateLoadsResult(string filePath)
        //{
        //    return GetFormatUltimateLoadsResult(filePath, headerU2);
        //}

        //public string[,] GetFormatTowerTopUltimateLoadsResult(string filePath)
        //{
        //    return GetFormatUltimateLoadsResult(filePath, headerU1);
        //}

        //public string[,] GetTowerBaseUltimateLoadResult(string filePath)
        //{
        //    return GetFormatUltimateLoadsResult(filePath, headerU1);
        //}

        //private string[,] GetFormatEquivalentFatigueLoadsResult(string filePath, string[] header)
        //{
        //    string[,] formatToExcelResult = new string[11, 7];
        //    string[,] temp = GetEquivalentFatigueLoadsResultsFromFile(filePath).equivalentFatigueLoads.arrayResults;

        //    for(int j=0; j<7; j++)//添加表头
        //    {
        //        formatToExcelResult[0, j] = header[j];
        //    }
        //    for(int i=0; i<7; i++)//添加内容
        //    {
        //        for(int j=0; j<10; j++)
        //        {
        //            formatToExcelResult[j + 1, i] = temp[j, i];
        //        }
        //    }
        //    return formatToExcelResult;
        //}

        //public string[,] GetFormatBladedRootEquivalentFatigueLoadsResult(string filePath)
        //{
        //    return GetFormatEquivalentFatigueLoadsResult(filePath, headerF);
        //}

        //public string[,] GetFormatRotorHubEquivalentFatigueLoadsResult(string filePath)
        //{
        //    return GetFormatEquivalentFatigueLoadsResult(filePath, headerF);
        //}

        //public string[,] GetFormatStationHubEquivalentFatigueLoadsResult(string filePath)
        //{
        //    return GetFormatEquivalentFatigueLoadsResult(filePath, headerF);
        //}

        //public string[,] GetFormatTowerTopEquivalentFatigueLoadsResult(string filePath)
        //{
        //    return GetFormatEquivalentFatigueLoadsResult(filePath, headerF);
        //}

        //public string[,] GetFormatTowerBaseEquivalentFatigueLoadsResult(string filePath)
        //{
        //    return GetFormatEquivalentFatigueLoadsResult(filePath, headerF);
        //}

        public List<BladedData> GetMainCompinentLoadsResult(List<BladedData> bladedPath)
        {
            List<BladedData> bladedDlist = new List<BladedData>();

            foreach (var dd in bladedPath)  //打开每个地址文件读取各部件的地址
            {
                BladedData bladedDataTemp = new BladedData();
                bladedDataTemp.ultimateLoads.path = dd.ultimateLoads.path;
                bladedDataTemp.equivalentFatigueLoads.path = dd.equivalentFatigueLoads.path;
                bladedDataTemp.turbineMainCompenontResult.turbineID = dd.turbineMainCompenontResult.turbineID;

                var ComponentU = dd.turbineMainCompenontResult.results.ultmateData.component;
                var ComponentF = dd.turbineMainCompenontResult.results.equivalentFatigueData.component;
                foreach (var dic in ComponentU)
                {
                    dic.variableHeader = new string[headerU1.Length, 1];
                    if ((dic.name.Contains("Blade") || dic.name.Contains("Tower")))
                    {
                        int h = 0;
                        foreach (string hh in headerU1)
                        {
                            dic.variableHeader[h++, 0] = hh;
                        }
                        dic.resultMatrix = this.GetFormatResult(dic.path, headerU1, "UltimateLoads");
                    }

                    else if (dic.name.Contains("Hub"))
                    {
                        int h = 0;
                        foreach (string hh in headerU2)
                        {
                            dic.variableHeader[h++, 0] = hh;
                        }
                        dic.resultMatrix = this.GetFormatResult(dic.path, headerU2, "UltimateLoads");
                    }

                    else
                    {
                        Console.WriteLine("GetMainCompinentLoadsResult error!");
                    }
                    dic.mainDlcData = GetEachDlcMaxValueFromMainCom(dic); //向部件数据中添加热力图数据源                    
                    bladedDataTemp.turbineMainCompenontResult.results.ultmateData.component.Add(GetMatrixMaxValue(dic));
                }
                //疲劳载荷
                foreach (var dic in ComponentF)
                {
                    dic.resultMatrix = this.GetFormatResult(dic.path, headerF, "EquivalentFatigueLoads");
                    bladedDataTemp.turbineMainCompenontResult.results.equivalentFatigueData.component.Add(dic);
                }
                bladedDlist.Add(bladedDataTemp);
            }
            return bladedDlist;
        }

        List<BladedData.TurbineMainCompenontResult.Results.MainDlcDataStruct> GetEachDlcMaxValueFromMainCom(BladedData.TurbineMainCompenontResult.Results.MainComponentDataStruct mainCom)
        {
            List<string> dlcPathList;
            BladedData.TurbineMainCompenontResult.Results.MainDlcDataStruct mainDlcData;
            List<BladedData.TurbineMainCompenontResult.Results.MainDlcDataStruct> mainDlcDataList = new List<BladedData.TurbineMainCompenontResult.Results.MainDlcDataStruct>();
            string pjName = Path.GetFullPath(Directory.GetFiles(mainCom.path, "*.$PJ")[0]).Replace(".$PJ", ".$MX");
            string[] dlcPath = Directory.GetFiles(mainCom.path, "*.$MX");
            dlcPathList = new List<string>();
            string sameDlc = null;
            int endFlag = 0;
            foreach (string ss in dlcPath)
            {
                if (ss.Equals(pjName)) continue;
                else if(pjName.Contains("br.$MX") || pjName.Contains("hr.$MX") || pjName.Contains("hs.$MX"))
                {
                    string[] tempName = Path.GetFileNameWithoutExtension(ss).Split(new char[] { '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
                    if (string.IsNullOrEmpty(sameDlc)) sameDlc = tempName[0];
                    else if (sameDlc == tempName[0]) dlcPathList.Add(ss);
                    else
                    {
                        mainDlcData = new BladedData.TurbineMainCompenontResult.Results.MainDlcDataStruct();
                        mainDlcData.resultMatrix = GetUltimateLoadsDlcMaxValueFromFile(dlcPathList);
                        mainDlcData.dlcName = sameDlc.Replace(Path.GetFileNameWithoutExtension(pjName) + ".", "DLC");  //修改输出工况名称
                        mainDlcDataList.Add(mainDlcData);
                        sameDlc = tempName[0];
                        dlcPathList.Clear();
                        dlcPathList.Add(ss);
                    }
                }
                else //塔架的命名分割与其他不同
                {
                    string[] tempName = Path.GetFileNameWithoutExtension(ss).Split(new char[] { '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
                    string tempDlcName = tempName[0] + "_" + tempName[1];
                    if (string.IsNullOrEmpty(sameDlc)) sameDlc = tempDlcName;
                    else if (sameDlc == tempDlcName) dlcPathList.Add(ss);
                    else
                    {
                        mainDlcData = new BladedData.TurbineMainCompenontResult.Results.MainDlcDataStruct();
                        mainDlcData.resultMatrix = GetUltimateLoadsDlcMaxValueFromFile(dlcPathList);
                        mainDlcData.dlcName = sameDlc.Replace(Path.GetFileNameWithoutExtension(pjName) + ".", "DLC");  //修改输出工况名称
                        mainDlcDataList.Add(mainDlcData);
                        sameDlc = tempDlcName;
                        dlcPathList.Clear();
                        dlcPathList.Add(ss);
                    }
                }
                ++endFlag;
                if (dlcPath.Length == endFlag + 1) //文件遍历完后获取最后一部分数据+1为br.$MX
                {
                    mainDlcData = new BladedData.TurbineMainCompenontResult.Results.MainDlcDataStruct();
                    mainDlcData.resultMatrix = GetUltimateLoadsDlcMaxValueFromFile(dlcPathList);
                    mainDlcData.dlcName = sameDlc.Replace(Path.GetFileNameWithoutExtension(pjName) + ".", "DLC");  //修改输出工况名称
                    mainDlcDataList.Add(mainDlcData);
                }
            }
            return mainDlcDataList;
        }

        public static T Clone<T>(T RealObject)

        {
            using (Stream objectStream = new MemoryStream())
            {
                //利用 System.Runtime.Serialization序列化与反序列化完成引用对象的复制  
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, RealObject);
                objectStream.Seek(0, SeekOrigin.Begin);
                return (T)formatter.Deserialize(objectStream);
            }
        }
    }

    public interface IFilesOperation
    {
        List<Dictionary<string, string>> GetMainComponentPath();
    }

    public class FilesOperation : IFilesOperation
    {
        public FilesOperation()
        {
        }
        public List<Dictionary<string, string>> GetMainComponentPath()
        {
            List<Dictionary<string, string>> mainComponentPath = new List<Dictionary<string, string>>();
            string dirPath = Directory.GetCurrentDirectory();
            List<string> filesName = new List<string>();
            var filesPathTemp = Directory.GetFiles(dirPath, "*.txt");
            filesName.AddRange(filesPathTemp);
            filesName.Sort();

            foreach (string s in filesName) //依次打开文件
            {
                Dictionary<string, string> singleComponentPath = new Dictionary<string, string>();
                using (StreamReader reader = new StreamReader(s)) //依次打开单个文件中的每个部件的路径
                {
                    while (reader.EndOfStream == false)
                    {
                        var linePath = reader.ReadLine().Split(new[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);
                        singleComponentPath.Add(linePath[0], linePath[1]);
                    }
                    reader.Close();
                }
                mainComponentPath.Add(singleComponentPath);
            }
            return mainComponentPath;
        }
    }
}
