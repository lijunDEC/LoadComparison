using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoadComparison
{
    public class BladedData
    {
        public EquivalentFatigueLoads equivalentFatigueLoads;
        public UltimateLoads ultimateLoads;
        public TurbineMainCompenontResult turbineMainCompenontResult;
        public BladedData()
        {
            equivalentFatigueLoads = new EquivalentFatigueLoads();
            ultimateLoads = new UltimateLoads();
            turbineMainCompenontResult = new TurbineMainCompenontResult();
        }
        #region EquivalentFatigueLoads
        public class EquivalentFatigueLoads
        {
            public string path;

            public class EquivalentFatigueLoadsVariable
            {
                public Dictionary<int, float> variable;
                public EquivalentFatigueLoadsVariable()
                {
                    variable = new Dictionary<int, float>();
                }
            }
            public EquivalentFatigueLoadsVariable Mx, My, Mz, Fx, Fy, Fz;
            //            public List<Dictionary<int, float>> variable;
            public string[,] arrayResults;
 
            public EquivalentFatigueLoads()
            {
                Mx = new EquivalentFatigueLoadsVariable();
                My = new EquivalentFatigueLoadsVariable();
                Mz = new EquivalentFatigueLoadsVariable();
                Fx = new EquivalentFatigueLoadsVariable();
                Fy = new EquivalentFatigueLoadsVariable();
                Fz = new EquivalentFatigueLoadsVariable();
                arrayResults = new string[10, 7];
            }
        }
        #endregion EquivalentFatigueLoads

        #region UltimateLoads
        public class UltimateLoads
        {
            public UltimateLoadsvariable Mx, My, Mxy, Mz, Myz, Fx, Fy, Fxy, Fz, Fyz;
            public string[,] arrayResults;
            public string path;
            public UltimateLoads()
            {
                Mx = new UltimateLoadsvariable();
                My = new UltimateLoadsvariable();
                Mxy = new UltimateLoadsvariable();
                Mz = new UltimateLoadsvariable();
                Myz = new UltimateLoadsvariable();
                Fx = new UltimateLoadsvariable();
                Fy = new UltimateLoadsvariable();
                Fxy = new UltimateLoadsvariable();
                Fz = new UltimateLoadsvariable();
                Fyz = new UltimateLoadsvariable();
                arrayResults = new string[18, 12];
            }

            public class maxMinValue
            {
                public string loadCase;
                public float Mx, My, Mxy, Mz, Myz, Fx, Fy, Fxy, Fz, Fyz;
            }

            public class UltimateLoadsvariable : maxMinValue
            {
                public maxMinValue maxValue;
                public maxMinValue minValue;
                public UltimateLoadsvariable()
                {
                    maxValue = new maxMinValue();
                    minValue = new maxMinValue();
                }
            }
        }
        #endregion UltimateLoads

        public class TurbineMainCompenontResult
        {
            public string turbineID;
            public Results results = new Results();
            
            public class Results
            {
                public UltmateLoad ultmateData;
                public EquivalentFatigueLoad equivalentFatigueData;

                public Results()
                {
                    ultmateData = new UltmateLoad();
                    equivalentFatigueData = new EquivalentFatigueLoad();
                }

                public class UltmateLoad
                {

                    public List<MainComponentDataStruct> component = new List<MainComponentDataStruct>(); 

                    public UltmateLoad()
                    {
                        component = new List<MainComponentDataStruct>();
                    }
                }

                public class EquivalentFatigueLoad : UltmateLoad
                {
                }

                public class MainComponentDataStruct
                {
                    public string name;
                    public string[,] variableHeader;//Mx, My, ...
                    public string path;
                    public string[,] resultMatrix;
                    public List<string> dlcNameList;
                    public float[,] dlcMaxValueList;
                    public List<MainDlcDataStruct> mainDlcData;
                }
                public class MainDlcDataStruct

                {
                    public string dlcName;

                    public float[,] resultMatrix; 
                }
            }
        }
    }
}
