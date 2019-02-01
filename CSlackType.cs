using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace excelParcer
{
    class CSlackType
    {
        

        //Тип рычага
        public string SlactType { get; set; }
        //Торговая марка
        public string TrandMark { get; set; }
        
        //Список с данными для распознования
        private Dictionary<string, string[,]> Types_TrandMarks = new Dictionary<string, string[,]>();
        
        //Конструктор класса
        public CSlackType()
        {
            this.SlactType = "";
            this.initTypes_TrandMarks();
        }

        //Перегруженный конструктор класса
        public CSlackType(string slackId)
        {
            
            this.SlactType = "";
            this.initTypes_TrandMarks();
            this.GetDataFromId(slackId);
            
        }

        private void initTypes_TrandMarks()
        {
            Types_TrandMarks.Add("4W", new[,] { { "ASA" }, { "MEI" } });
            Types_TrandMarks.Add("6Y", new[,] { { "S-ASA" }, { "MEI" } });
            Types_TrandMarks.Add("4Y", new[,] { { "MSA" }, { "MEI" } });
            Types_TrandMarks.Add("QA", new[,] { { "ASA" }, { "QAS" } });
            Types_TrandMarks.Add("AS", new[,] { { "ASA" }, { "QAS" } });
            Types_TrandMarks.Add("SS", new[,] { { "S-ASA" }, { "QAS" } });
            Types_TrandMarks.Add("21", new[,] { { "MSA" }, { "DeWe" } });
        }
        public void GetDataFromId(string id)
        {
            this.SlactType = Types_TrandMarks[id.Substring(0, 2)].GetValue(0,0).ToString();
            this.TrandMark = Types_TrandMarks[id.Substring(0, 2)].GetValue(1,0).ToString();
        }

    }
}
