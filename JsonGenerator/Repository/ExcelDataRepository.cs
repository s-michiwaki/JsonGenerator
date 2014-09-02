using JsonGenerator.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonGenerator.Repository
{
    public class ExcelDataRepository
    {
        public static readonly Dictionary<string, Type> ExcelDataTypes;

        static ExcelDataRepository()
        {
            ExcelDataTypes = GetExcelDataModelList().ToDictionary(x => x.Name, x => x);
        }

        public static IReadOnlyCollection<Type> GetExcelDataModelList()
        {
            System.Reflection.Assembly[] assemblies;
            assemblies = System.AppDomain.CurrentDomain.GetAssemblies();
            List<Type> excelDataModels = new List<Type>();
            foreach (System.Reflection.Assembly assembly in assemblies)
            {
                foreach (var type in assembly.GetTypes())
                {
                    if (type.GetCustomAttributes(typeof(ExcelDataAttribute), false).Length > 0)
                    {
                        excelDataModels.Add(type);
                    }
                }
            }
            return excelDataModels;
        }
    }
}
