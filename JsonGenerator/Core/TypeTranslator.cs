using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonGenerator.Core
{
    public class TypeTranslator
    {
        //public static object Sample(string name)
        //{
        //    // 型名から型を取得。  
        //    var type = GetTypeFromName(name);

        //    // Typeからインスタンス作成。  
        //    var instance = Activator.CreateInstance(type);

        //    return instance;
        //}

        /// <summary>  
        /// 特定の名前を持つ型を取得します。  
        /// </summary>  
        /// <param name="name">型名</param>  
        /// <returns>取得した型</returns>  
        public static Type GetTypeFromName(String name)
        {
            return GetTypeFromName(name, GetCurrentDomainAssembliesType());
        }

        /// <summary>  
        /// 特定の名前を持つ型を取得します。  
        /// </summary>  
        /// <param name="name">型名</param>  
        /// <param name="types">型のリスト</param>  
        /// <returns>取得した型</returns>  
        public static Type GetTypeFromName(String name, IEnumerable<Type> types)
        {
            foreach (Type type in types)
            {
                if (String.Equals(type.FullName, name) || String.Equals(type.AssemblyQualifiedName, name))
                    return type;
            }
            return null;
        }

        /// <summary>  
        /// CurrentDomainのアセンブリの型一覧を取得します。  
        /// </summary>  
        /// <returns>型一覧</returns>  
        public static IList<Type> GetCurrentDomainAssembliesType()
        {
            System.Reflection.Assembly[] assemblies;
            assemblies = System.AppDomain.CurrentDomain.GetAssemblies();
            List<Type> assemblyTypes = new List<Type>();
            foreach (System.Reflection.Assembly assembly in assemblies)
            {
                assemblyTypes.AddRange(assembly.GetTypes());
            }
            return assemblyTypes;
        }  
    }
}
