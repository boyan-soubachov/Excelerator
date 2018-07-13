/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Runtime.Serialization;

namespace ExcelIdea
{
    [Serializable]
    class PipeMessage
    {
        public List<Dictionary<string, string>> dict_loader_formulas;
        public List<Dictionary<string, float>> dict_loader_constants;
        public List<Dictionary<string, string>> dict_loader_string_constants;
        public List<Dictionary<string, Tools.ArrayFormulaInfo>> dict_loader_array_formulas;
        public Dictionary<string, int> dict_sheet_name_map;

        public PipeMessage()
        {
        }
    }

    sealed class MessageAssemblyBinder : SerializationBinder
    {
        public override Type BindToType(string assemblyName, string typeName)
        {
            Type typeToDeserialize = null;

            String currentAssembly = Assembly.GetExecutingAssembly().FullName;

            // In this case we are always using the current assembly
            assemblyName = currentAssembly;

            // Get the type using the typeName and assemblyName
            typeToDeserialize = Type.GetType(String.Format("{0}, {1}",
                typeName, assemblyName));

            return typeToDeserialize;
        }
    }
}
