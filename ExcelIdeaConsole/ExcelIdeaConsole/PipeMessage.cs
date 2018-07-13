/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.Reflection;

namespace ExcelIdea
{
    [Serializable]
    class PipeMessage
    {
        public List<Dictionary<string, string>> dict_loader_formulas = null;
        public List<Dictionary<string, float>> dict_loader_constants = null;
        public List<Dictionary<string, string>> dict_loader_string_constants = null;
        public List<Dictionary<string, Tools.ArrayFormulaInfo>> dict_loader_array_formulas = null;
        public Dictionary<string, int> dict_sheet_name_map = null;

        public PipeMessage()
        {
        }
    }

    //sealed class MessageAssemblyBinder : SerializationBinder
    //{
    //    public override Type BindToType(string assemblyName, string typeName)
    //    {
    //        Type typeToDeserialize = null;
    //        String currentAssembly = Assembly.GetExecutingAssembly().FullName;

    //        assemblyName = currentAssembly;
    //        typeName = typeName.Replace("ExcelIdea", "ExcelIdeaConsole");
    //        typeToDeserialize = Type.GetType(String.Format("{0}, {1}",
    //            typeName, assemblyName));

    //        return typeToDeserialize;
    //    }
    //}
}
