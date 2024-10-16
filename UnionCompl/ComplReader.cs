using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnionCompl
{
    internal class ComplReader
    {
        /// <summary>
        /// Словарь с элементами комплектации
        /// </summary>
        public Dictionary<string, Dictionary<string, int>> Elements;

        public string file_name;

        public ComplReader()
        {
            Elements = new Dictionary<string, Dictionary<string, int>>();
            file_name = "";
        }

        public ComplReader(string file_name)
        {
            Elements = new Dictionary<string, Dictionary<string, int>>();
            this.file_name = file_name;
        }

    }
}
