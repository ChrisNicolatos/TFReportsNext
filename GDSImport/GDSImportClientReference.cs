using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDSImport
{
    public class GDSImportClientReference
    {
        public int Id { get; set; }
        public string Element { get; set; }
        public string Description { get; set; }
        public string Required { get; set; }
        public bool LimitToLookUp { get; set; }
        public string RawValue { get; set; }
        public string Value { get; set; }
        public string ValueFound { get; set; }
        public string Found { get; set; }

    }
}
