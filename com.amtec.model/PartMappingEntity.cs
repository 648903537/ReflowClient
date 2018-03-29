using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace com.amtec.model
{
    public class PartMappingEntity
    {
        public string PartNo { get; set; }

        public string PartDesc { get; set; }

        public string MeasurementType { get; set; }

        public string MeasurementUnit { get; set; }

        public string MinValue { get; set; }

        public string MaxValue { get; set; }

        public string TestCount { get; set; }
    }
}
