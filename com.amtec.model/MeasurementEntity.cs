using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace com.amtec.model
{
    public class MeasurementEntity
    {
        public string MeasurementName { get; set; }
        public string MaxValue { get; set; }
        public string MinValue { get; set; }
        public string TargetValue { get; set; }
    }
}
