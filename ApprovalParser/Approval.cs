using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCompare.ApprovalParser
{
    public class Approval
    {
		public DateTime CalendarDate { get; set; }
		public Double DimOperationalMeasureKey { get; set; }
		public Double RevisedVoulmeDry { get; set; }
		public Double PrelinVolume   {get; set; }
		public Double PrelinDrynessFactor { get; set; }
		public Double RevisedVoulmeWet { get; set; }
		public Double RevisedDrynessFactor { get; set; }
		public Double CrushedOreAdjusted { get; set; }
		public string Comments { get; set; }
		public DateTime CreatedDate { get; set; }
		public string CreatedBy { get; set; }
		public Double DimDateKey { get; set; }
		public DateTime UpdatedDate { get; set; }
		public string UpdatedBy { get; set; }
		public string CreatedByEmailId { get; set; }
		public string UpdatedByEmailId { get; set; }
		public Double ApprovedDimOperationalMeasureKey { get; set; }

	}
}

