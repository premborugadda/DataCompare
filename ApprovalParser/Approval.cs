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
		public string DimOperationalMeasureKey { get; set; }
		public string RevisedVoulmeDry { get; set; }
		public string PrelinVolume   {get; set; }
		public string PrelinDrynessFactor { get; set; }
		public string RevisedVoulmeWet { get; set; }
		public string RevisedDrynessFactor { get; set; }
		public string CrushedOreAdjusted { get; set; }
		public string Comments { get; set; }
		public string CreatedDate { get; set; }
		public string CreatedBy { get; set; }
		public string DimDateKey { get; set; }
		public string UpdatedDate { get; set; }
		public string UpdatedBy { get; set; }
		public string CreatedByEmailId { get; set; }
		public string UpdatedByEmailId { get; set; }
		public string ApprovedDimOperationalMeasureKey { get; set; }

	}
}

