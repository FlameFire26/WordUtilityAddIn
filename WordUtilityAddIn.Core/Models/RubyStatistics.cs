using System.Collections.Generic;

namespace WordUtilityAddIn.Core.Models
{
    public sealed class RubyStatistics
    {
        public RubyStatistics()
        {
            PartRubyCounts = new Dictionary<string, int>();
        }

        public int RubyCount { get; internal set; }

        public int? DominantPronunciationHalfPoints { get; internal set; }

        public int? DominantBaseTextHalfPoints { get; internal set; }

        public int? DominantRaiseHalfPoints { get; internal set; }

        public IReadOnlyDictionary<string, int> PartRubyCounts { get; internal set; }

        public string ReadingFingerprint { get; internal set; }
    }
}
