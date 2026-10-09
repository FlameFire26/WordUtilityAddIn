namespace WordUtilityAddIn.Core.Models
{
    public sealed class RubyUpdateResult
    {
        public int ModifiedRubyCount { get; set; }

        public RubyStatistics Before { get; set; }

        public RubyStatistics After { get; set; }

        public string BackupPath { get; set; }
    }
}
