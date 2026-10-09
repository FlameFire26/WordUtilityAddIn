using System;

namespace WordUtilityAddIn.Core.Models
{
    public sealed class RubyFormatOptions
    {
        public int PronunciationHalfPoints { get; set; }

        public string PronunciationFont { get; set; }

        public bool AutoRaise { get; set; }

        public int ManualRaiseHalfPoints { get; set; }

        public void Validate()
        {
            if (PronunciationHalfPoints < 2 || PronunciationHalfPoints > 200)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(PronunciationHalfPoints),
                    "拼音字号必须在 1 到 100 磅之间。");
            }

            if (!AutoRaise && (ManualRaiseHalfPoints < 0 || ManualRaiseHalfPoints > 400))
            {
                throw new ArgumentOutOfRangeException(
                    nameof(ManualRaiseHalfPoints),
                    "上移参数必须在 0 到 400 个半磅单位之间。");
            }
        }
    }
}

