using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using WordUtilityAddIn.Core.Models;

namespace WordUtilityAddIn.Core.Services
{
    public sealed class OoxmlRubyProcessor
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static readonly string[] RubyPropertyOrder =
        {
            "rubyAlign", "hps", "hpsRaise", "hpsBaseText", "lid", "dirty"
        };

        private static readonly string[] RunPropertyOrder =
        {
            "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
            "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
            "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
            "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
            "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"
        };

        public RubyStatistics ScanPackage(string packagePath)
        {
            if (string.IsNullOrWhiteSpace(packagePath))
            {
                throw new ArgumentException("文档路径不能为空。", nameof(packagePath));
            }

            using (Package package = Package.Open(packagePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                List<XmlPartDocument> documents = LoadWordXmlParts(package);
                return ScanDocuments(documents);
            }
        }

        public RubyStatistics ScanXml(string xml)
        {
            XDocument document = ParseXml(xml);
            return ScanDocuments(new[] { new XmlPartDocument("selection", document) });
        }

        public RubyUpdateResult UpdatePackage(string packagePath, RubyFormatOptions options)
        {
            if (string.IsNullOrWhiteSpace(packagePath))
            {
                throw new ArgumentException("文档路径不能为空。", nameof(packagePath));
            }

            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            options.Validate();
            RubyStatistics before;
            int modifiedCount = 0;

            using (Package package = Package.Open(packagePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                List<XmlPartDocument> documents = LoadWordXmlParts(package);
                before = ScanDocuments(documents);
                int? fallbackBaseHalfPoints = before.DominantBaseTextHalfPoints;

                foreach (XmlPartDocument item in documents)
                {
                    int changed = UpdateDocument(item.Document, options, fallbackBaseHalfPoints);
                    if (changed == 0)
                    {
                        continue;
                    }

                    WritePart(item.Part, item.Document);
                    modifiedCount += changed;
                }
            }

            RubyStatistics after = ScanPackage(packagePath);
            ValidateIdentity(before, after);
            ValidatePackageFormatting(packagePath, options, after.DominantBaseTextHalfPoints);

            return new RubyUpdateResult
            {
                ModifiedRubyCount = modifiedCount,
                Before = before,
                After = after
            };
        }

        public string UpdateXml(string xml, RubyFormatOptions options, out RubyUpdateResult result)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            options.Validate();
            XDocument document = ParseXml(xml);
            RubyStatistics before = ScanDocuments(new[] { new XmlPartDocument("selection", document) });
            int modifiedCount = UpdateDocument(document, options, before.DominantBaseTextHalfPoints);
            RubyStatistics after = ScanDocuments(new[] { new XmlPartDocument("selection", document) });

            ValidateIdentity(before, after);
            ValidateDocumentFormatting(document, options, after.DominantBaseTextHalfPoints, "selection");

            result = new RubyUpdateResult
            {
                ModifiedRubyCount = modifiedCount,
                Before = before,
                After = after
            };

            return SerializeXml(document);
        }

        public RubyStatistics ValidateXmlFormatting(string xml, RubyFormatOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            options.Validate();
            XDocument document = ParseXml(xml);
            RubyStatistics statistics = ScanDocuments(
                new[] { new XmlPartDocument("selection", document) });
            ValidateDocumentFormatting(
                document,
                options,
                statistics.DominantBaseTextHalfPoints,
                "selection");
            return statistics;
        }

        private static List<XmlPartDocument> LoadWordXmlParts(Package package)
        {
            var documents = new List<XmlPartDocument>();

            foreach (PackagePart part in package.GetParts()
                .Where(IsWordXmlPart)
                .OrderBy(item => item.Uri.OriginalString, StringComparer.OrdinalIgnoreCase))
            {
                using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read))
                {
                    var settings = new XmlReaderSettings
                    {
                        DtdProcessing = DtdProcessing.Prohibit,
                        XmlResolver = null
                    };

                    using (XmlReader reader = XmlReader.Create(stream, settings))
                    {
                        XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
                        documents.Add(new XmlPartDocument(part.Uri.OriginalString, document, part));
                    }
                }
            }

            return documents;
        }

        private static bool IsWordXmlPart(PackagePart part)
        {
            string uri = part.Uri.OriginalString;
            string contentType = part.ContentType ?? string.Empty;
            return uri.StartsWith("/word/", StringComparison.OrdinalIgnoreCase)
                && contentType.IndexOf("xml", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static XDocument ParseXml(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml))
            {
                throw new ArgumentException("Word Open XML 内容为空。", nameof(xml));
            }

            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            };

            using (var stringReader = new StringReader(xml))
            using (XmlReader reader = XmlReader.Create(stringReader, settings))
            {
                return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            }
        }

        private static RubyStatistics ScanDocuments(IEnumerable<XmlPartDocument> documents)
        {
            var pronunciationSizes = new List<int>();
            var baseTextSizes = new List<int>();
            var raises = new List<int>();
            var partCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var fingerprintSource = new StringBuilder();
            int total = 0;

            foreach (XmlPartDocument item in documents)
            {
                List<XElement> rubies = item.Document.Descendants(W + "ruby").ToList();
                if (rubies.Count > 0)
                {
                    partCounts[item.Name] = rubies.Count;
                }

                fingerprintSource.Append(item.Name.Length).Append(':').Append(item.Name).Append('|');
                foreach (XElement ruby in rubies)
                {
                    total++;
                    AddIfPresent(pronunciationSizes, ReadPronunciationHalfPoints(ruby));
                    AddIfPresent(baseTextSizes, ReadBaseTextHalfPoints(ruby));
                    AddIfPresent(raises, ReadRubyProperty(ruby, "hpsRaise"));

                    string reading = string.Concat(
                        ruby.Element(W + "rt")?
                            .Descendants(W + "t")
                            .Select(element => element.Value)
                        ?? Enumerable.Empty<string>());

                    fingerprintSource.Append(reading.Length).Append(':').Append(reading).Append('|');
                }
            }

            return new RubyStatistics
            {
                RubyCount = total,
                DominantPronunciationHalfPoints = GetDominant(pronunciationSizes),
                DominantBaseTextHalfPoints = GetDominant(baseTextSizes),
                DominantRaiseHalfPoints = GetDominant(raises),
                PartRubyCounts = partCounts,
                ReadingFingerprint = ComputeSha256(fingerprintSource.ToString())
            };
        }

        private static int UpdateDocument(
            XDocument document,
            RubyFormatOptions options,
            int? fallbackBaseHalfPoints)
        {
            int changed = 0;

            foreach (XElement ruby in document.Descendants(W + "ruby").ToList())
            {
                XElement rubyProperties = ruby.Element(W + "rubyPr");
                if (rubyProperties == null)
                {
                    rubyProperties = new XElement(W + "rubyPr");
                    ruby.AddFirst(rubyProperties);
                }

                int? baseHalfPoints = ReadBaseTextHalfPoints(ruby) ?? fallbackBaseHalfPoints;
                int raiseHalfPoints;
                if (options.AutoRaise)
                {
                    if (!baseHalfPoints.HasValue)
                    {
                        throw new InvalidDataException("存在无法确定正文基准字号的拼音标注，不能计算自动上移参数。");
                    }

                    SetOrderedVal(rubyProperties, "hpsBaseText", baseHalfPoints.Value, RubyPropertyOrder);
                    raiseHalfPoints = checked(baseHalfPoints.Value + 2);
                }
                else
                {
                    raiseHalfPoints = options.ManualRaiseHalfPoints;
                }

                SetOrderedVal(rubyProperties, "hps", options.PronunciationHalfPoints, RubyPropertyOrder);
                SetOrderedVal(rubyProperties, "hpsRaise", raiseHalfPoints, RubyPropertyOrder);

                XElement rubyText = ruby.Element(W + "rt");
                List<XElement> runs = rubyText == null
                    ? new List<XElement>()
                    : rubyText.Descendants(W + "r").ToList();

                if (runs.Count == 0)
                {
                    throw new InvalidDataException("发现不含 w:rt/w:r 的拼音标注，已停止处理。");
                }

                foreach (XElement run in runs)
                {
                    XElement runProperties = run.Element(W + "rPr");
                    if (runProperties == null)
                    {
                        runProperties = new XElement(W + "rPr");
                        run.AddFirst(runProperties);
                    }

                    SetOrderedVal(runProperties, "sz", options.PronunciationHalfPoints, RunPropertyOrder);
                    SetOrderedVal(runProperties, "szCs", options.PronunciationHalfPoints, RunPropertyOrder);

                    if (!string.IsNullOrWhiteSpace(options.PronunciationFont))
                    {
                        XElement fonts = EnsureOrderedElement(runProperties, "rFonts", RunPropertyOrder);
                        string font = options.PronunciationFont.Trim();
                        fonts.SetAttributeValue(W + "ascii", font);
                        fonts.SetAttributeValue(W + "hAnsi", font);
                        fonts.SetAttributeValue(W + "eastAsia", font);
                        fonts.SetAttributeValue(W + "cs", font);
                    }
                }

                changed++;
            }

            return changed;
        }

        private static void ValidatePackageFormatting(
            string packagePath,
            RubyFormatOptions options,
            int? fallbackBaseHalfPoints)
        {
            using (Package package = Package.Open(packagePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                foreach (XmlPartDocument item in LoadWordXmlParts(package))
                {
                    ValidateDocumentFormatting(item.Document, options, fallbackBaseHalfPoints, item.Name);
                }
            }
        }

        private static void ValidateDocumentFormatting(
            XDocument document,
            RubyFormatOptions options,
            int? fallbackBaseHalfPoints,
            string sourceName)
        {
            foreach (XElement ruby in document.Descendants(W + "ruby"))
            {
                int? hps = ReadRubyProperty(ruby, "hps");
                if (hps != options.PronunciationHalfPoints)
                {
                    throw new InvalidDataException(sourceName + " 中的 w:rubyPr/w:hps 校验失败。");
                }

                int? baseHalfPoints = ReadBaseTextHalfPoints(ruby) ?? fallbackBaseHalfPoints;
                int expectedRaise = options.AutoRaise
                    ? checked((baseHalfPoints ?? throw new InvalidDataException("校验时无法确定正文基准字号。")) + 2)
                    : options.ManualRaiseHalfPoints;

                if (ReadRubyProperty(ruby, "hpsRaise") != expectedRaise)
                {
                    throw new InvalidDataException(sourceName + " 中的 w:rubyPr/w:hpsRaise 校验失败。");
                }

                XElement rubyText = ruby.Element(W + "rt");
                IEnumerable<XElement> runs = rubyText == null
                    ? Enumerable.Empty<XElement>()
                    : rubyText.Descendants(W + "r");

                foreach (XElement run in runs)
                {
                    XElement runProperties = run.Element(W + "rPr");
                    int? runSize = ReadVal(runProperties?.Element(W + "sz"));
                    if (runSize != options.PronunciationHalfPoints)
                    {
                        throw new InvalidDataException(sourceName + " 中的 w:rt/w:rPr/w:sz 校验失败。");
                    }

                    if (!string.IsNullOrWhiteSpace(options.PronunciationFont))
                    {
                        string expectedFont = options.PronunciationFont.Trim();
                        XElement fonts = runProperties?.Element(W + "rFonts");
                        if ((string)fonts?.Attribute(W + "ascii") != expectedFont
                            || (string)fonts?.Attribute(W + "hAnsi") != expectedFont)
                        {
                            throw new InvalidDataException(sourceName + " 中的拼音字体校验失败。");
                        }
                    }
                }
            }
        }

        private static void ValidateIdentity(RubyStatistics before, RubyStatistics after)
        {
            if (before.RubyCount != after.RubyCount)
            {
                throw new InvalidDataException(
                    string.Format("注音数量校验失败：修改前 {0}，修改后 {1}。", before.RubyCount, after.RubyCount));
            }

            if (!string.Equals(before.ReadingFingerprint, after.ReadingFingerprint, StringComparison.Ordinal))
            {
                throw new InvalidDataException("拼音读音校验失败：修改前后的读音内容不一致。");
            }
        }

        private static int? ReadPronunciationHalfPoints(XElement ruby)
        {
            int? value = ReadRubyProperty(ruby, "hps");
            if (value.HasValue)
            {
                return value;
            }

            return ReadVal(
                ruby.Element(W + "rt")?
                    .Descendants(W + "rPr")
                    .Elements(W + "sz")
                    .FirstOrDefault());
        }

        private static int? ReadBaseTextHalfPoints(XElement ruby)
        {
            int? value = ReadRubyProperty(ruby, "hpsBaseText");
            if (value.HasValue)
            {
                return value;
            }

            return ReadVal(
                ruby.Element(W + "rubyBase")?
                    .Descendants(W + "rPr")
                    .Elements(W + "sz")
                    .FirstOrDefault());
        }

        private static int? ReadRubyProperty(XElement ruby, string localName)
        {
            return ReadVal(ruby.Element(W + "rubyPr")?.Element(W + localName));
        }

        private static int? ReadVal(XElement element)
        {
            int value;
            return element != null
                && int.TryParse((string)element.Attribute(W + "val"), out value)
                ? value
                : (int?)null;
        }

        private static void SetOrderedVal(
            XElement parent,
            string localName,
            int value,
            IReadOnlyList<string> order)
        {
            XElement element = EnsureOrderedElement(parent, localName, order);
            element.SetAttributeValue(W + "val", value);
        }

        private static XElement EnsureOrderedElement(
            XElement parent,
            string localName,
            IReadOnlyList<string> order)
        {
            XElement existing = parent.Element(W + localName);
            if (existing != null)
            {
                return existing;
            }

            var created = new XElement(W + localName);
            int newIndex = IndexOf(order, localName);
            XElement insertBefore = parent.Elements()
                .FirstOrDefault(child =>
                {
                    int childIndex = IndexOf(order, child.Name.LocalName);
                    return childIndex >= 0 && childIndex > newIndex;
                });

            if (insertBefore == null)
            {
                parent.Add(created);
            }
            else
            {
                insertBefore.AddBeforeSelf(created);
            }

            return created;
        }

        private static int IndexOf(IReadOnlyList<string> values, string value)
        {
            for (int index = 0; index < values.Count; index++)
            {
                if (string.Equals(values[index], value, StringComparison.Ordinal))
                {
                    return index;
                }
            }

            return -1;
        }

        private static void WritePart(PackagePart part, XDocument document)
        {
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write))
            {
                var settings = new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false),
                    Indent = false,
                    OmitXmlDeclaration = false,
                    CloseOutput = false
                };

                using (XmlWriter writer = XmlWriter.Create(stream, settings))
                {
                    document.Save(writer);
                }
            }
        }

        private static string SerializeXml(XDocument document)
        {
            var builder = new StringBuilder();
            var settings = new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = false,
                OmitXmlDeclaration = false
            };

            using (XmlWriter writer = XmlWriter.Create(builder, settings))
            {
                document.Save(writer);
            }

            return builder.ToString();
        }

        private static int? GetDominant(IEnumerable<int> values)
        {
            var groups = values
                .GroupBy(value => value)
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key)
                .FirstOrDefault();
            return groups == null ? (int?)null : groups.Key;
        }

        private static void AddIfPresent(ICollection<int> values, int? value)
        {
            if (value.HasValue)
            {
                values.Add(value.Value);
            }
        }

        private static string ComputeSha256(string value)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value));
                return string.Concat(bytes.Select(item => item.ToString("x2")));
            }
        }

        private sealed class XmlPartDocument
        {
            public XmlPartDocument(string name, XDocument document, PackagePart part = null)
            {
                Name = name;
                Document = document;
                Part = part;
            }

            public string Name { get; private set; }

            public XDocument Document { get; private set; }

            public PackagePart Part { get; private set; }
        }
    }
}
