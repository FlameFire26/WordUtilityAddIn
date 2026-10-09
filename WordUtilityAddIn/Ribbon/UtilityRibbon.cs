using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using WordUtilityAddIn.Core.Models;
using WordUtilityAddIn.Services;

namespace WordUtilityAddIn.Ribbon
{
    [ComVisible(true)]
    public sealed class UtilityRibbon : Office.IRibbonExtensibility
    {
        private const string RibbonXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""tabWordUtility"" label=""实用工具"">
        <group id=""groupRubyGuide"" label=""拼音指南"">
          <dropDown id=""scopeDropDown"" label=""处理范围"" getSelectedItemIndex=""GetScopeIndex"" onAction=""OnScopeChanged"">
            <item id=""scopeDocument"" label=""全文"" />
            <item id=""scopeSelection"" label=""当前选区"" />
          </dropDown>
          <editBox id=""pronunciationSize"" label=""拼音字号（磅）"" sizeString=""100.0"" getText=""GetPronunciationSize"" onChange=""OnPronunciationSizeChanged"" />
          <editBox id=""pronunciationFont"" label=""拼音字体（留空保持）"" sizeString=""Microsoft YaHei"" getText=""GetPronunciationFont"" onChange=""OnPronunciationFontChanged"" />
          <checkBox id=""autoRaise"" label=""自动匹配上移距离"" getPressed=""GetAutoRaise"" onAction=""OnAutoRaiseChanged"" />
          <editBox id=""manualRaise"" label=""上移参数（半磅）"" sizeString=""000"" getText=""GetManualRaise"" getEnabled=""GetManualRaiseEnabled"" onChange=""OnManualRaiseChanged"" />
          <button id=""buttonScan"" label=""扫描"" size=""large"" imageMso=""FindDialog"" onAction=""OnScan"" getEnabled=""GetDocumentAvailable"" />
          <button id=""buttonApply"" label=""应用"" size=""large"" imageMso=""ApplyStyles"" onAction=""OnApply"" getEnabled=""GetApplyEnabled"" />
          <labelControl id=""statusCount"" getLabel=""GetCountLabel"" />
          <labelControl id=""statusPronunciationSize"" getLabel=""GetCurrentSizeLabel"" />
          <labelControl id=""statusBaseSize"" getLabel=""GetBaseSizeLabel"" />
          <labelControl id=""statusRaise"" getLabel=""GetRaiseLabel"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        private readonly WordDocumentService documentService;
        private Office.IRibbonUI ribbonUi;
        private RubyScope scope = RubyScope.FullDocument;
        private string pronunciationSizeText = "9";
        private string pronunciationFont = string.Empty;
        private bool autoRaise = true;
        private string manualRaiseText = "58";
        private RubyStatistics statistics;

        internal UtilityRibbon(WordDocumentService documentService)
        {
            this.documentService = documentService
                ?? throw new ArgumentNullException(nameof(documentService));
        }

        public string GetCustomUI(string ribbonId)
        {
            return RibbonXml;
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbon)
        {
            ribbonUi = ribbon;
        }

        public int GetScopeIndex(Office.IRibbonControl control)
        {
            return scope == RubyScope.FullDocument ? 0 : 1;
        }

        public void OnScopeChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            scope = selectedIndex == 1 ? RubyScope.Selection : RubyScope.FullDocument;
            statistics = null;
            InvalidateAll();
        }

        public string GetPronunciationSize(Office.IRibbonControl control)
        {
            return pronunciationSizeText;
        }

        public void OnPronunciationSizeChanged(Office.IRibbonControl control, string text)
        {
            try
            {
                ParsePointSize(text);
                pronunciationSizeText = text.Trim();
            }
            catch (Exception exception)
            {
                ShowError(exception.Message);
                InvalidateControl("pronunciationSize");
            }
        }

        public string GetPronunciationFont(Office.IRibbonControl control)
        {
            return pronunciationFont;
        }

        public void OnPronunciationFontChanged(Office.IRibbonControl control, string text)
        {
            pronunciationFont = (text ?? string.Empty).Trim();
        }

        public bool GetAutoRaise(Office.IRibbonControl control)
        {
            return autoRaise;
        }

        public void OnAutoRaiseChanged(Office.IRibbonControl control, bool pressed)
        {
            autoRaise = pressed;
            if (autoRaise && statistics?.DominantBaseTextHalfPoints != null)
            {
                manualRaiseText = (statistics.DominantBaseTextHalfPoints.Value + 2)
                    .ToString(CultureInfo.InvariantCulture);
            }

            InvalidateControl("manualRaise");
            InvalidateControl("statusRaise");
        }

        public string GetManualRaise(Office.IRibbonControl control)
        {
            return manualRaiseText;
        }

        public bool GetManualRaiseEnabled(Office.IRibbonControl control)
        {
            return !autoRaise;
        }

        public void OnManualRaiseChanged(Office.IRibbonControl control, string text)
        {
            int value;
            if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
                || value < 0
                || value > 400)
            {
                ShowError("上移参数必须是 0 到 400 之间的整数半磅单位。");
                InvalidateControl("manualRaise");
                return;
            }

            manualRaiseText = value.ToString(CultureInfo.InvariantCulture);
        }

        public bool GetDocumentAvailable(Office.IRibbonControl control)
        {
            try
            {
                return Globals.ThisAddIn?.Application?.Documents.Count > 0;
            }
            catch (COMException)
            {
                return false;
            }
        }

        public bool GetApplyEnabled(Office.IRibbonControl control)
        {
            if (!GetDocumentAvailable(control))
            {
                return false;
            }

            if (scope == RubyScope.FullDocument)
            {
                return true;
            }

            try
            {
                return Globals.ThisAddIn.Application.Selection != null
                    && Globals.ThisAddIn.Application.Selection.Range.Start
                        != Globals.ThisAddIn.Application.Selection.Range.End;
            }
            catch (COMException)
            {
                return false;
            }
        }

        public void OnScan(Office.IRibbonControl control)
        {
            try
            {
                statistics = documentService.Scan(scope);
                if (autoRaise && statistics.DominantBaseTextHalfPoints.HasValue)
                {
                    manualRaiseText = (statistics.DominantBaseTextHalfPoints.Value + 2)
                        .ToString(CultureInfo.InvariantCulture);
                }

                InvalidateAll();
            }
            catch (Exception exception)
            {
                ShowException("扫描失败", exception);
            }
        }

        public void OnApply(Office.IRibbonControl control)
        {
            try
            {
                RubyFormatOptions options = CreateOptions();
                RubyUpdateResult result = documentService.Apply(scope, options);
                statistics = result.After;
                InvalidateAll();

                string backupLine = string.IsNullOrWhiteSpace(result.BackupPath)
                    ? string.Empty
                    : Environment.NewLine + "备份：" + result.BackupPath;
                MessageBox.Show(
                    string.Format(
                        "已修改 {0} 个拼音标注。注音数量和读音校验通过。{1}",
                        result.ModifiedRubyCount,
                        backupLine),
                    "Word 实用工具",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                ShowException("应用失败", exception);
            }
        }

        public string GetCountLabel(Office.IRibbonControl control)
        {
            return statistics == null
                ? "注音数量：请先扫描"
                : "注音数量：" + statistics.RubyCount.ToString(CultureInfo.InvariantCulture);
        }

        public string GetCurrentSizeLabel(Office.IRibbonControl control)
        {
            return "当前拼音字号：" + FormatHalfPoints(statistics?.DominantPronunciationHalfPoints);
        }

        public string GetBaseSizeLabel(Office.IRibbonControl control)
        {
            return "正文基准字号：" + FormatHalfPoints(statistics?.DominantBaseTextHalfPoints);
        }

        public string GetRaiseLabel(Office.IRibbonControl control)
        {
            if (autoRaise && statistics?.DominantBaseTextHalfPoints != null)
            {
                int value = statistics.DominantBaseTextHalfPoints.Value + 2;
                return "上移参数：自动 " + value.ToString(CultureInfo.InvariantCulture) + "（半磅）";
            }

            if (statistics?.DominantRaiseHalfPoints != null)
            {
                return "当前上移参数：" + statistics.DominantRaiseHalfPoints.Value
                    .ToString(CultureInfo.InvariantCulture) + "（半磅）";
            }

            return autoRaise ? "上移参数：等待扫描" : "上移参数：手动 " + manualRaiseText + "（半磅）";
        }

        internal void NotifyDocumentChanged()
        {
            statistics = null;
            InvalidateAll();
        }

        internal void NotifySelectionChanged()
        {
            InvalidateControl("buttonApply");
        }

        private RubyFormatOptions CreateOptions()
        {
            int manualRaise;
            if (!int.TryParse(
                manualRaiseText,
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out manualRaise))
            {
                throw new InvalidOperationException("手动上移参数无效。");
            }

            return new RubyFormatOptions
            {
                PronunciationHalfPoints = ParsePointSize(pronunciationSizeText),
                PronunciationFont = pronunciationFont,
                AutoRaise = autoRaise,
                ManualRaiseHalfPoints = manualRaise
            };
        }

        private static int ParsePointSize(string text)
        {
            double points;
            bool parsed = double.TryParse(text, NumberStyles.Number, CultureInfo.CurrentCulture, out points)
                || double.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out points);
            if (!parsed || points < 1d || points > 100d)
            {
                throw new InvalidOperationException("拼音字号必须是 1 到 100 磅之间的数字。");
            }

            double halfPoints = points * 2d;
            int rounded = (int)Math.Round(halfPoints, MidpointRounding.AwayFromZero);
            if (Math.Abs(halfPoints - rounded) > 0.0001d)
            {
                throw new InvalidOperationException("拼音字号只能使用 0.5 磅的增量。");
            }

            return rounded;
        }

        private static string FormatHalfPoints(int? halfPoints)
        {
            if (!halfPoints.HasValue)
            {
                return "未知";
            }

            return string.Format(
                CultureInfo.CurrentCulture,
                "{0:0.#} 磅（OOXML {1}）",
                halfPoints.Value / 2d,
                halfPoints.Value);
        }

        private void InvalidateAll()
        {
            ribbonUi?.Invalidate();
        }

        private void InvalidateControl(string controlId)
        {
            ribbonUi?.InvalidateControl(controlId);
        }

        private static void ShowException(string title, Exception exception)
        {
            ShowError(title + "：" + Environment.NewLine + GetUsefulMessage(exception));
        }

        private static string GetUsefulMessage(Exception exception)
        {
            Exception current = exception;
            while (current.InnerException != null)
            {
                current = current.InnerException;
            }

            return current.Message;
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(
                message,
                "Word 实用工具",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
