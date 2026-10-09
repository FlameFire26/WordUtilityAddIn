using System;
using System.IO;
using System.Runtime.InteropServices;
using WordUtilityAddIn.Core.Models;
using WordUtilityAddIn.Core.Services;
using Word = Microsoft.Office.Interop.Word;

namespace WordUtilityAddIn.Services
{
    internal enum RubyScope
    {
        FullDocument,
        Selection
    }

    internal sealed class WordDocumentService
    {
        private readonly Func<Word.Application> applicationAccessor;
        private readonly OoxmlRubyProcessor processor;

        public WordDocumentService(Func<Word.Application> applicationAccessor)
        {
            this.applicationAccessor = applicationAccessor
                ?? throw new ArgumentNullException(nameof(applicationAccessor));
            processor = new OoxmlRubyProcessor();
        }

        public RubyStatistics Scan(RubyScope scope)
        {
            Word.Application application = GetApplication();
            Word.Document document = GetActiveDocument(application);

            if (scope == RubyScope.Selection)
            {
                Word.Range range = GetSupportedSelection(application);
                return processor.ScanXml(range.WordOpenXML);
            }

            return processor.ScanXml(document.WordOpenXML);
        }

        public RubyUpdateResult Apply(RubyScope scope, RubyFormatOptions options)
        {
            return scope == RubyScope.Selection
                ? ApplyToSelection(options)
                : ApplyToFullDocument(options);
        }

        private RubyUpdateResult ApplyToFullDocument(RubyFormatOptions options)
        {
            Word.Application application = GetApplication();
            Word.Document document = GetActiveDocument(application);
            string originalPath = GetSavedSupportedPath(document);

            if (document.ReadOnly)
            {
                throw new InvalidOperationException("当前文档为只读，不能执行全文修改。");
            }

            if (document.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                throw new InvalidOperationException("当前文档处于保护状态，请先取消保护。");
            }

            document.Save();
            string backupPath = CreateTimestampedSiblingPath(originalPath, ".backup-");
            SaveCopy(document, backupPath);

            string extension = Path.GetExtension(originalPath);
            string directory = Path.GetDirectoryName(originalPath);
            string workingPath = Path.Combine(
                directory,
                ".WordUtilityAddIn-" + Guid.NewGuid().ToString("N") + extension);
            string rollbackPath = Path.Combine(
                directory,
                ".WordUtilityAddIn-rollback-" + Guid.NewGuid().ToString("N") + extension);

            bool documentClosed = false;
            bool replaced = false;
            try
            {
                File.Copy(originalPath, workingPath, false);
                RubyUpdateResult result = processor.UpdatePackage(workingPath, options);

                CloseWithoutSaving(document);
                documentClosed = true;

                File.Replace(workingPath, originalPath, rollbackPath, true);
                replaced = true;

                application.Documents.Open(FileName: originalPath, AddToRecentFiles: true);
                result.BackupPath = backupPath;
                return result;
            }
            catch
            {
                if (documentClosed)
                {
                    TryOpenDocument(application, originalPath);
                }

                throw;
            }
            finally
            {
                TryDelete(workingPath);
                if (replaced)
                {
                    TryDelete(rollbackPath);
                }
            }
        }

        private RubyUpdateResult ApplyToSelection(RubyFormatOptions options)
        {
            Word.Application application = GetApplication();
            Word.Document document = GetActiveDocument(application);
            Word.Range range = GetSupportedSelection(application);
            string sourceXml = range.WordOpenXML;
            RubyUpdateResult intendedResult;
            string updatedXml = processor.UpdateXml(sourceXml, options, out intendedResult);

            if (intendedResult.ModifiedRubyCount == 0)
            {
                return intendedResult;
            }

            Word.UndoRecord undoRecord = application.UndoRecord;
            bool undoRecordStarted = false;
            bool inserted = false;
            try
            {
                undoRecord.StartCustomRecord("修改拼音指南格式");
                undoRecordStarted = true;

                object transformSource = Type.Missing;
                range.InsertXML(updatedXml, ref transformSource);
                inserted = true;

                undoRecord.EndCustomRecord();
                undoRecordStarted = false;

                RubyStatistics actual = processor.ValidateXmlFormatting(range.WordOpenXML, options);
                if (actual.RubyCount != intendedResult.Before.RubyCount
                    || !string.Equals(
                        actual.ReadingFingerprint,
                        intendedResult.Before.ReadingFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidDataException("选区替换后的注音数量或读音校验失败。");
                }

                intendedResult.After = actual;
                return intendedResult;
            }
            catch (Exception operationException)
            {
                if (undoRecordStarted)
                {
                    try
                    {
                        undoRecord.EndCustomRecord();
                    }
                    catch (COMException)
                    {
                    }
                }

                if (inserted)
                {
                    bool restored = TryUndo(document);
                    if (!restored)
                    {
                        restored = TryRestoreRangeXml(range, sourceXml);
                    }

                    if (!restored)
                    {
                        throw new InvalidOperationException(
                            "选区修改校验失败，而且自动回滚未成功。请立即关闭文档并选择不保存。",
                            operationException);
                    }
                }

                throw;
            }
        }

        private static Word.Range GetSupportedSelection(Word.Application application)
        {
            Word.Selection selection = application.Selection;
            if (selection == null || selection.Range == null || selection.Range.Start == selection.Range.End)
            {
                throw new InvalidOperationException("请先选择包含拼音标注的正文内容。");
            }

            Word.Range range = selection.Range.Duplicate;
            if (range.StoryType != Word.WdStoryType.wdMainTextStory)
            {
                throw new NotSupportedException("第一版选区处理仅支持主文档正文；页眉、页脚等内容请使用全文模式。");
            }

            return range;
        }

        private static string GetSavedSupportedPath(Word.Document document)
        {
            if (string.IsNullOrWhiteSpace(document.Path))
            {
                throw new InvalidOperationException("全文处理前请先保存文档。");
            }

            string path = document.FullName;
            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension != ".docx"
                && extension != ".docm"
                && extension != ".dotx"
                && extension != ".dotm")
            {
                throw new NotSupportedException("全文处理仅支持 docx、docm、dotx 和 dotm 文件。");
            }

            return path;
        }

        private Word.Application GetApplication()
        {
            Word.Application application = applicationAccessor();
            if (application == null)
            {
                throw new InvalidOperationException("无法连接到 Word 应用程序。");
            }

            return application;
        }

        private static Word.Document GetActiveDocument(Word.Application application)
        {
            if (application.Documents.Count == 0)
            {
                throw new InvalidOperationException("当前没有打开的 Word 文档。");
            }

            return application.ActiveDocument;
        }

        private static string CreateTimestampedSiblingPath(string originalPath, string marker)
        {
            string directory = Path.GetDirectoryName(originalPath);
            string name = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string candidate = Path.Combine(directory, name + marker + timestamp + extension);
            int suffix = 1;

            while (File.Exists(candidate))
            {
                candidate = Path.Combine(
                    directory,
                    string.Format("{0}{1}{2}-{3}{4}", name, marker, timestamp, suffix, extension));
                suffix++;
            }

            return candidate;
        }

        private static void SaveCopy(Word.Document document, string backupPath)
        {
            object fileName = backupPath;
            document.SaveCopyAs(ref fileName);
        }

        private static void CloseWithoutSaving(Word.Document document)
        {
            object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = Type.Missing;
            object routeDocument = false;
            document.Close(ref saveChanges, ref originalFormat, ref routeDocument);
        }

        private static void TryOpenDocument(Word.Application application, string path)
        {
            try
            {
                application.Documents.Open(FileName: path, AddToRecentFiles: false);
            }
            catch (COMException)
            {
            }
        }

        private static bool TryUndo(Word.Document document)
        {
            try
            {
                object times = 1;
                return document.Undo(ref times);
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static bool TryRestoreRangeXml(Word.Range range, string sourceXml)
        {
            try
            {
                object transformSource = Type.Missing;
                range.InsertXML(sourceXml, ref transformSource);
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
    }
}
