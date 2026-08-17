using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordMoveToHeading
{
    internal sealed class MoveToHeadingMenu : IDisposable
    {
        private const string RootTag = "WordMoveToHeading.Root";
        private const string ItemTagPrefix = "WordMoveToHeading.Item.";
        private const int MaxHeadings = 200;

        private static readonly string[] ContextMenuNames = { "Text", "Table Text" };

        private readonly Word.Application _application;
        private readonly List<Office.CommandBarPopup> _rootMenus = new List<Office.CommandBarPopup>();
        private readonly List<Office.CommandBarButton> _headingButtons = new List<Office.CommandBarButton>();
        private bool _isRefreshing;
        private bool _isMoving;
        private bool _disposed;

        public MoveToHeadingMenu(Word.Application application)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }
            _application = application;
        }

        public void Start()
        {
            _application.WindowSelectionChange += OnWindowSelectionChange;
            _application.DocumentChange += OnDocumentChange;
            EnsureRootMenus();
            RefreshMenus();
        }

        private void OnWindowSelectionChange(Word.Selection selection)
        {
            RefreshMenus();
        }

        private void OnDocumentChange()
        {
            EnsureRootMenus();
            RefreshMenus();
        }

        private void EnsureRootMenus()
        {
            if (_rootMenus.Count > 0)
            {
                return;
            }

            foreach (string contextMenuName in ContextMenuNames)
            {
                Office.CommandBar commandBar = null;
                try
                {
                    commandBar = _application.CommandBars[contextMenuName];
                    DeleteExistingRoot(commandBar);

                    Office.CommandBarPopup popup = (Office.CommandBarPopup)commandBar.Controls.Add(
                        Office.MsoControlType.msoControlPopup,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);

                    popup.Caption = "Move to";
                    popup.Tag = RootTag;
                    popup.BeginGroup = true;
                    _rootMenus.Add(popup);
                }
                catch (COMException)
                {
                    // Một số bản Word không cung cấp mọi CommandBar theo tên trên.
                }
                finally
                {
                    ReleaseComObject(commandBar);
                }
            }
        }

        private static void DeleteExistingRoot(Office.CommandBar commandBar)
        {
            for (int index = commandBar.Controls.Count; index >= 1; index--)
            {
                Office.CommandBarControl control = null;
                try
                {
                    control = commandBar.Controls[index];
                    if (string.Equals(control.Tag, RootTag, StringComparison.Ordinal))
                    {
                        control.Delete(false);
                    }
                }
                finally
                {
                    ReleaseComObject(control);
                }
            }
        }

        private void RefreshMenus()
        {
            if (_disposed || _isRefreshing || _isMoving)
            {
                return;
            }

            _isRefreshing = true;
            try
            {
                ReleaseHeadingButtons();
                ClearRootMenuItems();

                Word.Document document = TryGetActiveDocument();
                if (document == null)
                {
                    AddDisabledItem("Không có tài liệu đang mở");
                    return;
                }

                try
                {
                    List<HeadingInfo> headings = ReadHeadings(document);
                    if (headings.Count == 0)
                    {
                        AddDisabledItem("Không tìm thấy Heading");
                        return;
                    }

                    foreach (HeadingInfo heading in headings)
                    {
                        AddHeadingItem(heading);
                    }

                    if (headings.Count == MaxHeadings)
                    {
                        AddDisabledItem("Chỉ hiển thị 200 Heading đầu tiên");
                    }
                }
                finally
                {
                    ReleaseComObject(document);
                }
            }
            catch (COMException)
            {
                // Word có thể đang ở trạng thái modal; lần SelectionChange sau sẽ thử lại.
            }
            finally
            {
                _isRefreshing = false;
            }
        }

        private Word.Document TryGetActiveDocument()
        {
            try
            {
                return _application.Documents.Count == 0 ? null : _application.ActiveDocument;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static List<HeadingInfo> ReadHeadings(Word.Document document)
        {
            var result = new List<HeadingInfo>();
            Word.Paragraphs paragraphs = null;
            try
            {
                paragraphs = document.Paragraphs;
                int count = paragraphs.Count;
                for (int index = 1; index <= count && result.Count < MaxHeadings; index++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        Word.WdOutlineLevel outlineLevel = paragraph.OutlineLevel;
                        if (outlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
                        {
                            continue;
                        }

                        range = paragraph.Range;
                        string text = CleanParagraphText(range.Text);
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            text = "(Heading không có nội dung)";
                        }

                        string listNumber = string.Empty;
                        try
                        {
                            listNumber = range.ListFormat.ListString;
                        }
                        catch (COMException)
                        {
                            // Heading không đánh số.
                        }

                        result.Add(new HeadingInfo(
                            result.Count + 1,
                            range.Start,
                            (int)outlineLevel,
                            text,
                            listNumber));
                    }
                    finally
                    {
                        ReleaseComObject(range);
                        ReleaseComObject(paragraph);
                    }
                }
            }
            finally
            {
                ReleaseComObject(paragraphs);
            }

            return result;
        }

        private void AddHeadingItem(HeadingInfo heading)
        {
            foreach (Office.CommandBarPopup root in _rootMenus)
            {
                Office.CommandBarButton button = (Office.CommandBarButton)root.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);

                button.Caption = BuildCaption(heading);
                button.Tag = ItemTagPrefix + heading.Start;
                button.Parameter = heading.Start.ToString(System.Globalization.CultureInfo.InvariantCulture);
                button.TooltipText = "Chuyển vùng chọn đến cuối phần “" + heading.Text + "”";
                button.Click += OnHeadingClick;
                _headingButtons.Add(button);
            }
        }

        private void AddDisabledItem(string caption)
        {
            foreach (Office.CommandBarPopup root in _rootMenus)
            {
                Office.CommandBarButton button = (Office.CommandBarButton)root.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                button.Caption = caption;
                button.Enabled = false;
                _headingButtons.Add(button);
            }
        }

        private void OnHeadingClick(Office.CommandBarButton control, ref bool cancelDefault)
        {
            cancelDefault = true;
            int headingStart;
            if (!int.TryParse(control.Parameter, out headingStart))
            {
                return;
            }

            MoveSelectionToHeading(headingStart);
        }

        internal string BuildDynamicMenuXml()
        {
            Word.Document document = TryGetActiveDocument();
            var xml = new StringBuilder();
            xml.Append("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");

            if (document == null)
            {
                xml.Append("<button id=\"WordMoveToHeading.NoDocument\" label=\"Không có tài liệu đang mở\" enabled=\"false\" />");
            }
            else
            {
                try
                {
                    List<HeadingInfo> headings = ReadHeadings(document);
                    if (headings.Count == 0)
                    {
                        xml.Append("<button id=\"WordMoveToHeading.NoHeadings\" label=\"Không tìm thấy Heading\" enabled=\"false\" />");
                    }
                    else
                    {
                        foreach (HeadingInfo heading in headings)
                        {
                            xml.Append("<button id=\"WordMoveToHeading.Heading.");
                            xml.Append(heading.Ordinal);
                            xml.Append("\" label=\"");
                            xml.Append(EscapeXml(BuildCaption(heading)));
                            xml.Append("\" tag=\"");
                            xml.Append(heading.Start);
                            xml.Append("\" onAction=\"OnMoveToHeading\" />");
                        }

                        if (headings.Count == MaxHeadings)
                        {
                            xml.Append("<button id=\"WordMoveToHeading.Limit\" label=\"Chỉ hiển thị 200 Heading đầu tiên\" enabled=\"false\" />");
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(document);
                }
            }

            xml.Append("</menu>");
            return xml.ToString();
        }

        private static string EscapeXml(string value)
        {
            return System.Security.SecurityElement.Escape(value) ?? string.Empty;
        }

        internal void MoveSelectionToHeading(int headingStart)
        {
            Word.Document document = null;
            Word.Selection selection = null;
            Word.Range source = null;
            Word.Range destination = null;
            Word.Range inserted = null;
            Word.UndoRecord undoRecord = null;
            Word.Window activeWindow = null;
            Word.Pane activePane = null;
            bool undoStarted = false;
            int verticalScroll = 0;
            int horizontalScroll = 0;
            bool hasScrollPosition = false;

            _isMoving = true;
            try
            {
                document = _application.ActiveDocument;
                selection = _application.Selection;

                if (selection == null || selection.Range.Start == selection.Range.End)
                {
                    ShowMessage("Hãy bôi đen nội dung cần chuyển trước khi chọn Move to.");
                    return;
                }

                if (selection.StoryType != Word.WdStoryType.wdMainTextStory)
                {
                    ShowMessage("Add-in hiện chỉ chuyển nội dung trong phần chính của tài liệu.");
                    return;
                }

                if (document.ReadOnly || document.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    ShowMessage("Tài liệu đang ở chế độ chỉ đọc hoặc được bảo vệ.");
                    return;
                }

                source = selection.Range.Duplicate;
                int destinationPosition = FindSectionEnd(document, headingStart);

                try
                {
                    activeWindow = _application.ActiveWindow;
                    activePane = activeWindow.ActivePane;
                    verticalScroll = activePane.VerticalPercentScrolled;
                    horizontalScroll = activePane.HorizontalPercentScrolled;
                    hasScrollPosition = true;
                }
                catch (COMException)
                {
                    hasScrollPosition = false;
                }

                if (destinationPosition >= source.Start && destinationPosition <= source.End)
                {
                    ShowMessage("Không thể chuyển vùng chọn vào chính vùng đó.");
                    return;
                }

                undoRecord = _application.UndoRecord;
                undoRecord.StartCustomRecord("Move to heading");
                undoStarted = true;

                destination = document.Range(destinationPosition, destinationPosition);

                // Nếu phần cuối tài liệu chưa kết thúc bằng paragraph mark, tạo một đoạn mới.
                if (destinationPosition > 0)
                {
                    Word.Range previousCharacter = document.Range(destinationPosition - 1, destinationPosition);
                    try
                    {
                        if (previousCharacter.Text != "\r" && previousCharacter.Text != "\a")
                        {
                            destination.InsertBefore("\r");
                            destination.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(previousCharacter);
                    }
                }

                int insertionStart = destination.Start;
                destination.FormattedText = source.FormattedText;
                int insertionEnd = destination.End;

                inserted = document.Range(insertionStart, insertionEnd);
                source.Delete();

                // Stay at the reading position instead of jumping to the target.
                // After Delete, Word collapses the source range at the old location.
                source.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                source.Select();
                if (hasScrollPosition && activePane != null)
                {
                    activePane.HorizontalPercentScrolled = horizontalScroll;
                    activePane.VerticalPercentScrolled = verticalScroll;
                }
            }
            catch (Exception exception)
            {
                ShowMessage("Word không thể chuyển vùng chọn: " + exception.Message);
            }
            finally
            {
                if (undoStarted && undoRecord != null)
                {
                    try
                    {
                        undoRecord.EndCustomRecord();
                    }
                    catch (COMException)
                    {
                    }
                }

                ReleaseComObject(undoRecord);
                ReleaseComObject(inserted);
                ReleaseComObject(destination);
                ReleaseComObject(source);
                ReleaseComObject(selection);
                ReleaseComObject(document);
                ReleaseComObject(activePane);
                ReleaseComObject(activeWindow);
                _isMoving = false;
                RefreshMenus();
            }
        }

        internal void DetectAndAssignOutlineLevels()
        {
            Word.Document document = null;
            Word.Paragraphs paragraphs = null;
            Word.UndoRecord undoRecord = null;
            bool undoStarted = false;
            int changed = 0;
            int alreadyCorrect = 0;
            int skipped = 0;

            try
            {
                document = _application.ActiveDocument;
                if (document == null)
                {
                    ShowMessage("Không có tài liệu đang mở.");
                    return;
                }

                if (document.ReadOnly || document.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    ShowMessage("Tài liệu đang ở chế độ chỉ đọc hoặc được bảo vệ.");
                    return;
                }

                undoRecord = _application.UndoRecord;
                undoRecord.StartCustomRecord("Auto-detect heading levels");
                undoStarted = true;

                paragraphs = document.Paragraphs;
                for (int index = 1; index <= paragraphs.Count; index++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        range = paragraph.Range;
                        string paragraphText = CleanParagraphText(range.Text).TrimStart();
                        if (paragraphText.Length == 0)
                        {
                            continue;
                        }

                        string listLabel = string.Empty;
                        try
                        {
                            listLabel = (range.ListFormat.ListString ?? string.Empty).Trim();
                        }
                        catch (COMException)
                        {
                        }

                        int level = DetectOutlineLevel(paragraphText, listLabel);
                        if (level == 0)
                        {
                            continue;
                        }

                        if ((int)paragraph.OutlineLevel == level)
                        {
                            alreadyCorrect++;
                            continue;
                        }

                        paragraph.OutlineLevel = (Word.WdOutlineLevel)level;
                        changed++;
                    }
                    catch (COMException)
                    {
                        skipped++;
                    }
                    finally
                    {
                        ReleaseComObject(range);
                        ReleaseComObject(paragraph);
                    }
                }

                ShowMessage(
                    "Đã nhận diện Heading.\r\n\r\n" +
                    "Đã cập nhật: " + changed + "\r\n" +
                    "Đã đúng level: " + alreadyCorrect + "\r\n" +
                    "Không thể cập nhật: " + skipped + "\r\n\r\n" +
                    "Có thể nhấn Ctrl+Z để hoàn tác toàn bộ.");
            }
            catch (COMException exception)
            {
                ShowMessage("Word không thể nhận diện Heading: " + exception.Message);
            }
            finally
            {
                if (undoStarted && undoRecord != null)
                {
                    try
                    {
                        undoRecord.EndCustomRecord();
                    }
                    catch (COMException)
                    {
                    }
                }

                ReleaseComObject(paragraphs);
                ReleaseComObject(undoRecord);
                ReleaseComObject(document);
            }
        }

        private static int DetectOutlineLevel(string paragraphText, string listLabel)
        {
            string marker = string.IsNullOrWhiteSpace(listLabel) ? paragraphText : listLabel;

            if (Regex.IsMatch(marker, @"^\s*[IVXLCDM]+\.?(?:\s|$)", RegexOptions.CultureInvariant))
            {
                return 1;
            }

            Match compoundNumber = Regex.Match(
                marker,
                @"^\s*(\d+(?:\.\d+)+)\.?\s*(?:\S|$)",
                RegexOptions.CultureInvariant);
            if (compoundNumber.Success)
            {
                int components = compoundNumber.Groups[1].Value.Split('.').Length;
                return Math.Min(9, components + 1);
            }

            if (Regex.IsMatch(marker, @"^\s*\d+[\.)](?:\s|$)", RegexOptions.CultureInvariant))
            {
                return 2;
            }

            if (Regex.IsMatch(marker, @"^\s*[a-zA-Z]\)(?:\s|$)", RegexOptions.CultureInvariant))
            {
                return 4;
            }

            return 0;
        }

        private static int FindSectionEnd(Word.Document document, int headingStart)
        {
            Word.Paragraphs paragraphs = null;
            try
            {
                paragraphs = document.Paragraphs;
                int headingLevel = -1;
                bool foundHeading = false;

                for (int index = 1; index <= paragraphs.Count; index++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        range = paragraph.Range;

                        if (!foundHeading)
                        {
                            if (range.Start == headingStart)
                            {
                                headingLevel = (int)paragraph.OutlineLevel;
                                foundHeading = true;
                            }
                            continue;
                        }

                        Word.WdOutlineLevel level = paragraph.OutlineLevel;
                        if (level != Word.WdOutlineLevel.wdOutlineLevelBodyText && (int)level <= headingLevel)
                        {
                            return range.Start;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(range);
                        ReleaseComObject(paragraph);
                    }
                }

                if (!foundHeading)
                {
                    throw new InvalidOperationException("Heading đích không còn tồn tại.");
                }

                Word.Range content = document.Content;
                try
                {
                    return Math.Max(content.Start, content.End - 1);
                }
                finally
                {
                    ReleaseComObject(content);
                }
            }
            finally
            {
                ReleaseComObject(paragraphs);
            }
        }

        private static string BuildCaption(HeadingInfo heading)
        {
            string prefix = string.IsNullOrWhiteSpace(heading.ListNumber)
                ? heading.Ordinal + "."
                : heading.ListNumber;
            string caption = prefix + "  " + heading.Text + "  [H" + heading.Level + "]";
            return caption.Length <= 100 ? caption : caption.Substring(0, 97) + "...";
        }

        private static string CleanParagraphText(string value)
        {
            return (value ?? string.Empty).TrimEnd('\r', '\a', '\n', ' ', '\t');
        }

        private void ClearRootMenuItems()
        {
            foreach (Office.CommandBarPopup root in _rootMenus)
            {
                while (root.Controls.Count > 0)
                {
                    Office.CommandBarControl control = null;
                    try
                    {
                        control = root.Controls[1];
                        control.Delete(false);
                    }
                    finally
                    {
                        ReleaseComObject(control);
                    }
                }
            }
        }

        private void ReleaseHeadingButtons()
        {
            foreach (Office.CommandBarButton button in _headingButtons)
            {
                try
                {
                    button.Click -= OnHeadingClick;
                }
                catch (COMException)
                {
                }
                ReleaseComObject(button);
            }
            _headingButtons.Clear();
        }

        private void ShowMessage(string message)
        {
            System.Windows.Forms.MessageBox.Show(
                message,
                "Move to Heading",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;

            _application.WindowSelectionChange -= OnWindowSelectionChange;
            _application.DocumentChange -= OnDocumentChange;
            ReleaseHeadingButtons();

            foreach (Office.CommandBarPopup root in _rootMenus)
            {
                try
                {
                    root.Delete(false);
                }
                catch (COMException)
                {
                }
                ReleaseComObject(root);
            }
            _rootMenus.Clear();
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }

        private sealed class HeadingInfo
        {
            public HeadingInfo(int ordinal, int start, int level, string text, string listNumber)
            {
                Ordinal = ordinal;
                Start = start;
                Level = level;
                Text = text;
                ListNumber = listNumber;
            }

            public int Ordinal { get; private set; }
            public int Start { get; private set; }
            public int Level { get; private set; }
            public string Text { get; private set; }
            public string ListNumber { get; private set; }
        }
    }
}

