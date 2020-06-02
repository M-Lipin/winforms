// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Windows.Forms.Primitives.Resources;
using static Interop;
using static Interop.Gdi32;
using static Interop.UiaCore;
using static Interop.User32;

namespace System.Windows.Forms.Automation
{
    [ComVisible(true)]
    internal class UiaTextRange : ITextRangeProvider
    {
        private readonly IRawElementProviderSimple _enclosingElement;

        private readonly UiaTextProvider _provider;

        private int _start = 0;
        private int _end = 0;

        // Edit controls always use "\r\n" as the line separator, not "\n".
        // This string is a non-localizable string.
        private const string _lineSeparator = "\r\n";

        public UiaTextRange(IRawElementProviderSimple enclosingElement, UiaTextProvider provider, int start, int end)
        {
            if (enclosingElement is null || provider is null)
            {
                throw new ArgumentNullException(SR.UiaTextRangeConstructorNullParameterExceptionMessage);
            }

            _enclosingElement = enclosingElement;
            _provider = provider;

            if (start > 0)
            {
                _start = start;
                _end = start;
            }

            if (end > _start)
            {
                _end = end;
            }
        }

        /// <summary>
        ///  Last caret position of this text range.
        /// </summary>
        internal int End
        {
            get => _end;
            set
            {
                // Ensure that we never accidentally get a negative index.
                if (value < 0)
                {
                    _end = 0;
                }
                else
                {
                    _end = value;
                }

                // Ensure that end never moves before start.
                if (_end < _start)
                {
                    _start = _end;
                }
            }
        }

        internal int Length
        {
            get
            {
                if (Start < 0 || End < 0 || Start > End)
                {
                    return 0;
                }

                // The subtraction of caret positions returns the length of the text.
                return End - Start;
            }
        }

        /// <summary>
        ///  First caret position of this text range.
        /// </summary>
        internal int Start
        {
            get => _start;
            set
            {
                // Ensure that start never moves after end.
                if (value > _end)
                {
                    _end = value;
                }

                // Ensure that we never accidentally get a negative index.
                if (value < 0)
                {
                    _start = 0;
                }
                else
                {
                    _start = value;
                }
            }
        }

        /// <remarks>
        ///  <para>Strictly only needs to be == since never should _start &gt; _end.</para>
        /// </remarks>
        private bool IsDegenerate => _start > _end ? throw new ArgumentException("_start mustn't exceed _end") : _start == _end;

        ITextRangeProvider ITextRangeProvider.Clone() => new UiaTextRange(_enclosingElement, _provider, Start, End);

        BOOL ITextRangeProvider.Compare(ITextRangeProvider range)
        {
            // Ranges come from the same element. Only need to compare endpoints.
            UiaTextRange editRange = (UiaTextRange)range;

            bool result = editRange.Start == Start && editRange.End == End;
            return result.ToBOOL();
        }

        int ITextRangeProvider.CompareEndpoints(
            TextPatternRangeEndpoint endpoint,
            ITextRangeProvider targetRange,
            TextPatternRangeEndpoint targetEndpoint)
        {
            UiaTextRange editRange = (UiaTextRange)targetRange;
            int e1 = (endpoint == (int)TextPatternRangeEndpoint.Start) ? Start : End;
            int e2 = (targetEndpoint == (int)TextPatternRangeEndpoint.Start) ? editRange.Start : editRange.End;

            return e1 - e2;
        }

        void ITextRangeProvider.ExpandToEnclosingUnit(TextUnit unit)
        {
            switch (unit)
            {
                case TextUnit.Character:
                    // Leave it as it is except the case with 0-range.
                    if (IsDegenerate)
                    {
                        End = MoveEndpointForward(End, TextUnit.Character, 1, out int moved);
                    }
                    break;

                case TextUnit.Word:
                    {
                        // Get the word boundaries.
                        string text = _provider.Text;
                        ValidateEndpoints();

                        // Move start left until we reach a word boundary.
                        while (!AtWordBoundary(text, Start))
                        {
                            Start--;
                        }

                        // Move end right until we reach word boundary (different from Start).
                        End = Math.Min(Math.Max(End, Start + 1), text.Length);

                        while (!AtWordBoundary(text, End))
                        {
                            End++;
                        }
                    }
                    break;

                case TextUnit.Line:
                    {
                        if (_provider.LinesCount != 1)
                        {
                            int startLine = _provider.GetLineFromCharIndex(Start);
                            int startIndex = _provider.GetLineIndex(startLine);

                            int endLine = _provider.GetLineFromCharIndex(End);
                            int endIndex;
                            if (endLine < _provider.LinesCount - 1)
                            {
                                endLine++;
                                endIndex = _provider.GetLineIndex(endLine);
                            }
                            else
                            {
                                endIndex = _provider.TextLength;
                            }

                            MoveTo(startIndex, endIndex);
                        }
                        else
                        {
                            MoveTo(0, _provider.TextLength);
                        }
                    }
                    break;

                case TextUnit.Paragraph:
                    {
                        // Get the paragraph boundaries.
                        string text = _provider.Text;
                        ValidateEndpoints();

                        // Move start left until we reach a paragraph boundary.
                        while (!AtParagraphBoundary(text, Start))
                        {
                            Start--;
                        }

                        // Move end right until we reach a paragraph boundary (different from Start).
                        End = Math.Min(Math.Max(End, Start + 1), text.Length);
                        while (!AtParagraphBoundary(text, End))
                        {
                            End++;
                        }
                    }
                    break;

                case TextUnit.Format:
                case TextUnit.Page:
                case TextUnit.Document:
                    MoveTo(0, _provider.TextLength);
                    break;

                default:
                    throw new InvalidEnumArgumentException(nameof(unit), (int)unit, typeof(TextUnit));
            }
        }

        ITextRangeProvider ITextRangeProvider.FindAttribute(int attributeId, object val, BOOL backwards)
        {
            return null;
        }

        ITextRangeProvider ITextRangeProvider.FindText(string text, BOOL backwards, BOOL ignoreCase)
        {
            if (text is null)
            {
                Debug.Fail("Invalid text range argument. 'text' should not be null.");
                return null;
            }

            if (text.Length == 0)
            {
                Debug.Fail("Invalid text range argument. 'text' length should be more than 0.");
                return null;
            }

            string rangeText = _provider.Text;
            ValidateEndpoints();
            rangeText = rangeText.Substring(Start, Length);

            if (ignoreCase.IsTrue())
            {
                rangeText = rangeText.ToLowerInvariant();
                text = text.ToLowerInvariant();
            }

            // Do a case-sensitive search for the text inside the range.
            int i = backwards.IsTrue() ? rangeText.LastIndexOf(text, StringComparison.Ordinal) : rangeText.IndexOf(text, StringComparison.Ordinal);

            // If the text was found then create a new range covering the found text.
            return i >= 0 ? new UiaTextRange(_enclosingElement, _provider, Start + i, Start + i + text.Length) : null;
        }

        object ITextRangeProvider.GetAttributeValue(int attributeId)
        {
            return GetAttributeValue((TextAttributeIdentifier)attributeId);
        }

        double[] ITextRangeProvider.GetBoundingRectangles()
        {
            // Return zero rectangles for a degenerate-range. We don't return an empty,
            // but properly positioned, rectangle for degenerate ranges.
            if (IsDegenerate)
            {
                return Array.Empty<double>();
            }

            string text = _provider.Text;
            ValidateEndpoints();

            Rectangle ownerBounds = Drawing.Rectangle.Empty;
            object boundsPropertyValue = _enclosingElement.GetPropertyValue(UIA.BoundingRectanglePropertyId);
            if (boundsPropertyValue != null)
            {
                ownerBounds = (Rectangle)boundsPropertyValue;
            }

            // Get the mapping from client coordinates to screen coordinates.
            Point mapClientToScreen = new Point(ownerBounds.X, ownerBounds.Y);

            // Clip the rectangles to the edit control's formatting rectangle.
            Rectangle clippingRectangle = _provider.BoundingRectangle;

            // We accumulate rectangles onto a list.
            Collections.ArrayList rectangles;

            if (_provider.IsMultiline)
            {
                rectangles = GetMultilineBoundingRectangles(text, mapClientToScreen, clippingRectangle);
            }
            else
            {
                rectangles = new Collections.ArrayList(1);

                // Figure out the rectangle for this one line.
                Point startPoint = _provider.GetPositionFromChar(Start);
                Point endPoint = _provider.GetPositionFromCharUR(End - 1, text);
                Rectangle rectangle = new Rectangle(startPoint.X, startPoint.Y, endPoint.X - startPoint.X, clippingRectangle.Height);
                rectangle.Intersect(clippingRectangle);

                if (rectangle.Width > 0 && rectangle.Height > 0)
                {
                    rectangle.Offset(mapClientToScreen.X, mapClientToScreen.Y);
                    rectangles.Add(rectangle);
                }
            }

            // Convert the list of rectangles into an array for returning.
            Rectangle[] rectArray = new Rectangle[rectangles.Count];
            rectangles.CopyTo(rectArray);

            return _provider.RectArrayToDoubleArray(rectArray);
        }

        IRawElementProviderSimple ITextRangeProvider.GetEnclosingElement()
        {
            return _enclosingElement;
        }

        string ITextRangeProvider.GetText(int maxLength)
        {
            if (maxLength == -1)
            {
                maxLength = End + 1;
            }

            string text = _provider.Text;
            ValidateEndpoints();

            maxLength = maxLength >= 0 ? Math.Min(Length, maxLength) : Length;

            if (text.Length < maxLength - Start)
            {
                return text.Substring(Start);
            }

            return text.Substring(Start, maxLength);
        }

        int ITextRangeProvider.Move(TextUnit unit, int count)
        {
            // Positive count means move forward. Negative count means move backwards.
            int moved;

            if (count > 0)
            {
                // If the range is non-degenerate then we need to collapse the range.
                // (See the discussion of Count for ITextRange::Move)
                if (!IsDegenerate)
                {
                    // If the count is greater than zero, collapse the range at its end point.
                    Start = End;
                }

                // Move the degenerate range forward by the number of units.
                int start = Start;
                Start = MoveEndpointForward(Start, unit, count, out moved);

                // If the start did not change then no move was done.
                if (start != Start)
                {
                    return moved;
                }
            }

            if (count < 0)
            {
                // If the range is non-degenerate then we need to collapse the range.
                if (!IsDegenerate)
                {
                    // If the count is less than zero, collapse the range at the starting point.
                    End = Start;
                }

                // Move the degenerate range backward by the number of units.
                int end = End;
                End = MoveEndpointBackward(End, unit, count, out moved);

                // If the end did not change then no move was done.
                if (end != End)
                {
                    return moved;
                }
            }

            // Moving zero of any unit has no effect.
            return 0;
        }

        int ITextRangeProvider.MoveEndpointByUnit(
            TextPatternRangeEndpoint endpoint,
            TextUnit unit, int count)
        {
            // Positive count means move forward. Negative count means move backwards.
            bool moveStart = endpoint == TextPatternRangeEndpoint.Start;
            int moved;
            int start = Start;
            int end = End;

            if (count > 0)
            {
                if (moveStart)
                {
                    Start = MoveEndpointForward(Start, unit, count, out moved);

                    // If the start did not change then no move was done.
                    if (start == Start)
                    {
                        return 0;
                    }

                    return moved;
                }

                End = MoveEndpointForward(End, unit, count, out moved);

                // If the end did not change then no move was done.
                if (end == End)
                {
                    return 0;
                }

                return moved;
            }

            if (count < 0)
            {
                if (moveStart)
                {
                    Start = MoveEndpointBackward(Start, unit, count, out moved);

                    // If the start did not change then no move was done.
                    if (start == Start)
                    {
                        return 0;
                    }

                    return moved;
                }

                End = MoveEndpointBackward(End, unit, count, out moved);

                // If the end did not change then no move was done.
                if (end == End)
                {
                    return 0;
                }

                return moved;
            }

            // Moving zero of any unit has no effect.
            return 0;
        }

        void ITextRangeProvider.MoveEndpointByRange(
            TextPatternRangeEndpoint endpoint,
            ITextRangeProvider targetRange,
            TextPatternRangeEndpoint targetEndpoint)
        {
            UiaTextRange textRange = (UiaTextRange)targetRange;
            int e = (targetEndpoint == TextPatternRangeEndpoint.Start)
                ? textRange.Start
                : textRange.End;

            if (endpoint == TextPatternRangeEndpoint.Start)
            {
                Start = e;
            }
            else
            {
                End = e;
            }
        }

        void ITextRangeProvider.Select()
        {
            _provider.SetSelection(Start, End);
        }

        void ITextRangeProvider.AddToSelection()
        {
            // Do nothing. Do not throw exception.
        }

        void ITextRangeProvider.RemoveFromSelection()
        {
            // Do nothing. Do not throw exception.
        }

        void ITextRangeProvider.ScrollIntoView(BOOL alignToTop)
        {
            if (_provider.IsMultiline)
            {
                int newFirstLine;

                if (alignToTop.IsTrue())
                {
                    newFirstLine = _provider.GetLineFromCharIndex(Start);
                }
                else
                {
                    newFirstLine = Math.Max(0, _provider.GetLineFromCharIndex(End) - _provider.LinesPerPage + 1);
                }

                _provider.LineScroll(Start, newFirstLine - _provider.FirstVisibleLine);

                return;
            }

            if (_provider.IsScrollable)
            {
                _provider.GetVisibleRangePoints(out int visibleStart, out int visibleEnd);

                short key;

                if (_provider.IsReadingRTL)
                {
                    key = VK.LEFT;

                    if (Start > visibleStart)
                    {
                        key = VK.RIGHT;
                    }

                    if (Start > visibleStart || Start < visibleEnd)
                    {
                        _provider.SendKeyboardInputVK(key, true);
                        _provider.GetVisibleRangePoints(out visibleStart, out visibleEnd);
                    }

                    return;
                }

                key = VK.RIGHT;

                if (Start < visibleStart)
                {
                    key = VK.LEFT;
                }

                if (Start < visibleStart || Start > visibleEnd)
                {
                    _provider.SendKeyboardInputVK(key, true);
                    _provider.GetVisibleRangePoints(out visibleStart, out visibleEnd);
                }
            }
        }

        IRawElementProviderSimple[] ITextRangeProvider.GetChildren()
        {
            // We don't have any children so return an empty array.
            return Array.Empty<IRawElementProviderSimple>();
        }

        private static bool AtParagraphBoundary(string text, int index)
        {
            // Returns true if index identifies a paragraph boundary within text.
            return string.IsNullOrWhiteSpace(text) || index <= 0 || index >= text.Length || ((text[index - 1] == '\n') && (text[index] != '\n'));
        }

        private static bool AtWordBoundary(string text, int index)
        {
            // Returns true if index identifies a word boundary within text.
            // Following richedit & word precedent the boundaries are at the leading edge of the word
            // so the span of a word includes trailing whitespace.

            // We are at a word boundary if we are at the beginning or end of the text.
            if (string.IsNullOrWhiteSpace(text) || index <= 0 || index >= text.Length)
            {
                return true;
            }

            if (AtParagraphBoundary(text, index))
            {
                return true;
            }

            char ch1 = text[index - 1];
            char ch2 = text[index];

            // An apostrophe does *not* break a word if it follows or precedes characters.
            if ((char.IsLetterOrDigit(ch1) && IsApostrophe(ch2))
                || (IsApostrophe(ch1) && char.IsLetterOrDigit(ch2) && index >= 2 && char.IsLetterOrDigit(text[index - 2])))
            {
                return false;
            }

            // The following transitions mark boundaries.
            // Note: these are constructed to include trailing whitespace.
            return (char.IsWhiteSpace(ch1) && !char.IsWhiteSpace(ch2))
                || (char.IsLetterOrDigit(ch1) && !char.IsLetterOrDigit(ch2))
                || (!char.IsLetterOrDigit(ch1) && char.IsLetterOrDigit(ch2))
                || (char.IsPunctuation(ch1) && char.IsWhiteSpace(ch2));
        }

        private static bool IsApostrophe(char ch)
        {
            return ch == '\'' ||
                   ch == (char)0x2019; // Unicode Right Single Quote Mark
        }

        private object GetAttributeValue(TextAttributeIdentifier textAttributeIdentifier)
        {
            switch (textAttributeIdentifier)
            {
                case TextAttributeIdentifier.BackgroundColorAttributeId:
                    return GetBackgroundColor();

                case TextAttributeIdentifier.CapStyleAttributeId:
                    return GetCapStyle(_provider.EditStyle);

                case TextAttributeIdentifier.FontNameAttributeId:
                    return GetFontName(_provider.Logfont);

                case TextAttributeIdentifier.FontSizeAttributeId:
                    return GetFontSize(_provider.Logfont);

                case TextAttributeIdentifier.FontWeightAttributeId:
                    return GetFontWeight(_provider.Logfont);

                case TextAttributeIdentifier.ForegroundColorAttributeId:
                    return GetForegroundColor();

                case TextAttributeIdentifier.HorizontalTextAlignmentAttributeId:
                    return GetHorizontalTextAlignment(_provider.EditStyle);

                case TextAttributeIdentifier.IsItalicAttributeId:
                    return GetItalic(_provider.Logfont);

                case TextAttributeIdentifier.IsReadOnlyAttributeId:
                    return GetReadOnly();

                case TextAttributeIdentifier.StrikethroughStyleAttributeId:
                    return GetStrikethroughStyle(_provider.Logfont);

                case TextAttributeIdentifier.UnderlineStyleAttributeId:
                    return GetUnderlineStyle(_provider.Logfont);

                default:
                    return UiaGetReservedNotSupportedValue();
            }
        }

        /// <summary>
        /// Helper function to accumulate a list of bounding rectangles for a potentially mult-line range.
        /// </summary>
        private Collections.ArrayList GetMultilineBoundingRectangles(string text, Drawing.Point mapClientToScreen, Drawing.Rectangle clippingRectangle)
        {
            // Remember the line height.
            int height = Math.Abs(_provider.Logfont.lfHeight);

            // Get the starting and ending lines for the range.
            int start = Start;
            int end = End;

            int startLine = _provider.GetLineFromCharIndex(start);
            int endLine = _provider.GetLineFromCharIndex(end - 1);

            // Adjust the start based on the first visible line.
            int firstVisibleLine = _provider.FirstVisibleLine;

            if (firstVisibleLine > startLine)
            {
                startLine = firstVisibleLine;
                start = _provider.GetLineIndex(startLine);
            }

            // Adjust the end based on the last visible line.
            int lastVisibleLine = firstVisibleLine + _provider.LinesPerPage - 1;

            if (lastVisibleLine < endLine)
            {
                endLine = lastVisibleLine;
                end = _provider.GetLineIndex(endLine) - 1;
            }

            // Adding a rectangle for each line.
            Collections.ArrayList rects = new Collections.ArrayList(Math.Max(endLine - startLine + 1, 0));
            int nextLineIndex = _provider.GetLineIndex(startLine);

            for (int i = startLine; i <= endLine; i++)
            {
                // Determine the starting coordinate on this line.
                Point startPoint;

                if (i == startLine)
                {
                    startPoint = _provider.GetPositionFromChar(start);
                }
                else
                {
                    startPoint = _provider.GetPositionFromChar(nextLineIndex);
                }

                // Determine the ending coordinate on this line.
                Point endPoint;

                if (i == endLine)
                {
                    endPoint = _provider.GetPositionFromCharUR(end - 1, text);
                }
                else
                {
                    nextLineIndex = _provider.GetLineIndex(i + 1);
                    endPoint = _provider.GetPositionFromChar(nextLineIndex - 1);
                }

                // Add a bounding rectangle for this line if it is nonempty.
                Rectangle rect = new Rectangle(startPoint.X, startPoint.Y, endPoint.X - startPoint.X, height);
                rect.Intersect(clippingRectangle);

                if (rect.Width > 0 && rect.Height > 0)
                {
                    rect.Offset(mapClientToScreen.X, mapClientToScreen.Y);
                    rects.Add(rect);
                }
            }

            return rects;
        }

        private HorizontalTextAlignment GetHorizontalTextAlignment(ES editStyle)
        {
            if (editStyle.IsBitSet(ES.CENTER))
            {
                return HorizontalTextAlignment.Centered;
            }

            if (editStyle.IsBitSet(ES.RIGHT))
            {
                return HorizontalTextAlignment.Right;
            }

            return HorizontalTextAlignment.Left;
        }

        private CapStyle GetCapStyle(ES editStyle)
        {
            return editStyle.IsBitSet(ES.UPPERCASE) ? CapStyle.AllCap : CapStyle.None;
        }

        private bool GetReadOnly()
        {
            return _provider.IsReadOnly;
        }

        private static object GetBackgroundColor()
        {
            return GetSysColor(COLOR.WINDOW);
        }

        private static object GetFontName(LOGFONTW logfont)
        {
            return logfont.FaceName.ToString();
        }

        private object GetFontSize(LOGFONTW logfont)
        {
            // Note: this assumes integral point sizes. violating this assumption would confuse the user
            // because they set something to 7 point but reports that it is, say 7.2 point, due to the rounding.
            IntPtr hdc = GetDC(IntPtr.Zero);

            if (hdc == IntPtr.Zero)
            {
                return null;
            }

            int lpy = GetDeviceCaps(hdc, DeviceCapability.LOGPIXELSY);
            _provider.ReleaseDC(hdc);

            return Math.Round((double)(-logfont.lfHeight) * 72 / lpy);
        }

        private static object GetFontWeight(LOGFONTW logfont)
        {
            return logfont.lfWeight;
        }

        private static object GetForegroundColor()
        {
            return GetSysColor(COLOR.WINDOWTEXT);
        }

        private static object GetItalic(LOGFONTW logfont)
        {
            return logfont.lfItalic != 0;
        }

        private static object GetStrikethroughStyle(LOGFONTW logfont)
        {
            return logfont.lfStrikeOut != 0 ? TextDecorationLineStyle.Single : TextDecorationLineStyle.None;
        }

        private static object GetUnderlineStyle(LOGFONTW logfont)
        {
            return logfont.lfUnderline != 0 ? TextDecorationLineStyle.Single : TextDecorationLineStyle.None;
        }

        /// <summary>
        /// Moves an endpoint forward a certain number of units.
        /// </summary>
        private int MoveEndpointForward(int index, TextUnit unit, int count, out int moved)
        {
            switch (unit)
            {
                case TextUnit.Character:
                    {
                        int limit = _provider.TextLength;
                        ValidateEndpoints();

                        moved = Math.Min(count, limit - index);
                        index = index + moved;

                        index = index > limit ? limit : index;
                    }
                    break;

                case TextUnit.Word:
                    {
                        string text = _provider.Text;
                        ValidateEndpoints();

                        moved = 0;
                        while (moved < count && index < text.Length)
                        {
                            index++;

                            while (!AtWordBoundary(text, index))
                            {
                                index++;
                            }

                            moved++;
                        }
                    }
                    break;

                case TextUnit.Line:
                    {
                        // Figure out what line we are on. If we are in the middle of a line and
                        // are moving left then we'll round up to the next line so that we move
                        // to the beginning of the current line.
                        int line = _provider.GetLineFromCharIndex(index);

                        // Limit the number of lines moved to the number of lines available to move
                        // Note lineMax is always >= 1.
                        int lineMax = _provider.LinesCount;
                        moved = Math.Min(count, lineMax - line - 1);

                        if (moved > 0)
                        {
                            // move the endpoint to the beginning of the destination line.
                            index = _provider.GetLineIndex(line + moved);
                        }
                        else if (moved == 0 && lineMax == 1)
                        {
                            // There is only one line so get the text length as endpoint.
                            index = _provider.TextLength;
                            moved = 1;
                        }
                    }
                    break;

                case TextUnit.Paragraph:
                    {
                        // Just like moving words but we look for paragraph boundaries instead of
                        // word boundaries.
                        string text = _provider.Text;
                        ValidateEndpoints();

                        moved = 0;
                        while (moved < count && index < text.Length)
                        {
                            index++;

                            while (!AtParagraphBoundary(text, index))
                            {
                                index++;
                            }

                            moved++;
                        }
                    }
                    break;

                case TextUnit.Format:
                case TextUnit.Page:
                case TextUnit.Document:
                    {
                        // Since edit controls are plain text moving one uniform format unit will
                        // take us all the way to the end of the document, just like
                        // "pages" and document.
                        int limit = _provider.TextLength;
                        ValidateEndpoints();

                        // We'll move 1 format unit if we aren't already at the end of the
                        // document.  Otherwise, we won't move at all.
                        moved = index < limit ? 1 : 0;
                        index = limit;
                    }
                    break;

                default:
                    throw new InvalidEnumArgumentException(nameof(unit), (int)unit, typeof(TextUnit));
            }

            return index;
        }

        /// <summary>
        /// Moves an endpoint backward a certain number of units.
        /// </summary>
        private int MoveEndpointBackward(int index, TextUnit unit, int count, out int moved)
        {
            switch (unit)
            {
                case TextUnit.Character:
                    {
                        int limit = _provider.TextLength;
                        ValidateEndpoints();

                        int oneBasedIndex = index + 1;

                        moved = Math.Max(count, -oneBasedIndex);
                        index = index + moved;

                        index = index < 0 ? 0 : index;
                    }
                    break;

                case TextUnit.Word:
                    {
                        string text = _provider.Text;
                        ValidateEndpoints();

                        for (moved = 0; moved > count && index > 0; moved--)
                        {
                            index--;

                            while (!AtWordBoundary(text, index))
                            {
                                index--;
                            }
                        }
                    }
                    break;

                case TextUnit.Line:
                    {
                        // Note count < 0.

                        // Get 1-based line.
                        int line = _provider.GetLineFromCharIndex(index) + 1;

                        int lineMax = _provider.LinesCount;

                        // Truncate the count to the number of available lines.
                        int actualCount = Math.Max(count, -line);

                        moved = actualCount;

                        if (actualCount == -line)
                        {
                            // We are moving by the maximum number of possible lines,
                            // so we know the resulting index will be 0.
                            index = 0;

                            // If a line other than the first consists of only "\r\n",
                            // you can move backwards past this line and the position changes,
                            // hence this is counted.  The first line is special, though:
                            // if it is empty, and you move say from the second line back up
                            // to the first, you cannot move further; however if the first line
                            // is nonempty, you can move from the end of the first line to its
                            // beginning!  This latter move is counted, but if the first line
                            // is empty, it is not counted.

                            // Recalculate the value of "moved".
                            // The first line is empty if it consists only of
                            // a line separator sequence.
                            bool firstLineEmpty =
                                (lineMax == 0 || (lineMax > 1 && _provider.GetLineIndex(1) == _lineSeparator.Length));

                            if (moved < 0 && firstLineEmpty)
                            {
                                ++moved;
                            }
                        }
                        else // actualCount > -line
                        {
                            // Move the endpoint to the beginning of the following line,
                            // then back by the line separator length to get to the end
                            // of the previous line, since the Edit control has
                            // no method to get the character index of the end
                            // of a line directly.
                            index = _provider.GetLineIndex(line + actualCount) - _lineSeparator.Length;
                        }
                    }
                    break;

                case TextUnit.Paragraph:
                    {
                        // Just like moving words but we look for paragraph boundaries instead of
                        // word boundaries.
                        string text = _provider.Text;
                        ValidateEndpoints();

                        for (moved = 0; moved > count && index > 0; moved--)
                        {
                            index--;
                            while (!AtParagraphBoundary(text, index))
                            {
                                index--;
                            }
                        }
                    }
                    break;

                case TextUnit.Format:
                case TextUnit.Page:
                case TextUnit.Document:
                    {
                        // Since edit controls are plain text moving one uniform format unit will
                        // take us all the way to the beginning of the document, just like
                        // "pages" and document.

                        // We'll move 1 format unit if we aren't already at the beginning of the
                        // document.  Otherwise, we won't move at all.
                        moved = index > 0 ? -1 : 0;
                        index = 0;
                    }
                    break;

                default:
                    throw new InvalidEnumArgumentException(nameof(unit), (int)unit, typeof(TextUnit));
            }

            return index;
        }

        /// <summary>
        /// Method to set both endpoints simultaneously.
        /// </summary>
        private void MoveTo(int start, int end)
        {
            if (start < 0 || end < start)
            {
                throw new InvalidOperationException(SR.NoOffetTextRangeInvalidOperationExceptionMessage);
            }

            _start = start;
            _end = end;
        }

        private void ValidateEndpoints()
        {
            int limit = _provider.TextLength;

            if (Start > limit && limit > 0)
            {
                Start = limit;
            }

            if (End > limit && limit > 0)
            {
                End = limit;
            }
        }
    }
}
