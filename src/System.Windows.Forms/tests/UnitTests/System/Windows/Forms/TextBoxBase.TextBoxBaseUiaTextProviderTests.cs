// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms.Automation;
using Moq;
using Xunit;
using static System.Windows.Forms.TextBoxBase;
using static Interop;
using static Interop.User32;

namespace System.Windows.Forms.Tests
{
    public class TextBoxBaseUiaTextProviderTests : IClassFixture<ThreadExceptionFixture>
    {
        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_ctor_DoesntCreateControlHandle()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            Assert.False(textBoxBase.IsHandleCreated);

            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Assert.False(textBoxBase.IsHandleCreated);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_IsMultiline_IsCorrect()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);

            textBoxBase.Multiline = false;
            Assert.False(provider.IsMultiline);

            textBoxBase.Multiline = true;
            Assert.True(provider.IsMultiline);
        }

        [WinFormsTheory]
        [InlineData(true)]
        [InlineData(false)]
        public void TextBoxBaseUiaTextProvider_IsReadOnly_IsCorrect(bool readOnly)
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);

            textBoxBase.ReadOnly = readOnly;
            Assert.Equal(readOnly, provider.IsReadOnly);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_IsScrollable_IsCorrect()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);

            Assert.True(provider.IsScrollable);
        }

        [WinFormsFact]
        public void TextBoxBaseTextProvider_GetWindowStyle_ReturnsNoneForNotInitializedControl()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Assert.Equal(WS.OVERLAPPED, provider.WindowStyle);
        }

        [WinFormsTheory]
        [InlineData(RightToLeft.Yes, true)]
        [InlineData(RightToLeft.No, false)]
        public void TextBoxBaseUiaTextProvider_IsReadingRTL_ReturnsCorrectValue(RightToLeft rightToLeft, bool expectedResult)
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            textBoxBase.CreateControl();
            textBoxBase.RightToLeft = rightToLeft;

            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Assert.Equal(expectedResult, provider.IsReadingRTL);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_LinesPerPage_IsCorrect()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Assert.Equal(textBoxBase.LinesPerPage, provider.LinesPerPage);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_OwningControl_IsCorrect()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Assert.Equal(textBoxBase, provider.OwningControl);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_DocumentRange_IsNotNull()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Assert.NotNull(provider.DocumentRange);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_SupportedTextSelection_IsNotNull()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            UiaCore.SupportedTextSelection uiaTextRange = provider.SupportedTextSelection;
            Assert.Equal(UiaCore.SupportedTextSelection.Single, uiaTextRange);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_GetCaretRange_IsNotNull()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            UiaCore.ITextRangeProvider uiaTextRange = provider.GetCaretRange(out _);
            Assert.NotNull(uiaTextRange);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_GetFirstVisibleLine_ReturnsCorrectValue()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);

            int line = provider.FirstVisibleLine;
            Assert.Equal(0, line);

            textBoxBase.Multiline = true;

            textBoxBase.Size = new Size(30, 100);
            line = provider.FirstVisibleLine;
            Assert.Equal(0, line);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_GetLineCount_ReturnsCorrectValue()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Assert.Equal(1, provider.LinesCount);

            textBoxBase.Multiline = true;
            textBoxBase.Size = new Size(30, 50);
            Assert.Equal(1, provider.LinesCount);

            textBoxBase.Text += "1\r\n";
            Assert.Equal(2, provider.LinesCount);

            textBoxBase.Text += "2\r\n";
            Assert.Equal(3, provider.LinesCount);

            textBoxBase.Text += "3\r\n";
            Assert.Equal(4, provider.LinesCount);

            textBoxBase.Text += "4\r\n";
            Assert.Equal(5, provider.LinesCount);
        }

        public static IEnumerable<object[]> TextBoxBase_GetLineFromCharIndex_TestData()
        {
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Multiline = false }, 0, 0 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Multiline = false }, 50, 0 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 50), Multiline = true }, 50, 3 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 50), Multiline = true }, 50, 8 };
        }

        [WinFormsTheory]
        [MemberData(nameof(TextBoxBase_GetLineFromCharIndex_TestData))]
        public void TextBoxBaseUiaTextProvider_GetLineFromCharIndex_ReturnsCorrectValue(TextBoxBase textBoxBase, int charIndex, int expectedLine)
        {
            textBoxBase.Text = "Some test text for testing GetLineFromCharIndex method";
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            int actualLine = provider.GetLineFromCharIndex(charIndex);
            Assert.Equal(expectedLine, actualLine);
        }

        public static IEnumerable<object[]> TextBoxBase_GetLineIndex_TestData()
        {
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Multiline = false }, 0, 0 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Multiline = false }, 3, 0 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 50), Multiline = true }, 3, 19 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 50), Multiline = true }, 3, 40 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 50), Multiline = true }, 100, -1 };
        }

        [WinFormsTheory]
        [MemberData(nameof(TextBoxBase_GetLineIndex_TestData))]
        public void TextBoxBaseUiaTextProvider_GetLineIndex_ReturnsCorrectValue(TextBoxBase textBoxBase, int lineIndex, int expectedIndex)
        {
            textBoxBase.Text = "Some test text for testing GetLineIndex method";
            int actualIndex = textBoxBase.GetLineIndex(lineIndex);
            Assert.Equal(expectedIndex, actualIndex);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_ReleaseDC_ReturnsCorrectValue()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            IntPtr hdc = Interop.User32.GetDC(IntPtr.Zero);
            int actual = provider.ReleaseDC(hdc);
            Assert.Equal(0x00000001, actual);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_GetLogfont_ReturnsCorrectValue()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            textBoxBase.CreateControl();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);

            User32.LOGFONTW expected = textBoxBase.GetLogfont();
            User32.LOGFONTW actual = provider.Logfont;
            Assert.False(string.IsNullOrEmpty(actual.FaceName.ToString()));
            Assert.Equal(expected, actual);

            // GetLogfont method uses GetFont method which mustn't return IntPtr.Zero
            Assert.NotEqual(IntPtr.Zero, textBoxBase.GetFont());
        }

        public static IEnumerable<object[]> TextBoxBase_GetPositionFromChar_TestData()
        {
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Text = "Some test text for testing", Multiline = false }, 0, new Point(1, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Text = "Some test text for testing", Multiline = false }, 15, new Point(79, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Text = "Some test text for testing", Multiline = true }, 15, new Point(27, 30) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "This is a\r\nlong long text\r\nfor testing\r\nGetPositionFromChar method", Multiline = true }, 0, new Point(4, 1) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "This is a\r\nlong long text\r\nfor testing\r\nGetPositionFromChar method", Multiline = true }, 6, new Point(31, 1) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "This is a\r\nlong long text\r\nfor testing\r\nGetPositionFromChar method", Multiline = true }, 26, new Point(78, 16) };
        }

        [WinFormsTheory]
        [MemberData(nameof(TextBoxBase_GetPositionFromChar_TestData))]
        public void TextBoxBaseUiaTextProvider_GetPositionFromChar_ReturnsCorrectValue(TextBoxBase textBoxBase, int charIndex, Point expectedPoint)
        {
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Point actualPoint = provider.GetPositionFromChar(charIndex);
            Assert.True(actualPoint.X >= expectedPoint.X - 1 || actualPoint.X <= expectedPoint.X + 1);
            Assert.True(actualPoint.Y >= expectedPoint.Y - 1 || actualPoint.Y <= expectedPoint.Y + 1);
        }

        public static IEnumerable<object[]> TextBoxBaseUiaTextProvider_GetPositionFromCharUR_ReturnsCorrectValue_TestData()
        {
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Text = "", Multiline = false }, 0, new Point(0, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Text = "Some test text", Multiline = false }, 100, new Point(0, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Text = "Some test text", Multiline = false }, -1, new Point(0, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Text = "Some test text", Multiline = false }, 12, new Point(71, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20), Text = "Some test text", Multiline = true }, 12, new Point(19, 30) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "Some test \n text", Multiline = false }, 10, new Point(56, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "Some test \n text", Multiline = true }, 10, new Point(59, 1) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "Some test \r\n text", Multiline = false }, 10, new Point(56, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "Some test \r\n text", Multiline = true }, 10, new Point(59, 1) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "Some test \r\n text", Multiline = false }, 12, new Point(60, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "Some test \r\n text", Multiline = true }, 12, new Point(7, 16) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "Some test \t text", Multiline = false }, 10, new Point(57, 0) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 60), Text = "Some test \t text", Multiline = true }, 10, new Point(60, 1) };
            yield return new object[] { new SubTextBoxBase { Size = new Size(40, 60), Text = "Some test \t text", Multiline = true }, 12, new Point(8, 46) };
        }

        [WinFormsTheory]
        [MemberData(nameof(TextBoxBaseUiaTextProvider_GetPositionFromCharUR_ReturnsCorrectValue_TestData))]
        public void TextBoxBaseUiaTextProvider_GetPositionFromCharUR_ReturnsCorrectValue(TextBoxBase textBoxBase, int charIndex, Point expectedPoint)
        {
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Point actualPoint = provider.GetPositionFromCharUR(charIndex, textBoxBase.Text);
            Assert.True(actualPoint.X >= expectedPoint.X - 1 || actualPoint.X <= expectedPoint.X + 1);
            Assert.True(actualPoint.Y >= expectedPoint.Y - 1 || actualPoint.Y <= expectedPoint.Y + 1);
        }

        public static IEnumerable<object[]> TextBoxBase_GetRectangle_TestData()
        {
            yield return new object[] { new SubTextBoxBase { Multiline = false, Size = new Size(0, 0) }, new Rectangle(1, 0, 78, 16) };
            yield return new object[] { new SubTextBoxBase { Multiline = false, Size = new Size(50, 50) }, new Rectangle(1, 1, 44, 15) };
            yield return new object[] { new SubTextBoxBase { Multiline = false, Size = new Size(250, 100) }, new Rectangle(1, 1, 244, 15) };
            yield return new object[] { new SubTextBoxBase { Multiline = true, Size = new Size(0, 0) }, new Rectangle(4, 0, 72, 16) };
            yield return new object[] { new SubTextBoxBase { Multiline = true, Size = new Size(50, 50) }, new Rectangle(4, 1, 38, 30) };
            yield return new object[] { new SubTextBoxBase { Multiline = true, Size = new Size(250, 100) }, new Rectangle(4, 1, 238, 90) };
        }

        [WinFormsTheory]
        [MemberData(nameof(TextBoxBase_GetRectangle_TestData))]
        public void TextBoxBaseUiaTextProvider_GetRectangle_ReturnsCorrectValue(TextBoxBase textBoxBase, Rectangle expectedRectangle)
        {
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);

            Rectangle boxRectangle = textBoxBase.GetRectangle();
            Rectangle providerRectangle = provider.BoundingRectangle;
            Assert.Equal(boxRectangle, providerRectangle);
            Assert.Equal(expectedRectangle, providerRectangle);
        }

        [WinFormsTheory]
        [InlineData("")]
        [InlineData("Text")]
        [InlineData("Some test text")]
        public void TextBoxBaseUiaTextProvider_Text_ReturnsCorrectValue(string text)
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            textBoxBase.Text = text;
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            textBoxBase.CreateControl();
            string expected = textBoxBase.Text;
            string actual = provider.Text.Trim('\0');
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData("")]
        [InlineData("Text")]
        [InlineData("Some test text")]
        public void TextBoxBaseUiaTextProvider_TextLength_ReturnsCorrectValue(string text)
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            textBoxBase.Text = text;
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            int expected = textBoxBase.Text.Length;
            int actual = provider.TextLength;
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData(false, true, 0x00000200)]
        [InlineData(true, true, 0x00000000)]
        [InlineData(true, false, 0x00000000)]
        [InlineData(false, false, 0x00000000)]
        public void TextBoxBaseUiaTextProvider_WindowExStyle_ReturnsCorrectValue(bool nullOwner, bool createHandle, uint expected)
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();

            if (createHandle)
            {
                textBoxBase.CreateControl();
            }

            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(nullOwner ? null : textBoxBase);
            User32.WS_EX actual = provider.WindowExStyle;
            Assert.Equal((User32.WS_EX)expected, actual);
        }

        [WinFormsTheory]
        [InlineData(false, true, 0x560100c0)]
        [InlineData(true, true, 0x0000)]
        [InlineData(true, false, 0x0000)]
        [InlineData(false, false, 0x0000)]
        public void TextBoxBaseUiaTextProvider_EditStyle_ReturnsCorrectValue(bool nullOwner, bool createHandle, uint expected)
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();

            if (createHandle)
            {
                textBoxBase.CreateControl();
            }

            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(nullOwner ? null : textBoxBase);
            User32.ES actual = provider.EditStyle;
            Assert.Equal((User32.ES)expected, actual);
        }

        public static IEnumerable<object[]> TextBoxBase_GetVisibleRangePoints_TestData()
        {
            yield return new object[] { new SubTextBoxBase { Size = new Size(0, 0) }, 0, 0 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(0, 20) }, 0, 0 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(50, 20) }, 0, 7 };
            yield return new object[] { new SubTextBoxBase { Size = new Size(120, 20) }, 0, 23 };
        }

        [WinFormsTheory]
        [MemberData(nameof(TextBoxBase_GetVisibleRangePoints_TestData))]
        public void TextBoxBaseUiaTextProvider_GetVisibleRangePoints_ReturnsCorrectValue(TextBoxBase textBoxBase, int expectedStart, int expectedEnd)
        {
            textBoxBase.Text = "Some test text for testing";
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            provider.GetVisibleRangePoints(out int providerVisibleStart, out int providerVisibleEnd);
            textBoxBase.GetVisibleRangePoints(out int textBoxBaseVisibleStart, out int textBoxBaseVisibleEnd);
            Assert.Equal(textBoxBaseVisibleStart, providerVisibleStart);
            Assert.Equal(textBoxBaseVisibleEnd, providerVisibleEnd);

            Assert.True(providerVisibleStart >= 0);
            Assert.True(providerVisibleStart < textBoxBase.Text.Length);
            Assert.True(providerVisibleEnd >= 0);
            Assert.True(providerVisibleEnd < textBoxBase.Text.Length);

            Assert.Equal(expectedStart, providerVisibleStart);
            Assert.Equal(expectedEnd, providerVisibleEnd);
        }

        public static IEnumerable<object[]> TextBoxBase_GetVisibleRanges_TestData()
        {
            yield return new object[] { new SubTextBoxBase { Size = new Size(0, 0) } };
            yield return new object[] { new SubTextBoxBase { Size = new Size(100, 20) } };
        }

        [WinFormsTheory]
        [MemberData(nameof(TextBoxBase_GetVisibleRanges_TestData))]
        public void TextBoxBaseUiaTextProvider_GetVisibleRanges_ReturnsCorrectValue(TextBoxBase textBoxBase)
        {
            textBoxBase.Text = "Some test text for testing";
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            Assert.NotNull(provider.GetVisibleRanges());
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_RangeFromAnnotation_DoesntThrowAnException()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);

            // RangeFromAnnotation doesn't throw an exception
            UiaCore.ITextRangeProvider range = provider.RangeFromAnnotation(textBoxBase.AccessibilityObject);
            // RangeFromAnnotation implementation can be changed so this test can be changed too
            Assert.NotNull(range);
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_RangeFromChild_DoesntThrowAnException()
        {
            using (new NoAssertContext())
            {
                using TextBoxBase textBoxBase = new SubTextBoxBase();
                TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);

                // RangeFromChild doesn't throw an exception
                UiaCore.ITextRangeProvider range = provider.RangeFromChild(textBoxBase.AccessibilityObject);
                // RangeFromChild implementation can be changed so this test can be changed too
                Assert.Null(range);
            }
        }

        [WinFormsFact]
        public void TextBoxBaseUiaTextProvider_RangeFromPoint_DoesntThrowAnException()
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);

            UiaTextRange textRangeProvider = provider.RangeFromPoint(Point.Empty) as UiaTextRange;
            Assert.NotNull(textRangeProvider);

            textRangeProvider = provider.RangeFromPoint(new Point(10, 10)) as UiaTextRange;
            Assert.NotNull(textRangeProvider);
        }

        [WinFormsTheory]
        [InlineData(2, 5)]
        [InlineData(0, 10)]
        public void TextBoxBaseUiaTextProvider_SetSelection_IsCorrectAction(int start, int end)
        {
            using TextBoxBase textBoxBase = new SubTextBoxBase();
            textBoxBase.Text = "Some test text for testing";
            TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
            provider.SetSelection(start, end);
            UiaCore.ITextRangeProvider[] selection = provider.GetSelection();
            Assert.NotNull(selection);

            UiaTextRange textRange = selection[0] as UiaTextRange;
            Assert.NotNull(textRange);

            Assert.Equal(start, textRange.Start);
            Assert.Equal(end, textRange.End);
        }

        [WinFormsTheory]
        [InlineData(-5, 10)]
        [InlineData(5, 100)]
        public void TextBoxBaseUiaTextProvider_SetSelection_DoesntSelectText_IfIncorrectArguments(int start, int end)
        {
            using (new NoAssertContext())
            {
                using TextBoxBase textBoxBase = new SubTextBoxBase();
                textBoxBase.Text = "Some test text for testing";
                TextBoxBaseUiaTextProvider provider = new TextBoxBaseUiaTextProvider(textBoxBase);
                provider.SetSelection(start, end);
                UiaCore.ITextRangeProvider[] selection = provider.GetSelection();
                Assert.NotNull(selection);

                UiaTextRange textRange = selection[0] as UiaTextRange;
                Assert.NotNull(textRange);

                Assert.Equal(0, textRange.Start);
                Assert.Equal(0, textRange.End);
            }
        }

        private class SubTextBoxBase : TextBoxBase
        { }
    }
}
