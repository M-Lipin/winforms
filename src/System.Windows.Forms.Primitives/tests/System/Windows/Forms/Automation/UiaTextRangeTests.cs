// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms.Automation;
using Xunit;
using static System.Windows.Forms.TextBoxBase;
using static Interop;
using static Interop.UiaCore;
using static Interop.Gdi32;
using System.Runtime.InteropServices;
using Moq;
using static Interop.User32;

namespace System.Windows.Forms.Primitives.Tests.Automation
{
    public class UiaTextRangeTests
    {
        private UiaTextProvider GetProvider(TextBoxBase textBox)
        {
            Type type = typeof(TextBoxBase).GetNestedType("TextBoxBaseUiaTextProvider", BindingFlags.NonPublic);
            return (UiaTextProvider)Activator.CreateInstance(type, textBox);
        }

        [WinFormsTheory]
        [InlineData(0, 0)]
        [InlineData(0, 5)]
        [InlineData(5, 10)]
        [InlineData(1, 1)]
        public void UiaTextRange_Constructor_InitializesProvider_And_CorrectEndpoints(int start, int end)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            Assert.Equal(start, textRange.Start);
            Assert.Equal(end, textRange.End);
            Assert.Equal(textBox.AccessibilityObject, ((ITextRangeProvider)textRange).GetEnclosingElement());

            object actual = typeof(UiaTextRange).GetField("_provider", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(textRange);
            Assert.Equal(provider, actual);
        }

        [WinFormsTheory]
        [InlineData(-10, 0, 0, 0)]
        [InlineData(0, -10, 0, 0)]
        [InlineData(5, 0, 5, 5)]
        [InlineData(-1, -1, 0, 0)]
        [InlineData(10, 5, 10, 10)]
        [InlineData(-5, 5, 0, 5)]
        [InlineData(5, -5, 5, 5)]
        public void UiaTextRange_Constructor_InitializesProvider_And_CorrectEndpoints_IfEndpointsincorrect(int start, int end, int expectedStart, int expectedEnd)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);

            Assert.Equal(expectedStart, textRange.Start);
            Assert.Equal(expectedEnd, textRange.End);
        }

        [WinFormsFact]
        public void UiaTextRange_Constructor_Provider_Null_ThrowsException()
        {
            using TextBox textBox = new TextBox();
            Assert.Throws<ArgumentNullException>(() => new UiaTextRange(textBox.AccessibilityObject, null, 0, 5));
        }

        [WinFormsFact]
        public void UiaTextRange_Constructor_Control_Null_ThrowsException()
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            Assert.Throws<ArgumentNullException>(() => new UiaTextRange(null, provider, 0, 5));
        }

        [WinFormsTheory]
        [InlineData(3, -5)]
        [InlineData(-5, 3)]
        [InlineData(-3, -5)]
        public void UiaTextRange_Constructor_SetCorrectValues_IfNegativeStartEnd(int start, int end)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange range = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            Assert.True(range.Start >= 0);
            Assert.True(range.End >= 0);
        }

        [WinFormsTheory]
        [InlineData(0)]
        [InlineData(5)]
        [InlineData(int.MaxValue)]
        public void UiaTextRange_End_Get_ReturnsCorrectValue(int end)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, end);
            Assert.Equal(end, textRange.End);
        }

        [WinFormsTheory]
        [InlineData(0)]
        [InlineData(5)]
        [InlineData(int.MaxValue)]
        public void UiaTextRange_End_SetCorrectly(int end)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, /*end*/ 0);
            textRange.End = end;
            int actual = textRange.End < textRange.Start ? textRange.Start : textRange.End;
            Assert.Equal(end, actual);
        }

        [WinFormsFact]
        public void UiaTextRange_End_SetCorrect_IfValueIncorrect()
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 5, /*end*/ 10);
            textRange.End = 3;  /*Incorrect value*/
            Assert.Equal(textRange.Start, textRange.End);

            textRange.End = 6;
            Assert.Equal(6, textRange.End);

            textRange.End = -10; /*Incorrect value*/
            Assert.Equal(textRange.Start, textRange.End);
        }

        [WinFormsTheory]
        [InlineData(0, 0, 0)]
        [InlineData(5, 5, 0)]
        [InlineData(3, 15, 12)]
        [InlineData(0, 10, 10)]
        [InlineData(6, 10, 4)]
        public void UiaTextRange_Length_ReturnsCorrectValue(int start, int end, int expected)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            Assert.Equal(expected, textRange.Length);
        }

        [WinFormsTheory]
        [InlineData(-5, 0)]
        [InlineData(0, -5)]
        [InlineData(-5, -5)]
        [InlineData(10, 5)]
        public void UiaTextRange_Length_ReturnsCorrectValue_IfIncorrectStartEnd(int start, int end)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 3, 10);
            var startField = textRange.GetType().
                GetField("_start", BindingFlags.NonPublic | BindingFlags.Instance);
            startField.SetValue(textRange, start);

            var endField = textRange.GetType().
                GetField("_end", BindingFlags.NonPublic | BindingFlags.Instance);
            endField.SetValue(textRange, end);

            Assert.Equal(0, textRange.Length);
        }

        [WinFormsTheory]
        [InlineData(0)]
        [InlineData(5)]
        [InlineData(int.MaxValue)]
        public void UiaTextRange_Start_Get_ReturnsCorrectValue(int start)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/0, /*end*/ 0);
            var startField = textRange.GetType().
                GetField("_start", BindingFlags.NonPublic | BindingFlags.Instance);
            startField.SetValue(textRange, start);
            Assert.Equal(start, textRange.Start);
        }

        [WinFormsTheory]
        [InlineData(0)]
        [InlineData(5)]
        [InlineData(int.MaxValue)]
        public void UiaTextRange_Start_SetCorrectly(int start)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, /*end*/ 0);
            textRange.Start = start;
            int actual = textRange.Start < textRange.End ? textRange.End : textRange.Start;
            Assert.Equal(start, actual);
        }

        [WinFormsFact]
        public void UiaTextRange_Start_Set_Correct_IfValueIncorrect()
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 4, /*end*/ 8);
            textRange.Start = -10;
            Assert.Equal(0, textRange.Start);
            Assert.Equal(8, textRange.End);
        }

        [WinFormsFact]
        public void UiaTextRange_Start_Set_Correct_IfValueMoreThanEnd()
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 4, /*end*/ 10);
            textRange.Start = 15; // More than End = 10
            Assert.True(textRange.Start <= textRange.End);
        }

        [WinFormsFact]
        public void UiaTextRange_ITextRangeProvider_Clone_ReturnsCorrectValue()
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 3, /*end*/ 9);
            UiaTextRange actual = (UiaTextRange)((ITextRangeProvider)textRange).Clone();
            Assert.Equal(textRange.Start, actual.Start);
            Assert.Equal(textRange.End, actual.End);
        }

        [WinFormsTheory]
        [InlineData(3, 9, true)]
        [InlineData(0, 2, false)]
        public void UiaTextRange_ITextRangeProvider_Compare_ReturnsCorrectValue(int start, int end, bool expected)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange1 = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 3, /*end*/ 9);
            UiaTextRange textRange2 = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ start, /*end*/ end);
            bool actual = ((ITextRangeProvider)textRange1).Compare(textRange2).IsTrue();
            Assert.Equal(expected, actual);
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_CompareEndpoints_ReturnsCorrectValue_TestData()
        {
            yield return new object[] { TextPatternRangeEndpoint.Start, 3, 9, TextPatternRangeEndpoint.Start, 0 };
            yield return new object[] { TextPatternRangeEndpoint.End, 3, 9, TextPatternRangeEndpoint.Start, 6 };
            yield return new object[] { TextPatternRangeEndpoint.Start, 3, 9, TextPatternRangeEndpoint.End, -6 };
            yield return new object[] { TextPatternRangeEndpoint.End, 3, 9, TextPatternRangeEndpoint.End, 0 };
            yield return new object[] { TextPatternRangeEndpoint.Start, 0, 0, TextPatternRangeEndpoint.Start, 3 };
            yield return new object[] { TextPatternRangeEndpoint.End, 0, 0, TextPatternRangeEndpoint.Start, 9 };
            yield return new object[] { TextPatternRangeEndpoint.End, 1, 15, TextPatternRangeEndpoint.End, -6 };
            yield return new object[] { TextPatternRangeEndpoint.Start, 1, 15, TextPatternRangeEndpoint.End, -12 };
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_CompareEndpoints_ReturnsCorrectValue_TestData))]
        public void UiaTextRange_ITextRangeProvider_CompareEndpoints_ReturnsCorrectValue(
            int endpoint,
            int targetStart,
            int targetEnd,
            int targetEndpoint,
            int expected)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 3, /*end*/ 9);
            UiaTextRange targetRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ targetStart, /*end*/ targetEnd);
            int actual = ((ITextRangeProvider)textRange).CompareEndpoints((TextPatternRangeEndpoint)endpoint, targetRange, (TextPatternRangeEndpoint)targetEndpoint);
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData(2, 2, 2, 3)]
        [InlineData(8, 9, 8, 9)]
        [InlineData(0, 3, 0, 3)]
        public void UiaTextRange_ITextRangeProvider_ExpandToEnclosingUnit_ExpandsToCharacter(int start, int end, int expandedStart, int expandedEnd)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "words, words, words";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            ((ITextRangeProvider)textRange).ExpandToEnclosingUnit(TextUnit.Character);
            Assert.Equal(expandedStart, textRange.Start);
            Assert.Equal(expandedEnd, textRange.End);
        }

        [WinFormsTheory]
        [InlineData(2, 3, 0, 5)]
        [InlineData(8, 8, 7, 12)]
        [InlineData(16, 17, 14, 19)]
        public void UiaTextRange_ITextRangeProvider_ExpandToEnclosingUnit_ExpandsToWord(int start, int end, int expandedStart, int expandedEnd)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "words, words, words";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            ((ITextRangeProvider)textRange).ExpandToEnclosingUnit(TextUnit.Word);
            Assert.Equal(expandedStart, textRange.Start);
            Assert.Equal(expandedEnd, textRange.End);
        }

        [WinFormsTheory]
        [InlineData(2, 4, 0, 12)]
        [InlineData(15, 16, 12, 25)]
        [InlineData(27, 28, 25, 36)]
        public void UiaTextRange_ITextRangeProvider_ExpandToEnclosingUnit_ExpandsToLine(int start, int end, int expandedStart, int expandedEnd)
        {
            using TextBox textBox = new TextBox();
            textBox.Multiline = true;
            textBox.Text =
@"First line
second line
third line.";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            ((ITextRangeProvider)textRange).ExpandToEnclosingUnit(TextUnit.Line);
            Assert.Equal(expandedStart, textRange.Start);
            Assert.Equal(expandedEnd, textRange.End);
        }

        [WinFormsTheory]
        [InlineData(2, 4, 0, 24)]
        [InlineData(30, 30, 24, 49)]
        [InlineData(49, 60, 49, 73)]
        public void UiaTextRange_ITextRangeProvider_ExpandToEnclosingUnit_ExpandsToParagraph(int start, int end, int expandedStart, int expandedEnd)
        {
            using TextBox textBox = new TextBox();
            textBox.Text =
@"This is the first line
this is the second line
this is the third line.";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            ((ITextRangeProvider)textRange).ExpandToEnclosingUnit(TextUnit.Paragraph);
            Assert.Equal(expandedStart, textRange.Start);
            Assert.Equal(expandedEnd, textRange.End);
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_ExpandToEnclosingUnit_ExpandsToAllText_TestData()
        {
            yield return new object[] { 5, 8, TextUnit.Page, 0, 72 };
            yield return new object[] { 10, 10, TextUnit.Format, 0, 72 };
            yield return new object[] { 10, 10, TextUnit.Document, 0, 72 };
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_ExpandToEnclosingUnit_ExpandsToAllText_TestData))]
        internal void UiaTextRange_ITextRangeProvider_ExpandToEnclosingUnit_ExpandsToAllText(int start, int end, TextUnit textUnit, int expandedStart, int expandedEnd)
        {
            using TextBox textBox = new TextBox();
            textBox.Text =
@"This is the first line
this is the second line
this is the third line.";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            ((ITextRangeProvider)textRange).ExpandToEnclosingUnit(textUnit);
            Assert.Equal(expandedStart, textRange.Start);
            Assert.Equal(expandedEnd, textRange.End);
        }

        [WinFormsTheory]
        [InlineData(true)]
        [InlineData(false)]
        internal void UiaTextRange_ITextRangeProvider_FindAttribute_Returns_null(bool backward)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, /*end*/ 0);
            var textAttributeIdentifiers = Enum.GetValues(typeof(TextAttributeIdentifier));

            foreach (int textAttributeIdentifier in textAttributeIdentifiers)
            {
                var actual = ((ITextRangeProvider)textRange).FindAttribute(textAttributeIdentifier, new object(), backward.ToBOOL());
                Assert.Null(actual);
            }
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_FindText_Returns_Correct_TestData()
        {
            yield return new object[] { "Test text to find something.", "text", "text", BOOL.FALSE, BOOL.FALSE };
            yield return new object[] { "Test text to find something.", "other", null, BOOL.FALSE, BOOL.FALSE };
            yield return new object[] { "Test text to find something.", "TEXT", "text", BOOL.FALSE, BOOL.TRUE };
            yield return new object[] { "Test text to find something.", "TEXT", null, BOOL.FALSE, BOOL.FALSE };
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_FindText_Returns_Correct_TestData))]
        internal void UiaTextRange_ITextRangeProvider_FindText_Returns_Correct(string text, string textToSearch, string foundText, BOOL backward, BOOL ignoreCase)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = text;
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, /*end*/ 28);

            var actual = ((ITextRangeProvider)textRange).FindText(textToSearch, backward, ignoreCase);

            if (foundText != null)
            {
                Assert.Equal(foundText, actual.GetText(5000));
            }
            else
            {
                Assert.Null(actual);
            }
        }

        [WinFormsFact]
        internal void UiaTextRange_ITextRangeProvider_FindText_ReturnsNull_IfTextNull()
        {
            using (new NoAssertContext())
            {
                using TextBox textBox = new TextBox();
                textBox.Text = "Some long test text";
                UiaTextProvider provider = GetProvider(textBox);
                UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, /*end*/ 28);
                var actual = ((ITextRangeProvider)textRange).FindText(null, BOOL.TRUE, BOOL.TRUE);
                Assert.Null(actual);
            }
        }

        private static object notSupportedValue;

        [DllImport(Libraries.UiaCore, ExactSpelling = true)]
        private static extern int UiaGetReservedNotSupportedValue([MarshalAs(UnmanagedType.IUnknown)] out object notSupportedValue);

        public static object UiaGetReservedNotSupportedValue()
        {
            if (notSupportedValue == null)
            {
                UiaGetReservedNotSupportedValue(out notSupportedValue);
            }

            return notSupportedValue;
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_GetAttributeValue_Returns_Correct_TestData()
        {
            yield return new object[] { TextAttributeIdentifier.BackgroundColorAttributeId, GetSysColor(COLOR.WINDOW) };
            yield return new object[] { TextAttributeIdentifier.CapStyleAttributeId, CapStyle.None };
            yield return new object[] { TextAttributeIdentifier.FontNameAttributeId, "Segoe UI" };
            yield return new object[] { TextAttributeIdentifier.FontSizeAttributeId, 9.0 };
            yield return new object[] { TextAttributeIdentifier.FontWeightAttributeId, FW.NORMAL };
            yield return new object[] { TextAttributeIdentifier.ForegroundColorAttributeId, 0 };
            yield return new object[] { TextAttributeIdentifier.HorizontalTextAlignmentAttributeId, HorizontalTextAlignment.Left };
            yield return new object[] { TextAttributeIdentifier.IsItalicAttributeId, false };
            yield return new object[] { TextAttributeIdentifier.IsReadOnlyAttributeId, false };
            yield return new object[] { TextAttributeIdentifier.StrikethroughStyleAttributeId, TextDecorationLineStyle.None };
            yield return new object[] { TextAttributeIdentifier.UnderlineStyleAttributeId, TextDecorationLineStyle.None };

            yield return new object[] { TextAttributeIdentifier.AnimationStyleAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.BulletStyleAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.CultureAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.IndentationFirstLineAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.IndentationLeadingAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.IndentationTrailingAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.IsHiddenAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.IsSubscriptAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.IsSuperscriptAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.MarginBottomAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.MarginLeadingAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.MarginTopAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.MarginTrailingAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.OutlineStylesAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.OverlineColorAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.OverlineStyleAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.StrikethroughColorAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.TabsAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.TextFlowDirectionsAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.UnderlineColorAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.AnnotationTypesAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.AnnotationObjectsAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.StyleNameAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.StyleIdAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.LinkAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.IsActiveAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.SelectionActiveEndAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.CaretPositionAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.CaretBidiModeAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.LineSpacingAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.BeforeParagraphSpacingAttributeId, UiaGetReservedNotSupportedValue() };
            yield return new object[] { TextAttributeIdentifier.AfterParagraphSpacingAttributeId, UiaGetReservedNotSupportedValue() };
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_GetAttributeValue_Returns_Correct_TestData))]
        internal void UiaTextRange_ITextRangeProvider_GetAttributeValue_Returns_Correct(int attributeId, object attributeValue)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "Some text to set for the TextBox to test.";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, /*end*/ 28);
            var actual = ((ITextRangeProvider)textRange).GetAttributeValue(attributeId);

            Assert.Equal(attributeValue, actual);
        }

        [WinFormsFact]
        public void UiaTextRange_ITextRangeProvider_GetBoundingRectangles_ReturnsEmpty_for_DegenerateRange()
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, /*end*/ 0);
            var actual = ((ITextRangeProvider)textRange).GetBoundingRectangles();

            Assert.Empty(actual);
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_GetBoundingRectangles_ReturnsCorrectValue_for_SingleLine_TestData()
        {
            yield return new object[] { 0, 4, new double[] { 11, 34, 20, 14 } };
            yield return new object[] { 3, 6, new double[] { 27, 34, 11, 14 } };
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_GetBoundingRectangles_ReturnsCorrectValue_for_SingleLine_TestData))]
        public void UiaTextRange_ITextRangeProvider_GetBoundingRectangles_ReturnsCorrectValue_for_SingleLine(int start, int end, double[] expected)
        {
            using TextBox textBox = new TextBox();
            string text = "Test text.";
            textBox.Text = text;
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            double[] actual = ((ITextRangeProvider)textRange).GetBoundingRectangles();

            // Acceptable deviation of 1 px.
            for (int i = 0; i < actual.Length; i++)
            {
                Assert.True(actual[i] >= 0 && actual[i] >= expected[i] - 1 && actual[i] <= expected[i] + 1);
            }
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_GetBoundingRectangles_ReturnsCorrectValue_for_MultiLine_TestData()
        {
            yield return new object[] { 18, 30, new double[] { 14, 49, 9, 12, 14, 64, 39, 12 } };
            yield return new object[] { 32, 35, new double[] { 60, 64, 17, 12 } };
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_GetBoundingRectangles_ReturnsCorrectValue_for_MultiLine_TestData))]
        public void UiaTextRange_ITextRangeProvider_GetBoundingRectangles_ReturnsCorrectValue_for_MultiLine(int start, int end, double[] expected)
        {
            using TextBox textBox = new TextBox();
            textBox.Multiline = true;
            textBox.Height = 60;
            string text =
@"Test text on line 1.
Test text on line 2.";
            textBox.Text = text;
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            var actual = ((ITextRangeProvider)textRange).GetBoundingRectangles();

            Assert.Equal(expected, actual);
        }

        [WinFormsFact]
        public void UiaTextRange_ITextRangeProvider_GetEnclosingElement_ReturnsCorrectValue()
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, /*end*/ 0);
            IRawElementProviderSimple actual = ((ITextRangeProvider)textRange).GetEnclosingElement();
            Assert.Equal(textBox.AccessibilityObject, actual);
        }

        [WinFormsTheory]
        [InlineData(0, 0, 0, "")]
        [InlineData(0, 0, 5, "")]
        [InlineData(0, 10, -5, "Some long ")]
        [InlineData(0, 10, 0, "")]
        [InlineData(0, 10, 10, "Some long ")]
        [InlineData(0, 10, 20, "Some long ")]
        [InlineData(0, 25, 7, "Some lo")]
        [InlineData(0, 300, 400, "Some long long test text")]
        [InlineData(5, 15, 7, "long lo")]
        [InlineData(5, 15, 25, "long long ")]
        [InlineData(5, 15, 300, "long long ")]
        [InlineData(5, 24, 400, "long long test text")]
        [InlineData(5, 25, 0, "")]
        [InlineData(5, 25, 7, "long lo")]
        [InlineData(5, 300, -5, "long long test text")]
        [InlineData(5, 300, 7, "long lo")]
        [InlineData(5, 300, 300, "long long test text")]
        public void UiaTextRange_ITextRangeProvider_GetText_ReturnsCorrectValue(int start, int end, int maxLength, string expected)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "Some long long test text";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            string actual = ((ITextRangeProvider)textRange).GetText(maxLength);
            Debug.WriteLine(expected);
            Assert.Equal(expected, actual);
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_Move_MovesCorrectly_TestData()
        {
            yield return new object[] { 0, 5, TextUnit.Character, 1, 6, 6 };
            yield return new object[] { 1, 6, TextUnit.Character, 5, 11, 11 };
            yield return new object[] { 0, 5, TextUnit.Character, -2, 0, 0 };
            yield return new object[] { 3, 6, TextUnit.Character, -2, 1, 1 };
            yield return new object[] { 1, 2, TextUnit.Word, 1, 4, 4 };
            yield return new object[] { 1, 2, TextUnit.Word, 5, 11, 11 };
            yield return new object[] { 12, 14, TextUnit.Word, -2, 8, 8 };
            yield return new object[] { 12, 14, TextUnit.Word, -10, 0, 0 };
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_Move_MovesCorrectly_TestData))]
        internal void UiaTextRange_ITextRangeProvider_Move_MovesCorrectly(int start, int end, TextUnit unit, int count, int expectedStart, int expectedEnd)
        {
            using TextBox textBox = new TextBox();
            textBox.Multiline = true;
            textBox.Text =
@"This is the text to move on - line 1
This is the line 2
This is the line 3";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            int result = ((ITextRangeProvider)textRange).Move(unit, count);

            Assert.Equal(expectedStart, textRange.Start);
            Assert.Equal(expectedEnd, textRange.End);
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_MoveEndpointByUnit_MovesCorrectly_TestData()
        {
            yield return new object[] { 0, 5, TextPatternRangeEndpoint.Start, TextUnit.Character, 1, 1, 5 };
            yield return new object[] { 1, 6, TextPatternRangeEndpoint.Start, TextUnit.Character, 5, 6, 6 };
            yield return new object[] { 0, 5, TextPatternRangeEndpoint.Start, TextUnit.Character, -2, 0, 5 };
            yield return new object[] { 3, 6, TextPatternRangeEndpoint.Start, TextUnit.Character, -2, 1, 6 };
            yield return new object[] { 3, 6, TextPatternRangeEndpoint.End, TextUnit.Character, 1, 3, 7 };
            yield return new object[] { 3, 6, TextPatternRangeEndpoint.End, TextUnit.Character, -1, 3, 5 };
            yield return new object[] { 1, 2, TextPatternRangeEndpoint.Start, TextUnit.Word, 1, 4, 4 };
            yield return new object[] { 1, 2, TextPatternRangeEndpoint.Start, TextUnit.Word, 5, 11, 11 };
            yield return new object[] { 12, 14, TextPatternRangeEndpoint.Start, TextUnit.Word, -1, 11, 14 };
            yield return new object[] { 12, 14, TextPatternRangeEndpoint.Start, TextUnit.Word, -2, 8, 14 };
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_MoveEndpointByUnit_MovesCorrectly_TestData))]
        internal void UiaTextRange_ITextRangeProvider_MoveEndpointByUnit_MovesCorrectly(int start, int end, TextPatternRangeEndpoint endpoint, TextUnit unit, int count, int expectedStart, int expectedEnd)
        {
            using TextBox textBox = new TextBox();
            textBox.Multiline = true;
            textBox.Text =
@"This is the text to move on - line 1
This is the line 2
This is the line 3";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            int result = ((ITextRangeProvider)textRange).MoveEndpointByUnit(endpoint, unit, count);

            Assert.Equal(expectedStart, textRange.Start);
            Assert.Equal(expectedEnd, textRange.End);
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_MoveEndpointByRange_MovesCorrectly_TestData()
        {
            yield return new object[] { 0, 5, TextPatternRangeEndpoint.Start, 7, 10, TextPatternRangeEndpoint.Start, 7, 7 };
            yield return new object[] { 0, 5, TextPatternRangeEndpoint.Start, 7, 10, TextPatternRangeEndpoint.End, 10, 10 };
            yield return new object[] { 0, 5, TextPatternRangeEndpoint.End, 7, 10, TextPatternRangeEndpoint.Start, 0, 7 };
            yield return new object[] { 0, 5, TextPatternRangeEndpoint.End, 7, 10, TextPatternRangeEndpoint.End, 0, 10 };
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_MoveEndpointByRange_MovesCorrectly_TestData))]
        internal void UiaTextRange_ITextRangeProvider_MoveEndpointByRange_MovesCorrectly(int start, int end, TextPatternRangeEndpoint endpoint, int targetRangeStart, int targetRangeEnd, TextPatternRangeEndpoint targetEndpoint, int expectedStart, int expectedEnd)
        {
            using TextBox textBox = new TextBox();
            textBox.Multiline = true;
            textBox.Text =
@"This is the text to move on - line 1
This is the line 2
This is the line 3";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            UiaTextRange targetRange = new UiaTextRange(textBox.AccessibilityObject, provider, targetRangeStart, targetRangeEnd);
            ((ITextRangeProvider)textRange).MoveEndpointByRange(endpoint, targetRange, targetEndpoint);

            Assert.Equal(expectedStart, textRange.Start);
            Assert.Equal(expectedEnd, textRange.End);
        }

        [WinFormsTheory]
        [InlineData(0, 0)]
        [InlineData(0, 10)]
        [InlineData(5, 10)]
        public void UiaTextRange_ITextRangeProvider_Select_ReturnsCorrectValue(int start, int end)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "Some long long test text";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            ((ITextRangeProvider)textRange).Select();
            ITextRangeProvider[] actual = provider.GetSelection();
            Assert.Single(actual);
            bool equal = ((ITextRangeProvider)textRange).Compare(actual[0]).IsTrue();
            Assert.True(equal);
        }

        [WinFormsFact]
        public void UiaTextRange_ITextRangeProvider_AddToSelection_DoesntThrowException()
        {
            // Check an app doesn't crash when calling AddToSelectio method.
            using TextBox textBox = new TextBox();
            textBox.Text = "Some long long test text";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 3, 7);
            ((ITextRangeProvider)textRange).AddToSelection();
            ITextRangeProvider[] actual = provider.GetSelection();
            Assert.Single(actual);
        }

        [WinFormsFact]
        public void UiaTextRange_ITextRangeProvider_RemoveFromSelection_DoesntThrowException()
        {
            // Check an app doesn't crash when calling RemoveFromSelection method.
            using TextBox textBox = new TextBox();
            textBox.Text = "Some long long test text";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 3, 7);
            ((ITextRangeProvider)textRange).RemoveFromSelection();
            ITextRangeProvider[] actual = provider.GetSelection();
            Assert.Single(actual);
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_ScrollIntoView_Multiline_CallsLineScrollCorrectly_TestData()
        {
            yield return new object[] { 30 /* start */, 35 /* end */, 30 /* charIndex */, 1 /* lineForCharIndex */, 30 /* charactersHorizontal */, 0 /* linesVertical */, 1 /* firstVisibleLine */};
            yield return new object[] { 60 /* start */, 65 /* end */, 60 /* charIndex */, 2 /* lineForCharIndex */, 60 /* charactersHorizontal */, 2 /* linesVertical */, 0 /* firstVisibleLine */};
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_ScrollIntoView_Multiline_CallsLineScrollCorrectly_TestData))]
        public void UiaTextRange_ITextRangeProvider_ScrollIntoView_Multiline_CallsLineScrollCorrectly(int start, int end, int charIndex, int lineForCharIndex, int charactersHorizontal, int linesVertical, int firstVisibleLine)
        {
            using TextBox textBox = new TextBox();
            textBox.Text =
@"This is the text - line 1
This is the line 2
This is the line 3";
            UiaTextProvider provider = GetProvider(textBox);

            var mockProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            mockProvider
                .Setup(p => p.IsMultiline)
                .Returns(true);
            mockProvider
                .Setup(p => p.GetLineFromCharIndex(charIndex))
                .Returns(lineForCharIndex);
            mockProvider
                .Setup(p => p.LineScroll(charactersHorizontal, linesVertical))
                .Returns(true);
            mockProvider
                .Setup(p => p.FirstVisibleLine)
                .Returns(firstVisibleLine);

            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, mockProvider.Object, start, end);

            ((ITextRangeProvider)textRange).ScrollIntoView(BOOL.TRUE);
            mockProvider.Verify(e => e.LineScroll(charactersHorizontal, linesVertical), Times.Once());
        }

        public static IEnumerable<object[]> UiaTextRange_ITextRangeProvider_ScrollIntoView_SingleLine_ExecutesCorrectly_TestData()
        {
            yield return new object[] { 0 /* start */, 30 /* end */, true /* scrollable */, false /* readingRTL */};
            yield return new object[] { 70 /* start */, 85 /* end */, true /* scrollable */, false /* readingRTL */};
        }

        [WinFormsTheory]
        [MemberData(nameof(UiaTextRange_ITextRangeProvider_ScrollIntoView_SingleLine_ExecutesCorrectly_TestData))]
        public void UiaTextRange_ITextRangeProvider_ScrollIntoView_SingleLine_ExecutesCorrectly(int start, int end,
            bool scrollable, bool readingRTL)
        {
            using TextBox textBox = new TextBox();
            textBox.Text =
@"This is the text - line 1
This is the line 2
This is the line 3";
            UiaTextProvider provider = GetProvider(textBox);

            int visibleStart = 40;
            int visibleEnd = 60;

            var mockProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            mockProvider
                .Setup(p => p.IsMultiline)
                .Returns(false);
            mockProvider
                .Setup(p => p.IsScrollable)
                .Returns(scrollable);
            mockProvider
                .Setup(p => p.IsReadingRTL)
                .Returns(readingRTL);
            mockProvider
                .Setup(p => p.GetVisibleRangePoints(out visibleStart, out visibleEnd));

            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, mockProvider.Object, start, end);

            ((ITextRangeProvider)textRange).ScrollIntoView(BOOL.TRUE);
            mockProvider.Verify(p => p.GetVisibleRangePoints(out visibleStart, out visibleEnd), Times.Exactly(2));
        }

        [WinFormsFact]
        public void UiaTextRange_ITextRangeProvider_GetChildren_ReturnsCorrectValue()
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, /*start*/ 0, /*end*/ 0);
            IRawElementProviderSimple[] actual = ((ITextRangeProvider)textRange).GetChildren();
            Assert.Empty(actual);
        }

        [WinFormsTheory]
        [InlineData("", 0, true)]
        [InlineData("", 5, true)]
        [InlineData("", -5, true)]
        [InlineData("Some text", 0, true)]
        [InlineData("Some text", 5, false)]
        [InlineData("Some text", 6, false)]
        [InlineData("Some text", 99, true)]
        [InlineData("Some text", -5, true)]
        [InlineData("Some, text", 4, false)]
        [InlineData("Some text", 4, false)]
        [InlineData("1dsf'21gj", 3, false)]
        [InlineData("1dsf'21gj", 4, false)]
        [InlineData("1dsf'21gj", 6, false)]
        [InlineData("1d??sf'21gj", 6, false)]
        public void UiaTextRange_private_AtParagraphBoundary_ReturnsCorrectValue(string text, int index, bool expected)
        {
            bool actual = typeof(UiaTextRange).TestAccessor().Dynamic.AtParagraphBoundary(text, index);
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData("", 0, true)]
        [InlineData("", 5, true)]
        [InlineData("", -5, true)]
        [InlineData("Some text", 0, true)]
        [InlineData("Some text", 5, true)]
        [InlineData("Some text", 6, false)]
        [InlineData("Some text", 99, true)]
        [InlineData("Some text", -5, true)]
        [InlineData("Some, text", 4, true)]
        [InlineData("Some text", 4, true)]
        [InlineData("1dsf'21gj", 3, false)]
        [InlineData("1dsf'21gj", 4, false)]
        [InlineData("1dsf'21gj", 6, false)]
        [InlineData("1d??sf'21gj", 6, false)]
        public void UiaTextRange_private_AtWordBoundary_ReturnsCorrectValue(string text, int index, bool expected)
        {
            bool actual = typeof(UiaTextRange).TestAccessor().Dynamic.AtWordBoundary(text, index);
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData('\'', true)]
        [InlineData((char)0x2019, true)]
        [InlineData('\t', false)]
        [InlineData('t', false)]
        public void UiaTextRange_private_IsApostrophe_ReturnsCorrectValue(char ch, bool expected)
        {
            bool actual = typeof(UiaTextRange).TestAccessor().Dynamic.IsApostrophe(ch);
            Assert.Equal(expected, actual);
        }

        [WinFormsFact]
        public void UiaTextRange_private_GetAttributeValue_ReturnsCorrectValue()
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 0, 0);

            foreach (TextAttributeIdentifier identifier in Enum.GetValues(typeof(TextAttributeIdentifier)))
            {
                object value = textRange.TestAccessor().Dynamic.GetAttributeValue(identifier);
                Assert.NotNull(value);
            }
        }

        [WinFormsTheory]
        [InlineData(ES.CENTER, HorizontalTextAlignment.Centered)]
        [InlineData(ES.LEFT, HorizontalTextAlignment.Left)]
        [InlineData(ES.RIGHT, HorizontalTextAlignment.Right)]
        public void UiaTextRange_private_GetHorizontalTextAlignment_ReturnsCorrectValue(object style, HorizontalTextAlignment expected)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 0, 0);
            HorizontalTextAlignment actual = textRange.TestAccessor().Dynamic.GetHorizontalTextAlignment(style);
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData(ES.UPPERCASE | ES.LEFT | ES.MULTILINE | ES.READONLY | ES.AUTOHSCROLL, CapStyle.AllCap)]
        [InlineData(ES.LOWERCASE | ES.LEFT | ES.MULTILINE | ES.READONLY | ES.AUTOHSCROLL, CapStyle.None)]
        public void UiaTextRange_private_GetCapStyle_ReturnsExpectedValue(object editStyle, CapStyle expected)
        {
            using TextBox textBox = new TextBox();
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 0, 0);
            CapStyle actual = textRange.TestAccessor().Dynamic.GetCapStyle(editStyle);
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData(true)]
        [InlineData(false)]
        public void UiaTextRange_private_GetReadOnly_ReturnsCorrectValue(bool readOnly)
        {
            using TextBox textBox = new TextBox();
            textBox.ReadOnly = readOnly;
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 0, 0);
            bool actual = textRange.TestAccessor().Dynamic.GetReadOnly();
            Assert.Equal(readOnly, actual);
        }

        [WinFormsFact]
        public void UiaTextRange_private_GetBackgroundColor_ReturnsExpectedValue()
        {
            int actual = typeof(UiaTextRange).TestAccessor().Dynamic.GetBackgroundColor();
            int expected = 0x00ffffff; // WINDOW system color
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("Some test text")]
        public void UiaTextRange_private_GetFontName_ReturnsExpectedValue(string faceName)
        {
            LOGFONTW logfont = new LOGFONTW
            {
                FaceName = faceName
            };
            string actual = typeof(UiaTextRange).TestAccessor().Dynamic.GetFontName(logfont);
            Assert.Equal(faceName ?? "", actual);
        }

        [WinFormsTheory]
        [InlineData(1, 1)]
        [InlineData(5, 5)]
        [InlineData(5.3, 5)]
        [InlineData(9.5, 10)]
        [InlineData(18, 18)]
        [InlineData(18.8, 19)]
        [InlineData(100, 100)]
        public void UiaTextRange_private_GetFontSize_ReturnsCorrectValue(float fontSize, double expected)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "Some long long test text";
            textBox.Font = new Font("Arial", fontSize, FontStyle.Regular) { };
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 5, 20);
            double actual = textRange.TestAccessor().Dynamic.GetFontSize(provider.Logfont);
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData(FW.BLACK)]
        [InlineData(FW.BOLD)]
        [InlineData(FW.DEMIBOLD)]
        [InlineData(FW.DONTCARE)]
        [InlineData(FW.EXTRABOLD)]
        [InlineData(FW.EXTRALIGHT)]
        [InlineData(FW.LIGHT)]
        [InlineData(FW.MEDIUM)]
        [InlineData(FW.NORMAL)]
        [InlineData(FW.THIN)]
        public void UiaTextRange_private_GetFontWeight_ReturnsCorrectValue(object fontWeight)
        {
            LOGFONTW logfont = new LOGFONTW() { lfWeight = (FW)fontWeight };
            FW actual = typeof(UiaTextRange).TestAccessor().Dynamic.GetFontWeight(logfont);
            Assert.Equal(fontWeight, actual);
        }

        [WinFormsFact]
        public void UiaTextRange_private_GetForegroundColor_ReturnsCorrectValue()
        {
            int actual = typeof(UiaTextRange).TestAccessor().Dynamic.GetForegroundColor();
            Assert.Equal(0, actual);
        }

        [WinFormsTheory]
        [InlineData(0, false)]
        [InlineData(5, true)]
        public void UiaTextRange_private_GetItalic_ReturnsCorrectValue(byte ifItalic, bool expected)
        {
            LOGFONTW logfont = new LOGFONTW() { lfItalic = ifItalic };
            bool actual = typeof(UiaTextRange).TestAccessor().Dynamic.GetItalic(logfont);
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData(0, TextDecorationLineStyle.None)]
        [InlineData(5, TextDecorationLineStyle.Single)]
        public void UiaTextRange_private_GetStrikethroughStyle_ReturnsCorrectValue(byte ifStrikeOut, TextDecorationLineStyle expected)
        {
            LOGFONTW logfont = new LOGFONTW() { lfStrikeOut = ifStrikeOut };
            TextDecorationLineStyle actual = typeof(UiaTextRange).TestAccessor().Dynamic.GetStrikethroughStyle(logfont);
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData(0, TextDecorationLineStyle.None)]
        [InlineData(5, TextDecorationLineStyle.Single)]
        public void UiaTextRange_private_GetUnderlineStyle_ReturnsCorrectValue(byte ifUnderline, TextDecorationLineStyle expected)
        {
            LOGFONTW logfont = new LOGFONTW() { lfUnderline = ifUnderline };
            TextDecorationLineStyle actual = typeof(UiaTextRange).TestAccessor().Dynamic.GetUnderlineStyle(logfont);
            Assert.Equal(expected, actual);
        }

        [WinFormsTheory]
        [InlineData(0, 0)]
        [InlineData(0, 10)]
        [InlineData(5, 10)]
        [InlineData(5, 100)]
        [InlineData(100, 100)]
        [InlineData(100, 200)]
        public void UiaTextRange_private_MoveTo_SetValuesCorrectly(int start, int end)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "Some long long test text";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 0, 0);
            textRange.TestAccessor().Dynamic.MoveTo(start, end);
            Assert.Equal(start, textRange.Start);
            Assert.Equal(end, textRange.End);
        }

        [WinFormsTheory]
        [InlineData(-5, 0)]
        [InlineData(0, -5)]
        [InlineData(-10, -10)]
        [InlineData(10, 5)]
        public void UiaTextRange_private_MoveTo_ThrowsException_IfIncorrectParameters(int start, int end)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "Some long long test text";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, 0, 0);
            Assert.ThrowsAny<Exception>(() => textRange.TestAccessor().Dynamic.MoveTo(start, end));
        }

        [WinFormsTheory]
        [InlineData(0, 0, 0, 0)]
        [InlineData(0, 10, 0, 10)]
        [InlineData(5, 10, 5, 10)]
        [InlineData(5, 100, 5, 24)]
        [InlineData(100, 100, 24, 24)]
        [InlineData(100, 200, 24, 24)]
        public void UiaTextRange_private_ValidateEndpoints_SetValuesCorrectly(int start, int end, int expectedStart, int expectedEnd)
        {
            using TextBox textBox = new TextBox();
            textBox.Text = "Some long long test text";
            UiaTextProvider provider = GetProvider(textBox);
            UiaTextRange textRange = new UiaTextRange(textBox.AccessibilityObject, provider, start, end);
            textRange.TestAccessor().Dynamic.ValidateEndpoints();
            Assert.Equal(expectedStart, textRange.Start);
            Assert.Equal(expectedEnd, textRange.End);
        }
    }
}
