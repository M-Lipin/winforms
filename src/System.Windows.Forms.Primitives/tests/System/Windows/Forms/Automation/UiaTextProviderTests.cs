// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Drawing;
using System.Windows.Forms.Automation;
using Moq;
using Xunit;
using static Interop.User32;

namespace System.Windows.Forms.Primitives.Tests.Automation
{
    public class UiaTextProviderTests
    {
        [WinFormsTheory]
        [InlineData(0)]
        [InlineData(5)]
        [InlineData(15)]
        [InlineData(99)]
        public void UiaTextProvider_GetText_ReturnsControlText(int length)
        {
            using TextBoxBase textBoxBase = new TextBox();
            Assert.NotEqual(IntPtr.Zero, textBoxBase.Handle);

            string textBoxText = "This is the textbox text.";
            textBoxBase.Text = textBoxText;
            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);

            string text = mockUiaTextProvider.Object.GetText(textBoxBase.Handle, length).TrimEnd('\0');
            string expected = length > textBoxText.Length ? textBoxText : textBoxText.Substring(0, length);
            Assert.Equal(expected, text);
        }

        [WinFormsTheory]
        [InlineData(int.MaxValue)]
        [InlineData(int.MinValue)]
        [InlineData(-1)]
        [InlineData(-5)]
        public void UiaTextProvider_GetText_ThrowsExceptionIfLengthIsIncorrect(int length)
        {
            using TextBoxBase textBoxBase = new TextBox();
            Assert.NotEqual(IntPtr.Zero, textBoxBase.Handle);

            textBoxBase.Text = "This is the textbox text.";
            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            Assert.Throws<ArgumentException>(new Action(() => mockUiaTextProvider.Object.GetText(textBoxBase.Handle, length)));
        }

        [WinFormsTheory]
        [InlineData(0)]
        [InlineData(-1)]
        [InlineData(-2)]
        [InlineData(-3)]
        [InlineData(-235453)]
        public void UiaTextProvider_GetText_ReturnsUnexpectedStringIfHandleIsIncorrect(int handle)
        {
            using TextBoxBase textBoxBase = new TextBox();
            Assert.NotEqual(IntPtr.Zero, textBoxBase.Handle);

            textBoxBase.Text = "This is the textbox text.";
            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            IntPtr hwnd = (IntPtr)handle;
            string text = mockUiaTextProvider.Object.GetText(hwnd, textBoxBase.TextLength);
            Assert.NotEqual(text, textBoxBase.Text);
        }

        [WinFormsFact]
        public void UiaTextProvider_GetEditStyle_ReturnsMultilineStyle_ForMultilineTextBox()
        {
            using TextBoxBase textBoxBase = new TextBox();
            textBoxBase.Multiline = true;
            textBoxBase.CreateControl();

            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            ES editStyle = mockUiaTextProvider.Object.GetEditStyle(textBoxBase.Handle);

            Assert.True(editStyle.IsBitSet(ES.MULTILINE));
        }

        [WinFormsFact]
        public void UiaTextProvider_GetEditStyle_DoesntReturnMultilineStyle_ForSinglelineTextBox()
        {
            using TextBoxBase textBoxBase = new TextBox();
            textBoxBase.Multiline = false;
            textBoxBase.CreateControl();

            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            ES editStyle = mockUiaTextProvider.Object.GetEditStyle(textBoxBase.Handle);

            Assert.False(editStyle.IsBitSet(ES.MULTILINE));
        }

        [WinFormsFact]
        public void UiaTextProvider_GetWindowStyle_ReturnsNoneForNotInitializedControl()
        {
            using TextBoxBase textBoxBase = new TextBox();
            textBoxBase.CreateControl();

            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            WS windowStyle = mockUiaTextProvider.Object.GetWindowStyle(textBoxBase.Handle);

            Assert.True(windowStyle.IsBitSet(WS.VISIBLE));
        }

        [WinFormsFact]
        public void UiaTextProvider_GetWindowExStyle_ReturnsNoneForNotInitializedControl()
        {
            using TextBoxBase textBoxBase = new TextBox();
            textBoxBase.CreateControl();

            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            WS_EX windowStyle = mockUiaTextProvider.Object.GetWindowExStyle(textBoxBase.Handle);

            Assert.True(windowStyle.IsBitSet(WS_EX.CLIENTEDGE));
        }

        [WinFormsFact]
        public void UiaTextProvider_RectArrayToDoubleArray_ReturnsCorrectValue()
        {
            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            double[] result = mockUiaTextProvider.Object.RectArrayToDoubleArray(new Drawing.Rectangle[]
            {
                new Drawing.Rectangle(0, 0, 10, 5),
                new Drawing.Rectangle(10, 10, 20, 30)
            });

            Assert.Equal(8, result.Length);
            Assert.Equal(0, result[0]);
            Assert.Equal(0, result[1]);
            Assert.Equal(10, result[2]);
            Assert.Equal(5, result[3]);
            Assert.Equal(10, result[4]);
            Assert.Equal(10, result[5]);
            Assert.Equal(20, result[6]);
            Assert.Equal(30, result[7]);
        }

        [WinFormsFact]
        public void UiaTextProvider_RectArrayToDoubleArray_NullParameter_ReturnsNull()
        {
            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            double[] result = mockUiaTextProvider.Object.RectArrayToDoubleArray(null);
            Assert.Null(result);
        }

        [WinFormsFact]
        public void UiaTextProvider_RectArrayToDoubleArray_EmptyArrayParameter_ReturnsEmptyArrayResult()
        {
            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            double[] result = mockUiaTextProvider.Object.RectArrayToDoubleArray(new Rectangle[0]);
            Assert.Empty(result);
        }

        [WinFormsFact]
        public unsafe void UiaTextProvider_SendInput_SendsOneInput()
        {
            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            INPUT keyboardInput = new INPUT();

            int result = mockUiaTextProvider.Object.SendInput(1, ref keyboardInput, sizeof(INPUT));

            Assert.Equal(1, result);
        }

        [WinFormsFact]
        public void UiaTextProvider_SendKeyboardInputVK_SendsOneInput()
        {
            var mockUiaTextProvider = new Mock<UiaTextProvider>(MockBehavior.Strict);
            int result = mockUiaTextProvider.Object.SendKeyboardInputVK(VK.LEFT, true);

            Assert.Equal(1, result);
        }
    }
}
