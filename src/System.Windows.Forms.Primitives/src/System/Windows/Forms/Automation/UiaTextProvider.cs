// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms.Primitives.Resources;
using static Interop.UiaCore;
using static Interop.User32;

namespace System.Windows.Forms.Automation
{
    internal abstract class UiaTextProvider : ITextProvider
    {
        public abstract ITextRangeProvider[] GetSelection();

        public abstract ITextRangeProvider[] GetVisibleRanges();

        public abstract ITextRangeProvider RangeFromChild(IRawElementProviderSimple childElement);

        public abstract ITextRangeProvider RangeFromPoint(Point screenLocation);

        public abstract ITextRangeProvider DocumentRange { get; }

        public abstract SupportedTextSelection SupportedTextSelection { get; }

        public abstract Rectangle BoundingRectangle { get; }

        public abstract ES EditStyle { get; }

        public abstract int FirstVisibleLine { get; }

        public abstract bool IsMultiline { get; }

        public abstract bool IsReadingRTL { get; }

        public abstract bool IsReadOnly { get; }

        public abstract bool IsScrollable { get; }

        public abstract int LinesPerPage { get; }

        public abstract int LinesCount { get; }

        public abstract LOGFONTW Logfont { get; }

        public abstract string Text { get; }

        public abstract int TextLength { get; }

        public abstract WS_EX WindowExStyle { get; }

        public abstract WS WindowStyle { get; }

        public abstract int GetLineFromCharIndex(int charIndex);

        public abstract int GetLineIndex(int line);

        public abstract Point GetPositionFromChar(int charIndex);

        public abstract Point GetPositionFromCharUR(int startCharIndex, string text);

        public abstract void GetVisibleRangePoints(out int visibleStart, out int visibleEnd);

        public abstract bool LineScroll(int charactersHorizontal, int linesVertical);

        public abstract int ReleaseDC(IntPtr hdc);

        public abstract void SetSelection(int start, int end);

        public string GetText(IntPtr hWnd, int length)
        {
            if (length < 0 || length > int.MaxValue - 1 /* 1 for terminating null */ )
            {
                throw new ArgumentException(string.Format(SR.UiaTextProviderTextLengthInvalidExceptionMessageTemplate, int.MaxValue));
            }

            SendMessageW(hWnd, WM.GETTEXT, (IntPtr)length + 1, out string text);
            return text;
        }

        public ES GetEditStyle(IntPtr hWnd)
        {
            return (ES)GetWindowLong(new HandleRef(null, hWnd), GWL.STYLE);
        }

        public WS_EX GetWindowExStyle(IntPtr hWnd)
        {
            return (WS_EX)GetWindowLong(new HandleRef(null, hWnd), GWL.EXSTYLE);
        }

        public WS GetWindowStyle(IntPtr hWnd)
        {
            return (WS)GetWindowLong(new HandleRef(this, hWnd), GWL.STYLE);
        }

        public double[] RectArrayToDoubleArray(Rectangle[] rectArray)
        {
            if (rectArray == null)
            {
                return null;
            }

            double[] doubles = new double[rectArray.Length * 4];
            int scan = 0;

            for (int i = 0; i < rectArray.Length; i++)
            {
                doubles[scan++] = rectArray[i].X;
                doubles[scan++] = rectArray[i].Y;
                doubles[scan++] = rectArray[i].Width;
                doubles[scan++] = rectArray[i].Height;
            }

            return doubles;
        }

        public int SendInput(int inputs, ref INPUT input, int size)
        {
            Span<INPUT> currentInput = stackalloc INPUT[1];
            currentInput[0] = input;

            return (int)Interop.User32.SendInput((uint)inputs, currentInput, size);
        }

        public unsafe int SendKeyboardInputVK(short vk, bool press)
        {
            INPUT keyboardInput = new INPUT();

            keyboardInput.type = INPUTENUM.KEYBOARD;
            keyboardInput.inputUnion.ki.wVk = (ushort)vk;
            keyboardInput.inputUnion.ki.wScan = 0;
            keyboardInput.inputUnion.ki.dwFlags = press ? 0 : KEYEVENTF.KEYUP;

            if (IsExtendedKey(vk))
            {
                keyboardInput.inputUnion.ki.dwFlags |= KEYEVENTF.EXTENDEDKEY;
            }

            keyboardInput.inputUnion.ki.time = 0;
            keyboardInput.inputUnion.ki.dwExtraInfo = new IntPtr(0);

            return SendInput(1, ref keyboardInput, sizeof(INPUT));
        }

        private static bool IsExtendedKey(short vk)
        {
            // From the SDK:
            // The extended-key flag indicates whether the keystroke message originated from one of
            // the additional keys on the enhanced keyboard. The extended keys consist of the ALT and
            // CTRL keys on the right-hand side of the keyboard; the INS, DEL, HOME, END, PAGE UP,
            // PAGE DOWN, and arrow keys in the clusters to the left of the numeric keypad; the NUM LOCK
            // key; the BREAK (CTRL+PAUSE) key; the PRINT SCRN key; and the divide (/) and ENTER keys in
            // the numeric keypad. The extended-key flag is set if the key is an extended key.
            //
            // - docs appear to be incorrect. Use of Spy++ indicates that break is not an extended key.
            // Also, menu key and windows keys also appear to be extended.
            return vk == unchecked((short)VK.RMENU) ||
                   vk == unchecked((short)VK.RCONTROL) ||
                   vk == unchecked((short)VK.NUMLOCK) ||
                   vk == unchecked((short)VK.INSERT) ||
                   vk == unchecked((short)VK.DELETE) ||
                   vk == unchecked((short)VK.HOME) ||
                   vk == unchecked((short)VK.END) ||
                   vk == unchecked((short)VK.PRIOR) ||
                   vk == unchecked((short)VK.NEXT) ||
                   vk == unchecked((short)VK.UP) ||
                   vk == unchecked((short)VK.DOWN) ||
                   vk == unchecked((short)VK.LEFT) ||
                   vk == unchecked((short)VK.RIGHT) ||
                   vk == unchecked((short)VK.APPS) ||
                   vk == unchecked((short)VK.RWIN) ||
                   vk == unchecked((short)VK.LWIN);
            // Note that there are no distinct values for the following keys:
            // numpad divide
            // numpad enter
        }
    }
}
