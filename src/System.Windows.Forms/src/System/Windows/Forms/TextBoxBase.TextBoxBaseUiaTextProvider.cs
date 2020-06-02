// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms.Automation;
using static Interop;

namespace System.Windows.Forms
{
    public abstract partial class TextBoxBase
    {
        internal class TextBoxBaseUiaTextProvider : UiaTextProvider2
        {
            private readonly TextBoxBase _owner;

            public TextBoxBaseUiaTextProvider(TextBoxBase owner)
            {
                _owner = owner;
            }

            public override UiaCore.ITextRangeProvider[] GetSelection()
            {
                _owner.GetSelection(out int start, out int end);

                var internalAccessibleObject = new InternalAccessibleObject(_owner.AccessibilityObject);
                return new UiaCore.ITextRangeProvider[] { new UiaTextRange(internalAccessibleObject, this, start, end) };
            }

            public override UiaCore.ITextRangeProvider[] GetVisibleRanges()
            {
                GetVisibleRangePoints(out int start, out int end);

                var internalAccessibleObject = new InternalAccessibleObject(_owner.AccessibilityObject);
                return new UiaCore.ITextRangeProvider[] { new UiaTextRange(internalAccessibleObject, this, start, end) };
            }

            public override UiaCore.ITextRangeProvider RangeFromChild(UiaCore.IRawElementProviderSimple childElement)
            {
                // We don't have any children so this call returns null.
                Debug.Fail("Text edit control cannot have a child element.");
                return null;
            }

            /// <summary>
            /// Returns the degenerate (empty) text range nearest to the specified screen coordinates.
            /// </summary>
            /// <param name="screenlocation">The location in screen coordinates.</param>
            /// <returns>A degenerate range nearest the specified location. Null is never returned.</returns>
            public override UiaCore.ITextRangeProvider RangeFromPoint(Point screenlocation)
            {
                Point clientLocation = screenlocation;

                // Convert screen to client coordinates.
                // (Essentially ScreenToClient but MapWindowPoints accounts for window mirroring using WS_EX_LAYOUTRTL.)
                if (User32.MapWindowPoints(new HandleRef(null, IntPtr.Zero), new HandleRef(this, _owner.Handle), ref clientLocation, 1) == 0)
                {
                    return new UiaTextRange(new InternalAccessibleObject(_owner.AccessibilityObject), this, 0, 0);
                }

                // We have to deal with the possibility that the coordinate is inside the window rect
                // but outside the client rect. In that case we just scoot it over so it is at the nearest
                // point in the client rect.
                RECT clientRectangle = _owner.ClientRectangle;

                clientLocation.X = Math.Max(clientLocation.X, clientRectangle.left);
                clientLocation.X = Math.Min(clientLocation.X, clientRectangle.right);
                clientLocation.Y = Math.Max(clientLocation.Y, clientRectangle.top);
                clientLocation.Y = Math.Min(clientLocation.Y, clientRectangle.bottom);

                // Get the character at those client coordinates.
                int start = _owner.GetCharIndexFromPosition(clientLocation);

                return new UiaTextRange(new InternalAccessibleObject(_owner.AccessibilityObject), this, start, start);
            }

            public override UiaCore.ITextRangeProvider DocumentRange => new UiaTextRange(new InternalAccessibleObject(_owner.AccessibilityObject), this, 0, TextLength);

            public override UiaCore.SupportedTextSelection SupportedTextSelection => UiaCore.SupportedTextSelection.Single;

            public override UiaCore.ITextRangeProvider GetCaretRange(out BOOL isActive)
            {
                isActive = BOOL.FALSE;

                var hasKeyboardFocus = OwningControl.AccessibilityObject.GetPropertyValue(UiaCore.UIA.HasKeyboardFocusPropertyId);
                if (hasKeyboardFocus is bool && (bool)hasKeyboardFocus)
                {
                    isActive = BOOL.TRUE;
                }

                var internalAccessibleObject = new InternalAccessibleObject(_owner.AccessibilityObject);
                return new UiaTextRange(internalAccessibleObject, this, _owner.SelectionStart, _owner.SelectionStart);
            }

            /// <summary>
            /// Exposes a text range that contains the text that is the target of the annotation associated with the specified annotation element.
            /// </summary>
            /// <param name="annotationElement">
            /// The provider for an element that implements the IAnnotationProvider interface.
            /// The annotation element is a sibling of the element that implements the <see cref="Interop.UiaCore.ITextProvider2"/> interface for the document.
            /// </param>
            /// <returns>
            /// A text range that contains the annotation target text.
            /// </returns>
            public override UiaCore.ITextRangeProvider RangeFromAnnotation(UiaCore.IRawElementProviderSimple annotationElement)
            {
                var internalAccessibleObject = new InternalAccessibleObject(_owner.AccessibilityObject);
                return new UiaTextRange(internalAccessibleObject, this, 0, 0);
            }

            public override Rectangle BoundingRectangle => _owner.GetRectangle();

            public override int FirstVisibleLine => _owner.GetFirstVisibleLine();

            public override bool IsMultiline => _owner.Multiline;

            public override bool IsReadingRTL =>
                (_owner != null && _owner.IsHandleCreated) ? WindowExStyle.IsBitSet(User32.WS_EX.RTLREADING) : false;

            public override bool IsReadOnly => _owner.ReadOnly;

            public override bool IsScrollable => _owner.Scrollable;

            public override int LinesCount =>
                unchecked((int)(long)User32.SendMessageW(new HandleRef(this, _owner.Handle), (User32.WM)User32.EM.GETLINECOUNT));

            public override int LinesPerPage => _owner.LinesPerPage;

            public override User32.LOGFONTW Logfont => _owner.GetLogfont();

            public Control OwningControl => _owner;

            public override string Text =>
                (_owner != null && _owner.IsHandleCreated) ? GetText(_owner.Handle, _owner.TextLength) ?? string.Empty : string.Empty;

            public override int TextLength => _owner.GetTextLength();

            public override User32.WS_EX WindowExStyle =>
                (_owner != null && _owner.IsHandleCreated) ? GetWindowExStyle(_owner.Handle) : User32.WS_EX.LEFT;

            public override User32.WS WindowStyle =>
                (_owner != null && _owner.IsHandleCreated) ? GetWindowStyle(_owner.Handle) : User32.WS.OVERLAPPED;

            public override User32.ES EditStyle =>
                (_owner != null && _owner.IsHandleCreated) ? GetEditStyle(_owner.Handle) : User32.ES.LEFT;

            public override int GetLineFromCharIndex(int charIndex) => _owner.GetLineFromCharIndex(charIndex);

            public override int GetLineIndex(int line) => _owner.GetLineIndex(line);

            public override Point GetPositionFromChar(int charIndex) => _owner.GetPositionFromCharIndex(charIndex);

            public override Point GetPositionFromCharUR(int startCharIndex, string text) => _owner.GetPositionFromCharUR(startCharIndex, text);

            public override void GetVisibleRangePoints(out int visibleStart, out int visibleEnd) => _owner.GetVisibleRangePoints(out visibleStart, out visibleEnd);

            public override bool LineScroll(int charactersHorizontal, int linesVertical) => _owner.LineScroll(charactersHorizontal, linesVertical);

            public override int ReleaseDC(IntPtr hdc)
            {
                return User32.ReleaseDC(_owner.Handle, hdc);
            }

            public override void SetSelection(int start, int end)
            {
                if (start < 0 || start > TextLength)
                {
                    Debug.Fail("SetSelection start is out of text range.");
                    return;
                }

                if (end < 0 || end > TextLength)
                {
                    Debug.Fail("SetSelection end is out of text range.");
                    return;
                }

                User32.SendMessageW(new HandleRef(this, _owner.Handle), (User32.WM)User32.EM.SETSEL, (IntPtr)start, (IntPtr)end);
            }
        }
    }
}
