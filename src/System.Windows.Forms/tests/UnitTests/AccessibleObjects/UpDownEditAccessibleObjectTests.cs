// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using Xunit;
using static Interop.UiaCore;

namespace System.Windows.Forms.Tests.AccessibleObjects
{
    public class UpDownEditAccessibleObjectTests
    {
        [WinFormsFact]
        public void UpDownEditAccessibleObject_ctor_default()
        {
            using UpDownBase upDown = new SubUpDownBase();
            using UpDownBase.UpDownEdit upDownEdit = new UpDownBase.UpDownEdit(upDown);
            Assert.NotNull(upDownEdit.AccessibilityObject);
        }

        [WinFormsTheory]
        [InlineData((int)UIA.IsTextPatternAvailablePropertyId)]
        [InlineData((int)UIA.IsTextPattern2AvailablePropertyId)]
        public void TextBoxBaseAccessibleObject_TextPatternAvailable(int propertyId)
        {
            using UpDownBase upDown = new SubUpDownBase();
            using UpDownBase.UpDownEdit upDownEdit = new UpDownBase.UpDownEdit(upDown);
            AccessibleObject textBoxAccessibleObject = upDownEdit.AccessibilityObject;

            // Interop.UiaCore.UIA accessible level (internal) is less than the test level (public) so it needs boxing and unboxing
            Assert.True((bool)textBoxAccessibleObject.GetPropertyValue((UIA)propertyId));
        }

        [WinFormsTheory]
        [InlineData((int)UIA.TextPatternId)]
        [InlineData((int)UIA.TextPattern2Id)]
        public void TextBoxBaseAccessibleObject_TextPatternSupported(int patternId)
        {
            using UpDownBase upDown = new SubUpDownBase();
            using UpDownBase.UpDownEdit upDownEdit = new UpDownBase.UpDownEdit(upDown);
            AccessibleObject textBoxAccessibleObject = upDownEdit.AccessibilityObject;

            // Interop.UiaCore.UIA accessible level (internal) is less than the test level (public) so it needs boxing and unboxing
            Assert.True(textBoxAccessibleObject.IsPatternSupported((UIA)patternId));
        }

        private class SubUpDownBase : UpDownBase
        {
            public override void DownButton() { }

            public override void UpButton() { }

            protected override void UpdateEditText() { }
        }
    }
}
