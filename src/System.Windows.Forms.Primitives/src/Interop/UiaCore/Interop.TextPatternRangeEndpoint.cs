// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Runtime.InteropServices;

internal static partial class Interop
{
    internal static partial class UiaCore
    {
        [Guid("7D54FF7A-5B63-4EFA-86C3-112A948ED93E")]
        public enum TextPatternRangeEndpoint
        {
            Start = 0,
            End = 1
        }
    }
}
