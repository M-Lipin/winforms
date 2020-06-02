﻿// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Runtime.InteropServices;

internal static partial class Interop
{
    internal static partial class UiaCore
    {
        [Guid("1DC6C721-5F21-4AAB-845A-7F8E714AD322")]
        public enum TextUnit
        {
            Character = 0,
            Format = 1,
            Word = 2,
            Line = 3,
            Paragraph = 4,
            Page = 5,
            Document = 6
        }
    }
}
