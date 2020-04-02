// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

#nullable disable

using static Interop;
using static Interop.ComCtl32;

namespace System.Windows.Forms
{
    public partial class MonthCalendar
    {
        /// <summary>
        /// Represents the calendar body accessible object.
        /// </summary>
        internal class CalendarBodyAccessibleObject : CalendarChildAccessibleObject
        {
            private const int ChildId = 4; // TODO: This should be adjusted for multicalendar calendar control.

            public CalendarBodyAccessibleObject(MonthCalendarAccessibleObject calendarAccessibleObject, int calendarIndex)
                : base(calendarAccessibleObject, calendarIndex, CalendarChildType.CalendarBody)
            {
            }

            protected override RECT CalculateBoundingRectangle()
            {
                _calendarAccessibleObject.GetCalendarPartRectangle(_calendarIndex, MCGIP.CALENDARBODY, 0, 0, out RECT calendarPartRectangle);
                return calendarPartRectangle;
            }

            internal override int ColumnCount
            {
                get
                {
                    _calendarAccessibleObject.GetCalendarGridInfo(
                        MCGIF.RECT,
                        MCGIP.CALENDARBODY,
                        _calendarIndex,
                        -1,
                        -1,
                        out RECT calendarBodyRectangle,
                        out Kernel32.SYSTEMTIME endDate,
                        out Kernel32.SYSTEMTIME startDate);

                    int columnCount = 0;
                    bool success = true;
                    while (success)
                    {
                        success = _calendarAccessibleObject.GetCalendarGridInfo(
                            MCGIF.RECT,
                            MCGIP.CALENDARCELL,
                            _calendarIndex,
                            0,
                            columnCount,
                            out RECT calendarPartRectangle,
                            out endDate,
                            out startDate);

                        // Out of the body, so this is out of the grid column.
                        if (calendarPartRectangle.right > calendarBodyRectangle.right)
                        {
                            break;
                        }

                        columnCount++;
                    }

                    return columnCount;
                }
            }

            internal override int RowCount
            {
                get
                {
                    _calendarAccessibleObject.GetCalendarGridInfo(
                        MCGIF.RECT,
                        MCGIP.CALENDARBODY,
                        _calendarIndex,
                        -1,
                        -1,
                        out RECT calendarBodyRectangle,
                        out Kernel32.SYSTEMTIME endDate,
                        out Kernel32.SYSTEMTIME startDate);

                    int rowCount = 0;
                    bool success = true;
                    while (success)
                    {
                        success = _calendarAccessibleObject.GetCalendarGridInfo(
                            MCGIF.RECT,
                            MCGIP.CALENDARCELL,
                            _calendarIndex,
                            rowCount,
                            0,
                            out RECT calendarPartRectangle,
                            out endDate,
                            out startDate);

                        // Out of the body, so this is out of the grid row.
                        if (calendarPartRectangle.bottom > calendarBodyRectangle.bottom)
                        {
                            break;
                        }

                        rowCount++;
                    }

                    return rowCount;
                }
            }

            internal override int GetChildId() => ChildId;

            public bool HasHeaderRow
            {
                get
                {
                    bool result = _calendarAccessibleObject.GetCalendarGridInfoText(MCGIP.CALENDARCELL, _calendarIndex, -1, 0, out string text);
                    if (!result || string.IsNullOrEmpty(text))
                    {
                        return false;
                    }

                    return true;
                }
            }

            internal override UiaCore.IRawElementProviderFragment FragmentNavigate(UiaCore.NavigateDirection direction) =>
                direction switch
                {
                    UiaCore.NavigateDirection.NextSibling => new Func<AccessibleObject>(() =>
                    {
                        MonthCalendar owner = (MonthCalendar)_calendarAccessibleObject.Owner;
                        return owner.ShowToday ? _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.TodayLink) : null;
                    })(),
                    UiaCore.NavigateDirection.PreviousSibling => _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarHeader),
                    UiaCore.NavigateDirection.FirstChild => _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarRow, this, HasHeaderRow ? -1 : 0),
                    UiaCore.NavigateDirection.LastChild => _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarRow, this, _calendarAccessibleObject.RowCount - 1),
                    _ => base.FragmentNavigate(direction),

                };

            public CalendarChildAccessibleObject GetFromPoint(MCHITTESTINFO hitTestInfo)
            {
                switch ((MCHT)hitTestInfo.uHit)
                {
                    case MCHT.CALENDARDAY:
                    case MCHT.CALENDARWEEKNUM:
                    case MCHT.CALENDARDATE:
                        AccessibleObject rowAccessibleObject = _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarRow, this, hitTestInfo.iRow);
                        return _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarCell, rowAccessibleObject, hitTestInfo.iCol);
                }

                return this;
            }

            internal override object GetPropertyValue(UiaCore.UIA propertyID) =>
                propertyID switch
                {
                    UiaCore.UIA.NamePropertyId => SR.MonthCalendarBodyAccessibleName,
                    UiaCore.UIA.IsGridPatternAvailablePropertyId => true,
                    UiaCore.UIA.IsTablePatternAvailablePropertyId => true,
                    _ => base.GetPropertyValue(propertyID)
                };

            internal override UiaCore.IRawElementProviderSimple[] GetRowHeaders() => null;

            internal override UiaCore.IRawElementProviderSimple[] GetColumnHeaderItems()
            {
                if (!HasHeaderRow)
                {
                    return null;
                }

                UiaCore.IRawElementProviderSimple[] headers =
                    new UiaCore.IRawElementProviderSimple[MonthCalendarAccessibleObject.MAX_DAYS];
                AccessibleObject headerRowAccessibleObject = _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarRow, this, -1);
                for (int columnIndex = 0; columnIndex < MonthCalendarAccessibleObject.MAX_DAYS; columnIndex++)
                {
                    headers[columnIndex] = _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarCell, headerRowAccessibleObject, columnIndex);
                }

                return headers;
            }

            internal override UiaCore.IRawElementProviderSimple GetItem(int row, int column)
            {
                AccessibleObject rowAccessibleObject = _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarRow, this, row);

                if (rowAccessibleObject == null)
                {
                    return null;
                }

                return _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarCell, rowAccessibleObject, column);
            }

            private AccessibleObject GetCalendarChildAccessibleObject(DateTime selectionStart, DateTime selectionEnd)
            {
                int columnCount = ColumnCount;

                AccessibleObject bodyAccessibleObject = _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarBody);
                for (int row = 0; row < RowCount; row++)
                {
                    AccessibleObject rowAccessibleObject = _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarRow, bodyAccessibleObject, row);
                    for (int column = 0; column < columnCount; column++)
                    {
                        bool success = _calendarAccessibleObject.GetCalendarGridInfo(
                            MCGIF.DATE,
                            MCGIP.CALENDARCELL,
                            _calendarIndex,
                            row,
                            column,
                            out RECT calendarPartRectangle,
                            out Kernel32.SYSTEMTIME systemEndDate,
                            out Kernel32.SYSTEMTIME systemStartDate);

                        if (!success)
                        {
                            continue;
                        }

                        AccessibleObject cellAccessibleObject = _calendarAccessibleObject.GetCalendarChildAccessibleObject(_calendarIndex, CalendarChildType.CalendarCell, rowAccessibleObject, column);
                        if (cellAccessibleObject == null)
                        {
                            continue;
                        }

                        DateTime endDate = DateTimePicker.SysTimeToDateTime(systemEndDate);
                        DateTime startDate = DateTimePicker.SysTimeToDateTime(systemStartDate);

                        if (DateTime.Compare(selectionEnd, endDate) <= 0 &&
                            DateTime.Compare(selectionStart, startDate) >= 0)
                        {
                            return cellAccessibleObject;
                        }
                    }
                }

                return null;
            }
        }
    }
}
