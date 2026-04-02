using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
    private FrameworkElement BuildTimeTagCard(BoardItemModel item)
    {
        var dueText = item.TimeTagDueAt.HasValue
            ? item.TimeTagDueAt.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
            : "未设置";

        var recurrenceText = FormatTimeTagRecurrence(item.TimeTagRecurrence, item.TimeTagMonthlyDays);

        var container = new Border
        {
            Background = new SolidColorBrush(ColorHelper.FromArgb(210, 20, 28, 38)),
            Padding = new Thickness(12),
            Child = new StackPanel
            {
                Spacing = 8,
                Children =
                {
                    new TextBlock
                    {
                        Text = string.IsNullOrWhiteSpace(item.Content) ? "时间标签" : item.Content,
                        Foreground = ItemTitleBrush,
                        FontWeight = SemiBoldWeight,
                        FontSize = 15,
                        TextWrapping = TextWrapping.WrapWholeWords
                    },
                    new TextBlock
                    {
                        Text = $"到期时间：{dueText}",
                        Foreground = ItemTextBrush,
                        TextWrapping = TextWrapping.WrapWholeWords
                    },
                    new TextBlock
                    {
                        Text = $"提醒：{(item.TimeTagReminderEnabled ? "开启" : "关闭")}",
                        Foreground = SecondaryTextBrush
                    },
                    new TextBlock
                    {
                        Text = $"循环：{recurrenceText}",
                        Foreground = SecondaryTextBrush,
                        TextWrapping = TextWrapping.WrapWholeWords
                    }
                }
            }
        };

        if (container.Child is StackPanel panel)
        {
            var editButton = new Button
            {
                HorizontalAlignment = HorizontalAlignment.Left,
                Content = "设置时间标签"
            };
            editButton.Click += async (_, _) => await ShowTimeTagEditorAsync(item);
            panel.Children.Add(editButton);
        }

        return container;
    }

    private async Task ShowTimeTagEditorAsync(BoardItemModel item)
    {
        var due = item.TimeTagDueAt?.ToLocalTime() ?? DateTimeOffset.Now.AddHours(1);
        var localOffset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);

        var titleBox = new TextBox
        {
            Header = "标签标题",
            Text = string.IsNullOrWhiteSpace(item.Content) ? "时间标签" : item.Content
        };

        var datePicker = new CalendarDatePicker
        {
            Header = "日期",
            Language = "zh-Hans-CN",
            DateFormat = "{year.full}-{month.integer(2)}-{day.integer(2)}",
            Date = new DateTimeOffset(due.Year, due.Month, due.Day, 0, 0, 0, localOffset)
        };

        var timePicker = new TimePicker
        {
            Header = "时间",
            ClockIdentifier = "24HourClock",
            Time = due.TimeOfDay
        };

        var reminderSwitch = new ToggleSwitch
        {
            Header = "到期提醒",
            IsOn = item.TimeTagReminderEnabled
        };

        var monday = new CheckBox { Content = "周一", IsChecked = item.TimeTagRecurrence.HasFlag(TimeTagRecurrence.Monday) };
        var tuesday = new CheckBox { Content = "周二", IsChecked = item.TimeTagRecurrence.HasFlag(TimeTagRecurrence.Tuesday) };
        var wednesday = new CheckBox { Content = "周三", IsChecked = item.TimeTagRecurrence.HasFlag(TimeTagRecurrence.Wednesday) };
        var thursday = new CheckBox { Content = "周四", IsChecked = item.TimeTagRecurrence.HasFlag(TimeTagRecurrence.Thursday) };
        var friday = new CheckBox { Content = "周五", IsChecked = item.TimeTagRecurrence.HasFlag(TimeTagRecurrence.Friday) };
        var saturday = new CheckBox { Content = "周六", IsChecked = item.TimeTagRecurrence.HasFlag(TimeTagRecurrence.Saturday) };
        var sunday = new CheckBox { Content = "周日", IsChecked = item.TimeTagRecurrence.HasFlag(TimeTagRecurrence.Sunday) };
        var monthlyDaysBox = new TextBox
        {
            Header = "每月日期（支持 1-5,8,15）",
            PlaceholderText = "例如：1-5,8,15",
            Text = item.TimeTagMonthlyDays ?? string.Empty
        };

        var recurrencePanel = new StackPanel
        {
            Spacing = 6,
            Children =
            {
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 8,
                    Children = { monday, tuesday, wednesday, thursday }
                },
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 8,
                    Children = { friday, saturday, sunday }
                }
            }
        };

        var dialog = new ContentDialog
        {
            Title = "设置时间标签",
            PrimaryButtonText = "保存",
            CloseButtonText = "取消",
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = RootGrid.XamlRoot,
            Content = new ScrollViewer
            {
                MaxHeight = 460,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = new StackPanel
                {
                    Width = 340,
                    Spacing = 10,
                    Children =
                    {
                        titleBox,
                        datePicker,
                        timePicker,
                        reminderSwitch,
                        new TextBlock { Text = "按周循环（可多选）", Foreground = SecondaryTextBrush },
                        recurrencePanel,
                        monthlyDaysBox
                    }
                }
            }
        };

        var result = await dialog.ShowAsync();
        if (result != ContentDialogResult.Primary)
        {
            return;
        }

        var recurrence = TimeTagRecurrence.None;
        if (monday.IsChecked == true) recurrence |= TimeTagRecurrence.Monday;
        if (tuesday.IsChecked == true) recurrence |= TimeTagRecurrence.Tuesday;
        if (wednesday.IsChecked == true) recurrence |= TimeTagRecurrence.Wednesday;
        if (thursday.IsChecked == true) recurrence |= TimeTagRecurrence.Thursday;
        if (friday.IsChecked == true) recurrence |= TimeTagRecurrence.Friday;
        if (saturday.IsChecked == true) recurrence |= TimeTagRecurrence.Saturday;
        if (sunday.IsChecked == true) recurrence |= TimeTagRecurrence.Sunday;

        var pickedDate = datePicker.Date ?? new DateTimeOffset(due.Year, due.Month, due.Day, 0, 0, 0, localOffset);
        var pickedTime = timePicker.Time;
        var dueAt = new DateTimeOffset(
            pickedDate.Year,
            pickedDate.Month,
            pickedDate.Day,
            pickedTime.Hours,
            pickedTime.Minutes,
            0,
            TimeZoneInfo.Local.GetUtcOffset(DateTime.Now));

        item.Content = string.IsNullOrWhiteSpace(titleBox.Text) ? "时间标签" : titleBox.Text.Trim();
        item.TimeTagDueAt = dueAt;
        item.TimeTagReminderEnabled = reminderSwitch.IsOn;
        item.TimeTagRecurrence = recurrence;
        item.TimeTagMonthlyDays = string.IsNullOrWhiteSpace(monthlyDaysBox.Text) ? null : monthlyDaysBox.Text.Trim();
        item.TimeTagLastReminderAt = null;
        RefreshBoardItemView(item);
    }

    private static string FormatTimeTagRecurrence(TimeTagRecurrence recurrence, string? monthlyDays)
    {
        if (recurrence == TimeTagRecurrence.None && string.IsNullOrWhiteSpace(monthlyDays))
        {
            return "不循环";
        }

        var labels = new List<string>();
        if (recurrence.HasFlag(TimeTagRecurrence.Monday)) labels.Add("周一");
        if (recurrence.HasFlag(TimeTagRecurrence.Tuesday)) labels.Add("周二");
        if (recurrence.HasFlag(TimeTagRecurrence.Wednesday)) labels.Add("周三");
        if (recurrence.HasFlag(TimeTagRecurrence.Thursday)) labels.Add("周四");
        if (recurrence.HasFlag(TimeTagRecurrence.Friday)) labels.Add("周五");
        if (recurrence.HasFlag(TimeTagRecurrence.Saturday)) labels.Add("周六");
        if (recurrence.HasFlag(TimeTagRecurrence.Sunday)) labels.Add("周日");
        if (!string.IsNullOrWhiteSpace(monthlyDays)) labels.Add($"每月 {monthlyDays}");
        return string.Join(" / ", labels);
    }
}
