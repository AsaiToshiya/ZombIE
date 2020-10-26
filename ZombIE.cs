/**
 * ZombIE 1.0.0
 * Copyright (c) 2020 Asai Toshiya.
 *
 * License: The MIT License (https://opensource.org/licenses/MIT)
 */

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;

namespace YourNamespace
{
    public class ZombIE
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        private static readonly string DefaultDownloadDirectory = GetDefaultDownloadDirectory();

        private static readonly Regex DialogDownloadFileNameRegex = new Regex(@"(.*?) で行う操作を選んでください。");

        private static readonly Regex NotificationBarDownloadFileNameRegex = new Regex(@".*? から (.*?) (\(.*?\) )?を開くか、または保存しますか\?");

        public static TimeSpan DownloadTimeout { get; set; } = TimeSpan.Zero;

        public static TimeSpan FindElementTimeout { get; set; } = TimeSpan.FromSeconds(30);

        public static void DownloadFileTo(string path)
        {
            var downloadFileName = Timeout<string>(
                () => FromDialog() ?? FromNoticationBar(),
                (x) => x != null,
                FindElementTimeout
            );

            var downloadedFileName = WaitUntilDownloadCompleted(downloadFileName);
            var downloadedFile = Path.Combine(DefaultDownloadDirectory, downloadedFileName);

            File.Move(downloadedFile, path);
        }

        private static void Click(AutomationElement ae)
        {
            (ae.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern).Invoke();
        }

        private static AutomationElement FindElementByName(AutomationElement parent, string name)
        {
            return parent.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, name));
        }

        private static AutomationElement FindNotificationBar(IntPtr hWnd)
        {
            var hWndNotificationBar = FindWindowEx(hWnd, IntPtr.Zero, "Frame Notification Bar", string.Empty);
            if (hWnd == IntPtr.Zero)
            {
                return null;
            }
            return AutomationElement.FromHandle(hWndNotificationBar);
        }

        private static AutomationElement FindText(AutomationElement parent)
        {
            return parent.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
        }

        private static string FromDialog()
        {
            var hWnd = FindWindow(null, "Internet Explorer");
            if (hWnd == IntPtr.Zero)
            {
                return null;
            }

            var window = AutomationElement.FromHandle(hWnd);

            var text = FindText(window);
            if (text == null)
            {
                return null;
            }

            var name = text.Current.Name;
            var match = DialogDownloadFileNameRegex.Match(name);
            if (!match.Success)
            {
                return null;
            }

            var saveButton = FindElementByName(window, "保存(S)");
            if (saveButton == null)
            {
                return null;
            }

            var fileName = match.Groups[1].Value;
            Click(saveButton);

            return fileName;
        }

        private static string FromNoticationBar()
        {
            var processes = Process.GetProcessesByName("iexplore");
            foreach (var process in processes)
            {
                var hWnd = process.MainWindowHandle;

                var notificationBar = FindNotificationBar(hWnd);
                if (notificationBar == null)
                {
                    continue;
                }

                var text = FindText(notificationBar);
                if (text == null)
                {
                    continue;
                }

                var textValue = GetValue(text);
                var match = NotificationBarDownloadFileNameRegex.Match(textValue);
                if (!match.Success)
                {
                    continue;
                }

                var saveButton = FindElementByName(notificationBar, "保存");
                if (saveButton == null)
                {
                    continue;
                }

                var fileName = match.Groups[1].Value;
                Click(saveButton);

                return fileName;
            }
            return null;
        }

        private static string GetDefaultDownloadDirectory()
        {
            // c# - How to find browser download folder path - Stack Overflow
            // https://stackoverflow.com/a/24673279
            var path = string.Empty;
            var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main");
            if (key != null)
            {
                path = (string)key.GetValue("Default Download Directory");
            }
            if (string.IsNullOrEmpty(path))
            {
                path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads";
            }

            return path;
        }

        private static string GetValue(AutomationElement ae)
        {
            return (ae.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern).Current.Value;
        }

        private static T Timeout<T>(Func<T> action, Predicate<T> predicate, TimeSpan timeout)
        {
            var endTime = timeout == TimeSpan.Zero ? DateTime.MaxValue : DateTime.Now + timeout;
            while (DateTime.Now < endTime)
            {
                var x = action();
                if (predicate(x))
                {
                    return x;
                }
                Thread.Sleep(100);
            }
            throw new TimeoutException();
        }

        private static string WaitUntilDownloadCompleted(string downloadFileName)
        {
            var movedOrDeletedText = downloadFileName + " は移動または削除された可能性があります。";

            var index = downloadFileName.LastIndexOf('.');
            var fileNamePattern = index == -1
                ? Regex.Escape(downloadFileName) + @"( \(\d*?\))*"
                : Regex.Escape(downloadFileName.Substring(0, index)) + @"( \(\d*?\))*\." + downloadFileName.Substring(index + 1);
            var downloadCompletedRegex = new Regex("(" + fileNamePattern + ") のダウンロードが完了しました。");

            var downloadedFileName = Timeout<string>(
                () => {
                    var processes = Process.GetProcessesByName("iexplore");
                    foreach (var process in processes)
                    {
                        var hWnd = process.MainWindowHandle;

                        var notificationBar = FindNotificationBar(hWnd);
                        if (notificationBar == null)
                        {
                            continue;
                        }

                        var text = FindText(notificationBar);
                        if (text == null)
                        {
                            continue;
                        }

                        var textValue = GetValue(text);

                        if (textValue == movedOrDeletedText)
                        {
                            var retryButton = FindElementByName(notificationBar, "再試行");
                            if (retryButton != null)
                            {
                                Click(retryButton);
                            }
                            continue;
                        }

                        var match = downloadCompletedRegex.Match(textValue);
                        if (!match.Success)
                        {
                            continue;
                        }

                        var closeButton = FindElementByName(notificationBar, "閉じる");
                        if (closeButton == null)
                        {
                            continue;
                        }

                        var fileName = match.Groups[1].Value;
                        Click(closeButton);

                        return fileName;
                    }
                    return null;
                },
                (x) => x != null,
                DownloadTimeout
            );
            return downloadedFileName;
        }
    }
}
