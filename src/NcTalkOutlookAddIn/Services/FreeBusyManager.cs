// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using Microsoft.Win32;
using Outlook = Microsoft.Office.Interop.Outlook;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
        // The local IFB (Free/Busy) HTTP listener has been removed — Exchange handles Free/Busy
        // natively in this deployment. This class now only restores Outlook's native free/busy
        // registry paths, in case an older build ever ran with the listener enabled and hijacked them.
    internal sealed class FreeBusyManager
    {
        private Outlook.Application _application;

        internal void Initialize(Outlook.Application application)
        {
            _application = application;
        }

        internal void ApplySettings(AddinSettings settings)
        {            if (_application == null || settings == null)
            {
                return;
            }
            RestoreFreeBusyPath(settings);
        }

        private void RestoreFreeBusyPath(AddinSettings settings)
        {
            string previous = settings.IfbPreviousFreeBusyPath ?? string.Empty;
            TryRestoreKey(BuildCalendarOptionsKey(), "FreeBusySearchPath", previous);
            TryRestoreKey(BuildPolicyCalendarOptionsKey(), "FreeBusySearchPath", previous);
            TryRestoreKey(BuildInternetFreeBusyKey(), "Read URL", previous);
            TryRestoreKey(BuildPolicyInternetFreeBusyKey(), "Read URL", previous);
        }

        private string BuildCalendarOptionsKey()
        {
            return @"Software\Microsoft\Office\" + GetOutlookVersionSegment() + @"\Outlook\Options\Calendar";
        }

        private string BuildInternetFreeBusyKey()
        {
            return BuildCalendarOptionsKey() + @"\Internet Free/Busy";
        }

        private string BuildPolicyCalendarOptionsKey()
        {
            return @"Software\Policies\Microsoft\Office\" + GetOutlookVersionSegment() + @"\Outlook\Options\Calendar";
        }

        private string BuildPolicyInternetFreeBusyKey()
        {
            return BuildPolicyCalendarOptionsKey() + @"\Internet Free/Busy";
        }

        private string GetOutlookVersionSegment()
        {
            string version = "16.0";
            try
            {
                string versionString = _application != null ? _application.Version : null;
                if (!string.IsNullOrEmpty(versionString))
                {
                    string[] parts = versionString.Split('.');
                    if (parts.Length >= 2)
                    {
                        version = parts[0] + "." + parts[1];
                    }
                    else if (parts.Length == 1)
                    {
                        version = parts[0] + ".0";
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to read Outlook version. Falling back to default version segment.", ex);
                version = "16.0";
            }
            return version;
        }

        private void TryRestoreKey(string path, string valueName, string previous)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }
            string existing;
            if (TryReadValue(path, valueName, out existing))
            {
                if (!string.IsNullOrEmpty(previous))
                {
                    if (string.Equals(existing, previous, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }
                }
                else if (string.IsNullOrEmpty(existing))
                {
                    return;
                }
            }
            try
            {
                using (var key = OpenOrCreateSubKey(path))
                {                    if (key == null)
                    {
                        DiagnosticsLogger.Log(LogCategories.Ifb, "No access to registry '" + path + "' while restoring.");
                        return;
                    }
                    if (!string.IsNullOrEmpty(previous))
                    {
                        key.SetValue(valueName, previous, RegistryValueKind.String);
                    }
                    else
                    {
                        try
                        {
                            key.DeleteValue(valueName, false);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to delete registry value '" + valueName + "' at '" + path + "'. Falling back to setting empty string.", ex);
                            key.SetValue(valueName, string.Empty, RegistryValueKind.String);
                        }
                    }
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                DiagnosticsLogger.Log(LogCategories.Ifb, "No access to registry '" + path + "' while restoring: " + ex.Message);
            }
            catch (System.Security.SecurityException ex)
            {
                DiagnosticsLogger.Log(LogCategories.Ifb, "Security exception for registry '" + path + "' while restoring: " + ex.Message);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to restore registry value '" + valueName + "' at '" + path + "'.", ex);
            }
        }

        private bool TryReadValue(string path, string valueName, out string value)
        {
            value = string.Empty;
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(path, false))
                {                    if (key == null)
                    {
                        return false;
                    }

                    object raw = key.GetValue(valueName, string.Empty);
                    value = raw as string ?? string.Empty;
                    return true;
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "No access to registry '" + path + "' while reading '" + valueName + "'.", ex);
                return false;
            }
            catch (System.Security.SecurityException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Security exception while reading registry '" + path + "' value '" + valueName + "'.", ex);
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to read registry '" + path + "' value '" + valueName + "'.", ex);
                return false;
            }
        }

        private static RegistryKey OpenOrCreateSubKey(string path)
        {
            try
            {
                var key = Registry.CurrentUser.OpenSubKey(path, true);                if (key != null)
                {
                    return key;
                }
                return Registry.CurrentUser.CreateSubKey(path);
            }
            catch (UnauthorizedAccessException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "No access to registry '" + path + "' (open/create).", ex);
                return null;
            }
            catch (System.Security.SecurityException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Security exception for registry '" + path + "' (open/create).", ex);
                return null;
            }
        }
    }
}
