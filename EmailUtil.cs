using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.TeamFoundation.Common;

namespace EdiUtilities.ConfigurationChangeGates
{
    public static class EmailUtil
    {
        public static IEnumerable<string> GetEmailIds(this IList<string> userList)
        {
            return userList == null ? 
                Enumerable.Empty<string>() :
                userList.SelectMany(GetEmailIds);
        }

        public static IEnumerable<string> GetEmailIds(string userList)
        {
            if (userList.IsNullOrEmpty())
            {
                return Enumerable.Empty<string>();
            }

            char[] separators = { ',', ';' };

            return userList
                .Split(separators, StringSplitOptions.RemoveEmptyEntries)
                .Select(u => u.Trim())
                .Where(u => !string.IsNullOrEmpty(u) && !u.Contains(' '))
                .Select(AppendAtMicrosoftDotComIfRequired);
        }

        private static string AppendAtMicrosoftDotComIfRequired(string mailAlias)
        {
            return mailAlias == null || mailAlias.Contains("@") ? mailAlias : $"{mailAlias}@microsoft.com";
        }
    }
}
