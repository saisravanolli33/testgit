using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.Office.Telemetry.Properties;

namespace Microsoft.Office.Telemetry
{

    public static class ErrorHandling
    {
        /// <summary>
        /// Indicates to Code Analysis that a method validates a particular parameter.
        /// This is an undocumented VS feature.
        /// </summary>
        [AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false, Inherited = false)]
        private sealed class ValidatedNotNullAttribute : Attribute { }

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Caller* attributes only work on parameters with default values")]
        public static void VerifyArgumentNotNull<T>(
            [ValidatedNotNull] T parameter,
            string parameterName,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0) where T : class
        {
            if (parameter == null)
            {
                ThrowException<ArgumentNullException>(parameterName, memberName, sourceFilePath, sourceLineNumber);
            }
        }

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Caller* attributes only work on parameters with default values")]
        public static void VerifyArgumentNotNullNorEmpty<T>(
            [ValidatedNotNull] IEnumerable<T> parameter,
            string parameterName,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0)
        {
            if (parameter == null)
            {
                ThrowException<ArgumentNullException>(parameterName, memberName, sourceFilePath, sourceLineNumber);
            }
            else if (!parameter.Any())
            {
                ThrowException<ArgumentException>(String.Format(CultureInfo.InvariantCulture, Resources.ErrorHandling_VerifyArgumentNotNullNorEmpty_Argument_should_not_be_empty, parameterName), memberName, sourceFilePath, sourceLineNumber);
            }
        }

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Caller* attributes only work on parameters with default values")]
        public static void VerifyStringNotNullNorEmpty(
            [ValidatedNotNull] string parameter,
            string parameterName,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0)
        {
            if (string.IsNullOrEmpty(parameter))
            {
                if (parameter == null)
                {
                    ThrowException<ArgumentNullException>(parameterName, memberName, sourceFilePath, sourceLineNumber);
                }
                else
                {
                    ThrowException<ArgumentException>(String.Format(CultureInfo.InvariantCulture, "Argument should not be empty : {0}", parameterName), memberName, sourceFilePath, sourceLineNumber);
                }
            }
        }

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Caller* attributes only work on parameters with default values")]
        public static void VerifyStringNotNullNorWhiteSpace(
            [ValidatedNotNull] string parameter,
            string parameterName,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0)
        {
            if (string.IsNullOrWhiteSpace(parameter))
            {
                if (parameter == null)
                {
                    ThrowException<ArgumentNullException>(parameterName, memberName, sourceFilePath, sourceLineNumber);
                }
                else
                {
                    ThrowException<ArgumentException>(String.Format(CultureInfo.InvariantCulture, Resources.ErrorHandling_VerifyStringNotNullNorWhiteSpace,parameterName), memberName, sourceFilePath, sourceLineNumber);
                }
            }
        }

        [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Caller* attributes only work on parameters with default values")]
        public static void ThrowExceptionIf<T>(
            bool condition,
            string message = null,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0
            ) where T : Exception
        {
            if (condition)
            {
                ThrowException<T>(message, memberName, sourceFilePath, sourceLineNumber);
            }
        }

        [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Caller* attributes only work on parameters with default values")]
        public static void ThrowException<T>(
            string message = null,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0
            ) where T : Exception
        {
            var exceptionType = typeof(T);
            var exceptionInstance = (T)Activator.CreateInstance(exceptionType, message);
            exceptionInstance.Data.Add("CallerMemberName", memberName);
            exceptionInstance.Data.Add("CallerFilePath", sourceFilePath );
            exceptionInstance.Data.Add("CallerLineNumber", sourceLineNumber);

            throw exceptionInstance;
        }
    }
}
