using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Automation;

namespace HeadlessWindowsAutomation
{
    /// <summary>
    /// Interface for custom conditions used to evaluate <see cref="AutomationElement"/>.
    /// </summary>
    internal interface ICustomCondition
    {
        /// <summary>
        /// Evaluates the specified <see cref="AutomationElement"/> against the condition.
        /// </summary>
        /// <param name="element">The <see cref="AutomationElement"/> to evaluate.</param>
        /// <returns>True if the element meets the condition; otherwise, false.</returns>
        bool Evaluate(AutomationElement element);
    }

    /// <summary>
    /// Represents a custom condition that evaluates an <see cref="AutomationElement"/> property against a regular expression.
    /// </summary>
    internal sealed class CustomPropertyRegexCondition : ICustomCondition
    {
        internal AutomationProperty Property { get; }
        internal Regex Pattern { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomPropertyRegexCondition"/> class.
        /// </summary>
        /// <param name="property">The <see cref="AutomationProperty"/> to evaluate.</param>
        /// <param name="pattern">The regular expression pattern to match.</param>
        /// <param name="options">Optional. The <see cref="RegexOptions"/> to use.</param>
        internal CustomPropertyRegexCondition(AutomationProperty property, string pattern, RegexOptions options = RegexOptions.None)
        {
            Property = property;
            Pattern = new Regex(pattern, RegexOptions.Compiled | options);
        }

        /// <summary>
        /// Evaluates the specified <see cref="AutomationElement"/> against the regular expression condition.
        /// </summary>
        /// <param name="element">The <see cref="AutomationElement"/> to evaluate.</param>
        /// <returns>True if the element's property value matches the regular expression; otherwise, false.</returns>
        public bool Evaluate(AutomationElement element)
        {
            object propertyValue = element.GetCurrentPropertyValue(Property);
            return propertyValue != null && Pattern.IsMatch(propertyValue.ToString());
        }

        /// <summary>
        /// Helper class for parsing regular expression strings.
        /// </summary>
        internal class RegexStringParser
        {
            private readonly string _value;
            internal string Pattern { get; private set; }
            private string _flags;
            internal RegexOptions Options { get; private set; } = RegexOptions.None;
            internal Dictionary<char, RegexOptions> FlagMapping { get; } = new Dictionary<char, RegexOptions>() {
                { 'i', RegexOptions.IgnoreCase },       // Case-insensitive matching
                { 'm', RegexOptions.Multiline },        // Multiline mode
                { 's', RegexOptions.Singleline },       // Dot matches newline
                // Global (g) and Unicode (u) handled differently in .NET. Sticky (y) not supported
                { 'g', RegexOptions.None },
                { 'u', RegexOptions.None },
                { 'y', RegexOptions.None }
            };

            /// <summary>
            /// Initializes a new instance of the <see cref="RegexStringParser"/> class.
            /// </summary>
            /// <param name="value">The regular expression string to parse.</param>
            internal RegexStringParser(string value)
            {
                _value = value.Trim();
            }

            /// <summary>
            /// Parses the regular expression string.
            /// </summary>
            /// <returns>True if the parsing is successful; otherwise, false.</returns>
            internal bool Parse()
            {

                if (_value.StartsWith("/"))
                {
                    // Find the last '/' to determine the end of the pattern and start of the flags
                    int lastSlashIndex = _value.LastIndexOf('/');
                    if (lastSlashIndex > 0)
                    {
                        // Extract potential flags
                        this._flags = _value.Substring(lastSlashIndex + 1);

                        // Validate flags
                        foreach (char c in this._flags)
                        {
                            if (!FlagMapping.ContainsKey(c))
                            {
                                return false; // Invalid flag found
                            }
                            else
                            {
                                // Make Options from flags using FlagMapping
                                if (FlagMapping.TryGetValue(c, out RegexOptions option))
                                {
                                    this.Options |= option;
                                }
                            }
                        }

                        // Extract pattern
                        this.Pattern = _value.Substring(1, lastSlashIndex - 1);
                        return this.Pattern.Length > 0;
                    }
                }
                return false;
            }
        }
    }

    /// <summary>
    /// Represents a custom condition that evaluates an <see cref="AutomationElement"/> property against a specific value.
    /// </summary>
    internal sealed class CustomPropertyCondition : ICustomCondition 
    {
        private readonly PropertyCondition _condition;

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomPropertyCondition"/> class.
        /// </summary>
        /// <param name="property">The <see cref="AutomationProperty"/> to evaluate.</param>
        /// <param name="value">The value to match.</param>
        internal CustomPropertyCondition(AutomationProperty property, object value)
        {
            this._condition = new PropertyCondition(property, value);
        }

        /// <summary>
        /// Evaluates the specified <see cref="AutomationElement"/> against the property condition.
        /// </summary>
        /// <param name="element">The <see cref="AutomationElement"/> to evaluate.</param>
        /// <returns>True if the element's property value matches the specified value; otherwise, false.</returns>
        public bool Evaluate(AutomationElement element) 
        {
            var found = element.FindFirst(TreeScope.Element, this._condition);
            return found != null;
        }
    }

    /// <summary>
    /// Represents a custom condition that evaluates multiple conditions using a logical AND operation.
    /// </summary>
    internal sealed class CustomAndCondition : ICustomCondition
    {
        private readonly ICustomCondition[] _conditions;

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomAndCondition"/> class.
        /// </summary>
        /// <param name="conditions">The array of conditions to evaluate.</param>
        internal CustomAndCondition(params ICustomCondition[] conditions)
        {
            _conditions = conditions;
        }

        /// <summary>
        /// Evaluates the specified <see cref="AutomationElement"/> against all the conditions.
        /// </summary>
        /// <param name="element">The <see cref="AutomationElement"/> to evaluate.</param>
        /// <returns>True if the element meets all the conditions; otherwise, false.</returns>
        public bool Evaluate(AutomationElement element)
        {
            foreach (var condition in this._conditions)
            {
                if (!condition.Evaluate(element))
                {
                    return false;
                }
            }
            return true;
        }
    }

    /// <summary>
    /// Represents a custom condition that evaluates to a boolean value.
    /// </summary>
    internal sealed class CustomBoolCondition : ICustomCondition
    {
        internal static readonly CustomBoolCondition TrueCondition = new CustomBoolCondition(true);
        internal static readonly CustomBoolCondition FalseCondition = new CustomBoolCondition(false);

        private readonly bool _value;

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomBoolCondition"/> class.
        /// </summary>
        /// <param name="value">The boolean value to evaluate.</param>
        internal CustomBoolCondition(bool value)
        {
            _value = value;
        }

        /// <summary>
        /// Evaluates the specified <see cref="AutomationElement"/> against the boolean condition.
        /// </summary>
        /// <param name="element">The <see cref="AutomationElement"/> to evaluate.</param>
        /// <returns>The boolean value of the condition.</returns>
        public bool Evaluate(AutomationElement element)
        {
            return this._value;
        }
    }
}