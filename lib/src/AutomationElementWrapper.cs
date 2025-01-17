using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;

namespace HeadlessWindowsAutomation
{
    /// <summary>
    /// Wrapper class for <see cref="AutomationElement"/> to provide additional automation functionalities.
    /// </summary>
    public class AutomationElementWrapper
    {
        /// <summary>
        /// Gets the associated <see cref="AutomationElement"/>.
        /// </summary>
        public AutomationElement Element { get; }

        /// <summary>
        /// Gets or sets the parent <see cref="AutomationElementWrapper"/>.
        /// </summary>
        public AutomationElementWrapper Parent { get; set; }

        /// <summary>
        /// Gets the Keyboard instance for simulating keyboard interactions.
        /// </summary>
        public Keyboard Keyboard { get; }

        /// <summary>
        /// Gets the configuration settings.
        /// </summary>
        public readonly Config Config = new Config();

        /// <summary>
        /// An <see cref="AutomationElementWrapper"/> representing the root window.
        /// </summary>
        public AutomationElementWrapper RootWindow { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutomationElementWrapper"/> class with the specified <see cref="AutomationElement"/>.  
        /// The <see cref="AutomationElementWrapper.RootWindow" /> is the new instance.
        /// </summary>
        /// <param name="element">The <see cref="AutomationElement"/> to wrap.</param>
        public AutomationElementWrapper(AutomationElement element)
        {
            this.Keyboard = new Keyboard(this);
            this.Element = element;
            this.RootWindow = this;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutomationElementWrapper"/> class with the specified window handle.
        /// The <see cref="AutomationElementWrapper.RootWindow" /> is the new instance.
        /// </summary>
        /// <param name="hwnd">The window handle.</param>
        /// <exception cref="ArgumentException">Thrown when the <see cref="AutomationElement"/> is not found for the given handle.</exception>
        public AutomationElementWrapper(IntPtr hwnd)
        {
            this.Element = WindowsAPIHelper.GetAutomationElement(hwnd);
            if (this.Element == null) throw new ArgumentException($"AutomationElement not found, invalid handle {hwnd}");
            this.Keyboard = new Keyboard(this);
            this.RootWindow = this;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutomationElementWrapper"/> class with the specified <see cref="AutomationElement"/> and parent.
        /// </summary>
        /// <param name="element">The <see cref="AutomationElement"/> to wrap.</param>
        /// <param name="parent">The parent <see cref="AutomationElementWrapper"/>.</param>
        public AutomationElementWrapper(AutomationElement element, AutomationElementWrapper parent) : this(element)
        {
            this.Parent = parent;
            this.Config.CopyFrom(parent.Config);
            this.RootWindow = parent.RootWindow;
        }

        /// <summary>
        /// Returns a string representation of the <see cref="AutomationElementWrapper"/>.
        /// </summary>
        /// <returns>A string representing the <see cref="AutomationElementWrapper"/>.</returns>
        public override string ToString()
        {
            return $"{this.Element.Current.Name} of {this.Element.Current.ControlType.ProgrammaticName}";
        }

        /// <summary>
        /// Gets the first <see cref="AutomationElementWrapper"/> with a window handle.  
        /// It will start from the current element then look into the ancestors.
        /// </summary>
        /// <returns>The first <see cref="AutomationElementWrapper"/> with a window handle.</returns>
        /// <exception cref="InvalidOperationException">Thrown when no window handle is found and <see cref="Config.ShowError"/> is true.</exception>
        public AutomationElementWrapper GetFirstElementWithHwnd()
        {
            int hwnd = this.Element.Current.NativeWindowHandle;
            if (hwnd != 0) return this;
            else if (this.Parent != null)
            {
                return this.Parent.GetFirstElementWithHwnd();
            }

            if (this.Config.ShowError) throw new InvalidOperationException($"{this} has no window handle and no parent");
            else return null;
        }

        /// <summary>
        /// Gets the window handle for the <see cref="AutomationElement"/>.
        /// </summary>
        /// <returns>The window handle as <see cref="IntPtr"/>.</returns>
        public IntPtr GetHwnd()
        {
            AutomationElementWrapper elementWithHwnd = this.GetFirstElementWithHwnd();  // throw when not found
            return new IntPtr(elementWithHwnd.Element.Current.NativeWindowHandle);
        }

        /// <summary>
        /// Sets focus to the current <see cref="AutomationElement"/>.  
        /// Doesn't work in headless.
        /// </summary>
        public AutomationElementWrapper SetFocus()
        {
            this.Element.SetFocus();
            return this;
        }

        /// <summary>
        /// Posts a Windows message to the specified window handle.
        /// </summary>
        /// <param name="hwnd">The window handle.</param>
        /// <param name="msg">The message to post.</param>
        /// <param name="wParam">The WPARAM value.</param>
        /// <param name="lParam">The LPARAM value.</param>
        /// <returns>The result of the message processing.</returns>
        internal void PostMessageWin32(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                bool result = WindowsAPIHelper.PostMessage(hwnd, msg, wParam, lParam);
                if (!result && this.Config.ShowError) throw new Exception($"Call to {msg} failed, error: {Marshal.GetLastWin32Error()}");
            }
            catch (Exception e)
            {
                if (this.Config.ShowError) throw new Exception($"Failed to send message to {this}, error: {e}");
            }
        }

        /// <summary>
        /// Posts a Windows message to the current window handle.
        /// </summary>
        /// <param name="hwnd">The window handle.</param>
        /// <param name="msg">The message to post.</param>
        /// <param name="wParam">The WPARAM value.</param>
        /// <param name="lParam">The LPARAM value.</param>
        /// <returns>The result of the message processing.</returns>
        internal void PostMessageWin32(uint msg, IntPtr wParam, IntPtr lParam)
        {
            var hwnd = this.GetHwnd();
            if (hwnd != IntPtr.Zero)
            {
                this.PostMessageWin32(hwnd, msg, wParam, lParam);
            }
            else
            {
                if (this.Config.ShowError) throw new Exception($"Handle is not defined for {this}");
            }
        }

        /// <summary>
        /// Sends a Windows message to the specified window handle.
        /// </summary>
        /// <param name="hwnd">The window handle.</param>
        /// <param name="msg">The message to send.</param>
        /// <param name="wParam">The WPARAM value.</param>
        /// <param name="lParam">The LPARAM value.</param>
        /// <returns>The result of the message processing.</returns>
        /// <exception cref="Exception">Thrown when the message sending fails and <see cref="Config.ShowError"/> is true.</exception>
        internal IntPtr SendMessageWin32(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                return WindowsAPIHelper.SendMessage(hwnd, msg, wParam, lParam);
            }
            catch (Exception e)
            {
                if (this.Config.ShowError) throw new Exception($"Failed to send message to {this}, error: {e}");
                else return IntPtr.Zero;
            }
        }

        /// <summary>
        /// Sends a Windows message to the current window handle.
        /// </summary>
        /// <param name="msg">The message to send.</param>
        /// <param name="wParam">The WPARAM value.</param>
        /// <param name="lParam">The LPARAM value.</param>
        /// <returns>The result of the message processing.</returns>
        /// <exception cref="Exception">Thrown when the handle is not defined and <see cref="Config.ShowError"/> is true.</exception>
        internal IntPtr SendMessageWin32(uint msg, IntPtr wParam, IntPtr lParam)
        {
            var hwnd = this.GetHwnd();
            if (hwnd != IntPtr.Zero)
            {
                return this.SendMessageWin32(hwnd, msg, wParam, lParam);
            }
            else
            {
                if (this.Config.ShowError) throw new Exception($"Handle is not defined for {this}");
                else return IntPtr.Zero;
            }
        }

        /// <summary>
        /// Sets the text of the current <see cref="AutomationElement"/>.
        /// </summary>
        /// <param name="text">The text to set.</param>
        /// <returns>The current <see cref="AutomationElementWrapper"/> instance.</returns>
        public AutomationElementWrapper SetText(string text)
        {
            IntPtr hGlobal = Marshal.StringToHGlobalAuto(text);
            this.SendMessageWin32(WindowsAPIHelper.WM_SETTEXT, IntPtr.Zero, hGlobal);
            Marshal.FreeHGlobal(hGlobal);
            return this;
        }

        /// <summary>
        /// Gets the text of the current <see cref="AutomationElement"/>.
        /// Text is the value property, the text property, or the name property. In that order.
        /// </summary>
        /// <returns>The text of the current <see cref="AutomationElement"/>.</returns>
        public string GetText()
        {
            object patternObj;
            if (this.Element.TryGetCurrentPattern(ValuePattern.Pattern, out patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                return valuePattern.Current.Value;
            }
            else if (this.Element.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
            {
                var textPattern = (TextPattern)patternObj;
                return textPattern.DocumentRange.GetText(-1).TrimEnd('\r'); // often there is an extra '\r' hanging off the end.
            }
            else
            {
                return this.Element.Current.Name;
            }
        }

        /// <summary>
        /// Gets the ID of the control.
        /// </summary>
        /// <returns>The hash code of the <see cref="AutomationElement"/>.</returns>
        public int GetIdOfControl()
        {
            return this.Element.GetHashCode();
        }

        /// <summary>
        /// Clicks the current <see cref="AutomationElement"/> at the specified location.  
        /// No location is required if the current <see cref="AutomationElement"/> has a window handle.
        /// </summary>
        /// <param name="x">Optional. Relative X.</param>
        /// <param name="y">Optional. Relative Y.</param>
        /// <param name="absX">Optional. Absolute X.</param>
        /// <param name="absY">Optional. Absolute Y.</param>
        /// <param name="isMouseCaptured">Optional. Allows to send mouse messages.</param>
        /// <returns>The current <see cref="AutomationElementWrapper"/> instance.</returns>
        public AutomationElementWrapper Click(int x = 0, int y = 0, int absX = 0, int absY = 0, bool isMouseCaptured = true)
        {
            IntPtr hwnd = this.GetHwnd();
            IntPtr lParam = WindowsAPIHelper.MakeLParam(x, y);
            if (!isMouseCaptured)
            {
                this.SendMessageWin32(hwnd, WindowsAPIHelper.WM_MOUSEMOVE, (IntPtr)WindowsAPIHelper.MK_LBUTTON, lParam);
                IntPtr absLParam = WindowsAPIHelper.MakeLParam(absX, absY);
                var ncHit = this.SendMessageWin32(hwnd, WindowsAPIHelper.WM_NCHITTEST, IntPtr.Zero, absLParam);
                this.SendMessageWin32(hwnd, WindowsAPIHelper.WM_SETCURSOR, hwnd, WindowsAPIHelper.MakeLParam((int)ncHit, (int)WindowsAPIHelper.WM_MOUSEMOVE));
            }
            this.PostMessageWin32(hwnd, WindowsAPIHelper.WM_LBUTTONDOWN, (IntPtr)WindowsAPIHelper.MK_LBUTTON, lParam);
            this.PostMessageWin32(hwnd, WindowsAPIHelper.WM_LBUTTONUP, (IntPtr)WindowsAPIHelper.MK_LBUTTON, lParam);

            this.NotifyParent(lParam);

            return this;
        }

        /// <summary>
        /// Double-clicks the current <see cref="AutomationElement"/> at the specified location.  
        /// No location is required if the current <see cref="AutomationElement"/> has a window handle.
        /// </summary>
        /// <param name="x">Optional. Relative X.</param>
        /// <param name="y">Optional. Relative Y.</param>
        /// <param name="absX">Optional. Absolute X.</param>
        /// <param name="absY">Optional. Absolute Y.</param>
        /// <param name="isMouseCaptured">Optional. Allows to send mouse messages.</param>
        /// <returns>The current <see cref="AutomationElementWrapper"/> instance.</returns>
        public AutomationElementWrapper DoubleClick(int x = 0, int y = 0, int absX = 0, int absY = 0, bool isMouseCaptured = true)
        {
            IntPtr hwnd = this.GetHwnd();
            IntPtr lParam = WindowsAPIHelper.MakeLParam(x, y);
            if (!isMouseCaptured)
            {
                this.SendMessageWin32(hwnd, WindowsAPIHelper.WM_MOUSEMOVE, (IntPtr)WindowsAPIHelper.MK_LBUTTON, lParam);
                IntPtr absLParam = WindowsAPIHelper.MakeLParam(absX, absY);
                var ncHit = this.SendMessageWin32(hwnd, WindowsAPIHelper.WM_NCHITTEST, IntPtr.Zero, absLParam);
                this.SendMessageWin32(hwnd, WindowsAPIHelper.WM_SETCURSOR, hwnd, WindowsAPIHelper.MakeLParam((int)ncHit, (int)WindowsAPIHelper.WM_MOUSEMOVE));
            }
            this.PostMessageWin32(hwnd, WindowsAPIHelper.WM_LBUTTONDBLCLK, (IntPtr)WindowsAPIHelper.MK_LBUTTON, lParam);

            this.NotifyParent(lParam);

            return this;
        }

        /// <summary>
        /// Gets the relative location of the current <see cref="AutomationElement"/> compared to an absolute location.
        /// </summary>
        /// <param name="absoluteLocation">The absolution location.</param>
        /// <returns>The relative location as a <see cref="Point"/>.</returns>
        public Point GetRelativeLocation(Point absoluteLocation)
        {
            AutomationElementWrapper window = this.GetFirstElementWithHwnd();
            if (window.Parent != null)  // On top level: relative is absolute as relative to the screen
            {
                var rect = window.Element.Current.BoundingRectangle;
                return new Point(absoluteLocation.X - (int)rect.X, absoluteLocation.Y - (int)rect.Y);
            }
            else return absoluteLocation;
        }

        /// <summary>
        /// Gets the clickable point of the current <see cref="AutomationElement"/>.  
        /// Based only on BoundingRectangle property, no visiblity check.
        /// </summary>
        /// <returns>The clickable point as a <see cref="Point"/>.</returns>
        public Point GetHeadlessClickablePoint()
        {
            var boundingRect = this.Element.Current.BoundingRectangle;
            if (boundingRect.IsEmpty)
            {
                if (this.Config.ShowError) throw new Exception($"{this} has no BoundingRectangle");
                else return default;
            }
            return new Point(boundingRect.X + boundingRect.Width / 2, boundingRect.Y + boundingRect.Height / 2);
        }

        /// <summary>
        /// Clicks on the location of the current <see cref="AutomationElement"/>.
        /// </summary>
        /// <returns>The current <see cref="AutomationElementWrapper"/> instance.</returns>
        public AutomationElementWrapper ClickOnLocation()
        {
            Point location = this.GetHeadlessClickablePoint();
            Point relativeLoc = this.GetRelativeLocation(location);
            return this.Click((int)relativeLoc.X, (int)relativeLoc.Y, (int)location.X, (int)location.Y, false);
        }

        /// <summary>
        /// Double-clicks on the location of the current <see cref="AutomationElement"/>.
        /// </summary>
        /// <returns>The current <see cref="AutomationElementWrapper"/> instance.</returns>
        public AutomationElementWrapper DoubleClickOnLocation()
        {
            Point location = this.GetHeadlessClickablePoint();
            Point relativeLoc = this.GetRelativeLocation(location);
            return this.DoubleClick((int)relativeLoc.X, (int)relativeLoc.Y, (int)location.X, (int)location.Y, false);
        }

        /// <summary>
        /// Simulates pressing the Enter key.
        /// </summary>
        /// <returns>The current <see cref="AutomationElementWrapper"/> instance.</returns>
        public AutomationElementWrapper Enter()
        {
            this.Keyboard.PressKey(Keys.Enter, new Keyboard.Options { SetFocus = true });
            return this;
        }

        /// <summary>
        /// Finds the first element where the given property is of the expected value.
        /// </summary>
        /// <param name="property">Property to check.</param>
        /// <param name="value">Expected value for the property.</param>
        /// <param name="scope">Optional. Specifies the search scope. Default to Children.</param>
        /// <returns>The found <see cref="AutomationElementWrapper"/> as a child of the current instance.</returns>
        public AutomationElementWrapper FindElement(AutomationProperty property, object value, TreeScope scope = TreeScope.Children)
        {
            ICustomCondition condition = this.CreateCustomPropertyCondition(property, value);
            AutomationElement found = null;
            AutomationElementFinder finder = new AutomationElementFinder(this.Element);
            if (this.Config.FindWaitForElement) found = this.RetryFindElement(() =>
            {
                return finder.FindFirst(scope, condition);
            });
            else
            {
                found = finder.FindFirst(scope, condition);
            }

            if (found != null) return new AutomationElementWrapper(found, this);
            else
            {
                if (this.Config.ShowError) Console.Error.WriteLine($"Failed to find the AutomationElement using {property.ProgrammaticName} = {value} inside {this}");
                return null;
            }
        }

        private ICustomCondition CreateCustomPropertyCondition(AutomationProperty property, object value)
        {
            if (this.Config.UseRegex && value is string stringValue)
            {
                var regexStringParser = new CustomPropertyRegexCondition.RegexStringParser(stringValue);
                if (regexStringParser.Parse())
                {
                    return new CustomPropertyRegexCondition(property, regexStringParser.Pattern, regexStringParser.Options);
                }
            }

            return new CustomPropertyCondition(property, value);
        }

        /// <summary>
        /// Finds the first element that matches all the conditions.
        /// </summary>
        /// <param name="propertyConditions">Collection of condition to match.</param>
        /// <param name="scope">Optional. Specifies the search scope. Default to Children.</param>
        /// <returns>The found <see cref="AutomationElementWrapper"/> as a child of the current instance.</returns>
        public AutomationElementWrapper FindElement(Dictionary<AutomationProperty, object> propertyConditions, TreeScope scope = TreeScope.Children)
        {
            if (propertyConditions.Count <= 0) throw new ArgumentException("Invalid condition.", nameof(propertyConditions));
            if (propertyConditions.Count == 1) return this.FindElement(propertyConditions.First().Key, propertyConditions.First().Value, scope);

            List<ICustomCondition> conditions = new List<ICustomCondition>();

            foreach (var propertyCondition in propertyConditions)
            {
                conditions.Add(this.CreateCustomPropertyCondition(propertyCondition.Key, propertyCondition.Value));
            }

            var combinedCondition = new CustomAndCondition(conditions.ToArray());
            AutomationElement found = null;
            AutomationElementFinder finder = new AutomationElementFinder(this.Element);
            if (this.Config.FindWaitForElement) found = this.RetryFindElement(() =>
            {
                return finder.FindFirst(scope, combinedCondition);
            });
            else
            {
                found = finder.FindFirst(scope, combinedCondition);
            }

            if (found != null) return new AutomationElementWrapper(found, this);
            else
            {
                if (this.Config.ShowError) Console.Error.WriteLine($"Failed to find the AutomationElement using the provided conditions inside {this}");
                return null;
            }
        }

        /// <summary>
        /// Finds the first element that matches the specified expression.  
        /// It only looks at nested elements inside the current <see cref="AutomationElementWrapper"/>, without any "path".
        /// </summary>
        /// <param name="expr">Expression as XPath condition: "ControlType[@attribute='value' and @attribute2='value' ...]".  
        /// Example: Edit[@HelpText='foo']</param>
        /// <param name="scope">Specifies the search scope. Children for the element's immediate children, Descendants for anywhere inside the element's.</param>
        /// <returns>The found <see cref="AutomationElementWrapper"/> as a child of the current instance.</returns>
        public AutomationElementWrapper FindElement(string expr, TreeScope scope = TreeScope.Children)
        {
            var propertyConditions = this.ParseConditionExpression(expr);
            return this.FindElement(propertyConditions, scope);
        }

        /// <summary>
        /// Parses an XPath condition expression.
        /// </summary>
        /// <param name="expr">The condition expression.</param>
        /// <returns>Collection of property,value corresponding to the expression.</returns>
        private Dictionary<AutomationProperty, object> ParseConditionExpression(string expr)
        {
            var propertyConditions = new Dictionary<AutomationProperty, object>();

            // Separate ControlType from the content inside the brackets
            int openBracketIndex = expr.IndexOf('[');
            var controlTypePart = expr.Substring(0, openBracketIndex > 0 ? openBracketIndex : expr.Length).Trim();
            string conditionsPart;
            if (openBracketIndex > 0) conditionsPart = expr.Substring(openBracketIndex + 1, expr.LastIndexOf(']') - openBracketIndex - 1);
            else conditionsPart = "";

            // Handle ControlType
            if (!string.IsNullOrEmpty(controlTypePart))
            {
                var controlType = GetControlTypeFromName(controlTypePart);
                if (controlType != null)
                {
                    propertyConditions[AutomationElement.ControlTypeProperty] = controlType;
                }
                else if (this.Config.ShowError) Console.Error.WriteLine($"ControlType not found for {controlTypePart} in expression {expr}");
            }

            if (!String.IsNullOrEmpty(conditionsPart))
            {
                // Split the conditions inside the brackets
                var conditions = conditionsPart.Split(new[] { " and " }, StringSplitOptions.None);
                foreach (var condition in conditions)
                {
                    var parts = condition.Split(new[] { '=' }, 2);
                    if (parts.Length == 2)
                    {
                        var propertyName = parts[0].Trim().TrimStart('@');
                        var propertyValue = parts[1].Trim().Trim('\'');
                        AutomationProperty property = GetAutomationPropertyByName(propertyName);
                        if (property != null)
                        {
                            propertyConditions[property] = propertyValue;
                        }
                    }
                }
            }

            return propertyConditions;
        }

        /// <summary>
        /// Gets an <see cref="AutomationProperty"/> by its name.
        /// </summary>
        /// <param name="name">The name of the property.</param>
        /// <returns>The <see cref="AutomationProperty"/> if found; otherwise, null.</returns>
        protected AutomationProperty GetAutomationPropertyByName(string name)
        {
            string propertyName = $"{name}Property";
            foreach (var field in typeof(AutomationElement).GetFields(System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public))
            {
                if (field.GetValue(null) is AutomationProperty propertyInfo)
                {
                    if (field.Name == propertyName) return propertyInfo;
                }
            }
            return null;
        }

        /// <summary>
        /// Gets the control type from its name.
        /// </summary>
        /// <param name="name">The name of the control type.</param>
        /// <returns>The <see cref="ControlType"/>.</returns>
        protected ControlType GetControlTypeFromName(string name)
        {
            // ControlType obj are available as fields (variables) of ControlType
            string progName = $"ControlType.{name}";
            foreach (var field in typeof(ControlType).GetFields(System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public))
            {
                // Foreach ControlType field, check the ProgrammaticName
                if (field.GetValue(null) is ControlType controlType && controlType.ProgrammaticName == progName)
                {
                    return controlType;
                }
            }
            return null;
        }

        /// <summary>
        /// Finds all elements including hidden ones based on the specified scope and condition.
        /// </summary>
        /// <param name="parent">The parent <see cref="AutomationElement"/>.</param>
        /// <param name="scope">The <see cref="TreeScope"/> to search within.</param>
        /// <param name="condition">The <see cref="ICustomCondition"/> to match.</param>
        /// <returns>A list of <see cref="AutomationElement"/> that match the condition.</returns>
        internal List<AutomationElement> FindAllWithHidden(AutomationElement parent, TreeScope scope, ICustomCondition condition)
        {
            AutomationElementFinder finder = new AutomationElementFinder(parent);
            List<AutomationElement> found = finder.FindAll(scope, condition);

            // Hidden
            var childrenWindows = this.GetChildrenWindows(parent, scope == TreeScope.Children);
            foreach (AutomationElement element in childrenWindows)
            {
                if (!ListContainsElement(found, element))
                {
                    AutomationElementFinder hiddenFinder = new AutomationElementFinder(element);
                    var descendants = finder.FindAll(scope, condition);
                    this.MergeLists(ref found, descendants);
                }
            }

            return found;
        }

        /// <summary>
        /// Finds all the elements where the given property is of the expected value.
        /// </summary>
        /// <param name="property">Property to check.</param>
        /// <param name="value">Expected value for the property.</param>
        /// <param name="scope">Optional. Specifies the search scope. Default to Children.</param>
        /// <returns>The found <see cref="AutomationElementWrapper"/>s as children of the current instance.</returns>
        public List<AutomationElementWrapper> FindElements(AutomationProperty property, object value, TreeScope scope = TreeScope.Children)
        {
            ICustomCondition condition = this.CreateCustomPropertyCondition(property, value);
            List<AutomationElement> found = null;
            if (this.Config.FindWaitForElement) found = this.RetryFindElement(() => {
                var _found = this.FindAllWithHidden(this.Element, scope, condition);
                if (_found.Count == 0) return null;
                else return _found;
            });
            else found = this.FindAllWithHidden(this.Element, scope, condition);

            if (found != null && found.Count >= 0)
            {
                return found.Cast<AutomationElement>().Select(element => new AutomationElementWrapper(element, this)).ToList();
            }
            else
            {
                if (this.Config.ShowError) Console.Error.WriteLine($"Failed to find the AutomationElement using {property.ProgrammaticName} = {value} inside {this}");
                return null;
            }
        }

        /// <summary>
        /// Finds all the elements that match all the conditions.
        /// </summary>
        /// <param name="propertyConditions">Collection of condition to match.</param>
        /// <param name="scope">Optional. Specifies the search scope. Default to Children.</param>
        /// <returns>The found <see cref="AutomationElementWrapper"/>s as children of the current instance.</returns>
        public List<AutomationElementWrapper> FindElements(Dictionary<AutomationProperty, object> propertyConditions, TreeScope scope = TreeScope.Children)
        {
            if (propertyConditions.Count <= 0) throw new ArgumentException("Invalid condition.", nameof(propertyConditions));
            if (propertyConditions.Count == 1) return this.FindElements(propertyConditions.First().Key, propertyConditions.First().Value, scope);

            List<ICustomCondition> conditions = new List<ICustomCondition>();

            foreach (var propertyCondition in propertyConditions)
            {
                conditions.Add(this.CreateCustomPropertyCondition(propertyCondition.Key, propertyCondition.Value));
            }

            var combinedCondition = new CustomAndCondition(conditions.ToArray());
            List<AutomationElement> found = null;
            if (this.Config.FindWaitForElement) found = this.RetryFindElement(() => {
                var _found = this.FindAllWithHidden(this.Element, scope, combinedCondition);
                if (_found.Count == 0) return null;
                else return _found;
            });
            else found = this.FindAllWithHidden(this.Element, scope, combinedCondition);

            if (found != null && found.Count >= 0)
            {
                return found.Cast<AutomationElement>().Select(element => new AutomationElementWrapper(element, this)).ToList();
            }
            else
            {
                if (this.Config.ShowError) Console.Error.WriteLine($"Failed to find the AutomationElement using the provided conditions inside {this}");
                return null;
            }
        }

        /// <summary>
        /// Finds all the elements that match the specified expression.  
        /// It only looks at nested elements inside the current <see cref="AutomationElementWrapper"/>, without any "path".
        /// </summary>
        /// <param name="expr">Expression as XPath condition: "ControlType[@attribute='value' and @attribute2='value' ...]".  
        /// Example: Edit[@HelpText='foo']</param>
        /// <param name="scope">Optional. Specifies the search scope. Default to Children.</param>
        /// <returns>The found <see cref="AutomationElementWrapper"/>s as children of the current instance.</returns>
        public List<AutomationElementWrapper> FindElements(string expr, TreeScope scope = TreeScope.Children)
        {
            var propertyConditions = this.ParseConditionExpression(expr);
            return this.FindElements(propertyConditions, scope);
        }

        /// <summary>
        /// Finds the first element that matches the specified XPath expression.
        /// </summary>
        /// <param name="xpath">XPath expression to find the element.</param>
        /// <returns>The found <see cref="AutomationElementWrapper"/>. All required instances will be created according to the path.</returns>
        public AutomationElementWrapper FindElementByXPath(string xpathExpr)
        {
            // Remove white-spaces
            xpathExpr = xpathExpr.Trim();

            // Define the regex pattern
            string pattern = @"^\.?\/{1,2}\w+(\[.*?\])?(\/{1,2}\w+(\[.*?\])?)*$";

            // Check if the xpathExpr matches the pattern
            if (!Regex.IsMatch(xpathExpr, pattern))
            {
                if (this.Config.ShowError) throw new ArgumentException("Invalid XPath expression format.", nameof(xpathExpr));
                else return null;
            }

            // If xpathExpr starts with . then we search from this else from root
            bool isRelativeSearch = xpathExpr.StartsWith(".");
            // Split the rest by path / (immediate child, Children scope) or // (anywhere, Descendants scope)
            var matches = Regex.Matches(xpathExpr, @"(/{1,2})([^\[\]/]+(\[[^\]]*\])?)");
            (TreeScope scope, string node)[] nodes = matches.Cast<Match>().Select(match =>
            {
                string delimiter = match.Groups[1].Value;
                string node = match.Groups[2].Value;
                TreeScope scope = delimiter == "//" ? TreeScope.Descendants : TreeScope.Children;
                string cleanNode = node.TrimStart('/');
                return (scope, cleanNode);
            }).ToArray();

            // The found (leaf) element must respect the path
            // Depth-First Search using FindElements to match each node of the path
            AutomationElementWrapper startingElement = isRelativeSearch ? this : (this.Config.SearchInAllTopWindows ? null : this.RootWindow);
            AutomationElementWrapper result = null;
            if (this.Config.FindWaitForElement) result = this.RetryFindElement(() =>
            {
                if (this.Config.SearchInAllTopWindows) return this.ExecuteFindFromTopLevels(nodes);
                else return FindElementByXPathRecursive(startingElement, nodes, 0);
            });
            else 
            {
                if (this.Config.SearchInAllTopWindows) result = this.ExecuteFindFromTopLevels(nodes);
                else result = FindElementByXPathRecursive(startingElement, nodes, 0);
            }

            if (result == null && this.Config.ShowError) Console.Error.WriteLine($"Element not found using expression {xpathExpr} inside {this}");
            return result;
        }

        private AutomationElementWrapper ExecuteFindFromTopLevels((TreeScope scope, string node)[] nodes)
        {
            AutomationElementWrapper found = null;

            WindowsAPIHelper.VisitTopWindows((AutomationElement element) => {
                AutomationElementWrapper wrapper = new AutomationElementWrapper(element);
                found = this.FindElementByXPathRecursive(wrapper, nodes, 0);
                return found == null;
            });

            return found;
        }

        /// <summary>
        /// Recursively finds an element by its XPath.
        /// </summary>
        /// <param name="currentElement"><see cref="AutomationElementWrapper"/> where to start the search.</param>
        /// <param name="nodes">Condition for each part of the path with its scope.</param>
        /// <param name="index">Index of the current recurse.</param>
        /// <returns>The found <see cref="AutomationElementWrapper"/>. All required instances will be created according to the path.</returns>
        private AutomationElementWrapper FindElementByXPathRecursive(AutomationElementWrapper currentElement, (TreeScope scope, string node)[] nodes, int index)
        {
            if (currentElement == null) return null;

            var savedConfig = this.Config.Clone();
            currentElement.Config.ShowError = false;    // don't display/throw error
            currentElement.Config.FindWaitForElement = false;   // don't retry

            if (index >= nodes.Length)
            {
                // At leaf, returns it if it matches
                var leafCond = nodes.Last();
                // Okay if match not null but don't use it! As the Parent link is wrong due to TreeScope.Element (itself)
                var match = currentElement.FindElement(leafCond.node, TreeScope.Element);
                currentElement.Config.CopyFrom(savedConfig);
                if (match != null) return currentElement;
                else return null;
            }

            var (scope, node) = nodes[index];
            var elements = currentElement.FindElements(node, scope);
            currentElement.Config.CopyFrom(savedConfig);

            if (elements == null || elements.Count == 0)
            {
                return null;
            }

            foreach (var element in elements)
            {
                var foundElement = FindElementByXPathRecursive(element, nodes, index + 1);
                if (foundElement != null)
                {
                    return foundElement;
                }
            }

            return null;
        }

        /// <summary>
        /// Converts an <see cref="AutomationElementCollection"/> to a <see cref="List{AutomationElement}"/>.
        /// </summary>
        /// <param name="collection">The <see cref="AutomationElementCollection"/> to convert.</param>
        /// <returns>A <see cref="List{AutomationElement}"/> containing the elements from the collection.</returns>
        protected List<AutomationElement> AutomationCollectionToList(AutomationElementCollection collection)
        {
            List<AutomationElement> result = new List<AutomationElement>();
            foreach (AutomationElement element in collection)
            {
                result.Add(element);
            }

            return result;
        }

        /// <summary>
        /// Writes the details of the specified elements to the console.
        /// </summary>
        /// <param name="elements">The list of <see cref="AutomationElement"/> to write.</param>
        private void PrintElements(List<AutomationElement> elements)
        {
            foreach (AutomationElement element in elements)
            {
                Console.WriteLine($"> Name: {element.Current.Name}, ControlType: {element.Current.ControlType.ProgrammaticName}, AutomationId: {element.Current.AutomationId}");
            }
        }

        /// <summary>
        /// Gets the children windows using Windows API.  
        /// This allows to find the "hidden" descendants which are not available with the "Find" of <see cref="AutomationElement"/>.
        /// </summary>
        /// <param name="parent">The parent <see cref="AutomationElement"/>.</param>
        /// <param name="immediateChild"></param>
        /// <returns>The list of <see cref="AutomationElement"/> added to "result". (delta)</returns>
        private List<AutomationElement> GetChildrenWindows(AutomationElement parent, bool immediateChild)
        {
            List<AutomationElement> result = new List<AutomationElement>();
            IntPtr parentHwnd = new IntPtr(parent.Current.NativeWindowHandle);
            if (this.Config.ForceChildrenWindows || parentHwnd != IntPtr.Zero)  // Do not add if no handle, else it will take all the windows of Windows
            {
                WindowsAPIHelper.VisitWindows(parentHwnd, immediateChild, (childWindowElt) => {
                    result.Add(childWindowElt);
                    return true;
                });
            }
            return result;
        }

        /// <summary>
        /// Print all the children of the current <see cref="AutomationElement"/>.
        /// Useful for debugging.
        /// </summary>
        public void PrintAllChildren()
        {
            Console.WriteLine($"Children of {this}");
            var found = this.FindAllWithHidden(this.Element, TreeScope.Children, CustomBoolCondition.TrueCondition);
            if (found.Count == 0) Console.WriteLine("> No children found");
            else
            {
                this.PrintElements(found);
            }
        }

        /// <summary>
        /// Displays all descendants of the current <see cref="AutomationElement"/>.
        /// Useful for debugging.
        /// </summary>
        public void PrintAllDescendants()
        {
            Console.WriteLine($"Descendants of {this}");
            var found = this.FindAllWithHidden(this.Element, TreeScope.Descendants, CustomBoolCondition.TrueCondition);
            if (found.Count == 0) Console.WriteLine("> No descendants found");
            else
            {
                this.PrintElements(found);
            }
        }

        /// <summary>
        /// Merges two lists of <see cref="AutomationElement"/>.
        /// </summary>
        /// <param name="dest">The result list to merge into.</param>
        /// <param name="source">The collection to merge from.</param>
        /// <returns>The result list.</returns>
        private List<AutomationElement> MergeLists(ref List<AutomationElement> dest, List<AutomationElement> source)
        {
            foreach (AutomationElement element in source)
            {
                if (!this.ListContainsElement(dest, element)) dest.Add(element);
            }
            return dest;
        }

        /// <summary>
        /// Checks if a list contains a specific <see cref="AutomationElement"/>.
        /// Note that it forces unique RuntimeId and AutomationId within the list.
        /// </summary>
        /// <param name="list">The list to check.</param>
        /// <param name="element">The element to check for.</param>
        /// <returns>True if the list contains the element; otherwise, false.</returns>
        private bool ListContainsElement(List<AutomationElement> list, AutomationElement element)
        {
            int[] id = element.GetRuntimeId();
            string automationId = element.Current.AutomationId;
            var found = list.Find((current) => {
                int[] curId = current.GetRuntimeId();
                if (id.SequenceEqual(curId)) return true;
                else
                {
                    string currentAutomationId = current.Current.AutomationId;
                    if (!String.IsNullOrEmpty(automationId) && automationId == currentAutomationId) return true;
                    else return false;
                }
            });

            return found != null;
        }

        /// <summary>
        /// Gets the notification code for the current <see cref="AutomationElement"/>.  
        /// Used by <see cref="NotifyParent"/> to identify the current element.
        /// </summary>
        /// <returns>The notification code.</returns>
        protected int GetNotificationCode()
        {
            if (this.Element != null)
            {
                // Determine the control type and return the appropriate notification code
                ControlType controlType = this.Element.Current.ControlType;
                if (controlType == ControlType.Button || controlType == ControlType.CheckBox || controlType == ControlType.RadioButton || controlType == ControlType.SplitButton)
                {
                    return WindowsAPIHelper.BN_CLICKED;
                }
                else if (controlType == ControlType.List || controlType == ControlType.DataGrid
                    || controlType == ControlType.Hyperlink || controlType == ControlType.ListItem || controlType == ControlType.DataItem)
                {
                    return WindowsAPIHelper.LBN_SELCHANGE;
                }
                else if (controlType == ControlType.ComboBox)
                {
                    return WindowsAPIHelper.CBN_SELCHANGE;
                }
                else if (controlType == ControlType.Edit || controlType == ControlType.Document)
                {
                    return WindowsAPIHelper.EN_CHANGE;
                }
                else if (controlType == ControlType.Tab || controlType == ControlType.TabItem)
                {
                    return WindowsAPIHelper.TCN_SELCHANGE;
                }
                else if (controlType == ControlType.Menu || controlType == ControlType.MenuItem || controlType == ControlType.MenuBar)
                {
                    return WindowsAPIHelper.MN_SELECTITEM;
                }
                else if (controlType == ControlType.Slider)
                {
                    return WindowsAPIHelper.TBM_SETPOS;
                }
                else if (controlType == ControlType.ScrollBar)
                {
                    return WindowsAPIHelper.SBM_SETPOS;
                }
                else if (controlType == ControlType.Tree || controlType == ControlType.TreeItem)
                {
                    return WindowsAPIHelper.TVN_SELCHANGED;
                }
                else if (controlType == ControlType.Image)
                {
                    return WindowsAPIHelper.STN_CLICKED;
                }
                else if (controlType == ControlType.ToolTip)
                {
                    return WindowsAPIHelper.TTN_SHOW;
                }
                else if (controlType == ControlType.StatusBar)
                {
                    return WindowsAPIHelper.SB_SETTEXT;
                }
                else if (controlType == ControlType.ProgressBar)
                {
                    return WindowsAPIHelper.PBM_SETPOS;
                }
                else if (controlType == ControlType.Spinner)
                {
                    return WindowsAPIHelper.UDN_DELTAPOS;
                }
                else if (controlType == ControlType.Header || controlType == ControlType.HeaderItem)
                {
                    return WindowsAPIHelper.HDN_ITEMCHANGED;
                }
                else if (controlType == ControlType.ToolBar)
                {
                    return WindowsAPIHelper.TBN_DROPDOWN;
                }
                else if (controlType == ControlType.Pane || controlType == ControlType.Window || controlType == ControlType.TitleBar || controlType == ControlType.Separator || controlType == ControlType.Group || controlType == ControlType.Custom)
                {
                    return WindowsAPIHelper.WM_USER;
                }
            }
            // Default notification code if control type is not specifically handled
            return 0;
        }

        /// <summary>
        /// Hovers over the current <see cref="AutomationElement"/> at the specified location.
        /// </summary>
        /// <param name="x">X</param>
        /// <param name="y">Y</param>
        /// <returns>The current <see cref="AutomationElementWrapper"/> instance.</returns>
        public AutomationElementWrapper Hover(int x, int y)
        {
            IntPtr lParam = WindowsAPIHelper.MakeLParam(x, y);
            IntPtr hwnd = this.GetHwnd();
            this.SendMessageWin32(hwnd, WindowsAPIHelper.WM_MOUSEMOVE, IntPtr.Zero, lParam);

            this.NotifyParent(lParam);

            return this;
        }

        /// <summary>
        /// Hovers over the current <see cref="AutomationElement"/>.
        /// </summary>
        /// <returns>The current <see cref="AutomationElementWrapper"/> instance.</returns>
        public AutomationElementWrapper Hover()
        {
            Point location = this.GetHeadlessClickablePoint();
            Point relativeLoc = this.GetRelativeLocation(location);
            return this.Hover((int)relativeLoc.X, (int)relativeLoc.Y);
        }

        /// <summary>
        /// Tries to notify the parent of an event.
        /// </summary>
        /// <param name="lParam">Optional. lParam for the notify.</param>
        protected void NotifyParent(IntPtr lParam = default)
        {
            // Send WM_COMMAND message to the parent window
            if (this.Parent != null)
            {
                try
                {
                    int notificationCode = this.GetNotificationCode();
                    IntPtr parentHwnd = this.Parent.GetHwnd();
                    this.PostMessageWin32(parentHwnd, WindowsAPIHelper.WM_COMMAND, WindowsAPIHelper.MakeWParam(this.GetIdOfControl(), notificationCode), lParam);
                }
                catch (Exception e)
                {
                    if (this.Config.ShowError) Console.Error.WriteLine($"Failed to notify the parent, {e}");
                }
            }
        }

        /// <summary>
        /// Retries a specified find until it succeeds or a timeout is reached.
        /// </summary>
        /// <typeparam name="T">What to find.</typeparam>
        /// <param name="callback">Function executing the find then returning the result. The find is successful if the result is not null.</param>
        /// <returns>The result of the find.</returns>
        protected T RetryFindElement<T>(Func<T> callback)
        {
            int timeCounter = 0;
            const int sleepTime = 500;
            while (timeCounter <= this.Config.WaitTimeoutMS)
            {
                try
                {
                    T result = callback();
                    if (result != null) return result;
                }
                catch (Exception e) { }
                System.Threading.Thread.Sleep(sleepTime);
                timeCounter += sleepTime;
            }
            if (this.Config.ShowError) Console.Error.WriteLine($"Failed to find the Element during {this.Config.WaitTimeoutMS} ms");
            return default(T);
        }

        /// <summary>
        /// Finds an element owned by the current Window Handle that matches the specified conditions.
        /// </summary>
        /// <param name="rootCondition">The XPath condition to match the owned window.</param>
        /// <param name="subExpr">Optional. The XPath to match the sub-element inside the given owned window.</param>
        /// <returns>An <see cref="AutomationElementWrapper"/> for the owned element if found; otherwise, null.</returns>
        public AutomationElementWrapper FindOwnedElement(string rootCondition, string subExpr = "")
        {
            AutomationElementWrapper Aux()
            {
                AutomationElementWrapper found = null;
                WindowsAPIHelper.VisitTopWindows((AutomationElement element) => {
                    AutomationElementWrapper wrapper = new AutomationElementWrapper(element);
                    wrapper.Config.ShowError = false;
                    wrapper.Config.FindWaitForElement = false;

                    var rootMatches = wrapper.FindElement(rootCondition, TreeScope.Element);  // Check self
                    if (rootMatches != null)
                    {
                        // We've found a matching top level owned window
                        if (String.IsNullOrEmpty(subExpr))
                        {
                            wrapper.Config.CopyFrom(this.Config);
                            found = wrapper;
                            return false;
                        }
                        else
                        {
                            // Check for sub element with xpath
                            var subElement = wrapper.FindElementByXPath(subExpr);
                            if (subElement != null)
                            {
                                subElement.Config.CopyFrom(this.Config);
                                found = subElement;
                                return false;
                            }
                        }
                    }
                    return true;
                });
                return found;
            }

            AutomationElementWrapper result = null;
            if (this.Config.FindWaitForElement) result = this.RetryFindElement(() => {
                return Aux();
            });
            else result = Aux();

            if (result == null && this.Config.ShowError) Console.Error.WriteLine($"Owned element not found with {rootCondition} and {subExpr} with owner {this}");
            return result;
        }

        /// <summary>
        /// Finds a top-level window that matches the specified condition.
        /// </summary>
        /// <param name="condition">The XPath condition to match the top-level window.</param>
        /// <param name="config">The configuration settings for finding the window.</param>
        /// <returns>An <see cref="AutomationElementWrapper"/> for the top-level window if found; otherwise, null.</returns>
        public static AutomationElementWrapper FindTopLevelWindow(string condition, Config config)
        {
            AutomationElementWrapper Aux()
            {
                AutomationElementWrapper _result = null;
                WindowsAPIHelper.VisitTopWindows((element) =>
                {
                    var wrapper = new AutomationElementWrapper(element);
                    wrapper.Config.ShowError = false;
                    wrapper.Config.FindWaitForElement = false;
                    var found = wrapper.FindElement(condition, TreeScope.Element);
                    if (found != null)
                    {
                        wrapper.Config.CopyFrom(config);
                        _result = found;
                    }
                    return true;
                });
                return _result;
            }

            AutomationElementWrapper result = Aux();
            if (config.FindWaitForElement && result == null)
            {
                var dummy = new AutomationElementWrapper(null);
                dummy.Config.CopyFrom(config);
                result = dummy.RetryFindElement(() =>
                {
                    return Aux();
                });
            }
            if (result == null && config.ShowError) Console.Error.WriteLine($"Top level window not found with {condition}");
            return result;
        }

        /// <summary>
        /// Retrieves the background color of the associated automation element.
        /// </summary>
        /// <returns>
        /// A <see cref="System.Drawing.Color"/> structure that represents the background color of the associated automation element.
        /// Returns <see cref="System.Drawing.Color.Empty"/> upon errors.
        /// </returns>
        /// <remarks>
        /// This method uses the <see cref="WindowsAPIHelper.GetBackgroundColor"/> method to retrieve the background color
        /// of the window associated with the automation element.
        /// </remarks>
        public System.Drawing.Color GetBackgroundColor()
        {
            return WindowsAPIHelper.GetBackgroundColor(this.GetHwnd());
        }

        /// <summary>
        /// Takes a screenshot of the current window and saves it to the specified file path.
        /// </summary>
        /// <param name="filePath">The file path to save the screenshot.</param>
        /// <param name="format">The format of the image.</param>
        public void TakeScreenshot(string filePath, ImageFormat format)
        {
            IntPtr hWnd = this.GetHwnd();
            if (hWnd == IntPtr.Zero)
            {
                if (this.Config.ShowError) Console.Error.WriteLine("Cannot take screenshot, window handle is not defined.");
                return;
            }

            WindowsAPIHelper.RECT rect;
            if (!WindowsAPIHelper.GetWindowRect(hWnd, out rect))
            {
                if (this.Config.ShowError) Console.Error.WriteLine("Failed to get window rectangle.");
                return;
            }

            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            using (System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(width, height, PixelFormat.Format32bppArgb))
            {
                using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(bitmap))
                {
                    IntPtr hdcBitmap = graphics.GetHdc();
                    IntPtr hdcWindow = WindowsAPIHelper.GetWindowDC(hWnd);

                    WindowsAPIHelper.BitBlt(hdcBitmap, 0, 0, width, height, hdcWindow, 0, 0, WindowsAPIHelper.SRCCOPY);

                    graphics.ReleaseHdc(hdcBitmap);
                    WindowsAPIHelper.ReleaseDC(hWnd, hdcWindow);
                }

                bitmap.Save(filePath, format);
            }
            Console.WriteLine($"Screenshot saved at '{filePath}'");
        }
    }

    /// <summary>
    /// Simulates keyboard interactions with the associated element.  
    /// The keyboard requires the associated element to be visible! Not headless.
    /// The methods use "PostMessage" instead of "SendMessage" to not lock the process after some keys input (like a shortcut opening a menu).  
    /// "PostMessage" does not wait for the result.
    /// </summary>
    public class Keyboard
    {
        /// <summary>
        /// Associated <see cref="AutomationElementWrapper"/>.
        /// </summary>
        private readonly AutomationElementWrapper _elementWrapper;

        /// <summary>
        /// Initializes a new instance of the <see cref="Keyboard"/> class.
        /// </summary>
        /// <param name="elementWrapper">The <see cref="AutomationElementWrapper"/> associated with this keyboard.</param>
        public Keyboard(AutomationElementWrapper elementWrapper)
        {
            _elementWrapper = elementWrapper;
        }

        /// <summary>
        /// <see cref="Keyboard"/> options.
        /// </summary>
        public class Options
        {
            /// <summary>
            /// Gets or sets a value indicating whether to automatically call <see cref="AutomationElementWrapper.SetFocus"/>.
            /// </summary>
            public bool SetFocus { get; set; } = false;
            /// <summary>
            /// Gets or sets a value indicating if a system key is currently pressed.
            /// </summary>
            public bool IsSystemKey { get; set; } = false;
        }

        /// <summary>
        /// Get the lparam for the keyboard messages.  
        /// Cf. doc like https://learn.microsoft.com/en-us/windows/win32/inputdev/wm-keydown.  
        /// Main difference is system message or not.
        /// </summary>
        /// <param name="key">The virtual-key code of the key.</param>
        /// <param name="message">The message being sent.</param>
        /// <returns>LParam flag</returns>
        private IntPtr GetLParam(Keys key, uint message)
        {
            bool isSystem = message == WindowsAPIHelper.WM_SYSKEYDOWN || message == WindowsAPIHelper.WM_SYSKEYUP || message == WindowsAPIHelper.WM_SYSCHAR || message == WindowsAPIHelper.WM_SYSDEADCHAR;
            bool isUp = message == WindowsAPIHelper.WM_KEYUP || message == WindowsAPIHelper.WM_SYSKEYUP;

            const int initialValue = 0;
            int repeatMask = 0x01;  // 1 = message sent once
            int scanCodeMask = 0;
            // Extended key flag bit 24 to 1, same no matter the message
            int extendedKeyMask = WindowsAPIHelper.IsExtendedKey(key) ? (0x01 << 24) : 0;
            // Context code bit 29 must be 1 on ALT (upon system)
            int contextCodeMask = isSystem ? (0x01 << 29) : 0;
            // Previous state bit 30, 1 for UP
            int previousStateMask = isUp ? (0x01 << 30) : 0;
            // Transition state bit 31, 1 for UP
            int transitionStateMask = isUp ? (0x01 << 31) : 0;
            int value = initialValue | repeatMask | scanCodeMask | extendedKeyMask | contextCodeMask | previousStateMask | transitionStateMask;
            return new IntPtr(value);
        }

        /// <summary>
        /// Sends a key down event.
        /// </summary>
        /// <param name="keyCode">Virtual key code of the key.</param>
        /// <param name="isSystem">True if a system key is currently pressed.</param>
        private void SendKeyDown(Keys keyCode, bool isSystem)
        {
            IntPtr keyPtr = new IntPtr((uint)keyCode);
            uint msg = isSystem ? WindowsAPIHelper.WM_SYSKEYDOWN : WindowsAPIHelper.WM_KEYDOWN;
            this._elementWrapper.PostMessageWin32(msg, keyPtr, this.GetLParam(keyCode, msg));
        }

        /// <summary>
        /// Sends a key up event.
        /// </summary>
        /// <param name="keyCode">Virtual key code of the key</param>
        /// <param name="isSystem">True if a system key is currently pressed.</param>
        private void SendKeyUp(Keys keyCode, bool isSystem)
        {
            IntPtr keyPtr = new IntPtr((uint)keyCode);
            uint msg = isSystem ? WindowsAPIHelper.WM_SYSKEYUP : WindowsAPIHelper.WM_KEYUP;
            this._elementWrapper.PostMessageWin32(msg, keyPtr, this.GetLParam(keyCode, msg));
        }

        /// <summary>
        /// Simulates a key press. It will send key down and key up messages, the char message is handled by translate message.
        /// </summary>
        /// <param name="keyCode">Virtual key code to press.</param>
        /// <param name="options">Optional options.</param>
        /// <returns>The associated <see cref="AutomationElementWrapper"/>.</returns>
        public AutomationElementWrapper PressKey(Keys keyCode, Options options = null)
        {
            if (options == null) options = new Options();
            if (options.SetFocus) this._elementWrapper.SetFocus();

            bool isSystem = options.IsSystemKey || WindowsAPIHelper.IsSystemKey(keyCode);

            // Key down
            this.SendKeyDown(keyCode, isSystem);

            // Key up
            this.SendKeyUp(keyCode, isSystem);

            return this._elementWrapper;
        }

        /// <summary>
        /// Simulates a sequential keys press.
        /// </summary>
        /// <param name="options"><see cref="Options"/> options.</param>
        /// <param name="keys">Virtual key codes to press.</param>
        /// <returns>The associated <see cref="AutomationElementWrapper"/>.</returns>
        /// <exception cref="ArgumentException">if no keys are provided.</exception>
        public AutomationElementWrapper PressKeys(Options options, params Keys[] keys)
        {
            if (keys == null || keys.Length == 0)
            {
                throw new ArgumentException("At least one key must be provided.", nameof(keys));
            }
            if (options == null) options = new Options();
            if (options.SetFocus) this._elementWrapper.SetFocus();

            for (int i = 0; i < keys.Length; i++)
            {
                var key = keys[i];
                this.PressKey(key, new Options { SetFocus = false, IsSystemKey = options.IsSystemKey });
            }

            return this._elementWrapper;
        }

        /// <summary>
        /// Simulates a sequential keys press.
        /// </summary>
        /// <param name="keys">Virtual key codes to press.</param>
        /// <returns>The associated <see cref="AutomationElementWrapper"/>.</returns>
        public AutomationElementWrapper PressKeys(params Keys[] keys)
        {
            return this.PressKeys(null, keys);
        }

        /// <summary>
        /// Holds a key.
        /// </summary>
        /// <param name="keyCode">Virtual key code to hold.param>
        /// <param name="options"><see cref="Options"/> options.</param>
        /// <returns>The associated <see cref="AutomationElementWrapper"/>.</returns>
        public AutomationElementWrapper HoldKey(Keys keyCode, Options options = null)
        {
            if (options == null) options = new Options();
            if (options.SetFocus) this._elementWrapper.SetFocus();

            bool isSystem = options.IsSystemKey || WindowsAPIHelper.IsSystemKey(keyCode);

            // Key down
            this.SendKeyDown(keyCode, isSystem);

            return this._elementWrapper;
        }

        /// <summary>
        /// Releases a key.
        /// </summary>
        /// <param name="keyCode">Virtual key code to hold.param>
        /// <param name="options"><see cref="Options"/> options.</param>
        /// <returns>The associated <see cref="AutomationElementWrapper"/>.</returns>
        public AutomationElementWrapper ReleaseKey(Keys keyCode, Options options = null)
        {
            if (options == null) options = new Options();
            if (options.SetFocus) this._elementWrapper.SetFocus();

            bool isSystem = options.IsSystemKey;

            // Key up
            this.SendKeyUp(keyCode, isSystem);

            return this._elementWrapper;
        }

        /// <summary>
        /// Holds then releases keys sequentially.  
        /// The release is made in reverse order from the last hold key.
        /// </summary>
        /// <param name="options"><see cref="Options"/> options.</param>
        /// <param name="keys">Virtual key codes to hold then release.</param>
        /// <returns>The associated <see cref="AutomationElementWrapper"/>.</returns>
        public AutomationElementWrapper HoldAndReleaseKeys(Options options, params Keys[] keys)
        {
            // Keys combination, hold everything then release in reverse order
            if (keys == null || keys.Length == 0)
            {
                throw new ArgumentException("At least one key must be provided.", nameof(keys));
            }
            if (options == null) options = new Options();
            if (options.SetFocus) this._elementWrapper.SetFocus();

            // Hold each key
            int altIndex = -1;
            for (int i = 0; i < keys.Length; i++)
            {
                var key = keys[i];
                if (altIndex == -1 && WindowsAPIHelper.IsAltKey(key)) altIndex = i;
                this.HoldKey(key, new Options { SetFocus = false, IsSystemKey = altIndex >= 0 });
            }

            // Release the keys in reverse order
            for (int i = keys.Length - 1; i >= 0; i--)
            {
                // keyup on alt
                this.ReleaseKey(keys[i], new Options { SetFocus = false, IsSystemKey = altIndex >= 0 && i > altIndex });
            }

            return this._elementWrapper;
        }

        /// <summary>
        /// Holds then releases keys sequentially.  
        /// The release is made in reverse order from the last hold key.
        /// </summary>
        /// <param name="keys">Virtual key codes to hold then release.</param>
        /// <returns>The associated <see cref="AutomationElementWrapper"/>.</returns>
        public AutomationElementWrapper HoldAndReleaseKeys(params Keys[] keys)
        {
            return this.HoldAndReleaseKeys(null, keys);
        }
    }

    /// <summary>
    /// Configuration settings for the <see cref="AutomationElementWrapper"/>.
    /// </summary>
    public class Config
    {
        /// <summary>
        /// Gets or sets a value indicating whether to show or throw errors.
        /// </summary>
        public bool ShowError { get; set; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether to wait for the element to be found in the context of a "find".
        /// </summary>
        public bool FindWaitForElement { get; set; } = true;

        /// <summary>
        /// Gets or sets the timeout in milliseconds for waiting for an element in the context of a "find".
        /// </summary>
        public int WaitTimeoutMS { get; set; } = 5000;

        /// <summary>
        /// Gets or sets a value indicating whether to force the inclusion of child windows.  
        /// The force bypass the window handle check. If there is no window handle then it will add all the windows of the system!
        /// </summary>
        public bool ForceChildrenWindows { get; set; } = false;

        /// <summary>
        /// Gets or sets a value indicating whether to execute the <see cref="AutomationElementWrapper.FindElementByXPath"/> using all the top level windows.
        /// This option is only available for non-relative search. Note that the find will be significantly slower.
        /// </summary>
        public bool SearchInAllTopWindows { get; set; } = false;

        /// <summary>
        /// Gets or sets a value indicating whether to check for value as Regex inside the conditions.
        /// </summary>
        public bool UseRegex { get; set; } = false;

        /// <summary>
        /// Copies the configuration settings from another <see cref="Config"/> instance.
        /// </summary>
        /// <param name="other">The other <see cref="Config"/> instance to copy from.</param>
        public void CopyFrom(Config other)
        {
            this.ShowError = other.ShowError;
            this.FindWaitForElement = other.FindWaitForElement;
            this.WaitTimeoutMS = other.WaitTimeoutMS;
            this.ForceChildrenWindows = other.ForceChildrenWindows;
            this.SearchInAllTopWindows = other.SearchInAllTopWindows;
            this.UseRegex = other.UseRegex;
        }

        /// <summary>
        /// Creates a deep copy of the current <see cref="Config"/> instance.
        /// </summary>
        /// <returns>A new <see cref="Config"/> instance with the same values.</returns>
        public Config Clone()
        {
            return new Config
            {
                ShowError = this.ShowError,
                FindWaitForElement = this.FindWaitForElement,
                WaitTimeoutMS = this.WaitTimeoutMS,
                ForceChildrenWindows = this.ForceChildrenWindows,
                SearchInAllTopWindows = this.SearchInAllTopWindows,
                UseRegex = this.UseRegex,
            };
        }
    }
}
