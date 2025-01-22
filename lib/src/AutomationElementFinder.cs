using System;
using System.Collections.Generic;
using System.Windows.Automation;

namespace HeadlessWindowsAutomation
{
    /// <summary>
    /// Provides methods to find <see cref="AutomationElement"/>s within a specified scope.
    /// </summary>
    internal sealed class AutomationElementFinder
    {
        private readonly TreeWalker _walker = TreeWalker.ControlViewWalker;
        internal AutomationElement Element { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutomationElementFinder"/> class.
        /// </summary>
        /// <param name="element">The root <see cref="AutomationElement"/> to start the search from.</param>
        internal AutomationElementFinder(AutomationElement element)
        {
            Element = element;
        }

        private bool CheckSelf(ICustomCondition condition)
        {
            return condition.Evaluate(this.Element);
        }

        private void VisitChildren(AutomationElement ofElement, Func<AutomationElement, bool> visitor)
        {
            AutomationElement child = this._walker.GetFirstChild(ofElement);
            while (child != null)
            {
                if (!visitor(child)) break;
                else
                {
                    child = this._walker.GetNextSibling(child);
                }
            }
        }

        private void VisitDescendants(AutomationElement ofElement, Func<AutomationElement, bool> visitor)
        {
            List<AutomationElement> stack = new List<AutomationElement> { ofElement };
            List<AutomationElement> queue = new List<AutomationElement>();
            bool stopLoop = false;
            bool isOfElement = true; // Don't visit the given element

            do
            {
                for (int i = 0; !stopLoop && i < stack.Count; i++)
                {
                    var element = stack[i];
                    if (!isOfElement && !visitor(element))
                    {
                        stopLoop = true;
                    }
                    else
                    {
                        if (isOfElement) isOfElement = false;

                        // Get the children to then visit them
                        this.VisitChildren(element, child =>
                        {
                            queue.Add(child);
                            return true;
                        });
                    }
                }
                if (!stopLoop)
                {
                    // Swap stack and queue, clear queue
                    (queue, stack) = (stack, queue);
                    queue.Clear();
                }
            } while (stack.Count > 0 && !stopLoop);
        }

        /// <summary>
        /// Finds the first <see cref="AutomationElement"/> that matches the specified condition within the given scope.
        /// </summary>
        /// <param name="scope">The <see cref="TreeScope"/> to search within.</param>
        /// <param name="condition">The <see cref="ICustomCondition"/> to match.</param>
        /// <returns>The first matching <see cref="AutomationElement"/> if found; otherwise, null.</returns>
        internal AutomationElement FindFirst(TreeScope scope, ICustomCondition condition)
        {
            if (scope == TreeScope.Element)
            {
                if (this.CheckSelf(condition)) return this.Element;
                else return null;
            }
            else if (scope == TreeScope.Children)
            {
                AutomationElement found = null;
                this.VisitChildren(this.Element, child => {
                    if (condition.Evaluate(child))
                    {
                        found = child;
                        return false;
                    }
                    return true;
                });
                return found;
            }
            else if (scope == TreeScope.Descendants || scope == TreeScope.Subtree)
            {
                if (scope == TreeScope.Subtree && this.CheckSelf(condition)) return this.Element;

                AutomationElement found = null;
                this.VisitDescendants(this.Element, child => {
                    if (condition.Evaluate(child))
                    {
                        found = child;
                        return false;
                    }
                    return true;
                });
                return found;
            }

            return null;
        }

        /// <summary>
        /// Finds all <see cref="AutomationElement"/>s that match the specified condition within the given scope.
        /// </summary>
        /// <param name="scope">The <see cref="TreeScope"/> to search within.</param>
        /// <param name="condition">The <see cref="ICustomCondition"/> to match.</param>
        /// <returns>A list of matching <see cref="AutomationElement"/>s.</returns>
        internal List<AutomationElement> FindAll(TreeScope scope, ICustomCondition condition)
        {
            List<AutomationElement> found = new List<AutomationElement>();

            if (scope == TreeScope.Element)
            {
                if (this.CheckSelf(condition)) found.Add(this.Element);
                else return null;
            }
            else if (scope == TreeScope.Children)
            {
                this.VisitChildren(this.Element, child => {
                    if (condition.Evaluate(child))
                    {
                        found.Add(child);
                    }
                    return true;
                });
            }
            else if (scope == TreeScope.Descendants || scope == TreeScope.Subtree)
            {
                if (scope == TreeScope.Subtree && this.CheckSelf(condition)) found.Add(this.Element);

                this.VisitDescendants(this.Element, child => {
                    if (condition.Evaluate(child))
                    {
                        found.Add(child);
                    }
                    return true;
                });
            }

            return found;
        }
    }
}