using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace Automation.Core
{
    class Program
    {
        static void Main(string[] args)
        {
            Process proc = Process.GetProcessesByName("QBW32")[0];

            Thread.Sleep(1000);
            // Attach to the main window. 
            AutomationElement rootElement = AutomationElement.FromHandle(proc.MainWindowHandle);

            // Create the AndCondition to find the menuBar. 
            AndCondition menuBarFind =
                new AndCondition(
                    new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.MenuBar),
                    new PropertyCondition(
                        AutomationElement.AutomationIdProperty, "MenuBar"));

            // Find the menuBar, which is a child of the main window.
            AutomationElement menuBarElement = rootElement.FindFirst(TreeScope.Children, menuBarFind);

            //Create the AndCondition to find the File menu.
            AndCondition fileMenuEmployee = new AndCondition(
                new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.MenuItem),
                new PropertyCondition(AutomationElement.AutomationIdProperty, "Item 10"));

            // Find the File menu.
            AutomationElement employeeMenuElement = menuBarElement.FindFirst(TreeScope.Children, fileMenuEmployee);

            // Get the control pattern for ExpandCollapse and do the
            // Expand to get the children.

            ExpandCollapseMenuItem(employeeMenuElement);
        }

        ///--------------------------------------------------------------------
        /// <summary>
        /// Obtains an ExpandCollapsePattern control pattern from an 
        /// automation element.
        /// </summary>
        /// <param name="targetControl">
        /// The automation element of interest.
        /// </param>
        /// <returns>
        /// A ExpandCollapsePattern object.
        /// </returns>
        ///--------------------------------------------------------------------
        private static ExpandCollapsePattern GetExpandCollapsePattern(
            AutomationElement targetControl)
        {
            ExpandCollapsePattern expandCollapsePattern = null;

            try
            {
                expandCollapsePattern =
                    targetControl.GetCurrentPattern(
                    ExpandCollapsePattern.Pattern)
                    as ExpandCollapsePattern;
            }
            // Object doesn't support the ExpandCollapsePattern control pattern.
            catch (InvalidOperationException)
            {
                return null;
            }

            return expandCollapsePattern;
        }

        ///--------------------------------------------------------------------
        /// <summary>
        /// Programmatically expand or collapse a menu item.
        /// </summary>
        /// <param name="menuItem">
        /// The target menu item.
        /// </param>
        ///--------------------------------------------------------------------
        private static void ExpandCollapseMenuItem(
            AutomationElement menuItem)
        {
            if (menuItem == null)
            {
                throw new ArgumentNullException(
                    "AutomationElement argument cannot be null.");
            }

            ExpandCollapsePattern expandCollapsePattern =
                GetExpandCollapsePattern(menuItem);

            if (expandCollapsePattern == null)
            {
                return;
            }

            if (expandCollapsePattern.Current.ExpandCollapseState ==
                ExpandCollapseState.LeafNode)
            {
                return;
            }

            try
            {
                if (expandCollapsePattern.Current.ExpandCollapseState == ExpandCollapseState.Expanded)
                {
                    // Collapse the menu item.
                    expandCollapsePattern.Collapse();
                }
                else if (expandCollapsePattern.Current.ExpandCollapseState == ExpandCollapseState.Collapsed ||
                    expandCollapsePattern.Current.ExpandCollapseState == ExpandCollapseState.PartiallyExpanded)
                {
                    // Expand the menu item.
                    expandCollapsePattern.Expand();
                }
            }
            // Control is not enabled
            catch (ElementNotEnabledException e)
            {
                // TO DO: error handling.
            }
            // Control is unable to perform operation.
            catch (InvalidOperationException e)
            {
                // TO DO: error handling.
            }
        }
    }
}
