using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using TestStack.White;
using TestStack.White.Factory;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.WindowStripControls;
//using System.Windows.Forms;

namespace Automation.Core
{
    class Program
    {

        static void Main(string[] args)
        {
            Process proc = Process.GetProcessesByName("QBW32").FirstOrDefault();

            TestStack.White.Application app = TestStack.White.Application.Attach(proc);

            SearchCriteria sc = SearchCriteria.ByClassName("MauiFrame");

            var mainWindow = app.GetWindow(sc, InitializeOption.WithCache);

            mainWindow.Focus();

            SendKeys.SendWait("%+{y}{P}{U}");

            Thread.Sleep(3000);

            var modelWindows = mainWindow.ModalWindows()
                .Where(w => w.Title == "Enter Payroll Information")
                .FirstOrDefault();

            //var table = GetTableElement(modelWindows);

            //AutomationElement elm = modelWindows
            //    .GetElement(SearchCriteria.ByControlType(ControlType.HeaderItem));

            //SearchCriteria tableSearch = SearchCriteria.ByControlType(ControlType.HeaderItem);


            //var table = modelWindows.Get<TestStack.White.UIItems.TableItems.TableHeader>(tableSearch);

            var x = 0;
        }

        // -------------------------------------------------------------------
        /// <summary>
        /// Obtain the table control of interest from the target application.
        /// </summary>
        /// <param name="targetApp">
        /// The target application.
        /// </param>
        /// <returns>
        /// An AutomationElement representing a table control.
        /// </returns>
        /// -------------------------------------------------------------------
        private static AutomationElement GetTableElement(AutomationElement targetApp)
        {
            // The control type we're looking for; in this case 'Document'
            PropertyCondition cond1 =
                new PropertyCondition(
                AutomationElement.ControlTypeProperty,
                ControlType.Table);

            // The control pattern of interest; in this case 'TextPattern'.
            PropertyCondition cond2 =
                new PropertyCondition(
                AutomationElement.IsTablePatternAvailableProperty,
                true);

            AndCondition tableCondition = new AndCondition(cond1, cond2);

            AutomationElement targetTableElement =
                targetApp.FindFirst(TreeScope.Descendants, tableCondition);

            // If targetTableElement is null then a suitable table control 
            // was not found.
            return targetTableElement;
        }

    }
}
