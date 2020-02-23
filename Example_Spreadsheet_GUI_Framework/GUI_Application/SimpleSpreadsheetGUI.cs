﻿/// <summary>
///   Original Author: Joe Zachary
///   Further Authors: H. James de St. Germain
///   
///   Dates          : 2012-ish - Original 
///                    2020     - Updated for use with ASP Core
///                    
///   This code represents a Windows Form element for a Spreadsheet. It includes
///   a Menu Bar with two operations (close/new) as well as the GRID of the spreadsheet.
///   The GRID is a separate class found in SpreadsheetGridWidget
///   
///   This code represents manual elements added to the GUI as well as the ability
///   to show a pop up of information, and the event handlers for a New window and to Close the window.
///
///   See the SimpleSpreadsheetGUIExample.Designer.cs for "auto-generated" code.
///   
///   This code relies on the SpreadsheetPanel "widget"
///  
/// </summary>

using SpreadsheetGrid_Framework;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace CS3500_Spreadsheet_GUI_Example
{
    public partial class SimpleSpreadsheetGUI : Form
    {
        public SimpleSpreadsheetGUI()
        {
            this.grid_widget = new SpreadsheetGridWidget();

            // Call the AutoGenerated code
            InitializeComponent();


            // Add event handler and select a start cell
            grid_widget.SelectionChanged += DisplaySelection;
            grid_widget.SetSelection(2, 3, false);


        }

        /// <summary>
        /// Given a spreadsheet, find the current selected cell and
        /// create a popup that contains the information from that cell
        /// </summary>
        /// <param name="ss"></param>
        private void DisplaySelection(SpreadsheetGridWidget ss)
        {
            int row, col;

            string value;

            ss.GetSelection(out col, out row);
            ss.GetValue(col, row, out value);

            //if (value == "")
            //{
            //    ss.SetValue(col, row, DateTime.Now.ToLocalTime().ToString("T"));
            //    ss.GetValue(col, row, out value);
            //   // MessageBox.Show("Selection: column " + col + " row " + row + " value " + value);
            //}

            cellName.Enabled = false;
            cellName.Text = getCellName(col, row);
        }

        // Deals with the New menu
        private void NewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Tell the application context to run the form on the same
            // thread as the other forms.
            Spreadsheet_Window.getAppContext().RunForm(new SimpleSpreadsheetGUI());
        }

        // Deals with the Close menu
        private void CloseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Example of how to use a button
        /// </summary>
        /// <param name="sender"> not used </param>
        /// <param name="e"> not used </param>
        private void sample_button_Click(object sender, EventArgs e)
        {
            grid_widget.SetValue(4, 5, "hello");
            grid_widget.SetSelection(4, 5);
        }

        /// <summary>
        /// Checkbox handler
        /// </summary>
        /// <param name="sender"> the checkbox (note the casting operator as)</param>
        /// <param name="e">not used</param>
        private void sample_checkbox_CheckedChanged(object sender, EventArgs e)
        {
            if ((sender as CheckBox).Checked)
            {
                grid_widget.SetValue(1, 1, "checked");
            }
            else
            {
                grid_widget.SetValue(1, 1, "not checked");
            }

        }

        /// <summary>
        /// Textbox handler
        /// </summary>
        /// <param name="sender"> the textbox </param>
        /// <param name="e">not used</param>
        private void sample_textbox_TextChanged(object sender, EventArgs e)
        {
            TextBox box = sender as TextBox;

            grid_widget.SetValue(2, 2, box.Text);

        }

        /// <summary>
        /// a helper method to get the cell name
        /// </summary>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private string getCellName(int col, int row)
        {
            string r = (row + 1).ToString(); // get the row number
            string cellName = ((char)(col + 'A')) + r; // get the cell name location

            return cellName; // TODO name normalized?
        }

        private void cellName_TextChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// a button for giving instructions on the spreadsheet and the additional featurs.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void helpButton_Click(object sender, EventArgs e)
        {
            //TODO add instructions for additional features
            string msg = "use the mouse to navigate through the spreadsheet";
            string caption = "Instructions";

            MessageBox.Show(msg, caption);
        }
    }
}
