﻿/// <summary>
///   Original Author: Joe Zachary
///   Further Authors: H. James de St. Germain
///   
///   Dates          : 2012-ish - Original 
///                    2020     - Updated for use with ASP Core
///                    
///   This code represents a Windows Form element for a Spreadsheet
///   
///   This code is the "auto-generated" portion of the SimpleSpreadsheetGUI.
///   See the SimpleSpreadsheetGUI.cs for "hand-written" code.
///  
/// </summary>

using SpreadsheetGrid_Framework;
using System.Drawing;
using System.Windows.Forms;

namespace CS3500_Spreadsheet_GUI_Example
{
    partial class SimpleSpreadsheetGUI
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.MainControlArea = new System.Windows.Forms.FlowLayoutPanel();
            this.sample_button = new System.Windows.Forms.Button();
            this.sample_checkbox = new System.Windows.Forms.CheckBox();
            this.sample_textbox = new System.Windows.Forms.TextBox();
            this.grid_widget = new SpreadsheetGrid_Framework.SpreadsheetGridWidget();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.cellName = new System.Windows.Forms.TextBox();
            this.helpButton = new System.Windows.Forms.Button();
            this.menuStrip.SuspendLayout();
            this.MainControlArea.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            //
            // menuStrip
            //
            this.menuStrip.GripMargin = new System.Windows.Forms.Padding(2, 2, 0, 2);
            this.menuStrip.ImageScalingSize = new System.Drawing.Size(48, 48);
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Padding = new System.Windows.Forms.Padding(19, 6, 0, 6);
            this.menuStrip.Size = new System.Drawing.Size(1849, 64);
            this.menuStrip.TabIndex = 1;
            this.menuStrip.Text = "menuStrip1";
            //
            // fileToolStripMenuItem
            //
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newToolStripMenuItem,
            this.closeToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(103, 52);
            this.fileToolStripMenuItem.Text = "File";
            //
            // newToolStripMenuItem
            //
            this.newToolStripMenuItem.Name = "newToolStripMenuItem";
            this.newToolStripMenuItem.Size = new System.Drawing.Size(304, 66);
            this.newToolStripMenuItem.Text = "New";
            this.newToolStripMenuItem.Click += new System.EventHandler(this.NewToolStripMenuItem_Click);
            //
            // closeToolStripMenuItem
            //
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(304, 66);
            this.closeToolStripMenuItem.Text = "Close";
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.CloseToolStripMenuItem_Click);
            //
            // MainControlArea
            //
            this.MainControlArea.AutoSize = true;
            this.MainControlArea.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.MainControlArea.BackColor = System.Drawing.Color.Coral;
            this.MainControlArea.Controls.Add(this.sample_button);
            this.MainControlArea.Controls.Add(this.sample_checkbox);
            this.MainControlArea.Controls.Add(this.sample_textbox);
            this.MainControlArea.Controls.Add(this.cellName);
            this.MainControlArea.Controls.Add(this.helpButton);
            this.MainControlArea.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.MainControlArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.MainControlArea.Location = new System.Drawing.Point(10, 9);
            this.MainControlArea.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.MainControlArea.MinimumSize = new System.Drawing.Size(317, 285);
            this.MainControlArea.Name = "MainControlArea";
            this.MainControlArea.Size = new System.Drawing.Size(1829, 285);
            this.MainControlArea.TabIndex = 4;
            //
            // sample_button
            //
            this.sample_button.Location = new System.Drawing.Point(10, 9);
            this.sample_button.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.sample_button.Name = "sample_button";
            this.sample_button.Size = new System.Drawing.Size(238, 65);
            this.sample_button.TabIndex = 0;
            this.sample_button.Text = "button1";
            this.sample_button.UseVisualStyleBackColor = true;
            this.sample_button.Click += new System.EventHandler(this.sample_button_Click);
            //
            // sample_checkbox
            //
            this.sample_checkbox.AutoSize = true;
            this.sample_checkbox.Location = new System.Drawing.Point(268, 9);
            this.sample_checkbox.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.sample_checkbox.Name = "sample_checkbox";
            this.sample_checkbox.Size = new System.Drawing.Size(216, 41);
            this.sample_checkbox.TabIndex = 1;
            this.sample_checkbox.Text = "checkBox1";
            this.sample_checkbox.UseVisualStyleBackColor = true;
            this.sample_checkbox.CheckedChanged += new System.EventHandler(this.sample_checkbox_CheckedChanged);
            //
            // sample_textbox
            //
            this.sample_textbox.Location = new System.Drawing.Point(504, 9);
            this.sample_textbox.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.sample_textbox.Name = "sample_textbox";
            this.sample_textbox.Size = new System.Drawing.Size(308, 44);
            this.sample_textbox.TabIndex = 2;
            this.sample_textbox.TextChanged += new System.EventHandler(this.sample_textbox_TextChanged);
            //
            // grid_widget
            //
            this.grid_widget.AutoSize = true;
            this.grid_widget.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.grid_widget.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grid_widget.Location = new System.Drawing.Point(10, 294);
            this.grid_widget.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.grid_widget.MaximumSize = new System.Drawing.Size(6650, 5692);
            this.grid_widget.Name = "grid_widget";
            this.grid_widget.Size = new System.Drawing.Size(1829, 660);
            this.grid_widget.TabIndex = 0;
            //
            // tableLayoutPanel1
            //
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.MainControlArea, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.grid_widget, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 64);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 285F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 57F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1849, 963);
            this.tableLayoutPanel1.TabIndex = 6;
            //
            // cellName
            //
            this.cellName.Location = new System.Drawing.Point(825, 3);
            this.cellName.Name = "cellName";
            this.cellName.Size = new System.Drawing.Size(184, 44);
            this.cellName.TabIndex = 3;
            this.cellName.TextChanged += new System.EventHandler(this.cellName_TextChanged);
            //
            // helpButton
            //
            this.helpButton.Location = new System.Drawing.Point(1015, 3);
            this.helpButton.Name = "helpButton";
            this.helpButton.Size = new System.Drawing.Size(155, 58);
            this.helpButton.TabIndex = 4;
            this.helpButton.Text = "Help";
            this.helpButton.UseVisualStyleBackColor = true;
            this.helpButton.Click += new System.EventHandler(this.helpButton_Click);
            //
            // SimpleSpreadsheetGUI
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(19F, 37F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1849, 1027);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.menuStrip);
            this.MainMenuStrip = this.menuStrip;
            this.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.Name = "SimpleSpreadsheetGUI";
            this.Text = "Sample GUI - Copy/Modify/Profit";
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.MainControlArea.ResumeLayout(false);
            this.MainControlArea.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion


        private SpreadsheetGridWidget grid_widget;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;

        private FlowLayoutPanel MainControlArea;
        private TableLayoutPanel tableLayoutPanel1;
        private Button sample_button;
        private CheckBox sample_checkbox;
        private TextBox sample_textbox;
        private TextBox cellName;
        private Button helpButton;
    }
}

