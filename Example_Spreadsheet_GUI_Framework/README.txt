
Author: Jim de St. Germain
Date:   Spring 2020

The spreadsheet GUI code was originally written by Professor Joe Zachary.
It has been refactored and renamed to make some elements clearer.

As with most windows GUIs, the structure of the project is as follows:

   1) Main Executable Program - GUI_Application

           This is the "executable" with the Main method.  It uses the SimpleSpreadsheetGUI
           form class to build the GUI.

   2) Overarching GUI elements - SimpleSpreadsheetGUI.cs
   
           This code contains menu items, a (placeholder for the) spreadsheet grid, etc.
           It replies on the Grid from the SpreadsheetGridWidget class

   3) Specialized GUI widgets - SpreadsheetGridWidget
   
           This code builds a scrollable grid of "cells" representing the values computed
           in each cell of the spreadsheet.

Notes:

   o) (as of Spring 2020) The ASP Core designer is not powerful enough to show a visual WYSIWYG 
      rendering of the GUI elements.  This means you will have to manually add new widgets
      as necessary.

      Hint: it should be possible to use the Example_Spreadsheet_GUI_Framework example to add
      functionality using the WYSIWYG and then paste that code into the CORE version.
      