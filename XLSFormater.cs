/*
MIT License

Copyright (c) 2020 Aiekick

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSVFormater
{
    public class XLSFormater
    {
        private Excel.Application ExcelApp;
        private Excel.Workbook objBook;
        private Excel.Worksheet objSheet;
        private string currentExcelFileNameWithoutExt = "";
        
        public bool MultipleFiles;
        
        private object missing = System.Reflection.Missing.Value;
            
        /*
         * ExcelApp = new Excel.ApplicationClass();
			ExcelApp.Visible = true;
			objBook = ExcelApp.Workbooks.Add(missing);
			objSheet = (Excel.Worksheet)objBook.Sheets["Sheet1"];
			objSheet.Name = "It's Me";

			objSheet.Cells[1, 1] = "Details";
			objSheet.Cells[2, 1] = "Name : "+ array[0].ToString();
			objSheet.Cells[3, 1] = "Age : "+ array[1].ToString();
			objSheet.Cells[4, 1] = "Designation : " + array[2].ToString();
			objSheet.Cells[5, 1] = "Company : " + array[3].ToString();
			objSheet.Cells[6, 1] = "Place : " + array[4].ToString();
			objSheet.Cells[7, 1] = "Email : "+array[5].ToString();
         * 
         */
        public XLSFormater()
        {
       //     ExcelApp = new Excel.ApplicationClass();
       //     ExcelApp.Visible = true;

         //   objBook = ExcelApp.Workbooks.Add(missing);

        	MultipleFiles = false;

        }

        public void OpenExcelApp(bool AppendOnCurrentFile)
        {
        	if ( AppendOnCurrentFile == false ) 
        	{
        		if( MultipleFiles == true ) // on merge dans la meme appli
        		{
        			if ( ExcelApp == null ) 
        				ExcelApp = new Excel.ApplicationClass();
        		}
        		else
        		{
		    		ExcelApp = new Excel.ApplicationClass();
        		}
        	}
        	else
        	{
	        	try
			    {
			       ExcelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
			    }
			    catch (System.Runtime.InteropServices.COMException /*ex*/)
			    {
			       ExcelApp = new Excel.ApplicationClass();
			    }
        	}
        	ExcelApp.Visible = true;
        	if( MultipleFiles == true ) // on merge dans la meme appli
        	{
        		if ( objBook == null ) 
        			objBook = ExcelApp.Workbooks.Add(missing);
        	}
        	else
        	{
            	objBook = ExcelApp.Workbooks.Add(missing);
        	}
        }

        public void OpenXlsFile(string fPath)
        {
            ExcelApp = new Excel.ApplicationClass();
            ExcelApp.Visible = true;
            
            currentExcelFileNameWithoutExt = Path.GetFileNameWithoutExtension(fPath);
            
            // create the workbook object by opening the excel file.
            objBook = ExcelApp.Workbooks.Open(fPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        }

        public void SetActiveSheet(string name)
        {
            objSheet = (Excel.Worksheet)objBook.Sheets[name];
            objSheet.Select(missing);
        }

        public void AddSheet()
        {
            objSheet = (Excel.Worksheet)objBook.Sheets.Add(missing, missing, missing, missing);
        }

        public void RenameSheet(string name)
        {
           objSheet.Name = name;
        }
        
        public string GetCurrentFileNameWithoutExt()
        {
	  		if ( objBook == null ) return "";
        	
			return currentExcelFileNameWithoutExt;
        }
        
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////
        // LA PREMIERE CELL DE EXCEL A POUR COORD (1,1) et pas (0,0) et c'est (row, col) et pas (col, row) ////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        public void AddCell(int col, int row, string value)
        {
            if ( objSheet == null ) return;
        	
        	objSheet.Cells[row + 1, col + 1] = value;
        }

	  	public string GetCell(int col, int row)
        {
	  		if ( objSheet == null ) return "";
        	
        	Excel.Range objRange = (Excel.Range)objSheet.Cells[row + 1, col + 1]; 	
			return objRange.Text.ToString();
        }

        public void InsertColAfter(int col)
        {
            if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.Cells[1,col+1];
            Excel.Range column = rng.EntireColumn;
            column.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, false);
        }

        public void InsertRowAfter(int row)
        {
            if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.Cells[row + 1,1];
            Excel.Range Rowumn = rng.EntireRow;
            Rowumn.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
        }

        public void SetSizeOfCol(int col, int size)
        {
            if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.Cells[1, col + 1];
            rng.ColumnWidth = size;
        }

        public void SortRangeByOneColInOrder(int left, int right, int top, int bottom, int col1, bool alpha1)
        {
            if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[top + 1, left + 1], objSheet.Cells[bottom + 1, right + 1]);

            Excel.XlSortOrder sortOrder1;
            if (alpha1 == true) sortOrder1 = Excel.XlSortOrder.xlAscending;
            else sortOrder1 = Excel.XlSortOrder.xlDescending;

            rng.Sort(
                            rng.Columns[col1 + 1 - left, missing], sortOrder1, // la colonne est en relaitf au debut du range
                            missing, missing, sortOrder1,
                            missing, sortOrder1,
                            Excel.XlYesNoGuess.xlNo, missing, missing,
                            Excel.XlSortOrientation.xlSortColumns,
                            Excel.XlSortMethod.xlPinYin,
                            Excel.XlSortDataOption.xlSortNormal,
                            Excel.XlSortDataOption.xlSortNormal,
                            Excel.XlSortDataOption.xlSortNormal
                    );
        }
        
        public void SortRangeByTwoColInOrder(int left, int right, int top, int bottom, int col1, bool alpha1, int col2, bool alpha2)
        {
            if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[top + 1, left + 1], objSheet.Cells[bottom + 1, right + 1]);

            Excel.XlSortOrder sortOrder1;
            if (alpha1 == true) sortOrder1 = Excel.XlSortOrder.xlAscending;
            else sortOrder1 = Excel.XlSortOrder.xlDescending;

            Excel.XlSortOrder sortOrder2;
            if (alpha2 == true) sortOrder2 = Excel.XlSortOrder.xlAscending;
            else sortOrder2 = Excel.XlSortOrder.xlDescending;

            rng.Sort(
                            rng.Columns[col1 + 1 - left, missing], sortOrder1, // la colonne est en relaitf au debut du range
                            rng.Columns[col2 + 1 - left, missing], missing, sortOrder2,
                            missing, sortOrder1,
                            Excel.XlYesNoGuess.xlNo, missing, missing,
                            Excel.XlSortOrientation.xlSortColumns,
                            Excel.XlSortMethod.xlPinYin,
                            Excel.XlSortDataOption.xlSortNormal,
                            Excel.XlSortDataOption.xlSortNormal,
                            Excel.XlSortDataOption.xlSortNormal
                    );
        }
        
        public void SortRangeByThreeColInOrder(int left, int right, int top, int bottom, int col1, bool alpha1, int col2, bool alpha2, int col3, bool alpha3)
        {
            if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[top + 1, left + 1], objSheet.Cells[bottom + 1, right + 1]);

            Excel.XlSortOrder sortOrder1;
            if (alpha1 == true) sortOrder1 = Excel.XlSortOrder.xlAscending;
            else sortOrder1 = Excel.XlSortOrder.xlDescending;

            Excel.XlSortOrder sortOrder2;
            if (alpha2 == true) sortOrder2 = Excel.XlSortOrder.xlAscending;
            else sortOrder2 = Excel.XlSortOrder.xlDescending;

            Excel.XlSortOrder sortOrder3;
            if (alpha3 == true) sortOrder3 = Excel.XlSortOrder.xlAscending;
            else sortOrder3 = Excel.XlSortOrder.xlDescending;

            rng.Sort(
                            rng.Columns[col1 + 1 - left, missing], sortOrder1, // la colonne est en relaitf au debut du range
                            rng.Columns[col2 + 1 - left, missing], missing, sortOrder2,
                            rng.Columns[col3 + 1 - left, missing], sortOrder3,
                            Excel.XlYesNoGuess.xlNo, missing, missing,
                            Excel.XlSortOrientation.xlSortColumns,
                            Excel.XlSortMethod.xlPinYin,
                            Excel.XlSortDataOption.xlSortNormal,
                            Excel.XlSortDataOption.xlSortNormal,
                            Excel.XlSortDataOption.xlSortNormal
                    );
        }
        
        public void Replace( int row, int col, string pattern1, string pattern2)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[row + 1, col + 1], objSheet.Cells[row + 1, col + 1]);
			objSheet.Cells.Replace(pattern1, pattern2, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByColumns, false, Type.Missing, false, false);
        }
        
        public void SetRowColor(int row, int colorIndex)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.Cells[row + 1, 1];
			rng.EntireRow.Interior.ColorIndex = colorIndex;
        }
        
        public void SetColColor(int col, int colorIndex)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.Cells[1, col + 1];
			rng.EntireColumn.Interior.ColorIndex = colorIndex;
        }
        
        public void SetRangeColor(int left, int right, int top, int bottom, int colorIndex)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[top + 1, left + 1], objSheet.Cells[bottom + 1, right + 1]);
        	rng.Interior.ColorIndex = colorIndex;
        }
        
        public void GetColorIndexInNewExcelSheet()
        {
        	if ( objSheet == null ) return;
        	
        	/*	Set objExcel = CreateObject("Excel.Application")
			objExcel.Visible = True
			Set objWorkbook = objExcel.Workbooks.Add()
			Set objWorksheet = objWorkbook.Worksheets(1)
			
			For i = 1 to 14
			    objExcel.Cells(i, 1).Value = i
			    objExcel.Cells(i, 2).Interior.ColorIndex = i
			Next
			
			For i = 15 to 28
			    objExcel.Cells(i - 14, 3).Value = i
			    objExcel.Cells(i - 14, 4).Interior.ColorIndex = i
			Next
			
			For i = 29 to 42
			    objExcel.Cells(i - 28, 5).Value = i
			    objExcel.Cells(i - 28, 6).Interior.ColorIndex = i
			Next
			
			For i = 43 to 56
			    objExcel.Cells(i - 42, 7).Value = i
			    objExcel.Cells(i - 42, 8).Interior.ColorIndex = i
			Next
		*/
        }
        // BorderWeight = {1=Hairline;2=Medium;3=Thick;4=Thin}
        public void SetRangeBordure(int left, int right, int top, int bottom, bool EdgeTop, bool EdgeBottom, bool EdgeLeft, bool EdgeRight, bool insideHoriz, bool insideVert, string BorderWeight)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[top + 1, left + 1], objSheet.Cells[bottom + 1, right + 1]);
        	
        	if ( EdgeTop == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( EdgeBottom == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( EdgeLeft == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( EdgeRight == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( insideHoriz == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( insideVert == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        }
        // BorderWeight = {1=Hairline;2=Medium;3=Thick;4=Thin}
        public void SetRowBordure(int row, bool EdgeTop, bool EdgeBottom, bool EdgeLeft, bool EdgeRight, bool insideHoriz, bool insideVert, string BorderWeight)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.Cells[row + 1, 1];
        	
        	if ( EdgeTop == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( EdgeBottom == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( EdgeLeft == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( EdgeRight == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( insideHoriz == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( insideVert == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireRow.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        }
        // BorderWeight = {Hairline;Medium;Thick;Thin;None}
        public void SetColBordure(int col, bool EdgeTop, bool EdgeBottom, bool EdgeLeft, bool EdgeRight, bool insideHoriz, bool insideVert, string BorderWeight)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.Cells[1, col + 1];
        	
        	if ( EdgeTop == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( EdgeBottom == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( EdgeLeft == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( EdgeRight == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( insideHoriz == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        	if ( insideVert == true ) 
        	{
        		if ( BorderWeight == "Hairline" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlHairline;
        		if ( BorderWeight == "Medium" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium;
        		if ( BorderWeight == "Thick" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlThick;
        		if ( BorderWeight == "Thin" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlThin;
        		if ( BorderWeight == "None" ) rng.EntireColumn.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        	}
        }
        
        public void AutoFitCols(int colStart, int colEnd)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[1, colStart + 1], objSheet.Cells[1, colEnd + 1]);
        	rng.EntireColumn.AutoFit();
        }
        	
        public void AutoFitRows(int rowStart, int rowEnd)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[rowStart + 1, 1], objSheet.Cells[rowEnd + 1, 1]);
        	rng.EntireRow.AutoFit();	
        }
        
        // VerticalAlignement = {AlignBottom;AlignCenter;AlignTop}
		// HorizontalAlignement = {AlignRight;AlignLeft;AlignCenter}
        public void SetColsAlignement(int colStart, int colEnd, string VerticalAlignement, string HorizontalAlignement)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[1, colStart + 1], objSheet.Cells[1, colEnd + 1]);
        	
       		if ( VerticalAlignement == "AlignBottom" ) rng.EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
       		if ( VerticalAlignement == "AlignCenter" ) rng.EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
       		if ( VerticalAlignement == "AlignTop" ) rng.EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
       		if ( HorizontalAlignement == "AlignRight" ) rng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
       		if ( HorizontalAlignement == "AlignLeft" ) rng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
       		if ( HorizontalAlignement == "AlignCenter" ) rng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }
        // VerticalAlignement = {AlignBottom;AlignCenter;AlignTop}
		// HorizontalAlignement = {AlignRight;AlignLeft;AlignCenter}
        public void SetRowsAlignement(int rowStart, int rowEnd, string VerticalAlignement, string HorizontalAlignement)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[rowStart + 1, 1], objSheet.Cells[rowEnd + 1, 1]);
        	
       		if ( VerticalAlignement == "AlignBottom" ) rng.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
       		if ( VerticalAlignement == "AlignCenter" ) rng.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
       		if ( VerticalAlignement == "AlignTop" ) rng.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
       		if ( HorizontalAlignement == "AlignRight" ) rng.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
       		if ( HorizontalAlignement == "AlignLeft" ) rng.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
       		if ( HorizontalAlignement == "AlignCenter" ) rng.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }
        // VerticalAlignement = {AlignBottom;AlignCenter;AlignTop}
		// HorizontalAlignement = {AlignRight;AlignLeft;AlignCenter}
        public void SetRangeAlignement(int left, int right, int top, int bottom, string VerticalAlignement, string HorizontalAlignement)
        {
        	if ( objSheet == null ) return;
        	
        	Excel.Range rng = (Excel.Range)objSheet.get_Range(objSheet.Cells[top + 1, left + 1], objSheet.Cells[bottom + 1, right + 1]);
 
       		if ( VerticalAlignement == "AlignBottom" ) rng.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
       		if ( VerticalAlignement == "AlignCenter" ) rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
       		if ( VerticalAlignement == "AlignTop" ) rng.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
       		if ( HorizontalAlignement == "AlignRight" ) rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
       		if ( HorizontalAlignement == "AlignLeft" ) rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
       		if ( HorizontalAlignement == "AlignCenter" ) rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }
    }
}
