-- Available functions (from XLSFormater Class)
-- SetProgressCalcul1ValueForLua(int val)
-- SetProgressCalcul2ValueForLua(int val)
-- IsAppendOnCurrentExcelFile()
-- Fonction pour manipulation excel
-- first row and first col = 0
-- void OpenExcelApp(bool NewAppIfNot)
-- void OpenXlsFile(string fPath)
-- void SetActiveSheet(string name)
-- void AddSheet()
-- void RenameSheet(string name)
-- void AddCell(int col, int row, string value)
-- string GetCell(int col, int row)
-- void InsertColAfter(int col)
-- void InsertRowAfter(int row)
-- void SetSizeOfCol(int col, int size)
-- void SortRangeByOneColInOrder(int left, int right, int top, int bottom, int col1, bool alpha1)
-- void SortRangeByTwoColInOrder(int left, int right, int top, int bottom, int col1, bool alpha1, int col2, bool alpha2)
-- void SortRangeByThreeColInOrder(int left, int right, int top, int bottom, int col1, bool alpha1, int col2, bool alpha2, int col3, bool alpha3)
-- void Replace( int row, int col, string pattern1, string pattern2)
-- void SetRowColor(int row, int colorIndex)
-- void SetColColor(int col, int colorIndex)
-- void SetRangeColor(int left, int right, int top, int bottom, int colorIndex)
-- void GetColorIndexInNewExcelSheet(
-- // BorderWeight = {Hairline;Medium;Thick;Thin;None}
-- void SetRangeBordure(int left, int right, int top, int bottom, bool EdgeTop, bool EdgeBottom, bool EdgeLeft, bool EdgeRight, bool insideHoriz, bool insideVert, string BorderWeight)
-- void SetRowBordure(int row, bool EdgeTop, bool EdgeBottom, bool EdgeLeft, bool EdgeRight, bool insideHoriz, bool insideVert, string BorderWeight)
-- void SetColBordure(int col, bool EdgeTop, bool EdgeBottom, bool EdgeLeft, bool EdgeRight, bool insideHoriz, bool insideVert, string BorderWeight)
-- void AutoFitCols(int colStart, int colEnd)
-- void AutoFitRows(int rowStart, int rowEnd)
-- // VerticalAlignement = {AlignBottom;AlignCenter;AlignTop}
-- // HorizontalAlignement = {AlignRight;AlignLeft;AlignCenter}
-- void SetRangeAlignement(int left, int right, int top, int bottom, string VerticalAlignement, string HorizontalAlignement)
-- void SetRowsAlignement(int rowStart, int rowEnd, string VerticalAlignement, string HorizontalAlignement)
-- void SetColsAlignement((int colStart, int colEnd, string VerticalAlignement, string HorizontalAlignement)

function Init()
  	SetAuthor("Aiekick"); -- define the script authot ( will be added in file propriety )
	SetSeparator(";"); -- define the separtor used in csv file
	SetBufferForCurrentLine("curLine"); -- curLine will be filled with the current line
	SetBufferForLastLine("lastLine"); -- lastLine will be filled with the last line
	SetFunctionForEachLine("eachRow"); -- eachRow will be called after each row filling of current ilne
	SetFunctionForEndFile("endFile"); -- endFile will be callaed at end of the current csv file
end

curLine = {}; -- Current Line
lastLine = {}; -- Last Line
function eachRow()
	local row = GetCurrentRowIndex();
end

function endFile()

end


