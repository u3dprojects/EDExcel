using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

/// <summary>
/// 类名 : EPPlus_Excel_Sheet 实体类
/// 作者 : Canyon
/// 日期 : 2017-08-14 20:10
/// 功能 : XSSF
/// </summary>
public class EX_Sheet {

	// 排序对象
	static EX_Sort_Cell _sort = new EX_Sort_Cell();
	
	public EX_WorkBook m_exWB;

	public ExcelWorksheet m_sheet;

	public string sheetName;
	public int sheetIndex;
	public int NumberOfRows;
	public int NumberOfColumns;

	public List<EX_SheetCell> m_tableList = new List<EX_SheetCell> ();

	protected ExcelPackage pre_exPackage;

	public EX_Sheet(string path,int sheetIndex){
		InitFirst (path,sheetIndex);
		Init ();
	}

	public EX_Sheet(string path,string sheetName){
		InitFirst (path,sheetName);
		Init ();
	}

	~EX_Sheet()
	{
		Clear ();
	}

	void InitFirst(string path,int sheetIndex){
		this.m_exWB = new EX_WorkBook (path);
		this.m_sheet = m_exWB.GetSheet (sheetIndex);
	}

	void InitFirst(string path,string sheetName){
		this.m_exWB = new EX_WorkBook (path);
		this.m_sheet = m_exWB.GetSheet (sheetName);
	}

	void Init(){
		this.sheetName = this.m_sheet.Name;
		this.sheetIndex = this.m_sheet.Index;
		this.NumberOfRows = this.m_sheet.Dimension.Rows;
		this.NumberOfColumns = this.m_sheet.Dimension.Columns;

		EX_SheetCell tmpCell;
		for (int row = 1; row <= NumberOfRows; row++)
		{
			for (int column = 1; column <= NumberOfColumns; column++)
			{
				tmpCell = EX_SheetCell.NewCell (this, row, column);
				m_tableList.Add(tmpCell);
			}
		}

		m_tableList.Sort(_sort);
	}

	public ExcelRange GetCell(int row,int column,bool isNew = false){
		ExcelRange ret = null;
		try {
			ret =  m_sheet.Cells[row,column];
		} catch (System.Exception ex) {
			if (!isNew) 
				UnityEngine.Debug.LogError ( "excel name = " + this.m_exWB.m_filename + ",maxRow = " + this.NumberOfRows +",row = " + row + ",column = " + column + "\n" + ex);
		}
		if (isNew && ret == null)
		{
			m_sheet.Cells [row, column].Value = "";
			ret =  m_sheet.Cells[row,column];
		}
		return ret;
	}

	public object GetObject(int row,int column)
	{
		if (row < 1 || row > this.NumberOfRows || column < 1 || column > this.NumberOfColumns)
			return null;

		ExcelRange cell = GetCell (row, column);
		if (cell == null) {
			return null;
		}
		return cell.Value;
	}

	public string GetString(int row, int column)
	{
		object obj = GetObject(row, column);
		if (obj != null)
			return obj.ToString();
		return "";
	}

	public int GetInt(int row, int column)
	{
		object obj = GetObject(row, column);
		try
		{
			if (obj != null)
				return int.Parse(obj.ToString());
		}
		catch (System.Exception)
		{
		}
		return 0;
	}

	public float GetFloat(int row, int column)
	{
		object obj = GetObject(row, column);
		try
		{
			if (obj != null)
				return float.Parse(obj.ToString());
		}
		catch (System.Exception)
		{
		}
		return 0.0f;
	}

	public double GetDouble(int row, int column)
	{
		object obj = GetObject(row, column);
		try
		{
			if (obj != null)
				return double.Parse(obj.ToString());
		}
		catch (System.Exception)
		{
		}
		return 0.0d;
	}

	#region === 保存 ===

	public void SaveValue(int row, int column, string val)
	{
		ExcelRange cell = GetCell(row, column, true);
		cell.Value = val;
	}

	public void SaveValue(int row, int column, int val)
	{
		ExcelRange cell = GetCell(row, column, true);
		cell.Value = val;
	}

	public void SaveValue(int row, int column, float val)
	{
		ExcelRange cell = GetCell(row, column, true);
		cell.Value = val;
	}

	public void SaveValue(int row, int column, double val)
	{
		ExcelRange cell = GetCell(row, column, true);
		cell.Value = val;
	}

	public void SaveValue(int row,int column,bool val)
	{
		ExcelRange cell = GetCell(row, column, true);
		cell.Value = val;
	}

	public void SaveValue(int row, int column, System.DateTime val)
	{
		ExcelRange cell = GetCell(row, column, true);
		cell.Value = val;
	}

	public void SaveValueToExcel(int row, int column,object val)
	{
		if (val == null)
			return;
		ExcelRange cell = GetCell(row, column, true);
		cell.Value = val;
	}

	public void SaveValueToExcel(EX_SheetCell one)
	{
		if (one == null)
			return;
		SaveValueToExcel(one.row, one.column, one.val);
	}

	public void SaveValueToCache(int row, int column, object val)
	{
		if (val == null)
			return;

		EX_SheetCell exCell = GetExCell(row, column);
		if(exCell == null)
		{
			exCell = EX_SheetCell.NewCell(this, row, column);
			m_tableList.Add(exCell);
		}
		exCell.val = val;
	}

	#endregion

	public EX_SheetCell GetExCell(int row,int column)
	{
		if (m_tableList.Count <= 0)
			return null;

		int lens = m_tableList.Count;
		EX_SheetCell tmpCell = null;
		for (int i = 0; i < lens; i++)
		{
			tmpCell = m_tableList[i];
			if (row == tmpCell.row && column == tmpCell.column)
			{
				return tmpCell;
			}
		}
		return null;
	}

	public ExcelPackage ToExcel(bool isNew = false){
		ExcelPackage ep = null;
		this.pre_exPackage = this.m_exWB.m_ep;
		if (isNew) {
			ep = new ExcelPackage ();
			this.m_exWB.m_ep = ep;
			this.m_exWB.m_wb = ep.Workbook;
			this.m_sheet = ep.Workbook.Worksheets.Add (this.sheetName);
		} else {
			ep = this.m_exWB.m_ep;
		}

		int lens = m_tableList.Count;
		for (int i = 0; i < lens; i++)
		{
			SaveValueToExcel(m_tableList[i]);
		}
		return ep;
	}

	/// <summary>
	/// 保存Excel
	/// </summary>
	public void SaveToExcel(string filePath,bool isNew = false)
	{
		ExcelPackage ep = ToExcel (isNew);
		FileInfo output = new FileInfo(filePath);
		ep.SaveAs (output);
	}

	public void Clear(){
		m_exWB = null;
		m_sheet = null;
		pre_exPackage = null;
	}
}
