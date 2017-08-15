using UnityEngine;
using System.Collections;

/// <summary>
/// 类名 : EPPlus_Excel_Sheet_Cell 实体类
/// 作者 : Canyon
/// 日期 : 2017-08-14 20:15
/// 功能 : XSSF
/// </summary>
public class EX_SheetCell{
	
	public int row { get; set; }
	public int column { get; set; }
	public object val { get; set; }

	public EX_Sheet m_sheet { get; set; }

	public string val2Str{
		get
		{
			if (val != null)
				return val.ToString();
			return "";
		}
	}

	public EX_SheetCell(EX_Sheet sheet,int row,int column){
		this.m_sheet = sheet;
		this.row = row;
		this.column = column;
		this.val = this.m_sheet.GetObject (row, column);
	}

	static public EX_SheetCell NewCell(EX_Sheet nsheet, int row, int column)
	{
		return new EX_SheetCell(nsheet, row, column);
	}
}

public class EX_Sort_Cell : System.Collections.Generic.IComparer<EX_SheetCell>
{
	public int Compare(EX_SheetCell x, EX_SheetCell y)
	{
		if (x.row == y.row)
			return x.column < y.column ? -1 : 1;
		return x.row < y.row ? -1 : 1;
	}
}