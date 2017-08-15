
/// <summary>
/// 类名 : EPPlus_Excel_Row 实体类
/// 作者 : Canyon
/// 日期 : 2017-08-15 09:10
/// 功能 : sheetIndex,rowIndex,columnIndex 都是从1开始
/// </summary>
public class EX_Row{

	/// <summary>
	/// 所属 Sheet
	/// </summary>
	public EX_Sheet sheet;

	/// <summary>
	/// 所在行
	/// </summary>
	public int rowIndex = 1;

	/// <summary>
	/// 唯一标识
	/// </summary>
	public int uniqueID = 0;

	/// <summary>
	/// 列元素
	/// </summary>
	protected virtual object[] Columns
	{
		get{
			return null;
		}
	}

	public bool IsCanCache
	{
		get{
			if (sheet == null || rowIndex < 0)
				return false;
			return true;
		}
	}

	public void ToCache()
	{
		if (!IsCanCache)
			return;
		
		object[] _columns = Columns;
		if (_columns == null || _columns.Length <= 0)
			return;

		int lens = _columns.Length;
		object val = null;
		for (int i = 0; i < lens; i++) {
			val = _columns [i];
			this.sheet.SaveValueToCache (this.rowIndex, i + 1, val);
		}
	}

	public virtual void Init(EX_Sheet sheet,int rowIndex)
	{
		this.sheet = sheet;
		this.rowIndex = rowIndex;
	}
}
