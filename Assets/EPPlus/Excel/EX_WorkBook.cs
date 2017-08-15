using OfficeOpenXml;
using System.IO;

/// <summary>
/// 类名 : EPPlus_Excel_WorkBook 实体类
/// 作者 : Canyon
/// 日期 : 2017-08-14 19:40
/// 功能 : XSSF
/// </summary>
public class EX_WorkBook {

	protected string m_path;

	public string m_foloder;
	public string m_filename;
	public string m_suffix;

	public ExcelPackage m_ep;
	public ExcelWorkbook m_wb; 

	public EX_WorkBook(string path){
		Init (path);
	}

	~EX_WorkBook()
	{
		Clear ();
	}

	void Init(string path){
		this.m_path = path;
		this.m_foloder = Path.GetDirectoryName (m_path);
		this.m_filename = Path.GetFileNameWithoutExtension (m_path);
		this.m_suffix = Path.GetExtension (m_path);
		this.m_suffix = this.m_suffix.Replace(".", "");

		FileInfo file = new FileInfo(m_path);
		this.m_ep = new ExcelPackage(file);
		this.m_wb = this.m_ep.Workbook;
	}

	/// <summary>
	/// Gets the sheet.
	/// </summary>
	/// <returns>The sheet.</returns>
	/// <param name="sheetIndex">Sheet index 从 1 开始.</param>
	public ExcelWorksheet GetSheet(int sheetIndex){
		if(sheetIndex > 0 && sheetIndex <= this.m_wb.Worksheets.Count)
			return this.m_wb.Worksheets [sheetIndex];
		
		return null;
	}

	public ExcelWorksheet GetSheet(string sheetName){
		int lens = this.m_wb.Worksheets.Count;
		ExcelWorksheet ret = null;
		for (int i = 0; i < lens; i++) {
			ret = this.m_wb.Worksheets [i];
			if (ret.Name.Equals (sheetName)) {
				return ret;
			}
		}
		return null;
	}

	/// <summary>
	/// 保存Excel
	/// </summary>
	public void SaveToExcel(string filePath)
	{
		FileInfo output = new FileInfo(filePath);
		this.m_ep.SaveAs (output);
	}

	public void Clear(){
		m_wb = null;
		if (m_ep != null)
			m_ep.Dispose ();
		m_ep = null;
	}
}
