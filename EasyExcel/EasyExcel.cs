using Newtonsoft.Json.Linq;
using NPOI.HSSF.Record.CF;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.IO;


/// <summary>
/// 建议把下面的namespace名字改为您的插件名字
/// </summary>
namespace EasyExcel
{
    /// <summary>
    /// 定义一个插件函数时，必须先在这个interface里面声明
    /// </summary>
    public interface IExcelPlugin
    {
        /// <summary>
        /// 创建Excel
        /// </summary>
        /// <returns>workbook</returns>
        IWorkbook CreateExcel();

        /// <summary>
        /// 创建Sheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns>sheet</returns>
        ISheet CreateSheet(IWorkbook workbook, string sheetName);

        /// <summary>
        /// 获取sheet
        /// </summary>
        /// <param name="workbook">workbook</param>
        /// <param name="sheet">sheet名称或者索引</param>
        /// <returns>sheet</returns>
        ISheet GetSheet(IWorkbook workbook, dynamic sheet);

        IWorkbook DeleteSheet(IWorkbook workbook, dynamic sheet);

        /// <summary>
        /// 写入行数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRowNum"></param>
        /// <param name="fromColumnNum"></param>
        /// <param name="array"></param>
        /// <returns></returns>
        bool WriteRow(ISheet sheet, int fromRowNum, int fromColumnNum, JArray array);

        /// <summary>
        /// 写入区域数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRowNum"></param>
        /// <param name="fromColumnNum"></param>
        /// <param name="array"></param>
        /// <returns></returns>
        bool WriteRange(ISheet sheet, int fromRowNum, int fromColumnNum, JArray array);

        /// <summary>
        /// 保存 workbook
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="path">路径</param>
        /// <returns></returns>
        bool Save(IWorkbook workbook, string path);

        bool Close(IWorkbook workbook, bool isSave);

        IWorkbook OpenExcel(string filePath);

        object ReadCell(ISheet sheet, int rowNum, int cellNum);

        bool WriteCell(ISheet sheet, int rowNum, int cellNum, dynamic value);

        JArray ReadRow(ISheet sheet, int rowNum);

        JArray ReadRange(ISheet sheet, string range);

        int GetRowsCount(ISheet sheet);

        int GetColumsCount(ISheet sheet);
    }

    /// <summary>
    /// 实现插件
    /// </summary>
    public class ExcelPlugin : IExcelPlugin
    {
        private Easylog.Log log = null;

        private static string ExcelFilePath = string.Empty;

        public ExcelPlugin()
        {
            log = new Easylog.Log();
        }

        /// <summary>
        /// 创建 workbook
        /// </summary>
        /// <returns></returns>
        public IWorkbook CreateExcel()
        {
            try
            {
                return new XSSFWorkbook();
            }
            catch (Exception ex)
            {
                log.Error("CreateExcel", ex);
                return null;
            }
        }

        /// <summary>
        /// 创建 sheet
        /// </summary>
        /// <param name="workbook">excel workbook</param>
        /// <param name="sheetName">sheet name</param>
        /// <returns></returns>
        public ISheet CreateSheet(IWorkbook workbook, string sheetName)
        {
            try
            {
                if (workbook == null)
                {
                    log.Info("[CreateSheet]=>workbook不能为空");
                    return null;
                }

                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    sheetName = "sheet1";
                }

                //防止sheet名字重复
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var name = workbook.GetSheetName(i);
                    if (sheetName == name)
                    {
                        sheetName = string.Format("{0}-r{1}", sheetName, new Random().Next(10, 99)) + DateTime.Now.Ticks; //添加随机数
                        break;
                    }
                }

                return workbook.CreateSheet(sheetName);
            }
            catch (Exception ex)
            {
                log.Error("CreateSheet", ex);
                return null;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public IWorkbook DeleteSheet(IWorkbook workbook, dynamic sheet)
        {
            try
            {
                if (workbook == null)
                {
                    log.Info("[DeleteSheet]=>workbook不能为空");
                }

                if (sheet is int)
                {
                    workbook.RemoveSheetAt(sheet);
                }
                else
                {
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        var name = workbook.GetSheetName(i);
                        if (sheet.ToString() == name)
                        {
                            workbook.RemoveSheetAt(i);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("DeleteSheet", ex);
            }

            return workbook;
        }

        /// <summary>
        /// 写入行数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRowNum">起始行,索引从0开始</param>
        /// <param name="fromColumnNum">起始列,索引从0开始</param>
        /// <param name="array">['Small','Medium','Large']</param>
        /// <returns></returns>
        public bool WriteRow(ISheet sheet, int fromRowNum, int fromColumnNum, JArray array)
        {
            try
            {
                if (sheet == null)
                {
                    log.Info("[WriteRow]=>sheet不能为空");
                    return false;
                }

                if (array == null)
                {
                    log.Info("[WriteRow]=>尝试写入空数据");
                    return false;
                }

                var row = sheet.CreateRow(fromRowNum);
                for (int i = 0; i < array.Count; i++)
                {
                    row.CreateCell(fromColumnNum + i).SetCellValue(array[i].ToString());
                }

                return true;
            }
            catch (Exception ex)
            {
                log.Error("WriteRow", ex);
                return false;
            }
        }

        /// <summary>
        /// 写入区域数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRowNum">起始行,索引从0开始</param>
        /// <param name="fromColumnNum">起始列,索引从0开始</param>
        /// <param name="array">[['a1','b1','c1'],['a2','b2','c2']]>
        /// <returns></returns>
        public bool WriteRange(ISheet sheet, int fromRowNum, int fromColumnNum, JArray array)
        {
            try
            {
                if (sheet == null)
                {
                    log.Info("[WriteRange]=>sheet不能为空");
                    return false;
                }

                if (array == null)
                {
                    log.Info("[WriteRange]=>尝试写入空数据");
                    return false;
                }


                for (int i = 0; i < array.Count; i++)
                {
                    var row = array[i] as JArray;
                    this.WriteRow(sheet, fromRowNum + i, fromColumnNum, row);
                }

                return true;
            }
            catch (Exception ex)
            {
                log.Error("WriteRow", ex);
                return false;
            }
        }

        /// <summary>
        /// 保存 workbook
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool Save(IWorkbook workbook, string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path))
                {
                    path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "auto-save" + DateTime.Now.ToString("yyyyMMdd") + "-r" + new Random().Next(10, 99).ToString() + ".xlsx");
                }

                //创建目录
                string dir = Path.GetDirectoryName(path);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                //创建文件
                using (var fs = File.Create(path))
                {
                    workbook.Write(fs);
                }

                return true;
            }
            catch (Exception ex)
            {
                log.Error("Save", ex);
                return false;
            }
        }

        public bool Close(IWorkbook workbook, bool isSave)
        {
            try
            {
                if (isSave)
                {
                    using (var fs = File.Create(ExcelFilePath))
                    {
                        workbook.Write(fs);
                    }
                }

                workbook.Close();
                return true;
            }
            catch (Exception ex)
            {
                log.Error("Save", ex);
                return false;
            }
        }

        #region 读取数据


        /// <summary>
        /// 打开Excel
        /// </summary>
        /// <param name="filePath">excel文件地址</param>
        /// <returns></returns>
        public IWorkbook OpenExcel(string filePath)
        {
            try
            {
                if (Path.GetExtension(filePath) == ".xls")
                {
                    ExcelFilePath = filePath;
                    return WorkbookFactory.Create(filePath);
                }

                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    ExcelFilePath = filePath;
                    return new XSSFWorkbook(file);
                }
            }
            catch (Exception ex)
            {
                log.Error("OpenExcel", ex);
                return null;
            }
        }

        /// <summary>
        /// 获取sheet
        /// </summary>
        /// <param name="workbook">workbook</param>
        /// <param name="sheet">sheet名称或者索引</param>
        /// <returns>sheet</returns>
        public ISheet GetSheet(IWorkbook workbook, dynamic sheet)
        {
            try
            {
                if (sheet is int)
                {
                    return workbook.GetSheetAt(sheet);
                }
                else
                {
                    return workbook.GetSheet(sheet);
                }

            }
            catch (Exception ex)
            {
                log.Error("GetSheet", ex);
                return null;
            }
        }

        //读取单元格
        public object ReadCell(ISheet sheet, int rowNum, int cellNum)
        {
            try
            {
                return sheet.GetRow(rowNum).GetCell(cellNum);
            }
            catch (Exception ex)
            {
                log.Error("ReadCell", ex);
                return string.Empty;
            }
        }

        //写入单元格
        public bool WriteCell(ISheet sheet, int rowNum, int cellNum, dynamic value)
        {
            try
            {
                if (sheet.GetRow(rowNum) == null || sheet.GetRow(rowNum).GetCell(cellNum) == null)
                {
                    sheet.CreateRow(rowNum).CreateCell(cellNum).SetCellValue(value);
                }
                else
                {
                    sheet.GetRow(rowNum).GetCell(cellNum).SetCellValue(value);
                }

                return true;
            }
            catch (Exception ex)
            {
                log.Error("SetCell", ex);
                return false;
            }
        }

        //创建单元格
        public bool CreateCell(ISheet sheet, int rowNum, int cellNum, dynamic value)
        {
            try
            {
                sheet.CreateRow(rowNum).CreateCell(cellNum).SetCellValue(value);
                return true;
            }
            catch (Exception ex)
            {
                log.Error("SetCell", ex);
                return false;
            }
        }

        public ICell[,] GetRange(ISheet sheet, string range)
        {
            string[] cellStartStop = range.Split(':');

            CellReference cellRefStart = new CellReference(cellStartStop[0]);
            CellReference cellRefStop = new CellReference(cellStartStop[1]);

            ICell[,] cells = new ICell[cellRefStop.Row - cellRefStart.Row + 1, cellRefStop.Col - cellRefStart.Col + 1];

            for (int i = cellRefStart.Row; i < cellRefStop.Row + 1; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                {
                    row = sheet.CreateRow(i);
                }

                for (int j = cellRefStart.Col; j < cellRefStop.Col + 1; j++)
                {
                    cells[i - cellRefStart.Row, j - cellRefStart.Col] = row.GetCell(j);
                }
            }

            return cells;
        }

        //读取区域
        public JArray ReadRange(ISheet sheet, string range)
        {
            JArray array = new JArray();
            try
            {
                var cells = GetRange(sheet, range);
                foreach (var item in cells)
                {
                    if (item != null)
                    {
                        array.Add(item.ToString());
                    }
                    else
                    {
                        array.Add("");
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("ReadRange", ex);
            }

            return array;
        }

        //读取行
        public JArray ReadRow(ISheet sheet, int rowNum)
        {
            JArray array = new JArray();
            try
            {
                foreach (var item in sheet.GetRow(rowNum).Cells)
                {
                    array.Add(item.ToString());
                }
            }
            catch (Exception ex)
            {
                log.Error("ReadRow", ex);
            }

            return array;
        }

        //读取列

        //获取行数
        public int GetRowsCount(ISheet sheet)
        {
            try
            {
                return sheet.LastRowNum;
            }
            catch (Exception ex)
            {
                log.Error("GetRowsCount", ex);
            }

            return 0;
        }

        //获取列数
        public int GetColumsCount(ISheet sheet)
        {
            try
            {
                return sheet.GetRow(0).LastCellNum;
            }
            catch (Exception ex)
            {
                log.Error("GetColumsCount", ex);
            }

            return 0;
        }

        //合并单元格
        public bool MergeRange(ISheet sheet)
        {
            return true;
        }

        //拆分单元格

        //写入行
        //删除行
        //插入行
        //插入列
        //插入图片
        //删除图片
        //写入区域
        //选中区域
        //清除区域
        //删除区域

        #endregion
    }
}
