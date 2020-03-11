using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Diagnostics;
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

        /// <summary>
        /// 保存 workbook
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        bool Save(IWorkbook workbook);
    }

    /// <summary>
    /// 实现插件
    /// </summary>
    public class ExcelPlugin : IExcelPlugin
    {
        private Easylog.Log log = null;

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

                return workbook.CreateSheet(sheetName);
            }
            catch (Exception ex)
            {
                log.Error("CreateSheet", ex);
                return null;
            }
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
                    path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "auto-save" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
                }

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

        /// <summary>
        /// 保存 workbook,默认保存在桌面
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public bool Save(IWorkbook workbook)
        {
            try
            {
                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "auto-save" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
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
    }
}
