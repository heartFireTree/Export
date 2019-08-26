using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HPSF;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Aspose.Cells;

namespace ExcelHelp
{
    public class ToExcel
    {
       public static string CompanyName = "自定义";
         
          /// <summary>
          ///  导出excel
          /// </summary>
          /// <param name="dtSource">table数据</param>
          /// <param name="strHeaderText">标题</param>
          /// <param name="strFileName">文件名</param>
          /// <param name="IsDoubleHead">是否双表头</param>
          /// <param name="heads">表头信息</param>
          public static void ExportByWeb(List<DataTable> dtSource, List<string> strHeaderText, List<string> heads, string strFileName)
          {
              HttpContext curContext = HttpContext.Current;
              string fileName = strFileName;
              //处理乱码兼容性
              if (curContext.Request.UserAgent.ToLower().IndexOf("firefox") > -1)
              {

              }
              else
              {
                  fileName = HttpUtility.UrlEncode(strFileName, System.Text.Encoding.UTF8);
              }


              // 设置编码和附件格式  
              curContext.Response.ContentType = "application/vnd.ms-excel";
              curContext.Response.ContentEncoding = Encoding.UTF8;
              curContext.Response.Charset = "";
              curContext.Response.AppendHeader("Content-Disposition",
                  "attachment;filename=" + fileName);
              if (dtSource == null && dtSource.Count == 0)
              {
                  return;
              }
              curContext.Response.BinaryWrite(ExportMultipleSheet(dtSource, strHeaderText, heads, fileName).GetBuffer());

              curContext.Response.End();
          }
        
         public static void ExportByWeb(List<DataTable> dtSource, List<string> strHeaderText, List<string> heads, List<string> pageNames, string strFileName)
         {
            HttpContext curContext = HttpContext.Current;
            string fileName = strFileName;
            //处理乱码兼容性
            if (curContext.Request.UserAgent.ToLower().IndexOf("firefox") > -1)
            {

            }
            else
            {
                fileName = HttpUtility.UrlEncode(strFileName, System.Text.Encoding.UTF8);
            }
            // 设置编码和附件格式  
            curContext.Response.ContentType = "application/vnd.ms-excel";
            curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.Charset = "";
            curContext.Response.AppendHeader("Content-Disposition",
                "attachment;filename=" + fileName);
            if (dtSource == null && dtSource.Count == 0)
            {
                return;
            }
            curContext.Response.BinaryWrite(ExportByAspose(dtSource, strHeaderText, heads, pageNames).GetBuffer());
            curContext.Response.End();
         }
        
        /// <summary>
        /// Aspose批量Sheet导出
        /// </summary>
        /// <param name="dtSource">多个sheet</param>
        /// <param name="HeadTitle">标题</param>
        /// <param name="HeadTitletow">二标题</param>
        /// <param name="pageName">每个sheet名称</param>
        /// <returns></returns>
        public static MemoryStream ExportByAspose(List<DataTable> dtSource, List<string> HeadTitle, List<string> HeadTitletow, List<string> pageName)
        {
            Workbook workbook = new Aspose.Cells.Workbook();
            //清除页先 要不然 新建就有一个sheet
            workbook.Worksheets.Clear();
            #region 样式
            //标题样式
            Style styleTitle = workbook.Styles[workbook.Styles.Add()];
            styleTitle.HorizontalAlignment = TextAlignmentType.Center;
            styleTitle.Font.Name = "宋体";
            styleTitle.Font.Size = 18;
            styleTitle.Font.IsBold = true;

            //样式2
            Style style2 = workbook.Styles[workbook.Styles.Add()];
            style2.HorizontalAlignment = TextAlignmentType.Right;
            style2.Font.IsBold = false;
            style2.IsTextWrapped = true;

            //列样式
            Style styleColumn = workbook.Styles[workbook.Styles.Add()];
            styleColumn.HorizontalAlignment = TextAlignmentType.Left;
            styleColumn.VerticalAlignment = TextAlignmentType.Center;
            styleColumn.Font.IsBold = true;
            styleColumn.Font.IsNormalizeHeights = true;
            #endregion

            for (int i = 0; i < dtSource.Count; i++)
            {
                string sheetName = pageName[i].Trim();
                sheetName = sheetName.Length > 31 ? sheetName.Substring(0,31) : sheetName;
                //创建sheet(excel限制最大长度为31)

                workbook.Worksheets.Add(sheetName);
                //获取当前sheet
                Worksheet sheet = workbook.Worksheets[i];
                Cells cells = sheet.Cells;

                int rowIndex = 0;
                if (dtSource[i] != null)
                {
                    //标题行
                    int colNum = dtSource[i].Columns.Count;//表格列数
                    int rowNum = dtSource[i].Rows.Count;//表格行数
                    cells.Merge(0, 0, 1, colNum);//合并单元格//生成行1 标题行 
                    cells[0, 0].PutValue(HeadTitle[i].ToString().Trim());//填写标题内容
                    cells[0, 0].SetStyle(styleTitle);//标题样式
                    cells.SetRowHeight(0, 30);//设置行高

                    //标题第二行
                    cells.Merge(1, 0, 1, colNum);//合并单元格
                    cells[1,0].PutValue(HeadTitletow[i] + "| 导出时间:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                    cells[1, 0].SetStyle(style2);
                    cells.SetRowHeight(1, 18);

                    //列名行
                    for (int ct = 0; ct < colNum; ct++)
                    {
                        cells[2, ct].PutValue(dtSource[i].Columns[ct].ColumnName);
                        cells[2, ct].SetStyle(styleColumn);
                    }
                    cells.SetRowHeight(2, 18);

                    rowIndex = 3;
                    //生成数据行
                    for (int r = 0; r < rowNum; r++)
                    {
                        //数据列遍历
                        for (int c = 0; c < colNum; c++)
                        {
                            string drValue = dtSource[i].Rows[r][c].ToString();

                            double doubV = 0;

                            //设置单元格数据根据值类型转换
                            if (double.TryParse(drValue,out doubV))
                            {
                                cells[rowIndex, c].PutValue(doubV);
                            }
                            else
                            {
                                cells[rowIndex, c].PutValue(drValue);
                            }
                        }
                        cells.SetRowHeight(rowIndex, 18);
                        rowIndex++;
                    }
                }
                //sheet.AutoFitColumns();//自动填充列宽(不完美)
                setColumnWithAuto(sheet);
            }

            MemoryStream ms = workbook.SaveToStream();
            return ms;
        }

        /// <summary>          
        /// Aspose设置表页的列宽度自适应          
        /// </summary>         
        /// /// <param name="sheet">worksheet对象</param>       
        public static void setColumnWithAuto(Worksheet sheet)
        {
            Cells cells = sheet.Cells;
            //获取表页的最大列数    
            int columnCount = cells.MaxColumn + 1;
            //获取表页的最大行数   
            int rowCount = cells.MaxRow;
                              
            for (int col = 0; col < columnCount; col++)
            {
                sheet.AutoFitColumn(col, 0, rowCount);
            }
            for (int col = 0; col < columnCount; col++)
            {
                int pixel = cells.GetColumnWidthPixel(col) + 30;
                if (pixel > 255)
                {
                    cells.SetColumnWidthPixel(col, 255);
                }
                else
                {
                    cells.SetColumnWidthPixel(col, pixel);
                }
            }
        }
          
          public static MemoryStream Export(DataTable dtSource, string strHeaderText, string hidText, string FileName = null)
          {
              IWorkbook workbook = null;
              if (!string.IsNullOrWhiteSpace(FileName))
              {
                  if (FileName.IndexOf(".xlsx") > 0)
                      workbook = new XSSFWorkbook();
                  // 2003版本  
                  else if (FileName.IndexOf(".xls") > 0)
                      workbook = new HSSFWorkbook();
              }
              else
              {
                  workbook = new HSSFWorkbook();
              }

              new HSSFWorkbook();
              ISheet sheet = workbook.CreateSheet();

              #region 右击文件 属性信息
              {
                  DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();

                  dsi.Company = CompanyName;

                  //workbook.DocumentSummaryInformation = dsi;

                  SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                  si.Author = CompanyName; //填加xls文件作者信息  
                  si.ApplicationName = CompanyName; //填加xls文件创建程序信息  
                  si.LastAuthor = CompanyName; //填加xls文件最后保存者信息  
                  si.Comments = CompanyName; //填加xls文件作者信息  
                  si.Title = CompanyName; //填加xls文件标题信息  
                  si.Subject = CompanyName;//填加文件主题信息  
                  si.CreateDateTime = DateTime.Now;
                  //workbook.SummaryInformation = si;
              }
              #endregion

              ICellStyle dateStyle = workbook.CreateCellStyle();
              IDataFormat format = workbook.CreateDataFormat();
              dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
              //取得列宽  
              int[] arrColWidth = new int[dtSource.Columns.Count];
              foreach (DataColumn item in dtSource.Columns)
              {
                  arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
              }
              for (int i = 0; i < dtSource.Rows.Count; i++)
              {
                  for (int j = 0; j < dtSource.Columns.Count; j++)
                  {
                      var intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                      if (intTemp > arrColWidth[j])
                      {
                          arrColWidth[j] = intTemp;
                      }
                  }
              }
              int rowIndex = 0;
              ICellStyle tdStyle = null;
              if (!string.IsNullOrWhiteSpace(FileName))
              {
                  //2007版本
                  if (FileName.IndexOf(".xlsx") > 0)
                      tdStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                  // 2003版本  
                  else if (FileName.IndexOf(".xls") > 0)
                      tdStyle = (HSSFCellStyle)workbook.CreateCellStyle();
              }
              else
              {
                  tdStyle = (HSSFCellStyle)workbook.CreateCellStyle();
              }



              foreach (DataRow row in dtSource.Rows)
              {
                  #region 新建表，填充表头，填充列头，样式
                  if (rowIndex == 65535 || rowIndex == 0)
                  {
                      if (rowIndex != 0)
                      {
                          sheet = workbook.CreateSheet();
                      }
                      #region 表头及样式
                      {
                          IRow headerRow = sheet.CreateRow(0);
                          headerRow.HeightInPoints = 30;
                          headerRow.CreateCell(0).SetCellValue(strHeaderText);
                          ICellStyle headStyle = workbook.CreateCellStyle();
                          headStyle.Alignment = HorizontalAlignment.Center;
                          IFont font = workbook.CreateFont();
                          font.FontHeightInPoints = 20;
                          font.Boldweight = 700;
                          headStyle.SetFont(font);
                          headerRow.GetCell(0).CellStyle = headStyle;
                          sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1));
                      }
                      #endregion
                      {
                          IRow headerRow = sheet.CreateRow(1);
                          if (hidText != "")
                          {
                              hidText += " | ";
                          }
                          headerRow.CreateCell(0).SetCellValue(hidText + "导出时间:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                          headerRow.HeightInPoints = 18;
                          ICellStyle headStyle = workbook.CreateCellStyle();
                          headStyle.Alignment = HorizontalAlignment.Right;
                          headStyle.Indention = 50;
                          headerRow.GetCell(0).CellStyle = headStyle;
                          sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, dtSource.Columns.Count - 1));
                      }
                      #region 列头及样式
                      {
                          IRow headerRow = sheet.CreateRow(2);
                          headerRow.HeightInPoints = 18;
                          ICellStyle headStyle = workbook.CreateCellStyle();
                          headStyle.Alignment = HorizontalAlignment.Left;
                          headStyle.VerticalAlignment = VerticalAlignment.Center;
                          //headStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREEN.index;
                          //headStyle.FillPattern = NPOI.SS.UserModel.FillPatternType.SOLID_FOREGROUND;
                          //headStyle.BorderBottom = BorderStyle.THIN;
                          //headStyle.BorderLeft = BorderStyle.THIN;
                          //headStyle.BorderRight = BorderStyle.THIN;
                          //headStyle.BorderTop = BorderStyle.THIN;
                          IFont font = workbook.CreateFont();
                          font.FontHeightInPoints = 10;
                          font.Boldweight = 700;
                          //font.Color = NPOI.HSSF.Util.HSSFColor.WHITE.index;
                          headStyle.SetFont(font);

                          foreach (DataColumn column in dtSource.Columns)
                          {
                              headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                              headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                              //设置列宽   
                              if (column.ColumnName == "款号图片")
                              {
                                  sheet.SetColumnWidth(column.Ordinal, 20 * 256);
                              }
                              else
                              {
                                  var columnWidth = (arrColWidth[column.Ordinal] + 1) * 256;
                                  sheet.SetColumnWidth(column.Ordinal, columnWidth > (255 * 256) ? 6000 : columnWidth);
                                  //sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                              }
                          }
                      }
                      #endregion
                      rowIndex = 3;
                  }
                  #endregion

                  #region 填充内容
                  IRow dataRow = sheet.CreateRow(rowIndex);
                  dataRow.HeightInPoints = 18;
                  //if (rowIndex % 2 == 0)
                  //    tdStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index;
                  //else
                  //    tdStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.WHITE.index;
                  //tdStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.WHITE.index;
                  //tdStyle.FillPattern = NPOI.SS.UserModel.FillPatternType.SOLID_FOREGROUND;
                  ////设置单元格边框 
                  //tdStyle.BorderBottom = BorderStyle.THIN;
                  //tdStyle.BorderLeft = BorderStyle.THIN;
                  //tdStyle.BorderRight = BorderStyle.THIN;
                  //tdStyle.BorderTop = BorderStyle.THIN;
                  tdStyle.VerticalAlignment = VerticalAlignment.Center;
                  short datet = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
                  bool isCell = dtSource.Rows.Count < 4000;//4000行bug
                  InitRowData(dtSource, tdStyle, row, dataRow, isCell, workbook, rowIndex, strHeaderText);
                  #endregion
                  rowIndex++;
              }

              ExcelStream stream = new ExcelStream();
              workbook.Write(stream);
              stream.CanDispose = true;

              return stream;
          }

          /// <summary>
          /// 分表导出到多个Sheet
          /// </summary>
          /// <param name="dtSource"></param>
          /// <param name="strHeaderText"></param>
          /// <param name="hidText"></param>
          /// <param name="FileName"></param>
          /// <returns></returns>
          public static MemoryStream ExportMultipleSheet(List<DataTable> dtSource, List<string> strHeaderText, List<string> hidText, string FileName = null)
          {
              IWorkbook workbook = null;
              if (!string.IsNullOrWhiteSpace(FileName))
              {
                  if (FileName.IndexOf(".xlsx") > 0)
                      workbook = new XSSFWorkbook();
                  // 2003版本  
                  else if (FileName.IndexOf(".xls") > 0)
                      workbook = new HSSFWorkbook();
              }
              else
              {
                  workbook = new HSSFWorkbook();
              }

              #region 右击文件 属性信息
              {
                  DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();

                  dsi.Company = CompanyName;

                  //workbook.DocumentSummaryInformation = dsi;

                  SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                  si.Author = CompanyName; //填加xls文件作者信息  
                  si.ApplicationName = CompanyName; //填加xls文件创建程序信息  
                  si.LastAuthor = CompanyName; //填加xls文件最后保存者信息  
                  si.Comments = CompanyName; //填加xls文件作者信息  
                  si.Title = CompanyName; //填加xls文件标题信息  
                  si.Subject = CompanyName;//填加文件主题信息  
                  si.CreateDateTime = DateTime.Now;
                  //workbook.SummaryInformation = si;
              }
              #endregion

              //表头样式
              ICellStyle headStyleOne = workbook.CreateCellStyle();
              headStyleOne.Alignment = HorizontalAlignment.Center;
              IFont font = workbook.CreateFont();
              font.FontHeightInPoints = 18;
              font.Boldweight = 700;
              headStyleOne.SetFont(font);
              //表头1
              ICellStyle headStyleTwo = workbook.CreateCellStyle();
              headStyleTwo.Alignment = HorizontalAlignment.Right;
              headStyleTwo.Indention = 50;

              //列头样式
              ICellStyle headStyleThree = workbook.CreateCellStyle();
              headStyleThree.Alignment = HorizontalAlignment.Left;
              headStyleThree.VerticalAlignment = VerticalAlignment.Center;
              IFont fonts = workbook.CreateFont();
              fonts.FontHeightInPoints = 10;
              fonts.Boldweight = 700;
              headStyleThree.SetFont(fonts);

              //每个表头名
              var HeadText = string.Empty;
              //每个表头二
              var HeadTextTwo = string.Empty;
              //遍历多张表(每一张表创建一个Sheet)
              for (int i = 0; i < dtSource.Count; i++)
              {
                  HeadText = strHeaderText[i];
                  HeadTextTwo = hidText[i];
                  #region 处理cell样式和大小
                  ICellStyle dateStyle = workbook.CreateCellStyle();

                  IDataFormat format = workbook.CreateDataFormat();

                  dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
                  //取得列宽  
                  int[] arrColWidth = new int[dtSource[i].Columns.Count];
                  foreach (DataColumn item in dtSource[i].Columns)
                  {
                      arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
                  }
                  for (int r = 0; r < dtSource[i].Rows.Count; r++)
                  {
                      for (int j = 0; j < dtSource[i].Columns.Count; j++)
                      {
                          var intTemp = Encoding.GetEncoding(936).GetBytes(dtSource[i].Rows[r][j].ToString()).Length;
                          if (intTemp > arrColWidth[j])
                          {
                              arrColWidth[j] = intTemp;
                          }
                      }
                  }
                  ICellStyle tdStyle = null;

                  if (!string.IsNullOrWhiteSpace(FileName))
                  {
                      //2007版本
                      if (FileName.IndexOf(".xlsx") > 0)
                          tdStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                      // 2003版本  
                      else if (FileName.IndexOf(".xls") > 0)
                          tdStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                  }
                  else
                  {
                      tdStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                  }
                  #endregion

                  int rowIndex = 0;

                  ISheet sheet = workbook.CreateSheet();

                  foreach (DataRow row in dtSource[i].Rows)
                  {
                      #region 新建表，填充表头，填充列头，样式
                      if (rowIndex == 65535 || rowIndex == 0)
                      {
                          if (rowIndex != 0)
                          {
                              sheet = workbook.CreateSheet();
                          }
                          #region 表头及样式
                          {
                              IRow headerRow = sheet.CreateRow(0);
                              headerRow.HeightInPoints = 30;
                              headerRow.CreateCell(0).SetCellValue(HeadText);
                              headerRow.GetCell(0).CellStyle = headStyleOne;
                              sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dtSource[i].Columns.Count - 1));
                          }
                          #endregion
                          {
                              IRow headerRow = sheet.CreateRow(1);

                              headerRow.CreateCell(0).SetCellValue(HeadTextTwo + "| 导出时间:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                              headerRow.HeightInPoints = 18;

                              headerRow.GetCell(0).CellStyle = headStyleTwo;
                              sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, dtSource[i].Columns.Count - 1));
                          }
                          #region 列头及样式
                          {
                              IRow headerRow = sheet.CreateRow(2);
                              headerRow.HeightInPoints = 18;

                              foreach (DataColumn column in dtSource[i].Columns)
                              {
                                  headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                                  headerRow.GetCell(column.Ordinal).CellStyle = headStyleThree;
                                  //设置列宽   
                                  if (column.ColumnName == "款号图片")
                                  {
                                      sheet.SetColumnWidth(column.Ordinal, 20 * 256);
                                  }
                                  else
                                  {
                                      var columnWidth = (arrColWidth[column.Ordinal] + 1) * 256;
                                      sheet.SetColumnWidth(column.Ordinal, columnWidth > (255 * 256) ? 6000 : columnWidth);
                                      //sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                                  }
                              }
                          }
                          #endregion
                          rowIndex = 3;
                      }
                      #endregion

                      #region 填充内容
                      IRow dataRow = sheet.CreateRow(rowIndex);
                      dataRow.HeightInPoints = 18;
                      //if (rowIndex % 2 == 0)
                      //    tdStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index;
                      //else
                      //    tdStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.WHITE.index;
                      //tdStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.WHITE.index;
                      //tdStyle.FillPattern = NPOI.SS.UserModel.FillPatternType.SOLID_FOREGROUND;
                      ////设置单元格边框 
                      //tdStyle.BorderBottom = BorderStyle.THIN;
                      //tdStyle.BorderLeft = BorderStyle.THIN;
                      //tdStyle.BorderRight = BorderStyle.THIN;
                      //tdStyle.BorderTop = BorderStyle.THIN;
                      tdStyle.VerticalAlignment = VerticalAlignment.Center;
                      short datet = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
                      bool isCell = dtSource[i].Rows.Count < 4000;//4000行bug
                      InitRowData(dtSource[i], tdStyle, row, dataRow, isCell, workbook, rowIndex, HeadText);
                      #endregion
                      rowIndex++;
                  }
              }

              ExcelStream stream = new ExcelStream();
              workbook.Write(stream);
              stream.CanDispose = true;

              return stream;
          }

          private static void InitRowData(DataTable dtSource, ICellStyle tdStyle, DataRow row, IRow dataRow, bool isCell, IWorkbook workbook = null, int rowNum = 0, string reportType = "")
          {
              tdStyle.WrapText = true;
              if (ExportPictureTiltes.Contains(reportType))
              {
                  dataRow.Height = 2000;
              }
              for (int i = 0; i < dtSource.Columns.Count; i++)
              {
                  var column = dtSource.Columns[i];
                  ICell newCell = dataRow.CreateCell(column.Ordinal);
                  if (isCell) newCell.CellStyle = tdStyle;
                  string drValue = row[column].ToString();
                  
                  newCell.SetCellValue(drValue);

                  if (column.ColumnName == "款号图片" && !string.IsNullOrWhiteSpace(drValue))
                  {
                      #region 款号图片
                      System.Drawing.Image originalImage = null;
                      ////新建一个bmp图片         
                      //System.Drawing.Image bitmap = null;
                      ////新建一个画板       
                      //System.Drawing.Graphics g = null;
                      //double toheight = 100;
                      //double towidth = 100;
                      if (File.Exists(drValue))
                      {
                          originalImage = System.Drawing.Image.FromFile(drValue);
                          //double proportion1;
                          //double proportion2;
                          //int x = 0;
                          //int y = 0;
                          ////原图的宽   
                          //int ow = originalImage.Width;
                          ////原图的高  
                          //int oh = originalImage.Height;
                          //towidth = toheight * ow / oh;
                          //proportion1 = toheight / Convert.ToDouble(oh);
                          //proportion2 = towidth / Convert.ToDouble(ow);

                          ////新建一个bmp图片  
                          //bitmap = new System.Drawing.Bitmap(Convert.ToInt32(towidth), Convert.ToInt32(toheight));
                          ////新建一个画板 
                          //g = System.Drawing.Graphics.FromImage(bitmap);
                          ////设置高质量插值法     
                          //g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
                          ////设置高质量,低速度呈现平滑程度      
                          //g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                          ////清空画布并以透明背景色填充    
                          //g.Clear(System.Drawing.Color.Transparent);
                          ////在指定位置并且按指定大小绘制原图片的指定部分   
                          //g.DrawImage(originalImage, new System.Drawing.Rectangle(0, 0, Convert.ToInt32(towidth), Convert.ToInt32(toheight)), new System.Drawing.Rectangle(x, y, ow, oh), System.Drawing.GraphicsUnit.Pixel);
                          ////以jpg格式保存缩略图WebControls 
                          ////File.Delete(thumbnailPath);  

                          //MemoryStream stream = new MemoryStream();
                          //bitmap.Save(stream, ImageFormat.Jpeg);
                          //bitmap.Dispose();

                          ImageConverter imgconv = new ImageConverter();
                          byte[] imgByte = (byte[])imgconv.ConvertTo(originalImage, typeof(byte[]));
                          int pictureIdx = workbook.AddPicture(imgByte, PictureType.PNG);
                          IDrawing patriarch = newCell.Sheet.CreateDrawingPatriarch();
                          HSSFClientAnchor anchor = new HSSFClientAnchor(5, 5, 1023, 250, i, rowNum, i, rowNum);
                          HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                          pict.LineWidth = 100;
                          //pict.Resize();
                      }
                      else
                      {
                          newCell.SetCellValue("图片不存在");
                      }

                      #endregion
                  }
              }
          }

          /// <summary>
          /// 将泛类型集合List类转换成DataTable
          /// </summary>
          /// <param name="list">泛类型集合</param>
          /// <param name="list">需要保留的字段</param>
          /// <param name="list">字段對應的中文名稱</param>
          /// <returns></returns>
          public static DataTable ListToDataTable<T>(IList<T> entitys, string[] field, string[] name)
          {
              //检查实体集合不能为空
              if (entitys == null || entitys.Count < 1)
              {
                  return null;
              }
              //取出第一个实体的所有Propertie
              Type entityType = entitys[0].GetType();
              PropertyInfo[] entityProperties = entityType.GetProperties();

              //生成DataTable的structure
              //生产代码中，应将生成的DataTable结构Cache起来，此处略
              DataTable dt = new DataTable();
              int[] index = new int[field.Length];
              for (int i = 0; i < field.Length; i++)
              {
                  for (int j = 0; j < entityProperties.Length; j++)
                  {
                      if (entityProperties[j].Name == field[i])
                      {
                          index[dt.Columns.Count] = j;
                          //dt.Columns.Add(entityProperties[i].Name, entityProperties[i].PropertyType);
                          dt.Columns.Add(name[i]);
                          break;
                      }
                  }
              }
              //将所有entity添加到DataTable中
              foreach (object entity in entitys)
              {
                  //检查所有的的实体都为同一类型
                  if (entity.GetType() != entityType)
                  {
                      throw new Exception("要转换的集合元素类型不一致");
                  }
                  object[] entityValues = new object[name.Length];
                  for (int i = 0; i < name.Length; i++)
                  {
                      entityValues[i] = entityProperties[index[i]].GetValue(entity, null);
                  }
                  dt.Rows.Add(entityValues);
              }
              return dt;
          }

          /// <summary>
          /// 对dt进行处理，改中文列名，列排序，设置列格式
          /// </summary>
          /// <param name="dt">dt</param>
          /// <param name="names">中文列名</param>
          /// <param name="field">列名（可以传null）</param>
          /// <param name="hs">列名和对应的数据类型(可传null)</param>
          /// <returns></returns>
          public static DataTable SetDataTableColmName(DataTable dt, string[] names, string[] field, Hashtable hs)
          {
              DataTable table = new DataTable();
              if (dt.Rows.Count > 0)
              {

                  if (field != null && field.Length > 0) //如果field不为空 则按照field的列的顺序进行排序，且field列和names的中文列名顺序一一对应
                  {
                      DataTable newDt = new DataTable();
                      //建表结构
                      for (int i = 0; i < field.Length; i++)
                      {
                          if (!dt.Columns.Contains(field[i].ToString()))
                          {
                              continue;
                          }
                          DataColumn cl = new DataColumn();
                          cl.ColumnName = field[i].ToString();

                          if (dt.Columns[cl.ColumnName].DataType != typeof(System.DateTime))
                          {
                              cl.DataType = dt.Columns[cl.ColumnName].DataType;
                          }
                          else
                          {
                              cl.DataType = typeof(String);
                          }

                          if (hs != null)
                          {
                              foreach (DictionaryEntry deHB in hs)
                              {
                                  if (deHB.Key.ToString().ToLower() == field[i].ToString().ToLower())
                                  {
                                      cl.DataType = (Type)deHB.Value;
                                      break;
                                  }
                              }
                          }
                          newDt.Columns.Add(cl);
                      }

                      //填数据
                      foreach (DataRow row in dt.Rows)
                      {
                          DataRow r = newDt.NewRow();
                          foreach (DataColumn clm in newDt.Columns)
                          {

                              r[clm] = row[clm.ColumnName];
                          }
                          newDt.Rows.Add(r);
                      }
                      //改列名
                      foreach (DataColumn dc in newDt.Columns)
                      {
                          dc.ColumnName = names[dc.Ordinal];
                      }
                      table = newDt;
                  }
                  else //field为空，列的顺序就按照dt的顺序，且names的中文名和dt的列顺序一一对应
                  {
                      DataTable dt1 = dt.Copy();

                      if (names.Length != dt1.Columns.Count)
                      {
                          throw new Exception("必须为给个列一一对应的中文名称！");
                      }
                      foreach (DataColumn dc in dt1.Columns)
                      {
                          //foreach (DictionaryEntry deHB in hs)
                          //{
                          //    if (deHB.Key.ToString() == dc.ColumnName.ToString())
                          //    {
                          //        dc.DataType = deHB.Value.GetType();
                          //    }
                          //}
                          dc.ColumnName = names[dc.Ordinal];
                      }
                      dt1.AcceptChanges();
                      table = dt1;
                  }
              }
              return table;

          }

          /// <summary>
          /// 合并单元格
          /// </summary>
          /// <param name="sheet">要合并单元格所在的sheet</param>
          /// <param name="rowstart">开始行的索引</param>
          /// <param name="rowend">结束行的索引</param>
          /// <param name="colstart">开始列的索引</param>
          /// <param name="colend">结束列的索引</param>
          public static void SetCellRangeAddress(ISheet sheet, int rowstart, int rowend, int colstart, int colend)
          {
              CellRangeAddress cellRangeAddress = new CellRangeAddress(rowstart, rowend, colstart, colend);
              sheet.AddMergedRegion(cellRangeAddress);
          }


          public static bool VerifyFileExist(string fileUrl)
          {
              bool isHave = false;
              if (File.Exists(fileUrl))
              {
                  isHave = true;
              }
              return isHave;
          }
    }
  
  /// <summary>
    /// 表头
    /// </summary>
    public class ReportTableHead
    {
        public string Title { get; set; }
        public int startCell { get; set; }
        public int endCell { get; set; }

        public int rowspan { get; set; }

    }

    /// <summary>
    /// 这是为了避免流被NPOI关闭而实现的流
    /// 当CanDispose为false时此流的dispose接口无效，仅当该值为true时有效
    /// </summary>
    public class ExcelStream : System.IO.MemoryStream
    {
        protected override void Dispose(bool disposing)
        {
            if (CanDispose)
            {
                base.Dispose(disposing);
            }
        }

        public bool CanDispose { get; set; }
    }
}
