      private void mItemExport_click(object sender, RoutedEventArgs e)
      {
          try
          {
              if (_dgGroup.SelectedIndex == -1)
              {
                  MBox.Show("请从右下侧列表中选择要导出的分组");
                  return;
              }
              var group = _groups[_dgGroup.SelectedIndex];

              string path = RTHelper.openDirectoryDialog();
              if (path == null || path == "")
                  return;
              gt_tip.show();
              gt_tip.setCancelVisibility(false);
              gt_tip.setMainTip("导出开始……");
              gt_tip.setSubTip("");

              Action dltExport = delegate
              {
                  try
                  {
                      gt_tip.setMainTip("导出中……");

                      // 创建一个新的Excel工作簿
                      NPOI.SS.UserModel.IWorkbook workbook = null;
                      workbook = getExcelData();

                      if (workbook == null)
                      {
                          MBox.Show("警告", "导出失败", MBoxType.Alarm);
                          return;
                      }
                      string name = group.BeginTime.ToString("yyyyMMdd HH.mm.ss") + ".xlsx";
                      path = path + "\\" + name;
                      if (System.IO.File.Exists(path))
                          System.IO.File.Delete(path);
                      // 将工作簿保存到文件
                      using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                      {
                          workbook.Write(fs);
                      }

                      if (MBox.Show("提示", "导出成功，是否打开？", new string[] { "是", "否" }) == MBoxResult.Yes)
                      {
                          System.Diagnostics.Process.Start("explorer.exe", path);
                      }
                  }
                  catch (Exception ex)
                  {
                      RunLog.save(ex.ToString());
                      MBox.Show("警告", "导出失败", MBoxType.Alarm);
                  }
                  finally
                  {
                      gt_tip.hide();
                  }
              };

              dltExport.BeginInvoke(null, null);
          }
          catch (Exception ex)
          {
              RunLog.save(ex.ToString());
          }
      }





  private NPOI.SS.UserModel.IWorkbook getExcelData()
  {
      //使用主线程 这样是为了配合gt_tip的显示和控件获取
      NPOI.SS.UserModel.IWorkbook workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
      NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet("Sheet1");

      Application.Current.Dispatcher.Invoke(delegate
      {
          int maxColumn = 0;
          int mergeColumn = 1; //被合并的数量
          DataView viewDetail = ds_detail.getDataGrid().ItemsSource as DataView;
          DataTable dataTableDetail = viewDetail.Table;
          var group = _groups[_dgGroup.SelectedIndex];
          maxColumn = ds_detail.getDataGrid().Columns.Count(x => x.Visibility == Visibility.Visible);
          if (dataTableDetail.Rows.Count == 0 || maxColumn <= 1)
          {
              workbook = null;
          }
          else
          {
              var detailColumns = ds_detail.getDataGrid().Columns;
              {

                  // 数据准备
                  string title = "历史数据";
                  string beginTime = group.BeginTime.ToString("yyyy/MM/dd HH:mm:ss");
                  string specification = group.Specification;
                  string machine = group.Machine;
                  string turn = group.Turn;

                  // 写入数据到单元格
                  int rowIndex = 0;
                  sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(title);
                  sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);//占位 边框

                  rowIndex++;
                  sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("测量时间");
                  sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(beginTime);
                  sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

                  rowIndex++;
                  sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("牌号");
                  sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(specification);
                  sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

                  rowIndex++;
                  sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("机台");
                  sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(machine);
                  sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

                  rowIndex++;
                  sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("班次");
                  sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(turn);
                  sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

                  // 合并单元格并设置格式
                  CellRangeAddress titleRange = new CellRangeAddress(0, 0, 0, maxColumn - 1);
                  CellRangeAddress beginTimeRange = new CellRangeAddress(1, 1, 1, maxColumn - 1);
                  CellRangeAddress specificationRange = new CellRangeAddress(2, 2, 1, maxColumn - 1);
                  CellRangeAddress machineRange = new CellRangeAddress(3, 3, 1, maxColumn - 1);
                  CellRangeAddress turnRange = new CellRangeAddress(4, 4, 1, maxColumn - 1);

                  sheet.AddMergedRegion(titleRange);
                  sheet.AddMergedRegion(beginTimeRange);
                  sheet.AddMergedRegion(specificationRange);
                  sheet.AddMergedRegion(machineRange);
                  sheet.AddMergedRegion(turnRange);

                  // 初始化行和列的索引
                  rowIndex++;
                  int headerRowIndex = rowIndex;
                  int nextRow = rowIndex;//5
                  int columnsIndex = 0;

                  // 写入列头
                  NPOI.SS.UserModel.IRow headerRow = sheet.CreateRow(nextRow);
                  foreach (var column in detailColumns)
                  {
                      if (column.Visibility == Visibility.Visible)
                      {
                          headerRow.CreateCell(columnsIndex).SetCellValue((string)column.Header);
                          columnsIndex++;
                      }
                  }

                  // 写入数据
                  foreach (DataRow row in dataTableDetail.Rows)
                  {
                      nextRow++;
                      columnsIndex = 0;
                      NPOI.SS.UserModel.IRow dataRow = sheet.CreateRow(nextRow);
                      foreach (var column in detailColumns)
                      {
                          if (column.Visibility == Visibility.Visible)
                          {
                              dataRow.CreateCell(columnsIndex).SetCellValue(row[(string)column.Header].ToString());
                              columnsIndex++;
                          }
                      }
                  }

                  nextRow++;
                  int mergeBeginRowIndex = nextRow;

                  // 假设 ds_statistics.getDataGrid().ItemsSource 返回 DataView
                  DataView viewStatistics = ds_statistics.getDataGrid().ItemsSource as DataView;
                  DataTable dataTableStatistics = viewStatistics.Table;
                  var statisticsColumns = ds_statistics.getDataGrid().Columns;
                  columnsIndex = 0;

                  sheet.CreateRow(nextRow);
                  // 写入列头
                  for (int i = 0; i < statisticsColumns.Count; i++)
                  {
                      if (statisticsColumns[i].Visibility == Visibility.Visible)
                      {
                          sheet.GetRow(nextRow).CreateCell(columnsIndex).SetCellValue(statisticsColumns[i].Header.ToString());
                          if (columnsIndex == 0)
                          {
                              columnsIndex = columnsIndex + 1 + mergeColumn;//合并2>1
                          }
                          else
                          {
                              columnsIndex = columnsIndex + 1;
                          }
                      }
                  }

                  // 写入数据
                  for (int i = 0; i < dataTableStatistics.Rows.Count; i++)
                  {
                      nextRow++;
                      columnsIndex = 0;
                      sheet.CreateRow(nextRow);
                      for (int j = 0; j < statisticsColumns.Count; j++)
                      {
                          if (statisticsColumns[j].Visibility == Visibility.Visible)
                          {
                              if (dataTableStatistics.Rows[i][0].ToString().Contains("合格"))
                              {
                                  sheet.GetRow(nextRow).CreateCell(columnsIndex).SetCellValue(" " + dataTableStatistics.Rows[i][j].ToString());
                              }
                              else
                              {
                                  sheet.GetRow(nextRow).CreateCell(columnsIndex).SetCellValue(dataTableStatistics.Rows[i][j].ToString());
                              }
                              if (columnsIndex == 0)
                              {
                                  columnsIndex = columnsIndex + 1 + mergeColumn;//合并2>1
                              }
                              else
                              {
                                  columnsIndex = columnsIndex + 1;
                              }
                          }
                      }
                  }

                  // 合并单元格
                  for (int i = mergeBeginRowIndex; i <= nextRow; i++)
                  {
                      sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(i, i, 0, mergeColumn));
                  }
               

                  // 设置边框
                  var borderStyle = NPOI.SS.UserModel.BorderStyle.Medium;
                  var black = NPOI.SS.UserModel.IndexedColors.Black.Index;
                  for (int i = 0; i <= nextRow; i++)
                  {
                      var row = sheet.GetRow(i);
                      //for (int j = 0; j < row.PhysicalNumberOfCells; j++)//合并
                      for (int j = 0; j < row.LastCellNum; j++)//不合并   
                      {
                          //var cell = row.GetCell(j);//不合并
                          var cell = row.GetCell(j, NPOI.SS.UserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK); //边框用 写法       
                          if (cell != null)
                          {
                              var cellStyle = workbook.CreateCellStyle();
                              cellStyle.BorderBottom = borderStyle;
                              cellStyle.BorderLeft = borderStyle;
                              cellStyle.BorderRight = borderStyle;
                              cellStyle.BorderTop = borderStyle;
                              cellStyle.BottomBorderColor = black;
                              cellStyle.LeftBorderColor = black;
                              cellStyle.RightBorderColor = black;
                              cellStyle.TopBorderColor = black;
                              // 设置水平居中
                              cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                              cell.CellStyle = cellStyle;
                          }
                      }
                  }

                  // 设置字体大小和字体
                  var font = workbook.CreateFont();
                  font.FontHeightInPoints = 12;
                  font.FontName = "微软雅黑";

                  for (int i = 0; i <= nextRow; i++)
                  {
                      sheet.GetRow(i).Cells.ForEach(c => c.CellStyle.SetFont(font));
                  }

                  // 自适应列宽
                  // 灵活调整 参数  数字为最多列位置
                  AutoColumnWidth(sheet, sheet.GetRow(headerRowIndex).PhysicalNumberOfCells);

                  //// 加粗内容
                  //var fontBold = workbook.CreateFont();
                  //fontBold.FontHeightInPoints = 12;
                  //fontBold.FontName = "微软雅黑";
                  //fontBold.IsBold = true;
                  //for (int i = 0; i <= nextRow; i++)
                  //{
                  //    excelWS.GetRow(i).Cells[0].CellStyle.SetFont(fontBold);
                  //}

              }
          }

      });

      return workbook;
  }


  public void AutoColumnWidth(NPOI.SS.UserModel.ISheet sheet, int cols)
  {
      for (int col = 0; col <= cols; col++)
      {
          sheet.AutoSizeColumn(col);//自适应宽度，但是其实还是比实际文本要宽
          int columnWidth = (int)(sheet.GetColumnWidth(col) / 256);//获取当前列宽度
          for (int rowIndex = 2; rowIndex <= sheet.LastRowNum; rowIndex++)
          {
              NPOI.SS.UserModel.IRow row = sheet.GetRow(rowIndex);
              NPOI.SS.UserModel.ICell cell = row.GetCell(col);
              if (cell != null)
              {
                  int contextLength = Encoding.UTF8.GetBytes(cell.ToString()).Length;//获取当前单元格的内容宽度
                  columnWidth = columnWidth < contextLength ? contextLength : columnWidth;
              }
          }
          sheet.SetColumnWidth(col, columnWidth * 256 + 500);// 调整值

      }
  }

