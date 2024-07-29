  static public void exportExcel(Group group, TotalStatistics totalStatistics, bool? isInner, string specification, string strMachine, string strClass, string serialNumber)
  {
      if (totalStatistics.TotalCnt == 0)
          return;


      // 创建一个新的Excel工作簿
      NPOI.SS.UserModel.IWorkbook workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
      NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet("Sheet1");

      int maxColumn = 12;
      int mergeColumn = 1; //被合并的数量

      Specification sp = Helper.getSpecification(specification);

      // 数据准备

      // 写入数据到单元格
      int rowIndex = 0;
      sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("烟用爆珠质量检测系统");

      rowIndex++;
      sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("测量日期");
      sheet.GetRow(rowIndex).CreateCell(1).SetCellValue((totalStatistics.BeginTime.ToString("yyyy/MM/dd")));
      sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

      rowIndex++;
      sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("测量时间");
      sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(totalStatistics.BeginTime.ToString("HH:mm:ss") + " - " + totalStatistics.EndTime.ToString("HH:mm:ss"));
      sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

      rowIndex++;
      sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("牌号");
      sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(specification);
      sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

      rowIndex++;
      sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("机台");
      sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(strMachine);
      sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

      rowIndex++;
      sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("班次");
      sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(strClass);
      sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

      if (string.IsNullOrEmpty(serialNumber) == false)
      {
          rowIndex++;
          sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("序号");
          sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(serialNumber);
          sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);
      }

      rowIndex++;
      sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("爆珠总数");
      sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(totalStatistics.TotalCnt);
      sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

      rowIndex++;
      sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("不合格数");
      sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(totalStatistics.UnQualifiedCnt);
      sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

      rowIndex++;
      sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("合格率");
      sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(Math.Round((totalStatistics.QualifiedCnt / (double)totalStatistics.TotalCnt) * 100.0, 2).ToString("0.00") + "%");
      sheet.GetRow(rowIndex).CreateCell(maxColumn - 1);

      //int titleIndexFinal = rowIndex + 1;

      //// 合并单元格并设置格式
      //CellRangeAddress titleRange = new CellRangeAddress(0, 0, 0, maxColumn - 1);
      //sheet.AddMergedRegion(titleRange);
      //for (int i = 1; i < titleIndexFinal; i++)
      //{
      //    titleRange = new CellRangeAddress(i, i, 1, maxColumn - 1);
      //    sheet.AddMergedRegion(titleRange);
      //}

      // 初始化行和列的索引
      //rowIndex++;
      int maxColumnsRowIndex = rowIndex;
      //int nextRow = rowIndex;//5
      int columnsIndex = 0;


      rowIndex++;
      columnsIndex = 0;
      NPOI.SS.UserModel.IRow dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(columnsIndex++).SetCellValue("项目");
      dataRow.CreateCell(columnsIndex++).SetCellValue("均值");
      dataRow.CreateCell(columnsIndex++).SetCellValue("最大");
      dataRow.CreateCell(columnsIndex++).SetCellValue("最小");

      columnsIndex = 0;
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(columnsIndex++).SetCellValue("长轴");
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MeanLongAxis, RemainCntHelper.Diameter.RemainCnt).ToString(RemainCntHelper.Diameter.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MaxLongAxis, RemainCntHelper.Diameter.RemainCnt).ToString(RemainCntHelper.Diameter.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MinLongAxis, RemainCntHelper.Diameter.RemainCnt).ToString(RemainCntHelper.Diameter.StrFormat));

      columnsIndex = 0;
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(columnsIndex++).SetCellValue("短轴");
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MeanShortAxis, RemainCntHelper.Diameter.RemainCnt).ToString(RemainCntHelper.Diameter.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MaxShortAxis, RemainCntHelper.Diameter.RemainCnt).ToString(RemainCntHelper.Diameter.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MinShortAxis, RemainCntHelper.Diameter.RemainCnt).ToString(RemainCntHelper.Diameter.StrFormat));

      columnsIndex = 0;
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(columnsIndex++).SetCellValue("圆度");
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MeanCircleRate, RemainCntHelper.CircleRate.RemainCnt).ToString(RemainCntHelper.CircleRate.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MaxCircleRate, RemainCntHelper.CircleRate.RemainCnt).ToString(RemainCntHelper.CircleRate.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MinCircleRate, RemainCntHelper.CircleRate.RemainCnt).ToString(RemainCntHelper.CircleRate.StrFormat));

      columnsIndex = 0;
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(columnsIndex++).SetCellValue("粒径");
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MeanParticleSize, RemainCntHelper.Diameter.RemainCnt).ToString(RemainCntHelper.Diameter.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MaxParticleSize, RemainCntHelper.Diameter.RemainCnt).ToString(RemainCntHelper.Diameter.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MinParticleSize, RemainCntHelper.Diameter.RemainCnt).ToString(RemainCntHelper.Diameter.StrFormat));

      columnsIndex = 0;
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(columnsIndex++).SetCellValue("圆整度");
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MeanRoundness, RemainCntHelper.Roundness.RemainCnt).ToString(RemainCntHelper.Roundness.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MaxRoundness, RemainCntHelper.Roundness.RemainCnt).ToString(RemainCntHelper.Roundness.StrFormat));
      dataRow.CreateCell(columnsIndex++).SetCellValue(Math.Round(totalStatistics.MinRoundness, RemainCntHelper.Roundness.RemainCnt).ToString(RemainCntHelper.Roundness.StrFormat));


      columnsIndex = 0;
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("圆度异常");
      dataRow.CreateCell(1).SetCellValue( totalStatistics.EcllipseCnt);
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("圆度合格率" );
      dataRow.CreateCell(1).SetCellValue( Math.Round((totalStatistics.TotalCnt - totalStatistics.EcllipseCnt) / (double)totalStatistics.TotalCnt * 100.0, 2).ToString("0.00") + "%");
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("圆整度异常");
      dataRow.CreateCell(1).SetCellValue( totalStatistics.RoundnessCnt);
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("圆整度合格率" );
      dataRow.CreateCell(1).SetCellValue(Math.Round((totalStatistics.TotalCnt - totalStatistics.RoundnessCnt) / (double)totalStatistics.TotalCnt * 100.0, 2).ToString("0.00") + "%");
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("爆珠过大" );
      dataRow.CreateCell(1).SetCellValue( totalStatistics.BiggerCnt);
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("爆珠过小");
      dataRow.CreateCell(1).SetCellValue( totalStatistics.SmallerCnt);
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("粒径过大" );
      dataRow.CreateCell(1).SetCellValue( totalStatistics.ParticleSizeBiggerCnt);
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("粒径过小" );
      dataRow.CreateCell(1).SetCellValue( totalStatistics.ParticleSizeSmallerCnt);
      if (sp != null)
      {
          rowIndex++;
          dataRow = sheet.CreateRow(rowIndex);
          dataRow.CreateCell(0).SetCellValue(("粒径"
           ));
          dataRow.CreateCell(1).SetCellValue(( " 上允差:" + Math.Round(sp.ParticleSizeUp - sp.ParticleSize, 5).ToString()
           + " 下允差:" + Math.Round(sp.ParticleSize - sp.ParticleSizeDown, 5).ToString()
           ));
      }
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("粒径合格");
      dataRow.CreateCell(1).SetCellValue( totalStatistics.ParticleSizeCnt);
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("粒径合格率" );
      dataRow.CreateCell(1).SetCellValue(Math.Round((totalStatistics.ParticleSizeCnt) / (double)totalStatistics.TotalCnt * 100.0, 2).ToString("0.00") + "%");
      if (sp != null)
      {
          rowIndex++;
          dataRow = sheet.CreateRow(rowIndex);
          dataRow.CreateCell(0).SetCellValue("粒径标准2"
             );
          dataRow.CreateCell(1).SetCellValue( " 上允差:" + Math.Round(sp.ParticleSizeUp2 - sp.ParticleSize, 5).ToString()
              + " 下允差:" + Math.Round(sp.ParticleSize - sp.ParticleSizeDown2, 5).ToString());
      }
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("粒径超标2" );
      dataRow.CreateCell(1).SetCellValue(totalStatistics.ParticleSize2UnQualifiedCnt);
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("粒径合格率2" );
      dataRow.CreateCell(1).SetCellValue( Math.Round((totalStatistics.TotalCnt - totalStatistics.ParticleSize2UnQualifiedCnt) / (double)totalStatistics.TotalCnt * 100.0, 2).ToString("0.00") + "%");
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue(StringDescription.Eccentricity );
      dataRow.CreateCell(1).SetCellValue( totalStatistics.EccentricityCnt);
      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue(StringDescription.HeteroSharp );
      dataRow.CreateCell(1).SetCellValue( totalStatistics.HeteroMorphismCnt);
      var xClassification = ConfigHelper.getRootConfig().SelectSingleNode("classification");
      if (xClassification.Attributes["t1"].Value == "1")
      {
          rowIndex++;
          dataRow = sheet.CreateRow(rowIndex);
          dataRow.CreateCell(0).SetCellValue(StringDescription.Classified1);
          dataRow.CreateCell(1).SetCellValue( totalStatistics.Classified1Cnt);
          rowIndex++;
          dataRow = sheet.CreateRow(rowIndex);
          dataRow.CreateCell(0).SetCellValue(StringDescription.Classified1 + "率"  );
          dataRow.CreateCell(1).SetCellValue( Math.Round((totalStatistics.Classified1Cnt) / (double)totalStatistics.TotalCnt * 100.0, 2).ToString("0.0") + "%");
      }
      if (xClassification.Attributes["t2"].Value == "1")
      {
          rowIndex++;
          dataRow = sheet.CreateRow(rowIndex);
          dataRow.CreateCell(0).SetCellValue(StringDescription.Classified2  );
          dataRow.CreateCell(1).SetCellValue( totalStatistics.Classified2Cnt);
      }
      if (xClassification.Attributes["t3"].Value == "1")
      {
          rowIndex++;
          dataRow = sheet.CreateRow(rowIndex);
          dataRow.CreateCell(0).SetCellValue(StringDescription.Classified3 );
          dataRow.CreateCell(1).SetCellValue(totalStatistics.Classified3Cnt);
      }
      if (xClassification.Attributes["t4"].Value == "1")
      {
          rowIndex++;
          dataRow = sheet.CreateRow(rowIndex);
          dataRow.CreateCell(0).SetCellValue(StringDescription.Classified4 );
          dataRow.CreateCell(1).SetCellValue(totalStatistics.Classified4Cnt);
      }

      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue(StringDescription.Fill );
      dataRow.CreateCell(1).SetCellValue(totalStatistics.FillCnt);

      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue(StringDescription.Blank);
      dataRow.CreateCell(1).SetCellValue( totalStatistics.BlankCnt);

      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue(StringDescription.HeteroColor );
      dataRow.CreateCell(1).SetCellValue( totalStatistics.HeteroColorCnt);

      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue(StringDescription.ConvexHeightCnt);
      dataRow.CreateCell(1).SetCellValue( totalStatistics.ConvexHeightCnt);

      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue(StringDescription.ConvexHeightCnt + "率");
      dataRow.CreateCell(1).SetCellValue( Math.Round((totalStatistics.ConvexHeightCnt) / (double)totalStatistics.TotalCnt * 100.0, 2).ToString("0.0") + "%");

      rowIndex++;
      dataRow = sheet.CreateRow(rowIndex);
      dataRow.CreateCell(0).SetCellValue("详细数据");


      List<Ball> lsBallsExist = group.getBalls();
      List<Ball> lsBalls = new List<Ball>();
      foreach (Ball b in lsBallsExist)
      {
          if (!b.IsExist)
              continue;
          lsBalls.Add(b);
      }
      if (lsBalls.Count != 0)
      {
          columnsIndex = 0;
          rowIndex++;
          dataRow = sheet.CreateRow(rowIndex);

          dataRow.CreateCell(columnsIndex++).SetCellValue("序号");
          dataRow.CreateCell(columnsIndex++).SetCellValue("时间");
          dataRow.CreateCell(columnsIndex++).SetCellValue("长轴");
          dataRow.CreateCell(columnsIndex++).SetCellValue("短轴");
          dataRow.CreateCell(columnsIndex++).SetCellValue("圆度");
          dataRow.CreateCell(columnsIndex++).SetCellValue("粒径");
          dataRow.CreateCell(columnsIndex++).SetCellValue("圆整度");
          dataRow.CreateCell(columnsIndex++).SetCellValue("分类");
          dataRow.CreateCell(columnsIndex++).SetCellValue("异色");
          dataRow.CreateCell(columnsIndex++).SetCellValue("灰度");
          dataRow.CreateCell(columnsIndex++).SetCellValue(StringDescription.HeteroSharp);
          dataRow.CreateCell(columnsIndex++).SetCellValue(StringDescription.Eccentricity);
          dataRow.CreateCell(columnsIndex++).SetCellValue("拖尾");

          int count = 0;
          foreach (Ball b in lsBalls)
          {
              columnsIndex = 0;
              rowIndex++;
              dataRow = sheet.CreateRow(rowIndex);
              count++;

              dataRow.CreateCell(columnsIndex++).SetCellValue(count);
              dataRow.CreateCell(columnsIndex++).SetCellValue(b.TestTime.Value.ToString("HH:mm:ss.fff"));
              dataRow.CreateCell(columnsIndex++).SetCellValue(b.LongAxis.ToString(RemainCntHelper.Diameter.StrFormat));
              dataRow.CreateCell(columnsIndex++).SetCellValue(b.ShortAxis.ToString(RemainCntHelper.Diameter.StrFormat));
              dataRow.CreateCell(columnsIndex++).SetCellValue(b.Oval.ToString(RemainCntHelper.CircleRate.StrFormat));
              dataRow.CreateCell(columnsIndex++).SetCellValue(b.ParticleSize.ToString(RemainCntHelper.Diameter.StrFormat));
              dataRow.CreateCell(columnsIndex++).SetCellValue(b.Roundness.ToString(RemainCntHelper.Roundness.StrFormat));

              string typeRet = "";
              if (b.IsClassified1)
                  typeRet = StringDescription.Classified1 + " ";
              if (b.IsClassified2)
                  typeRet += StringDescription.Classified2 + " ";
              if (b.IsClassified3)
                  typeRet += StringDescription.Classified3 + " ";
              if (b.IsClassified4)
                  typeRet += StringDescription.Classified4 + " ";
              if (typeRet == "")
                  typeRet = "正常";

              dataRow.CreateCell(columnsIndex++).SetCellValue(typeRet);
              dataRow.CreateCell(columnsIndex++).SetCellValue(b.IsHeteroColor ? "是      " : "否      ");

              string grayRet = b.GrayLv.ToString();
              string morphismStr = Math.Round(b.IsHetoeroMorphism, 0).ToString();
              string eccentricityStr = b.Eccentricity.ToString();
              string chromatismStr = b.Chromatism.ToString();
              string convexHeightStr = b.ConvexHeight.ToString();
              if (sp != null)
              {
                  if (b.GrayLv < sp.GrayFillThreshold)
                      grayRet = "实心" + " " + grayRet;
                  else if (b.GrayLv > sp.GrayBlankThreshold)
                      grayRet = "空心" + " " + grayRet;
                  morphismStr = (b.IsHetoeroMorphism < sp.MorphismThreshold ? "否" : "是") + " " + morphismStr;
                  eccentricityStr = (b.Eccentricity < sp.EccentricityThreshold ? "否" : "是") + " " + eccentricityStr;
                  chromatismStr = (b.Chromatism < sp.ChromatismThreshold ? "否" : "是") + " " + chromatismStr;
                  convexHeightStr = (b.ConvexHeight < sp.ConvexHeightThreshold ? "否" : "是") + " " + convexHeightStr;
              }

              dataRow.CreateCell(columnsIndex++).SetCellValue(grayRet);
              dataRow.CreateCell(columnsIndex++).SetCellValue(morphismStr);
              dataRow.CreateCell(columnsIndex++).SetCellValue(eccentricityStr);
              dataRow.CreateCell(columnsIndex++).SetCellValue(convexHeightStr);

          }

      }
      maxColumnsRowIndex = rowIndex;

      //// 设置边框
      //var borderStyle = NPOI.SS.UserModel.BorderStyle.Medium;
      //var black = NPOI.SS.UserModel.IndexedColors.Black.Index;
      for (int i = 0; i <= rowIndex; i++)
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
                  //cellStyle.BorderBottom = borderStyle;
                  //cellStyle.BorderLeft = borderStyle;
                  //cellStyle.BorderRight = borderStyle;
                  //cellStyle.BorderTop = borderStyle;
                  //cellStyle.BottomBorderColor = black;
                  //cellStyle.LeftBorderColor = black;
                  //cellStyle.RightBorderColor = black;
                  //cellStyle.TopBorderColor = black;
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

      for (int i = 0; i <= rowIndex; i++)
      {
          sheet.GetRow(i).Cells.ForEach(c => c.CellStyle.SetFont(font));
      }

      // 自适应列宽
      // 灵活调整 参数  数字为最多列位置
      AutoColumnWidth(sheet, sheet.GetRow(maxColumnsRowIndex).PhysicalNumberOfCells);

      if (workbook == null)
      {
          MBox.Show("警告", "导出失败", MBoxType.Alarm);
          return;
      }
      string name = totalStatistics.BeginTime.ToString("yyyy_MM_dd__HH_mm_ss");
      if (string.IsNullOrEmpty(serialNumber) == false)
      {
          name += "__" + serialNumber;
      }
      name = name + ".xlsx";

      string path = AppDomain.CurrentDomain.BaseDirectory + "excel";
      Directory.CreateDirectory(path);
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

  public static void AutoColumnWidth(NPOI.SS.UserModel.ISheet sheet, int cols)
  {
      for (int col = 0; col <= cols; col++)
      {
          sheet.AutoSizeColumn(col);//自适应宽度，但是其实还是比实际文本要宽
          int columnWidth = (int)(sheet.GetColumnWidth(col) / 256);//获取当前列宽度
          for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
          {
              NPOI.SS.UserModel.IRow row = sheet.GetRow(rowIndex);
              NPOI.SS.UserModel.ICell cell = row.GetCell(col);
              if (cell != null)
              {
                  int contextLength = Encoding.UTF8.GetBytes(cell.ToString()).Length;//获取当前单元格的内容宽度
                  columnWidth = contextLength >columnWidth  ? contextLength : columnWidth;
              }
          }
          sheet.SetColumnWidth(col, columnWidth * 256 + 500);// 调整值

      }
  }
