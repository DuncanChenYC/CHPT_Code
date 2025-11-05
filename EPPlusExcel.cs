using System;
using System.Data;
using System.Linq;
using System.Text;
using System.IO;
using DataProviderInfrastructure;
namespace ERP
{
	public class EPPlusExcel
	{
		public OfficeOpenXml.ExcelPackage pck;
		public OfficeOpenXml.ExcelWorksheet[] workSheetAry;
		public System.Drawing.Color dtHeaderColor;
		private string UnlockPath = "";

		public EPPlusExcel()
		{
			UnlockPath = GetUnlockPath();

            this.Init(1);
		}

		public EPPlusExcel(int sheetCount)
		{
			this.Init(sheetCount);
		}

		private void Init(int sheetCount)
		{
			dtHeaderColor = System.Drawing.Color.Yellow;
			workSheetAry = new OfficeOpenXml.ExcelWorksheet[sheetCount];
		}

		/// <summary>
		/// 先解密再讀取Excel內容
		/// </summary>
		/// <param name="isDeBugMode">【True】回傳執行幾次後，才成功</param>
		/// <param name="FileUpload">FileUpload元件</param>
		/// <param name="Sleepms">毫秒</param>
		/// <param name="WhileCount">執行While最大次數</param>
		/// <param name="ExcelSheetName">Excel的Sheet名稱</param>
		/// <param name="sbMsg">回覆錯誤訊息</param>
		/// <returns></returns>
		public System.Data.DataTable getDataTableFromExcelBeforeUnlock(bool isDeBugMode, System.Web.UI.WebControls.FileUpload FileUpload, int Sleepms, int WhileCount, string ExcelSheetName, ref System.Text.StringBuilder sbMsg)
		{
			//string UnlockPath = @"\\192.168.1.222\erp\ERPAP\App_Data\Unlock\";
			return getDataTableFromExcelBeforeUnlock(isDeBugMode, UnlockPath, FileUpload, Sleepms, WhileCount, ExcelSheetName, ref sbMsg);
		}
		/// <summary>
		/// 先解密再讀取Excel內容
		/// </summary>
		/// <param name="IsDeBugMode">【True】回傳執行幾次後，才成功</param>
		/// <param name="UnlockPath">解密路徑【\\192.168.1.222\erp\ERPAP\App_Data\Unlock\】</param>
		/// <param name="FileUpload">FileUpload元件</param>
		/// <param name="Sleepms">毫秒</param>
		/// <param name="WhileCount">執行While最大次數</param>
		/// <param name="ExcelSheetName">Excel的Sheet名稱</param>
		/// <param name="sbMsg">回覆錯誤訊息</param>
		/// <returns></returns>
		public System.Data.DataTable getDataTableFromExcelBeforeUnlock(bool IsDeBugMode, string UnlockPath, System.Web.UI.WebControls.FileUpload FileUpload, int Sleepms, int WhileCount, string ExcelSheetName, ref System.Text.StringBuilder sbMsg)
		{
			System.Data.DataTable dt = new DataTable();
			EmuSPUserClass MyEmuSPUserClass = new EmuSPUserClass();
			try
			{
				if (Sleepms <= 0)
					Sleepms = 500;

				if (MyEmuSPUserClass.EmuSPUser("CHPT"))
				{
					string ServerPath = UnlockPath;
					string FilePath = ServerPath + FileUpload.FileName;
					FileUpload.SaveAs(FilePath);

					int index = 1;
					while (index <= WhileCount)
					{
						try
						{
							System.Threading.Thread.Sleep(Sleepms);

							dt = getDataTableFromExcel(FilePath, ExcelSheetName);

							if (IsDeBugMode)
								sbMsg.AppendFormat("執行次數【{0}】", index);
							index = WhileCount + 1;
						}
						catch (Exception ex)
						{
							if (index == WhileCount)
								sbMsg.AppendFormat("{0}\n執行次數【{1}】已達上限！", ex.Message, index);

							if (ex.Message == "檢查Sheet名稱是否不存在!")
							{
								index = WhileCount + 1;
								sbMsg.Append(ex.Message);
							}
						}
						index++;
					}

					System.IO.File.Delete(FilePath);
				}
			}
			catch (Exception ex)
			{
				sbMsg.Append(ex.Message);
			}
			finally
			{
				//釋放模擬權限
				MyEmuSPUserClass.UndoEmuSPUser();
			}

			return dt;
		}

        /// <summary>
        /// 先解密再讀取Excel內容
        /// </summary>
        /// <param name="isDeBugMode">【True】回傳執行幾次後，才成功</param>
        /// <param name="FileUpload">FileUpload元件</param>
        /// <param name="Sleepms">毫秒</param>
        /// <param name="WhileCount">執行While最大次數</param>
        /// <param name="ExcelSheetName">Excel的Sheet名稱</param>
        /// <param name="sbMsg">回覆錯誤訊息</param>
        /// <returns></returns>
        public System.Data.DataSet getDataSetFromExcelBeforeUnlock(bool isDeBugMode, System.Web.UI.WebControls.FileUpload FileUpload, int Sleepms, int WhileCount, string[] ExcelSheetName, ref System.Text.StringBuilder sbMsg)
        {
            //string UnlockPath = @"\\192.168.1.222\erp\ERPAP\App_Data\Unlock\";
            return getDataSetFromExcelBeforeUnlock(isDeBugMode, UnlockPath, FileUpload, Sleepms, WhileCount, ExcelSheetName, ref sbMsg);
        }

        /// <summary>
        /// 先解密再讀取Excel內容
        /// </summary>
        /// <param name="IsDeBugMode">【True】回傳執行幾次後，才成功</param>
        /// <param name="UnlockPath">解密路徑【\\192.168.1.222\erp\ERPAP\App_Data\Unlock\】</param>
        /// <param name="FileUpload">FileUpload元件</param>
        /// <param name="Sleepms">毫秒</param>
        /// <param name="WhileCount">執行While最大次數</param>
        /// <param name="ExcelSheetName">Excel的Sheet名稱</param>
        /// <param name="sbMsg">回覆錯誤訊息</param>
        /// <returns></returns>
        private System.Data.DataSet getDataSetFromExcelBeforeUnlock(bool IsDeBugMode, string UnlockPath, System.Web.UI.WebControls.FileUpload FileUpload, int Sleepms, int WhileCount, string[] ExcelSheetName, ref System.Text.StringBuilder sbMsg)
        {
            DataSet ds = new DataSet();
            ds.DataSetName = "dsExcel";

            EmuSPUserClass MyEmuSPUserClass = new EmuSPUserClass();
            try
            {
                if (Sleepms <= 0)
                    Sleepms = 500;

                if (MyEmuSPUserClass.EmuSPUser("CHPT"))
                {
                    string ServerPath = UnlockPath;
                    string FilePath = ServerPath + FileUpload.FileName;
                    FileUpload.SaveAs(FilePath);

                    int index = 1;
                    while (index <= WhileCount)
                    {
                        try
                        {
                            System.Threading.Thread.Sleep(Sleepms);

                            for (int i = 0; i < ExcelSheetName.Length; i++)
                            {
                                System.Data.DataTable dt = new DataTable();

                                dt = getDataTableFromExcel(FilePath, ExcelSheetName[i]);
                                if (ds.Tables.Contains(ExcelSheetName[i]) == false) //判斷是否DataSet中,有包含該DataTable Name
                                {
                                    ds.Tables.Add(dt);
                                    dt.TableName = ExcelSheetName[i]; //設定Sheet名稱
                                }
                            }

                            if (IsDeBugMode)
                                sbMsg.AppendFormat("執行次數【{0}】", index);
                            index = WhileCount;
                        }
                        catch (Exception ex)
                        {
                            if (index == WhileCount)
                                sbMsg.AppendFormat("{0}\n執行次數【{1}】已達上限！", ex.Message, index);

                            if (ex.Message == "檢查Sheet名稱是否不存在!" || ex.Message.Contains("檢查Excel欄位是否不存在!") == true)
                            {
                                index = WhileCount;
                                sbMsg.Append(ex.Message);
                            }
                        }
                        index++;
                    }

                    System.IO.File.Delete(FilePath);
                }
            }
            catch (Exception ex)
            {
                sbMsg.Append(ex.Message);
            }
            finally
            {
                //釋放模擬權限
                MyEmuSPUserClass.UndoEmuSPUser();
            }

            return ds;
        }

		/// <summary>
		/// 讀2007↑Excel檔案到DataTable
		/// </summary>
		/// <param name="path">Excel路徑</param>
		/// <returns></returns>
		public System.Data.DataTable getDataTableFromExcel(string path, string sheetName)
		{
			using (var pck = new OfficeOpenXml.ExcelPackage())
			{
				using (var stream = System.IO.File.OpenRead(path))
				{
					pck.Load(stream);
				}
				var ws = pck.Workbook.Worksheets[sheetName];
				System.Data.DataTable tb = new System.Data.DataTable();
				if (ws != null)
				{
					bool hasHeader = true;
                    try 
                    {
                        foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                        {
                            tb.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                        }
                    }
                    catch (Exception ex)
                    {
                    throw new Exception("Sheet[" + sheetName + "]:檢查Excel欄位是否不存在!");
                    }
					
					var startRow = hasHeader ? 2 : 1;
					for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
					{
						var wsRow = ws.Cells[rowNum, 1, rowNum, tb.Columns.Count];
						var row = tb.NewRow();
						foreach (var cell in wsRow)
						{
							row[cell.Start.Column - 1] = cell.Text;
						}
						tb.Rows.Add(row);
					}
				}
				else
					throw new Exception("檢查Sheet名稱是否不存在!");
				return tb;
			}
		}

		public System.Data.DataTable getDataTableFromExcelBeforeUnlock(bool IsDeBugMode, string strPath, int Sleepms, int WhileCount, string ExcelSheetName, ref System.Text.StringBuilder sbMsg)
		{
			//string UnlockPath = @"\\192.168.1.222\erp\ERPAP\App_Data\Unlock\";

			System.Data.DataTable dt = new DataTable();
			EmuSPUserClass MyEmuSPUserClass = new EmuSPUserClass();
			try
			{
				if (Sleepms <= 0)
					Sleepms = 500;

				if (MyEmuSPUserClass.EmuSPUser("CHPT"))
				{
					System.IO.FileInfo objFile = new System.IO.FileInfo(strPath);
					string ServerPath = UnlockPath;
					string FilePath = ServerPath + objFile.Name;
					File.Copy(strPath, FilePath);

					int index = 1;
					while (index <= WhileCount)
					{
						try
						{
							System.Threading.Thread.Sleep(Sleepms);

							dt = getDataTableFromExcel(FilePath, ExcelSheetName);

							if (IsDeBugMode)
								sbMsg.AppendFormat("執行次數【{0}】", index);
							index = WhileCount + 1;
						}
						catch (Exception ex)
						{
							if (index == WhileCount)
								sbMsg.AppendFormat("{0}\n執行次數【{1}】已達上限！", ex.Message, index);

                            if (ex.Message == "檢查Sheet名稱是否不存在!" || ex.Message.Contains("檢查Excel欄位是否不存在!") == true)
							{
								index = WhileCount + 1;
								sbMsg.Append(ex.Message);
							}
						}
						index++;
					}

					System.IO.File.Delete(FilePath);
				}
			}
			catch (Exception ex)
			{
				sbMsg.Append(ex.Message);
			}
			finally
			{
				//釋放模擬權限
				MyEmuSPUserClass.UndoEmuSPUser();
			}

			return dt;
		}

		/// <summary>
		/// 讀2007↑Excel檔案到DataTable
		/// </summary>
		/// <param name="path">Excel路徑</param>
		/// <param name="startRowIndex">開始列的Index</param>
		/// <param name="hasHeader">是否依Excel的StartRow建立欄位</param>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public System.Data.DataTable getDataTableFromExcel(string path, int startRowIndex, bool hasHeader, string sheetName)
		{
			using (var pck = new OfficeOpenXml.ExcelPackage())
			{
				using (var stream = System.IO.File.OpenRead(path))
				{
					pck.Load(stream);
				}
				var ws = pck.Workbook.Worksheets[sheetName];
				return this.getDataTable(startRowIndex, hasHeader, ws);
			}
		}

		public System.Data.DataTable getDataTableFromExcel(string path, int startRowIndex, bool hasHeader, int sheetIndex)
		{
			using (var pck = new OfficeOpenXml.ExcelPackage())
			{
				using (var stream = System.IO.File.OpenRead(path))
				{
					pck.Load(stream);
				}
				var ws = pck.Workbook.Worksheets[sheetIndex];
				return this.getDataTable(startRowIndex, hasHeader, ws);
			}
		}

		public void getWorkSheetFromExcelTemplate(string path, int sheetIndex)
		{
			if (pck == null)
				pck = new OfficeOpenXml.ExcelPackage();

			using (var stream = System.IO.File.OpenRead(path))
			{
				pck.Load(stream);
			}
			workSheetAry[0] = pck.Workbook.Worksheets[sheetIndex];
		}

		private System.Data.DataTable getDataTable(int startRowIndex, bool hasHeader, OfficeOpenXml.ExcelWorksheet ws)
		{
			System.Data.DataTable tb = new System.Data.DataTable();
			if (ws != null)
			{
				foreach (var firstRowCell in ws.Cells[startRowIndex, 1, startRowIndex, ws.Dimension.End.Column])
				{
					tb.Columns.Add(hasHeader ? firstRowCell.Text.Replace(" ", "").Replace(".", "") : string.Format("Column {0}", firstRowCell.Start.Column));
				}
				var startRow = startRowIndex + (hasHeader ? 1 : 0);
				for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
				{
					var wsRow = ws.Cells[rowNum, 1, rowNum, tb.Columns.Count];
					var row = tb.NewRow();
					foreach (var cell in wsRow)
					{
						row[cell.Start.Column - 1] = cell.Text.Replace("'", "");
					}
					tb.Rows.Add(row);
				}
			}
			else
				throw new Exception("檢查Sheet名稱是否不存在!");
			return tb;
		}

		/// <summary>
		/// 建立一個新的Sheet
		/// </summary>
		/// <param name="sheetName"></param>
		public void SetExcelWorkSheet(string sheetName)
		{
			this.SetExcelWorkSheet(0, sheetName);
		}

		public void SetExcelWorkSheet(int sheetIndex, string sheetName)
		{
			if (pck == null)
				pck = new OfficeOpenXml.ExcelPackage();
			workSheetAry[sheetIndex] = pck.Workbook.Worksheets.Add(sheetName);
		}

		public void GetExcelFileInfo(string sheetName, string filePath)
		{
			/*
			if (filePath != "")
			{
				System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.ReadWrite);
				pck = new OfficeOpenXml.ExcelPackage(fs);
				workSheetAry[0] = pck.Workbook.Worksheets.SingleOrDefault(x => x.Name == sheetName);
			}
			*/
			
			if (filePath != "")
            {
                System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.ReadWrite);

                try
                {
                    pck = new OfficeOpenXml.ExcelPackage(fs);
                    workSheetAry[0] = pck.Workbook.Worksheets.SingleOrDefault(x => x.Name == sheetName);
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    fs.Close();
                }
            }			
		}


		public void GetExcelFileInfo(string sheetName)
		{
			if (pck != null)
			{
				workSheetAry[0] = pck.Workbook.Worksheets.SingleOrDefault(x => x.Name == sheetName);
			}
		}

		public object GetCellValue(
		  int fromRowIndex,
		  int fromColomnIndex,
		  int toRowIndex,
		  int toColumnIndex)
		{
			object obj = new object();
			obj = this.workSheetAry[0].Cells[fromRowIndex, fromColomnIndex, toRowIndex, toColumnIndex].Value;
			return obj;
		}

		public object GetCellValue(int rowIndex, int columnIndex)
		{
			return this.GetCellValue(rowIndex, columnIndex, rowIndex, columnIndex);
		}

		public object GetCellValue(string cellAddress)
		{
			object obj = new object();

			obj = this.workSheetAry[0].Cells[cellAddress].Value;
			return obj;
		}

		public void GetUnlockExcelFileInfo(System.Web.UI.WebControls.FileUpload FileUpload, string sheetName, ref StringBuilder sbMsg)
		{
			//string UnlockPath = @"\\192.168.1.222\erp\ERPAP\App_Data\Unlock\";
			EmuSPUserClass MyEmuSPUserClass = new EmuSPUserClass();
			try
			{
				int Sleepms = 500;
				int WhileCount = 5;
				if (MyEmuSPUserClass.EmuSPUser("CHPT"))
				{
					string ServerPath = UnlockPath;
					string FilePath = ServerPath + FileUpload.FileName;
					FileUpload.SaveAs(FilePath);

					int index = 1;
					while (index <= WhileCount)
					{
						System.Threading.Thread.Sleep(Sleepms);

						System.IO.FileStream fs = new System.IO.FileStream(FilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);

						try
						{
							pck = new OfficeOpenXml.ExcelPackage(fs);
							if (pck.Workbook.Worksheets.Count > 0)
								workSheetAry[0] = pck.Workbook.Worksheets.SingleOrDefault(x => x.Name == sheetName);

							index = WhileCount + 1;
						}
						catch (Exception ex)
						{
							if (index == WhileCount)
								sbMsg.AppendFormat("{0}\n執行次數【{1}】已達上限！", ex.Message, index);

                            if (ex.Message == "檢查Sheet名稱是否不存在!")
							{
								index = WhileCount + 1;
								sbMsg.Append(ex.Message);
							}
						}
						finally
						{
							fs.Close();
						}
						index++;
					}

					System.IO.File.Delete(FilePath);
				}
			}
			catch (Exception ex)
			{
				sbMsg.Append(ex.Message);
			}
			finally
			{
				//釋放模擬權限
				MyEmuSPUserClass.UndoEmuSPUser();
			}


		}

		public void GetExcelFileInfo(string sheetName, System.IO.FileStream fs)
		{
			pck = new OfficeOpenXml.ExcelPackage(fs);
			workSheetAry[0] = pck.Workbook.Worksheets.SingleOrDefault(x => x.Name == sheetName);
		}



		public void SetWorkSheetTabColor(System.Drawing.Color color)
		{
			this.SetWorkSheetTabColor(0, color);
		}

		public void SetWorkSheetTabColor(int sheetIndex, System.Drawing.Color color)
		{
			workSheetAry[sheetIndex].TabColor = color;
		}

		/// <summary>
		/// 設定凍結欄位
		/// 使用說明：
		/// 1.若需凍結第一列(Row)，則Row需設定2 Ex：(2, 1)
		/// 2.若需凍結第一欄(Column)，則Column需設定2 Ex：(1, 2)
		/// 3.若需凍結2列+5欄，則(3, 6)
		/// </summary>
		/// <param name="row"></param>
		/// <param name="column"></param>
		public void SetFreezePanes(int row, int column)
		{
			//workSheet.View.FreezePanes(row, column);
			this.SetFreezePanes(0, row, column);
		}

		public void SetFreezePanes(int sheetIndex, int row, int column)
		{
			workSheetAry[sheetIndex].View.FreezePanes(row, column);
		}

		public void AddHeaderColumnByUserDefin(string headerName, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string cellColorName, bool isBorder, bool isTextCenter)
		{
			this.AddHeaderColumnByUserDefin(0, headerName, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, cellColorName, isBorder, isTextCenter);
		}

		public void AddHeaderColumnByUserDefin(int sheetIndex, string headerName, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string cellColorName, bool isBorder, bool isTextCenter)
		{
			using (var header = workSheetAry[sheetIndex].Cells[fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex])
			{
				header.Value = headerName;
				header.Merge = true;

				//設定框線
				if (isBorder)
					this.setBorderStyle(header);
				//置中
				if (isTextCenter)
					this.setCellCenter(header);
				//背景顏色
				if (cellColorName != null && cellColorName != "")
				{
					System.Drawing.Color cellColor = System.Drawing.Color.FromName(cellColorName);
					header.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					header.Style.Fill.BackgroundColor.SetColor(cellColor);
				}
			}
		}

		public void SetCellValue(string cellValue, string cellAddress, string hexColor)
		{
			this.workSheetAry[0].Cells[cellAddress].Value = cellValue;
			if (!string.IsNullOrEmpty(hexColor))
			{
				System.Drawing.Color cellColor = System.Drawing.ColorTranslator.FromHtml(hexColor);
				this.workSheetAry[0].Cells[cellAddress].Style.Fill.BackgroundColor.SetColor(cellColor);
			}
		}

		public void SetCellValue(string cellValue, string cellAddress)
		{
			this.workSheetAry[0].Cells[cellAddress].Value = cellValue;
		}

		public void SetCellValue(string cellValue, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string cellColorName, bool isBorder, bool isTextCenter)
		{
			this.AddHeaderColumnByUserDefin(cellValue, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, cellColorName, isBorder, isTextCenter);
		}

		public void SetCellValue(int sheetIndex, string cellValue, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string cellColorName, bool isBorder, bool isTextCenter)
		{
			this.AddHeaderColumnByUserDefin(sheetIndex, cellValue, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, cellColorName, isBorder, isTextCenter);
		}

		public void SetCellComment(string cellComment, int rowIndex, int columnIndex)
		{
			this.SetCellComment(0, cellComment, rowIndex, columnIndex);
		}

		public void SetCellComment(int sheetIndex, string cellComment, int rowIndex, int columnIndex)
		{
			using (var cell = workSheetAry[sheetIndex].Cells[rowIndex, columnIndex, rowIndex, columnIndex])
			{
				cell.AddComment(cellComment, "Tab");
				cell.Comment.AutoFit = true;
			}
		}

		/// <summary>
		/// 設定DataTable Header的顏色
		/// </summary>
		/// <param name="color"></param>
		public void SetDataTableHeaderColor(System.Drawing.Color color)
		{
			dtHeaderColor = color;
		}

		/// <summary>
		/// 載入DataTable的值
		/// </summary>
		/// <param name="startRowIndex"></param>
		/// <param name="dtSource"></param>
		/// <param name="isAddHeaderColumn"></param>
		/// <param name="isBorder"></param>
		/// <param name="isTextCenter"></param>
		public void AddContentData(int startRowIndex, System.Data.DataTable dtSource, bool isAddHeaderColumn, bool isBorder, bool isTextCenter)
		{
			this.AddContentData(startRowIndex, dtSource, isAddHeaderColumn, isBorder, isTextCenter, "", "", "");
		}

		public void AddContentData(int sheetIndex, int startRowIndex, System.Data.DataTable dtSource, bool isAddHeaderColumn, bool isBorder, bool isTextCenter)
		{
			this.AddContentData(sheetIndex, startRowIndex, dtSource, isAddHeaderColumn, isBorder, isTextCenter, "", "", "");
		}

		public void AddContentData(int startRowIndex, System.Data.DataTable dtSource, bool isAddHeaderColumn, bool isBorder, bool isTextCenter, string highlightDifferenceFontColorName, string highlightDifferenceBackgroundColorName, string otherBackgroundColorName)
		{
			this.AddContentData(0, startRowIndex, dtSource, isAddHeaderColumn, isBorder, isTextCenter, highlightDifferenceFontColorName, highlightDifferenceBackgroundColorName, otherBackgroundColorName);
		}

		public void AddContentData(int sheetIndex, int startRowIndex, System.Data.DataTable dtSource, bool isAddHeaderColumn, bool isBorder, bool isTextCenter, string highlightDifferenceFontColorName, string highlightDifferenceBackgroundColorName, string otherBackgroundColorName)
		{
			if (dtSource != null && dtSource.Rows.Count > 0)
			{
				#region 設定欄位名稱

				if (isAddHeaderColumn)
				{
					int startColumnIndex = 1;
					foreach (System.Data.DataColumn dc in dtSource.Columns)
					{
						using (var header = workSheetAry[sheetIndex].Cells[startRowIndex, startColumnIndex++])
						{
							header.Value = dc.ColumnName;
							//設定框線
							if (isBorder)
								this.setBorderStyle(header);
							//置中
							if (isTextCenter)
								this.setCellCenter(header);

							header.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
							header.Style.Fill.BackgroundColor.SetColor(dtHeaderColor);
						}
					}
					//換下一行
					startRowIndex++;
				}

				#endregion 設定欄位名稱

				#region 設定DataTable的內容

				bool isFontColor, isBackgroundColor;
				foreach (System.Data.DataRow dr in dtSource.Rows)
				{
					//每次將啟始欄位歸零
					int startColumnIndex = 1;
					foreach (object obj in dr.ItemArray)
					{
						using (var cellContent = workSheetAry[sheetIndex].Cells[startRowIndex, startColumnIndex])
						{
							if (otherBackgroundColorName != "")
							{
								cellContent.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
								cellContent.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromName(otherBackgroundColorName));
							}

							/*過濾HTML格式*/
							System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(obj.ToString(), "<[^>]*>");
							if (match.Success)
							{
								cellContent.Value = System.Text.RegularExpressions.Regex.Replace(obj.ToString(), "<[^>]*>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
							}
							else
								cellContent.Value = obj;

							#region 換行

							if (cellContent.Value.ToString().IndexOf("\n") > -1)
								cellContent.Style.WrapText = true;
							else if (cellContent.Value.ToString().IndexOf("\\n") > -1)
							{
								cellContent.Value = cellContent.Value.ToString().Replace("\\n", "\n");
								cellContent.Style.WrapText = true;
							}

							#endregion 換行

							isFontColor = highlightDifferenceFontColorName != "";
							isBackgroundColor = highlightDifferenceBackgroundColorName != "";
							if (isFontColor || isBackgroundColor)
							{
								if (obj.ToString().IndexOf("→") > 0)
								{
									if (isFontColor)
										cellContent.Style.Font.Color.SetColor(System.Drawing.Color.FromName(highlightDifferenceFontColorName));

									if (isBackgroundColor)
									{
										cellContent.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
										cellContent.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromName(highlightDifferenceBackgroundColorName));
									}
								}
							}

							//設定框線
							if (isBorder)
								this.setBorderStyle(cellContent);
							//置中
							if (isTextCenter)
								this.setCellCenter(cellContent);
						}
						startColumnIndex++;
					}
					//換下一行
					startRowIndex++;
				}

				#endregion 設定DataTable的內容
			}
		}

		/// <summary>
		/// 水平對齊位置
		/// </summary>
		/// <param name="fromRowIndex"></param>
		/// <param name="fromColumnIndex"></param>
		/// <param name="toRowIndex"></param>
		/// <param name="toColumnIndex"></param>
		/// <param name="horizontalAlignment"></param>
		public void SetHorizontalAlignment(int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, OfficeOpenXml.Style.ExcelHorizontalAlignment horizontalAlignment)
		{
			this.SetHorizontalAlignment(0, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, horizontalAlignment);
		}

		public void SetHorizontalAlignment(int sheetIndex, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, OfficeOpenXml.Style.ExcelHorizontalAlignment horizontalAlignment)
		{
			using (var range = workSheetAry[sheetIndex].Cells[fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex])
			{
				range.Style.HorizontalAlignment = horizontalAlignment;
			}
		}

		/// <summary>
		/// 垂直對齊位置
		/// </summary>
		/// <param name="sheetIndex"></param>
		/// <param name="fromRowIndex"></param>
		/// <param name="fromColumnIndex"></param>
		/// <param name="toRowIndex"></param>
		/// <param name="toColumnIndex"></param>
		/// <param name="verticalAlignment"></param>
		public void SetVerticalAlignment(int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, OfficeOpenXml.Style.ExcelVerticalAlignment verticalAlignment)
		{
			this.SetVerticalAlignment(0, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, verticalAlignment);
		}

		public void SetVerticalAlignment(int sheetIndex, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, OfficeOpenXml.Style.ExcelVerticalAlignment verticalAlignment)
		{
			using (var range = workSheetAry[sheetIndex].Cells[fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex])
			{
				range.Style.VerticalAlignment = verticalAlignment;
			}
		}

		public enum NumberFormatType
		{
			Int,
			Double
		}

		/// <summary>
		/// 設定欄位為數字型態
		/// </summary>
		/// <param name="fromRowIndex"></param>
		/// <param name="fromColumnIndex"></param>
		/// <param name="toRowIndex"></param>
		/// <param name="toColumnIndex"></param>
		/// <param name="numType">Int或Double</param>
		/// <param name="cellFormat">定義欄位格式</param>
		/// <param name="colorName"></param>
		public void SetNumberFormat(int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, NumberFormatType numType, string cellFormat, string colorName)
		{
			this.SetNumberFormat(0, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, numType, cellFormat, colorName);
		}

		public void SetNumberFormat(int sheetIndex, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, NumberFormatType numType, string cellFormat, string colorName)
		{
			for (int i = fromRowIndex; i <= toRowIndex; i++)
			{
				for (int j = fromColumnIndex; j <= toColumnIndex; j++)
				{
					using (var cell = workSheetAry[sheetIndex].Cells[i, j, i, j])
					{
						if (cell == null || cell.Value.ToString() == "")
						{
							cell.Value = 0;
						}
						else
						{
							if (numType == NumberFormatType.Int)
							{
								int intValue;
								cell.Value = int.TryParse(cell.Value.ToString(), out intValue) ? intValue : 0;
							}
							else if (numType == NumberFormatType.Double)
							{
								double dbValue;
								cell.Value = double.TryParse(cell.Value.ToString(), out dbValue) ? dbValue : 0;
							}
						}
					}
				}
			}

			using (var range = workSheetAry[sheetIndex].Cells[fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex])
			{
				range.Style.Numberformat.Format = cellFormat;
				if (colorName != null && colorName != "")
				{
					System.Drawing.Color cellColor = System.Drawing.Color.FromName(colorName);
					range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					range.Style.Fill.BackgroundColor.SetColor(cellColor);
				}
			}
		}

        public void SetDateFormat(int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string cellFormat, string colorName)
        {
            this.SetDateFormat(0, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, cellFormat, colorName);
        }

        public void SetDateFormat(int sheetIndex, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string cellFormat, string colorName)
        {
            for (int i = fromRowIndex; i <= toRowIndex; i++)
            {
                for (int j = fromColumnIndex; j <= toColumnIndex; j++)
                {
                    using (var cell = workSheetAry[sheetIndex].Cells[i, j, i, j])
                    {
                        DateTime dtime;
                        cell.Value = DateTime.TryParse(cell.Value.ToString(), out dtime) ? dtime.ToString(cellFormat) : "";
                    }
                }
            }

            using (var range = workSheetAry[sheetIndex].Cells[fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex])
            {
                if (colorName != null && colorName != "")
                {
                    System.Drawing.Color cellColor = System.Drawing.Color.FromName(colorName);
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(cellColor);
                }
            }
        }

		/// <summary>
		/// 設定AutoFit
		/// </summary>
		/// <param name="fromColumnIndex"></param>
		/// <param name="toColumnIndex"></param>
		public void SetAutoFit(int fromColumnIndex, int toColumnIndex)
		{
			this.SetAutoFit(0, fromColumnIndex, toColumnIndex);
		}

		public void SetAutoFit(int sheetIndex, int fromColumnIndex, int toColumnIndex)
		{
			try
			{
				for (int i = fromColumnIndex; i <= toColumnIndex; i++)
				{
					workSheetAry[sheetIndex].Column(i).AutoFit();
				}
			}
			catch (Exception ex)
			{
				string Msg = ex.Message;
			}
		}

		public void SetAutoFit(int fromColumnIndex, int toColumnIndex, double minWidth, double maxWidth)
		{
			this.SetAutoFit(0, fromColumnIndex, toColumnIndex, minWidth, maxWidth);
		}

		public void SetAutoFit(int sheetIndex, int fromColumnIndex, int toColumnIndex, double minWidth, double maxWidth)
		{
			for (int i = fromColumnIndex; i <= toColumnIndex; i++)
			{
				workSheetAry[sheetIndex].Column(i).AutoFit(minWidth, maxWidth);
			}
		}

		/// <summary>
		/// 設定自動換列
		/// </summary>
		/// <param name="fromRowIndex"></param>
		/// <param name="fromColumnIndex"></param>
		/// <param name="toRowIndex"></param>
		/// <param name="toColumnIndex"></param>
		public void SetContentAutoRows(int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex)
		{
			this.SetContentAutoRows(0, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex);
		}

		public void SetContentAutoRows(int sheetIndex, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex)
		{
			using (var range = workSheetAry[sheetIndex].Cells[fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex])
			{
				range.Style.WrapText = true;
			}
		}

		/// <summary>
		/// 設定版面配置 直向或橫向
		/// </summary>
		/// <param name="orientation">直向(Portrait)、橫向(Landscape)</param>
		public void SetOrientation(OfficeOpenXml.eOrientation orientation)
		{
			this.SetOrientation(0, orientation);
		}

		public void SetOrientation(int sheetIndex, OfficeOpenXml.eOrientation orientation)
		{
			workSheetAry[sheetIndex].PrinterSettings.Orientation = orientation;
		}

		/// <summary>
		/// 設定列印紙張大小
		/// </summary>
		/// <param name="paperSize">紙張大小</param>
		/// <param name="isBlackAndWhite">是否設定為黑白列印</param>
		public void SetPrintPaperSize(OfficeOpenXml.ePaperSize paperSize, bool isBlackAndWhite)
		{
			this.SetPrintPaperSize(0, paperSize, isBlackAndWhite);
		}

		public void SetPrintPaperSize(int sheetIndex, OfficeOpenXml.ePaperSize paperSize, bool isBlackAndWhite)
		{
			workSheetAry[sheetIndex].PrinterSettings.PaperSize = paperSize;
			workSheetAry[sheetIndex].PrinterSettings.BlackAndWhite = isBlackAndWhite;
		}

		/// <summary>
		/// 設定列印符合一頁寬
		/// </summary>
		public void SetPrintFitToOnePage()
		{
			this.SetPrintFitToOnePage(0);
		}

		public void SetPrintFitToOnePage(int sheetIndex)
		{
			this.SetPrintFitToOnePage(sheetIndex, 0, 1);
		}

		public void SetPrintFitToOnePage(int FitToHeight, int FitToWidth)
		{
			this.SetPrintFitToOnePage(0, FitToHeight, FitToWidth);
		}

		public void SetPrintFitToOnePage(int sheetIndex, int FitToHeight, int FitToWidth)
		{
			workSheetAry[sheetIndex].PrinterSettings.FitToPage = true;
			workSheetAry[sheetIndex].PrinterSettings.FitToWidth = FitToWidth;
			workSheetAry[sheetIndex].PrinterSettings.FitToHeight = FitToHeight;
		}

		/// <summary>
		/// 設定列印邊界
		/// </summary>
		/// <param name="HeaderMargin">頁首</param>
		/// <param name="FooterMargin">頁尾</param>
		/// <param name="TopMargin">上</param>
		/// <param name="BottomMargin">下</param>
		/// <param name="LeftMargin">左</param>
		/// <param name="RightMargin">右</param>
		public void SetPrintMargin(decimal HeaderMargin, decimal FooterMargin, decimal TopMargin, decimal BottomMargin, decimal LeftMargin, decimal RightMargin)
		{
			this.SetPrintMargin(0, HeaderMargin, FooterMargin, TopMargin, BottomMargin, LeftMargin, RightMargin);
		}

		public void SetPrintMargin(int sheetIndex, decimal HeaderMargin, decimal FooterMargin, decimal TopMargin, decimal BottomMargin, decimal LeftMargin, decimal RightMargin)
		{
			workSheetAry[sheetIndex].PrinterSettings.HeaderMargin = HeaderMargin;
			workSheetAry[sheetIndex].PrinterSettings.FooterMargin = FooterMargin;

			workSheetAry[sheetIndex].PrinterSettings.TopMargin = TopMargin;
			workSheetAry[sheetIndex].PrinterSettings.BottomMargin = BottomMargin;
			workSheetAry[sheetIndex].PrinterSettings.LeftMargin = LeftMargin;
			workSheetAry[sheetIndex].PrinterSettings.RightMargin = RightMargin;

			//workSheetAry[sheetIndex].View.ZoomScale = 70;
		}

		public enum EPPlusHorizontalAlign
		{
			Left,
			Center,
			Right
		}

		/// <summary>
		/// 設定頁碼
		/// </summary>
		/// <param name="hAlign">顯示在Footer那個位置</param>
		public void SetFooterPageNumber(EPPlusHorizontalAlign hAlign)
		{
			this.SetFooterPageNumber(0, hAlign);
		}

		public void SetFooterPageNumber(int sheetIndex, EPPlusHorizontalAlign hAlign)
		{
			string pageNumber = string.Format("第 {0} 頁，共 {1} 頁", OfficeOpenXml.ExcelHeaderFooter.PageNumber, OfficeOpenXml.ExcelHeaderFooter.NumberOfPages);
			switch (hAlign)
			{
				case EPPlusHorizontalAlign.Left:
					workSheetAry[sheetIndex].HeaderFooter.OddFooter.LeftAlignedText = pageNumber;
					break;

				case EPPlusHorizontalAlign.Center:
					workSheetAry[sheetIndex].HeaderFooter.OddFooter.CenteredText = pageNumber;
					break;

				case EPPlusHorizontalAlign.Right:
					workSheetAry[sheetIndex].HeaderFooter.OddFooter.RightAlignedText = pageNumber;
					break;
			}
		}

		/// <summary>
		/// 設定Footer顯示的值
		/// </summary>
		/// <param name="hAlign">顯示在Footer那個位置</param>
		/// <param name="value">顯示的值</param>
		public void SetFooterValue(EPPlusHorizontalAlign hAlign, string value)
		{
			this.SetFooterValue(0, hAlign, value);
		}

		public void SetFooterValue(int sheetIndex, EPPlusHorizontalAlign hAlign, string value)
		{
			switch (hAlign)
			{
				case EPPlusHorizontalAlign.Left:
					workSheetAry[sheetIndex].HeaderFooter.OddFooter.LeftAlignedText = value;
					break;

				case EPPlusHorizontalAlign.Center:
					workSheetAry[sheetIndex].HeaderFooter.OddFooter.CenteredText = value;
					break;

				case EPPlusHorizontalAlign.Right:
					workSheetAry[sheetIndex].HeaderFooter.OddFooter.RightAlignedText = value;
					break;
			}
		}

		public void SetFontStyle(int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, float fontSize, bool fontBold, string fontColor)
		{
			this.SetFontStyle(0, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, fontSize, fontBold, fontColor);
		}

		public void SetFontStyle(int sheetIndex, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, float fontSize, bool fontBold, string fontColor)
		{
			using (var range = workSheetAry[sheetIndex].Cells[fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex])
			{
				range.Style.Font.Size = fontSize;
				range.Style.Font.Bold = fontBold;

				//背景顏色
				if (fontColor != null && fontColor != "")
				{
					System.Drawing.Color cellColor = System.Drawing.Color.FromName(fontColor);
					range.Style.Font.Color.SetColor(cellColor);
				}
			}
		}

		public void SetRowHeight(int rowIndex, int height)
		{
			this.SetRowHeight(0, rowIndex, height);
		}

		public void SetRowHeight(int sheetIndex, int rowIndex, int height)
		{
			workSheetAry[sheetIndex].Row(rowIndex).Height = height;
		}

		/// <summary>
		/// 設定重覆列印的Row
		/// </summary>
		/// <param name="fromRowIndex"></param>
		/// <param name="toRowIndex"></param>
		public void SetPrintRepeatRows(int fromRowIndex, int toRowIndex)
		{
			this.SetPrintRepeatRows(0, fromRowIndex, toRowIndex);
		}

		/// <summary>
		/// 設定重覆列印的Row
		/// </summary>
		/// <param name="sheetIndex"></param>
		/// <param name="fromRowIndex"></param>
		/// <param name="toRowIndex"></param>
		public void SetPrintRepeatRows(int sheetIndex, int fromRowIndex, int toRowIndex)
		{
			workSheetAry[sheetIndex].PrinterSettings.RepeatRows = new OfficeOpenXml.ExcelAddress(String.Format("{0}:{1}", fromRowIndex, toRowIndex));
		}

		/// <summary>
		/// 匯出至Excel
		/// </summary>
		/// <param name="page"></param>
		/// <param name="fileName"></param>
		/// <param name="sbErrMsg"></param>
		/// <returns></returns>
		public bool ExportToExcel(System.Web.UI.Page page, string fileName, out System.Text.StringBuilder sbErrMsg)
		{
			sbErrMsg = new System.Text.StringBuilder();

			bool isSuccess = (pck != null && this.CheckWorkSheet());
			if (isSuccess)
			{
				isSuccess = false;
				this.ResponseToExcel(page, pck, fileName);
				isSuccess = true;
			}
			else
			{
				sbErrMsg.Append("請先建立WorkSheet，再執行匯出動作!\\n");
			}
			return isSuccess;
		}

		public bool ExportToExcel(System.Web.HttpContext context, string fileName, out System.Text.StringBuilder sbErrMsg)
		{
			sbErrMsg = new System.Text.StringBuilder();

			bool isSuccess = (pck != null && this.CheckWorkSheet());
			if (isSuccess)
			{
				isSuccess = false;
				this.ResponseToExcel(context, pck, fileName);
				isSuccess = true;
			}
			else
			{
				sbErrMsg.Append("請先建立WorkSheet，再執行匯出動作!\\n");
			}
			return isSuccess;
		}

        public bool ExportToExcelBatch(System.Web.UI.Page page, string fileName, string path, out System.Text.StringBuilder sbErrMsg)
        {
            sbErrMsg = new System.Text.StringBuilder();

            bool isSuccess = (pck != null && this.CheckWorkSheet());
            if (isSuccess)
            {
                isSuccess = false;
                this.ResponseToExcelBatch(page, pck, fileName, path);
                isSuccess = true;
            }
            else
            {
                sbErrMsg.Append("請先建立WorkSheet，再執行匯出動作!\\n");
            }
            return isSuccess;
        }

        private bool CheckWorkSheet()
		{
			bool isOK = true;
			for (int i = 0; i < workSheetAry.Length; i++)
			{
				if (workSheetAry[i] == null)
				{
					isOK = false;
					break;
				}
			}
			return isOK;
		}

		/// <summary>
		/// 儲存Excel至指定路徑
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="sbErrMsg"></param>
		/// <returns></returns>
		public bool ExportToSaveFile(string filePath, out System.Text.StringBuilder sbErrMsg)
		{
			sbErrMsg = new System.Text.StringBuilder();

			bool isSuccess = (pck != null && this.CheckWorkSheet());
			if (isSuccess)
			{
				isSuccess = false;
				System.IO.File.WriteAllBytes(filePath, pck.GetAsByteArray());
				isSuccess = true;
			}
			else
			{
				sbErrMsg.Append("請先建立WorkSheet，再執行匯出動作!\\n");
			}
			return isSuccess;
		}

		/// <summary>
		/// 設定Total的公式
		/// </summary>
		/// <param name="ws"></param>
		/// <param name="startRow"></param>
		/// <param name="endRow"></param>
		/// <param name="startColumn"></param>
		/// <param name="endColumn"></param>
		public void SetTotalFormula(int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string cellFormat, string colorName, bool isBorder, bool isTextCenter)
		{
			this.SetTotalFormula(0, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, cellFormat, colorName, isBorder, isTextCenter);
		}

		public void SetTotalFormula(int sheetIndex, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string cellFormat, string colorName, bool isBorder, bool isTextCenter)
		{
			using (var formula = workSheetAry[sheetIndex].Cells[toRowIndex + 1, fromColumnIndex, toRowIndex + 1, toColumnIndex])
			{
				formula.Formula = string.Format("Sum({0})", new OfficeOpenXml.ExcelAddress(fromRowIndex, fromColumnIndex, toRowIndex, fromColumnIndex).Address);

				if (cellFormat != "")
					formula.Style.Numberformat.Format = cellFormat;

				//設定框線
				if (isBorder)
					this.setBorderStyle(formula);
				//置中
				if (isTextCenter)
					this.setCellCenter(formula);
				//背景顏色
				if (colorName != null && colorName != "")
				{
					System.Drawing.Color cellColor = System.Drawing.Color.FromName(colorName);
					formula.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					formula.Style.Fill.BackgroundColor.SetColor(cellColor);
				}
			}
		}

		/// <summary>
		/// 數字欄位 使用者可自訂公式
		/// </summary>
		/// <param name="fromRowIndex"></param>
		/// <param name="fromColumnIndex"></param>
		/// <param name="toRowIndex"></param>
		/// <param name="toColumnIndex"></param>
		/// <param name="formulaString"></param>
		/// <param name="cellFormat"></param>
		/// <param name="colorName"></param>
		/// <param name="isBorder"></param>
		/// <param name="isTextCenter"></param>
		public void SetUserDefinFormula(int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string formulaString, string cellFormat, string colorName, bool isBorder, bool isTextCenter)
		{
			this.SetUserDefinFormula(0, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, formulaString, cellFormat, colorName, isBorder, isTextCenter);
		}

		public void SetUserDefinFormula(int sheetIndex, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, string formulaString, string cellFormat, string colorName, bool isBorder, bool isTextCenter)
		{
			using (var formula = workSheetAry[sheetIndex].Cells[fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex])
			{
				formula.Formula = formulaString;
				formula.Style.Numberformat.Format = cellFormat;

				//設定框線
				if (isBorder)
					this.setBorderStyle(formula);
				//置中
				if (isTextCenter)
					this.setCellCenter(formula);
				//背景顏色
				if (colorName != null && colorName != "")
				{
					System.Drawing.Color cellColor = System.Drawing.Color.FromName(colorName);
					formula.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					formula.Style.Fill.BackgroundColor.SetColor(cellColor);
				}
			}
		}

		/// <summary>
		/// 設定Cell背景
		/// </summary>
		/// <param name="fromRowIndex"></param>
		/// <param name="fromColumnIndex"></param>
		/// <param name="toRowIndex"></param>
		/// <param name="toColumnIndex"></param>
		/// <param name="color"></param>
		public void SetBackgroundColor(int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, System.Drawing.Color color)
		{
			this.SetBackgroundColor(0, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex, color);
		}

		public void SetBackgroundColor(int sheetIndex, int fromRowIndex, int fromColumnIndex, int toRowIndex, int toColumnIndex, System.Drawing.Color color)
		{
			using (var range = workSheetAry[sheetIndex].Cells[fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex])
			{
				//背景顏色
				range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				range.Style.Fill.BackgroundColor.SetColor(color);
			}
		}

		/// <summary>
		/// 垂直、水平置中
		/// </summary>
		/// <param name="range"></param>
		private void setCellCenter(OfficeOpenXml.ExcelRange range)
		{
			//垂直位置
			range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
			//水平位置
			range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
		}

		/// <summary>
		/// Response Excel檔案
		/// </summary>
		/// <param name="page"></param>
		/// <param name="pck"></param>
		/// <param name="ExcelFileName"></param>
		private void ResponseToExcel(System.Web.UI.Page page, OfficeOpenXml.ExcelPackage pck, string ExcelFileName)
		{
			System.Web.HttpBrowserCapabilities browser = page.Request.Browser;
			string urlEncodeFileName;
			if (browser.Browser.ToUpper() != "IE")
				urlEncodeFileName = string.Format("{0}.xlsx", ExcelFileName);
			else
				urlEncodeFileName = System.Web.HttpUtility.UrlEncode(string.Format("{0}.xlsx", ExcelFileName), Encoding.UTF8);

			//匯出Excel
			page.Response.Clear();
			page.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", urlEncodeFileName));
			page.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			page.Response.BinaryWrite(pck.GetAsByteArray());
			page.Response.End();
		}

		private void ResponseToExcel(System.Web.HttpContext context, OfficeOpenXml.ExcelPackage pck, string ExcelFileName)
		{
			string urlEncodeFileName = string.Format("{0}.xlsx", ExcelFileName);
			//匯出Excel
			context.Response.Clear();
			context.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", urlEncodeFileName));
			context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			context.Response.BinaryWrite(pck.GetAsByteArray());
			context.Response.End();
		}

        private void ResponseToExcelBatch(System.Web.UI.Page page, OfficeOpenXml.ExcelPackage pck, string ExcelFileName, string path)
        {
            System.Web.HttpBrowserCapabilities browser = page.Request.Browser;
            string urlEncodeFileName;
            if (browser.Browser.ToUpper() != "IE")
                urlEncodeFileName = string.Format("{0}.xlsx", ExcelFileName);
            else
                urlEncodeFileName = System.Web.HttpUtility.UrlEncode(string.Format("{0}.xlsx", ExcelFileName), Encoding.UTF8);


            path = path + ExcelFileName + ".xlsx";
            Stream stream = File.Create(path);
            pck.SaveAs(stream);
            stream.Close();



            //////匯出Excel
            //page.Response.Clear();
            //page.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", urlEncodeFileName));
            //page.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //page.Response.BinaryWrite(pck.GetAsByteArray());
            //page.Response.Flush();


        }

        /// <summary>
        /// 設定Cell框的格線
        /// </summary>
        /// <param name="rangeStyle"></param>
        private void setBorderStyle(OfficeOpenXml.ExcelRange rangeStyle)
		{
			//設定格線
			rangeStyle.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
			rangeStyle.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
			rangeStyle.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
			rangeStyle.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
		}

		public System.Data.DataTable Compare2DataTable(System.Data.DataTable dtMain, System.Data.DataTable dtSecond, string keyColumn, string orderByColumn, ref System.Text.StringBuilder sbMsg)
		{
			return this.Compare2DataTable(dtMain, dtSecond, keyColumn, orderByColumn, null, "", ref sbMsg);
		}

		public System.Data.DataTable Compare2DataTable(System.Data.DataTable dtMain, System.Data.DataTable dtSecond, string keyColumn, string orderByColumn, string[] CheckWholeColumnAry, string SplitSymbol, ref System.Text.StringBuilder sbMsg)
		{
			System.Data.DataTable dtDifference = new System.Data.DataTable();
			if (dtSecond.Rows.Count > 0)
			{
				dtDifference.Columns.Add("DefineDifference");

				#region 比對2個Table欄位名稱是否相符

				foreach (System.Data.DataColumn dcM in dtMain.Columns)
				{
					dtDifference.Columns.Add(dcM.ColumnName);

					bool isCheckOK = false;
					foreach (System.Data.DataColumn dcS in dtSecond.Columns)
					{
						if (dcM.ColumnName == dcS.ColumnName)
						{
							isCheckOK = true;
							break;
						}
					}
					if (!isCheckOK)
						sbMsg.AppendFormat("需比對的資料(Second Table)無【{0}】欄位，無法進行比對!\\n", dcM.ColumnName);
				}

				#endregion 比對2個Table欄位名稱是否相符

				if (sbMsg.Length == 0)
				{
					try
					{
						#region 刪除

						System.Collections.ArrayList alMainNo = this.getKeyColumnValue(dtMain, keyColumn);
						var linqDelete = from l in dtSecond.AsEnumerable()
										 where l[keyColumn].ToString().ToUpper() != "" && !alMainNo.Contains(l[keyColumn].ToString().ToUpper())
										 orderby l[orderByColumn]
										 select l;
						System.Data.DataTable dtDelete = new DataTable();
						if (linqDelete.Any())
							dtDelete = linqDelete.CopyToDataTable();

						#endregion 刪除

						#region 新增

						System.Collections.ArrayList alSecond = this.getKeyColumnValue(dtSecond, keyColumn);
						System.Data.DataTable dtAddNew = new DataTable();
						var linqAddNew = from l in dtMain.AsEnumerable()
										 where l[keyColumn].ToString().ToUpper() != "" && !alSecond.Contains(l[keyColumn].ToString().ToUpper())
										 orderby l[orderByColumn]
										 select l;
						if (linqAddNew.Any())
							dtAddNew = linqAddNew.CopyToDataTable();

						#endregion 新增

						#region 差異

						var linqTempSecond = from l in dtSecond.AsEnumerable()
											 where l[keyColumn].ToString().ToUpper() != "" && alMainNo.Contains(l[keyColumn].ToString().ToUpper())
											 orderby l[orderByColumn]
											 select l;
						System.Data.DataTable dtTempSecond = new DataTable();
						if (linqTempSecond.Any())
							dtTempSecond = linqTempSecond.CopyToDataTable();

						var linqTempMain = from l in dtMain.AsEnumerable()
										   where l[keyColumn].ToString().ToUpper() != "" && alSecond.Contains(l[keyColumn].ToString().ToUpper())
										   orderby l[orderByColumn]
										   select l;
						System.Data.DataTable dtTempMain = new DataTable();
						if (linqTempMain.Any())
							dtTempMain = linqTempMain.CopyToDataTable();

						char charSymbol = ' ';
						if (SplitSymbol != "")
							charSymbol = SplitSymbol.ToCharArray()[0];

						string mStr = "", sStr = "";
						System.Data.DataTable dtModify = dtDifference.Clone();
						foreach (System.Data.DataRow drM in dtTempMain.Rows)
						{
							System.Data.DataRow drSelected = null;
							foreach (System.Data.DataRow drS in dtTempSecond.Rows)
							{
								if (drM[keyColumn].ToString() == drS[keyColumn].ToString())
								{
									drSelected = drS;

									bool isAdd = false;
									System.Data.DataRow drNew = dtModify.NewRow();
									foreach (System.Data.DataColumn dcM in dtTempMain.Columns)
									{
										mStr = drM[dcM.ColumnName].ToString().Trim();
										sStr = drS[dcM.ColumnName].ToString().Trim();
										if (mStr.ToUpper() == sStr.ToUpper())
											drNew[dcM.ColumnName] = mStr;
										else
										{
											string value;
											if (SplitSymbol != "" && CheckWholeColumnAry != null && CheckWholeColumnAry.Contains(dcM.ColumnName))
											{
												//完整比較 不管順序只要有出現就算符合
												if (this.CheckWholeAllString(mStr, sStr, charSymbol, out value))
													drNew[dcM.ColumnName] = mStr;
												else
													drNew[dcM.ColumnName] = string.Format("【{0}】→【{1}】" + System.Environment.NewLine + "{2}", sStr, mStr, value);
											}
											else
												drNew[dcM.ColumnName] = string.Format("【{0}】→【{1}】", sStr, mStr);
											isAdd = true;
										}
									}
									if (isAdd)
										dtModify.Rows.Add(drNew);
									break;
								}
							}
							if (drSelected != null)
								dtTempSecond.Rows.Remove(drSelected);
						}

						#endregion 差異

						this.CombineDataTable(dtAddNew, "新增", ref dtDifference);
						this.CombineDataTable(dtModify, "修改", ref dtDifference);
						this.CombineDataTable(dtDelete, "刪除", ref dtDifference);
					}
					catch (Exception ex)
					{
						sbMsg.Append(ex.Message);
					}
				}

				if (sbMsg.Length > 0)
					dtDifference.Columns.Clear();
			}
			return dtDifference;
		}

		private bool CheckWholeAllString(string str1, string str2, char splitSymbol, out string outMsg)
		{
			bool isCheckOK = false;
			System.Text.StringBuilder sbComment = new System.Text.StringBuilder();

			if (splitSymbol != null)
			{
				System.Collections.ArrayList alBefore = new System.Collections.ArrayList(str1.Split(splitSymbol));
				System.Collections.ArrayList alAfter = new System.Collections.ArrayList(str2.Split(splitSymbol));

				var onlyInBefore = from string l in alBefore
								   where !alAfter.Contains(l)
								   select l;
				var onlyInAfter = from string l in alAfter
								  where !alBefore.Contains(l)
								  select l;

				if (onlyInBefore.Count() > 0)
				{
					sbComment.Append("新增：" + System.Environment.NewLine);
					int cycle = 1;
					foreach (string s in onlyInBefore)
					{
						if (cycle > 20)
						{
							cycle = 1;
							sbComment.Append(System.Environment.NewLine);
						}
						sbComment.AppendFormat("{0},", s);
						cycle++;
					}
				}
				if (sbComment.Length > 0)
				{
					sbComment.Remove(sbComment.Length - 1, 1);
					sbComment.Append(System.Environment.NewLine);
				}

				if (onlyInAfter.Count() > 0)
				{
					sbComment.Append("刪除：" + System.Environment.NewLine);
					int cycle = 1;
					foreach (string s in onlyInAfter)
					{
						if (cycle > 20)
						{
							cycle = 1;
							sbComment.Append(System.Environment.NewLine);
						}
						sbComment.AppendFormat("{0},", s);
						cycle++;
					}
				}
				if (sbComment.Length > 0)
					sbComment.Remove(sbComment.Length - 1, 1);
			}

			outMsg = "";
			if (sbComment.Length > 0)
				outMsg = sbComment.ToString();
			else
				isCheckOK = true;

			return isCheckOK;
		}

		private void CombineDataTable(System.Data.DataTable dtDifference, string differenceString, ref System.Data.DataTable dtCombine)
		{
			foreach (System.Data.DataRow dr in dtDifference.Rows)
			{
				System.Data.DataRow drNew = dtCombine.NewRow();
				foreach (System.Data.DataColumn dc in dtCombine.Columns)
					drNew[dc.ColumnName] = dc.ColumnName == "DefineDifference" ? differenceString : dr[dc.ColumnName];
				dtCombine.Rows.Add(drNew);
			}
		}

		private System.Collections.ArrayList getKeyColumnValue(System.Data.DataTable dt, string keyColumn)
		{
			System.Collections.ArrayList alNo = new System.Collections.ArrayList();
			foreach (System.Data.DataRow dr in dt.Rows)
			{
				if (dr[keyColumn] != null && dr[keyColumn].ToString() != "")
					alNo.Add(dr[keyColumn].ToString());
			}
			return alNo;
		}

        public string GetUnlockPath()
        {
            string path = "";
            bool isTestMode = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["TestMode"]);

            if (isTestMode)
            {
                path = @"\\weberp\C$\inetpub\wwwroot\UploadFiles\";
            }
            else
            {
                // \\FS1\erp\ERPAP\UploadFiles\
                string strSQL = @"
SELECT TOP 1 ParameterValue FROM AD0EM WHERE ParameterName = 'FileServerPath(ERP)'
";
				ERP.DBDao dao = new DBDao();

                path = dao.SqlSelectToString(strSQL, "ParameterValue");
            }

            // \\192.168.1.222\erp\ERPAP\App_Data\Unlock\
            return path+@"Unlock\";
        }
    }

	public class NPOIExcel
	{
		public NPOIExcel()
		{
		}

		public System.Data.DataTable getDataTableFromExcel(string path, int startRowIndex)
		{
			NPOI.HSSF.UserModel.HSSFWorkbook workbook;
			NPOI.HSSF.UserModel.HSSFSheet sheet;
			System.Data.DataTable dtExcel = new System.Data.DataTable();
			try
			{
				using (var stream = System.IO.File.OpenRead(path))
				{
					workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(stream);
				}
				// 讀取第一個 工作表
				sheet = (NPOI.HSSF.UserModel.HSSFSheet)workbook.GetSheetAt(0);
				// 取得 第一列 也就是表頭列
				NPOI.HSSF.UserModel.HSSFRow headerRow = (NPOI.HSSF.UserModel.HSSFRow)sheet.GetRow(startRowIndex);
				// 取得 欄數
				int cellcount = headerRow.LastCellNum;
				// 從第一欄到最後一欄
				for (int i = headerRow.FirstCellNum; i < cellcount; i++)
				{
					// 建立 DataTable 的欄位
					System.Data.DataColumn column = new System.Data.DataColumn(headerRow.GetCell(i).StringCellValue.Replace(" ", "").Replace(".", ""));
					dtExcel.Columns.Add(column);
				}

				// 取得 列數
				int rowCount = sheet.LastRowNum;
				for (int i = (sheet.FirstRowNum + startRowIndex + 1); i <= sheet.LastRowNum; i++)
				{
					// 讀取 Excel 裡面的列值
					NPOI.HSSF.UserModel.HSSFRow row = (NPOI.HSSF.UserModel.HSSFRow)sheet.GetRow(i);
					// 建立 DataTable 的列
					System.Data.DataRow dataRow = dtExcel.NewRow();
					// 讀取欄位
					for (int j = row.FirstCellNum; j < cellcount; j++)
					{
						// 如果該列的值不為空
						if (row.GetCell(j) != null)
						{
							// 將資料欄位加入到 dataRow 內
							// dataRow[欄位]
							dataRow[j] = row.GetCell(j).ToString();
						}
					}
					dtExcel.Rows.Add(dataRow);
				}
			}
			catch (Exception ex)
			{
				ex.Message.ToString();
			}
			finally
			{
				// 釋放記憶體
				workbook = null;
				sheet = null;
			}
			return dtExcel;
		}
	}
}