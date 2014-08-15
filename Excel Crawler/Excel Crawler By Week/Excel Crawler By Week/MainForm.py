import System.Drawing
import System.Windows.Forms
import sys
sys.path.append(r'C:\Program Files (x86)\IronPython 2.7\Lib')
import os
import time
import datetime
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
excel.Visible = False

from System.Drawing import *
from System.Windows.Forms import *

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._saveLocationButton = System.Windows.Forms.Button()
		self._label1 = System.Windows.Forms.Label()
		self._label2 = System.Windows.Forms.Label()
		self._dirToSearchTextBox = System.Windows.Forms.TextBox()
		self._label3 = System.Windows.Forms.Label()
		self._dirToSearchButton = System.Windows.Forms.Button()
		self._saveLocationTextBox = System.Windows.Forms.TextBox()
		self._label4 = System.Windows.Forms.Label()
		self._startCrawl = System.Windows.Forms.Button()
		self._statusTextBox = System.Windows.Forms.RichTextBox()
		self._weekFrom = System.Windows.Forms.ComboBox()
		self._weekTo = System.Windows.Forms.ComboBox()
		self.SuspendLayout()
		# 
		# saveLocationButton
		# 
		self._saveLocationButton.Location = System.Drawing.Point(415, 252)
		self._saveLocationButton.Name = "saveLocationButton"
		self._saveLocationButton.Size = System.Drawing.Size(75, 23)
		self._saveLocationButton.TabIndex = 0
		self._saveLocationButton.Text = "Save"
		self._saveLocationButton.UseVisualStyleBackColor = True
		self._saveLocationButton.Click += self.saveDialog
		# 
		# label1
		# 
		self._label1.Location = System.Drawing.Point(12, 42)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(40, 14)
		self._label1.TabIndex = 2
		self._label1.Text = "From:"
		# 
		# label2
		# 
		self._label2.Location = System.Drawing.Point(12, 89)
		self._label2.Name = "label2"
		self._label2.Size = System.Drawing.Size(39, 15)
		self._label2.TabIndex = 4
		self._label2.Text = "To:"
		# 
		# dirToSearchTextBox
		# 
		self._dirToSearchTextBox.Location = System.Drawing.Point(12, 138)
		self._dirToSearchTextBox.Name = "dirToSearchTextBox"
		self._dirToSearchTextBox.ReadOnly = True
		self._dirToSearchTextBox.Size = System.Drawing.Size(478, 20)
		self._dirToSearchTextBox.TabIndex = 5
		self._dirToSearchTextBox.Text = "O:\\NPL Invoices\\2014 INVOICES"
		# 
		# label3
		# 
		self._label3.Location = System.Drawing.Point(12, 121)
		self._label3.Name = "label3"
		self._label3.Size = System.Drawing.Size(189, 14)
		self._label3.TabIndex = 6
		self._label3.Text = "Directory to Search"
		# 
		# dirToSearchButton
		# 
		self._dirToSearchButton.Location = System.Drawing.Point(415, 164)
		self._dirToSearchButton.Name = "dirToSearchButton"
		self._dirToSearchButton.Size = System.Drawing.Size(75, 23)
		self._dirToSearchButton.TabIndex = 7
		self._dirToSearchButton.Text = "Search"
		self._dirToSearchButton.UseVisualStyleBackColor = True
		self._dirToSearchButton.Click += self.dirToSearchDialog
		# 
		# saveLocationTextBox
		# 
		self._saveLocationTextBox.Location = System.Drawing.Point(13, 226)
		self._saveLocationTextBox.Name = "saveLocationTextBox"
		self._saveLocationTextBox.ReadOnly = True
		self._saveLocationTextBox.Size = System.Drawing.Size(477, 20)
		self._saveLocationTextBox.TabIndex = 8
		# 
		# label4
		# 
		self._label4.Location = System.Drawing.Point(12, 207)
		self._label4.Name = "label4"
		self._label4.Size = System.Drawing.Size(100, 16)
		self._label4.TabIndex = 9
		self._label4.Text = "File Save Location"
		# 
		# startCrawl
		# 
		self._startCrawl.Location = System.Drawing.Point(212, 286)
		self._startCrawl.Name = "startCrawl"
		self._startCrawl.Size = System.Drawing.Size(75, 23)
		self._startCrawl.TabIndex = 10
		self._startCrawl.Text = "GO!"
		self._startCrawl.UseVisualStyleBackColor = True
		self._startCrawl.Click += self.runLogic
		# 
		# statusTextBox
		# 
		self._statusTextBox.Location = System.Drawing.Point(12, 321)
		self._statusTextBox.Name = "statusTextBox"
		self._statusTextBox.ReadOnly = True
		self._statusTextBox.Size = System.Drawing.Size(478, 186)
		self._statusTextBox.TabIndex = 11
		self._statusTextBox.Text = ""
		# 
		# weekFrom
		# 
		self._weekFrom.FormattingEnabled = True
		self._weekFrom.Items.AddRange(System.Array[System.Object](
			["01",
			"02",
			"03",
			"04",
			"05",
			"06",
			"07",
			"08",
			"09",
			"10",
			"11",
			"12",
			"13",
			"14",
			"15",
			"16",
			"17",
			"18",
			"19",
			"20",
			"21",
			"22",
			"23",
			"24",
			"25",
			"26",
			"27",
			"28",
			"29",
			"30",
			"31",
			"32",
			"33",
			"34",
			"35",
			"36",
			"37",
			"38",
			"39",
			"40",
			"41",
			"42",
			"43",
			"44",
			"45",
			"46",
			"47",
			"48",
			"49",
			"50",
			"51",
			"52"]))
		self._weekFrom.Location = System.Drawing.Point(58, 42)
		self._weekFrom.Name = "weekFrom"
		self._weekFrom.Size = System.Drawing.Size(121, 21)
		self._weekFrom.TabIndex = 12
		# 
		# weekTo
		# 
		self._weekTo.FormattingEnabled = True
		self._weekTo.Items.AddRange(System.Array[System.Object](
			["01",
			"02",
			"03",
			"04",
			"05",
			"06",
			"07",
			"08",
			"09",
			"10",
			"11",
			"12",
			"13",
			"14",
			"15",
			"16",
			"17",
			"18",
			"19",
			"20",
			"21",
			"22",
			"23",
			"24",
			"25",
			"26",
			"27",
			"28",
			"29",
			"30",
			"31",
			"32",
			"33",
			"34",
			"35",
			"36",
			"37",
			"38",
			"39",
			"40",
			"41",
			"42",
			"43",
			"44",
			"45",
			"46",
			"47",
			"48",
			"49",
			"50",
			"51",
			"52"]))
		self._weekTo.Location = System.Drawing.Point(58, 83)
		self._weekTo.Name = "weekTo"
		self._weekTo.Size = System.Drawing.Size(121, 21)
		self._weekTo.TabIndex = 13
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(502, 519)
		self.Controls.Add(self._weekTo)
		self.Controls.Add(self._weekFrom)
		self.Controls.Add(self._statusTextBox)
		self.Controls.Add(self._startCrawl)
		self.Controls.Add(self._label4)
		self.Controls.Add(self._saveLocationTextBox)
		self.Controls.Add(self._dirToSearchButton)
		self.Controls.Add(self._label3)
		self.Controls.Add(self._dirToSearchTextBox)
		self.Controls.Add(self._label2)
		self.Controls.Add(self._label1)
		self.Controls.Add(self._saveLocationButton)
		self.Name = "MainForm"
		self.Text = "ExcelCrawler"
		self.ResumeLayout(False)
		self.PerformLayout()

	#search Directory dialog
	def dirToSearchDialog(self, sender, e):
		dialog = FolderBrowserDialog()
		dialog.RootFolder = System.Environment.SpecialFolder.MyComputer
		dialog.SelectedPath = "O:\\\\"
		
		if (dialog.ShowDialog(self) == DialogResult.OK):
			self._dirToSearchTextBox.Text = dialog.SelectedPath
	
	#save location dialog	
	def saveDialog(self, sender, e):
		dialog = SaveFileDialog()
		dialog.DefaultExt = "xlsx"
		dialog.FileName = "Excel Crawler Report"
		dialog.InitialDirectory = "%USERPROFILE%\\Desktop\\"
		dialog.AddExtension = "xlsx"
		
		if (dialog.ShowDialog(self) == DialogResult.OK):
			self._saveLocationTextBox.Text = dialog.FileName
			
			
	############Start of Logic###################################
	def runLogic(self, sender, e):
		headings = [
			'Total',
			'Date',
			'Invoice Number',
			'Pieces',
			'Weight',
			'Customer Code',
			'Facility Code',
			'File Name'
		]
	
		cell_write_control = [
			1,
			2,
			3,
			4,
			5,
			6,
			7,
			8
		]
		
		reportHeaders = [
			'Week',
			'ATL',
			'BDL',
			'BWI',
			'LAX',
			'MCO',
			'MIA',
			'TPA',
			'Total:',
			'NPLD'
		]
	
		def IsValidInvoice(path, filename):
			seconds = os.path.getctime(path)
			date = time.strftime('%Y%m%d', time.localtime(seconds))
			date = int(date)
			periodIndex = filename.IndexOf("4")
			if (periodIndex != -1):
				try:
					self.weekNumber = filename.Substring(periodIndex + 4, 2)
					
				except:
					self._statusTextBox.AppendText( "Error with " + filename + " is this a quote?" )
					self._statusTextBox.ScrollToCaret()
					return False
				
			else:
				self.weekNumber = 0
				
			fromDate = self._weekFrom.SelectedItem
			toDate = self._weekTo.SelectedItem
			temp_file_check = filename.find("~$")
			xls_check = filename.find('.xlsx')
			master_check = filename.lower()
			master_check = master_check.find("master")
			if xls_check == -1 and date > 20130729:
				self._statusTextBox.AppendText(  "Skipping " + filename + " Please Convert " + filename +  "to the correct format \n \n" )
				self._statusTextBox.ScrollToCaret()
				#os.system("pause")
			try:
				if date > 20130725 and temp_file_check == -1 and xls_check != -1 and master_check == -1 and self.weekNumber >= fromDate and self.weekNumber <= toDate and "VOID" not in filename :
					return True
				else:
					return False
				
			except:
				self._statusTextBox.AppendText( "Error with " + filename + " is this a quote?" )
				self._statusTextBox.ScrollToCaret()
				return False
		
		def WriteHeadings(titles, worksheet, save_loc):
			_c = 1
			for i in titles:
				WriteCell( 1, _c, worksheet, i)
				_c = _c + 1
				#SaveExcelFile(save_loc)
				
		def WriteCell(row, column, worksheet, data):
			worksheet.Cells [row, column] = data
			
		def representsInt(s):
			try:
				int(s)
				return True
			
			except ValueError:
				return False
			
		#############End of Logic Functions################################
		
		self.workbook = excel.Workbooks.Add()
		self.worksheet = self.workbook.ActiveSheet
		self.reportSheet = self.workbook.Worksheets.Add()
		self.reportSheet.Name = "Report By Fac"
		self.workbook.Application.DisplayAlerts = False
		r = 2
		c = 1
		saveLoc = self._saveLocationTextBox.Text
		dirToSearch = self._dirToSearchTextBox.Text
		if "NPL Dedicated Invoices" in dirToSearch:
			self.companyInvoiceType = "NPLD"
		else:
			self.companyInvoiceType = "NPL"
			
		if ( saveLoc == "" ):
			MessageBox.Show( "Please Choose a Save Location" )
			return
		
		fromDate = self._weekFrom.SelectedItem
		toDate = self._weekTo.SelectedItem
		count = fromDate
		count = int(count)
		self.atlWeekNumberList = []
		self.bdlWeekNumberList = []
		self.bwiWeekNumberList = []
		self.laxWeekNumberList = []
		self.mcoWeekNumberList = []
		self.miaWeekNumberList = []
		self.tpaWeekNumberList = []
		self.npldWeekNumberList = []
		
		self.atlTotalCostList = []
		self.bdlTotalCostList = []
		self.bwiTotalCostList = []
		self.laxTotalCostList = []
		self.mcoTotalCostList = []
		self.miaTotalCostList = []
		self.tpaTotalCostList = []
		self.npldTotalCostList = []
		time.sleep(5.5)
		while count < int(toDate) + 1:
			self.atlTotalCostList.append(0)
			self.bdlTotalCostList.append(0)
			self.bwiTotalCostList.append(0)
			self.laxTotalCostList.append(0)
			self.mcoTotalCostList.append(0)
			self.miaTotalCostList.append(0)
			self.tpaTotalCostList.append(0)
			self.npldTotalCostList.append(0)
			self.atlWeekNumberList.append(str(count))
			self.bdlWeekNumberList.append(str(count))
			self.bwiWeekNumberList.append(str(count))
			self.laxWeekNumberList.append(str(count))
			self.mcoWeekNumberList.append(str(count))
			self.miaWeekNumberList.append(str(count))
			self.tpaWeekNumberList.append(str(count))
			self.npldWeekNumberList.append(str(count))
			count = count + 1
			
		WriteHeadings(headings, self.worksheet, "test")
		WriteHeadings(reportHeaders, self.reportSheet, "")
		self._statusTextBox.Text = "Starting \n \n"
		for path, subdir, files in os.walk(dirToSearch):
			for filename in files:
				invoicePath = path + "\\" + filename
				valid = IsValidInvoice(invoicePath, filename)
				if valid == True:
					try:
						self._workbook = excel.Workbooks.Open(invoicePath)
						self._worksheet = self._workbook.ActiveSheet
					except:
						self._statusTextBox.AppendText( "Failed to Open " + invoicePath + "\n \n" )
						self._statusTextBox.ScrollToCaret()
						break
					
					self._statusTextBox.AppendText( invoicePath + "\n \n" )
					self._statusTextBox.ScrollToCaret()
					for i in cell_write_control:
						data = self._worksheet.Cells[1, i]
						data = data.Value2
						if i == 1:
							if representsInt( data ) == True:
								self.totalCost = data
								
							else:
								selftotalCost = 0
							
						if i == 3:
							data = str(data)
							customerCode = data
							
						if i == 6 and customerCode != "None":
							customerCode = customerCode[1:4]
							data = customerCode
							
						if i == 2:
							range = self.worksheet.Cells[r, i]
							range.NumberFormat = "MM/DD/YYYY"
							
						WriteCell( r, i, self.worksheet, data )
					
					fileNameColumn = len(headings)
					WriteCell( r, fileNameColumn, self.worksheet, invoicePath)
					print path
					if "NPL Dedicated" in path:
						index = self.npldWeekNumberList.index(self.weekNumber)
						self.npldTotalCostList[index] = self.npldTotalCostList[index] + int(self.totalCost)
						
					if "ATLANTA" in path:
						index = self.atlWeekNumberList.index(self.weekNumber)
						self.atlTotalCostList[index] = self.atlTotalCostList[index] + int(self.totalCost)
						
					if "BDL" in path:
						index = self.bdlWeekNumberList.index(self.weekNumber)
						self.bdlTotalCostList[index] = self.bdlTotalCostList[index] + int(self.totalCost)
						
					if "BALTIMORE" in path:
						index = self.bwiWeekNumberList.index(self.weekNumber)
						self.bwiTotalCostList[index] = self.bwiTotalCostList[index] + int(self.totalCost)
						
					if "LOS ANGELES" in path:
						index = self.laxWeekNumberList.index(self.weekNumber)
						self.laxTotalCostList[index] = self.laxTotalCostList[index] + int(self.totalCost)
						
					if "ORLANDO" in path:
						index = self.mcoWeekNumberList.index(self.weekNumber)
						self.mcoTotalCostList[index] = self.mcoTotalCostList[index] + int(self.totalCost)
						
					if "MIAMI" in path:
						index = self.miaWeekNumberList.index(self.weekNumber)
						self.miaTotalCostList[index] = self.miaTotalCostList[index] + int(self.totalCost)
						
					if "TAMPA" in path:
						index = self.tpaWeekNumberList.index(self.weekNumber)
						self.tpaTotalCostList[index] = self.tpaTotalCostList[index] + int(self.totalCost)
					
					r = r + 1
					self._workbook.Close(False)
		
		index = 0
		for i in self.npldWeekNumberList:
			WriteCell( index + 2, 1, self.reportSheet, self.npldWeekNumberList[index] )
			WriteCell( index + 2, 10, self.reportSheet, str( self.npldTotalCostList[index] ))
			WriteCell( index + 2, 2, self.reportSheet, str( self.atlTotalCostList[index] ))
			WriteCell( index + 2, 3, self.reportSheet, str( self.bdlTotalCostList[index] ))
			WriteCell( index + 2, 4, self.reportSheet, str( self.bwiTotalCostList[index] ))
			WriteCell( index + 2, 5, self.reportSheet, str( self.laxTotalCostList[index] ))
			WriteCell( index + 2, 6, self.reportSheet, str( self.mcoTotalCostList[index] ))
			WriteCell( index + 2, 7, self.reportSheet, str( self.miaTotalCostList[index] ))
			WriteCell( index + 2, 8, self.reportSheet, str( self.tpaTotalCostList[index] ))
			index += 1
		
		def totalCostLoop( list ):
			totalCostReturn = 0
			for item in list:
				totalCostReturn = totalCostReturn + item
			
			return totalCostReturn

		WriteCell( index + 2, 1, self.reportSheet, "Total:" )
		grandTotal = 0
		
		atlTotal = totalCostLoop( self.atlTotalCostList )
		#grandTotal = grandTotal + atlTotal
		WriteCell( index + 2, 2, self.reportSheet, atlTotal )
		
		bdlTotal = totalCostLoop( self.bdlTotalCostList )
		#grandTotal = grandTotal + bdlTotal
		WriteCell( index + 2, 3, self.reportSheet, bdlTotal )
		
		bwiTotal = totalCostLoop( self.bwiTotalCostList )
		#grandTotal = grandTotal + bwiTotal
		WriteCell( index + 2, 4, self.reportSheet, bwiTotal )
		
		laxTotal = totalCostLoop( self.laxTotalCostList )
		#grandTotal = grandTotal + laxTotal
		WriteCell( index + 2, 5, self.reportSheet, laxTotal )
		
		mcoTotal = totalCostLoop( self.mcoTotalCostList )
		#grandTotal = grandTotal + mcoTotal
		WriteCell( index + 2, 6, self.reportSheet, mcoTotal )
		
		miaTotal = totalCostLoop( self.miaTotalCostList )
		#grandTotal = grandTotal + miaTotal
		WriteCell( index + 2, 7, self.reportSheet, miaTotal )
		
		tpaTotal = totalCostLoop( self.tpaTotalCostList )
		#grandTotal = grandTotal + tpaTotal
		WriteCell( index + 2, 8, self.reportSheet, tpaTotal )
		
		npldTotal = totalCostLoop( self.npldTotalCostList )
		WriteCell( index + 2, 10, self.reportSheet, npldTotal )
		
		weekTotal = 0
		for( i, self.npldWeekNumberList ) in enumerate( self.npldWeekNumberList ):
			weekTotal = weekTotal + self.atlTotalCostList[ i ]
			weekTotal = weekTotal + self.bdlTotalCostList[ i ]
			weekTotal = weekTotal + self.bwiTotalCostList[ i ]
			weekTotal = weekTotal + self.laxTotalCostList[ i ]
			weekTotal = weekTotal + self.mcoTotalCostList[ i ]
			weekTotal = weekTotal + self.miaTotalCostList[ i ]
			weekTotal = weekTotal + self.tpaTotalCostList[ i ]
			WriteCell( i + 2, 9, self.reportSheet, weekTotal )
			grandTotal = grandTotal + weekTotal
			weekTotal = 0
			
		WriteCell( index + 2, 9, self.reportSheet, grandTotal )

		self.workbook.SaveAs(saveLoc)
		self.workbook.Close(False)
		System.Diagnostics.Process.Start(saveLoc)
		self._statusTextBox.AppendText( "Finished" )
		self._statusTextBox.ScrollToCaret()