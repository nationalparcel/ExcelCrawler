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
				if date > 20130725 and temp_file_check == -1 and xls_check != -1 and master_check == -1 and self.weekNumber >= fromDate and self.weekNumber <= toDate:
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
		self.totalCostList = []
		self.weekNumberList = []
		time.sleep(5.5)
		while count < int(toDate) + 1:
			self.totalCostList.append(0)
			self.weekNumberList.append(str(count))
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
							self.totalCost = data
							
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
					if self.companyInvoiceType == "NPLD":
						index = self.weekNumberList.index(self.weekNumber)
						self.totalCostList[index] = self.totalCostList[index] + self.totalCost
						print self.totalCostList
						print self.weekNumberList
						
						#WriteCell( r, 9, self.reportSheet, self.totalCost )
						#WriteCell( r, 1, self.reportSheet, self.weekNumber )
					
					r = r + 1
					self._workbook.Close(False)
		
		if self.companyInvoiceType == "NPLD":
			index = 0
			for i in self.weekNumberList:
				WriteCell( index + 2, 1, self.reportSheet, self.weekNumberList[index] )
				WriteCell( index + 2, 9, self.reportSheet, str( self.totalCostList[index] ))
				index += 1
		
		#else:
			
			
		self.workbook.SaveAs(saveLoc)
		self.workbook.Close(False)
		System.Diagnostics.Process.Start(saveLoc)
		self._statusTextBox.AppendText( "Finished" )
		self._statusTextBox.ScrollToCaret()