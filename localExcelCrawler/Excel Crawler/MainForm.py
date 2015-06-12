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
		self._fromDateTimePicker = System.Windows.Forms.DateTimePicker()
		self._label1 = System.Windows.Forms.Label()
		self._toDateTimePicker = System.Windows.Forms.DateTimePicker()
		self._label2 = System.Windows.Forms.Label()
		self._dirToSearchTextBox = System.Windows.Forms.TextBox()
		self._label3 = System.Windows.Forms.Label()
		self._dirToSearchButton = System.Windows.Forms.Button()
		self._saveLocationTextBox = System.Windows.Forms.TextBox()
		self._label4 = System.Windows.Forms.Label()
		self._startCrawl = System.Windows.Forms.Button()
		self._statusTextBox = System.Windows.Forms.RichTextBox()
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
		# fromDateTimePicker
		# 
		self._fromDateTimePicker.Location = System.Drawing.Point(58, 36)
		self._fromDateTimePicker.Name = "fromDateTimePicker"
		self._fromDateTimePicker.Size = System.Drawing.Size(200, 20)
		self._fromDateTimePicker.TabIndex = 1
		# 
		# label1
		# 
		self._label1.Location = System.Drawing.Point(12, 42)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(40, 14)
		self._label1.TabIndex = 2
		self._label1.Text = "From:"
		# 
		# toDateTimePicker
		# 
		self._toDateTimePicker.Location = System.Drawing.Point(58, 84)
		self._toDateTimePicker.Name = "toDateTimePicker"
		self._toDateTimePicker.Size = System.Drawing.Size(200, 20)
		self._toDateTimePicker.TabIndex = 3
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
		self._dirToSearchTextBox.Text = "O:\\Invoices\\2014 INVOICES"
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
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(502, 519)
		self.Controls.Add(self._statusTextBox)
		self.Controls.Add(self._startCrawl)
		self.Controls.Add(self._label4)
		self.Controls.Add(self._saveLocationTextBox)
		self.Controls.Add(self._dirToSearchButton)
		self.Controls.Add(self._label3)
		self.Controls.Add(self._dirToSearchTextBox)
		self.Controls.Add(self._label2)
		self.Controls.Add(self._toDateTimePicker)
		self.Controls.Add(self._label1)
		self.Controls.Add(self._fromDateTimePicker)
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
	
		def IsValidInvoice(path, filename):
			seconds = os.path.getctime(path)
			date = time.strftime('%Y%m%d', time.localtime(seconds))
			date = int(date)
			fromDate = self._fromDateTimePicker.Text
			fromDate = time.strptime(fromDate, '%Y%m%d')
			fromDate = time.strftime('%Y%m%d', fromDate)
			fromDate = int(fromDate)
			toDate = self._toDateTimePicker.Text
			toDate = time.strptime(toDate, '%Y%m%d')
			toDate = time.strftime('%Y%m%d', toDate)
			toDate = int(toDate)
			temp_file_check = filename.find("~$")
			xls_check = filename.find('.xlsx')
			master_check = filename.lower()
			master_check = master_check.find("master")
			if xls_check == -1 and date > 20130729:
				self._statusTextBox.AppendText(  "Skipping " + filename + " Please Convert " + filename +  "to the correct format \n \n" )
				self._statusTextBox.ScrollToCaret()
				#os.system("pause")
			if date > 20130725 and temp_file_check == -1 and xls_check != -1 and master_check == -1 and date >= fromDate and date <= toDate:
				return True
			else:
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
		self.workbook.Application.DisplayAlerts = False
		r = 2
		c = 1
		saveLoc = self._saveLocationTextBox.Text
		dirToSearch = self._dirToSearchTextBox.Text
		if ( saveLoc == "" ):
			MessageBox.Show( "Please Choose a Save Location" )
			return
		
		WriteHeadings(headings, self.worksheet, "test")
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
					r = r + 1
					self._workbook.Close(False)
		self.workbook.SaveAs(saveLoc)
		self.workbook.Close(False)
		System.Diagnostics.Process.Start(saveLoc)
		self._statusTextBox.AppendText( "Finished" )
		self._statusTextBox.ScrollToCaret()
		MessageBox.Show( "All Done! Shawn is awesome :)" )