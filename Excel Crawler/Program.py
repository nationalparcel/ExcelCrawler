﻿import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('Microsoft.Office.Interop.Excel')

from System.Windows.Forms import Application
import MainForm

Application.EnableVisualStyles()
form = MainForm.MainForm()
Application.Run(form)
