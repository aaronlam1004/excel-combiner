import os
from os import listdir
import ntpath
from pathlib import Path

import xlrd
import xlsxwriter
from xlutils.copy import copy

from tkinter import Tk, filedialog, Frame, Button, Listbox, messagebox, PhotoImage, Toplevel
from tkinter import *


class SubApplication(Toplevel):
	def __init__(self, parent, workbook):
		self.values = None
		Toplevel.__init__(self, parent)

		self.workbook = workbook
		options = [str(s) for s in self.workbook.sheet_names()]
		self.option = Listbox(self, width = 40, height = 20, selectmode = MULTIPLE)
		for op in options:
			self.option.insert(self.option.size(), op)
		self.option.pack()

		process = Button(self, text = "Select", command = self.processInput)
		process.pack()

		self.transient(parent)
		self.grab_set()
		parent.wait_window(self)
		
	def processInput(self):
		self.values = tuple(op for op in self.option.curselection())
		if len(self.values) != 0:
			self.destroy()

	def getValues(self):
		return self.values

class Application(Frame):
	def __init__(self, parent = None):
		Frame.__init__(self, parent)

		parent.title("Excel Copier")
		parent.geometry("500x500")

		self.pack()
		self.widgets()

		self.masterImportFile = ""
		self.importedFiles = []		

	def widgets(self):
		self.masterFile = Button(self, text = "Master file", command = self.getMasterFile)
		self.masterFile.pack()

		self.masterArea = Listbox(self, width = 40, height = 1)
		self.masterArea.pack()

		self.importFiles = Button(self, text = "+", command = self.gatherFiles)
		self.importFiles.pack(side = LEFT)

		self.deleteFiles = Button(self, text = "-", command = self.deleteFiles)
		self.deleteFiles.pack(side = LEFT)

		self.downButton = Button(self, text = "↓", command = self.moveFileDown)
		self.downButton.pack(side = RIGHT)

		self.upButton = Button(self, text = "↑", command = self.moveFileUp)
		self.upButton.pack(side = RIGHT)

		self.importArea = Listbox(self, width = 40, height = 20, selectmode = MULTIPLE)
		self.importArea.pack()

		self.processFiles = Button(self, text = "Create new file", command = self.processFiles)
		self.processFiles.pack()

		self.addFiles = Button(self, text = "Add to file", command = self.addToFile)
		self.addFiles.pack()

	def moveFileUp(self):
		fileIndex = self.importArea.curselection()
		if 0 not in fileIndex:
			for file in fileIndex:	
				self.importArea.insert(file - 1, ntpath.basename(self.importedFiles[file]))
				self.importArea.delete(file + 1)
				self.importedFiles.insert(file - 1, self.importedFiles[file])
				self.importedFiles.pop(file + 1)

	def moveFileDown(self):
		fileIndex = self.importArea.curselection()
		if len(self.importedFiles) - 1 not in fileIndex:
			for file in fileIndex[::-1]:
				self.importArea.insert(file + 2, ntpath.basename(self.importedFiles[file]))
				self.importArea.delete(file)
				self.importedFiles.insert(file + 2, self.importedFiles[file])
				self.importedFiles.pop(file)

	def getMasterFile(self):
		mFile = filedialog.askopenfilename(initialdir = os.getcwd(), title = "Select file")
		if os.path.splitext(mFile)[1] in (".xls", ".xlsx", ".xlsm"):
			self.masterImportFile = mFile
			self.masterArea.delete(0);
			self.masterArea.insert(0, ntpath.basename(self.masterImportFile))
		else:
			messagebox.showinfo("Error", "This program does not support " + os.path.splitext(mFile)[1] + " files. Only .xls, .xlsx, or .xlsm")

	def gatherFiles(self):
		imported = [file for file in filedialog.askopenfilenames(initialdir = os.getcwd(), title = "Select file(s)")]
		for file in imported:
			if os.path.splitext(file)[1] in (".xls", ".xlsx", ".xlsm"):
				self.importedFiles.append(file)
				self.importArea.insert(self.importArea.size(), ntpath.basename(file))
			else:
				messagebox.showinfo("Error", "This program does not support " + os.path.splitext(file)[1] + " files. Only .xls, .xlsx, or .xlsm")

	def deleteFiles(self):
		toDelete = self.importArea.curselection()
		deleteIndex = 0
		for delete in toDelete:
			delete = delete - deleteIndex
			self.importedFiles.pop(delete)
			self.importArea.delete(delete)
			deleteIndex += 1

	def addToFile(self):
		if self.masterImportFile is "":
			messagebox.showinfo("Error", "No \"Master file\" selected.")
		else:
			if os.path.splitext(self.masterImportFile)[1] == ".xlsx":
				self.xlsxCopy()
				messagebox.showinfo("Excel file created", "Finished adding into file: " + ntpath.basename(self.masterImportFile))	
				for i in range(len(self.importedFiles)):
					self.importArea.delete(0)
				self.importedFiles = []
			else:
				work = self.xlsWrite()
				work.save(self.masterImportFile)
				messagebox.showinfo("Excel file created", "Finished adding into file: " + ntpath.basename(self.masterImportFile))	
				for i in range(len(self.importedFiles)):
					self.importArea.delete(0)
				self.importedFiles = []

	def processFiles(self):
		if self.masterImportFile is "":
			messagebox.showinfo("Error", "No \"Master file\" selected.")
		else:
			work = self.xlsWrite()
			savefile = filedialog.asksaveasfilename(initialdir = os.getcwd(), title = "Save as")
			if os.path.splitext(savefile)[0] == "":
				messagebox.showinfo("Error", "No save file specified, please try again.")
			elif ntpath.basename(savefile + ".xls") in [f for f in listdir(Path(savefile).parent)]:
				overwrite = messagebox.askquestion("Save file that already exists", "A file with this name already exists, do you want to overwrite that file?", icon = "warning")
				if overwrite == "yes":
					work.save(savefile + ".xls")
					messagebox.showinfo("Excel file created", "Finished creating new Excel file.")
					for i in range(len(self.importedFiles)):
						self.importArea.delete(0)
						self.importedFiles = []
			elif os.path.splitext(savefile)[1] == "":
				work.save(savefile + ".xls")
				messagebox.showinfo("Excel file created", "Finished creating new Excel file.")
				for i in range(len(self.importedFiles)):
					self.importArea.delete(0)
				self.importedFiles = []
			else:
				work.save(savefile)
				messagebox.showinfo("Excel file created", "Finished creating new Excel file.")
				for i in range(len(self.importedFiles)):
					self.importArea.delete(0)
				self.importedFiles = []

	def xlsxCopy(self):
		readsheet = xlrd.open_workbook(self.masterImportFile).sheet_by_index(0)
		wb = xlsxwriter.Workbook(self.masterImportFile)
		wksht = wb.add_worksheet()
		for i in range(readsheet.nrows):
			for j in range(readsheet.ncols):
				wksht.write(i, j, readsheet.cell_value(i, j))

		self.writeImportedInfo(wksht)
		wb.close()


	def xlsWrite(self):
		work = copy(xlrd.open_workbook(self.masterImportFile))
		wksht = work.get_sheet(0)

		#workCols = xlrd.open_workbook(self.masterImportFile).sheet_by_index(0).ncols
		self.writeImportedInfo(wksht)
		return work

	def writeImportedInfo(self, wksht):
		rowIndex = xlrd.open_workbook(self.masterImportFile).sheet_by_index(0).nrows
		for file in self.importedFiles:
			readwb = xlrd.open_workbook(file)
			totalCount = 0
			skip = 0
			values = tuple()
			if len(readwb.sheets()) > 1:
				sb = SubApplication(self, readwb)
				values = sb.getValues()
				if values is not None and len(values) >= 1:
					for s in values:
						sheet = readwb.sheet_by_index(s)
						for i in range(sheet.nrows):
							if i == 3:
								totalCount += 1
								skip += 1
							else:
								blankSheet = 0
								for j in range(sheet.ncols):
									if sheet.cell_value(i, j) != "":
										wksht.write(i + rowIndex - skip, j, sheet.cell_value(i, j))
									else:
										blankSheet += 1
								if blankSheet == sheet.ncols:
									totalCount += 1
									skip += 1
						rowIndex += sheet.nrows - totalCount
			else:
				sheet = readwb.sheet_by_index(0)
				for i in range(sheet.nrows):
					if i == 3:
						totalCount += 1
						skip += 1
					else:
						blankSheet = 0
						for j in range(sheet.ncols):
							if sheet.cell_value(i, j) != "":
								wksht.write(i + rowIndex - skip, j, sheet.cell_value(i, j))
							else:
								blankSheet += 1
						if blankSheet == sheet.ncols:
							totalCount += 1
							skip += 1
				rowIndex += sheet.nrows - totalCount

root = Tk()
app = Application(parent = root)
app.mainloop()
