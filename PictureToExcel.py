from os import path, remove
from sys import argv, stdout
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QObject, qDebug, QDateTime
from PyQt5.QtGui import QPixmap, QImage, QColor
# excel
# python scripts\pywin32_postinstall.py -install
from win32com.client import Dispatch

class PictureToExcel(QObject):
	"""Write pixels picture to sheet excel cell = one pixel."""

	def __init__(self, picture, xls=None, visible = False, cellsquaresize = 20, parent=None):
		"""Init."""
		super(PictureToExcel, self).__init__(parent)
		self.parent = parent
		self.cellsquaresize = cellsquaresize
		self.xlsrationsquare = 7.5
		self.picture = picture
		# xls=None : same folder, name
		if xls is None:
			self.classeurname = path.join(path.dirname(self.picture) , path.splitext(path.basename(self.picture))[0] + '.xlsx')
		else:
			self.classeurname = xls
		self.prepare_worksheet(visible)
		self.write_picture()
		self.close_worksheet()

	def prepare_worksheet(self, visible = False):
		self.excel = Dispatch("Excel.Application")
		if visible:
			self.excel.Visible = 1
		self.classeur = self.excel.Workbooks.Add()
		self.classeur.Author = 'doubsman'
		self.feuille = self.classeur.Worksheets(1)
		self.feuille.name = 'PictureToExcel 1.0'
		#self.feuille.Cells(1,1).Value = "Hello, World"
		self.excel.Windows.Item(1).DisplayGridlines = False

	def write_picture(self):
		img = QImage(self.picture)
		width = img.width()
		height = img.height()
		ratio = width / (height * self.xlsrationsquare)
		bar = ProgressBar(width * height)
		for x in range(0, width):
			for y in range(0, height):
				# Square cells
				self.feuille.Cells(y + 1, x + 1).RowHeight = self.cellsquaresize 
				self.feuille.Cells(y + 1, x + 1).ColumnWidth = self.cellsquaresize * ratio 
				# copy pixel to cell
				imgpix = img.pixel(x,y)
				colors = QColor(imgpix).getRgb()
				self.feuille.Cells.Item(y + 1, x + 1).interior.color = self.rgb_to_hex(colors)
				# progress bar
				bar.update()

	def close_worksheet(self):
		if path.isfile(self.classeurname):
			remove(self.classeurname)
		self.classeur.SaveAs(self.classeurname)
		self.classeur.Close()
		self.excel.Quit()

	def rgb_to_hex(self, rgb):
		"""ws.Cells(1, i).Interior.color uses bgr in hex."""
		bgr = (rgb[2], rgb[1], rgb[0])
		strValue = '%02x%02x%02x' % bgr
		iValue = int(strValue, 16)
		return iValue


class ProgressBar:
	"""make easily a progress bar."""
	def __init__(self, steps, maxbar=100, title='Chargement'):
		if steps <= 0 or maxbar <= 0 or maxbar > 200:
			raise ValueError
		self.steps = steps
		self.maxbar = maxbar
		self.title = title
		self.perc = 0
		self._completed_steps = 0
		self.update(False)

	def update(self, increase=True):
		if increase:
			self._completed_steps += 1
		self.perc = int(self._completed_steps / self.steps * 100)
		if self._completed_steps > self.steps:
			self._completed_steps = self.steps
		steps_bar = int(self.perc / 100 * self.maxbar)
		if steps_bar == 0:
			visual_bar = self.maxbar * ' '
		else:
			visual_bar = (steps_bar - 1) * '=' + '>' + (self.maxbar - steps_bar) * ' '
		stdout.write('\r' + self.title + ' [' + visual_bar + '] ' + str(self.perc) + '%')
		stdout.flush()


if __name__ == '__main__':
	app = QApplication(argv)
	if len(argv)>1:
		# prod
		imgpath = argv[1]
	else:
		# test envt
		imgpath = r'\\Homerstation\_pro\Python\PictureToExcel\facelego.png'
	# class
	#xlspath = path.join(path.dirname(imgpath) , path.splitext(path.basename(imgpath))[0] + '.xlsx')
	BuildProcess = PictureToExcel(imgpath, None, False)