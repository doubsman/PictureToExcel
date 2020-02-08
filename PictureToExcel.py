from os import rename, path
from sys import argv, path as syspath
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QObject, qDebug, QDateTime
from PyQt5.QtGui import QPixmap, QImage, QColor
# log
syspath.append(path.dirname(path.dirname(path.abspath(__file__))))
from LogPrintFile.LogPrintFile import LogPrintFile
from FilesProcessing.FilesProcessing import FilesProcessing
# excel
# python scripts\pywin32_postinstall.py -install
from win32com.client import Dispatch

class PictureToExcel(QObject):
	"""Youtube list download."""
						
	def __init__(self, picture, xls, parent=None):
		"""Init."""
		super(PictureToExcel, self).__init__(parent)
		self.parent = parent
		self.picture = picture
		self.classeurname = xls
		self.prepare_worksheet()
		self.write_picture()


	def prepare_worksheet(self):
		# Excel prepare worksheet
		self.excel = Dispatch("Excel.Application")
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
		for x in range(0, width):
			for y in range(0, height):
				c = img.pixel(x,y)
				colors = QColor(c).getRgb()
				self.feuille.Cells.Item(y + 1, x + 1).interior.color = self.rgb_to_hex(colors)
		self.classeur.SaveAs(self.classeurname)
		self.classeur.Close()
		self.excel.Quit()

	def rgb_to_hex(self, rgb):
		"""ws.Cells(1, i).Interior.color uses bgr in hex."""
		bgr = (rgb[2], rgb[1], rgb[0])
		strValue = '%02x%02x%02x' % bgr
		iValue = int(strValue, 16)
		return iValue

if __name__ == '__main__':
	app = QApplication(argv)
	if len(argv)>1:
		# prod
		imgpath = argv[1]
		xlspath = argv[2]
	else:
		# test envt
		imgpath = r'R:\Python\PictureToExcel\totoro.png'
		xlspath = r'R:\Python\PictureToExcel\totoro.xls'
	# class
	BuildProcess = PictureToExcel(imgpath, xlspath)
