# Horas: Rellenador de control de jornada
# Copyright (C) 2019 by Alejandro Vicario
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

from pdfrw import PdfReader, PdfWriter, PdfDict, PdfObject
import locale
import datetime
import json

datefields = [
		'DIARow{}',
		'HORA ENTRADARow{}',
		'HORA SALIDARow{}',
		'HORA ENTRADARow{}_2',
		'HORA SALIDARow{}_2',
		'HORA ENTRADARow{}_3',
		'HORA SALIDARow{}_3',
		'TOTAL HORASRow{}',
		'INCIDENCIARow{}'
]

def fillPDF(pdf_file, fields, workId):
	pdf = PdfReader(pdf_file)
	anns = pdf.pages[0]['/Annots']
	for ann in anns:
		key = ann['/T'][1:-1]
		if key in fields:
			ann.update(PdfDict(V='{}'.format(fields[key])))

	filename = 'registro{}_s{}-{}{}.pdf'.format(workId, fields['Semana'], fields['Mes'], fields['Anio'])
	pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject('true')))
	PdfWriter().write(filename, pdf) 
	print(filename)
	
def readJSON(json_file, date):
	trabajadores = []
	with open(json_file) as f:
		d = json.load(f)
		for t in d['trabajadores']:
			fields = {}
			nombre = '{} {}'.format(t['nombre'], t['apellidos'])
			fields['Nombre'] = nombre
			fields['NIF'] = t['NIF']
			fields['Centro'] = t['centro']
			fields['NOMBRE Y APELLIDOSResponsable'] = t['responsable']
			fillDate(t, date, fields)
			trabajadores.append(fields)
	return trabajadores

def fillDate(t, date, fields):
	fields['Mes'] = date.strftime('%B').upper()
	fields['Anio'] = date.strftime('%Y')
	first = datetime.datetime(day=1, month=date.month, year=date.year)
	weeknumber = date.isocalendar()[1] - first.isocalendar()[1] + 1
	fields['Semana'] = str(weeknumber)
	week = getWeek(date)
	for day in week:
		wd = day.weekday() + 1
		fields[datefields[0].format(wd)] = day.strftime('%a (%d)').capitalize()
		fields[datefields[1].format(wd)] = t['hora_entrada1']
		fields[datefields[2].format(wd)] = t['hora_salida1']
		fields[datefields[3].format(wd)] = t['hora_entrada2']
		fields[datefields[4].format(wd)] = t['hora_salida2']
		fields[datefields[7].format(wd)] = str(t['total'])

def getWeek(date):
	monday = date-datetime.timedelta(days=date.weekday())
	week = []
	for i in range(5):
		day = monday + datetime.timedelta(days=i)
		if day.month == date.month:
			week.append(day)
	return week

def genRegistros(pdf, date, json_file):
	first = datetime.datetime(day=1, month=date.month, year=date.year)
	monday = first - datetime.timedelta(days=first.weekday())
	for i in range(5):
		trabajadores = []
		week = monday + datetime.timedelta(days=i*7)
		if i == 0:
			trabajadores = readJSON('horas.json', first)
		elif week.month == date.month:
			trabajadores = readJSON('horas.json', week)
		else:
			break
		for i in range(len(trabajadores)):
			fields = trabajadores[i]
			fillPDF(pdf, fields, i)

if __name__ == "__main__":
	print("Horas v1.0")
	print("Copyright (C) 2019 by Alejandro Vicario")
	print("This is free software, and you are welcome to redistribute it")
	print("under certain conditions")
	locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
	now = datetime.datetime.now()
	genRegistros('temp.pdf', now, 'horas.json')
	print("DONE!")
