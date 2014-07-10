#!/usr/bin/python3.3
import requests
import xlsxwriter
import sys
from model.movie import *

def main():
	total = len(sys.argv)
	url = sys.argv[1]
	salida = sys.argv[2]
	if(total!=3):
		print('Para ejecutar el script debe especificar la url del servidor xbmc:puerto y la rulta de salida del pdf')
		print('Ej: ./jsonxbmccall.py http://ruta:9080 /home/user/salida.pdf')
		return
	movielist2excel(getMovies(url),salida)
	print("FIN")

def movielist2excel(peliculas,salida):
	workbook = xlsxwriter.Workbook(salida)
	listatmejorablesws = workbook.add_worksheet()
	listatmejorablesws.name = 'Lista Peliculas'
	#CABECERAS
	listatmejorablesws.write(0, 0 ,'ID')
						listatmejorablesws.write(0, 1 , 'TITLE')
	listatmejorablesws.write(0, 2 , 'WIDTH')
	listatmejorablesws.write(0, 3 , 'HEIGHT')
	listatmejorablesws.write(0, 4 , 'RUTA')
	listatmejorablesws.write(0, 5 , 'IDIOMAS')
	
	#RELLENAMOS TODO
	pos = 1
	for mejo in peliculas:
		listatmejorablesws.write(pos, 0 , mejo.iden)
		listatmejorablesws.write(pos, 1 , mejo.title)
		listatmejorablesws.write(pos, 2 , mejo.reswidth)
		listatmejorablesws.write(pos, 3 , mejo.reshei)
		listatmejorablesws.write(pos, 4 , mejo.filepath)
		listatmejorablesws.write(pos, 5 , str(mejo.dual))
		pos = pos + 1
	workbook.close()	
	

def getMovies(urlServicio):
	movielist = []
	movieListRequest =	requests.get(urlServicio+'/jsonrpc?request={"jsonrpc": "2.0", "method": "VideoLibrary.GetMovies","params": {	"properties" : ["tag","imdbnumber"],"sort": { "order": "ascending", "method": "label", "ignorearticle": true } }, "id": "libMovies"}')
	lista  = movieListRequest.json()
	movieListJson=lista['result']['movies']
	for jsonMovie in movieListJson:
		movieID = jsonMovie['movieid']
		url =urlServicio+ '/jsonrpc?request={"jsonrpc": "2.0", "method": "VideoLibrary.GetMovieDetails", "params": {"movieid": %s, "properties": ["title", "imdbnumber", "streamdetails", "rating", "file"]}, "id": "1"}' % movieID
		movieInfoRequest = requests.get(url)
		movies =movieInfoRequest.json() 
		ide =movies['result']['moviedetails']['imdbnumber']
		name=movies['result']['moviedetails']['title']
		width = 0;
		height = 0;
		for media in movies['result']['moviedetails']['streamdetails']['video']:
			width = media['width'];
			height = media['height'];
		path=movies['result']['moviedetails']['file']
		#audiocad=len(movie['result']['moviedetails']['streamdetails']['audio'])
		audiocad = []
		for audio in movies['result']['moviedetails']['streamdetails']['audio']:
			audiocad.append(audio['language'])
		movielist.append(movie(ide, name, width, height, path, audiocad))
	return movielist
	
	
if __name__ == "__main__":
    main()


