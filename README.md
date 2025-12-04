Conversió d’Excel a CSV



Script en Python per convertir fitxers Excel (.xlsx / .xls) a CSV aplicant diversos processos de neteja i formatació.



✨ Funcionalitats principals



Converteix tots els fulls d’un arxiu Excel en fitxers CSV independents.



Elimina les primeres 10 files i la primera columna (configurable).



Substitueix salts de línia i caràcters especials.



Dona format a certes columnes de data (dd/mm/YYYY).



Reemplaça punts per comes en valors numèrics.



Genera noms de fitxer segurs i evita sobreescriptures automàticament.



Mostra missatges d’error o informació mitjançant Tkinter (o via consola si no hi ha entorn gràfic).



Permet seleccionar arxius manualment si no s’indiquen per paràmetres.



📂 Funcionament general



Per a cada full d’un Excel:



Llegeix el full ometent les primeres n files (10 per defecte).



Elimina la primera columna.



Neteja salts de línia i espais sobrants.



Dona format a columnes de data (posicions 7, 17 i 18).



Converteix els valors numèrics de format anglès (.) a format europeu (,).



Genera un CSV amb separador ;, sense capçalera ni índex.



Mostra un missatge confirmant la conversió o l’error.



🛠 Requisits



Python 3.7+



Llibreries:



pandas



tkinter (per a la selecció d’arxius i missatges)



argparse



csv



Instal·lació recomanada:



pip install pandas





Tkinter ve instal·lat de sèrie amb la majoria de distribucions de Python.



🚀 Ús des de la línia de comandes

Comanda bàsica

python convert.py fitxer.xlsx



Paràmetres disponibles

usage: convert.py \[-h] \[--out-dir OUT\_DIR] \[--overwrite] \[--skiprows SKIPROWS] \[files ...]



Paràmetre	Descripció

files	Un o més fitxers Excel a convertir. Si s’omet, s'obrirà un selector de fitxers.

--out-dir	Directori on es guardaran els CSV (per defecte: .).

--overwrite	Permet sobreescriure CSV existents (per defecte, genera noms únics).

--skiprows	Files que s’ometen al principi de cada full (per defecte: 10).

Exemple complet

python convert.py dades1.xlsx dades2.xlsx --out-dir exportats --skiprows 12 --overwrite



🖱 Ús sense línia de comandes



Si executes el script sense arguments:



python convert.py





S’obrirà un selector d’arxius per triar manualment els Excel.



📁 Sortida



Per cada full, es genera un CSV amb nom:



<nom\_arxiu>\_<nom\_fulla>.csv





Si ja existeix i no s’ha indicat --overwrite, es crearà automàticament una versió única:



<nom\_arxiu>\_<nom\_fulla>\_1.csv

<nom\_arxiu>\_<nom\_fulla>\_2.csv

…


Comandes  per a compilar el codi font a un executable amb PyInstaller:

 # pyinstaller --onedir --clean conversioCSV.py --> Crea amb carpeta dist
 # pyinstaller --onefile --clean conversioCSV.py --> Crea nomes un executble
