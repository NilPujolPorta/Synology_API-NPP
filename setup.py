from setuptools import setup

setup(name='SynologyAPI-NPP',
      version='1.6.9',
      description='API Basica per a synologys Active Backup For Buissiness',
      long_description="""
# Synology Active Backup for Business API-NPP
## Informació
- Per executar el programa s'ha de tenir instalat el python versio 3 o mes.
- Requeriments a requirements.txt.
- Configuració de la base de dades a `config/config.yaml`
- Logs de errors a `errorLogs/*txt`
- El fitxer compilar.bat transforma el .py en .pyc que es mes eficient i rapid.
- Executar amb opcio -h per veure mes opcions i funcionalitats.


## Estructura de la base de dades
```
"nom" Nom identificatiu, no es pot repetir. SENSE ESPAIS!!!!

"user" Usuari amb permisos d'administrador al active backup

"password" Contrassenya del NAS

"url" Enllaç quickconnect amb la barra final

"cookie" Per aconseguir la cookie anar al Chrome(o similar) entrar al enllaç anterior i fer inspeccionar elemento; Una vegada alla anem a l'apartat de network clickem CONTROL+R cliquem al resultat que ens sortira i busquem on esta cookie

"pandoraID" El numero identificador que te el grup de pandora per exemple Splendid foods  es 15
```
## Ús
- Executar el fitxer `synology_API.py` o `synology_API.cpython-39.pyc` amb les opcions adients. Llavors les dades es guardaran a `dadesSynology.json` i si la opcio de excel esta activada tambe es guardara a `revisio_copies_seguretat_synology_vs1.xlsx`

- Opcions
```
usage: PandoraFMS_API.cpython-39.pyc [-h] [-e] [-f RUTA] [-q] [-v] [-w URL]

Una API per a recullir informacio de la web de PandoraFMS.

optional arguments:
  -h, --help            show this help message and exit
  -e, --excel           Guardar la informacio a un excel, per defecte esta desactivat
  -f RUTA, --file RUTA  Especificar el fitxer de excel a on guardar. Per defecte es: PandoraResum.xlsx
  -q, --quiet           Nomes mostra els errors i el missatge de acabada per pantalla.
  -v, --versio          Mostra la versio
  -w URL, --web URL     Especificar la web de PandoraFMS a on accedir. Per defecte es: http://pandora.eio.cat/pandora_console/include/api.php
```
## Instal·lacio
- Utilitzant `pip`
```
pip install SynologyAPI-NPP
```
### Errors coneguts
- Si dona error per algun motiu, en els logs et donara un codi, que llavors pots mirar a errorLogs/0codisErrors.txt per saber el seu significat.
- Si s'interumpeix a mitges el excel pot quedar corromput, pero al borrar-lo  i executar-ho una altre vegada s'arregla.

### Proximament:
1. Afegir support per altres bases de dades a part de mysql

""",
      long_description_content_type='text/markdown',
      url='https://github.com/NilPujolPorta/Synology_API-NPP',
      author='Nil Pujol Porta',
      author_email='nilpujolporta@gmail.com',
      license='GNU',
      packages=['SynologyAPI'],
      install_requires=[
          'argparse',
          "setuptools>=42",
          "wheel",
          "openpyxl",
          "pyyaml",
          "requests",
          "mysql-connector-python",
          "tqdm"
      ],
	entry_points = {
        "console_scripts": ['SynologyAPI-NPP = SynologyAPI.synology_API:main']
        },
      zip_safe=False)
