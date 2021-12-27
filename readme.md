
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

### Errors coneguts
- Si dona error per algun motiu, en els logs et donara un codi, que llavors pots mirar a errorLogs/0codisErrors.txt per saber el seu significat.
- A vegades peta la primera vegada el access al excel, si passa tornar a executar(recomanat fer-ho sempre).
- Si s'interumpeix a mitges el excel pot quedar corromput, pero al borrar-lo  i executar-ho una altre vegada s'arregla.

### Proximament:
1. Afegir support per altres bases de dades a part de mysql
