
# Synology Active Backup for Business API-NPP

- Per executar el programa s'ha de tenir instalat el python versio 3 o mes.
- Requeriments a requirements.txt.

    
- Per afegir un nou dispositiu ves al fitxer data/dispositius.json i despres de l'ultim dispositiu afegim una coma
  i seguidament el seguent amb les dades que corresponguin:
```
    {
      "nom": "Nom identificatiu SENSE ESPAIS!!!!",
      "user": "usuari amb permisos d'administrador al active backup",
      "password": "contrassenya",
      "url": "Enllaç quickconnect amb la barra final",
      "cookie": "Per aconseguir la cookie anar al Chrome(o similar) entrar al enllaç anterior i fer inspeccionar elemento; 
       Una vegada alla anem a l'apartat de network clickem CONTROL+R cliquem al resultat que ens sortira i busquem on esta cookie"
      "pandoraID": "El numero identificador que te el grup de pandora per exemple Splendid foods  es 15"
    }
```
- El fitxer compilar.bat transforma el .py en .pyc que es mes eficient i rapid.
- Si dona error per algun motiu, en els logs et donara un codi, que llavors pots mirar a errorLogs/0codisErrors.txt per saber el seu significat.
- A vegades peta la primera vegada el access al excel, si passa tornar a executar(recomanat fer-ho sempre).
- Si s'interumpeix a mitges el excel pot quedar corromput, pero al borrar-lo  i executar-ho una altre vegada s'arregla.

