# Descripció
Aquest petit script llegeix els fitxers XLS amb dades de les avaluacions d'Esfera, que exporta EsferaPowerToys, i genera una nova pestanya amb la quantitat de RA avaluats i aprovats. 

# Com fer-lo servir:
Només s'ha provat sobre Linux Mint 22.3, que equival a Ubuntu 24.04 LTS.

## Requisits
```
apt install python3-pandas
```

## Execució
Cal deixar els fitxers XLS dins la carpeta **input** i executar l'script:
```
python3 analisi_ra.py
```

Els resultats apareixeran a la terminal i també a la carpeta **output**. 

