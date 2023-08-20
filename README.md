# VBA-Inspect

[![forthebadge](https://forthebadge.com/images/badges/made-with-python.svg)](http://forthebadge.com)  [![forthebadge](https://forthebadge.com/images/badges/built-by-developers.svg)](http://forthebadge.com)  [![forthebadge](https://forthebadge.com/images/badges/check-it-out.svg)](http://forthebadge.com)

VBA-Inspect est un **outil d'extraction du code VBA** présent principalement dans les fichiers Excel, avec pour objectif final de fournir rapidement une vue d'ensemble du code présent dans ces fichiers et commencer la chasse au Shadow-IT (extractions "sauvages", opérations de CRUD via des connexions OLEDB/ODBC, ...).

## Pour commencer

Pour inspecter des ressources, nous allons utiliser des scripts en Python.

### Pré-requis

Ce qu'il est requis pour commencer avec votre projet...

- [Python](https://www.python.org/downloads/windows/)

### Installation

Les étapes pour utiliser les scripts....

1. Installer Python sur votre machine
2. Clôner le dépôt [VBA-Inspect](https://github.com/evandycke/vba-inspect)
3. Installer OleTools

```bat
pip install -U oletools
```

4. Configurer l'analyse par le biais du fichier vba-inspect.ini

## Démarrage

La configuration de l'analyse s'effectue dans le fichier /config/vba-inspect.ini, dans la section DEFAULT. Vous indiquerez le dossier à analyser et le type de fichier à prendre en compte (*, *.xls, *.xlsx).

Pour réaliser un audit, il vous faudra exécuter le script vba-inspect.py

```bat
Python vba-inspect.py
```

Le script exposera dans le dossier /out/result.log le contenu VBA de chaque fichier analysé.
Les logs de l'analyse sont disponibles dans le dossier /log/vba-inspect.log

## Fabriqué avec

* [Python](https://www.python.org/downloads/windows/) - Langage de programmation
* [OleTools](http://www.decalage.info/python/oletools) - Outils développés en Python pour analyser des fichiers OLE et des fichiers Microsoft Office

## Contributing

Si vous souhaitez contribuer, lisez le fichier CONTRIBUTING.md pour savoir comment le faire.