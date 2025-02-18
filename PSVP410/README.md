# fcpCSV

## pacotes necessários

```sh
pip install pandas openpyxl pyinstaller

```

# testar por terminal

```sh
python .\PSVP410.py .\PDR410.csv
```

## gerar EXE

```sh
pyinstaller --onefile --clean --noconfirm --add-data "assets;assets" --icon=assets\icone.ico .\PSVP410.py
```
