::"C:\Program Files\Pandoc\pandoc.exe" README.md --pdf-engine=xelatex -V geometry:margin=1in -o README.pdf
pyinstaller --noconfirm --icon="GUI/icons/128x128.ico" -w Start.py -n SchoolReport_Excel2Word2PDF ^
        --add-data="img/*;img/." ^
        --add-data="README.md;." ^
        --add-data="README.pdf;." ^
        --add-data="Demo;Demo" ^
        --add-data="GUI/icons/*;GUI/icons/."
::pyinstaller --clean SchoolReport_Excel2Word2PDF.spec

