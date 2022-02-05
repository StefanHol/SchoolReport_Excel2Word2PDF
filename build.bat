pyinstaller --noconfirm -w Start.py -n SchoolReport_Excel2Word2PDF ^
        --add-data="img/*;img/." ^
        --add-data="SchoolReport Excel2Word2PDF.pdf;." ^
        --add-data="README.md;." ^
        --add-data="Demo;Demo" ^
        --add-data="GUI/icons/*;GUI/icons/."
::pyinstaller --clean SchoolReport_Excel2Word2PDF.spec