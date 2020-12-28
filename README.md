# File Converter ![APM](https://img.shields.io/apm/l/vim-mode?color=red&logo=xcv&logoColor=xcv) ![GitHub Pipenv locked Python version](https://img.shields.io/github/pipenv/locked/python-version/metabolize/rq-dashboard-on-heroku?color=gree) ![Conda](https://img.shields.io/conda/pn/conda-forge/python)
Proje **Python PyQT** kütüphanesini öğrenmek ve öğrencilik hayatında birçok ödevle uğraşırken **dosya dönüştürücülerine** ihtiyaç duyduğumuzdan dolayı hızlı bir şekilde elimin altında ücretsiz bir uygulamanın bulunması amaçlarıyla oluşturulmuştur. Projeye destek vermek isterseniz todo'lardaki yapılacaklara bakarak contributer olarak projeyi geliştirebilirsiniz. Belge dönüşümleri için kullandığım dosyaları ayrı python dosyalarında yazdım bunları kendi projelerinizde kullanabilirsiniz:
* [DOCX to PDF](https://github.com/enesonmez/File-Converter/blob/master/docx2pdf.py)
* [PPTX to PDF](https://github.com/enesonmez/File-Converter/blob/master/pptx2pdf.py)
* [TXT to PDF](https://github.com/enesonmez/File-Converter/blob/master/txt2pdf.py)
* [XLSX to PDF](https://github.com/enesonmez/File-Converter/blob/master/xlsx2pdf.py)
* [PDF Merge](https://github.com/enesonmez/File-Converter/blob/master/pdfmerge.py)

## Uygulama Neleri İçeriyor:question:
* DOCX (word belgesi) => PDF Dönüştürücü
* PPTX (powerpoint belgesi) => PDF Dönüştürücü
* TXT (metin belgesi) => PDF Dönüştürücü
* XLSX (excel belgesi) => PDF Dönüştürücü
* PDF Birleştirici

## ToDo's :scroll:
- [ ] PDF => DOCX Dönüştürücü
- [ ] PPTX => DOXC Dönüştürücü
- [ ] PDF => PPTX Dönüştürücü
- [ ] DOCX => PPTX Dönüştürücü

## Nasıl Kurulur?:computer:
1. **Linux için :** 
   - Projeyi kurmak istediğiniz dizine gidin ve terminali açın.
   - sudo apt-get install python3
   - sudo apt-get install python3-venv
   - python3 -m venv file-converter
   - cd file-converter
   - **git clone https://github.com/enesonmez/File-Converter.git**
   - source file-converter/bin/activate
   - pip install -r requirements.txt
   - python file_converter.py
   - Uygulamayı kapattıktan sonra terminalde **deactivate** yazın.
2. **Windows için :**
   - Bilgisayarınızda python yoksa **https://www.python.org/downloads/** giderek python indirin ve kurun.
   - Projeyi kurmak istediğiniz dizine gidin ve cmd'yi açın.
   - python get-pip.py
   - pip install virtualenv
   - **git clone https://github.com/enesonmez/File-Converter.git**
   - cd File-Converter
   - virtualenv -p python3 venv
   - venv\Scripts\activate
   - pip install -r requirements.txt
   - python file_converter.py
   - Uygulamayı kapattıktan sonra terminalde **venv\Scripts\deactivate** yazın.
