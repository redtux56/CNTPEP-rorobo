# Instalation sur debian 13
sudo apt install -y python3 python3-pip python3-venv python3-tk
## telecharger le deb de chrome ensuite apt install ./google_chrome.deb
sudo apt install -y libcairo2 libpango-1.0-0 libpangoft2-1.0-0 libgdk-pixbuf-2.0-0
sudo apt install -y fonts-dejavu-core fonts-liberation2
## cloner le repertoire
git clone https://github.com/redtux56/CNTPEP-rorobo
cd CNTPEP-rorobo
## creer venv et l'activer 
python3 -m venv .venv
source .venv/bin/activate
## installer pip et les dépendances
python -m pip install -U pip wheel
python -m pip install -r requirements.txt
## racourci




