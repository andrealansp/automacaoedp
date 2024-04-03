import os
from datetime import datetime

diretorio_raiz = os.path.join(os.getcwd(), str(datetime.today().year))
if not os.path.exists(diretorio_raiz):
    os.mkdir(diretorio_raiz)
    os.chdir(diretorio_raiz)
    for i in range(1, 13):
        os.mkdir(str(i))
