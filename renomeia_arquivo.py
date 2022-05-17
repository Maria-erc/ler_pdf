# RENOMEIA ARQUIVO

# Script 1
import os

file_oldname = os.path.join(r'C:\pastas')
file_newname_newfile = os.path.join(r'C:\pastas')

os.rename(file_oldname, file_newname_newfile)

# Script 2
# Se trocar a extenção pdf para docx, funciona, mas o conteúdo do documento word não carrega.
import shutil
import os

file_oldname = os.path.join(r'C:\pastas.pdf')
file_newname_newfile = os.path.join(r'C:\pastas.docx')

newFileName=shutil.move(file_oldname, file_newname_newfile)

print ("The renamed file has the name:",newFileName)