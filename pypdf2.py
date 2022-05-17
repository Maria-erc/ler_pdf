# PYPDF2

# Instala pacotes
# Promp anaconda: conda install -c conda-forge pypdf2
import PyPDF2
import re

arquivo_pdf = open('contrato.pdf','rb')

ler_pdf = PyPDF2.PdfFileReader(arquivo_pdf)

numero_de_paginas = ler_pdf.getNumPages()

pagina = ler_pdf.getPage(3)

conteudo_pagina = pagina.extractText().encode('utf-8').decode('utf-8')

parsed = ''.join(conteudo_pagina)

print('Sem eliminar quebras')
print(parsed)

parsed = re.sub(',', '', parsed)
parsed2 = re.sub('.', '', parsed)
print('Após eliminar as quebras')
print(parsed)

''''
count ={}
for word in parsed: # palavras a quantificar no texto
        #corpus = arquivo_pdf.split(' ')
        w_strp = word.strip() # retirar quebras de linha
        if w_strp != '' and w_strp not in count: # se ja a adicionamos nao vale a pena faze-lo outra vez
            count[w_strp] = parsed.count(w_strp)
print(count) # {'mas': 2, 'é': 2, 'luz': 4}
'''

# Contar a frequência de cada palavra
palavras = parsed.split()
palavras
contagem = {}
for palavra in palavras:
    count_palavra = palavras.count(palavra)
    contagem[palavra] = count_palavra
print(contagem)

# Ordenar a frequência de palavras
contagem_ordenada = {k: v for k, v in sorted(contagem.items(), key=lambda item: item[1], reverse=True)}
print(contagem_ordenada)
print(sorted(contagem.items(), key=lambda item: item[1], reverse=True))