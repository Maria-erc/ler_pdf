# PDFMINER

#https://pdfminersix.readthedocs.io/en/latest/tutorial/composable.html?msclkid=227e73fcaec511ec9397e5e6f525d3ef
from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
import pandas as pd
import os


# antes do loop
# criação de dataframe para concatenar as informações dentro dele
df_colunas = [
    'Arquivo',
    'Nome',
    'CPF/CNPJ',
    'Torre_Apart_Cota'
]

# dataframe de contrato padrão 1
df_contrato_p1 = pd.DataFrame(columns=df_colunas)

# dataframe de contrato de padrão 2
df_contrato_p2 = pd.DataFrame(columns=df_colunas)

# dataframe de contrato de nem padrão 1 nem de padrão 2
df_contrato_exc = pd.DataFrame(columns=df_colunas)

# dicionario erro
df_erro = pd.DataFrame(columns=df_colunas)


# ler cada um dos arquivos da pasta
# https://www.geeksforgeeks.org/os-module-python-examples/?msclkid=8d257d4caf6011ec8e535be26e480972
# https://www.bing.com/videos/search?q=keith+galli+pandas&&view=detail&mid=79A2939426D67640E9BF79A2939426D67640E9BF&&FORM=VRDGAR&ru=%2Fvideos%2Fsearch%3Fq%3Dkeith%2520galli%2520pandas%26qs%3Dn%26form%3DQBVDMH%26%3D%2525eManage%2520Your%2520Search%2520History%2525E%26sp%3D-1%26pq%3Dkeith%2520galli%2520pandas%26sc%3D3-18%26sk%3D%26cvid%3D880C9399048D44A9A61AA117C0830186
# https://www.delftstack.com/pt/howto/python/python-open-all-files-in-directory/?msclkid=e89fbebcaf5f11ecbe22883f88cec848
#arquivos = [arquivo for arquivo in os.listdir(r'\\diretorio')]
#arquivos = arquivos[0:11]
#diretorio = r'\\diretorio\\'
arquivos = [arquivo for arquivo in os.listdir(r'C:\pastas')]
arquivos = arquivos[0:55]
diretorio = r'C:\pastas\\'
print('Carregando...')
contador = 0


for arquivo in arquivos:
    doc_pdf = diretorio + arquivo
    def le_contrato(documento_pdf):
        '''Lê contrato em pdf'''
        global output_string
        output_string = StringIO()
        with open(doc_pdf, 'rb') as in_file:
            parser = PDFParser(in_file)
            doc = PDFDocument(parser)
            rsrcmgr = PDFResourceManager()
            device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            for page in PDFPage.create_pages(doc):
                interpreter.process_page(page)

    le_contrato(doc_pdf)
    contador += 1
    print('=-=-', contador)
    # armazena o texto
    texto = output_string.getvalue()
    try:
         # criação de dicionário
        dicionario_dados_pdf_excecao = {
            'Arquivo' : [],
            'Nome' : [],
            'palavrachave1' : [],
            'conjuntopalavrachave2' : [],
        }

        if 'Titulo do pdf' in texto:
            texto_lista = texto.split('\n')

            # extrair nome
            for elemento in texto_lista:
                if 'Nome: ' in elemento:
                    nome = elemento[6::]
                    nome = nome.split('(')
                    nome = nome[0]
                    break

            # extrair palavrachave1
            for elemento in texto_lista:
                if 'palavrachave1: ' in elemento:
                    palavrachave1 = elemento[5::]
                    break

            # extrair palavrachave2.1
            for elemento in texto_lista:
                if 'palavrachave2.1: ' in elemento:
                    palavrachave2_1 = elemento[7::]
                    break

            #extrair palavrachave2.2
            for elemento in texto_lista:
                if 'palavrachave2.2: ' in elemento:
                    palavrachave2_2 = elemento[13::]
                    break

            # extrair palavrachave2.3
            for elemento in texto_lista:
                if 'palavrachave2.3: ' in elemento:
                    palavrachave2_3 = elemento[16::]
                    break

            conjuntopalavrachave2 = palavrachave2_1 + '-' + palavrachave2_2 + '-' + palavrachave2_3


            dicionario_dados_pdf_excecao['Nome'] = nome
            dicionario_dados_pdf_excecao['palavrachave1'] = palavrachave1
            dicionario_dados_pdf_excecao['conjuntopalavrachave2'] = conjuntopalavrachave2

        elif 'palavraA' and 'palavraB'and 'palavraC' in texto:
            texto_lista = texto.split('\n')

            # extrair nome
            for elemento in texto_lista:
                if 'Nome: ' in elemento:
                    nome = elemento[6::]
                    nome = nome.split('(')
                    nome = nome[0]
                    print(nome)
                    break

            # extrair palavrachave1
            for elemento in texto_lista:
                if 'palavrachave1:' in elemento:
                    palavrachave1 = elemento[5:19]
                    break
                else:
                    palavrachave1 = ''

            # extrair palavrachave2.1
            for elemento in texto_lista:
                if 'palavrachave2.1:' in elemento:
                    palavrachave2_1 = elemento[6::]
                    print('palavrachave2.1:', palavrachave2_1)
                    break
                else:
                    palavrachave2_1 = ''

            # extrair palavrachave2.3
            for elemento in texto_lista:
                if 'palavrachave2.3:' in elemento:
                    palavrachave2_3_index = texto_lista.index(elemento)
                    palavrachave2_3  = texto_lista[palavrachave2_3_index+1]
                    print('palavrachave2.3:', palavrachave2_3)
                    break
                else:
                    palavrachave2_3 = ''


            for elemento in texto_lista:
                if 'palavrachave2.2' in elemento:
                    if ':' in elemento:
                        if '.' in elemento:
                            palavrachave2_2 = elemento
                            print('palavrachave2.2', palavrachave2_2)
                            break
                else:
                    palavrachave2_2 = ''
            conjuntopalavrachave2 = palavrachave2_1 + '-' + palavrachave2_3 + '-' + palavrachave2_2


            dicionario_dados_pdf_excecao['Nome'] = nome
            dicionario_dados_pdf_excecao['palavrachave1'] = palavrachave1
            dicionario_dados_pdf_excecao['conjuntopalavrachave2'] = conjuntopalavrachave2
                            
        
        # se não achar nome nem palavrachave1
        else: 
            dicionario_dados_pdf_excecao['Nome'] = 'ne'
            dicionario_dados_pdf_excecao['palavrachave1'] = 'ne'
            dicionario_dados_pdf_excecao['conjuntopalavrachave2'] = 'ne'

        # coloca informações extraídas do texto no dicionário
        dicionario_dados_pdf_excecao['Arquivo'] = arquivo

        df_dict_exc = pd.DataFrame([dicionario_dados_pdf_excecao])

        df_contrato_exc = pd.concat([df_contrato_exc,df_dict_exc], ignore_index=True)

        print('arquivo sem padrão = ', arquivo)
            
    # se der erro por "list out of range"
    except IndexError as e:
        print(e, 'arquivo = ', arquivo)

        dicionario_erro = {
            'Arquivo' : [],
            'Nome' : [],
            'palavrachave1' : [],
            'conjuntopalavrachave2' : [],
        }

        dicionario_erro['Arquivo'] = arquivo
        dicionario_erro['Nome'] = 'e'
        dicionario_erro['palavrachave1'] = 'e'
        dicionario_erro['conjuntopalavrachave2'] = 'e'

        df_dict_erro = pd.DataFrame([dicionario_erro])

        df_erro = pd.concat([df_erro, df_dict_erro])


    else:
            # concatena os dataframes 
            #https://pythonexamples.org/pandas-concatenate-dataframes/?msclkid=4110b667af5c11ecbec6a915e00be673
            df_completo = pd.concat([df_contrato_p1, df_contrato_p2,df_contrato_exc, df_erro], ignore_index=True)
    
# exporta para planilha em excel
df_completo.to_excel(r'C:pasta.xlsx', index=False)

print('Dataframe transferido para Excel.')

print('-----------------------')
print('Tabela resumo de documentos:')
print('\ndf_completo:')
print(df_completo)