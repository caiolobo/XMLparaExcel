import xmltodict
import os
import pandas as pd
import json

def pegar_infos(nome_arquivo, valores):
    # print(f"Tudo Certo! {nome_arquivo}")
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml:

        dic_arquivo = xmltodict.parse(arquivo_xml)

        # try: 
            # esse try é para mostrar em forma de json o erro
        if "Nfe" in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nf = dic_arquivo['nfeProc']["NFe"]["infNFe"]

        

        infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        numero_nota = infos_nf["@Id"]
        empresa_emissora = infos_nf["emit"]["xNome"]
        


        if "dest" in dic_arquivo["nfeProc"]["NFe"]["infNFe"]:
            nome_cliente = infos_nf["dest"]["xNome"]
        else:    
            nome_cliente = "Cliente não cadastrado!"

        if "dest" in dic_arquivo["nfeProc"]["NFe"]["infNFe"]:
            endereco = infos_nf["dest"]["enderDest"]
        else:
            endereco = "Endereço não cadastrado!"
        # if "protNFe" in dic_arquivo["nfeProc"]["protNFe"]:
        protocolo_nfe = dic_arquivo['nfeProc']["protNFe"]["infProt"]["nProt"]
        # else:
        #     protocolo_nfe = "Nota sem Protocolo de Autorização"

        peso = infos_nf["transp"]["vol"]
        valores.append([numero_nota, protocolo_nfe, empresa_emissora, nome_cliente, endereco, peso])
        # except Exception as e:
        #     print(e)
        #     print(json.dumps(dic_arquivo, indent=4))
            


        
lista_arquivos = os.listdir("nfs")

colunas = ["numero_nota", "protocolo_nfe", "empresa_emissora", "nome_cliente", "endereco","peso"]
valores = []

for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)
    


    

tabela = pd.DataFrame(columns=colunas, data=valores)

tabela.to_excel("NotasFiscais.xlsx", index=False)
