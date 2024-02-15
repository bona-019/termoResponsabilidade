from docx import Document
from datetime import date
import pandas as pd
from time import sleep

df = pd.read_excel("template/BASE_TERMO_RESPONSABILIDADE.xlsx")

def cpf_checker(cpf):
    if len(str(cpf)) < 11:
        cpf = cpf.zfill(11)
        cpf_formatado = '{}.{}.{}-{}'.format(str(cpf)[:3], str(cpf)[3:6], str(cpf)[6:9], str(cpf)[9:])
        return cpf_formatado
    else:
        cpf_formatado = '{}.{}.{}-{}'.format(str(cpf)[:3], str(cpf)[3:6], str(cpf)[6:9], str(cpf)[9:])
        return cpf_formatado

def fill_doc(nome, cpf, aparelho, marca, modelo, codigo, estado, data):
    doc = Document("template/TEMPLATE_TERMO_RESPONSABILIDADE.docx")
    for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if "{NOME}" in run.text:
                    run.text = run.text.replace("{NOME}", nome)
                    run.bold = True

                if "{CPF}" in run.text:
                    run.text = run.text.replace("{CPF}", cpf_checker(cpf))
                    
                if "{APARELHO}" in run.text:
                    run.text = run.text.replace("{APARELHO}", aparelho)

                if "{MARCA}" in run.text:
                    run.text = run.text.replace("{MARCA}", marca)

                if "{MODELO}" in run.text:
                    run.text = run.text.replace("{MODELO}", modelo)

                if "{CODIGO}" in run.text:
                    run.text = run.text.replace("{CODIGO}", codigo)
                    run.bold = True

                if "{ESTADO}" in run.text:
                    run.text = run.text.replace("{ESTADO}", estado)

                if "{DATA}" in run.text:
                    run.text = run.text.replace("{DATA}", data)

                elif "{NOME}" in run.text:
                    run.text = run.text.replace("{NOME}", nome)
                    run.bold = True

    doc.save(f"exportados/termo - {nome}.docx")
    fill_doc(nome=nome, cpf=cpf, aparelho=aparelho, marca=marca, modelo=modelo, codigo=codigo, estado=estado, data=data)

if __name__ == "__main__":
    for index, row in df.iterrows():
        fill_doc(
            nome = row['NOME'],
            cpf = row['CPF'],
            aparelho = row['APARELHO'],
            marca = row['MARCA'],
            modelo = row['MODELO'],
            codigo = row['CODIGO'],
            estado = row['ESTADO'],
            data = date.today().strftime("%d/%m/%Y")
        )