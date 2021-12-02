#Importando Libs;

from reportlab.pdfgen   import canvas
from reportlab.lib.units import inch, cm
from reportlab.lib.pagesizes import landscape,A4
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Paragraph
from reportlab.pdfbase.ttfonts  import TTFont
from reportlab.lib  import colors
from reportlab.pdfbase  import pdfmetrics
from reportlab.platypus import Paragraph
from datetime import date
import numpy as np


#Lib de Referência;
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Image, Paragraph, Frame, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

#Importando o Pandas;

import pandas as pd
import matplotlib.pyplot as plt


SEXO_CONSULTOR  = {
   'F': 'pela consultora',
   'M': 'pelo consultor'
}



MES = {
        '01':'Janeiro',
        '02':'Fevereiro',
        '03':'Março',
        '04':'Abril',
        '05':'Maio',
        '06':'Junho',
        '07':'Julho',
        '08':'Agosto',
        '09':'Setembro',
        '10':'Outubro',
        '11':'Novembro',
        '12':'Dezembro',
}

    
#Gerando PDF;

def gerar_pdf(nomepdf = False, data=False, aluno=False, instrutor=False, treinamento=False, sexo=False, carga_horaria=False):
    # Criando o PDF;
   

    pdf = canvas.Canvas(f'PDF/{nomepdf}.pdf')
    pdf.drawImage('Fundo.png', 0, 0, 29.7*cm, 21*cm) 
    pdf.setFont('Helvetica-Bold', 24)
    pdf.setFillColor(colors.white)


    #Variavéis;


    #data_conclusao = 'xx/xx/xxxx'


    #pdf.setFillColor('blue')

    styles = getSampleStyleSheet()
    estilo = ParagraphStyle(name='Normal_CENTER',
                              parent=styles['Normal'],
                              fontName='Helvetica',
                              wordWrap='LTR',
                              alignment=TA_LEFT,
                              fontSize=15,
                              leading=23,
                              textColor=colors.HexColor('#FFFFFF'),
                              borderPadding=0,
                              leftIndent=0,
                              rightIndent=0,
                              spaceAfter=0,
                              spaceBefore=0,
                              splitLongWords=True,
                              spaceShrinkage=0.05,
                            )




    nome = Paragraph(f'<font size=22 color=#FFFFFF>{aluno}</font> ')

    nome.wrapOn(pdf, 500, 150) #Quebra - Não mexe;
    nome.drawOn(pdf, 275, 400) #Posição (o y é altura)
    
    descricao = f' Certifico que participou do treinamento de \
                    <b>{treinamento}</b>, ministrado \
                       {sexo} \
                    <b>{instrutor}</b>, com carga horária de \
                    <b>{carga_horaria} h/a</b>.'
    
    p = Paragraph(descricao, estilo)

    p.wrapOn(pdf, 500, 200) #Quebra
    p.drawOn(pdf, 200, 280) #Posição



    # Data atual;


    d = Paragraph(f'<font size=14 color = #FFFFFF>{data}</font>')

    d.wrapOn(pdf, 355, 200) #Quebra
    d.drawOn(pdf, 300, 205) #Posição


    #Configurando Paisagem;
    canvas.Canvas.setPageSize(pdf, (landscape(A4)))


    #Salvando PDF;
    pdf.save()

#Leitura de Base;

lista = pd.read_excel('Lista.xlsx')
listadf = pd.DataFrame(lista)
listafim = listadf[['CPF','Nome','Email','Consultora','Carga','Treinamento','Sexo','Data']]


nome = listadf['Nome']
email = listadf['Email']
Consultora = listadf['Consultora']
Carga = listadf['Carga']
Treinamentos = listadf['Treinamento']
Sexo = listadf['Sexo']
Data = listadf['Data']
Cpf_aluno = listadf['CPF']


# Gerando o PDF;

import win32com.client as win32
from win32 import win32api

for index, row in listadf.iterrows():
    
   ano = row['Data'].strftime('%Y')
   mes = MES[row['Data'].strftime('%m')]
   dia = row['Data'].strftime('%d')
   sexo_consultor = SEXO_CONSULTOR[row['Sexo']]
   data = f'São Paulo, {dia} de {mes} de {ano}'
   gerar_pdf(row['CPF'], data, row['Nome'], row['Consultora'],row['Treinamento'], sexo_consultor, row['Carga'])


#Disparo - Email;
'''
   outlook = win32.Dispatch('outlook.application')
   email_chave = outlook.CreateItem(0)
   email_chave.To = row['Email'] 
   email_chave.Subject = "Certificado"
   email_chave.HTMLBody = f"""
        <p>Olá Colaborador(a), segue em anexo seu certificado de conclusão.</p>


        """
   caminho_arquivo = 'C:/Users/ln00925/Desktop/RH_CERTIFICADOS/PDF/' + str(row['CPF']) + '.pdf'
   email_chave.Attachments.Add(caminho_arquivo)
   email_chave.Send()
   print('Anexo Enviado')

'''
