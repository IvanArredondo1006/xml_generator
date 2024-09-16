import xml.etree.ElementTree as ET
import pandas as pd
import numpy as np
import xml.dom.minidom


col = ['tipoNovedadPago','codigoMotivoAbono','destinoAbono','fechaAplicacionPago','tipoCarteraId',
       'codigoIntermediario','numeroObligacion','tipoMonedaId','valorAbonoCapital','tipoDocumentoId',
       'numeroDocumento']

df = pd.read_excel(r'M:\Bancos\Banco Caja Social\Otros\ABONOS\Insumo_robot.xlsx',sheet_name='Hoja1',names=col,index_col=False)


nur =  input('Cuantas operaciones vas a cargar: ')
nur=str(nur)
Ni = len(df['tipoNovedadPago'])

ET.register_namespace('', "http://www.finagro.com.co/sit")
ET.register_namespace('xsi', "http://www.w3.org/2001/XMLSchema-instance")
ET.register_namespace('xsd', "http://www.w3.org/2001/XMLSchema")


df['fechaAplicacionPago'] = df['fechaAplicacionPago'].astype(str)
# Convertir la fecha al formato deseado
df['fechaAplicacionPago'] = pd.to_datetime(df['fechaAplicacionPago'], format='%d%m%Y').dt.strftime('%Y-%m-%d')


abonos = ET.Element('{http://www.finagro.com.co/sit}abonos',cifraDeControl = nur)
for i in range(Ni):
    tnp=df.iloc[i,0]
    cma=df.iloc[i,1]
    da=df.iloc[i,2]
    fap=df.iloc[i,3]
    tcid=df.iloc[i,4]
    ci=df.iloc[i,5]
    no=df.iloc[i,6]
    tmid=df.iloc[i,7]
    vac=df.iloc[i,8]
    tdid=df.iloc[i,9]
    nd=df.iloc[i,10]
    

    tnp=str(tnp)
    cma=str(cma)
    da=str(da)
    fap=str(fap)
    tcid=str(tcid)
    ci=str(ci)
    no=str(no)
    tmid=str(tmid)
    vac=str(vac)
    tdid=str(tdid)
    nd=str(nd)

    # if cma == "NEGOCIACIONES ESPECIALES":
    #     cma == str(11)

    abono = ET.SubElement(abonos,'{http://www.finagro.com.co/sit}abono',tipoNovedadPago = tnp,codigoMotivoAbono = cma,fechaAplicacionPago=fap)
    informacionObligacion = ET.SubElement(abono,'{http://www.finagro.com.co/sit}informacionObligacion',tipoCarteraId=tcid,codigoIntermediario=ci,numeroObligacion=no,tipoMonedaId=tmid)
    ET.SubElement(informacionObligacion,'{http://www.finagro.com.co/sit}informacionBeneficiario',tipoDocumentoId=tdid,numeroDocumento=nd)
    valorAbono = ET.SubElement(abono,'{http://www.finagro.com.co/sit}valorAbono')
    ET.SubElement(valorAbono,'{http://www.finagro.com.co/sit}valorAbonoCapital',xmlns="").text=vac
    

xml_str = ET.tostring(abonos, encoding='unicode')

# Utilizar minidom para formatear la cadena XML con indentaciones
dom = xml.dom.minidom.parseString(xml_str)
pretty_xml_as_string = dom.toprettyxml()

# Escribir el XML formateado en un archivo
with open(r'M:\Bancos\Banco Caja Social\Otros\ABONOS\PRUEBA 24-11-23\Abonos.xml', 'w', encoding='utf-8') as f:
    f.write(pretty_xml_as_string)
