import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
from dateutil.relativedelta import relativedelta
import xml.etree.ElementTree as ET
import xml.dom.minidom
import numpy as np

# -------------------------------------
# FUNCIONES AUXILIARES
# -------------------------------------

def convertir_a_formato_yyyy_mm_dd(fecha):
    try:
        return datetime.strptime(str(fecha), "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        try:
            return datetime.strptime(str(fecha), "%d-%m-%Y").strftime("%Y-%m-%d")
        except ValueError:
            try:
                return pd.to_datetime(fecha, errors='coerce').strftime("%Y-%m-%d")
            except:
                return ""

def calcular_fecha_aplicabilidad(fecha_final):
    hoy = datetime.today()
    dia_final = fecha_final.day
    mes = hoy.month + 1 if (hoy.day + 6 > dia_final) else hoy.month
    anio = hoy.year + 1 if mes == 1 and hoy.month == 12 else hoy.year
    return f"{anio}-{mes:02d}-{dia_final:02d}"

def format_xml(xml_str):
    return xml.dom.minidom.parseString(xml_str).toprettyxml()

# -------------------------------------
# PASO 1 - macro_excel
# -------------------------------------

def macro_excel(archivo, banco):
    col = ['IDENTIFICACION', 'NOMBRE COMPLETO', 'CODIGO MUN', 'MONTO INGRESOS', 'MONTO ACTIVOS',
           'VALOR DESEMBOLSADO', 'SALDO A CAPITAL DEL CREDITO', 'FECHA INICIAL DEL CREDITO', 'FECHA FINAL CREDITO',
           'AMORTIZACION', 'TASA FINAL', 'DIRECCION', 'TELEFONO', 'PAGARE', 'TIPO PRODUCTOR']
    
    df = pd.read_excel(archivo, names=col, header=0).fillna("")
    df['saldoCapitalCredito'] = np.ceil(df['SALDO A CAPITAL DEL CREDITO'].astype(float)).astype(int)
    df['VALOR DESEMBOLSADO'] = np.ceil(df['VALOR DESEMBOLSADO'].astype(float)).astype(int)
    df['fechaInicialEjecucion'] = pd.to_datetime(df['FECHA INICIAL DEL CREDITO'], errors='coerce')
    df['fechaFinalEjecucion'] = pd.to_datetime(df['FECHA FINAL CREDITO'], errors='coerce')
    df['plazoCredito'] = (df['fechaFinalEjecucion'] - df['fechaInicialEjecucion']).dt.days // 30
    print('plazo',df['plazoCredito'])
    df['valorCuota2'] = np.ceil(df['saldoCapitalCredito'] / df['plazoCredito']).astype(int)
    print('valor2',df['valorCuota2'])
    df['valorCuota1'] = np.ceil(df['saldoCapitalCredito'] - df['valorCuota2'] * (df['plazoCredito'] - 1)).astype(int)
    print('valor1',df['valorCuota1'])
    df['registro'] = '1'

    df['valorTotalCredito'] = df['saldoCapitalCredito'].astype(str)
    df['valorObligacion'] = df['saldoCapitalCredito'].astype(str)

    return df

# -------------------------------------
# PASO 2 - transformar_a_estructura_xml
# -------------------------------------


def transformar_a_estructura_xml(df, banco):
    df = df.copy()
    df['valorCuotaCapital'] = df['valorCuota1']
    df['fechaAplicacionHasta'] = df['fechaFinalEjecucion'].apply(lambda x: calcular_fecha_aplicabilidad(x) if pd.notnull(x) else "")

    prueba2 = pd.DataFrame(index=df.index)

    prueba2['tipoCartera'] = ["2"]*len(prueba2)
    prueba2['programaCredito'] = "732"
    prueba2['tipoOperacion'] = "1"
    prueba2['tipoMoneda'] = "1"
    prueba2['tipoAgrupamiento'] = "1"
    prueba2['numeroPagare'] = df['PAGARE'].astype(str)
    prueba2['numeroObligacionIntermediario'] = df['PAGARE'].astype(str)
    prueba2['fechaSuscripcion'] = df['FECHA INICIAL DEL CREDITO'].apply(convertir_a_formato_yyyy_mm_dd)
    prueba2['fechaDesembolso'] = df['FECHA INICIAL DEL CREDITO'].apply(convertir_a_formato_yyyy_mm_dd)
    prueba2['oficinaPagare'] = "100"
    prueba2['oficinaObligacion'] = "100"
    prueba2['codigo'] = "101059"
    prueba2['cantidad'] = "1"
    prueba2['correoElectronico'] = ""
    prueba2['cumpleCondicionesProductorAgrupacion'] = "true"
    prueba2['tipoAgrupacion'] = ""
    prueba2['tipoPersona'] = "1"
    prueba2['tipoProductor'] = df['TIPO PRODUCTOR'].astype(str)
    prueba2['actividadEconomica'] = ""
    prueba2['tipo'] = "2"
    prueba2['numeroIdentificacion'] = df['IDENTIFICACION'].astype(str)
    prueba2['digitoVerificacion'] = ""
    prueba2['primerApellido'] = ""
    prueba2['SegundoApellido'] = ""
    prueba2['PrimerNombre'] = ""
    prueba2['SegundoNombre'] = ""
    prueba2['Razonsocial'] = df['NOMBRE COMPLETO']
    prueba2['direccion'] = df['DIRECCION']
    prueba2['municipio'] = "11001"
    prueba2['prefijo'] = ""
    prueba2['numero'] = df['TELEFONO'].astype(str)
    prueba2['valor'] = df['MONTO ACTIVOS'].astype(str)
    prueba2['fechaCorte'] = "2021-12-30"
    prueba2['fechaInicialEjecucion'] = df['fechaInicialEjecucion'].dt.strftime('%Y-%m-%d')
    prueba2['fechaFinalEjecucion'] = df['fechaFinalEjecucion'].dt.strftime('%Y-%m-%d')
    prueba2['tipo2'] = "1"
    prueba2['municipio2'] = "11001"
    prueba2['direccion2'] = df['DIRECCION']
    prueba2['codigo2'] = "165000"
    prueba2['unidadesAFinanciar'] = "1"
    prueba2['costoInversion'] = df['VALOR DESEMBOLSADO'].astype(str)
    prueba2['valorAFinanciar'] = df['VALOR DESEMBOLSADO'].astype(str)
    prueba2['fechaVencimientoFinal'] = df['fechaFinalEjecucion'].dt.strftime('%Y-%m-%d')
    prueba2['plazoCredito'] = df['plazoCredito'].astype(str)
    prueba2['valorTotalCredito'] = df['saldoCapitalCredito'].astype(str)
    prueba2['porcentaje'] = "100"
    prueba2['valorObligacion'] = df['saldoCapitalCredito'].astype(str)
    prueba2['registro'] = "1"
    prueba2['fechaAplicacionHasta'] = df['fechaAplicacionHasta']
    prueba2['conceptoRegistroCuota'] = "K"
    prueba2['periodicidadIntereses'] = "PE"
    prueba2['periodicidadCapital'] = "PE"
    prueba2['tasaBaseBeneficiario'] = "5"
    prueba2['margenTasaBeneficiario'] = df['TASA FINAL'].astype(str)
    prueba2['valorCuotaCapital'] = df['valorCuota1'].astype(str)
    prueba2['porcentajeCapitalizacionIntereses'] = "0"
    prueba2['margenTasaRedescuento'] = "0"
    prueba2['valorIngresos'] = df['MONTO INGRESOS'].astype(str)
    prueba2['Anticipo'] = ""
    prueba2['valorTotalProyecto'] = ""
    prueba2['valorTotalAFinanciar'] = ""
    prueba2['plazoFinanciacion'] = df['plazoCredito'].astype(str)
    prueba2['numeroProyecto'] = ""
    prueba2['numeroDesembolso'] = ""
    prueba2['desembolsos'] = ""

    if banco != "Banco AV Villas":
        copia = prueba2.copy()
        copia['registro'] = "2"
        copia['valorCuotaCapital'] = df['valorCuota2'].astype(str)
        prueba2 = pd.concat([prueba2, copia], ignore_index=True)

    return prueba2


# -------------------------------------
# PASO 3 - procesar_excel (genera XML)
# -------------------------------------
def procesar_excel(tabla,num_operaciones,banco):

    vtotal = '0'
    ET.register_namespace('', "http://www.finagro.com.co/sit")
    Obligaciones = ET.Element('obligaciones', {'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
                                            'xmlns:xsd': 'http://www.w3.org/2001/XMLSchema',
                                            'cifraDeControl': str(num_operaciones),
                                            'cifraDeControlValor': vtotal,
                                            
                                            })

    #Obligaciones=ET.Element('{http://www.finagro.com.co/sit}obligaciones',cifraDeControlValor=vtotal,cifraDeControl=nur)
    #Obligaciones=ET.Element('{http://www.w3.org/2001/XMLSchema}obligaciones',cifraDeControlValor=vtotal,cifraDeControl=nur)
    num_operaciones= num_operaciones*2
    fechas_ok = 0
    fechas_no_concuerdan = 0

    for i in range(num_operaciones):

        tc=tabla.iloc[i,0]  #tipoCartera
        pc=tabla.iloc[i,1]  #programaCredito
        to=tabla.iloc[i,2]  #tipoOperacion
        tm=tabla.iloc[i,3]  #tipoMoneda
        ta=tabla.iloc[i,4]  #tipoAgrupamiento
        np=tabla.iloc[i,5]  #numeroPagare
        noi=tabla.iloc[i,6] #numeroObligacionIntermediario
        fs=tabla.iloc[i,7]  #fechaSuscripcion
        fd=tabla.iloc[i,8]  #fechaDesembolso
        op=tabla.iloc[i,9]  #oficinaPagare
        ofo=tabla.iloc[i,10]    #oficinaObligacion
        cod=tabla.iloc[i,11]    #codigo
        can=tabla.iloc[i,12]    #cantidad
        email=tabla.iloc[i,13]  #correoElectronico
        cpa=tabla.iloc[i,14]
        ta2=tabla.iloc[i,15]    #tipoAgrupacion
        tp=tabla.iloc[i,16]     #tipoPersona
        tpr=tabla.iloc[i,17]    #tipoProductor
        ae=tabla.iloc[i,18] #actividadEconomica
        tipo=tabla.iloc[i,19]   #tipo
        id=tabla.iloc[i,20] #numeroIdentificacion
        dv=tabla.iloc[i,21] #digitoVerificacion
        pa=tabla.iloc[i,22] #PrimerApellido
        sa=tabla.iloc[i,23] #SegundoApellido
        pn=tabla.iloc[i,24] #PrimerNombre
        sn=tabla.iloc[i,25] #SegundoNombre
        rs=tabla.iloc[i,26] #Razonsocial
        dir=tabla.iloc[i,27]    #direccion
        mun=tabla.iloc[i,28]    #municipio
        pref=tabla.iloc[i,29]   #prefijo
        num=tabla.iloc[i,30]    #numero
        va=tabla.iloc[i,31] #valor
        fc=tabla.iloc[i,32] #fechaCorte
        fie=tabla.iloc[i,33]    #fechaInicialEjecucion
        ffe=tabla.iloc[i,34]    #fechaFinalEjecucion
        tipo2=tabla.iloc[i,35]  #tipo2
        mun2=tabla.iloc[i,36]   #municipio3
        dir2=tabla.iloc[i,37]   #direccion4
        cod2=tabla.iloc[i,38]   #codigo5
        uf=tabla.iloc[i,39] #unidadesAFinanciar
        ci=tabla.iloc[i,40] #costoInversion
        vaf=tabla.iloc[i,41]    #valorAFinanciar
        fvf=tabla.iloc[i,42]    #fechaVencimientoFinal
        plazo=tabla.iloc[i,43]  #plazoCredito
        vtc=tabla.iloc[i,44]    #valorTotalCredito
        por=tabla.iloc[i,45]    #porcentaje
        vao=tabla.iloc[i,46]    #valorObligacion


        reg=tabla.iloc[i,47]    #registro
        fah=tabla.iloc[i,48]    #fechaAplicacionHasta
        crc=tabla.iloc[i,49]    #conceptoRegistroCuota
        pi=tabla.iloc[i,50] #periodicidadIntereses
        pec=tabla.iloc[i,51]    #periodicidadCapital
        tb=tabla.iloc[i,52] #tasaBaseBeneficiario
        mtb=tabla.iloc[i,53]    #margenTasaBeneficiario
        vcap=tabla.iloc[i,54]   #valorCuotaCapital
        pci=tabla.iloc[i,55]    #porcentajeCapitalizacionIntereses
        mtr=tabla.iloc[i,56]    #margenTasaRedescuento
        vai=tabla.iloc[i,57]
        #fci=tabla.iloc[i,58]
        pf=tabla.iloc[i,59] # La solicitud corresponde a  un proyecto financiado con varios desembolsos


        tc=str(tc)
        pc=str(pc)
        to=str(to)
        tm=str(tm)
        ta=str(ta)
        np=str(np)
        noi=str(noi)
        fs=str(fs)
        fd=str(fd)
        op=str(op)
        ofo=str(ofo)
        cod=str(cod)
        can=str(can)
        email=str(email)
        cpa=str(cpa)
        ta2=str(ta2)
        tp=str(tp)
        tpr=str(tpr)
        ae=str(ae)
        tipo=str(tipo)
        id=str(id)
        dv=str(dv)
        pn=str(pn)
        sn=str(sn)
        pa=str(pa)
        sa=str(sa)
        rs=str(rs)
        dir=str(dir)
        mun=str(mun)
        pref=str(pref)
        num=str(num)
        va=str(va)
        fc=str(fc)
        fie=str(fie)
        ffe=str(ffe)
        tipo2=str(tipo2)
        mun2=str(mun2)
        dir2=str(dir2)
        cod2=str(cod2)
        uf=str(uf)
        ci=str(ci)
        vaf=str(vaf)
        fvf=str(fvf)
        plazo=str(plazo)
        vtc=str(vtc)
        por=str(por)
        vao=str(vao)

        reg=str(reg)
        fah=str(fah)
        crc=str(crc)
        pi=str(pi)
        pec=str(pec)
        tb=str(tb)
        mtb=str(mtb)
        vcap=str(vcap)
        pci=str(pci)
        mtr=str(mtr)
        pf=str(pf)
        vai=str(vai)
        #fci=str(fci)
        print(tc)


        if reg == str(1):

            
            obligacion=ET.SubElement(Obligaciones,'{http://www.finagro.com.co/sit}obligacion',tipoCartera=tc,programaCredito=pc,tipoOperacion=to,tipoMoneda=tm,tipoAgrupamiento=ta,numeroPagare=np,numeroObligacionIntermediario=noi,fechaSuscripcion=fs,fechaDesembolso=fd)
            ET.SubElement(obligacion,'{http://www.finagro.com.co/sit}intermediario', oficinaPagare=op,oficinaObligacion=ofo,codigo=cod)
            beneficiarios=ET.SubElement(obligacion,'{http://www.finagro.com.co/sit}beneficiarios',cantidad=can)
            beneficiario=ET.SubElement(beneficiarios,'{http://www.finagro.com.co/sit}beneficiario', correoElectronico=email,cumpleCondicionesProductorAgrupacion="true",tipoAgrupacion=ta2,tipoPersona=tp,tipoProductor=tpr,actividadEconomica=ae)
            type(int(can))
            if dv == "":
                ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}identificacion', tipo=tipo, numeroIdentificacion=id)
                ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}nombre', primerNombre=pn, segundoNombre=sn, primerApellido=pa, segundoApellido=sa)

            else:
                ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}identificacion', tipo=tipo, numeroIdentificacion=id, digitoVerificacion=dv)
                ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}nombre', primerNombre=pn, segundoNombre=sn, primerApellido=pa, segundoApellido=sa,Razonsocial=rs)

            ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}direccionCorrespondencia', direccion=dir, municipio=mun)
            ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}numeroTelefono', prefijo=pref, numero=num)
            ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}valorActivos', valor=va, fechaCorte=fc, tipoDato="COP")
            ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}valorIngresos', valor=vai, fechaCorte='2024-12-31', tipoDato="COP")

            if pf == str(1): #(nuevo)
                proyecto=ET.SubElement(obligacion,'{http://www.finagro.com.co/sit}proyecto', fechaInicialEjecucion=fie, fechaFinalEjecucion=ffe)
                ET.SubElement(proyecto,'{http://www.finagro.com.co/sit}incentivo',inscripcionIncentivo="false")
                proyectosFinanciados=ET.SubElement(proyecto,'{http://www.finagro.com.co/sit}proyectosFinanciados',valorTotalProyecto="",valorTotalAFinanciar="",plazoFinanciacion="",numeroProyecto="",numeroDesembolso="99",desembolsos="")
                ET.SubElement(proyectosFinanciados,'{http://www.finagro.com.co/sit}destinosProyecto',codigoDestinoCredito="",valorAFinanciar="",unidadesAFinanciar="",costoInversion="")
                ET.SubElement(proyectosFinanciados,'{http://www.finagro.com.co/sit}municipios',municipio="")


            elif len(pf) == 0:
                proyecto=ET.SubElement(obligacion,'{http://www.finagro.com.co/sit}proyecto', fechaInicialEjecucion=fie, fechaFinalEjecucion=ffe)

            predios=ET.SubElement(obligacion,'{http://www.finagro.com.co/sit}predios')
            ET.SubElement(predios,'{http://www.finagro.com.co/sit}predio', tipo=tipo2, municipio=mun, direccion=dir)
            destinosCredito=ET.SubElement(obligacion,'{http://www.finagro.com.co/sit}destinosCredito')
            destinoCredito=ET.SubElement(destinosCredito,'{http://www.finagro.com.co/sit}destinoCredito', codigo=cod2, unidadesAFinanciar=uf, costoInversion=ci)
            destinoCreditoValorAFinanciar=ET.SubElement(destinoCredito,'{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar')
            ET.SubElement(destinoCreditoValorAFinanciar,'{http://www.finagro.com.co/sit}valorAFinanciar',xmlns="").text=vaf
            ET.SubElement(obligacion,'{http://www.finagro.com.co/sit}financiacion',fechaVencimientoFinal=fvf,plazoCredito=plazo,valorTotalCredito=vtc,porcentaje=por,valorObligacion=vao)
            planPagos=ET.SubElement(obligacion,'{http://www.finagro.com.co/sit}planPagos')
            ET.SubElement(planPagos,'{http://www.finagro.com.co/sit}registroCuota',registro=reg,fechaAplicacionHasta=fah,conceptoRegistroCuota=crc,periodicidadIntereses=pi,periodicidadCapital=pec,tasaBaseBeneficiario=tb,margenTasaBeneficiario=mtb,valorCuotaCapital=vcap,porcentajeCapitalizacionIntereses=pci,margenTasaRedescuento=mtr)

        elif reg == str(2) or reg == str(3):
            ET.SubElement(planPagos,'{http://www.finagro.com.co/sit}registroCuota',registro=reg,fechaAplicacionHasta=fah,conceptoRegistroCuota=crc,periodicidadIntereses=pi,periodicidadCapital=pec,tasaBaseBeneficiario=tb,margenTasaBeneficiario=mtb,valorCuotaCapital=vcap,porcentajeCapitalizacionIntereses=pci,margenTasaRedescuento=mtr)

        # elif >= 2:
        #     type(str(can))
        #     for j in range(len(can)):
        #         beneficiario=ET.SubElement(beneficiarios,'{http://www.finagro.com.co/sit}beneficiario', correoElectronico=email,cumpleCondicionesProductorAgrupacion="true", tipoAgrupacion=ta2, tipoPersona=tp, tipoProductor=tpr, actividadEconomica=ae)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}identificacion', tipo=tipo, numeroIdentificacion=id, digitoVerificacion=dv)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}nombre', primerNombre=pn, segundoNombre=sn, primerApellido=pa, segundoApellido=sa, Razonsocial=rs)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}direccionCorrespondencia', direccion=dir, municipio=mun)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}numeroTelefono', prefijo=pref, numero=num)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}valorActivos', valor=va, fechaCorte=fc)
        #
        if fah[:-2] ==  fvf[:-2]:
            fechas_ok += 1
        else:
            fechas_no_concuerdan +=1

        # elif >= 2:
        #     type(str(can))
        #     for j in range(len(can)):
        #         beneficiario=ET.SubElement(beneficiarios,'{http://www.finagro.com.co/sit}beneficiario', correoElectronico=email,cumpleCondicionesProductorAgrupacion="true", tipoAgrupacion=ta2, tipoPersona=tp, tipoProductor=tpr, actividadEconomica=ae)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}identificacion', tipo=tipo, numeroIdentificacion=id, digitoVerificacion=dv)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}nombre', primerNombre=pn, segundoNombre=sn, primerApellido=pa, segundoApellido=sa, Razonsocial=rs)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}direccionCorrespondencia', direccion=dir, municipio=mun)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}numeroTelefono', prefijo=pref, numero=num)
        #         ET.SubElement(beneficiario,'{http://www.finagro.com.co/sit}valorActivos', valor=va, fechaCorte=fc)
        #

    # xml = ET.tostring(Obligaciones)  # binary string
    # #my_arr=xml.decode()
    #
    # with open('output2.xml', 'w') as f:  # Write in XML file as utf-8
    #     f.write('<?xml version="1.0" encoding="UTF-8"?>' + xml.decode('utf-8'))

    # Generar el contenido del archivo XML y formatearlo
    xml_str = ET.tostring(Obligaciones, encoding='unicode')
    formatted_xml = format_xml('<?xml version="1.0" encoding="UTF-8"?>\n' + xml_str)
    xml_filename = datetime.today().strftime("%d%m%Y") + ".xml"
    
    # Crear el archivo ZIP en memoria
    zip_buffer = io.BytesIO()
    zip_name = datetime.today().strftime("%d-%m-%Y") + ".zip"  # Nombre del ZIP
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(xml_filename, formatted_xml)
    
    zip_buffer.seek(0)  # Mover el puntero al inicio
    return zip_buffer, zip_name

# -------------------------------------
# STREAMLIT INTERFAZ
# -------------------------------------

st.title("Generador de XML desde archivo Insumo")
st.write("Sube un archivo de Excel con información financiera para convertirlo a XML")

num_operaciones = st.number_input("Número de operaciones", min_value=1, step=1)
banco = st.selectbox("Selecciona el banco", ["Banco AV Villas", "Banco Caja social", "Banco Santander"])
archivo = st.file_uploader("Sube el archivo Excel", type=["xlsx"])


if archivo and num_operaciones > 0:
    with st.spinner("Procesando archivo..."):
        base = macro_excel(archivo, banco)
        prueba2 = transformar_a_estructura_xml(base, banco)
        st.subheader("Vista previa del DataFrame transformado (prueba2)")
        st.dataframe(prueba2)

        zip_buffer, zip_name = procesar_excel(prueba2, num_operaciones, banco)

    st.success("XML generado con éxito.")
    st.download_button("Descargar XML", zip_buffer, file_name=zip_name, mime="application/zip")
elif not archivo:
    st.warning("Por favor, sube un archivo Excel.")
elif num_operaciones <= 0:
    st.warning("Ingresa un número válido de operaciones.")
