import sys #Sirve para definir el directorio donde se encuentran las librerias
sys.path.append('C:\\Users\\Certeza\\Lib\\site-packages') #Directorio donde se encuentran las librerias de PIP

import os
import pandas as pd
#from unidecode import unidecode

directorio = '/Users/leonardorebollarruelas/Documents/GitHub/LeonardoRR95/Generador de Importadores'

os.chdir(directorio)

tabla = pd.read_excel('IMSS ENE - DIC 24.xlsx', sheet_name= 'CONCENTRADO - ENERO- DIC 2024')
tabla['NO DE POLIZA '] = tabla['NO DE POLIZA '].astype(str)
tabla['NO DE POLIZA '] = tabla['NO DE POLIZA '].str.replace(' ', '')
tabla['NO DE POLIZA '] = tabla['NO DE POLIZA '].str.replace(r'(-|ENCOMPAÑIA|POLIZAVIGENTE|PÓLIZAVIGENTE|ENESPERADEEMISION|PENDIENTEDECANCELACION)', '', regex=True)
tabla['NO DE POLIZA '] = tabla['NO DE POLIZA '].str.replace(r'(COMPAÑIA|CONOTROAGENTE|COTIZACIONSINVIGENCIA|PENDIENTERESPUESTADEEMISION|PENDIENTERESPUESTADETRAMITADOR)', '', regex=True)
tabla['NO DE POLIZA '] = tabla['NO DE POLIZA '].str.replace(r'(PENDIENTERESPUESTATRAMITADOR|POLIZADERENOVACIONCONVIGENCIADEDICIEMBRE|POLIZAMULTIANUALCONFALTADEPAGO|PENDIENTERESPUESTADEEMISION|RESPUESTAPENDIENTE)', '', regex=True)
tabla['NO DE POLIZA '] = tabla['NO DE POLIZA '].str.replace(r'(RESPUESTAPENDIENTEDETRAMITADOR|RESPUESTAPENDIENTEDETRAMITADORSIESAMPLIAOAMPLIAMAS|RESPUESTAPENDIENTEEEMISION|SERIEBLOQUEADASEMANDAACOMPAÑIA|TRAMITEDXNAUNNOSEPUEDEREALIZAREMISION)', '', regex=True)
tabla['NO DE POLIZA '] = tabla['NO DE POLIZA '].str.replace(r'(DETRAMITADOR|DETRAMITADORSIESAMPLIAOAMPLIAMAS|EEMISION|PENDIENTERESPUETADETRAMITADOR|PÓLIZAYFALTADEPAGO|SERIEBLOQUEADASEMANDAA|SIESAMPLIAOAMPLIAMAS)', '', regex=True)
tabla = tabla[tabla['NO DE POLIZA '] != '']
tabla = tabla.drop_duplicates(subset = 'NO DE POLIZA ')
tabla = tabla.reset_index(drop = True)
tabla['ASEGURADORA'] = tabla['ASEGURADORA'].str.replace(' ', '') 
tabla['CLAVE DE PROMOTOR'] = tabla['CLAVE DE PROMOTOR'].astype(str)
tabla['AGENTE PRODUCTOR ASOCIADO'] = tabla['AGENTE PRODUCTOR ASOCIADO'].astype(str)
tabla['AGENTE PRODUCTOR ASOCIADO'] = tabla['AGENTE PRODUCTOR ASOCIADO'].str.rstrip()

importador = pd.read_excel('SICAS - PolVehiculos_PRIV.xlsx', sheet_name= '02Pólizas de vehiculos')

tab_names = tabla.columns
imp_names = list(importador.columns)

imp = pd.DataFrame(columns = imp_names)

tabla_imp = pd.read_excel('IMSS ENE - DIC 24.xlsx', sheet_name= 'IMPORTADOR 2024')
tabla_imp = tabla_imp.astype(str)

tab_imp_nam = tabla_imp.columns

tabla_imp['Documento = No Poliza'] = tabla_imp['Documento = No Poliza'].str.replace(' ', '')
tabla_imp['Documento = No Poliza'] = tabla_imp['Documento = No Poliza'].str.replace(r'(-|ENCOMPAÑIA|POLIZAVIGENTE|PÓLIZAVIGENTE|ENESPERADEEMISION|PENDIENTEDECANCELACION)', '', regex=True)
tabla_imp['Documento = No Poliza'] = tabla_imp['Documento = No Poliza'].str.replace(r'(COMPAÑIA|CONOTROAGENTE|COTIZACIONSINVIGENCIA|PENDIENTERESPUESTADEEMISION|PENDIENTERESPUESTADETRAMITADOR)', '', regex=True)
tabla_imp['Documento = No Poliza'] = tabla_imp['Documento = No Poliza'].str.replace(r'(PENDIENTERESPUESTATRAMITADOR|POLIZADERENOVACIONCONVIGENCIADEDICIEMBRE|POLIZAMULTIANUALCONFALTADEPAGO|PENDIENTERESPUESTADEEMISION|RESPUESTAPENDIENTE)', '', regex=True)
tabla_imp['Documento = No Poliza'] = tabla_imp['Documento = No Poliza'].str.replace(r'(RESPUESTAPENDIENTEDETRAMITADOR|RESPUESTAPENDIENTEDETRAMITADORSIESAMPLIAOAMPLIAMAS|RESPUESTAPENDIENTEEEMISION|SERIEBLOQUEADASEMANDAACOMPAÑIA|TRAMITEDXNAUNNOSEPUEDEREALIZAREMISION)', '', regex=True)
tabla_imp['Documento = No Poliza'] = tabla_imp['Documento = No Poliza'].str.replace(r'(DETRAMITADOR|DETRAMITADORSIESAMPLIAOAMPLIAMAS|EEMISION|PENDIENTERESPUETADETRAMITADOR|PÓLIZAYFALTADEPAGO|SERIEBLOQUEADASEMANDAA|SIESAMPLIAOAMPLIAMAS)', '', regex=True)
tabla_imp = tabla_imp.drop_duplicates(subset = 'Documento = No Poliza')
tabla_imp = tabla_imp[tabla_imp['Documento = No Poliza'] != '']
tabla_imp = tabla_imp.reset_index(drop = True)

cosas = pd.read_excel('IMSS ENE - DIC 24.xlsx', sheet_name= 'Tablas a no borrar')

asesores = pd.read_excel('Catálogo 27.06.2024.xlsx')
asesores = asesores.astype(str)
asesores.iloc[:, 6] = asesores.iloc[:, 6].str.lower()

raul_f = ['', 'RAUL FLORES ESTRADA', '', 'oceaniasecc33@gmail.com', 'Jessica', '', 'nan']
#geor_r = ['', 'MARGARITA LEBRUN ALEJANDRO', '', '', 'Vania', '', 'nan']

asesores.loc[43] = raul_f 
#asesores.loc[44] = geor_r 

# =============================================================================
# Herramientas
# =============================================================================
    # Genera la tabla con las claves de agentes 

cve_a = cosas[['ASEGURADORAS', 'CONCEPTO', 'CLAVE']].loc[0:8]
cve_a['CLAVE'].loc[8] = 'A10504'
cve_a.loc[9] = ['CHUBB', '', '297009'] #CLAVES DE ASEGURADORAS

    # delegaciones
    
delegaciones = pd.DataFrame({'jubilado' : cosas['JUBILADO'],
                             'activo' : cosas['ACTIVO']})
delegaciones['ID'] = delegaciones['jubilado'].str[0:2]
delegaciones = delegaciones[0:38]

    # Formas de Pago
    
f_pago = ['Activo', 'Contado', 'Jubilado', 'Domiciliado']


# =============================================================================
# # =============================================================================
# # Genera la fecha de nacimiento 
# # =============================================================================
# =============================================================================

f_nac = pd.DataFrame({'Nombre' : tabla['ASEGURADO']})
f_nac['rfc'] = tabla['RFC']
f_nac['rfc'] = f_nac['rfc'].str.replace(' ', '')
f_nac['f_nac'] = f_nac['rfc'].str[4:10]
f_nac['año'] = f_nac['f_nac'].str[0:2]
f_nac['mes'] = f_nac['f_nac'].str[2:4]
f_nac['dia'] = f_nac['f_nac'].str[4:7]

for i in range(len(f_nac['año'])):
    #i = 0
    
    if pd.isna(f_nac['rfc'][i]):
        ''
    elif len(f_nac['rfc'][i]) > 13:
        f_nac.iloc[i, 2:6] = pd.NA
    else:
        ''

cumple = []

for i in range(len(f_nac['año'])):
    #i = 0
    try:
        
        año = pd.to_numeric(f_nac['año'][i])
        
        if año <= 10:
            f_nac['año'][i] = '200' + str(año)
        else:
            f_nac['año'][i] = '19' + str(año)
        
        año = f_nac['año'][i]
        mes = f_nac['mes'][i]
        dia = f_nac['dia'][i]
        
        f_cumple = dia + '/' + str(mes) + '/' + str(año)
    
        cumple.append(f_cumple)
    
    except Exception as e:
        f_cumple = ''
        cumple.append(f_cumple)
        continue
    
f_nac['nacimiento'] = cumple
f_nac['NO DE POLIZA '] = tabla['NO DE POLIZA ']

tabla = pd.merge(tabla, f_nac.iloc[:, [7,6]], on = 'NO DE POLIZA ', how = 'left')

# =============================================================================
# Generación de los grupos, sub grupos y sub sub grupos
# =============================================================================
año_c = pd.DataFrame({'rfc' : f_nac['rfc'], 
                      'año' : f_nac['año']})

grupos = pd.DataFrame({'Nombre' : tabla['ASEGURADO']})
grupos['No Emp'] = tabla['MATRICULA DE EMPLEADO']
grupos['Aseguradora'] = tabla['ASEGURADORA']
grupos['Forma P'] = tabla['FORMA DE PAGO']
grupos['Tipo'] = tabla['QUINCENA DE RECUPERACION']
grupos['delegacion'] = tabla['UBICACION']
grupos['trabajador'] = tabla['TRABAJADOR']
grupos['No de POL'] = tabla['NO DE POLIZA ']
grupos['f_na'] = imp['Fecha. Nac.']

grupos['clave_age'] = ''
grupos['grupo'] = ''
grupos['s_grupo'] = ''
grupos['s_s_grupo'] = ''

# =============================================================================
# clave_ag = []
# grupo = []
# s_grupo = []
# s_s_grupo = []
# =============================================================================

clave_qua = cve_a[cve_a['ASEGURADORAS'] == 'QUALITAS']

#aseguradora = pd.unique(grupos['Aseguradora'])
#forma = pd.unique(grupos['Forma P'])

for i in range(len(grupos['Nombre'])):
    #i = 0
    
    ne = grupos['No Emp'][i]
    aseg = grupos['Aseguradora'][i]
    form = grupos['Forma P'][i]
    tipo = grupos['Tipo'][i]
    deleg = str(grupos['delegacion'][i])
    trab = str(grupos['trabajador'][i])
    
    try: 
        f_na = pd.to_numeric(grupos['año'][i])
    
    except Exception as e:
        
        f_na = 0
    
    if deleg  == '-':
        ''
    elif len(str(grupos['delegacion'][i])) == 1:
        deleg = '0' + deleg
    else:
        ''
    valores = pd.DataFrame({'Datos' : [ne, aseg, form, tipo, deleg, trab, f_na]})
    
    if ne == 999999:
        
        grupos['No Emp'][i] = ''
        grupos['grupo'].loc[i] = 'TRADICIONAL'
        grupos['s_grupo'].loc[i] = ''
        grupos['s_s_grupo'].loc[i] = ''

        
        if aseg == 'GNP':
            grupos['clave_age'].loc[i] = cve_a.iloc[5,2]
        elif aseg == 'GS':
            grupos['clave_age'].loc[i] = cve_a.iloc[8,2]
        elif aseg == 'HDI':
            grupos['clave_age'].loc[i] = cve_a.iloc[6,2]
        elif aseg == 'AXA':
            grupos['clave_age'].loc[i] = cve_a.iloc[7,2]
        else: #PARA EL CASO DE QUALITAS
            if len(form) < 3:
                form = 'TRADICIONAL'
                grupos['clave_age'].loc[i] = clave_qua[clave_qua['CONCEPTO'] == form].iloc[0,2]
            else: 
                grupos['clave_age'].loc[i] = clave_qua[clave_qua['CONCEPTO'] == form].iloc[0,2]

    
    elif form == 'DXN' or form == 'MIXTO':
        
        grupos['grupo'].loc[i] ='IMSS'
        
        if tipo == 10 or f_na < 1979:
            grupos['s_grupo'].loc[i] = 'JUBILADO'
            
            if aseg == 'GNP':
                grupos['clave_age'].loc[i] = cve_a.iloc[5,2]
            elif aseg == 'GS':
                grupos['clave_age'].loc[i] = cve_a.iloc[8,2]
            elif aseg == 'HDI':
                grupos['clave_age'].loc[i] = cve_a.iloc[6,2]
            elif aseg == 'AXA':
                grupos['clave_age'].loc[i] = cve_a.iloc[7,2]
            else: #PARA EL CASO DE QUALITAS
                grupos['clave_age'].loc[i] = clave_qua[clave_qua['CONCEPTO'] == form].iloc[0,2]
            
            if delegaciones['ID'].isin([deleg]).any() == False or deleg == 'nan':
                grupos['s_s_grupo'].loc[i] = ''
            else:
                grupos['s_s_grupo'].loc[i] = delegaciones[delegaciones['ID'] == deleg].iloc[0,0]
            
        else: 
            
            grupos['s_grupo'].loc[i] = 'ACTIVO'
            
            if aseg == 'GNP':
                grupos['clave_age'].loc[i] = cve_a.iloc[5,2]
            elif aseg == 'GS':
                grupos['clave_age'].loc[i] = cve_a.iloc[8,2]
            elif aseg == 'HDI':
                grupos['clave_age'].loc[i] = cve_a.iloc[6,2]
            elif aseg == 'AXA':
                grupos['clave_age'].loc[i] = cve_a.iloc[7,2]
            else: #PARA EL CASO DE QUALITAS
                form = 'DXN'
                grupos['clave_age'].loc[i] = clave_qua[clave_qua['CONCEPTO'] == form].iloc[0,2]
                
            if delegaciones['ID'].isin([deleg]).any() == False or deleg == 'nan':
                grupos['s_s_grupo'].loc[i] = ''
            else:
                grupos['s_s_grupo'].loc[i] = delegaciones[delegaciones['ID'] == deleg].iloc[0,0]


# =============================================================================
# Formas de Pago
# =============================================================================

f_pago = ['Activo', 'Contado', 'Jubilado', 'Domiciliado']

grupos['Forma Pago'] = ''

for i in range(len(grupos['Nombre'])):
    #i = 0
    forma = grupos['Forma P'][i]
    
    if forma == 'DXN' or forma == 'MIXTO':
        
        if grupos['s_grupo'][i] == 'JUBILADO':
            
            grupos['Forma Pago'][i] = 'Jubilado'
            
        else:
            grupos['Forma Pago'][i] = 'Activo'
            
    elif forma == 'CONTADO':
        
        grupos['Forma Pago'][i] = 'Contado'
        
    elif forma == 'DOMICILIADO':
        
        grupos['Forma Pago'][i] = 'Domiciliado'


# =============================================================================
# Segmento para adicionar dirección
# =============================================================================

direccion = tabla_imp.iloc[:, [40, 30, 31, 32, 33, 34, 35, 36]]
direccion = direccion.rename(columns = {'Documento = No Poliza' : 'No de POL'})


grupos = pd.merge(grupos, direccion, on = 'No de POL', how = 'right')
grupos = grupos.rename(columns = {'No de POL' : 'NO DE POLIZA '})

tabla = pd.merge(tabla, grupos, on = 'NO DE POLIZA ', how = 'left')


# =============================================================================
# Derechos
# AXA 400, GNP 570, QUA 570, HDI 695, 
# =============================================================================

tabla['derechos'] = ''

for i in range(len(tabla['ASEGURADORA'])):
    #i = 0
    aseguradora = tabla['ASEGURADORA'][i]
    
    if aseguradora == 'QUALITAS':
        tabla['derechos'][i] = '570'
    elif aseguradora == 'GNP':
        tabla['derechos'][i] = '570'
    elif aseguradora == 'AXA':
        tabla['derechos'][i] = '400'
    elif aseguradora == 'HDI':
        tabla['derechos'][i] = '695'
    else:
        tabla['derechos'][i] = '950'


# =============================================================================
# Estatus
# Viegnte 0, Cancelada 1
# =============================================================================

tabla['estatus'] = ''

for i in range(len(tabla['ASEGURADORA'])):
    # i = 0
    
    estado = tabla['DOCTOS FALTANTES'][i]
    contiene_patron = pd.Series([estado]).str.contains(r'(CANCE|RECHA)', regex=True).iloc[0]
    
    if contiene_patron == True:
        tabla['estatus'][i] = '1'
    else: 
        tabla['estatus'][i] = '0'


# =============================================================================
# vendedor 
# =============================================================================

claves = pd.unique(tabla['CLAVE DE PROMOTOR'])
promotores = pd.unique(tabla['AGENTE PRODUCTOR ASOCIADO'].str.strip())


vend_certeza = pd.Series(pd.unique(tabla['AGENTE PRODUCTOR ASOCIADO'].str.contains(r'(CERTEZ|certez|Certez)')))
vend_certeza = pd.Series(pd.unique(tabla['AGENTE PRODUCTOR ASOCIADO'][tabla['AGENTE PRODUCTOR ASOCIADO'].str.contains(r'(CERTEZ|certez|Certez)', regex=True)]))
vend_certeza = pd.DataFrame({'vendedor' : vend_certeza})
clave_ag = ['1-9001', '2-9001', '3-9001', '4-9001', '5-9001', '6-9001', '0-9001', '7-9001']
vend_certeza['clave'] = clave_ag
vend_certeza = vend_certeza.iloc[:, [1,0]]

vendedor = []

for i in range(len(claves)):
    #i = 0
    vendedor.append(pd.unique(tabla[tabla['CLAVE DE PROMOTOR'] == claves[i]]['AGENTE PRODUCTOR ASOCIADO'])[0]) 

vendedores = pd.DataFrame({'clave' : claves, 'vendedor' : vendedor})
vendedores = vendedores[~vendedores['clave'].isin(['09001', '9001'])]

vendedores = pd.concat([vendedores, vend_certeza], ignore_index= True)

tabla['T_vendedores'] = ''

for i in range(len(tabla['CLAVE DE PROMOTOR'])):
    #i = 7
    clave_v = tabla['CLAVE DE PROMOTOR'][i]
    vend = tabla['AGENTE PRODUCTOR ASOCIADO'][i]
    
    if clave_v == '9001' or clave_v == '09001':
        
        vend = vend.replace('CERTEZA / ', '')
        tabla['T_vendedores'][i] = vendedores[vendedores['vendedor'].str.contains(vend[0:5])].iloc[0,1]
    
    else: 
        
        tabla['T_vendedores'][i] = vendedores[vendedores['clave'] == clave_v].iloc[0,1]
    

vendedores = vendedores.drop_duplicates(subset = 'vendedor')
vendedores = vendedores.reset_index(drop = True)
vendedores.iloc[23,0] = '9111'

vendedores = vendedores.rename(columns = {'vendedor' : 'T_vendedores'})

tabla['Correo_v'] = ''
tabla['Officina_v'] = ''
tabla['Ejec_com'] = ''
tabla['Activo_v'] = ''

for i in range(len(tabla['T_vendedores'])):
    #i = 7
    vend = tabla['T_vendedores'][i]
    
    esta = asesores['TRAMITADOR'].isin([vend]).any()
   
    if esta == True:
        
        if pd.Series([vend]).str.contains('CERTEZA')[0]:
            
            vend = vend.replace('CERTEZA / ', '')
            correo = asesores[asesores['TRAMITADOR'].str.contains(vend)].iloc[0, 3]
            oficina = asesores[asesores['TRAMITADOR'].str.contains(vend)].iloc[0, 2]
            comercial = asesores[asesores['TRAMITADOR'].str.contains(vend)].iloc[0, 4]
            
            if comercial == 'nan':
                
                comercial = 'Certeza'
                
            else:
                ''
            
            activo = asesores[asesores['TRAMITADOR'].str.contains(vend)].iloc[0, 6]
        
            if pd.Series([activo]).str.contains('baja')[0]:
                
                activo = 'Baja'
                
            else:
                
                activo = 'Activo'
        else:
            
            correo = asesores[asesores['TRAMITADOR'] == vend].iloc[0, 3]
            oficina = asesores[asesores['TRAMITADOR'] == vend].iloc[0, 2]
            comercial = asesores[asesores['TRAMITADOR'] == vend].iloc[0, 4]
            
            if comercial == 'nan':
                
                comercial = 'Certeza'
                
            else:
                ''
            
            activo = asesores[asesores['TRAMITADOR'].str.contains(vend)].iloc[0, 6]
        
            if pd.Series([activo]).str.contains('baja')[0]:
                activo = 'Baja'
            else:
                activo = 'Activo'
    else:
        
        correo = ""
        oficina = ""
        comercial = "Certeza"
        activo = "Activo"
    
    tabla['Correo_v'][i] = correo
    tabla['Officina_v'][i] = oficina
    tabla['Ejec_com'][i] = comercial
    tabla['Activo_v'][i] = activo    
    

#tabla = pd.merge(tabla, vendedores, on = 'T_vendedores', how = 'left')
    
# =============================================================================
# DETALLES DEL AUTO
# =============================================================================

autos = tabla_imp.iloc[:, [40, 83, 87, 88, 89, 90, 91, 95]]
autos = autos.rename(columns = {'Documento = No Poliza' : 'NO DE POLIZA '})

tabla = pd.merge(tabla, autos, on = 'NO DE POLIZA ', how = 'left')

# =============================================================================
# COBERTURAS
# =============================================================================

coberturas = tabla.iloc[:, [11, 13, 42]]
coberturas['COBERTURA'] = coberturas['COBERTURA'].str.replace(r'(-|c15045|)', '', regex=True)
c_p  = coberturas.drop_duplicates(subset = ['COBERTURA', 'ASEGURADORA'])

c_nue = pd.DataFrame({'COBERTURA' : pd.unique(coberturas['COBERTURA'])})
c_nue['nombre'] = ['Amplia Premium', 'Amplia', 'Integral', 'Amplia',
                   'Amplia Premium', 'Limitada', 'Responsabilidad Civil', 
                   'Amplia Integral', 'Limitada', 'Amplia Premium',
                   'Amplia Premium', 'Protegido', 'Protegido',
                   'Responsabilidad Civil', 'Amplia']

c_p = pd.merge(c_p, c_nue, on = 'COBERTURA', how = 'left')
c_p = c_p.iloc[:, [2,3]]

coberturas = pd.merge(coberturas, c_p, on = 'COBERTURA', how = 'left')
coberturas  = coberturas.drop_duplicates(subset = 'NO DE POLIZA ')
c_p  = coberturas.drop_duplicates(subset = ['COBERTURA', 'ASEGURADORA'])

coberturas = coberturas.iloc[:, [0, 3]]
coberturas = coberturas.rename(columns = {'nombre' : 'COBERTURA'})

tabla = pd.merge(tabla, coberturas, on = 'NO DE POLIZA ', how = 'left')

# =============================================================================
# EXTRAS 
# =============================================================================

extras = tabla_imp.iloc[:, [40, 106, 107, 108, 109, 110, 111]]
extras = extras.rename(columns = {'Documento = No Poliza' : 'NO DE POLIZA '})

tabla = pd.merge(tabla, extras, on = 'NO DE POLIZA ', how = 'left')


# =============================================================================
# Forma de pago tal como SICAS
# =============================================================================




# =============================================================================
# Generar la Tabla de importación
# =============================================================================

tab_names = tabla.columns


imp['Nombre'] = tabla['ASEGURADO']
imp['Entidad'] = '0'
imp['Fecha. Nac.'] = tabla['nacimiento']
imp['RFC'] = tabla['RFC']
imp['Correo1'] = tabla['Correo_v']
imp['Ejecutivo de Cuenta'] = tabla['Ejec_com']
imp['No. Empleado'] = tabla['No Emp'].astype(str)
imp['Agente'] = tabla['clave_age'].astype(str)
imp['Grupo'] = tabla['grupo']
imp['Sub Grupo'] = tabla['s_grupo']
imp['Sub Sub Grupo'] = tabla['s_s_grupo']
imp['Despacho'] = 'Certeza'
imp['Calle'] = tabla['Calle']
imp['No. Ext.'] = tabla['No. Ext.']
imp['No. Int.'] = tabla['No. Int.']
imp['CP'] = tabla['CP_y']
imp['Colonia'] = tabla['Colonia']
imp['Poblacion o Delegacion'] = tabla['Poblacion o Delegacion']
imp['Ciudad/Edo'] = tabla['Ciudad/Edo']
imp['Pais'] = 'México'
imp['Referencia1'] = tabla['NUMERO DE FOLIO'].astype(str)
imp['Tipo Documento'] = 'Poliza'
imp['Forma Pago'] = tabla['Forma Pago']
imp['Moneda'] = 'Pesos mexicanos'
imp['Documento'] = tabla['NO DE POLIZA ']
imp['T. Cambio'] = "1"
imp['Sub Ramo'] = 'Automóviles individuales' 
imp['Concepto'] = 'Automóviles individuales' 
imp['Vendedor'] = tabla['T_vendedores']
imp['Renovacion'] = '0' 
imp['Fecha Antigüedad'] = tabla['FECHA DE EMISION']
imp['Desde'] = tabla['INICIO DE VIGENCIA']
imp['Hasta'] = tabla['FIN DE VIGENCIA']
imp['Estatus'] = tabla['estatus']
imp['Prima Neta'] = pd.to_numeric(tabla['PRIMA NETA'], errors= 'coerce')
imp['Derechos'] = pd.to_numeric(tabla['derechos'], errors= 'coerce')
imp['Sub Total'] = imp['Prima Neta'] + imp['Derechos']
imp['IVA'] = imp['Sub Total'] * 0.16
imp['Prima Total'] = pd.to_numeric(tabla['PRIMA TOTAL'], errors = 'coerce')
imp['Recargos'] = (imp['Prima Total']/1.16) - (imp['Sub Total'])
imp['Sub Total'] = imp['Prima Neta'] + imp['Derechos'] + imp['Recargos']
imp['IVA'] = imp['Sub Total'] * 0.16
imp['Plan'] = tabla['COBERTURA_y']
imp['Marca'] = tabla['Marca']
imp['Transmisión'] = tabla['Transmisión ']
imp['Motor'] = tabla['NUMERO DE MOTOR'] ####
imp['Puertas'] = tabla['Puertas']
imp['Modelo'] = tabla['Modelo']
imp['Placas'] = tabla['Placas']
imp['Serie'] = tabla['Serie']
imp['Servicio'] = tabla['Servicio']
imp['Uso del Vehiculo'] = 'Particular'
imp['Ingreso Dig'] = tabla['Ingreso Dig']
imp['Ingreso Fisico'] = tabla['Ingreso Fisico']
imp['Mesa'] = tabla['Mesa']
imp['No Vale'] = tabla['No Vale']
imp['Vale recibido'] = tabla['Vale recibido']
imp['Reviso'] = tabla['Reviso']

imp = imp.astype(str)

imp.to_excel('importador 24 280624.xlsx', index = False)
