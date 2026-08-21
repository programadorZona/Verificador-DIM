import streamlit as st
import pandas as pd
import os
import re
import tempfile
from verificacion_dim import (
    ExtractorDIANSimplificado,
    ComparadorDatos, 
    ExtractorSubpartidas,
    ValidadorDeclaracionImportacionCompleto
)
from collections import Counter, defaultdict
import io
from contextlib import redirect_stdout

# Configuración de la página
st.set_page_config(
    page_title="SmartDIM",
    page_icon="🚀",
    layout="wide"
)

# Estilos CSS sin bordes punteados
st.markdown("""
<style>
    .file-info {
        background-color: #e9ecef;
        border-radius: 5px;
        padding: 8px;
        margin: 5px 0;
        font-size: 14px;
    }
    .result-section {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Inicializar estados de sesión si no existen
def inicializar_estados():
    if 'uploader_key_counter' not in st.session_state:
        st.session_state.uploader_key_counter = 0
    if 'procesamiento_completado' not in st.session_state:
        st.session_state.procesamiento_completado = False
    if 'download_counter' not in st.session_state:
        st.session_state.download_counter = 0
    # Estados para los datos
    if 'comparacion_data' not in st.session_state:
        st.session_state.comparacion_data = None
    if 'anexos_data' not in st.session_state:
        st.session_state.anexos_data = None
    if 'reporte_comparacion' not in st.session_state:
        st.session_state.reporte_comparacion = None
    if 'reporte_anexos' not in st.session_state:
        st.session_state.reporte_anexos = None
    if 'datos_dian' not in st.session_state:
        st.session_state.datos_dian = None
    if 'datos_subpartidas' not in st.session_state:
        st.session_state.datos_subpartidas = None
    # Nuevos estados para los resúmenes
    if 'datos_proveedor' not in st.session_state:
        st.session_state.datos_proveedor = None
    if 'resumen_codigos' not in st.session_state:
        st.session_state.resumen_codigos = None
    if 'estadisticas_validacion' not in st.session_state:
        st.session_state.estadisticas_validacion = None
    if 'validacion_integridad' not in st.session_state:
        st.session_state.validacion_integridad = None

# =============================================================================
# NUEVAS FUNCIONES PARA MOSTRAR RESULTADOS EN EL FORMATO ESPECÍFICO
# =============================================================================

def mostrar_resultados_validacion_formateados(datos_proveedor, resumen_codigos, estadisticas_validacion, validacion_integridad):
    """Muestra los resultados de validación en el formato específico solicitado"""
    
    # Información del Proveedor
    st.markdown("### 👤 Información del Proveedor")
    nit = datos_proveedor.get('nit', 'No disponible')
    nombre = datos_proveedor.get('nombre', 'No disponible')
    st.markdown(f"**📇 NIT:** {nit}  \n**🏢 Nombre:** {nombre}")
    
    # Resumen por código
    st.markdown("### 🗒️ Resumen por código:")
    if resumen_codigos:
        for codigo, info in resumen_codigos.items():
            cantidad = info.get('cantidad', 0)
            nombre_doc = info.get('nombre', 'DOCUMENTO')
            st.markdown(f"• **Código {codigo}:** {cantidad} - {nombre_doc}")

    
    # Validación de Integridad (si hay problemas críticos)
    tiene_problemas_criticos = False
    if validacion_integridad:
        st.markdown("### 🔍 VALIDACIÓN DE INTEGRIDAD:")
        
        if 'levantes_duplicados' in validacion_integridad:
            info = validacion_integridad['levantes_duplicados']
            st.markdown(f"❌ {info['cantidad']} Levantes duplicados: {info['numero']}")
            tiene_problemas_criticos = True
        
        if 'desbalance' in validacion_integridad:
            info = validacion_integridad['desbalance']
            st.markdown(f"❌ Desbalance: {info['di']} DI vs {info['levantes']} Levantes")
            tiene_problemas_criticos = True
    
    # Análisis de Integridad
    st.markdown("### 🔍 Análisis de Integridad")

    total_di_anexos = estadisticas_validacion.get('total_di', 0)
    total_di_procesadas = estadisticas_validacion.get('total_di_dian', 0)
    di_faltantes = total_di_anexos - total_di_procesadas
    
    st.markdown(
        f"""
        **📄 DI en Anexos:** {total_di_anexos}  
        **⚠️ Faltantes:** {di_faltantes}  
        **✅ DI Procesadas:** {total_di_procesadas} de {total_di_anexos} totales
        """
    )
    
    
    # Estado de la Validación
    st.markdown("### 📋 Estado de la Validación")
    
    if tiene_problemas_criticos:
        if 'desbalance' in validacion_integridad:
            info = validacion_integridad['desbalance']
            st.markdown(f"❌ Desbalance detectado: {info['di']} DI vs {info['levantes']} Levantes")
    else:
        # Calcular balance DI vs Levantes
        di_count = resumen_codigos.get('9', {}).get('cantidad', 0) if resumen_codigos else 0
        levantes_count = resumen_codigos.get('47', {}).get('cantidad', 0) if resumen_codigos else 0
        
        if di_count == levantes_count:
            st.markdown(f"✅ Balance correcto en anexos: {di_count} DI = {levantes_count} Levantes")
        else:
            st.markdown(f"❌ Desbalance detectado: {di_count} DI vs {levantes_count} Levantes")
            tiene_problemas_criticos = True
    
    if di_faltantes > 0:
        st.markdown(f"⚠️ Diferencia encontrada: {total_di_procesadas} DI procesadas vs {total_di_anexos} DI en anexos")
        st.markdown(f"   📝 Faltan por procesar: {di_faltantes} declaraciones de DI")
    
    # RESUMEN EJECUTIVO
    st.markdown("### 🗒️ RESUMEN EJECUTIVO")
    
    declaraciones_correctas = estadisticas_validacion.get('declaraciones_correctas', 0)
    # Corrección para evitar división por cero
    denom = total_di_anexos if total_di_anexos > 0 else (total_di_procesadas if total_di_procesadas > 0 else 1)
    eficiencia = (total_di_procesadas / denom * 100)
    
    st.markdown(f"[**DI Procesadas:** {total_di_procesadas}/{total_di_anexos} -{di_faltantes}] [**Validación:** {declaraciones_correctas}✅ Perfecto] [**Eficiencia:** {eficiencia:.1f}%]")
    
    # Estado Final del Proceso
    st.markdown("### 📋 Estado Final del Proceso")
    
    if tiene_problemas_criticos:
        st.markdown("🚨 **PROCESO COMPLETADO CON DIFERENCIAS**")
        st.markdown("Se detectaron inconsistencias en la validación de integridad")
    else:
        if di_faltantes > 0:
            st.markdown("⚠️ **PROCESO COMPLETADO CON INCOMPLETITUD**")
        else:
            st.markdown("✅ **PROCESO COMPLETADO EXITOSAMENTE**")
    
    st.markdown(f"📈 {total_di_procesadas} de {total_di_anexos} DI procesadas | ✅ {declaraciones_correctas} correctas | ❌ {estadisticas_validacion.get('declaraciones_con_errores', 0)} con diferencias")

def extraer_datos_de_consola_mejorado(consola_output):
    """Extrae datos del proveedor de la salida de consola - MEJORADO"""
    datos = {'nit': 'No disponible', 'nombre': 'No disponible'}
    
    lineas = consola_output.split('\n')
    for i, linea in enumerate(lineas):
        if 'NIT:' in linea:
            # Buscar NIT en la línea actual o siguiente
            nit_match = re.search(r'NIT:\s*([0-9]+)', linea)
            if nit_match:
                datos['nit'] = nit_match.group(1)
            elif i + 1 < len(lineas):
                nit_match = re.search(r'([0-9]{6,12})', lineas[i + 1])
                if nit_match:
                    datos['nit'] = nit_match.group(1)
        
        if 'Nombre:' in linea or 'Razón Social:' in linea:
            # Buscar nombre en la línea actual o siguiente
            nombre_match = re.search(r'(?:Nombre|Razón Social):\s*(.+)', linea)
            if nombre_match:
                datos['nombre'] = nombre_match.group(1).strip()
            elif i + 1 < len(lineas):
                nombre_texto = lineas[i + 1].strip()
                if nombre_texto and not nombre_texto.isdigit():
                    datos['nombre'] = nombre_texto
    
    return datos

def extraer_resumen_de_consola_mejorado(consola_output):
    """Extrae resumen de códigos Y validación de integridad de la salida de consola - MEJORADO"""
    resumen = {}
    validacion_integridad = {}
    
    lineas = consola_output.split('\n')
    en_resumen = False
    en_validacion = False
    
    for linea in lineas:
        # Capturar VALIDACIÓN DE INTEGRIDAD
        if 'VALIDACIÓN DE INTEGRIDAD:' in linea or 'VALIDACION DE INTEGRIDAD:' in linea:
            en_validacion = True
            continue
        
        if en_validacion:
            if '❌' in linea:
                if 'Levantes duplicados:' in linea or 'DI duplicadas:' in linea:
                    match = re.search(r'❌\s*(\d+)\s*(?:Levantes|DI)\s*duplicados?:\s*([0-9, ]+)', linea)
                    if match:
                        validacion_integridad['levantes_duplicados'] = {
                            'cantidad': match.group(1),
                            'numero': match.group(2).strip()
                        }
                elif 'Desbalance:' in linea:
                    match = re.search(r'❌\s*Desbalance:\s*(\d+)\s*DI\s*vs\s*(\d+)\s*Levantes', linea)
                    if match:
                        validacion_integridad['desbalance'] = {
                            'di': match.group(1),
                            'levantes': match.group(2)
                        }
            elif not linea.strip() or '📋 Declaraciones encontradas:' in linea:
                en_validacion = False
        
        # Capturar RESUMEN POR CÓDIGO
        if 'Resumen por código:' in linea or '🗒️ Resumen por código:' in linea:
            en_resumen = True
            continue
        
        if en_resumen and linea.strip().startswith('•'):
            match = re.search(r'•\s*Código\s+(\d+):\s*(\d+)\s*-\s*(.+)', linea)
            if match:
                codigo = match.group(1)
                cantidad = int(match.group(2))
                nombre = match.group(3).strip()
                resumen[codigo] = {'cantidad': cantidad, 'nombre': nombre}
        
        if en_resumen and not linea.strip().startswith('•') and linea.strip():
            en_resumen = False
    
    return resumen, validacion_integridad

def extraer_estadisticas_de_consola_mejorado(consola_output, datos_dian):
    """Extrae estadísticas REALES de la salida de consola - MEJORADO"""
    estadisticas = {
        'total_anexos': 0,
        'total_di': 0,
        'total_di_dian': len(datos_dian) if datos_dian is not None else 0,
        'declaraciones_con_errores': 0,
        'declaraciones_correctas': 0,
        'datos_dian': datos_dian
    }
    
    lineas = consola_output.split('\n')
    
    for linea in lineas:
        # Buscar total de anexos
        if 'anexos encontrados' in linea:
            match = re.search(r'✅\s*(\d+)\s*anexos', linea)
            if match:
                estadisticas['total_anexos'] = int(match.group(1))
        
        # Buscar total DI en anexos
        if 'Código 9:' in linea:
            match = re.search(r'Código\s*9:\s*(\d+)', linea)
            if match:
                estadisticas['total_di'] = int(match.group(1))
        
        # Buscar declaraciones con errores
        if 'Declaraciones con errores:' in linea:
            match = re.search(r'Declaraciones con errores:\s*(\d+)', linea)
            if match:
                estadisticas['declaraciones_con_errores'] = int(match.group(1))
        elif 'declaraciones con errores' in linea.lower():
            match = re.search(r'(\d+)\s*declaraciones con errores', linea.lower())
            if match:
                estadisticas['declaraciones_con_errores'] = int(match.group(1))
    
    # Calcular declaraciones correctas
    estadisticas['declaraciones_correctas'] = estadisticas['total_di_dian'] - estadisticas['declaraciones_con_errores']
    
    return estadisticas

# =============================================================================
# FUNCIONES AUXILIARES EXISTENTES (MODIFICADA SOLO LA LÓGICA DE CONTEO)
# =============================================================================

def mostrar_resumen_comparacion_simplificado(reporte_comparacion, reporte_anexos):
    """Muestra solo el resumen esencial de la comparación DIM vs Subpartidas - CORREGIDO"""
    
    if reporte_comparacion is None or reporte_comparacion.empty:
        return
    
    # 1. Obtener conjunto de DI con errores en Anexos
    dis_con_error_anexos = set()
    if reporte_anexos is not None and not reporte_anexos.empty:
        errores = reporte_anexos[reporte_anexos['Coincidencias'].str.contains('NO COINCIDE', na=False)]
        # Normalizar a string y quitar espacios para comparar
        dis_con_error_anexos = set(errores['Numero DI'].astype(str).str.strip())

    # 2. Filtrar solo filas individuales del reporte numérico
    di_individuales = reporte_comparacion[
        (reporte_comparacion['4. Número DI'] != 'VALORES ACUMULADOS') & 
        (reporte_comparacion['4. Número DI'] != 'VALORES ACUMULADOS (MÚLTIPLES SUBPARTIDAS)')
    ]
    
    total_di = len(di_individuales)
    con_diferencias = 0
    
    # 3. Contar diferencias cruzadas (Numérica OR Anexos)
    for _, row in di_individuales.iterrows():
        numero_di = str(row['4. Número DI']).strip()
        resultado_numerico = str(row['Resultado verificación'])
        
        # Es error si el Excel de comparación dice error, O si está en la lista de errores de anexos
        if '❌' in resultado_numerico or numero_di in dis_con_error_anexos:
            con_diferencias += 1
            
    conformes = total_di - con_diferencias
    
    st.markdown("### 🗒️ Resumen Comparación DIM vs Subpartidas")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total DI procesadas", total_di)
    with col2:
        st.metric("DI conformes", conformes)
    with col3:
        st.metric("DI con diferencias", con_diferencias)

# =============================================================================
# FUNCIONES PRINCIPALES ACTUALIZADAS
# =============================================================================

def procesar_conciliacion(dian_pdfs, excel_subpartidas, excel_anexos):
    """Procesa la conciliación con los archivos cargados - ACTUALIZADA"""
    
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Guardar archivos en temporal
            for pdf in dian_pdfs:
                with open(os.path.join(temp_dir, pdf.name), "wb") as f:
                    f.write(pdf.getbuffer())
            
            excel_sub_path = os.path.join(temp_dir, excel_subpartidas.name)
            with open(excel_sub_path, "wb") as f:
                f.write(excel_subpartidas.getbuffer())
            
            excel_anexos_path = os.path.join(temp_dir, excel_anexos.name)  
            with open(excel_anexos_path, "wb") as f:
                f.write(excel_anexos.getbuffer())

            # Procesar comparación DIM vs Subpartidas
            st.info("🔍 Comparando DIM vs Subpartidas...")
            
            extractor_dian = ExtractorDIANSimplificado()
            datos_dian = extractor_dian.procesar_multiples_dis(temp_dir)
            errores_extraccion = getattr(extractor_dian, 'errores_extraccion', [])

            if datos_dian is None or datos_dian.empty:
                st.error("❌ No se pudieron extraer datos de las DIM")
                if errores_extraccion:
                    st.markdown("#### Detalle de errores por archivo")
                    st.dataframe(
                        pd.DataFrame(errores_extraccion),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.warning("No se encontró ningún archivo PDF para procesar.")
                return None

            if errores_extraccion:
                st.warning("Algunos PDF no pudieron procesarse:")
                st.dataframe(
                    pd.DataFrame(errores_extraccion),
                    use_container_width=True,
                    hide_index=True
                )

            st.success(f"✅ {len(datos_dian)} declaraciones DIAN extraídas")
            
            extractor_subpartidas = ExtractorSubpartidas()
            datos_subpartidas = extractor_subpartidas.extraer_y_estandarizar(temp_dir)
            
            if datos_subpartidas.empty:
                st.error("❌ No se pudieron extraer datos del archivo de subpartidas")
                return None
            
            st.success(f"✅ Datos de subpartidas extraídos: {len(datos_subpartidas)} registros")
            
            # Detectar si hay múltiples subpartidas para mostrar info
            comparador = ComparadorDatos()
            multiples_subpartidas = comparador.detectar_multiples_subpartidas(datos_subpartidas)
            
            if multiples_subpartidas:
                st.info(f"🔍 Detectadas {len(datos_subpartidas)} subpartidas - Aplicando lógica de suma y comparación")
            else:
                st.info("🔍 Subpartida única detectada - Aplicando lógica estándar")
            
            output_comparacion = os.path.join(temp_dir, "comparacion_dim_subpartidas.xlsx")
            reporte_comparacion = comparador.generar_reporte_comparacion(
                datos_dian, datos_subpartidas, output_comparacion
            )

            # Procesar validación de anexos
            st.info("🔄 Validando anexos FMM...")
            
            validador = ValidadorDeclaracionImportacionCompleto()
            output_anexos = os.path.join(temp_dir, "validacion_anexos.xlsx")
            
            # CAPTURAR LA SALIDA DE CONSOLA DEL VALIDADOR
            output_buffer = io.StringIO()
            
            with redirect_stdout(output_buffer):
                resultado_validacion = validador.procesar_validacion_completa(temp_dir, output_anexos)
            
            # Obtener la salida de consola
            consola_output = output_buffer.getvalue()
            
            # EXTRAER DATOS REALES DEL PROCESAMIENTO CON LAS NUEVAS FUNCIONES
            datos_proveedor = extraer_datos_de_consola_mejorado(consola_output)
            resumen_codigos, validacion_integridad = extraer_resumen_de_consola_mejorado(consola_output)
            estadisticas_validacion = extraer_estadisticas_de_consola_mejorado(consola_output, datos_dian)
            
            # Si el validador retorna un diccionario, usarlo, sino usar los datos extraídos
            if isinstance(resultado_validacion, dict):
                reporte_anexos = resultado_validacion.get('reporte_anexos')
                # Combinar con datos extraídos de consola
                datos_proveedor = resultado_validacion.get('datos_proveedor', datos_proveedor)
                resumen_codigos = resultado_validacion.get('resumen_codigos', resumen_codigos)
                estadisticas_validacion = resultado_validacion.get('estadisticas_validacion', estadisticas_validacion)
            else:
                reporte_anexos = resultado_validacion

            # GUARDAR RESULTADOS EN SESSION_STATE - CLAVE PARA PERSISTENCIA
            with open(output_comparacion, "rb") as f:
                st.session_state.comparacion_data = f.read()
            if os.path.exists(output_anexos):
                with open(output_anexos, "rb") as f:
                    st.session_state.anexos_data = f.read()
            else:
                st.warning("⚠️ No se generó el reporte de anexos porque no se encontraron coincidencias entre el PDF y el Excel.")
                st.session_state.anexos_data = None
    
            # Guardar también los DataFrames completos para mostrar resultados
            st.session_state.reporte_comparacion = reporte_comparacion
            st.session_state.reporte_anexos = reporte_anexos
            st.session_state.datos_dian = datos_dian
            st.session_state.datos_subpartidas = datos_subpartidas
            # Guardar las variables de resumen
            st.session_state.datos_proveedor = datos_proveedor
            st.session_state.resumen_codigos = resumen_codigos
            st.session_state.estadisticas_validacion = estadisticas_validacion
            st.session_state.validacion_integridad = validacion_integridad

            return {
                'comparacion': reporte_comparacion is not None,
                'anexos': reporte_anexos is not None,
                'datos_dian': datos_dian,
                'datos_subpartidas': datos_subpartidas,
                'reporte_comparacion': reporte_comparacion,
                'reporte_anexos': reporte_anexos,
                'datos_proveedor': datos_proveedor,
                'resumen_codigos': resumen_codigos,
                'estadisticas_validacion': estadisticas_validacion,
                'validacion_integridad': validacion_integridad,
                'multiples_subpartidas': multiples_subpartidas  # Nuevo campo para tracking
            }

        except Exception as e:
            st.error(f"❌ Error en el procesamiento: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
            return None

def mostrar_resultados_en_pantalla():
    """Muestra los resultados detallados en pantalla usando session_state - ACTUALIZADA Y SIMPLIFICADA"""
    
    st.markdown("---")
    st.header("📋 Resultados de la Verificación")
    
    # MOSTRAR RESULTADOS DE VALIDACIÓN EN EL NUEVO FORMATO (PRIMERO)
    if (st.session_state.datos_proveedor is not None and 
        st.session_state.resumen_codigos is not None and 
        st.session_state.estadisticas_validacion is not None):
        
        mostrar_resultados_validacion_formateados(
            st.session_state.datos_proveedor,
            st.session_state.resumen_codigos,
            st.session_state.estadisticas_validacion,
            st.session_state.validacion_integridad
        )
    else:
        st.error("No se pudieron cargar los datos de validación")

    # RESULTADO COMPARACIÓN DIM vs SUBPARTIDAS - ACTUALIZADO CON LÓGICA DE CRUCE
    st.markdown("---")
    if st.session_state.reporte_comparacion is not None:
        mostrar_resumen_comparacion_simplificado(
            st.session_state.reporte_comparacion, 
            st.session_state.reporte_anexos # <--- SE AGREGA EL REPORTE DE ANEXOS PARA EL CRUCE
        )

    # Resultados de Validación de Anexos - TABLA DETALLADA (OPCIONAL)
    with st.expander("🔍 Ver Detalle de Validación de Anexos"):
        if st.session_state.reporte_anexos is not None:
            reporte_anexos = st.session_state.reporte_anexos
            
            if reporte_anexos is not None and not reporte_anexos.empty:
                # Mostrar tabla de validación con resaltado SOLO para diferencias
                st.markdown("**Detalle de Validación:**")
                
                def resaltar_solo_validacion_anexos(row):
                    """Resalta SOLO filas que no coinciden en la validación de anexos"""
                    if row['Coincidencias'] == '❌ NO COINCIDE':
                        return ['background-color: #ffcccc'] * len(row)  # Rojo claro solo para diferencias
                    else:
                        return [''] * len(row)  # Sin resaltado para coincidencias
                
                # Aplicar el estilo
                styled_anexos = reporte_anexos.style.apply(resaltar_solo_validacion_anexos, axis=1)
                
                st.dataframe(styled_anexos, use_container_width=True)

    # Resultados de Comparación DIM vs Subpartidas - TABLA DETALLADA (OPCIONAL)
    with st.expander("🔍 Ver Detalle de Comparación DIM vs Subpartidas"):
        if st.session_state.reporte_comparacion is not None:
            reporte = st.session_state.reporte_comparacion
            
            # Mostrar tabla de resultados con resaltado SOLO para diferencias
            st.markdown("**Detalle por Declaración:**")
            
            # Filtrar solo filas individuales (excluyendo totales acumulados)
            di_individuales = reporte[
                (reporte['4. Número DI'] != 'VALORES ACUMULADOS') & 
                (reporte['4. Número DI'] != 'VALORES ACUMULADOS (MÚLTIPLES SUBPARTIDAS)')
            ]
            
            def resaltar_solo_diferencias(row):
                """Resalta SOLO filas que tienen diferencias (❌)"""
                if '❌' in str(row['Resultado verificación']):
                    return ['background-color: #ffcccc'] * len(row)  # Rojo claro solo para diferencias
                else:
                    return [''] * len(row)  # Sin resaltado para conformes
            
            # Aplicar el estilo
            styled_reporte = di_individuales.style.apply(resaltar_solo_diferencias, axis=1)
            
            # Mostrar la tabla con estilos
            st.dataframe(styled_reporte, use_container_width=True)
            
            # MOSTRAR TOTALES ACUMULADOS (como estaba antes)
            fila_totales = reporte[reporte['4. Número DI'] == 'VALORES ACUMULADOS']
            if not fila_totales.empty:
                st.markdown("**Totales Acumulados:**")
                st.dataframe(fila_totales, use_container_width=True)
                
                # Resaltar también los totales si hay diferencias
                if '❌' in str(fila_totales.iloc[0]['Resultado verificación']):
                    st.warning("⚠️ Se detectaron diferencias en los totales acumulados")
            
            # MOSTRAR TOTALES MÚLTIPLES SUBPARTIDAS (si existe)
            fila_totales_multiples = reporte[reporte['4. Número DI'] == 'VALORES ACUMULADOS (MÚLTIPLES SUBPARTIDAS)']
            if not fila_totales_multiples.empty:
                st.markdown("**Totales Múltiples Subpartidas:**")
                st.dataframe(fila_totales_multiples, use_container_width=True)
                
                # Resaltar también los totales si hay diferencias
                if '❌' in str(fila_totales_multiples.iloc[0]['Resultado verificación']):
                    st.warning("⚠️ Se detectaron diferencias en los totales de múltiples subpartidas")

def mostrar_botones_descarga():
    """Muestra los botones para descargar los Excel"""
    
    st.markdown("---")
    st.markdown(
    "<h2 style='text-align: center;'>📥 Descargar Resultados Completos</h2>",
    unsafe_allow_html=True
)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.session_state.comparacion_data is not None:
            # Usar una key única dinámica basada en el contador
            download_key_comp = f"download_comparacion_{st.session_state.download_counter}"
            st.download_button(
                label="📥 Descargar Validación DIM vs Subpartidas (Excel)",
                data=st.session_state.comparacion_data,
                file_name="Comparacion_DIM_Subpartidas.xlsx",
                mime="application/vnd.ms-excel",
                use_container_width=True,
                key=download_key_comp
            )
        else:
            st.button(
                "📊 Comparación No Disponible",
                disabled=True,
                use_container_width=True
            )
    
    with col2:
        if st.session_state.anexos_data is not None:
            # Usar una key única dinámica basada en el contador
            download_key_anex = f"download_anexos_{st.session_state.download_counter}"
            st.download_button(
                label="📥 Descargar Comparación Anexos FMM (Excel)",
                data=st.session_state.anexos_data,
                file_name="Validacion_Anexos_FMM.xlsx", 
                mime="application/vnd.ms-excel",
                use_container_width=True,
                key=download_key_anex
            )
        else:
            st.button(
                "📋 Validación No Disponible",
                disabled=True,
                use_container_width=True
            )

def main():
    inicializar_estados()
    
    # Header principal
    st.title("Aplicación de Verificación DIM vs FMM - SmartDIM 🚀 ")

    # Estilos para el sidebar: fuente y tamaño
    st.markdown("""
    <style>
        /* Cambiar fuente y tamaño SOLO en el sidebar */
        [data-testid="stSidebar"] {
            font-family: 'Roboto', sans-serif;
            font-size: 14px;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Instrucciones en sidebar
    with st.sidebar:
        st.header("🧭 Instrucciones de uso")
        st.markdown("""
        1. Cargar Declaraciones PDF
        2. Cargar Excel Subpartidas
        3. Cargar Excel Anexos
        4. Ejecutar Verificación
        5. Ver resultados en pantalla
        6. Descargar Resultados
        """)
        
        # Botón de limpieza
        if st.button("🧹 Limpiar y Reiniciar", type="secondary", use_container_width=True):
            # Limpiar todo el estado específico
            st.session_state.comparacion_data = None
            st.session_state.anexos_data = None
            st.session_state.reporte_comparacion = None
            st.session_state.reporte_anexos = None
            st.session_state.datos_dian = None
            st.session_state.datos_subpartidas = None
            st.session_state.datos_proveedor = None
            st.session_state.resumen_codigos = None
            st.session_state.estadisticas_validacion = None
            st.session_state.validacion_integridad = None
            st.session_state.procesamiento_completado = False
            
            # Incrementar el contador para forzar nuevos file uploaders
            st.session_state.uploader_key_counter += 1
            st.session_state.download_counter += 1
            
            # Mensaje de confirmación
            st.sidebar.success("✅ Todo ha sido limpiado. Puedes cargar nuevos archivos.")
            
            # Forzar actualización
            st.rerun()

    # Sección de carga de archivos
    st.header("Cargar Archivos")

    # Usar el contador como parte de la key para forcear reset
    current_key = st.session_state.uploader_key_counter

    # Declaraciones PDF (DIAN)
    st.subheader("Declaraciones PDF (DIAN)")
    dian_pdfs = st.file_uploader(
        "Arrastre y suelte archivos PDF de DIAN aquí",
        type=['pdf'],
        accept_multiple_files=True,
        key=f"dian_pdfs_{current_key}"
    )
    st.caption("Límite: 200 MB por archivo • PDF")

    if dian_pdfs:
        st.markdown("**Archivos cargados:**")
        for pdf in dian_pdfs:
            st.markdown(f'<div class="file-info">📄 {pdf.name} ({pdf.size / 1024 / 1024:.1f} MB)</div>', 
                       unsafe_allow_html=True)

    # Excel de Subpartidas
    st.subheader("Archivo Excel (Subpartidas)")
    excel_subpartidas = st.file_uploader(
        "Arrastre y suelte Excel de subpartidas aquí",
        type=['xlsx'],
        key=f"excel_subpartidas_{current_key}"
    )
    st.caption("Formatos soportados: XLSX")

    if excel_subpartidas:
        st.markdown(f'<div class="file-info">📋 {excel_subpartidas.name} ({excel_subpartidas.size / 1024:.1f} KB)</div>', 
                   unsafe_allow_html=True)

    # Excel de Anexos/Proveedores
    st.subheader("Archivo Excel (Anexos FMM)")
    excel_anexos = st.file_uploader(
        "Arrastre y suelte Excel de anexos FMM aquí",
        type=['xlsx'],
        key=f"excel_anexos_{current_key}"
    )
    st.caption("Formatos soportados: XLSX")

    if excel_anexos:
        st.markdown(f'<div class="file-info">📋 {excel_anexos.name} ({excel_anexos.size / 1024:.1f} KB)</div>', 
                   unsafe_allow_html=True)

    st.markdown("---")

    # Proceso de conciliación
    st.header("Proceso: Verificación")

    # Verificar archivos mínimos para nuevo procesamiento
    archivos_cargados = (dian_pdfs and excel_subpartidas and excel_anexos)

    # Mostrar resultados existentes si los hay
    if st.session_state.procesamiento_completado and st.session_state.reporte_comparacion is not None:
        st.info("📈 Mostrando resultados de conciliación previa. Puedes descargar los archivos o cargar nuevos para reprocesar.")
        mostrar_resultados_en_pantalla()
        mostrar_botones_descarga()
        
        # Mostrar botón para nuevo procesamiento si hay archivos cargados
        if archivos_cargados:
            st.markdown("---")
            st.subheader("Reprocesar con nuevos archivos")
            if st.button("🔄 Ejecutar Nueva Verificación", type="primary", use_container_width=True):
                with st.spinner("Procesando nueva verificación..."):
                    resultados = procesar_conciliacion(dian_pdfs, excel_subpartidas, excel_anexos)
                    if resultados:
                        st.success("✅ Nueva verificación completada exitosamente")
                        st.rerun()
        return

    # Si no hay resultados previos, procesar normalmente
    if not archivos_cargados:
        st.warning("⚠️ Cargue todos los archivos requeridos para continuar")
        return

    # Mostrar resumen
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("PDFs DIAN", len(dian_pdfs))
    with col2:
        st.metric("Excel Subpartidas", "✓" if excel_subpartidas else "✗")
    with col3:
        st.metric("Excel Anexos", "✓" if excel_anexos else "✗")

    # Botón de procesamiento
    if st.button("🔄 Ejecutar Verificación", type="primary", use_container_width=True):
        with st.spinner("Procesando verificación..."):
            resultados = procesar_conciliacion(dian_pdfs, excel_subpartidas, excel_anexos)
            
            if resultados:
                st.session_state.procesamiento_completado = True
                st.success("✅ Verificación completada exitosamente")
                st.rerun()

if __name__ == "__main__":
    main()







