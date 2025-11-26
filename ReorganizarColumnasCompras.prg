*!*****************************************************************
*!* ReorganizarColumnasCompras.prg
*!* 
*!* Procedimiento para reorganizar automáticamente las columnas de
*!* alícuotas en registros de COMPRAS importados desde Excel.
*!* 
*!* IMPORTANTE: Esta lógica aplica SOLO para COMPRAS (CompraVenta=01)
*!* NO modificar la funcionalidad de VENTAS (CompraVenta=02)
*!*****************************************************************

*!*****************************************************************
*!* Función: ReorganizarColumnasCompras
*!* 
*!* Descripción: 
*!*   Reorganiza las columnas de alícuotas para que los valores
*!*   significativos queden en columnas consecutivas empezando 
*!*   desde la columna 1.
*!*
*!* Parámetros (pasados por referencia):
*!*   - Alícuotas: wAli1, wAli2, wAli3
*!*   - IVA: wIVA1, wIVA2, wIVA3
*!*   - Neto: wNeto1, wNeto2, wNeto3
*!*   - Exento: wExento1, wExento2, wExento3
*!*   - NoGrava: wNoGrava1, wNoGrava2, wNoGrava3
*!*   - Monot: wMonot1, wMonot2, wMonot3
*!*   - wTot2, wTot3 (totales a recalcular)
*!*
*!* Lógica:
*!*   1. Detectar columnas con valores significativos
*!*   2. Mover columnas activas a posiciones consecutivas
*!*   3. Preservar la alícuota correspondiente
*!*   4. Recalcular totales wTot2 y wTot3
*!*****************************************************************
PROCEDURE ReorganizarColumnasCompras
    LPARAMETERS ;
        wAli1, wIVA1, wNeto1, wExento1, wNoGrava1, wMonot1, ;
        wAli2, wIVA2, wNeto2, wExento2, wNoGrava2, wMonot2, ;
        wAli3, wIVA3, wNeto3, wExento3, wNoGrava3, wMonot3, ;
        wTot2, wTot3

    *-- Variables locales para determinar columnas activas
    LOCAL lCol1Activa, lCol2Activa, lCol3Activa
    LOCAL nColumnasActivas
    
    *-- Variables temporales para reorganización
    LOCAL tmpAli1, tmpIVA1, tmpNeto1, tmpExento1, tmpNoGrava1, tmpMonot1
    LOCAL tmpAli2, tmpIVA2, tmpNeto2, tmpExento2, tmpNoGrava2, tmpMonot2
    LOCAL tmpAli3, tmpIVA3, tmpNeto3, tmpExento3, tmpNoGrava3, tmpMonot3
    
    *-- Array para almacenar columnas activas en orden
    LOCAL ARRAY aColumnasActivas[3, 6]  && 3 columnas, 6 valores cada una (Ali, IVA, Neto, Exento, NoGrava, Monot)
    LOCAL nIndiceActiva
    
    *-- Inicializar contador
    nIndiceActiva = 0
    
    *-- ===============================================================
    *-- PASO 1: Detectar columnas con valores significativos
    *-- Una columna tiene valores significativos si:
    *-- IVA > 0 OR Neto > 0 OR Exento > 0 OR NoGrava > 0 OR Monot > 0
    *-- Las retenciones NO se consideran para determinar si está "activa"
    *-- ===============================================================
    
    *-- Verificar Columna 1
    lCol1Activa = (NVL(wIVA1, 0) > 0 OR NVL(wNeto1, 0) > 0 OR ;
                   NVL(wExento1, 0) > 0 OR NVL(wNoGrava1, 0) > 0 OR ;
                   NVL(wMonot1, 0) > 0)
    
    *-- Verificar Columna 2
    lCol2Activa = (NVL(wIVA2, 0) > 0 OR NVL(wNeto2, 0) > 0 OR ;
                   NVL(wExento2, 0) > 0 OR NVL(wNoGrava2, 0) > 0 OR ;
                   NVL(wMonot2, 0) > 0)
    
    *-- Verificar Columna 3
    lCol3Activa = (NVL(wIVA3, 0) > 0 OR NVL(wNeto3, 0) > 0 OR ;
                   NVL(wExento3, 0) > 0 OR NVL(wNoGrava3, 0) > 0 OR ;
                   NVL(wMonot3, 0) > 0)
    
    *-- Contar columnas activas
    nColumnasActivas = IIF(lCol1Activa, 1, 0) + IIF(lCol2Activa, 1, 0) + IIF(lCol3Activa, 1, 0)
    
    *-- ===============================================================
    *-- PASO 2: Si todas las columnas activas ya están consecutivas
    *-- desde Col1, no hay nada que reorganizar
    *-- ===============================================================
    
    *-- Casos que NO requieren reorganización:
    *-- - Ninguna columna activa
    *-- - Solo Col1 activa
    *-- - Col1 y Col2 activas
    *-- - Col1, Col2 y Col3 activas
    
    IF nColumnasActivas = 0
        *-- No hay valores significativos, nada que hacer
        RETURN
    ENDIF
    
    IF lCol1Activa AND NOT lCol2Activa AND NOT lCol3Activa
        *-- Solo Col1 activa, ya está en posición correcta
        RETURN
    ENDIF
    
    IF lCol1Activa AND lCol2Activa AND NOT lCol3Activa
        *-- Col1 y Col2 activas, ya están consecutivas
        RETURN
    ENDIF
    
    IF lCol1Activa AND lCol2Activa AND lCol3Activa
        *-- Todas activas, mantener tal cual
        RETURN
    ENDIF
    
    *-- ===============================================================
    *-- PASO 3: Reorganizar columnas que SÍ necesitan cambio
    *-- ===============================================================
    
    *-- Guardar valores originales en variables temporales
    tmpAli1 = wAli1
    tmpIVA1 = wIVA1
    tmpNeto1 = wNeto1
    tmpExento1 = wExento1
    tmpNoGrava1 = wNoGrava1
    tmpMonot1 = wMonot1
    
    tmpAli2 = wAli2
    tmpIVA2 = wIVA2
    tmpNeto2 = wNeto2
    tmpExento2 = wExento2
    tmpNoGrava2 = wNoGrava2
    tmpMonot2 = wMonot2
    
    tmpAli3 = wAli3
    tmpIVA3 = wIVA3
    tmpNeto3 = wNeto3
    tmpExento3 = wExento3
    tmpNoGrava3 = wNoGrava3
    tmpMonot3 = wMonot3
    
    *-- Recolectar columnas activas en orden
    *-- Columna 1 si está activa
    IF lCol1Activa
        nIndiceActiva = nIndiceActiva + 1
        aColumnasActivas[nIndiceActiva, 1] = tmpAli1
        aColumnasActivas[nIndiceActiva, 2] = tmpIVA1
        aColumnasActivas[nIndiceActiva, 3] = tmpNeto1
        aColumnasActivas[nIndiceActiva, 4] = tmpExento1
        aColumnasActivas[nIndiceActiva, 5] = tmpNoGrava1
        aColumnasActivas[nIndiceActiva, 6] = tmpMonot1
    ENDIF
    
    *-- Columna 2 si está activa
    IF lCol2Activa
        nIndiceActiva = nIndiceActiva + 1
        aColumnasActivas[nIndiceActiva, 1] = tmpAli2
        aColumnasActivas[nIndiceActiva, 2] = tmpIVA2
        aColumnasActivas[nIndiceActiva, 3] = tmpNeto2
        aColumnasActivas[nIndiceActiva, 4] = tmpExento2
        aColumnasActivas[nIndiceActiva, 5] = tmpNoGrava2
        aColumnasActivas[nIndiceActiva, 6] = tmpMonot2
    ENDIF
    
    *-- Columna 3 si está activa
    IF lCol3Activa
        nIndiceActiva = nIndiceActiva + 1
        aColumnasActivas[nIndiceActiva, 1] = tmpAli3
        aColumnasActivas[nIndiceActiva, 2] = tmpIVA3
        aColumnasActivas[nIndiceActiva, 3] = tmpNeto3
        aColumnasActivas[nIndiceActiva, 4] = tmpExento3
        aColumnasActivas[nIndiceActiva, 5] = tmpNoGrava3
        aColumnasActivas[nIndiceActiva, 6] = tmpMonot3
    ENDIF
    
    *-- ===============================================================
    *-- PASO 4: Asignar columnas activas a posiciones consecutivas
    *-- ===============================================================
    
    *-- Limpiar todas las columnas primero
    wAli1 = 0
    wIVA1 = 0
    wNeto1 = 0
    wExento1 = 0
    wNoGrava1 = 0
    wMonot1 = 0
    
    wAli2 = 0
    wIVA2 = 0
    wNeto2 = 0
    wExento2 = 0
    wNoGrava2 = 0
    wMonot2 = 0
    
    wAli3 = 0
    wIVA3 = 0
    wNeto3 = 0
    wExento3 = 0
    wNoGrava3 = 0
    wMonot3 = 0
    
    *-- Asignar primera columna activa a Col1
    IF nColumnasActivas >= 1
        wAli1 = aColumnasActivas[1, 1]
        wIVA1 = aColumnasActivas[1, 2]
        wNeto1 = aColumnasActivas[1, 3]
        wExento1 = aColumnasActivas[1, 4]
        wNoGrava1 = aColumnasActivas[1, 5]
        wMonot1 = aColumnasActivas[1, 6]
    ENDIF
    
    *-- Asignar segunda columna activa a Col2
    IF nColumnasActivas >= 2
        wAli2 = aColumnasActivas[2, 1]
        wIVA2 = aColumnasActivas[2, 2]
        wNeto2 = aColumnasActivas[2, 3]
        wExento2 = aColumnasActivas[2, 4]
        wNoGrava2 = aColumnasActivas[2, 5]
        wMonot2 = aColumnasActivas[2, 6]
    ENDIF
    
    *-- Asignar tercera columna activa a Col3
    IF nColumnasActivas >= 3
        wAli3 = aColumnasActivas[3, 1]
        wIVA3 = aColumnasActivas[3, 2]
        wNeto3 = aColumnasActivas[3, 3]
        wExento3 = aColumnasActivas[3, 4]
        wNoGrava3 = aColumnasActivas[3, 5]
        wMonot3 = aColumnasActivas[3, 6]
    ENDIF
    
    *-- ===============================================================
    *-- PASO 5: Manejar caso especial con NoGrava de columnas inactivas
    *-- Si una columna NO activa tenía NoGrava > 0, sumarlo a Col1
    *-- (como en el ejemplo donde Col1 original tenía solo NoGrava)
    *-- ===============================================================
    
    *-- Si Col1 original NO estaba activa pero tenía NoGrava
    IF NOT lCol1Activa AND NVL(tmpNoGrava1, 0) > 0
        wNoGrava1 = NVL(wNoGrava1, 0) + tmpNoGrava1
    ENDIF
    
    *-- Si Col2 original NO estaba activa pero tenía NoGrava
    IF NOT lCol2Activa AND NVL(tmpNoGrava2, 0) > 0
        wNoGrava1 = NVL(wNoGrava1, 0) + tmpNoGrava2
    ENDIF
    
    *-- Si Col3 original NO estaba activa pero tenía NoGrava
    IF NOT lCol3Activa AND NVL(tmpNoGrava3, 0) > 0
        wNoGrava1 = NVL(wNoGrava1, 0) + tmpNoGrava3
    ENDIF
    
    *-- ===============================================================
    *-- PASO 6: Recalcular wTot2 y wTot3
    *-- Estos totales deben reflejar la nueva organización
    *-- ===============================================================
    
    *-- wTot2 = Total de Col2 (IVA2 + Neto2 + Exento2 + NoGrava2 + Monot2)
    wTot2 = NVL(wIVA2, 0) + NVL(wNeto2, 0) + NVL(wExento2, 0) + ;
            NVL(wNoGrava2, 0) + NVL(wMonot2, 0)
    
    *-- wTot3 = Total de Col3 (IVA3 + Neto3 + Exento3 + NoGrava3 + Monot3)
    wTot3 = NVL(wIVA3, 0) + NVL(wNeto3, 0) + NVL(wExento3, 0) + ;
            NVL(wNoGrava3, 0) + NVL(wMonot3, 0)

ENDPROC


*!*****************************************************************
*!* Función: EsColumnaActiva
*!* 
*!* Descripción: 
*!*   Determina si una columna tiene valores significativos.
*!*   Función auxiliar para usar en otras partes del código.
*!*
*!* Parámetros:
*!*   - nIVA: Valor del IVA de la columna
*!*   - nNeto: Valor del Neto de la columna
*!*   - nExento: Valor del Exento de la columna
*!*   - nNoGrava: Valor del No Gravado de la columna
*!*   - nMonot: Valor del Monotributo de la columna
*!*
*!* Retorna:
*!*   .T. si la columna tiene valores significativos
*!*   .F. si la columna está vacía
*!*****************************************************************
FUNCTION EsColumnaActiva
    LPARAMETERS nIVA, nNeto, nExento, nNoGrava, nMonot
    
    RETURN (NVL(nIVA, 0) > 0 OR NVL(nNeto, 0) > 0 OR ;
            NVL(nExento, 0) > 0 OR NVL(nNoGrava, 0) > 0 OR ;
            NVL(nMonot, 0) > 0)
ENDFUNC


*!*****************************************************************
*!* EJEMPLO DE INTEGRACIÓN EN EL CÓDIGO EXISTENTE
*!*****************************************************************
*!* 
*!* Este código debe integrarse en la sección de COMPRAS (IF CompraVenta=01)
*!* específicamente en el bucle DO WHILE ! EOF() después de leer los 
*!* valores del cursor TempoExcel y ANTES de hacer los REPLACE en 
*!* RegistrosCompras_Reg.
*!*
*!* Ejemplo de integración:
*!*
*!* ```
*!* IF CompraVenta = 01  && COMPRAS
*!*     DO WHILE ! EOF()
*!*         *-- Leer valores desde TempoExcel
*!*         *-- ... código existente para leer wAli1, wIVA1, etc ...
*!*         
*!*         *-- *** NUEVA LINEA: Reorganizar columnas ***
*!*         DO ReorganizarColumnasCompras WITH ;
*!*             wAli1, wIVA1, wNeto1, wExento1, wNoGrava1, wMonot1, ;
*!*             wAli2, wIVA2, wNeto2, wExento2, wNoGrava2, wMonot2, ;
*!*             wAli3, wIVA3, wNeto3, wExento3, wNoGrava3, wMonot3, ;
*!*             wTot2, wTot3
*!*         
*!*         *-- Continuar con REPLACE en RegistrosCompras_Reg
*!*         *-- ... código existente para REPLACE ...
*!*         
*!*         SKIP
*!*     ENDDO
*!* ENDIF
*!* ```
*!*
*!*****************************************************************


*!*****************************************************************
*!* CASOS DE PRUEBA
*!*****************************************************************
*!*
*!* Caso 1: Solo Col3 tiene valores (ejemplo del issue)
*!* Entrada:
*!*   Col1: Ali=21, IVA=0, Neto=0, Exento=0, NoGrava=25229.53, Monot=0
*!*   Col2: Ali=10.5, IVA=0, Neto=0, Exento=0, NoGrava=0, Monot=0
*!*   Col3: Ali=27, IVA=29024.14, Neto=107496.83, Exento=0, NoGrava=0, Monot=0
*!* Resultado esperado:
*!*   Col1: Ali=27, IVA=29024.14, Neto=107496.83, Exento=0, NoGrava=25229.53, Monot=0
*!*   Col2: vacía
*!*   Col3: vacía
*!*
*!* Caso 2: Col1 y Col3 tienen valores
*!* Entrada:
*!*   Col1: Ali=21, IVA=1000, Neto=5000, Exento=0, NoGrava=0, Monot=0
*!*   Col2: Ali=10.5, vacía
*!*   Col3: Ali=27, IVA=500, Neto=2000, Exento=0, NoGrava=0, Monot=0
*!* Resultado esperado:
*!*   Col1: Ali=21, IVA=1000, Neto=5000, Exento=0, NoGrava=0, Monot=0
*!*   Col2: Ali=27, IVA=500, Neto=2000, Exento=0, NoGrava=0, Monot=0
*!*   Col3: vacía
*!*
*!* Caso 3: Solo Col2 tiene valores
*!* Entrada:
*!*   Col1: Ali=21, vacía
*!*   Col2: Ali=10.5, IVA=800, Neto=3000, Exento=0, NoGrava=0, Monot=0
*!*   Col3: Ali=27, vacía
*!* Resultado esperado:
*!*   Col1: Ali=10.5, IVA=800, Neto=3000, Exento=0, NoGrava=0, Monot=0
*!*   Col2: vacía
*!*   Col3: vacía
*!*
*!* Caso 4: Todas las columnas tienen valores (no reorganizar)
*!* Entrada:
*!*   Col1: Ali=21, IVA=1000, Neto=5000, Exento=0, NoGrava=0, Monot=0
*!*   Col2: Ali=10.5, IVA=800, Neto=3000, Exento=0, NoGrava=0, Monot=0
*!*   Col3: Ali=27, IVA=500, Neto=2000, Exento=0, NoGrava=0, Monot=0
*!* Resultado esperado:
*!*   Sin cambios - todas permanecen igual
*!*
*!*****************************************************************
