# Gestión de Recaudo y Asignación Diaria

## Objetivo
Aplicación local en Streamlit para configurar metas mensuales de recaudo, consolidar pagos de cartera castigada y vencida, identificar clientes que requieren gestión y distribuir obligaciones entre gestores según prioridad y porcentaje objetivo.

## Funcionalidades
La solución tiene tres pantallas: **Configuración mensual**, **Seguimiento de recaudo** y **Asignación diaria**. Mantiene el estado durante la sesión con `st.session_state`, permite mapear columnas cuando los nombres cambian, separa novedades de calidad sin detener todo el proceso y genera un Excel profesional descargable.

En los archivos reales suministrados se reconocen, entre otros, estos nombres: en recaudo `CED_SIN_DIG`, `OBLIGACION_ICS`, `ESTADO_OBLIG`, `SUM_of_VALOR_CONSIGNACION` y `FECHA_RECAUDO`; en la hoja `Base` de asignación `OBLIGACION`, `CEDULA SIN DIGITO`, `CLIENTE`, `SALDO_OBLIG`, `MARCA_CASTIGO` y `FILTRO VENCIDA`. La aplicación no depende de esos nombres: intenta detectarlos y permite corregir el mapeo.

## Estructura de archivos
- `app.py`: aplicación completa.
- `requirements.txt`: dependencias fijadas.
- `README.md`: instalación, uso y criterios técnicos.

## Requisitos previos
- Python 3.11 o superior.
- Acceso local a los Excel de recaudo y asignación.

## Crear entorno virtual
### Windows
```bash
python -m venv .venv
.venv\Scripts\activate
```

### macOS / Linux
```bash
python3 -m venv .venv
source .venv/bin/activate
```

## Instalar dependencias
```bash
pip install -r requirements.txt
```

## Ejecutar
```bash
streamlit run app.py
```

## Flujo de uso
### 1. Configuración mensual
Seleccione año y mes, registre la meta de cartera castigada y vencida y agregue los gestores. Cada gestor tiene prioridad, peso objetivo y estado activo/inactivo. Los pesos de gestores activos deben sumar 100%. El botón de distribución automática usa un peso decreciente según la prioridad y luego normaliza a 100%; el usuario puede modificarlo manualmente antes de guardar.

### 2. Seguimiento de recaudo
Cargue el Excel diario. La aplicación lee las hojas, propone el mapeo de obligación, cédula, tipo de cartera, recaudo y fecha de pago, y permite corregirlo. La fecha de corte se intenta obtener del nombre, incluyendo formatos como `DDMMAAAA`, `DDMESAAAA` (por ejemplo `01SEP2026`) y `AAAAMM`.

Los identificadores se convierten a texto, se eliminan espacios y terminaciones `.0`, los montos se convierten de forma segura y las fechas se filtran al mes configurado. Los duplicados exactos se excluyen antes de sumar. Se muestran metas, recaudo, brecha, cumplimiento, obligaciones/clientes con pago, días transcurridos/restantes, promedio diario y recaudo diario requerido.

### 3. Asignación diaria
Cargue el archivo mensual y seleccione la hoja; si existe `Base`, se propone como hoja inicial. Mapee obligación, cédula, cliente opcional, tipo de cartera, saldo y gestor actual opcional.

Por defecto, una regla de negocio adicional excluye **todas las obligaciones de una cédula cuando al menos una obligación de esa cédula registra pago en el mes**. Puede habilitar pagos parciales; en ese caso esas cédulas se mantienen, pero quedan después de las que no tienen pago.

Puede asignar todo, limitar por cantidad de obligaciones, por monto total o por máximo de obligaciones por gestor. El resultado muestra resumen por gestor, detalle y un botón de descarga.

## Estructura mínima esperada de los Excel
### Recaudo
Debe existir una columna para cada uno de estos conceptos:
- Número de obligación.
- Cédula / identificación.
- Tipo o estado de cartera.
- Valor del recaudo.
- Fecha del pago.

La clasificación se toma de los valores encontrados en el archivo. Valores que contengan equivalentes de `CAST`/`CASTIGADA` se clasifican como **Castigada** y valores equivalentes a `VENC`/`VENCIDA` como **Vencida**. Lo demás queda como **Sin clasificar**.

### Asignación mensual
Mínimo:
- Número de obligación.
- Cédula.
- Tipo de cartera.
- Saldo pendiente o saldo de capital.

Opcionales:
- Nombre del cliente.
- Gestor actual.

## Algoritmo de asignación
1. Se eliminan saldos no válidos, obligaciones duplicadas y registros que requieren revisión.
2. Se determina el pago principalmente por número de obligación; a partir de las obligaciones pagadas se construye el conjunto de cédulas con pago.
3. Por defecto, si una cédula tiene al menos una obligación con pago, se excluyen todas sus obligaciones.
4. Las obligaciones elegibles se ordenan primero por ausencia de pago, luego por menor cumplimiento de la meta de su tipo de cartera y finalmente por mayor saldo.
5. Se calcula el monto objetivo de cada gestor: `total a asignar × porcentaje objetivo`.
6. La primera ronda asigna los mayores saldos siguiendo el orden de prioridad de gestores.
7. Las obligaciones restantes se entregan al gestor con mayor déficit relativo: `(monto objetivo - monto actual) / monto objetivo`.
8. En empate gana el gestor con mayor prioridad.
9. Una obligación nunca se divide ni se asigna dos veces.
10. Se calcula la desviación entre monto real y objetivo. La proporcionalidad puede no ser exacta por la indivisibilidad de las obligaciones.

## Validaciones implementadas
- Metas y configuración previa.
- Existencia de gestores activos.
- Gestores y prioridades duplicadas.
- Pesos positivos y suma de 100%.
- Excel inválido o dañado.
- Mapeo de columnas indispensables.
- Obligaciones y cédulas vacías.
- Saldos y recaudos no numéricos.
- Fechas inválidas.
- Tipos de cartera no clasificables.
- Duplicados exactos de recaudo.
- Obligaciones duplicadas en asignación.
- Fecha de corte diferente al mes configurado.
- Ausencia de oportunidades elegibles.

Los errores parciales se registran como calidad de datos o exclusiones para no detener los registros válidos.

## Excel descargable
El archivo `Asignacion_Diaria_AAAAMMDD.xlsx` incluye:
- `Asignacion_Diaria`.
- `Resumen_Gestores`.
- `Resumen_Metas`.
- `Exclusiones`.
- `Calidad_Datos`.

Incluye filtros, paneles congelados, anchos ajustados, encabezados destacados y formatos monetarios/porcentuales.

## Manejo confidencial de datos
La aplicación funciona localmente y no utiliza base de datos, APIs externas ni servicios de pago. Los Excel se procesan en memoria durante la sesión de Streamlit. Se recomienda ejecutar la aplicación únicamente en equipos y redes autorizados, proteger los archivos de origen y cerrar/reiniciar la sesión al cambiar de período o usuario.

## Limitaciones conocidas
- La detección automática de columnas depende de nombres similares; por eso siempre se permite mapeo manual.
- La eliminación automática de duplicados de recaudo se limita a filas exactamente repetidas por obligación, cédula, tipo, valor y fecha. Movimientos distintos del mismo día se conservan porque podrían ser pagos legítimos.
- Si el archivo de origen trae valores acumulados y movimientos individuales mezclados sin un campo que los diferencie, la aplicación no puede inferir con certeza cuál debe prevalecer; esa situación debe revisarse en calidad de datos antes de usar el total.
- El cruce operativo usa obligación como llave principal. La cédula se usa para aplicar la regla de exclusión de todo el cliente cuando alguna de sus obligaciones registra pago; no se realiza un `merge` muchos-a-muchos.
- La asignación proporcional es aproximada porque una obligación no se divide entre gestores.
