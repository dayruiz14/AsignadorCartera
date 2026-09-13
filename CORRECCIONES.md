# Correcciones realizadas

## Hallazgo principal del recaudo
El nombre `FINESA_RECAUDO_01SEP2026.xlsx` corresponde a la fecha de generación o corte, pero todas las fechas de pago del archivo están entre el 1 y el 31 de agosto de 2026. Si la configuración de la app estaba en septiembre, el filtro original descartaba las 85 filas y mostraba recaudo cero.

La versión corregida detecta los meses presentes en la columna de fecha, selecciona por defecto el período encontrado cuando el configurado no existe y advierte la diferencia.

## Clasificación de cartera confirmada
- Recaudo: `ESTADO_OBLIG`, con `CAST` y `VENC`.
- Asignación, hoja `Base`: `MARCA_CASTIGO`, con `CAST` y `VENC`.
- En el archivo revisado, `FILTRO VENCIDA` está vacío.

## Totales observados en el archivo de recaudo
- 85 movimientos válidos.
- Castigada: $224.300.333,01.
- Vencida: $155.047.295,78.
- Total: $379.347.628,79.

## Asignación
La hoja `Base` contiene 2.854 obligaciones con saldo válido y un rango usado de Excel artificialmente extendido a 16.363 columnas. La versión original podía intentar leer ese rango completo, consumir demasiada memoria y fallar. Ahora la app limita la lectura a columnas con encabezados útiles y valida la configuración de gestores antes de distribuir.

En la prueba funcional con los archivos entregados se obtuvieron 2.617 obligaciones elegibles y 237 exclusiones bajo la regla predeterminada de excluir la cédula completa cuando registra al menos un pago en el mes. La cantidad puede variar si se activa la opción de admitir pagos parciales o se cambia el alcance.

## Uso recomendado
1. Configure las metas asociadas a agosto de 2026.
2. En recaudo, seleccione agosto de 2026 como período a contabilizar.
3. Mapee `ESTADO_OBLIG` como Tipo de cartera.
4. En la hoja `Base`, mapee `MARCA_CASTIGO` como Tipo de cartera y `SALDO_OBLIG` como saldo.
5. Registre gestores activos con prioridades únicas y pesos positivos que sumen 100%.
