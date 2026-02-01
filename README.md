# Generador de Reportes CSC - Renta Comercial y Fiscal

## Descripción General

Este archivo de Excel habilitado para macros (`CSC_reportes v3.xlsm`) automatiza la generación de reportes de **Renta Comercial y Fiscal (RCF)** a partir de balances de comprobación contables. La macro principal `Generar_estado_situacion_financiera()` procesa datos contables y genera automáticamente:

- Estado de Situación Financiera
- Anexo N°13 (Conciliación Fiscal)
- Base de Impuesto Diferido
- Cálculos de tributos (LOCTI, ISLR, Ley del Deporte)

## Propósito

Facilitar la preparación de reportes fiscales mediante la automatización de:
- Búsqueda y consolidación de cuentas contables específicas
- Cálculos de conciliación entre base contable y fiscal
- Generación de anexos fiscales requeridos
- Cálculo de provisiones tributarias

## Hojas de Trabajo Requeridas

El archivo debe contener las siguientes hojas:

| Hoja | Descripción |
|------|-------------|
| **BC** | Balance de Comprobación del período actual |
| **BC inicial** | Balance de Comprobación del cierre del ejercicio anterior |
| **RCF** | Reporte de Renta Comercial y Fiscal (salida) |
| **Anexo N13** | Conciliación Fiscal (salida) |
| **Base impuesto Diferido** | Cálculos de impuesto diferido (salida) |
| **Menu** | Panel de control con indicadores fiscales (salida) |

## Instrucciones de Uso

### Paso 1: Preparar los Datos de Entrada

1. **Hoja "BC inicial"**: Importar el Balance de Comprobación del **cierre del ejercicio fiscal anterior**
   - Columna A: Código de cuenta
   - Columna C: Saldo de la cuenta

2. **Hoja "BC"**: Importar el Balance de Comprobación del **mes actual** que deseas reportar
   - Columna A: Código de cuenta (para algunas búsquedas)
   - Columna B: Código de cuenta auxiliar (para búsquedas detalladas)
   - Columna D: Saldo de la cuenta

### Paso 2: Ejecutar la Macro

1. Abrir el archivo `CSC_reportes v3.xlsm`
2. Presionar `Alt + F8` para abrir el cuadro de diálogo de macros
3. Seleccionar `Generar_estado_situacion_financiera`
4. Hacer clic en **Ejecutar**
5. Esperar el mensaje de confirmación: *"Los montos han sido transferidos correctamente."*

### Paso 3: Revisar los Resultados

Los datos se generarán automáticamente en las hojas de salida:
- **RCF**: Verificar los montos en las celdas calculadas
- **Anexo N13**: Revisar las conciliaciones fiscales
- **Base impuesto Diferido**: Validar cálculos de impuesto diferido
- **Menu**: Consultar indicadores y provisiones tributarias

## Funcionalidades Principales

### 1. Procesamiento de Cuentas Contables

La macro busca y consolida grupos específicos de cuentas contables:

#### Grupo 1: Otros Ingresos Mercantiles (611xxx)
- **Cuentas**: 611011, 611012, 611021, 611022, 611031, etc.
- **Salida**: Celda `RCF!D12`
- **Propósito**: Consolidar ingresos mercantiles

#### Grupo 2-3: Cuentas Específicas de Ingresos
- **Salida**: Celdas `RCF!D14` y `RCF!D17`

#### Grupo 4: Gastos Operativos (711xxx, 712xxx, 713xxx, 714xxx)
- **Cuentas**: 711011, 711021, 712011, 713011, etc.
- **Salida**: Celda `RCF!G28`
- **Propósito**: Consolidar gastos de personal, financieros y otros

#### Grupo 5: Apartados y Reservas (714011)
- **Salida**: Celda `RCF!G34`
- **Nota**: El saldo se invierte (multiplicado por -1)

#### Grupo 6: Liberaciones de Reservas (612011)
- **Salida**: Celda `RCF!G36`

#### Grupos 7-10: Impuesto Sobre la Renta
- **Cuentas**: 714013004, 714013006, 714013007, 714013008
- **Salidas**: Celdas `RCF!I47` a `RCF!I50`
- **Propósito**: Provisiones de ISLR

### 2. Anexo N°13 - Conciliación Fiscal

Genera automáticamente las siguientes secciones:

| Sección | Rango | Descripción |
|---------|-------|-------------|
| Tributos Pagados y No Causados | A8:A9 | Compara BC inicial vs BC actual |
| Tributos Causados y No Pagados | A13:A21 | Provisiones tributarias |
| Diferencia en Cambio | A25:A141 | Diferencial cambiario realizado/no realizado |
| Remuneraciones | A146:A150 | Causadas y no pagadas |
| Arrendamientos | A154 | Devengados y no cobrados |
| Gastos de Administración | A158:A162 | Causados y no pagados |
| Gastos de Conservación | A166:A170 | Causados y no pagados |
| Gastos No Deducibles | A174:A201 | Gastos rechazados fiscalmente |
| Reservas | A206:A217 | Provisiones constituidas |
| Liberaciones de Reservas | A221:A224 | Reversiones de provisiones |

**Transferencias a RCF:**
- `RCF!G30` ← `Anexo N13!F22` (Tributos)
- `RCF!G20` ← `-Anexo N13!F143` (Diferencial cambiario)
- `RCF!G29` ← `-Anexo N13!F203` (Gastos no deducibles)
- `RCF!G32` ← `-Anexo N13!F151` (Remuneraciones)

### 3. Base de Impuesto Diferido

Calcula las diferencias temporarias para impuesto diferido:

- Tributos pagados/causados
- Diferencias en cambio
- Remuneraciones
- Arrendamientos
- Gastos de administración y conservación
- Provisiones
- Impuesto Diferido Activo (IDA) - Celda A196
- Impuesto Diferido Pasivo (IDP) - Celda A201

### 4. Cálculos Tributarios Automáticos (Hoja Menu)

#### LOCTI (Ley Orgánica de Ciencia, Tecnología e Innovación)
- **Cálculo**: `RCF!D12 × 0.5%`
- **Salida**: `Menu!M6`
- **Cuenta BC**: 713571001 → `Menu!N6`

#### Impuesto Diferido
- **Cálculo**: `Base ID!E204 × 100%`
- **Salida**: `Menu!M7`
- **Cuenta BC**: 714013006 → `Menu!N7`

#### ISLR (Impuesto Sobre la Renta)
- **Cálculo**: `RCF!H59 × 34%`
- **Salida**: `Menu!M8`
- **Cuenta BC**: 213011010 → `Menu!N8`

#### Ley del Deporte
- **Cálculo**: `RCF!J51 × 1%` (solo si J51 es negativo)
- **Salida**: `Menu!M9`
- **Cuenta BC**: 714011046 → `Menu!N9`

### 5. Consolidación por Prefijos de Cuenta

La macro también consolida cuentas por sus primeros 3 dígitos:

| Prefijo | Descripción | Salida |
|---------|-------------|--------|
| **711** | Gastos de Personal | `Menu!M17` |
| **712** | Gastos Financieros | `Menu!M18` |
| **713** | Otros Gastos Operativos | `Menu!M19` |
| **714** | Apartados y Reservas | `Menu!M20` |
| **611** | Otros Ingresos Mercantiles | `Menu!M24` |
| **612** | Liberaciones de Reservas | `Menu!M25` |

## Lógica de Procesamiento

### Búsqueda de Cuentas

```vba
Para cada grupo de cuentas:
  1. Recorrer columna A o B del Balance de Comprobación
  2. Buscar coincidencias exactas con códigos de cuenta
  3. Sumar los valores de la columna D (saldo)
  4. Escribir el total en la celda de salida correspondiente
```

### Comparación BC Inicial vs BC Actual

```vba
Para cada cuenta en Anexo N13:
  1. Buscar cuenta en "BC inicial" → Escribir en columna D
  2. Buscar cuenta en "BC" actual → Escribir en columna E
  3. La diferencia se calcula automáticamente en el Anexo
```

## Notas Importantes

> [!IMPORTANT]
> - **Formato de cuentas**: Los códigos de cuenta deben ser numéricos exactos
> - **Columnas fijas**: No modificar la estructura de columnas en las hojas BC y BC inicial
> - **Celdas de salida**: Las fórmulas en las hojas de salida pueden sobrescribirse; ejecutar la macro las regenera

> [!WARNING]
> - Si una cuenta no se encuentra en el Balance de Comprobación, se asigna valor **0**
> - Algunos valores se invierten (multiplicados por -1) para ajustes fiscales
> - La macro sobrescribe datos existentes en las celdas de salida

> [!TIP]
> - Mantener respaldos del archivo antes de ejecutar la macro
> - Verificar que los Balances de Comprobación estén completos antes de ejecutar
> - Revisar los totales en la hoja Menu como validación rápida

## 🔧 Requisitos Técnicos

- **Software**: Microsoft Excel 2010 o superior
- **Macros**: Deben estar habilitadas (Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza → Configuración de macros → Habilitar todas las macros)
- **Formato**: Archivo .xlsm (Excel habilitado para macros)

## Estructura del Archivo

```
CSC_reportes v3.xlsm
├── Hoja: Menu (Panel de control)
├── Hoja: BC (Balance actual - ENTRADA)
├── Hoja: BC inicial (Balance anterior - ENTRADA)
├── Hoja: RCF (Renta Comercial y Fiscal - SALIDA)
├── Hoja: Anexo N13 (Conciliación Fiscal - SALIDA)
├── Hoja: Base impuesto Diferido (SALIDA)
└── Módulo VBA: Generar_estado_situacion_financiera()
```

## Solución de Problemas

| Problema | Solución |
|----------|----------|
| La macro no se ejecuta | Verificar que las macros estén habilitadas |
| Valores en 0 | Verificar que los códigos de cuenta existan en el BC |
| Error de referencia | Verificar que todas las hojas requeridas existan |
| Resultados incorrectos | Verificar que los datos estén en las columnas correctas |

## Soporte

Para dudas o problemas con la macro, verificar:
1. Estructura de las hojas de trabajo
2. Formato de los códigos de cuenta
3. Ubicación de los datos en las columnas especificadas

---

**Versión**: 3.0  
**Última actualización**: Enero 2026  
**Tipo**: Macro VBA para Excel
