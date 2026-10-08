# ============================================================
# MVP - VALIDACIÓN AUTOMÁTICA DE INFORMACIÓN PARA REPORTERÍA
# ============================================================

# ------------------------------------------------------------
# 1. CARGAR LIBRERÍAS
# ------------------------------------------------------------

library(readxl)
library(dplyr)
library(openxlsx)

# ------------------------------------------------------------
# 2. CARGAR BASES DE INFORMACIÓN
# ------------------------------------------------------------

base_financiera <- read_excel("data/informacion_financiera.xlsx",skip = 3)

base_reporteria <- read_excel("data/informacion_reporteria.xlsx",skip = 3)

# ------------------------------------------------------------
# 3. REVISIÓN INICIAL DE LAS BASES
# ------------------------------------------------------------

cat("Registros base financiera:", nrow(base_financiera), "\n")
cat("Registros base reportería:", nrow(base_reporteria), "\n")

# ------------------------------------------------------------
# 4. HOMOLOGAR NOMBRES DE VARIABLES
# ------------------------------------------------------------

base_financiera <- base_financiera %>%
  rename(
    Patrimonio_Base = Patrimonio,
    Cuotas_Base = N_Cuotas,
    Valor_Cuota_Base = Valor_Cuota,
    Efectivo_Base = Efectivo)

base_reporteria <- base_reporteria %>%
  rename(
    Patrimonio_Reportado = Patrimonio,
    Cuotas_Reportadas = N_Cuotas,
    Valor_Cuota_Reportado = Valor_Cuota,
    Efectivo_Reportado = Efectivo)

# ------------------------------------------------------------
# 5. CRUZAR LAS DOS BASES
# ------------------------------------------------------------

validacion <- base_financiera %>%
  left_join(
    base_reporteria,
    by = c("ID_Fondo", "Fecha"),
    suffix = c("_Base", "_Reporteria"))

# ------------------------------------------------------------
# 6. CALCULAR DIFERENCIAS
# ------------------------------------------------------------

validacion <- validacion %>%
  mutate(
    Dif_Patrimonio = Patrimonio_Base - Patrimonio_Reportado,
    Dif_Cuotas = Cuotas_Base - Cuotas_Reportadas,
    Dif_Valor_Cuota = Valor_Cuota_Base - Valor_Cuota_Reportado,
    Dif_Efectivo = Efectivo_Base - Efectivo_Reportado)

# ------------------------------------------------------------
# 7. APLICAR REGLAS DE VALIDACIÓN
# ------------------------------------------------------------

# Tolerancia de 0,01% para diferencias en valor cuota
tolerancia_valor_cuota <- 0.0001

validacion <- validacion %>%
  mutate(

    # Control de existencia del registro
    Control_Registro = case_when(
      is.na(Nombre_Fondo_Reporteria) ~ "REVISAR",
      TRUE ~ "OK"),

    # Control de patrimonio
    Control_Patrimonio = case_when(
      is.na(Patrimonio_Reportado) ~ "REVISAR",
      abs(Dif_Patrimonio) > 0.01 ~ "REVISAR",
      TRUE ~ "OK"),

    # Control de número de cuotas
    Control_Cuotas = case_when(
      is.na(Cuotas_Reportadas) ~ "REVISAR",
      abs(Dif_Cuotas) > 0 ~ "REVISAR",
      TRUE ~ "OK"),

    # Control de valor cuota
    Control_Valor_Cuota = case_when(
      is.na(Valor_Cuota_Reportado) ~ "REVISAR",
      abs(Dif_Valor_Cuota / Valor_Cuota_Base) >
        tolerancia_valor_cuota ~ "REVISAR",
      TRUE ~ "OK"),

    # Control de efectivo
    Control_Efectivo = case_when(
      is.na(Efectivo_Reportado) ~ "REVISAR",
      abs(Dif_Efectivo) > 0.01 ~ "REVISAR",
      TRUE ~ "OK"))

# ------------------------------------------------------------
# 8. DETERMINAR ESTADO GENERAL Y MOTIVO DE REVISIÓN
# ------------------------------------------------------------

validacion <- validacion %>%
  rowwise() %>%
  mutate(
    Estado_General = if_else(
      any(c_across(starts_with("Control_")) == "REVISAR"),
      "REVISAR",
      "OK"),

    Motivo_Revision = paste(
      c(
        if (Control_Registro == "REVISAR")
          "Registro no encontrado en reportería",
        if (
          Control_Patrimonio == "REVISAR" &&
          Control_Registro == "OK")
          "Diferencia en patrimonio",
        if (
          Control_Cuotas == "REVISAR" &&
          Control_Registro == "OK")
          "Diferencia en número de cuotas",
        if (
          Control_Valor_Cuota == "REVISAR" &&
          Control_Registro == "OK")
          "Diferencia en valor cuota",
        if (
          Control_Efectivo == "REVISAR" &&
          Control_Registro == "OK")
          "Diferencia o dato faltante en efectivo"),collapse = "; ")) %>%
  ungroup()

# ------------------------------------------------------------
# 9. GENERAR REPORTE DE EXCEPCIONES
# ------------------------------------------------------------

excepciones <- validacion %>%
  filter(Estado_General == "REVISAR")

# ------------------------------------------------------------
# 10. GENERAR RESUMEN DEL PROCESO
# ------------------------------------------------------------

fondos_analizados <- nrow(validacion)

fondos_ok <- sum(
  validacion$Estado_General == "OK")

fondos_revisar <- sum(
  validacion$Estado_General == "REVISAR")

total_excepciones <- sum(
  validacion$Control_Registro == "REVISAR",
  validacion$Control_Patrimonio == "REVISAR",
  validacion$Control_Cuotas == "REVISAR",
  validacion$Control_Valor_Cuota == "REVISAR",
  validacion$Control_Efectivo == "REVISAR")

resumen <- data.frame(
  Indicador = c(
    "Fondos analizados",
    "Fondos sin diferencias",
    "Fondos con excepciones",
    "Total de excepciones detectadas"),
  Resultado = c(
    fondos_analizados,
    fondos_ok,
    fondos_revisar,
    total_excepciones))

# ------------------------------------------------------------
# 11. MOSTRAR RESULTADOS EN CONSOLA
# ------------------------------------------------------------

cat("\n")
cat("========================================\n")
cat("     RESUMEN DE VALIDACIÓN\n")
cat("========================================\n")

print(resumen)

cat("\nFondos que requieren revisión:\n")

print(
  excepciones %>%
    select(
      ID_Fondo,
      Nombre_Fondo_Base,
      Motivo_Revision))

# ------------------------------------------------------------
# 12. CREAR CARPETA DE RESULTADOS
# ------------------------------------------------------------

if (!dir.exists("output")) {dir.create("output")}

# ------------------------------------------------------------
# 13. CREAR REPORTE EN EXCEL
# ------------------------------------------------------------

wb <- createWorkbook()

# Hoja 1: Resumen
addWorksheet(wb, "Resumen")
writeData(
  wb,
  "Resumen",
  resumen)

# Hoja 2: Validación completa
addWorksheet(wb, "Validacion_Completa")
writeData(
  wb,
  "Validacion_Completa",
  validacion)

# Hoja 3: Excepciones
addWorksheet(wb, "Excepciones")
writeData(
  wb,
  "Excepciones",
  excepciones)

# ------------------------------------------------------------
# 14. FORMATO BÁSICO DEL REPORTE
# ------------------------------------------------------------

# Formato simple para los encabezados
estilo_encabezado <- createStyle(
  textDecoration = "bold",
  fgFill = "#D9EAF7",
  border = "Bottom")

# Aplicar formato a los encabezados
addStyle(
  wb,
  "Resumen",
  estilo_encabezado,
  rows = 1,
  cols = 1:2,
  gridExpand = TRUE)

addStyle(
  wb,
  "Validacion_Completa",
  estilo_encabezado,
  rows = 1,
  cols = 1:ncol(validacion),
  gridExpand = TRUE)

addStyle(
  wb,
  "Excepciones",
  estilo_encabezado,
  rows = 1,
  cols = 1:ncol(excepciones),
  gridExpand = TRUE)

# Agregar filtros
addFilter(
  wb,
  "Validacion_Completa",
  rows = 1,
  cols = 1:ncol(validacion))

addFilter(
  wb,
  "Excepciones",
  rows = 1,
  cols = 1:ncol(excepciones))

# Congelar primera fila
freezePane(
  wb,
  "Validacion_Completa",
  firstRow = TRUE)

freezePane(
  wb,
  "Excepciones",
  firstRow = TRUE)

# Ajustar ancho de columnas
setColWidths(
  wb,
  "Resumen",
  cols = 1:2,
  widths = "auto")

setColWidths(
  wb,
  "Validacion_Completa",
  cols = 1:ncol(validacion),
  widths = "auto")

setColWidths(
  wb,
  "Excepciones",
  cols = 1:ncol(excepciones),
  widths = "auto")

# ------------------------------------------------------------
# 15. GUARDAR ARCHIVO FINAL
# ------------------------------------------------------------

saveWorkbook(
  wb,
  "output/reporte_validacion.xlsx",
  overwrite = TRUE)

# ------------------------------------------------------------
# 16. MENSAJE FINAL
# ------------------------------------------------------------

cat("\nProceso finalizado correctamente.\n")
cat("Reporte generado en: output/reporte_validacion.xlsx\n")