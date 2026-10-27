# ============================================================
# app.R
# Aplicación Shiny para analizar notas_estudiantes.xlsx
# Datos FICTICIOS de demostración.
# ============================================================

library(shiny)
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(DT)
library(writexl)

# ============================================================
# Funciones reutilizables
# ============================================================

hojas_obligatorias <- c("estudiantes","notas","parametros","diccionario")
cols_obligatorias <- list(
  estudiantes = c("id_estudiante","nombre_ficticio","seccion"),
  notas       = c("id_estudiante","asignatura","periodo",
                  "evaluacion","peso","nota"),
  parametros  = c("parametro","valor"),
  diccionario = c("hoja","variable","descripcion",
                  "tipo_dato","valores_permitidos")
)

# Lee el Excel y devuelve lista con hojas o errores acumulados
leer_excel <- function(ruta) {
  errores <- character(0)
  hojas_presentes <- tryCatch(readxl::excel_sheets(ruta),
                              error = function(e) character(0))
  faltantes <- setdiff(hojas_obligatorias, hojas_presentes)
  if (length(faltantes) > 0) {
    errores <- c(errores, paste0("Faltan hojas: ", paste(faltantes, collapse = ", ")))
  }
  datos <- list()
  for (h in intersect(hojas_obligatorias, hojas_presentes)) {
    datos[[h]] <- tryCatch(
      readxl::read_excel(ruta, sheet = h),
      error = function(e) {
        errores <<- c(errores, paste0("No se pudo leer hoja '", h, "': ", conditionMessage(e)))
        NULL
      }
    )
    if (!is.null(datos[[h]])) {
      faltan_cols <- setdiff(cols_obligatorias[[h]], names(datos[[h]]))
      if (length(faltan_cols) > 0) {
        errores <- c(errores, paste0("Hoja '", h, "' sin columnas: ",
                                     paste(faltan_cols, collapse = ", ")))
      }
    }
  }
  list(datos = datos, errores = errores)
}

# Validación completa. Devuelve lista con errores y advertencias.
validar_datos <- function(d) {
  errores <- character(0); avisos <- character(0)
  req_h <- hojas_obligatorias
  if (!all(req_h %in% names(d))) {
    errores <- c(errores, "Estructura incompleta de hojas.")
    return(list(errores = errores, avisos = avisos))
  }
  
  est <- d$estudiantes; not <- d$notas; par <- d$parametros
  nota_min <- suppressWarnings(as.numeric(par$valor[par$parametro == "nota_minima"][1]))
  nota_max <- suppressWarnings(as.numeric(par$valor[par$parametro == "nota_maxima"][1]))
  if (is.na(nota_min) || is.na(nota_max))
    errores <- c(errores, "Parámetros nota_minima / nota_maxima no numéricos.")
  
  # IDs únicos
  if (anyDuplicated(est$id_estudiante))
    errores <- c(errores, "La hoja 'estudiantes' contiene id_estudiante duplicado.")
  
  # Llave foránea
  huerfanos <- setdiff(unique(not$id_estudiante), unique(est$id_estudiante))
  if (length(huerfanos) > 0)
    errores <- c(errores, paste0("Estudiantes en 'notas' ausentes en 'estudiantes': ",
                                 paste(head(huerfanos, 5), collapse = ", "),
                                 if (length(huerfanos) > 5) " ..." else ""))
  
  # Claves duplicadas
  clave <- not[, c("id_estudiante","asignatura","periodo","evaluacion")]
  if (anyDuplicated(clave))
    errores <- c(errores, "Hay claves duplicadas (id + asignatura + periodo + evaluacion).")
  
  # Notas no numéricas y fuera de escala
  nota_num <- suppressWarnings(as.numeric(not$nota))
  no_num   <- which(!is.na(not$nota) & is.na(nota_num))
  if (length(no_num) > 0)
    errores <- c(errores, paste0("Notas no numéricas en ", length(no_num), " filas."))
  if (!is.na(nota_min) && !is.na(nota_max)) {
    fuera <- which(!is.na(nota_num) & (nota_num < nota_min | nota_num > nota_max))
    if (length(fuera) > 0)
      errores <- c(errores, paste0("Notas fuera de escala en ", length(fuera), " filas."))
  }
  
  # Pesos
  peso_num <- suppressWarnings(as.numeric(not$peso))
  if (any(is.na(peso_num) & !is.na(not$peso)))
    errores <- c(errores, "Pesos no numéricos detectados.")
  if (any(peso_num <= 0, na.rm = TRUE))
    errores <- c(errores, "Hay pesos no positivos.")
  suma <- not %>% mutate(peso_num = peso_num) %>%
    group_by(id_estudiante, asignatura, periodo) %>%
    summarise(s = sum(peso_num, na.rm = TRUE), .groups = "drop")
  malos <- which(abs(suma$s - 1) > 1e-6)
  if (length(malos) > 0)
    errores <- c(errores, paste0("Pesos que no suman 1 en ",
                                 length(malos), " combinaciones."))
  
  # Evaluaciones esperadas ausentes (registro completo ausente, no solo NA)
  asigs <- sort(unique(not$asignatura)); pers <- sort(unique(not$periodo))
  evals <- sort(unique(not$evaluacion))
  ids   <- unique(est$id_estudiante)
  grid  <- expand_grid(id_estudiante = ids, asignatura = asigs,
                       periodo = pers, evaluacion = evals)
  comb_actual <- not %>% distinct(id_estudiante, asignatura, periodo, evaluacion)
  faltan_filas <- anti_join(grid, comb_actual,
                            by = c("id_estudiante","asignatura","periodo","evaluacion"))
  if (nrow(faltan_filas) > 0)
    errores <- c(errores, paste0("Faltan ", nrow(faltan_filas),
                                 " registros esperados en 'notas' (combinación estudiante/asig/per/ev)."))
  
  list(errores = errores, avisos = avisos)
}

# Promedio ponderado por (estudiante, asignatura, periodo)
# Regla: suma(nota*peso) / suma(pesos con nota disponible). NA si todas faltan.
calcular_promedios_asig <- function(not, nota_aprob) {
  not %>%
    mutate(nota_num = suppressWarnings(as.numeric(nota)),
           peso_num = suppressWarnings(as.numeric(peso))) %>%
    group_by(id_estudiante, asignatura, periodo) %>%
    summarise(
      n_eval_total    = n(),
      n_eval_valid    = sum(!is.na(nota_num)),
      n_eval_falta    = sum(is.na(nota_num)),
      suma_nota_peso  = sum(nota_num * peso_num, na.rm = TRUE),
      suma_peso_valid = sum(peso_num[!is.na(nota_num)]),
      promedio = ifelse(suma_peso_valid > 0,
                        suma_nota_peso / suma_peso_valid, NA_real_),
      estado   = case_when(
        n_eval_falta == 0 & !is.na(promedio) ~ "completo",
        is.na(promedio)                      ~ "sin_datos",
        TRUE                                 ~ "provisional"
      ),
      aprobado_def = if_else(estado == "completo",
                             promedio >= nota_aprob, NA),
      .groups = "drop"
    )
}

# Promedio general por estudiante y período (peso igual a cada asignatura)
# Completo solo si TODAS las asignaturas están completas (sin faltantes).
calcular_promedio_general <- function(prom_asig) {
  prom_asig %>%
    group_by(id_estudiante, periodo) %>%
    summarise(
      n_asig_total    = n(),
      n_asig_valid    = sum(!is.na(promedio)),
      n_asig_completa = sum(estado == "completo"),
      prom_general    = ifelse(n_asig_valid > 0,
                               mean(promedio[!is.na(promedio)]), NA_real_),
      estado_general  = case_when(
        n_asig_completa == n_asig_total ~ "completo",
        n_asig_valid == 0               ~ "sin_datos",
        TRUE                            ~ "provisional"
      ),
      .groups = "drop"
    )
}

# Estadísticos descriptivos (recibe el vector completo, con NA;
# calcula los faltantes internamente)
describir <- function(x) {
  n_falt <- sum(is.na(x))
  x <- x[!is.na(x)]
  n <- length(x)
  if (n == 0) {
    return(tibble(n_validas = 0L, n_faltantes = n_falt, media = NA_real_,
                  mediana = NA_real_, moda = NA_character_, minimo = NA_real_,
                  maximo = NA_real_, rango = NA_real_, varianza = NA_real_,
                  sd = NA_real_, q1 = NA_real_, q3 = NA_real_, iqr = NA_real_))
  }
  tb <- table(x)
  m <- as.numeric(names(tb)[tb == max(tb)])
  moda_txt <- if (length(m) == length(tb) && length(tb) > 1) {
    "No informativa (todos los valores tienen la misma frecuencia)"
  } else if (length(m) > 1) {
    paste0("Empate: ", paste(m, collapse = ", "))
  } else {
    as.character(m)
  }
  tibble(
    n_validas = n,
    n_faltantes = n_falt,
    media = mean(x), mediana = median(x),
    moda = moda_txt,
    minimo = min(x), maximo = max(x),
    rango = max(x) - min(x),
    varianza = if (n >= 2) var(x) else NA_real_,
    sd = if (n >= 2) sd(x) else NA_real_,
    q1 = if (n >= 2) quantile(x, 0.25, names = FALSE) else NA_real_,
    q3 = if (n >= 2) quantile(x, 0.75, names = FALSE) else NA_real_,
    iqr = if (n >= 2) IQR(x) else NA_real_
  )
}

# Descriptivos por grupo (df, nombre de columna de valor, nombre de columna de grupo)
desc_grupo <- function(df, col, grupo) {
  df %>%
    group_by(across(all_of(grupo))) %>%
    group_modify(~ describir(.x[[col]])) %>%
    ungroup()
}

# Redondeo de columnas numéricas para visualización
round_df <- function(df, digits = 3) {
  df %>% mutate(across(where(is.numeric), ~ round(.x, digits)))
}

# ============================================================
# UI
# ============================================================
ui <- fluidPage(
  titlePanel("Análisis de calificaciones — datos simulados (ficticios)"),
  sidebarLayout(
    sidebarPanel(
      fileInput("archivo", "Subir archivo Excel (notas_estudiantes.xlsx)",
                accept = c(".xlsx", ".xls")),
      helpText("Los datos son ficticios y solo para demostración."),
      hr(),
      uiOutput("ui_filtros")
    ),
    mainPanel(
      tabsetPanel(
        tabPanel("Resumen", h4("Indicadores"), tableOutput("tab_resumen"),
                 verbatimTextOutput("msg_estado")),
        tabPanel("Estadística descriptiva",
                 h4("General"), DTOutput("tab_desc_gen"),
                 h4("Por asignatura"), DTOutput("tab_desc_asig"),
                 h4("Por sección"),    DTOutput("tab_desc_sec"),
                 h4("Por período"),    DTOutput("tab_desc_per")),
        tabPanel("Gráficos",
                 plotOutput("g_hist", height = 380),
                 plotOutput("g_box_asig", height = 380),
                 plotOutput("g_box_sec", height = 380),
                 plotOutput("g_barras_asig", height = 380),
                 plotOutput("g_evol", height = 380),
                 plotOutput("g_aprob", height = 380)),
        tabPanel("Resultados por estudiante",
                 DTOutput("tab_estudiantes"),
                 h4("Promedios por asignatura y período"),
                 DTOutput("tab_prom_asig"),
                 h4("Promedio general por período"),
                 DTOutput("tab_prom_gen")),
        tabPanel("Calidad de datos",
                 h4("Errores de validación"),
                 verbatimTextOutput("val_errores"),
                 h4("Advertencias"),
                 verbatimTextOutput("val_avisos")),
        tabPanel("Tabla y descargas",
                 h4("Datos filtrados — notas individuales"),
                 DTOutput("tab_notas_filtradas"),
                 h4("Datos filtrados — promedios por asignatura/período"),
                 DTOutput("tab_prom_filtrados"),
                 hr(),
                 downloadButton("dl_filtrado", "Descargar datos filtrados (Excel)"),
                 downloadButton("dl_resumen",  "Descargar resúmenes estadísticos (Excel)"),
                 downloadButton("dl_est",      "Descargar resultados por estudiante (Excel)"))
      )
    )
  )
)

# ============================================================
# SERVER
# ============================================================
server <- function(input, output, session) {
  
  output$msg_estado <- renderText({
    if (is.null(input$archivo)) "Debe cargar el archivo Excel para comenzar."
    else ""
  })
  
  # ---- Carga ----
  datos_raw <- reactive({
    req(input$archivo)
    leer_excel(input$archivo$datapath)
  })
  
  # ---- Validación ----
  validacion <- reactive({
    r <- datos_raw()
    if (length(r$errores) > 0 && !all(hojas_obligatorias %in% names(r$datos)))
      return(list(errores = r$errores, avisos = character(0)))
    v <- validar_datos(r$datos)
    list(errores = c(r$errores, v$errores), avisos = v$avisos)
  })
  
  output$val_errores <- renderText({
    v <- validacion()
    if (length(v$errores) == 0) "Sin errores detectados."
    else paste(v$errores, collapse = "\n")
  })
  output$val_avisos <- renderText({
    v <- validacion()
    if (length(v$avisos) == 0) "Sin advertencias."
    else paste(v$avisos, collapse = "\n")
  })
  
  # Bloquea cálculos si hay errores
  datos_ok <- reactive({
    v <- validacion()
    validate(need(length(v$errores) == 0,
                  "Corrija los errores de la pestaña 'Calidad de datos'."))
    datos_raw()$datos
  })
  
  # ---- Parámetros ----
  params <- reactive({
    d <- datos_ok()
    par <- d$parametros
    list(
      nota_min   = as.numeric(par$valor[par$parametro == "nota_minima"][1]),
      nota_max   = as.numeric(par$valor[par$parametro == "nota_maxima"][1]),
      nota_aprob = as.numeric(par$valor[par$parametro == "nota_aprobatoria"][1])
    )
  })
  
  # ---- Cálculos base ----
  notas_prep <- reactive({
    d <- datos_ok()
    d$notas %>%
      mutate(nota = suppressWarnings(as.numeric(nota)),
             peso = suppressWarnings(as.numeric(peso))) %>%
      left_join(d$estudiantes, by = "id_estudiante")
  })
  
  prom_asig <- reactive({
    calcular_promedios_asig(datos_ok()$notas, params()$nota_aprob)
  })
  
  prom_gen <- reactive({
    calcular_promedio_general(prom_asig())
  })
  
  # ---- Filtros UI ----
  output$ui_filtros <- renderUI({
    req(datos_ok())
    est <- datos_ok()$estudiantes
    not <- datos_ok()$notas
    tagList(
      selectInput("f_sec", "Sección",
                  choices = c("Todos", sort(unique(est$seccion))),
                  selected = "Todos"),
      selectInput("f_asig", "Asignatura",
                  choices = c("Todos", sort(unique(not$asignatura))),
                  selected = "Todos"),
      selectInput("f_per", "Período",
                  choices = c("Todos", sort(unique(not$periodo))),
                  selected = "Todos"),
      selectInput("f_est", "Estudiante",
                  choices = c("Todos", sort(unique(est$id_estudiante))),
                  selected = "Todos"),
      radioButtons("f_nivel", "Nivel de análisis",
                   choices = c("Evaluaciones (notas individuales)" = "eval",
                               "Promedios por asignatura y período"  = "prom"),
                   selected = "eval")
    )
  })
  
  # ---- Aplicar filtros ----
  notas_filtradas <- reactive({
    not <- notas_prep()
    if (!is.null(input$f_sec)  && input$f_sec  != "Todos") not <- not %>% filter(seccion == input$f_sec)
    if (!is.null(input$f_asig) && input$f_asig != "Todos") not <- not %>% filter(asignatura == input$f_asig)
    if (!is.null(input$f_per)  && input$f_per  != "Todos") not <- not %>% filter(periodo == input$f_per)
    if (!is.null(input$f_est)  && input$f_est  != "Todos") not <- not %>% filter(id_estudiante == input$f_est)
    not
  })
  
  prom_asig_filtrado <- reactive({
    pr <- prom_asig() %>% left_join(datos_ok()$estudiantes, by = "id_estudiante")
    if (!is.null(input$f_sec)  && input$f_sec  != "Todos") pr <- pr %>% filter(seccion == input$f_sec)
    if (!is.null(input$f_asig) && input$f_asig != "Todos") pr <- pr %>% filter(asignatura == input$f_asig)
    if (!is.null(input$f_per)  && input$f_per  != "Todos") pr <- pr %>% filter(periodo == input$f_per)
    if (!is.null(input$f_est)  && input$f_est  != "Todos") pr <- pr %>% filter(id_estudiante == input$f_est)
    pr
  })
  
  prom_gen_filtrado <- reactive({
    pg <- prom_gen() %>% left_join(datos_ok()$estudiantes, by = "id_estudiante")
    if (!is.null(input$f_sec) && input$f_sec != "Todos") pg <- pg %>% filter(seccion == input$f_sec)
    if (!is.null(input$f_per) && input$f_per != "Todos") pg <- pg %>% filter(periodo == input$f_per)
    if (!is.null(input$f_est) && input$f_est != "Todos") pg <- pg %>% filter(id_estudiante == input$f_est)
    pg
  })
  
  # Serie "según nivel" (para resumen)
  serie_nivel <- reactive({
    if (input$f_nivel == "eval") {
      notas_filtradas() %>% filter(!is.na(nota)) %>% pull(nota)
    } else {
      prom_asig_filtrado() %>% filter(!is.na(promedio)) %>% pull(promedio)
    }
  })
  
  # ---- Indicadores ----
  output$tab_resumen <- renderTable({
    req(input$f_nivel)
    req(nrow(notas_filtradas()) > 0)
    
    not <- notas_filtradas()
    pg  <- prom_gen_filtrado()
    n_est <- not %>% distinct(id_estudiante) %>% nrow()
    n_reg <- nrow(not)
    n_falt <- sum(is.na(not$nota))
    pct_falt <- if (n_reg > 0) round(100 * n_falt / n_reg, 2) else NA
    x <- serie_nivel()
    media   <- if (length(x) > 0) round(mean(x), 3) else NA
    mediana <- if (length(x) > 0) round(median(x), 3) else NA
    sdv     <- if (length(x) >= 2) round(sd(x), 3) else NA
    
    # Clasificación solo sobre resultados completos (promedio general)
    completos <- pg %>% filter(estado_general == "completo", !is.na(prom_general))
    n_completos <- nrow(completos)
    n_prov  <- sum(pg$estado_general == "provisional")
    n_aprob <- sum(completos$prom_general >= params()$nota_aprob)
    n_desap <- sum(completos$prom_general <  params()$nota_aprob)
    pct_aprob <- if (n_completos > 0) round(100 * n_aprob / n_completos, 2) else NA
    pct_desap <- if (n_completos > 0) round(100 * n_desap / n_completos, 2) else NA
    excluidos <- nrow(pg) - n_completos
    
    data.frame(
      indicador = c(
        "Nivel de análisis",
        "Estudiantes únicos",
        "Notas registradas (filas)",
        "Notas faltantes (N)",
        "Notas faltantes (%) [denom. = filas filtradas]",
        "Media de la unidad analizada",
        "Mediana de la unidad analizada",
        "Desviación estándar (muestral)",
        "Resultados completos (promedio general)",
        "Resultados provisionales (promedio general)",
        "Aprobados (% sobre completos)",
        "Desaprobados (% sobre completos)",
        "Resultados excluidos de la clasificación"
      ),
      valor = as.character(c(
        ifelse(input$f_nivel == "eval", "Evaluaciones", "Promedios por asignatura/período"),
        n_est, n_reg, n_falt, pct_falt,
        media, mediana, sdv,
        n_completos, n_prov,
        paste0(n_aprob, " (", ifelse(is.na(pct_aprob), "NA", pct_aprob), "%)"),
        paste0(n_desap, " (", ifelse(is.na(pct_desap), "NA", pct_desap), "%)"),
        excluidos
      ))
    )
  }, rownames = FALSE)
  
  # ---- Estadística descriptiva ----
  # Se calcula sobre la serie del nivel analizado (eval o prom).
  
  output$tab_desc_gen <- renderDT({
    req(input$f_nivel)
    d <- if (input$f_nivel == "eval") describir(notas_filtradas()$nota)
    else describir(prom_asig_filtrado()$promedio)
    datatable(round_df(d), options = list(dom = "t"), rownames = FALSE)
  })
  
  tabla_por_grupo <- function(grupo) {
    if (input$f_nivel == "eval") {
      desc_grupo(notas_filtradas(), "nota", grupo)
    } else {
      desc_grupo(prom_asig_filtrado(), "promedio", grupo)
    }
  }
  
  output$tab_desc_asig <- renderDT({
    req(input$f_nivel)
    datatable(round_df(tabla_por_grupo("asignatura")),
              options = list(dom = "t"), rownames = FALSE)
  })
  output$tab_desc_sec <- renderDT({
    req(input$f_nivel)
    datatable(round_df(tabla_por_grupo("seccion")),
              options = list(dom = "t"), rownames = FALSE)
  })
  output$tab_desc_per <- renderDT({
    req(input$f_nivel)
    datatable(round_df(tabla_por_grupo("periodo")),
              options = list(dom = "t"), rownames = FALSE)
  })
  
  # ---- Gráficos ----
  # Todos respetan filtros y el nivel (eval/prom).
  
  output$g_hist <- renderPlot({
    req(input$f_nivel)
    if (input$f_nivel == "eval") {
      df <- notas_filtradas() %>% filter(!is.na(nota))
      validate(need(nrow(df) > 0, "Sin datos para el histograma."))
      ggplot(df, aes(x = nota)) +
        geom_histogram(binwidth = 1, boundary = 0,
                       fill = "steelblue", color = "white") +
        labs(title = "Histograma de notas individuales (evaluaciones)",
             x = "Nota (0-20)", y = "Frecuencia") +
        scale_x_continuous(limits = c(params()$nota_min, params()$nota_max))
    } else {
      df <- prom_asig_filtrado() %>% filter(!is.na(promedio))
      validate(need(nrow(df) > 0, "Sin datos para el histograma."))
      ggplot(df, aes(x = promedio)) +
        geom_histogram(binwidth = 1, boundary = 0,
                       fill = "darkorange", color = "white") +
        labs(title = "Histograma de promedios ponderados por asignatura y período",
             x = "Promedio (0-20)", y = "Frecuencia") +
        scale_x_continuous(limits = c(params()$nota_min, params()$nota_max))
    }
  })
  
  output$g_box_asig <- renderPlot({
    req(input$f_nivel)
    if (input$f_nivel == "eval") {
      df <- notas_filtradas() %>% filter(!is.na(nota))
      validate(need(nrow(df) > 0, "Sin datos."))
      ggplot(df, aes(x = asignatura, y = nota, fill = asignatura)) +
        geom_boxplot(show.legend = FALSE) +
        labs(title = "Diagrama de cajas — notas individuales por asignatura",
             x = NULL, y = "Nota (0-20)")
    } else {
      df <- prom_asig_filtrado() %>% filter(!is.na(promedio))
      validate(need(nrow(df) > 0, "Sin datos."))
      ggplot(df, aes(x = asignatura, y = promedio, fill = asignatura)) +
        geom_boxplot(show.legend = FALSE) +
        labs(title = "Diagrama de cajas — promedios por asignatura y período",
             x = NULL, y = "Promedio (0-20)")
    }
  })
  
  output$g_box_sec <- renderPlot({
    req(input$f_nivel)
    if (input$f_nivel == "eval") {
      df <- notas_filtradas() %>% filter(!is.na(nota))
      validate(need(nrow(df) > 0, "Sin datos."))
      ggplot(df, aes(x = seccion, y = nota, fill = seccion)) +
        geom_boxplot(show.legend = FALSE) +
        labs(title = "Diagrama de cajas — notas individuales por sección",
             x = "Sección", y = "Nota (0-20)")
    } else {
      df <- prom_asig_filtrado() %>% filter(!is.na(promedio))
      validate(need(nrow(df) > 0, "Sin datos."))
      ggplot(df, aes(x = seccion, y = promedio, fill = seccion)) +
        geom_boxplot(show.legend = FALSE) +
        labs(title = "Diagrama de cajas — promedios por sección",
             x = "Sección", y = "Promedio (0-20)")
    }
  })
  
  output$g_barras_asig <- renderPlot({
    req(input$f_nivel)
    if (input$f_nivel == "eval") {
      df <- notas_filtradas() %>% filter(!is.na(nota))
    } else {
      df <- prom_asig_filtrado() %>% filter(!is.na(promedio)) %>%
        rename(nota = promedio)
    }
    validate(need(nrow(df) > 0, "Sin datos."))
    res <- df %>%
      group_by(asignatura) %>%
      summarise(media = mean(nota), n = n(), .groups = "drop")
    ggplot(res, aes(x = reorder(asignatura, media), y = media)) +
      geom_col(fill = "steelblue") +
      geom_text(aes(label = paste0("n=", n)), vjust = -0.4, size = 3.5) +
      labs(title = paste0("Promedio por asignatura (",
                          ifelse(input$f_nivel == "eval",
                                 "notas individuales", "promedios"), ")"),
           x = NULL, y = "Promedio") +
      ylim(0, params()$nota_max)
  })
  
  output$g_evol <- renderPlot({
    req(input$f_nivel)
    if (input$f_nivel == "eval") {
      df <- notas_filtradas() %>% filter(!is.na(nota))
    } else {
      df <- prom_asig_filtrado() %>% filter(!is.na(promedio)) %>%
        rename(nota = promedio)
    }
    validate(need(nrow(df) > 0, "Sin datos."))
    res <- df %>% group_by(periodo) %>%
      summarise(media = mean(nota), n = n(), .groups = "drop")
    ggplot(res, aes(x = periodo, y = media, group = 1)) +
      geom_line(color = "darkred") +
      geom_point(size = 3, color = "darkred") +
      geom_text(aes(label = paste0("n=", n)), vjust = -1, size = 3.5) +
      labs(title = paste0("Evolución del promedio por período (",
                          ifelse(input$f_nivel == "eval",
                                 "notas individuales", "promedios"), ")"),
           x = "Período", y = "Promedio") +
      ylim(0, params()$nota_max)
  })
  
  output$g_aprob <- renderPlot({
    pg <- prom_gen_filtrado()
    validate(need(nrow(pg) > 0, "Sin datos."))
    res <- pg %>%
      filter(estado_general == "completo", !is.na(prom_general)) %>%
      mutate(clasif = ifelse(prom_general >= params()$nota_aprob,
                             "Aprobado", "Desaprobado")) %>%
      count(clasif)
    validate(need(nrow(res) > 0,
                  "No hay resultados completos para clasificar con los filtros actuales."))
    ggplot(res, aes(x = clasif, y = n, fill = clasif)) +
      geom_col(show.legend = FALSE) +
      geom_text(aes(label = n), vjust = -0.4) +
      labs(title = "Aprobados y desaprobados (solo resultados completos, promedio general)",
           x = NULL, y = "Cantidad de estudiantes-período") +
      scale_fill_manual(values = c("Aprobado" = "#2e7d32",
                                   "Desaprobado" = "#c62828"))
  })
  
  # ---- Resultados individuales ----
  output$tab_estudiantes <- renderDT({
    df <- notas_filtradas() %>%
      mutate(nota_mostrar = ifelse(is.na(nota), "(faltante)",
                                   format(round(nota, 2), nsmall = 2))) %>%
      select(id_estudiante, nombre_ficticio, seccion,
             asignatura, periodo, evaluacion, peso, nota = nota_mostrar)
    datatable(df, options = list(pageLength = 10), rownames = FALSE)
  })
  
  output$tab_prom_asig <- renderDT({
    df <- prom_asig_filtrado() %>%
      mutate(promedio = round(promedio, 2)) %>%
      select(id_estudiante, seccion, asignatura, periodo,
             n_eval_total, n_eval_valid, n_eval_falta,
             promedio, estado, aprobado_def)
    datatable(df, options = list(pageLength = 10), rownames = FALSE) %>%
      formatStyle("estado",
                  backgroundColor = styleEqual(
                    c("completo","provisional","sin_datos"),
                    c("#d0f0d0","#fff3cd","#f8d7da")))
  })
  
  output$tab_prom_gen <- renderDT({
    df <- prom_gen_filtrado() %>%
      mutate(prom_general = round(prom_general, 2)) %>%
      select(id_estudiante, seccion, periodo,
             n_asig_total, n_asig_valid, n_asig_completa,
             prom_general, estado_general)
    datatable(df, options = list(pageLength = 10), rownames = FALSE)
  })
  
  # ---- Tabla y descargas ----
  output$tab_notas_filtradas <- renderDT({
    df <- notas_filtradas() %>%
      select(id_estudiante, nombre_ficticio, seccion,
             asignatura, periodo, evaluacion, peso, nota)
    datatable(df, options = list(pageLength = 10), rownames = FALSE)
  })
  
  output$tab_prom_filtrados <- renderDT({
    df <- prom_asig_filtrado() %>%
      mutate(promedio = round(promedio, 2)) %>%
      select(id_estudiante, seccion, asignatura, periodo,
             n_eval_total, n_eval_valid, n_eval_falta,
             promedio, estado, aprobado_def)
    datatable(df, options = list(pageLength = 10), rownames = FALSE)
  })
  
  output$dl_filtrado <- downloadHandler(
    filename = function() paste0("datos_filtrados_", Sys.Date(), ".xlsx"),
    content = function(file) {
      writexl::write_xlsx(
        list(notas = notas_filtradas(),
             promedios_asig = prom_asig_filtrado()),
        path = file)
    }
  )
  
  output$dl_resumen <- downloadHandler(
    filename = function() paste0("resumenes_estadisticos_", Sys.Date(), ".xlsx"),
    content = function(file) {
      if (input$f_nivel == "eval") {
        df <- notas_filtradas(); col <- "nota"
      } else {
        df <- prom_asig_filtrado(); col <- "promedio"
      }
      writexl::write_xlsx(
        list(general        = describir(df[[col]]),
             por_asignatura = desc_grupo(df, col, "asignatura"),
             por_seccion    = desc_grupo(df, col, "seccion"),
             por_periodo    = desc_grupo(df, col, "periodo")),
        path = file)
    }
  )
  
  output$dl_est <- downloadHandler(
    filename = function() paste0("resultados_estudiantes_", Sys.Date(), ".xlsx"),
    content = function(file) {
      writexl::write_xlsx(
        list(promedios_asignatura = prom_asig_filtrado(),
             promedio_general = prom_gen_filtrado(),
             estudiantes = datos_ok()$estudiantes),
        path = file)
    }
  )
}

shinyApp(ui, server)