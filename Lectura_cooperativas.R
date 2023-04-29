
library(readxl)
library(tidyverse)

# Obtengo una lista de los archivos en la carpeta que los contiene
archivos <- list.files("Excels/Archivos_Analisis")

# Inicializo una lista vacia para rellenar con los tibbles traidos de los archivos
tibble_combinado <- tibble()

# Itero sobre la lista de archivos
for (archivo in archivos) {

  tibble_xls <- read_xls(
    path = str_c("Excels/Archivos_Analisis/", archivo),
    sheet = 1, 
    col_names = FALSE, 
  )
  
  # encontrar fila donde se encuentra "Departamento" en la primera columna
  fila_inicio <- which(grepl("Departamento(/Partido)?", tibble_xls$...1))[1]
                   
  # extraer las filas a partir de la fila donde se encuentra "Departamento" hasta el final del conjunto de datos
  tabla_datos <- tibble_xls[fila_inicio:nrow(tibble_xls), ]
  
  #Hago que todas las primeras celdas de la tabla sean Departamento
  tabla_datos[1,1] <- "Departamento"
  
  # Seteo el nombre de las columnas
  colnames(tabla_datos) <- slice(tabla_datos, 1)
  tabla_datos <- slice(tabla_datos, -1) 
  
  # Encontrar el índice de la columna sin nombre
  col_sin_nombre <- which(is.na(colnames(tabla_datos)))
  
  # Seleccionar todas las columnas excepto la sin nombre y eliminar los registros en donde la Variable "Ente" es NA
  tabla_datos <- tabla_datos %>%
    select(-col_sin_nombre)%>%
    filter(!is.na(Ente))
  
  # Convertir columnas numéricas desde la tercera columna en adelante a enteros
  tabla_datos <- tabla_datos %>% 
    mutate(across(3:13, ~ parse_double(.x))) %>% 
    filter(!grepl("^Total", Departamento))
  
  #AGREGAR LA VARIABLE DE LA PROVINCIA/ZONA
  area <- tibble_xls %>% 
            slice(2) %>% 
            pull(1)
  
  tabla_datos <- tabla_datos %>% 
                   mutate(Zona = area)
  
  # Concatenar los datos del archivo actual a la tabla total
  tibble_combinado <- bind_rows(tibble_combinado, tabla_datos)
}

# Mostrar los resultados
tibble_combinado
