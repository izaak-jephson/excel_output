### Convert tibble to named list. Defaults to take names from column 1
### and values from column 2.

tibble_to_list <- function(tibble, names = 1, values = 2) {
  setNames(
    as.list(
      tibble[[2]]), 
    tibble[[1]])
}


read_metadata <- function(filepath) {

  list (
    
    sheet_layout = 
      
      read_xlsx(filepath, 
                sheet = "sheet_layout"),
    
    table_list = 
      read_xlsx(filepath, 
                sheet = "tables"),
    
    notes = 
      read_xlsx(filepath, 
                sheet = "notes"),
    
    # convert options to named list
    options = 
      read_xlsx(filepath, 
                sheet = "options") %>% 
      tibble_to_list(),
    
    contents = 
      read_xlsx(filepath, 
                sheet = "contents") %>% 
      tibble_to_list()
  )
  
}