### Turn value to numeric where possible, i.e. ignore anything that cannot be
### parsed as a number
as_numeric_if_number <- function(string){
  if(is.numeric(x <- suppressWarnings(type.convert(string)))) x 
  else string
}



### Turn [c] figures in given table to negatives for easier formatting for given table

c_to_negative <- function(table){
  
  table %>% mutate(across(everything(), ~ str_replace(.,"\\[c\\]","-1")))

}

### Turn columns in given table to numeric where possible
cols_to_numeric <- function(table) {
  
  table %>% mutate(across(everything(), as_numeric_if_number))
  
}

### Apply suppression formatting to all tables in list and turn columns to 
### numeric where possible

tables_to_numeric <- function(table_list) {
  
  table_list %>% 
    map(c_to_negative) %>% 
    map(cols_to_numeric)
  
}