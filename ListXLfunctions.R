if (!require(pacman)) {install.packages("pacman")}
p_load(here,
       tidyverse,
       magrittr,
       tidyxl)

fnGetFunctionsFromXL <- function(Path, sheets) {
  xlsx_cells(Path, sheets) %>% 
    add_column(Path = str_remove_all(Path, root %>% 
                                       str_replace_all("\\\\", "\\\\\\\\"))) %>%
    select(Path, sheet, address, formula) %>% 
    drop_na(formula)  %>% 
    distinct(Path, sheet, formula) %>% 
    rowwise() %>% 
    mutate(xx = list(xlex(formula))) %>% 
    unnest(xx) %>% 
    filter(type == "function") %>% 
    distinct(Path, Function = token) %>% 
    arrange(Path, Function)
}

strRoot <- paste0(readRegistry("Environment", hive = "HCU", maxdepth = 2)$OneDrive, "\\zNICE\\models\\")
strFiles <- list.files(path       = strRoot,
                       pattern    = ".xls",
                       full.names = T)

tblWBs <- tibble(Path   = strFiles,
                 sheets = NA)

tblFuns <- tblWBs %>% 
  pmap_dfr(.f  = fnGetFunctionsFromXL,
           .id = "Path")

tblFuns %>% 
  group_by(Path) %>% 
  summarise(n = n())
