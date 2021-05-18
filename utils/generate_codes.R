library(tidyverse)
library(here)
library(ids)

# we'll use so-called proquints (PRO-nouncable QUINT-uplets),
# which are both pronoucable and easy memorable random strings
# see https://arxiv.org/abs/0901.4016v2 for more details


# make a tibble of schools
# you may use any existing database/file with the column called "name"
schools <- tibble(name = c("Testová škola 1", "Test School 2"))

# replace non_ascii characters with standard ones, replace whitespaces with underscores
schools <- schools %>% mutate(name = textclean::replace_non_ascii(name) %>%
  str_replace_all("[:blank:]", "_"))

# append ingeters to schools for later use
schools <- schools %>% rowid_to_column(var = "school_code")

# seed for reproducibility
set.seed(31024598)

# make tibble in long format, with school, class, & id
# make all combinations of vars above, so the result will be "expanded" to child-level
schools_long <- expand_grid(
  schools, # from the tibble above
  class_code = seq_len(8), # 8 classes per school
  child_code = seq_len(45) # 45 children per class
) %>% # optionally make human-readable name for the classes
  mutate(class = class_code %>% str_c(". sada"))

# generate proquints with default settings for every row (i.e. child)
schools_long <- schools_long %>% mutate(id = proquint(n = n()))

# check for code duplicates
stopifnot({
  schools_long %>%
    pull(id) %>%
    duplicated() %>%
    not() %>%
    all()
})

# create several columns for teacher's use, edit as you wish
sheet_header <- c(
  "jméno", "příjmení", "měsíc narození", "rok narození", "poznámka",
  "1. měření", "2. měření", "3. měření", "4. měření"
)
sheet_header2 <- rep(NA, length(sheet_header)) %>% set_names(sheet_header)

# split tibble into list per school with sublists per class
out <- schools_long %>%
  split(.$name) %>%
  map(~ .x %>% split(.x$class)) %>%
  map_depth(2, ~ .x %>%
    transmute("kód dítěte" = id) %>% # rename col to human-readable
    add_column(!!!sheet_header2))

# make the output folder, if not already present
if (!dir.exists(here("id_codes"))) dir.create(here("id_codes"))

# for safety reasons add .gitignore so nothing in this dir is pushed to GH
usethis::use_git_ignore("*", "id_codes")

# write the files, MS name as filename, the resulting paths are printed as well
out %>% imap(~ writexl::write_xlsx(.x, path = here("id_codes", paste0(.y, ".xlsx"))))
