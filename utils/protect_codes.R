library(tidyverse)
library(openxlsx)
library(here)

# WARNING!
# locking cells is highly ineffective
# and thus the operation may take a few minutes

# set-up the password
pswrd <- rstudioapi::askForSecret(
  title = "Spreadsheet password",
  message = "Enter password for a spreadsheet protection:",
  name = "pswrd"
)

# tip: || or && are for short-curcuit evaluation (second arg evaluated
# only if the first one is not sufficient to determine expression value)
if (!exists("pswrd") || is.null(pswrd)) stop("You have to set-up the password!")

# write it to your .Renviron for further use (FOR THIS SESSION ONLY!)
Sys.setenv(ID_CODES_PW = paste0(pswrd))

file_list <- list.files(here("id_codes"), full.names = TRUE)

# just progress bar
pb <- progress::progress_bar$new(
  total = length(file_list),
  format = "[:bar] :percent ~ :eta",
  callback = function(x) {
    message(
      "Everything protected!\n",
      "The password is `", pswrd,
      "` and is stored in your Environment."
    )
  }
)

file_list %>% walk(function(school_file) {
  pb$tick()
  Sys.sleep(0.01) # update progres bar

  # probably due to encoding, openxlsx cannon read the content, but it does
  # a good job concerning the structure (sheets), so we'll use it to load xlsx
  wb <- loadWorkbook(school_file)

  # get the data from respective sheet
  data_from_excel <-
    map(
      sheets(wb),
      ~ readxl::read_excel(school_file, sheet = .x)
    )

  # data dims
  data_dims <- data_from_excel[[1]] %>% dim()

  # write the data where it belongs
  walk2(
    names(wb), data_from_excel,
    ~ writeData(wb, .x, .y)
  )

  # prepare style for unlocking some area for schools to play with
  # (note that all cells are locked by default and this lock applies when
  # the seet itself is locked/protected)
  unlocked <- createStyle(locked = F)

  # apply sheet protection and unlock specified cells (200 x 100 area)
  walk(sheets(wb), function(x) {
    setColWidths(wb,
      sheet = x,
      cols = seq_len(data_dims[2]),
      widths = 15
    )

    protectWorksheet(
      wb,
      sheet = x,
      password = pswrd,
      lockSelectingLockedCells = FALSE,
      lockSelectingUnlockedCells = FALSE,
      lockFormattingCells = FALSE,
      lockFormattingColumns = FALSE,
      lockFormattingRows = FALSE,
      lockInsertingColumns = FALSE,
      lockInsertingRows = FALSE,
      lockInsertingHyperlinks = FALSE,
      lockDeletingColumns = FALSE,
      lockDeletingRows = FALSE,
      lockSorting = FALSE,
      lockAutoFilter = TRUE,
      lockPivotTables = FALSE,
      lockObjects = FALSE,
      lockScenarios = FALSE
    )

    addStyle(
      wb,
      sheet = x,
      style = unlocked,
      cols = seq(2, data_dims[2]),
      rows = seq(2, data_dims[1] + 1), # keep headers locked, +1 'cause headers
      gridExpand = T,
      stack = TRUE # retain the previous formatting
    )
  })

  # this ensures user cannot simply delete or rename or move any sheets
  protectWorkbook(wb, lockStructure = TRUE, password = pswrd)

  # write the protected files to the separate folder
  saveWorkbook(wb, school_file, overwrite = TRUE)
})
