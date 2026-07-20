#!/usr/bin/env Rscript
# ============================================================
# BAMA Assay Dashboard -- Shiny App
# Tabs: 1) Raw Data Upload  2) Helper Setup
# ============================================================

suppressPackageStartupMessages({
  library(shiny)
  library(shinydashboard)
  library(DT)
  library(openxlsx)
  library(dplyr)
  library(stringr)
  library(shinyjs)
})

# Optional packages -- install if missing then load
.install_if_missing <- function(pkg) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    message("Installing missing package: ", pkg)
    install.packages(pkg, repos = "https://cloud.r-project.org", quiet = TRUE)
  }
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}
.install_if_missing("ggplot2")
.install_if_missing("scales")
.install_if_missing("readxl")
.install_if_missing("jsonlite")
.install_if_missing("ggrepel")
.install_if_missing("tidyr")
# ragg provides high-quality PNG rendering without X11 (needed for headless servers)
tryCatch(.install_if_missing("ragg"), error = function(e) invisible(NULL))
tryCatch(.install_if_missing("DescTools"), error = function(e) invisible(NULL))
tryCatch(.install_if_missing("drc"),       error = function(e) invisible(NULL))

# Null-coalescing operator (must be at top level)
`%||%` <- function(a, b) if (!is.null(a) && length(a) > 0 && !all(is.na(a))) a else b

# -- Helper generation logic ---------------------------------------------------
# Output mirrors the universal helper file (plate_setup sheet) MINUS virus_id
# and virus_control_range.
#
# KEY DESIGN: one row per (analyte x sample_base_id).
#   plate_analyte       = individual antigen column name (e.g. "GT1.1 (26)")
#   sample_id/plate_id  = base sample name (_N suffix stripped, e.g. "VRC01")
#   plate_range         = all wells where that sample appears (e.g. "A1:B2")
#   no_mab_range  = well range of C-type control wells (same for all rows)
#   sample_type         = auto-classified from Type column:
#                           X*/U* -> "sample"
#                           C*    -> "control"
#                           S*    -> "standard_curve"
#                           B     -> "no_mab"
#                           other -> "no_sample"
#
# Output columns:
#   plate_analyte, experiment_id, sample_id, plate_range,
#   no_mab_range, sample_type, start_concentration,
#   dilution_factor, bama_type, instrument, scientist_id
# ---------------------------------------------------------------------------

# Convert well labels vector -> bounding-box range string (e.g. "A1:H2").
# Used for plate_range (samples, standards, no_mab, no_sample): the full
# rectangular bounding box of all wells belonging to that entry.
.wells_to_range <- function(wells) {
  wells <- as.character(wells)
  wells <- wells[!is.na(wells) & nchar(wells) > 0 & wells != "NA"]
  if (length(wells) == 0) return("")
  rows  <- gsub("[0-9]",   "", wells)
  cols  <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", wells)))
  ridx  <- match(toupper(rows), LETTERS)
  min_r <- LETTERS[min(ridx, na.rm = TRUE)]
  max_r <- LETTERS[max(ridx, na.rm = TRUE)]
  min_c <- min(cols, na.rm = TRUE)
  max_c <- max(cols, na.rm = TRUE)
  if (min_r == max_r && min_c == max_c) paste0(min_r, min_c)
  else paste0(min_r, min_c, ":", max_r, max_c)
}

# Convert well labels vector -> per-row continuous range string.
# Used exclusively for no_mab_range: groups wells by row letter and,
# within each row, emits one segment per CONTINUOUS column run.
# Example: A11,B11,C11,D11,E11,F11,A12,B12,C12,D12,E12,F12 -> "A11:F11, A12:F12"
# This correctly excludes non-contiguous wells (e.g. G11, H12 are NOT included).
.wells_to_ctrl_range <- function(wells) {
  wells <- as.character(wells)
  wells <- wells[!is.na(wells) & nchar(wells) > 0 & wells != "NA"]
  if (length(wells) == 0) return("")

  rows <- toupper(gsub("[0-9]", "", wells))
  cols <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", wells)))

  df <- data.frame(row = rows, col = cols, stringsAsFactors = FALSE)
  df <- df[order(match(df$row, LETTERS), df$col), ]
  df <- df[!is.na(df$col), , drop = FALSE]
  if (nrow(df) == 0) return("")

  segments <- character(0)
  i <- 1L
  while (i <= nrow(df)) {
    r       <- df$row[i]
    c_start <- df$col[i]
    c_end   <- c_start
    # Extend run while same row AND consecutive column
    while (i + 1L <= nrow(df) &&
           df$row[i + 1L] == r &&
           df$col[i + 1L] == c_end + 1L) {
      i     <- i + 1L
      c_end <- df$col[i]
    }
    seg <- if (c_start == c_end) paste0(r, c_start)
           else                  paste0(r, c_start, ":", r, c_end)
    segments <- c(segments, seg)
    i <- i + 1L
  }
  paste(segments, collapse = ", ")
}

# Classify instrument Type code -> sample_type label
.classify_type <- function(type_code) {
  t <- as.character(type_code)
  dplyr::case_when(
    grepl("^[XUxu]\\d*$", t) ~ "sample",
    grepl("^[Cc]\\d*$",   t) ~ "control",
    grepl("^[Ss]\\d*$",   t) ~ "standard_curve",
    grepl("^[Bb]$",       t) ~ "no_mab",
    TRUE                      ~ "no_sample"
  )
}

generate_helper <- function(input_path) {
  fname_base  <- tools::file_path_sans_ext(basename(input_path))
  fname_parts <- str_split(fname_base, "_", n = 3)[[1]]
  if (length(fname_parts) < 3)
    stop("Filename must follow the pattern  Date_Study_ScientistID.xlsx")

  raw_date     <- fname_parts[1]
  bama_type    <- fname_parts[2]
  scientist_id <- fname_parts[3]

  experimental_date <- as.Date(raw_date, format = "%Y%m%d")
  if (is.na(experimental_date))
    stop(paste("Cannot parse date from filename:", raw_date))

  wb_raw <- read.xlsx(input_path, sheet = 1, colNames = FALSE)

  header_row <- which(apply(wb_raw, 1,
                      function(r) any(r == "Well", na.rm = TRUE)))[1]
  if (is.na(header_row))
    stop("Cannot locate 'Well' column header in the sheet.")

  # ---- Instrument serial from metadata rows --------------------------------
  instrument_serial <- NA_character_
  if (header_row > 1) {
    for (r in seq_len(header_row - 1)) {
      row_str <- paste(as.character(wb_raw[r, ]), collapse = " ")
      if (grepl("Reader Serial Number", row_str, ignore.case = TRUE)) {
        m <- regmatches(row_str,
               regexpr("Reader Serial Number:\\s*(\\S+)", row_str,
                       ignore.case = TRUE))
        if (length(m) > 0)
          instrument_serial <- str_trim(
            sub("Reader Serial Number:\\s*", "", m, ignore.case = TRUE))
        break
      }
    }
  }

  pd           <- read.xlsx(input_path, sheet = 1,
                             startRow = header_row, colNames = TRUE)
  colnames(pd) <- str_trim(colnames(pd))

  # ---- Capture ORIGINAL column names from the raw header row ---------------
  # read.xlsx with colNames=TRUE mangles names (spaces/parens -> dots).
  # wb_raw[header_row, ] still holds the unmangled originals.
  orig_header <- as.character(wb_raw[header_row, ])
  orig_header <- str_trim(orig_header[!is.na(orig_header) & nchar(str_trim(orig_header)) > 0])
  # Build a map: R-mangled name -> original name
  r_names  <- colnames(pd)
  # Match by position: orig_header and r_names should have the same length
  if (length(orig_header) == length(r_names)) {
    orig_name_map <- setNames(orig_header, r_names)
  } else {
    # Fallback: use R names as-is
    orig_name_map <- setNames(r_names, r_names)
  }

  missing_cols <- setdiff(c("Well", "Type", "Description"), colnames(pd))
  if (length(missing_cols) > 0)
    stop(paste("Missing column(s):", paste(missing_cols, collapse = ", ")))

  pd$well <- as.character(pd$Well)
  pd$type <- as.character(pd$Type)
  pd$desc <- as.character(pd$Description)
  pd$desc[is.na(pd$Description) | pd$desc == "NA"] <- ""

  # ---- Identify analyte columns (position-based: between Description & trailing metadata) ----
  # Strategy: analyte columns always sit between Description and the first
  # trailing metadata column (Region / Gate / Total / Location / Plate ID /
  # Bead Count / Acquisition / Sampling / % Agg / X..Agg / Plate.ID).
  # This avoids fragile pattern matching on R-mangled column names.
  all_cols <- colnames(pd)

  # Index of Description column (use original raw name stored before rename)
  desc_idx <- which(tolower(all_cols) == "description")[1]
  if (is.na(desc_idx)) desc_idx <- 3L   # fallback: assume col 3

  # Trailing metadata sentinel pattern (matches both original and R-dotted names)
  trailing_pat <- paste0("^(Region|Gate|Total|Location|X\\\\.\\\\.Agg|X\\\\.\\\\..Agg|%|",
                         "Sampling|Plate\\\\.ID|Plate.ID|Bead\\\\.Count|",
                         "Bead.Count|Acquisition|X\\\\.\\\\.)")
  trailing_idx <- which(grepl(trailing_pat, all_cols, ignore.case = TRUE))[1]
  if (is.na(trailing_idx)) trailing_idx <- length(all_cols) + 1L

  # Candidate analyte cols: strictly between Description and first trailing col
  candidate_idx <- seq(desc_idx + 1L, trailing_idx - 1L)
  candidate_cols <- all_cols[candidate_idx]

  # Exclude BLANK column(s) -- original name contains "BLANK" or R-dotted "BLANK"
  is_blank_col <- grepl("^BLANK", candidate_cols, ignore.case = TRUE)
  ag_cols      <- candidate_cols[!is_blank_col]

  # Resolve R-mangled names back to original display names for plate_analyte
  ag_display   <- unname(sapply(ag_cols, function(rc) {
    orig <- orig_name_map[rc]
    if (!is.na(orig) && nchar(orig) > 0) orig else rc
  }))

  # Exclude "ghost" analyte columns -- bead regions the acquisition software
  # exported with no reagent/antigen name assigned, which show up in the raw
  # header as blank, "NaN", or "NA". These aren't real analytes and should
  # never become a plate_analyte value in the helper file.
  is_ghost_ag <- is.na(ag_display) | trimws(ag_display) == "" |
                 toupper(trimws(ag_display)) %in% c("NAN", "NA")
  if (any(is_ghost_ag)) {
    ag_cols    <- ag_cols[!is_ghost_ag]
    ag_display <- ag_display[!is_ghost_ag]
  }

  if (length(ag_cols) == 0)
    stop(paste0("No analyte columns detected between 'Description' (col ", desc_idx,
                ") and trailing metadata (col ", trailing_idx, ").",
                " Candidates were: ", paste(candidate_cols, collapse=", ")))

  # ---- Classify all wells --------------------------------------------------
  pd$sample_type <- .classify_type(pd$type)

  # ---- no_mab wells (B): used as no_mab_range for all rows -----------
  # The no_mab (B-type) wells define the no-antibody baseline; their
  # per-row continuous range is stored in no_mab_range (e.g. "A12:F12").
  # no_mab rows are NOT emitted as separate helper rows.
  blank_pd        <- pd[pd$sample_type == "no_mab", , drop = FALSE]
  no_mab_range    <- .wells_to_ctrl_range(blank_pd$well)  # per-row continuous

  # ---- Build one row per (analyte x sample_base_id) ------------------------
  # Strip trailing _N from description to get base sample id
  pd$base_id <- sub("_\\d+$", "", pd$desc)
  pd$base_id[pd$base_id == ""] <- "no_mab"

  exp_counter <- 0L   # reset per analyte inside the loop below
  all_rows    <- list()

  # -- One outer loop per analyte; exp_counter restarts at 1 for each --------
  samp_pd      <- pd[pd$sample_type == "sample",         , drop = FALSE]
  ctrl_pd      <- pd[pd$sample_type == "control",        , drop = FALSE]
  std_pd       <- pd[pd$sample_type == "standard_curve", , drop = FALSE]
  nosamp_pd    <- pd[pd$sample_type == "no_sample" &
                     nchar(trimws(pd$well)) > 0,          , drop = FALSE]

  unique_bases <- unique(samp_pd$base_id)
  unique_ctrls <- unique(ctrl_pd$desc)
  unique_stds  <- unique(std_pd$desc)

  for (ai in seq_along(ag_cols)) {
    ag_disp     <- ag_display[ai]
    exp_counter <- 0L   # reset for each analyte

    # SAMPLE rows
    for (bid in unique_bases) {
      exp_counter <- exp_counter + 1L
      grp <- samp_pd[samp_pd$base_id == bid, , drop = FALSE]
      all_rows[[length(all_rows) + 1L]] <- data.frame(
        plate_analyte       = ag_disp,
        experiment_id       = sprintf("EXP%03d", exp_counter),
        sample_id           = bid,
        plate_range         = .wells_to_range(grp$well),
        no_mab_range  = no_mab_range,
        sample_type         = "sample",
        start_concentration = "",
        dilution_factor     = "",
        bama_type           = bama_type,
        instrument          = instrument_serial %||% "",
        scientist_id        = scientist_id,
        samples             = "",
        stringsAsFactors    = FALSE
      )
    }

    # CONTROL rows
    for (cid in unique_ctrls) {
      exp_counter <- exp_counter + 1L
      grp <- ctrl_pd[ctrl_pd$desc == cid, , drop = FALSE]
      all_rows[[length(all_rows) + 1L]] <- data.frame(
        plate_analyte       = ag_disp,
        experiment_id       = sprintf("EXP%03d", exp_counter),
        sample_id           = cid,
        plate_range         = .wells_to_range(grp$well),
        no_mab_range  = no_mab_range,
        sample_type         = "control",
        start_concentration = "",
        dilution_factor     = NA_character_,
        bama_type           = bama_type,
        instrument          = instrument_serial %||% "",
        scientist_id        = scientist_id,
        samples             = "",
        stringsAsFactors    = FALSE
      )
    }

    # STANDARD CURVE rows
    for (sid in unique_stds) {
      exp_counter <- exp_counter + 1L
      grp <- std_pd[std_pd$desc == sid, , drop = FALSE]
      all_rows[[length(all_rows) + 1L]] <- data.frame(
        plate_analyte       = ag_disp,
        experiment_id       = sprintf("EXP%03d", exp_counter),
        sample_id           = sid,
        plate_range         = .wells_to_range(grp$well),
        no_mab_range  = no_mab_range,
        sample_type         = "standard_curve",
        start_concentration = "",
        dilution_factor     = "",
        bama_type           = bama_type,
        instrument          = instrument_serial %||% "",
        scientist_id        = scientist_id,
        samples             = "",
        stringsAsFactors    = FALSE
      )
    }


    # NO_SAMPLE rows
    for (k in seq_len(nrow(nosamp_pd))) {
      exp_counter <- exp_counter + 1L
      all_rows[[length(all_rows) + 1L]] <- data.frame(
        plate_analyte       = ag_disp,
        experiment_id       = sprintf("EXP%03d", exp_counter),
        sample_id           = nosamp_pd$desc[k],
        plate_range         = nosamp_pd$well[k],
        no_mab_range  = no_mab_range,
        sample_type         = "no_sample",
        start_concentration = "",
        dilution_factor     = "",
        bama_type           = bama_type,
        instrument          = instrument_serial %||% "",
        scientist_id        = scientist_id,
        samples             = "",
        stringsAsFactors    = FALSE
      )
    }
  } # end analyte loop

  if (length(all_rows) == 0)
    stop("No wells could be classified. Check Well/Type/Description columns.")

  helper <- do.call(rbind, all_rows)

  list(
    helper            = helper,
    experimental_date = format(experimental_date, "%Y-%m-%d"),
    bama_type         = bama_type,
    scientist_id      = scientist_id,
    instrument_serial = instrument_serial,
    filename          = fname_base,
    plate_analyte     = paste(ag_display, collapse = "; ")
  )
}




# -- UI -----------------------------------------------------------------------
ui <- dashboardPage(
  skin = "blue",

  # Logo note: Shiny serves static files from a www/ folder placed next to this app file.
  # To display the AIRU logo, create: <app_directory>/www/airu_logo.png
  # The onerror handler hides the img gracefully if the file is missing.
  
  dashboardHeader(
    title = tags$span(
      tags$img(src = "airu_logo.png",
               height = "38px",
               style = "margin-right:8px; vertical-align:middle;",
               onerror = "this.style.display='none'"),
      "BAMA Dashboard"
    ),
    titleWidth = 340
  ),

  dashboardSidebar(
    width = 270,
    sidebarMenu(
      id = "sidebar",
      menuItem("Overview",          tabName = "overview",           icon = icon("home")),
      menuItem("Plate Data Upload", tabName = "upload",             icon = icon("upload")),
      menuItem("Helper Setup",      tabName = "helper",             icon = icon("cog")),
      menuItem("Plate Review",      tabName = "review",             icon = icon("table")),
      menuItem("Processed",         tabName = "dataframe",          icon = icon("file-excel")),
      menuItem("Analysis",          tabName = "analysis",           icon = icon("chart-line"),
        tags$div(id = "sidebar_item_point",
          menuSubItem("Point-based", tabName = "analysis_point")
        ),
        tags$div(id = "sidebar_item_titration",
          menuSubItem("Titration", tabName = "analysis_titration")
        ),
        tags$div(id = "sidebar_item_quant",
          menuSubItem("Quantification",tabName = "analysis_quant")
        )
      ),
      menuItem("Export",            tabName = "export",             icon = icon("download"))
    ),
    tags$hr(),
    tags$div(
      style = "padding: 10px 15px;",
      fileInput("sidebar_helper_file", "Upload Complete Helper File (.xlsx)",
                accept = c(".xlsx", ".xls"),
                buttonLabel = "Browse...",
                placeholder = "universal_helper_..."),
      tags$strong(style = "color:#aaa; font-size:11px; text-transform:uppercase; letter-spacing:1px;", "RUN PARAMETERS"),
      tags$br(), tags$br(),
      tags$label(style = "color:#c8d8e8; font-size:12px;", "Analysis Date (YYYYMMDD)"),
      textInput("sidebar_analysis_date", NULL, value = format(Sys.Date(), "%Y%m%d")),
      tags$label(style = "color:#c8d8e8; font-size:12px;", "Analyst Name"),
      textInput("sidebar_analyst_name", NULL, placeholder = "First Last"),
      tags$br(),
      actionButton("sidebar_run_analysis", "Run Analysis",
                   class = "btn btn-primary btn-block",
                   style = "background:#1e88e5; border-color:#1e88e5; width:100%; font-weight:600;",
                   icon  = icon("play"))
    )
  ),

  dashboardBody(
    useShinyjs(),
    tags$head(
      tags$style(HTML("
        /* -- Global -------------------------------- */
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .content-wrapper { background: #f4f6f9; }

        /* -- Header -------------------------------- */
        .main-header .logo { background: #1a3a5c !important; font-weight: 700; }
        .main-header .navbar { background: #1e4976 !important; }

        /* -- Sidebar ------------------------------- */
        .main-sidebar { background: #1a2940 !important; }
        .sidebar-menu > li > a { color: #c8d8e8 !important; }
        .sidebar-menu > li.active > a,
        .sidebar-menu > li > a:hover { background: #0d6efd22 !important; color: #fff !important; border-left: 3px solid #4fc3f7; }

        /* -- Section boxes ------------------------- */
        .box { border-top-color: #1e88e5; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
        .box-title { font-weight: 600; font-size: 14px; }

        /* -- Status badges ------------------------- */
        .status-pill { display:inline-block; padding:3px 12px; border-radius:20px; font-size:12px; font-weight:600; }
        .pill-success { background:#d4edda; color:#155724; }
        .pill-waiting { background:#fff3cd; color:#856404; }
        .pill-error   { background:#f8d7da; color:#721c24; }

        /* -- Upload area --------------------------- */
        .upload-zone { border: 2px dashed #4fc3f7; border-radius:10px; padding:30px; text-align:center;
                       background:#eaf6ff; transition:all .2s; }
        .upload-zone:hover { background:#d0ecff; }
        .upload-icon { font-size:40px; color:#1e88e5; margin-bottom:10px; }

        /* -- Data table polish --------------------- */
        .dataTables_wrapper .dataTables_filter input { border-radius:20px; padding:4px 12px; border:1px solid #ccc; }
        .dataTables_wrapper .dataTables_length select { border-radius:4px; }
        table.dataTable thead th { background:#1e4976; color:#fff; }
        table.dataTable tbody tr:hover { background:#e8f4fd !important; }

        /* -- Helper form --------------------------- */
        .helper-meta-card { background:#fff; border-radius:8px; padding:18px 22px; margin-bottom:16px;
                            box-shadow:0 1px 6px rgba(0,0,0,.08); border-left:4px solid #1e88e5; }
        .helper-meta-card h4 { margin-top:0; color:#1e4976; font-size:15px; }
        .form-control { border-radius:5px; }
        .btn-run { background:#1e88e5; color:#fff; border:none; border-radius:6px;
                   padding:10px 28px; font-size:14px; font-weight:600; cursor:pointer; }
        .btn-run:hover { background:#1565c0; }
        .btn-dl  { background:#27ae60; color:#fff; border:none; border-radius:6px;
                   padding:8px 20px; font-size:13px; font-weight:600; cursor:pointer; }
        .btn-dl:hover  { background:#1e8449; }

        /* -- Editable-row highlight ---------------- */
        .edited-cell { background:#fffde7 !important; }

        /* -- Info banner --------------------------- */
        .info-banner { background:#e3f2fd; border-left:4px solid #1e88e5; border-radius:4px;
                       padding:10px 16px; font-size:13px; color:#0d3c61; margin-bottom:14px; }

        /* -- Instrument selector cards ------------- */
        .instrument-btn-group { display:flex; gap:12px; margin-top:8px; }
        .instrument-card {
          flex:1; border:2px solid #ccc; border-radius:10px;
          padding:16px 12px; text-align:center; cursor:pointer;
          background:#fff; transition:all .18s;
        }
        .instrument-card:hover { border-color:#1e88e5; background:#eaf6ff; }
        .instrument-card.selected-bioplex    { border-color:#1e88e5; background:#e3f2fd; }
        .instrument-card.selected-intelliflex { border-color:#6a1b9a; background:#f3e5f5; }
        .instrument-card .inst-icon  { font-size:26px; margin-bottom:6px; }
        .instrument-card .inst-label { font-weight:700; font-size:13px; }
        .instrument-card .inst-desc  { font-size:11px; color:#666; margin-top:3px; }
        .bioplex-badge {
          display:inline-block; background:#1e88e5; color:#fff;
          border-radius:12px; padding:2px 10px; font-size:11px; font-weight:700; margin-left:6px;
        }
        .intelliflex-required-badge {
          display:inline-block; background:#6a1b9a; color:#fff;
          border-radius:12px; padding:2px 10px; font-size:11px; font-weight:700; margin-left:6px;
        }

        /* -- Helper tab locked banner (INTELLIFLEX) - */
        .helper-locked-banner {
          background:#f3e5f5; border-left:5px solid #6a1b9a; border-radius:6px;
          padding:28px 32px; margin:30px auto; max-width:600px; text-align:center;
          box-shadow:0 2px 10px rgba(0,0,0,.08);
        }
        .helper-locked-banner .lock-icon { font-size:38px; color:#6a1b9a; margin-bottom:12px; }
        .helper-locked-banner h3 { color:#4a0072; font-size:18px; margin:0 0 8px 0; font-weight:700; }
        .helper-locked-banner p  { color:#4a0072; font-size:13.5px; margin:0; line-height:1.6; }

        /* -- CV% table cell highlights ------------- */
        .cv-warn  { background:#fff3cd !important; color:#856404; font-weight:700; }
        .cv-alert { background:#f8d7da !important; color:#721c24; font-weight:700; }
        .cv-ok    { background:#d4edda !important; color:#155724; }

        /* -- CV threshold alert box ---------------- */
        .cv-threshold-alert {
          background:#f39c12; color:#fff; border-radius:6px;
          padding:10px 16px; font-size:13px; font-weight:700;
          margin-top:12px; text-align:center;
        }
        .cv-threshold-ok {
          background:#27ae60; color:#fff; border-radius:6px;
          padding:10px 16px; font-size:13px; font-weight:700;
          margin-top:12px; text-align:center;
        }

        /* -- Avg-duplicates section header --------- */
        .avg-section-header {
          background: #eaf6ff;
          border-left: 4px solid #1e88e5;
          border-radius: 4px;
          padding: 8px 14px;
          font-size: 14px;
          font-weight: 700;
          color: #1a3a5c;
          margin: 18px 0 10px 0;
        }

        /* -- Plate border override ----------------- */
        .plate-outer-border { outline: 2px solid #000; }

        /* -- LOD/LOQ lock button ------------------- */
        .lock-btn {
          background: none; border: 1.5px solid #ccc; border-radius: 5px;
          padding: 4px 8px; cursor: pointer; font-size: 13px;
          transition: all .18s; margin-top: 2px;
        }
        .lock-btn.locked {
          background: #fff3cd; border-color: #f39c12; color: #856404;
        }
        .lock-btn.unlocked {
          background: #f8f9fa; border-color: #aaa; color: #555;
        }
        .lock-btn:hover { opacity: 0.8; }
        .lod-loq-wrap { display: flex; align-items: center; gap: 6px; }
        .lod-loq-wrap .form-group { margin-bottom: 0; }
        input[readonly] {
          background: #f5f5f5 !important;
          border-color: #ccc !important;
          color: #999 !important;
          cursor: not-allowed;
        }

        .no-run-banner {
          background: #fff8e1;
          border-left: 5px solid #f39c12;
          border-radius: 6px;
          padding: 28px 32px;
          margin: 30px auto;
          max-width: 640px;
          text-align: center;
          box-shadow: 0 2px 10px rgba(0,0,0,.08);
        }
        .no-run-banner .no-run-icon { font-size: 38px; color: #f39c12; margin-bottom: 12px; }
        .no-run-banner h3 { color: #7d5a00; font-size: 18px; margin: 0 0 8px 0; font-weight: 700; }
        .no-run-banner p  { color: #7d5a00; font-size: 13.5px; margin: 0; line-height: 1.6; }
        .no-run-banner .run-hint {
          display: inline-block; margin-top: 16px;
          background: #f39c12; color: #fff;
          border-radius: 20px; padding: 6px 18px;
          font-size: 12.5px; font-weight: 600;
        }

        /* -- Per-tab refresh button ------------------- */
        .tab-refresh-bar {
          display: flex;
          justify-content: flex-end;
          align-items: center;
          margin-bottom: 10px;
        }
        .tab-refresh-btn {
          background: #1e88e5;
          color: #fff;
          border: none;
          border-radius: 6px;
          padding: 5px 14px;
          font-size: 12px;
          font-weight: 600;
          cursor: pointer;
          transition: background .15s;
        }
        .tab-refresh-btn:hover { background: #1565c0; color: #fff; }

      ")),
    ),

    tabItems(

      # ==========================================================
      # TAB 1 -- Raw Data Upload
      # ==========================================================
      tabItem(
        tabName = "upload",

        tags$div(class = "tab-refresh-bar",
          actionButton("refresh_upload", label = tagList(icon("sync-alt"), " Refresh"),
            class = "tab-refresh-btn btn")
        ),
        # -- Step 1: Instrument selector -----------------------------------
        fluidRow(
          box(
            title = tags$span(icon("microscope"), " Step 1: Select Instrument"),
            status = "primary", solidHeader = TRUE, width = 12,
            tags$div(
              class = "instrument-btn-group",
              # BioPlex card
              tags$div(
                id    = "inst_card_bioplex",
                class = "instrument-card selected-bioplex",
                onclick = "Shiny.setInputValue('instrument_select','BIOPLEX',{priority:'event'})",
                tags$div(class = "inst-icon", icon("server")),
                tags$div(class = "inst-label",
                  "BioPlex",
                  tags$span(class = "bioplex-badge", "Active")
                ),
                tags$div(class = "inst-desc",
                  "Standard BAMA format -- 'Well' header scan.",
                  tags$br(), "Helper Setup tab available.")
              ),
              # INTELLIFLEX card
              tags$div(
                id    = "inst_card_intelliflex",
                class = "instrument-card",
                onclick = "Shiny.setInputValue('instrument_select','INTELLIFLEX',{priority:'event'})",
                tags$div(class = "inst-icon", icon("microchip")),
                tags$div(class = "inst-label",
                  "INTELLIFLEX",
                  tags$span(class = "intelliflex-required-badge", "Requires plate file")
                ),
                tags$div(class = "inst-desc",
                  "Reads 'Results' sheet -- auto-detects ANALYTE NAME / MEDIAN columns across all regions.",
                  tags$br(), "96-well plate layout is ", tags$strong("required."))
              )
            ),
            tags$script(HTML("
              Shiny.addCustomMessageHandler('updateInstrumentCards', function(val) {
                var bp = document.getElementById('inst_card_bioplex');
                var ix = document.getElementById('inst_card_intelliflex');
                if (val === 'BIOPLEX') {
                  bp.className = 'instrument-card selected-bioplex';
                  ix.className = 'instrument-card';
                } else {
                  bp.className = 'instrument-card';
                  ix.className = 'instrument-card selected-intelliflex';
                }
              });
            "))
          )
        ),

        # -- Step 2: Raw file upload ---------------------------------------
        fluidRow(
          box(
            title = uiOutput("upload_box_title"),
            status = "primary", solidHeader = TRUE, width = 12,
            fluidRow(
              column(6,
                tags$div(class = "upload-zone",
                  tags$div(class = "upload-icon", icon("file-excel")),
                  fileInput("raw_file", NULL,
                    accept      = c(".xlsx", ".xls"),
                    buttonLabel = "Browse Files...",
                    placeholder = "Drop .xlsx file here or click Browse"
                  ),
                  uiOutput("upload_filename_hint")
                )
              ),
              column(6,
                tags$div(class = "helper-meta-card",
                  tags$h4(icon("info-circle"), " File Parsing Status"),
                  uiOutput("upload_status"),
                  tags$hr(),
                  uiOutput("file_meta_display")
                )
              )
            )
          )
        ),

        # -- Step 3: 96-well plate layout (required/optional per instrument) --
        fluidRow(uiOutput("well96_box_ui")),

        fluidRow(
          box(
            title = "Raw Plate Data Preview",
            status = "primary", solidHeader = TRUE, width = 12,
            collapsible = TRUE,
            uiOutput("raw_data_note"),
            DTOutput("raw_data_table")
          )
        )
      ),

      # ==========================================================
      # TAB 2 -- Helper Setup  (BioPlex only -- locked for INTELLIFLEX)
      # ==========================================================
      tabItem(
        tabName = "helper",
        tags$div(class = "tab-refresh-bar",
          actionButton("refresh_helper", label = tagList(icon("sync-alt"), " Refresh"),
            class = "tab-refresh-btn btn")
        ),

        uiOutput("helper_prereq_banner"),

        fluidRow(
            box(
              title = "Run Parameters",
              status = "primary", solidHeader = TRUE, width = 4,
              tags$div(class = "helper-meta-card",
                tags$h4(icon("calendar"), " Experimental Metadata"),
                textInput("run_date",       "Run Date (YYYYMMDD)",      value = format(Sys.Date(), "%Y%m%d")),
                tags$div(style = "display:none;",
                  checkboxGroupInput("bama_type_select", NULL,
                    choices  = c("Point-based", "Titration", "Quantification"),
                    selected = "Point-based",
                    inline   = FALSE)),
                textInput("scientist_id",   "Scientist ID",             value = ""),
                textInput("instrument_sn",  "Instrument Serial Number", value = ""),
                tags$hr(),
                tags$div(
                  style = "text-align:center;",
                  actionButton("btn_generate", "Generate / Refresh Helper",
                    class = "btn-run", icon = icon("sync"))
                )
              )
            ),
            column(8,
              box(
                title = "Helper File Preview  (editable -- double-click a cell to edit)",
                status = "primary", solidHeader = TRUE, width = NULL,
                collapsible = TRUE,
                tags$div(class = "info-banner",
                  icon("pencil-alt"), " You can edit any cell below before downloading.",
                  "  Editable fields: ",
                  tags$strong("plate_analyte, sample_id, plate_range, no_mab_range, sample_type, start_concentration, dilution_factor, samples.")
                ),
                DTOutput("helper_table"),
                tags$br(),
                fluidRow(
                  column(4,
                    downloadButton("dl_helper_xlsx", "Download Helper (.xlsx)", class = "btn-dl")
                  ),
                  column(4,
                    downloadButton("dl_helper_csv",  "Download Helper (.csv)",  class = "btn-dl",
                      style = "background:#8e44ad !important; border-color:#8e44ad !important;")
                  ),
                  column(4, uiOutput("helper_row_count"))
                ),
                tags$hr(),
                tags$div(
                  class = "info-banner", style = "margin-top:8px;",
                  icon("file-alt"),
                  "  Need a blank template? Download an empty file with column headers only:"
                ),
                fluidRow(
                  column(4,
                    downloadButton("dl_empty_xlsx", "Empty Template (.xlsx)", class = "btn-dl",
                      style = "background:#e67e22 !important; border-color:#e67e22 !important;")
                  ),
                  column(4,
                    downloadButton("dl_empty_csv",  "Empty Template (.csv)",  class = "btn-dl",
                      style = "background:#c0392b !important; border-color:#c0392b !important;")
                  )
                )
              ),
              box(
                title = "Helper File Format Guide",
                status = "info", solidHeader = FALSE, width = NULL,
                collapsible = TRUE, collapsed = TRUE,
                tags$p(tags$strong("Column reference:")),
                tags$table(
                  class = "table table-sm table-bordered",
                  style = "font-size:12px;",
                  tags$thead(tags$tr(
                    tags$th("Column"), tags$th("Description"), tags$th("Auto-filled?")
                  )),
                  tags$tbody(
                    tags$tr(tags$td("plate_analyte"),       tags$td("Analyte/antigen name(s) from raw instrument file"), tags$td("\u2705 auto")),
                    tags$tr(tags$td("experiment_id"),       tags$td("Sequential EXP001, EXP002\u2026"),                tags$td("\u2705 auto")),
                    tags$tr(tags$td("sample_id"),           tags$td("Base sample name (_N suffix stripped)"),            tags$td("\u2705 auto")),
                    tags$tr(tags$td("plate_range"),         tags$td("Wells occupied by this sample (e.g. A1:H2)"),       tags$td("\u2705 auto")),
                    tags$tr(tags$td("no_mab_range"),  tags$td("Well range of C-type control wells"),               tags$td("\u2705 auto")),
                    tags$tr(tags$td("sample_type"),         tags$td("X*\u2192sample  C*\u2192control  S*\u2192standard_curve  B\u2192no_mab  other\u2192no_sample"), tags$td("\u2705 auto")),
                    tags$tr(tags$td("start_concentration"), tags$td("Starting concentration (\u03bcg/mL) \u2014 fill in"), tags$td("\u274c fill in")),
                    tags$tr(tags$td("dilution_factor"),     tags$td("Fold dilution between steps \u2014 fill in"),      tags$td("\u274c fill in")),
                    tags$tr(tags$td("bama_type"),           tags$td("Point-based / Titration / Quantification"),         tags$td("\u2705 auto")),
                    tags$tr(tags$td("instrument"),          tags$td("Instrument serial (parsed from file header)"),      tags$td("\u2705 auto")),
                    tags$tr(tags$td("scientist_id"),        tags$td("From filename suffix"),                             tags$td("\u2705 auto")),
                    tags$tr(tags$td("samples"),            tags$td("Sample type: serum or mab \u2014 fill in"),         tags$td("\u274c fill in"))
                  )
                )
              )
            )
          )
      ),

      # ==========================================================
      # TAB -- Overview
      # ==========================================================
      tabItem(tabName = "overview",
        tags$div(class = "tab-refresh-bar",
          actionButton("refresh_overview", label = tagList(icon("sync-alt"), " Refresh"),
            class = "tab-refresh-btn btn")
        ),

        # -- Stat boxes row -------------------------------------
        fluidRow(
          valueBox(
            value    = uiOutput("ov_plates"),
            subtitle = "Plates Processed",
            icon     = icon("flask"),
            color    = "blue",
            width    = 3
          ),
          valueBox(
            value    = uiOutput("ov_samples"),
            subtitle = "Unique Samples",
            icon     = icon("vials"),
            color    = "green",
            width    = 3
          ),
          valueBox(
            value    = uiOutput("ov_viruses"),
            subtitle = "Antigen(s) Tested",
            icon     = icon("virus"),
            color    = "red",
            width    = 3
          )
        ),

        # -- BAMA Type selector ----------------------------------------
        fluidRow(
          box(
            title = tags$span(icon("tag"), " BAMA Type -- Assay Selection"),
            status = "primary", solidHeader = TRUE, width = 12,
            tags$div(
              style = "display:flex; align-items:flex-start; gap:24px; flex-wrap:wrap;",
              tags$div(
                style = "flex:1; min-width:280px;",
                tags$p(style = "font-size:13px; color:#555; margin-bottom:10px;",
                  "Select the assay type(s) for this run. This controls which analysis
                   tabs are active and is recorded in the helper file.
                   \u2018Point-based\u2019 is always included; type to add a custom assay type."
                ),
                selectizeInput(
                  "bama_type_select_ov", NULL,
                  choices  = c("Point-based", "Titration", "Quantification"),
                  selected = "Point-based",
                  multiple = TRUE,
                  options  = list(
                    placeholder = "Type or select assay type(s)...",
                    plugins     = list("remove_button"),
                    create      = TRUE,
                    createOnBlur = TRUE
                  ),
                  width = "100%"
                ),
                uiOutput("bama_type_pills_display")
              )
            )
          )
        ),

        # -- Workflow Guide + About This Dashboard --------------
        fluidRow(
          box(
            title       = tags$span(style = "color:#1a3a5c; font-weight:700;", "Workflow Guide"),
            status      = "primary", solidHeader = FALSE,
            width       = 6,
            style       = "border-top: 3px solid #1e88e5;",
            tags$ol(
              style = "font-size:13.5px; line-height:2;",
              tags$li(HTML("Upload your <strong>Raw (.xlsx) File</strong> in the Plate Data Upload tab and Review Raw Plate Data")),
              tags$li(HTML("Review the <strong>Helper Setup</strong> tab to see your current pre-constructed Helper File and edit it completely")),
              tags$li(HTML("Upload the <strong> Complete Helper (.xlsx) Files</strong> in the Upload Complete Helper File (.xlsx). [<strong>Option 2</strong>]: is to download an Empty Helper File, Fill it in and Upload it on the Upload Complete Helper File sidebar")),
              tags$li(HTML("Click <strong>Run Analysis</strong>. For the Intelliflex instrument, before running the analysis, you will be requested to additionally upload the <strong>96-well plate format</strong> and <strong>select the controls from that plate</strong>")),
              tags$li(HTML("Explore Analysis tabs: <strong>Point-based</strong> , <strong>Titration</strong> , <strong>Quantification</strong>")),
              tags$li(HTML("Download XLSX from the <strong>Export</strong> tab"))
            )
          ),
          box(
            title       = tags$span(style = "color:#fff; font-weight:700;", "About This Dashboard"),
            status      = "info", solidHeader = TRUE,
            width       = 6,
            background  = "light-blue",
            tags$ul(
              style = "font-size:13px; line-height:1.9; padding-left:18px;",
              tags$li("Binding Antibody Multiplex Assay (BAMA) plate layout and antigen/bead-region mapping"),
              tags$li("Bead acquisition, negative/positive control, and % aggregation QC checks"),
              tags$li("Point-based, Titration, and Quantification analysis workflows"),
              tags$li("Titration curves (average MFI per analyte vs. dilution/concentration)"),
              tags$li("Quantification of sample concentration via standard curve back-calculation")
            )
          )
        ),

        # -- Recent Run Summary ---------------------------------
        fluidRow(
          box(
            title  = tags$span(style = "color:#fff; font-weight:700;", "Recent Run Summary"),
            status = "warning", solidHeader = TRUE,
            width  = 12,
            uiOutput("ov_run_summary")
          )
        )
      ),

      # ==========================================================
      # Placeholder tabs
      # ==========================================================
      tabItem(tabName = "review",

        tags$div(class = "tab-refresh-bar",
          actionButton("refresh_review", label = tagList(icon("sync-alt"), " Refresh"),
            class = "tab-refresh-btn btn")
        ),
        uiOutput("review_no_run_banner"),

        conditionalPanel(
          condition = "output.analysis_has_run",

          fluidRow(
            box(
              title = "Select Plate", status = "primary", solidHeader = TRUE, width = 12,
              uiOutput("plate_selector_ui")
            )
          ),

          fluidRow(
            box(
              title = uiOutput("plate_heatmap_title"),
              status = "primary", solidHeader = TRUE, width = 6,
              uiOutput("plate_rlu_heatmap_ui")
            ),
            box(
              title = uiOutput("plate_map_title"),
              status = "primary", solidHeader = TRUE, width = 6,
              uiOutput("plate_map_heatmap_ui")
            )
          ),

          fluidRow(
            box(
              title = tags$span(style="font-weight:700;",
                                "Sample Averages"),
              status = "primary", solidHeader = TRUE, width = 8,
              DTOutput("control_summary_table")
            ),
            box(
              title = tags$span(style="font-weight:700;", "CV% Thresholds"),
              status = "warning", solidHeader = TRUE, width = 4,
              tags$label(style="font-size:13px; font-weight:600;",
                         "CV% Warning Threshold:"),
              numericInput("cv_threshold", NULL,
                           value = 20, min = 1, max = 100, step = 1,
                           width = "100px"),
              uiOutput("cv_threshold_alert")
            )
          ),


        )
      ),

      tabItem(tabName = "export",
        tags$div(class = "tab-refresh-bar",
          actionButton("refresh_export", label = tagList(icon("sync-alt"), " Refresh"),
            class = "tab-refresh-btn btn")
        ),
        uiOutput("export_no_run_banner"),
        conditionalPanel(
          condition = "output.analysis_has_run",
          fluidRow(
            # -- Left column: download card + run metadata ------------------
            column(width = 7,
              box(
                title = tags$span(icon("download"), " Export Results"),
                status = "primary", solidHeader = TRUE, width = NULL,

                # Sheet list description (dynamic)
                uiOutput("export_sheet_list_ui"),

                # Download button
                downloadButton("dl_results_xlsx", label = tagList(icon("file-excel"), " Download Results.xlsx"),
                  class = "btn btn-success",
                  style = "font-size:15px; padding:10px 22px; margin-bottom:18px;"
                ),

                tags$hr(),

                # Run metadata verbatim block
                tags$h5(tags$strong("Run Metadata"), style = "margin-bottom:8px;"),
                uiOutput("export_run_metadata_ui")
              )
            ),

            # -- Right column: fields legend --------------------------------
            column(width = 5,
              box(
                title = uiOutput("export_legend_title_ui"),
                status = "info", solidHeader = FALSE, width = NULL,
                style = "border-top: 3px solid #5bc0de;",

                # Sheet selector (dynamic -- hides conditional sheets when not selected)
                uiOutput("export_legend_sheet_selector_ui"),

                tags$div(
                  style = "max-height:420px; overflow-y:auto;",
                  tableOutput("export_legend_table")
                )
              )
            )
          )
        )
      ),

      tabItem(tabName = "dataframe",
        tags$div(class = "tab-refresh-bar",
          actionButton("refresh_dataframe", label = tagList(icon("sync-alt"), " Refresh"),
            class = "tab-refresh-btn btn")
        ),
        uiOutput("dataframe_no_run_banner"),
        conditionalPanel(
          condition = "output.analysis_has_run",
          tabBox(
            title = NULL, id = "processed_subtabs", width = 12,

            # -- Tab 1: mAbs subtraction --------------------------------------
            tabPanel(
              title = tagList(icon("vial"), " no_mAbs subtraction"),
              tags$div(
                class = "info-banner",
                icon("info-circle"),
                HTML(" Raw MFI values with only the <strong>average no_mab subtraction</strong>
                      applied (blank-bead subtraction <em>not</em> included).")
              ),
              DTOutput("df_mabs_table"),
              tags$br(), tags$hr(),
              tags$div(
                style = "font-weight:700; font-size:14px; color:#1a3a5c; margin-bottom:8px;",
                icon("calculator"), "  Average for Duplicates"
              ),
              DTOutput("df_mabs_avg_table")
            ),

            # -- Tab 2: blank beads subtraction -------------------------------
            tabPanel(
              title = tagList(icon("circle"), " blank beads subtraction"),
              tags$div(
                class = "info-banner",
                icon("info-circle"),
                HTML(" Raw MFI values with only the <strong>blank bead (R44 / BLANK 44)
                      subtraction</strong> applied, negative values floored to zero
                      (no_mab subtraction <em>not</em> included).")
              ),
              DTOutput("df_blank_table"),
              tags$br(), tags$hr(),
              tags$div(
                style = "font-weight:700; font-size:14px; color:#1a3a5c; margin-bottom:8px;",
                icon("calculator"), "  Average for Duplicates"
              ),
              DTOutput("df_blank_avg_table")
            ),

            # -- Tab 3: background subtraction (full pipeline) -----------------
            tabPanel(
              title = tagList(icon("table"), " background subtraction"),
              tags$div(
                class = "info-banner",
                icon("info-circle"),
                HTML(" Columns: <strong>Well</strong>, <strong>Type</strong>,
                      <strong>Sample_ID</strong>, and MFI per antigen after
                      <strong>both blank-bead and no_mab subtraction</strong>.
                      For INTELLIFLEX, sample names are sourced from the 96-well layout file.
                      For BioPlex, names come from the raw file Description / helper.")
              ),
              DTOutput("dataframe_mfi_table"),
              tags$br(), tags$hr(),
              tags$div(
                style = "font-weight:700; font-size:14px; color:#1a3a5c; margin-bottom:8px;",
                icon("calculator"), "  Average MFI for Duplicates"
              ),
              tags$div(
                class = "info-banner",
                icon("info-circle"),
                HTML(" When a <strong>Sample_ID</strong> appears more than once (duplicates),
                      this table shows the mean MFI per antigen across those replicate wells.")
              ),
              DTOutput("dataframe_avg_table")
            )
          )
        )
      ),

      # ==========================================================
      # TAB -- Analysis (sub-tabs)
      # ==========================================================
      tabItem(tabName = "analysis_point",
        tags$div(class = "tab-refresh-bar",
          actionButton("refresh_point", label = tagList(icon("sync-alt"), " Refresh"),
            class = "tab-refresh-btn btn")
        ),
        uiOutput("analysis_point_no_run_banner"),
        uiOutput("analysis_point_bama_gate"),
        conditionalPanel(
          condition = "output.analysis_has_run && output.analysis_point_selected",

          # ---- QC Configuration (available for BioPlex and INTELLIFLEX) -------
          conditionalPanel(
            condition = "output.pb_qc_available",

            fluidRow(
              box(
                title = tags$span(icon("sliders-h"), " Point-based QC Configuration"),
                status = "primary", solidHeader = TRUE, width = 12,

                # -- BioPlex only: is this a BRILLIANT assay? ---------------------
                # BRILLIANT assays use established LOD/LOQ thresholds (editable
                # below); non-BRILLIANT BioPlex runs have no validated LOD/LOQ,
                # so both are forced to 0.
                conditionalPanel(
                  condition = "output.is_bioplex",
                  tags$div(
                    style = "display:flex; align-items:center; gap:14px; flex-wrap:wrap;
                             background:#eef3f9; border:1px solid #cdd9e6; border-radius:6px;
                             padding:8px 14px; margin-bottom:12px;",
                    tags$label(style="font-size:13px; font-weight:600; color:#1a3a5c; margin:0;",
                      "Does the analyst have LOD/LOQ values?"),
                    radioButtons("pb_brilliant_toggle", NULL,
                                 choices  = c("Yes" = "yes", "No" = "no"),
                                 selected = "yes", inline = TRUE, width = "auto"),
                    tags$span(style="font-size:11px; color:#555;",
                      "If no, LOD and LOQ are fixed at 0.")
                  )
                ),

                # -- Step 1/2: choose controls first, then apply them to every ----
                # plate/antigen that shares them (e.g. 10 of 15 antigens use the
                # same positive/negative controls) --------------------------------
                tags$div(
                  style = "background:#f7f9fc; border:1px solid #d8e1ec; border-radius:6px;
                           padding:10px 14px; margin-bottom:14px;",
                  tags$h5(icon("layer-group"), " Apply Controls to Multiple Plates / Antigens",
                          style = "margin-top:0; color:#1a3a5c; font-weight:700; font-size:14px;"),
                  tags$p(style = "font-size:12px; color:#555; margin-bottom:10px;",
                    "If several plates/antigens share the same positive and negative controls, ",
                    tags$strong("select the controls once (Step 1)"), ", then ",
                    tags$strong("choose which plates/antigens use them (Step 2)"), "."),
                  fluidRow(
                    column(3,
                      tags$label(style="font-size:12px; font-weight:600; color:#1a3a5c;",
                        "STEP 1a \u2014 ", icon("plus-circle", style="color:#155724;"), " Positive Controls"),
                      uiOutput("pb_bulk_pos_ctrl_ui")
                    ),
                    column(3,
                      tags$label(style="font-size:12px; font-weight:600; color:#1a3a5c;",
                        "STEP 1b \u2014 ", icon("minus-circle", style="color:#721c24;"), " Negative Controls"),
                      uiOutput("pb_bulk_neg_ctrl_ui")
                    ),
                    column(4,
                      tags$label(style="font-size:12px; font-weight:600; color:#1a3a5c;",
                        "STEP 2 \u2014 Apply to Plates / Antigens"),
                      uiOutput("pb_bulk_antigen_ui")
                    ),
                    column(2,
                      tags$label(style="font-size:12px; font-weight:600; color:#fff;", "."),
                      tags$div(
                        actionButton("pb_bulk_apply_ctrl",
                                     label = tagList(icon("check"), " Apply"),
                                     class = "btn btn-primary btn-sm", width = "100%")
                      )
                    )
                  )
                ),

                tags$hr(style="margin:6px 0 14px 0;"),

                tags$label(style="font-size:12px; font-weight:600; color:#555; text-transform:uppercase;",
                  "Fine-tune / Review a Single Plate / Antigen"),
                fluidRow(
                  column(3,
                    tags$label(style="font-size:13px; font-weight:600;", "Select Plate / Antigen"),
                    uiOutput("pb_antigen_selector_ui")
                  ),
                  column(9,
                    uiOutput("pb_config_row")
                  )
                )
              )
            ),

            # ---- QC Summary scorecards ----------------------------------------
            fluidRow(
              box(
                title = tags$span(icon("clipboard-check"), " QC Run -- 4 Quality Checks"),
                status = "success", solidHeader = TRUE, width = 12,
                uiOutput("pb_qc_scorecards")
              )
            ),

            # ---- Plate layout visual ------------------------------------------
            fluidRow(
              box(
                title = tags$span(icon("th"), " Plate Layout -- Readout"),
                status = "primary", solidHeader = TRUE, width = 7,
                fluidRow(
                  column(12,
                    tags$div(style="display:flex; align-items:center; gap:16px; margin-bottom:8px; flex-wrap:wrap;",
                      tags$strong("Readout key:"),
                      tags$span(style="display:inline-flex; align-items:center; gap:5px; background:#d4edda; border:1px solid #27ae60; border-radius:6px; padding:3px 10px; color:#155724; font-weight:700; font-size:12px;",
                        icon("circle"), " GREEN -- MFI > LOD (Detected / Reactive)"),
                      tags$span(style="display:inline-flex; align-items:center; gap:5px; background:#f8d7da; border:1px solid #e53935; border-radius:6px; padding:3px 10px; color:#721c24; font-weight:700; font-size:12px;",
                        icon("circle"), " RED -- MFI \u2264 LOD (Not Detected / Non-reactive)")
                    )
                  )
                ),
                uiOutput("pb_plate_map_ui")
              ),
              box(
                title = tags$span(icon("table"), " Per-Well MFI"),
                status = "info", solidHeader = TRUE, width = 5,
                DTOutput("pb_well_table")
              )
            ),

            # ---- QC Detail panels (sub-tabs: QC 1, 2, 3, 4) --------------------
            fluidRow(
              box(
                title = tags$span(icon("clipboard-list"), " QC Detail Panels"),
                status = "primary", solidHeader = TRUE, width = 12,
                uiOutput("pb_qc_antigen_label"),
                tabsetPanel(
                  id = "pb_qc_subtabs",
                  type = "tabs",

                  tabPanel(
                    title = tagList(icon("circle"), " QC 1: Bead Acquisition"),
                    tags$br(),
                    tags$div(class = "info-banner", icon("info-circle"),
                      HTML(" <strong>Criterion (INTELLIFLEX):</strong> &ge; Bead Count Threshold beads counted per well (default 50, configurable below).
                            <strong>Criterion (BioPlex):</strong> Sample wells signal &gt; LOD.
                            Blank (no_mab) wells are used as the reference baseline.")),
                    uiOutput("pb_qc1_result"),
                    tags$br(),
                    tags$h5(icon("table"), " All Samples — Bead Counts per Antigen",
                            style = "color:#1a3a5c; font-weight:bold; margin-top:10px;"),
                    tags$div(
                      style = "display:flex; align-items:center; gap:10px; margin-bottom:8px; flex-wrap:wrap;",
                      tags$label("% Agg Threshold:", style = "font-weight:600; margin:0;"),
                      numericInput("pb_agg_threshold", label = NULL, value = 5,
                                   min = 0, max = 100, step = 0.5, width = "100px"),
                      tags$span("\u2018% Agg Beads\u2019 values below this threshold are highlighted below.",
                                style = "font-size:12px; color:#666;"),
                      tags$label("Bead Count Threshold:", style = "font-weight:600; margin:0 0 0 16px;"),
                      numericInput("pb_bead_count_threshold", label = NULL, value = 50,
                                   min = 0, step = 1, width = "100px"),
                      tags$span("Bead count values \u2265 this threshold are highlighted (pass) below.",
                                style = "font-size:12px; color:#666;")
                    ),
                    DTOutput("pb_qc1_all_table")
                  ),

                  tabPanel(
                    title = tagList(icon("minus-circle"), " QC 2: Negative Controls"),
                    tags$br(),
                    tags$div(class = "info-banner", icon("info-circle"),
                      HTML(" <strong>Criteria:</strong> Blank Bead MFI &le; 1000 &nbsp;|&nbsp;
                            Blank Well MFI &minus; Background &lt; 1000.
                            Control wells with MFI &le; LOD are expected to read below the detection threshold.")),
                    uiOutput("pb_qc2_result"),
                    DTOutput("pb_qc2_table")
                  ),

                  tabPanel(
                    title = tagList(icon("plus-circle"), " QC 3: Positive Controls"),
                    tags$br(),
                    tags$div(class = "info-banner", icon("info-circle"),
                      HTML(" <strong>Criteria:</strong> &lt;30% CV for replicates &nbsp;|&nbsp;
                            MFI within 3 STDev from mean of LJ plots.
                            Positive control wells should have MFI &gt; LOD and show consistent replication.")),
                    uiOutput("pb_qc3_result"),
                    DTOutput("pb_qc3_table")
                  ),

                  tabPanel(
                    title = tagList(icon("check-double"), " QC 4: QC SUMMARY"),
                    tags$br(),
                    tags$div(class = "info-banner", icon("info-circle"),
                      HTML(" <strong>Criterion:</strong> All individual QC checks (Bead Acquisition,
                            Negative Controls, Positive Controls) must PASS.
                            System is suitable only when all criteria are met.")),
                    uiOutput("pb_qc4_result")
                  )
                )
              )
            )

          ) # end pb_qc_available conditionalPanel

        ) # end analysis_has_run conditionalPanel
      ),
      tabItem(tabName = "analysis_titration",
        tags$div(class = "tab-refresh-bar",
          actionButton("refresh_titration", label = tagList(icon("sync-alt"), " Refresh"),
            class = "tab-refresh-btn btn")
        ),
        uiOutput("analysis_titration_no_run_banner"),
        uiOutput("analysis_titration_bama_gate"),
        conditionalPanel(
          condition = "output.analysis_has_run && output.analysis_titration_selected",

          tabsetPanel(
            id   = "tit_subtabs",
            type = "tabs",

            # ================================================================
            # SUB-TAB 1: Titration
            # ================================================================
            tabPanel(
              title = tagList(icon("chart-line"), " Titration"),
              tags$br(),

              fluidRow(
                # -- Left filter panel ----------------------------------------
                column(3,
                  box(
                    title = tags$span(icon("sliders-h"), " Filters"),
                    status = "primary", solidHeader = TRUE, width = 12,

                    tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                               "Analyte:"),
                    uiOutput("tit_analyte_selector_ui"),

                    tags$br(),
                    tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                               "Control:"),
                    uiOutput("tit_control_selector_ui"),

                    tags$br(),
                    tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                               "Sample (sample_id):"),
                    uiOutput("tit_sample_selector_ui"),

                    tags$hr(style = "margin:10px 0;"),

                    tags$br(),
                    downloadButton("tit_save_png", label = tagList(icon("download"), " Save PNG"),
                                   class = "btn btn-default btn-sm",
                                   style = "width:100%; margin-top:6px;"),
                    tags$br(), tags$br(),
                    downloadButton("tit_save_ctrl_png", label = tagList(icon("download"), " Save Control PNG"),
                                   class = "btn btn-default btn-sm",
                                   style = "width:100%; margin-top:4px;")
                  )
                ),

                # -- Right: two stacked plots ----------------------------------
                column(9,
                  # Plot 1 -- titration curves (samples)
                  box(
                    title = uiOutput("tit_plot_title"),
                    status = "primary", solidHeader = TRUE, width = 12,
                    tags$div(
                      class = "info-banner",
                      icon("info-circle"),
                      HTML(" Titration curves -- average MFI per analyte vs concentration/dilution.
                            Analyte-specific; one curve per sample. Dotted line(s) show the
                            per-antigen LOD from Point-based QC Configuration.")
                    ),
                    plotOutput("tit_curve_plot", height = "380px")
                  ),

                  # Plot 2 -- QC plot per antigen (controls + samples, LOD line)
                  box(
                    title = tags$span(icon("dot-circle"), " QC Plot per Antigen"),
                    status = "info", solidHeader = TRUE, width = 12,
                    tags$div(
                      class = "info-banner",
                      icon("info-circle"),
                      HTML(" MFI values for <strong>control wells</strong>
                            (<span style='color:#e74c3c;font-weight:bold;'>&#9679; Positive</span> /
                             <span style='color:#6c6cdb;font-weight:bold;'>&#9679; Negative</span> -- classified from Point-based QC Configuration;
                             <span style='color:#95a5a6;'>&#9675; Unclassified</span> -- not yet assigned) and
                            <strong>sample MFIs</strong> plotted per analyte.
                            The <strong>LOD dashed line</strong> is drawn per analyte from the Point-based QC Configuration
                            (<span style='color:#c0392b;font-weight:bold;'>red</span> = locked,
                             <span style='color:#888;'>grey</span> = current unlocked value).")
                    ),
                    plotOutput("tit_ctrl_plot", height = "420px")
                  )
                )
              ),

              # -- Data table ---------------------------------------------------
              fluidRow(
                box(
                  title = tags$span(icon("table"), " Titration Data Table"),
                  status = "primary", solidHeader = TRUE, width = 12,
                  DTOutput("tit_data_table")
                )
              )
            ), # end Titration sub-tab

            # ================================================================
            # SUB-TAB 2: AUC
            # ================================================================
            tabPanel(
              title = tagList(icon("chart-bar"), " AUC"),
              tags$br(),

              fluidRow(
                # -- Left filter panel ----------------------------------------
                column(3,
                  box(
                    title = tags$span(icon("sliders-h"), " AUC Filters"),
                    status = "primary", solidHeader = TRUE, width = 12,

                    tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                               "Analyte:"),
                    uiOutput("auc_analyte_selector_ui"),

                    tags$br(),
                    tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                               "Sample Type:"),
                    uiOutput("auc_sample_type_selector_ui"),

                    tags$br(),
                    tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                               "Control:"),
                    uiOutput("auc_control_selector_ui"),

                    tags$br(),
                    tags$div(
                      class = "info-banner",
                      style = "font-size:11px; padding:8px 10px;",
                      icon("info-circle"),
                      HTML(" AUC = trapezoidal rule on<br>
                            log\u2081\u2080(concentration) \u00d7 MFI.
                            Calculated per sample per analyte.")
                    ),

                    tags$br(),
                    downloadButton("auc_save_png", label = tagList(icon("download"), " Save AUC PNG"),
                                   class = "btn btn-default btn-sm",
                                   style = "width:100%; margin-top:4px;")
                  )
                ),

                # -- Right: AUC bar chart --------------------------------------
                column(9,
                  box(
                    title = uiOutput("auc_plot_title"),
                    status = "success", solidHeader = TRUE, width = 12,
                    tags$div(
                      class = "info-banner",
                      icon("info-circle"),
                      HTML(" Area Under the Curve (AUC) per sample -- trapezoidal integration
                            on log\u2081\u2080(concentration) vs average MFI. Higher AUC = stronger
                            binding response.")
                    ),
                    plotOutput("auc_bar_plot", height = "420px")
                  )
                )
              ),

              # -- AUC summary table ---------------------------------------------
              fluidRow(
                box(
                  title = tags$span(icon("table"), " AUC Summary Table"),
                  status = "success", solidHeader = TRUE, width = 12,
                  DTOutput("auc_summary_table")
                )
              )
            ) # end AUC sub-tab

          ) # end tabsetPanel
        )
      ),
      tabItem(tabName = "analysis_quant",
        tags$div(class = "tab-refresh-bar",
          actionButton("refresh_quant", label = tagList(icon("sync-alt"), " Refresh"),
            class = "tab-refresh-btn btn")
        ),
        uiOutput("analysis_quant_no_run_banner"),
        uiOutput("analysis_quant_bama_gate"),
        conditionalPanel(
          condition = "output.analysis_has_run && output.analysis_quant_selected",

          fluidRow(
            # -- Left filter panel ----------------------------------------------
            column(3,
              box(
                title = tags$span(icon("sliders-h"), " Filters"),
                status = "primary", solidHeader = TRUE, width = 12,

                tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                           "Analyte:"),
                uiOutput("sc_analyte_selector_ui"),

                tags$br(),
                tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                           "Sample(s):"),
                uiOutput("sc_sample_selector_ui"),

                tags$hr(style = "margin:10px 0;"),

                tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                           "MFI subtraction method:"),
                tags$div(
                  style = "margin-top:4px;",
                  radioButtons(
                    "sc_subtraction_mode", NULL,
                    choices  = c("Background subtraction (blank bead + no_mAbs)" = "full",
                                 "no_mAbs subtraction only"                      = "mabs",
                                 "Blank beads subtraction only"                  = "blank"),
                    selected = "full",
                    inline   = FALSE
                  )
                ),
                tags$div(
                  class = "info-banner",
                  style = "font-size:11px; padding:6px 10px;",
                  icon("info-circle"),
                  HTML(" Uses the same Processed-data variants as the
                        <strong>Processed</strong> page. Standard curve fit
                        AND back-calculated sample concentrations both switch
                        to the selected method together.")
                ),

                tags$hr(style = "margin:10px 0;"),

                tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                           "X-axis scale:"),
                tags$div(
                  style = "margin-top:4px;",
                  radioButtons(
                    "sc_log_x", NULL,
                    choices  = c("Log\u2081\u2080 concentration" = "log",
                                 "Linear concentration"    = "linear"),
                    selected = "log",
                    inline   = FALSE
                  )
                ),

                tags$hr(style = "margin:10px 0;"),

                tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                           "X-axis labels:"),
                tags$div(
                  style = "margin-top:4px;",
                  radioButtons(
                    "sc_x_label_type", NULL,
                    choices  = c("Concentration (\u00b5g/mL)" = "conc",
                                 "Dilution ratio"            = "dilution"),
                    selected = "conc",
                    inline   = FALSE
                  )
                ),

                tags$hr(style = "margin:10px 0;"),

                tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                           "Fit range (log\u2081\u2080 concentration):"),
                tags$div(
                  class = "info-banner",
                  style = "font-size:11px; padding:6px 10px; margin-bottom:6px;",
                  icon("info-circle"),
                  HTML(" Drag the slider to select which points are included
                        in the line of best fit. Points outside the range are
                        still plotted but excluded from the fit &amp; R\u00b2.")
                ),
                uiOutput("sc_fit_range_ui"),

                tags$hr(style = "margin:10px 0;"),

                tags$label(style = "font-size:13px; font-weight:600; color:#1a3a5c;",
                           "Plot options:"),
                tags$div(
                  style = "margin-top:4px;",
                  checkboxInput("sc_show_r2", "Show R\u00b2 annotation", value = TRUE),
                  checkboxInput("sc_show_sample_labels",
                                "Show sample names on plot",
                                value = TRUE)
                ),

                tags$hr(style = "margin:10px 0;"),

                tags$div(
                  class = "info-banner",
                  style = "font-size:11px; padding:8px 10px;",
                  icon("info-circle"),
                  HTML(" One panel per sample within the selected analyte.<br>
                        Points at each dilution step are averaged replicates.<br>
                        Dashed line connects the standard curve points.<br>
                        Point size &prop; number of replicate wells.<br><br>
                        Requires <strong>sample_type = standard_curve</strong>
                        and <strong>start_concentration</strong> in the helper.")
                ),

                tags$br(),
                downloadButton("sc_save_png",
                               label = tagList(icon("download"), " Save PNG"),
                               class = "btn btn-default btn-sm",
                               style = "width:100%; margin-top:4px;")
              )
            ),

            # -- Right: standard curve plot -------------------------------------
            column(9,
              box(
                title = uiOutput("sc_plot_title"),
                status = "primary", solidHeader = TRUE, width = 12,
                plotOutput("sc_curve_plot", height = "560px")
              )
            )
          ),

          # -- Standard curve data table ---------------------------------------
          fluidRow(
            box(
              title = tags$span(icon("table"), " Standard Curve Data Table"),
              status = "primary", solidHeader = TRUE, width = 12,
              DTOutput("sc_data_table")
            )
          ),

          # -- Back-calculated sample concentrations ----------------------------
          fluidRow(
            box(
              title = tags$span(icon("calculator"), " Sample Concentration Estimates (Back-calculated from Standard Curve)"),
              status = "success", solidHeader = TRUE, width = 12,
              tags$div(
                class = "info-banner",
                icon("info-circle"),
                HTML(" For each <strong>sample</strong> well (sample_type = <em>sample</em>),
                      the average MFI is projected onto the selected standard curve fit
                      (within the current <strong>Fit range</strong>) to back-calculate
                      an estimated concentration (&micro;g/mL).<br>
                      &nbsp;&nbsp;&bull;&nbsp;<strong style='color:#c0392b;'>&gt; upper limit</strong>
                       &mdash; sample MFI is above the highest fitted standard curve point (above linear range).<br>
                      &nbsp;&nbsp;&bull;&nbsp;<strong style='color:#888;'>&lt; lower limit</strong>
                       &mdash; sample MFI is below the lowest fitted standard curve point.")
              ),
              DTOutput("sc_sample_conc_table")
            )
          )
        )
      )
    )
  )
)


# -- Server -------------------------------------------------------------------
server <- function(input, output, session) {

  # Utility: insert a new column into a data frame immediately after column index `after`
  append_col_after <- function(df, col_name, col_values, after) {
    if (after >= ncol(df)) {
      df[[col_name]] <- col_values
      return(df)
    }
    cbind(df[, seq_len(after), drop = FALSE],
          setNames(data.frame(col_values, stringsAsFactors = FALSE), col_name),
          df[, seq(after + 1, ncol(df)), drop = FALSE])
  }

  # Reactive storage
  rv <- reactiveValues(
    raw_df           = NULL,
    helper_df        = NULL,
    parsed_meta      = NULL,
    upload_error     = NULL,
    helper_edited    = NULL,
    well96_df        = NULL,
    well96_error     = NULL,
    plate_review_df  = NULL,
    calc_wb_path     = NULL,
    analysis_run     = FALSE,
    analysis_ts      = NULL,
    run_analyst      = NULL,
    run_date_used    = NULL,
    instrument       = "BIOPLEX",  # "BIOPLEX" or "INTELLIFLEX"
    ag_config        = list(),     # per-antigen locked config: list keyed by antigen col name
    brilliant_bioplex = TRUE       # BioPlex only: TRUE = BRILLIANT assay (LOD/LOQ editable),
                                    # FALSE = non-BRILLIANT (LOD/LOQ forced to 0)
  )

  # ---------------------------------------------------------------------------
  # Dynamic plate dimensions -- inferred from the raw file's Well column.
  # Returns list(n_rows, n_cols, row_letters, col_nums).
  # Defaults to 96-well (8x12) when no raw data is loaded.
  # ---------------------------------------------------------------------------
  plate_dims <- reactive({
    df <- rv$raw_df
    default <- list(n_rows = 8L, n_cols = 12L,
                    row_letters = LETTERS[1:8],
                    col_nums    = as.character(1:12))
    if (is.null(df) || !"Well" %in% colnames(df)) return(default)
    wells <- trimws(as.character(df[["Well"]]))
    wells <- wells[!is.na(wells) & nchar(wells) > 0]
    if (length(wells) == 0) return(default)

    row_ltrs <- toupper(gsub("[0-9]", "", wells))
    col_ints  <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", wells)))
    row_idx  <- match(row_ltrs, LETTERS)
    max_row  <- max(row_idx,  na.rm = TRUE)
    max_col  <- max(col_ints, na.rm = TRUE)

    if      (max_row <= 8L  && max_col <= 12L)  { max_row <- 8L;  max_col <- 12L }
    else if (max_row <= 16L && max_col <= 24L)  { max_row <- 16L; max_col <- 24L }

    list(n_rows      = max_row,
         n_cols      = max_col,
         row_letters = LETTERS[1:max_row],
         col_nums    = as.character(1:max_col))
  })

  # -- output flags consumed by conditionalPanel --
  output$analysis_has_run <- reactive({ isTRUE(rv$analysis_run) })
  outputOptions(output, "analysis_has_run", suspendWhenHidden = FALSE)

  # Reactive flags: TRUE when this BAMA type is currently selected on Overview
  # Used in conditionalPanel conditions to hide content when type not chosen
  # NOTE: checks BOTH bama_type_select_ov (Overview buttons) and bama_type_select
  # (helper panel hidden checkbox) so either input path unlocks the tab.
  output$analysis_point_selected <- reactive({
    sel_ov <- input$bama_type_select_ov
    sel_hp <- input$bama_type_select
    sel    <- unique(c(sel_ov, sel_hp))
    !is.null(sel) && length(sel) > 0 && "Point-based" %in% sel
  })
  output$analysis_titration_selected <- reactive({
    sel_ov <- input$bama_type_select_ov
    sel_hp <- input$bama_type_select
    sel    <- unique(c(sel_ov, sel_hp))
    !is.null(sel) && length(sel) > 0 && "Titration" %in% sel
  })
  output$analysis_quant_selected <- reactive({
    sel_ov <- input$bama_type_select_ov
    sel_hp <- input$bama_type_select
    sel    <- unique(c(sel_ov, sel_hp))
    !is.null(sel) && length(sel) > 0 && "Quantification" %in% sel
  })
  outputOptions(output, "analysis_point_selected",     suspendWhenHidden = FALSE)
  outputOptions(output, "analysis_titration_selected", suspendWhenHidden = FALSE)
  outputOptions(output, "analysis_quant_selected",     suspendWhenHidden = FALSE)

  output$is_bioplex <- reactive({ is.null(rv$instrument) || rv$instrument == "BIOPLEX" })
  outputOptions(output, "is_bioplex", suspendWhenHidden = FALSE)

  # Point-based QC Configuration is available for both BioPlex and
  # INTELLIFLEX -- both produce Beads_<antigen> bead counts (parsed from
  # ": RP1 COUNT" / ": RP2 COUNT" for INTELLIFLEX) used by the >=50 QC1 check,
  # and pos/neg controls are selected the same way for both instruments.
  output$pb_qc_available <- reactive({ TRUE })
  outputOptions(output, "pb_qc_available", suspendWhenHidden = FALSE)

  # -- Instrument selector ---------------------------------------------------
  observeEvent(input$instrument_select, {
    rv$instrument    <- input$instrument_select
    rv$analysis_run  <- FALSE
    rv$raw_df        <- NULL
    rv$upload_error  <- NULL
    rv$parsed_meta   <- NULL
    rv$helper_df     <- NULL
    rv$helper_edited <- NULL
    rv$well96_df     <- NULL
    rv$well96_error  <- NULL
    session$sendCustomMessage("updateInstrumentCards", input$instrument_select)
  }, ignoreInit = TRUE)


  # -- Sidebar: Upload Complete Helper File (.xlsx) ----------------------------
  observeEvent(input$sidebar_helper_file, {
    req(input$sidebar_helper_file)
    tryCatch({
      df <- suppressMessages(
        readxl::read_excel(input$sidebar_helper_file$datapath, sheet = 1)
      )
      df <- as.data.frame(df)
      # Trim leading/trailing whitespace from all character columns.
      # Prevents silent mismatches like "IgG " vs "IgG" in control lookups.
      df[] <- lapply(df, function(col) if (is.character(col)) trimws(col) else col)
      # Auto-populate sample_type from the type column.
      # Works whether sample_type already exists or is absent from the uploaded file.
      type_col <- which(tolower(trimws(colnames(df))) == "type")[1]
      st_col   <- which(tolower(trimws(colnames(df))) == "sample_type")[1]
      if (!is.na(type_col)) {
        tv <- toupper(trimws(as.character(df[[type_col]])))
        mapped <- dplyr::case_when(
          grepl("^[XUxu]\\d*$", tv)  ~ "sample",
          grepl("^[Bb]\\d*$",   tv)  ~ "blank bead",
          grepl("^[Ss]\\d*$",   tv)  ~ "standard curve",
          grepl("^[Cc]\\d*$",   tv)  ~ "control",
          TRUE                        ~ ""
        )
        if (is.na(st_col)) {
          # Column absent -- insert it right after the type column
          df <- append_col_after(df, "sample_type", mapped, after = type_col)
        } else {
          df[[st_col]] <- mapped
        }
      }
      # ---- Normalise to universal helper format --------------------------------
      lc_names <- tolower(trimws(colnames(df)))

      # Drop removed columns if present in old-format files
      drop_cols <- c("virus_id", "virus_control_range", "experiment_date", "control",
                     "plate_id", "type", "well", "concentration", "dilution_start")
      df <- df[, !(lc_names %in% drop_cols), drop = FALSE]
      lc_names <- tolower(trimws(colnames(df)))

      # Rename legacy column names to universal names
      rename_map <- c(
        "start_concentration" = "concentration",
        "instrument"          = "instrument_serial_number",
        "plate_range"         = "plate_range"   # already correct
      )
      for (new_nm in names(rename_map)) {
        old_nm <- rename_map[new_nm]
        if (old_nm %in% lc_names && !(new_nm %in% lc_names)) {
          colnames(df)[lc_names == old_nm] <- new_nm
          lc_names <- tolower(trimws(colnames(df)))
        }
      }

      # Add required columns if absent
      required_new <- c("plate_analyte", "plate_range", "no_mab_range",
                        "sample_type", "start_concentration", "dilution_factor",
                        "bama_type", "instrument", "scientist_id")
      for (rc in required_new) {
        if (!(rc %in% lc_names)) df[[rc]] <- ""
      }

      # Ensure plate_analyte is first column
      if ("plate_analyte" %in% colnames(df) && colnames(df)[1] != "plate_analyte") {
        df <- df[, c("plate_analyte", setdiff(colnames(df), "plate_analyte")), drop = FALSE]
      } else if ("plate_id" %in% colnames(df)) {
        # Rename legacy plate_id -> plate_analyte on old uploaded helpers
        colnames(df)[colnames(df) == "plate_id"] <- "plate_analyte"
      }

      # Pre-fill bama_type from checkbox if blank
      sel_bt <- input$bama_type_select_ov
      if (!is.null(sel_bt) && length(sel_bt) > 0) {
        bt_val <- paste(sel_bt, collapse = " + ")
        if ("bama_type" %in% colnames(df))
          df$bama_type[is.na(df$bama_type) | trimws(df$bama_type) == ""] <- bt_val
      }

      rv$helper_edited <- df
      showNotification(
        paste0("\u2714 Helper file loaded: ", nrow(rv$helper_edited), " rows"),
        type = "message", duration = 4
      )
    }, error = function(e) {
      showNotification(
        paste0("Could not read helper file: ", conditionMessage(e)),
        type = "error", duration = 6
      )
    })
  })

  observeEvent(input$sidebar_run_analysis, {
    # Validate prerequisites
    missing  <- c()
    is_ix    <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    if (is.null(rv$raw_df)) missing <- c(missing, "Raw plate data (.xlsx)")

    # For INTELLIFLEX: full helper file OR 96-well plate layout is sufficient.
    # For BioPlex: full helper file is required.
    if (is_ix) {
      if (is.null(rv$helper_edited) && is.null(rv$well96_df))
        missing <- c(missing, "Helper file (full helper .xlsx OR 96-well plate layout)")
    } else {
      if (is.null(rv$helper_edited)) missing <- c(missing, "Helper file")
    }

    analyst <- trimws(input$sidebar_analyst_name)
    if (nchar(analyst) == 0) missing <- c(missing, "Analyst Name")

    if (length(missing) > 0) {
      showNotification(
        paste0("Cannot run analysis. Missing: ", paste(missing, collapse = "; ")),
        type = "error", duration = 6
      )
      return()
    }

    # All checks passed -- mark analysis as run
    rv$analysis_run   <- TRUE
    rv$analysis_ts    <- format(Sys.time(), "%Y-%m-%d %H:%M:%S")
    rv$run_analyst    <- analyst
    rv$run_date_used  <- trimws(input$sidebar_analysis_date)

    showNotification(
      paste0("\u2714 Analysis run successfully at ", rv$analysis_ts),
      type = "message", duration = 5
    )
  })

  # -- Parse uploaded file ---------------------------------------------------
  observeEvent(input$raw_file, {
    rv$upload_error  <- NULL
    rv$parsed_meta   <- NULL
    rv$raw_df        <- NULL
    rv$helper_df     <- NULL
    rv$helper_edited <- NULL

    req(input$raw_file)
    path <- input$raw_file$datapath

    # Temporarily rename so filename parsing works
    tmp_path <- file.path(tempdir(), input$raw_file$name)
    file.copy(path, tmp_path, overwrite = TRUE)

    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    if (is_ix) {
      # ---- INTELLIFLEX: read the 'Results' sheet with flexible column detection ----
      tryCatch({
        # Find the Results sheet (case-insensitive)
        all_sheets    <- readxl::excel_sheets(tmp_path)
        results_sheet <- all_sheets[grepl("^results$", all_sheets, ignore.case = TRUE)][1]
        if (is.na(results_sheet))
          stop(paste0("No 'Results' sheet found. Available sheets: ",
                      paste(all_sheets, collapse = ", ")))

        raw_df <- suppressMessages(
          readxl::read_excel(tmp_path, sheet = results_sheet, col_names = TRUE)
        )
        raw_df <- as.data.frame(raw_df)
        cols   <- colnames(raw_df)

        # Normalise: strip leading region prefix (e.g. "R34: ") for matching
        cols_norm <- toupper(trimws(sub("^[A-Z0-9]+:\\s*", "", cols)))

        # Find WELL LOCATION column
        well_col_idx <- which(cols_norm == "WELL LOCATION")[1]
        if (is.na(well_col_idx))
          stop("Cannot locate 'WELL LOCATION' column in the Results sheet.")

        # Find all ANALYTE NAME columns (any region prefix)
        analyte_idx <- which(grepl("ANALYTE NAME$", cols_norm))
        # Find all MEDIAN columns: ends with "MEDIAN" but NOT "NET MEDIAN" or "AVERAGE MFI"
        median_idx  <- which(grepl("MEDIAN$", cols_norm) & !grepl("NET MEDIAN$", cols_norm))

        # Find all COUNT columns (RP1 COUNT / RP2 COUNT) -- bead counts per region
        count_idx <- which(grepl("COUNT$", cols_norm) & !grepl("CLASSIFIED|GATED|TOTAL", cols_norm))

        # Find TOTAL EVENTS / TOTAL GATED EVENTS -- well-level columns (one pair
        # per well, not per region) used to compute % Agg Beads:
        #   % Agg Beads = (1 - TOTAL GATED EVENTS / TOTAL EVENTS) * 100
        total_events_idx <- which(cols_norm == "TOTAL EVENTS")[1]
        total_gated_idx  <- which(cols_norm == "TOTAL GATED EVENTS")[1]

        if (length(analyte_idx) == 0)
          stop("Cannot locate any 'ANALYTE NAME' column in the Results sheet.")
        if (length(median_idx)  == 0)
          stop("Cannot locate any 'MEDIAN' column in the Results sheet.")

        # Rename WELL LOCATION -> Well (canonical name used by heatmap code)
        colnames(raw_df)[well_col_idx] <- "Well"

        # Match each COUNT column to its region prefix (e.g. "R34" from "R34: RP1 COUNT")
        count_region <- vapply(count_idx, function(idx) {
          m <- regmatches(cols[idx], regexpr("^[A-Z0-9]+(?=:\\s*)", cols[idx], perl = TRUE))
          if (length(m) > 0) m else NA_character_
        }, character(1))

        # Rename each analyte/median pair to app-canonical names
        for (i in seq_along(analyte_idx)) {
          a_idx <- analyte_idx[i]
          m_idx <- if (i <= length(median_idx)) median_idx[i] else NA_integer_

          # Extract region prefix tag (e.g. "R34" from "R34: RP1 ANALYTE NAME")
          prefix_match <- regmatches(cols[a_idx],
                           regexpr("^[A-Z0-9]+(?=:\\s*)", cols[a_idx], perl = TRUE))
          region_tag   <- if (length(prefix_match) > 0) prefix_match else paste0("R", i)

          # Analyte name column renamed to AnalyteName_<region>
          colnames(raw_df)[a_idx] <- paste0("AnalyteName_", region_tag)

          # Median column renamed to "<region> (NNN)" -- the "(NNN)" suffix triggers
          # antigen detection in plate_data_parsed()
          if (!is.na(m_idx)) {
            region_num <- suppressWarnings(as.integer(gsub("[^0-9]", "", region_tag)))
            region_num <- if (!is.na(region_num)) region_num else i * 10
            median_name <- paste0(region_tag, " (", region_num, ")")
            colnames(raw_df)[m_idx] <- median_name

            # Bead count for this region (": RP1 COUNT" / ": RP2 COUNT") ->
            # Beads_<median_name>, matching the naming used by QC1
            # (>= 50 beads) and the titration bead-count filter.
            c_idx <- count_idx[which(count_region == region_tag)][1]
            if (!is.na(c_idx))
              colnames(raw_df)[c_idx] <- paste0("Beads_", median_name)
          }
        }

        # ---- Per-analyte name + row count (mirrors BioPlex pre-helper logic) --
        # For BioPlex, generate_helper() loops "for each analyte" and records the
        # display name plus how many rows were built for it. INTELLIFLEX has no
        # sample/control classification available yet at this stage (that needs
        # the 96-well layout), so instead we capture, per ANALYTE NAME column,
        # the actual analyte name value and how many wells carry a parsed
        # (non-missing) MEDIAN result for it.
        analyte_name_cols <- grep("^AnalyteName_", colnames(raw_df), value = TRUE)
        analyte_summary <- lapply(analyte_name_cols, function(nc) {
          region_tag <- sub("^AnalyteName_", "", nc)
          median_col <- grep(paste0("^", region_tag, " \\("), colnames(raw_df), value = TRUE)[1]
          nm_vals    <- as.character(raw_df[[nc]])
          nm_vals    <- nm_vals[!is.na(nm_vals) & trimws(nm_vals) != ""]
          disp_name  <- if (length(nm_vals) > 0) nm_vals[1] else region_tag
          n_rows     <- if (!is.na(median_col)) sum(!is.na(raw_df[[median_col]])) else 0L
          list(name = disp_name, rows = n_rows)
        })
        analyte_names <- vapply(analyte_summary, function(x) x$name, character(1))
        analyte_rows  <- vapply(analyte_summary, function(x) x$rows, integer(1))
        names(analyte_rows) <- analyte_names

        # Provide Type/Description defaults (updated if 96-well layout is supplied)
        if (!"Type"        %in% colnames(raw_df)) raw_df$Type        <- "S"
        if (!"Description" %in% colnames(raw_df)) raw_df$Description <- as.character(raw_df[["Well"]])

        # ---- % Agg Beads (well-level QC metric) --------------------------------
        # INTELLIFLEX has no direct "% Agg Beads" column; it's derived from the
        # well-level TOTAL EVENTS / TOTAL GATED EVENTS columns captured above.
        # Same value applies to every antigen/region within that well.
        if (!is.na(total_events_idx) && !is.na(total_gated_idx)) {
          tot_events <- suppressWarnings(as.numeric(raw_df[[total_events_idx]]))
          tot_gated  <- suppressWarnings(as.numeric(raw_df[[total_gated_idx]]))
          raw_df$PctAggBeads <- ifelse(!is.na(tot_events) & tot_events > 0,
                                        (1 - (tot_gated / tot_events)) * 100,
                                        NA_real_)
        } else {
          raw_df$PctAggBeads <- NA_real_
        }

        rv$raw_df <- raw_df

        # Extract metadata from filename and file content
        fname_base  <- tools::file_path_sans_ext(basename(tmp_path))
        fname_parts <- str_split(fname_base, "_", n = 3)[[1]]
        exp_date    <- tryCatch(as.Date(fname_parts[1], "%Y%m%d"), error = function(e) Sys.Date())
        bama_type   <- if (length(fname_parts) >= 2) fname_parts[2] else "INTELLIFLEX"
        sci_id      <- if (length(fname_parts) >= 3) fname_parts[3] else ""

        # Serial number from SERIAL NUMBER column if present
        sn_col  <- which(toupper(trimws(colnames(raw_df))) == "SERIAL NUMBER")[1]
        inst_sn <- if (!is.na(sn_col) && nrow(raw_df) > 0)
                     as.character(raw_df[[sn_col]][1]) else NA_character_

        rv$parsed_meta <- list(
          filename           = fname_base,
          experimental_date  = format(exp_date, "%Y-%m-%d"),
          bama_type          = bama_type,
          scientist_id       = sci_id,
          instrument_serial  = inst_sn,
          helper             = NULL,
          plate_analyte      = paste(analyte_names, collapse = "; "),
          analyte_row_counts = analyte_rows
        )

        # Pre-fill Run Parameters fields (mirrors BioPlex behaviour)
        updateTextInput(session, "run_date",     value = gsub("-", "", format(exp_date, "%Y%m%d")))
        updateTextInput(session, "scientist_id", value = sci_id)
        if (!is.na(inst_sn)) updateTextInput(session, "instrument_sn", value = inst_sn)

      }, error = function(e) {
        rv$upload_error <- conditionMessage(e)
      })

    } else {
      # ---- BIOPLEX: original generate_helper() path ----
      tryCatch({
        meta       <- generate_helper(tmp_path)
        rv$parsed_meta  <- meta
        rv$helper_df    <- meta$helper
        rv$helper_edited <- meta$helper

        # Also read raw plate data for preview
        wb_raw     <- read.xlsx(tmp_path, sheet = 1, colNames = FALSE)
        header_row <- which(apply(wb_raw, 1, function(r) any(r == "Well", na.rm = TRUE)))[1]
        raw_plate  <- read.xlsx(tmp_path, sheet = 1, startRow = header_row, colNames = TRUE)
        rv$raw_df  <- raw_plate

        # Pre-fill Run Parameters
        updateTextInput(session, "run_date",      value = meta$experimental_date %>% gsub("-", "", .))
        # Pre-select matching BAMA type checkboxes from parsed filename --
        # merge with whatever is already selected (e.g. from the Overview
        # page) rather than overwriting it, so previously-checked types
        # like Titration / Quantification survive a raw-file upload.
        known_types  <- c("Point-based", "Titration", "Quantification")
        matched      <- known_types[tolower(known_types) == tolower(trimws(meta$bama_type))]
        existing_sel <- input$bama_type_select_ov %||% input$bama_type_select %||% character(0)
        combined_sel <- union(existing_sel, matched)
        if (length(combined_sel) == 0) combined_sel <- "Point-based"
        updateCheckboxGroupInput(session, "bama_type_select", selected = combined_sel)
        updateTextInput(session, "scientist_id",  value = meta$scientist_id)
        updateTextInput(session, "instrument_sn", value = meta$instrument_serial %||% "")

      }, error = function(e) {
        rv$upload_error <- conditionMessage(e)
      })
    }
  })

  # (%||% is defined at top level)

  # -- Upload status badge ---------------------------------------------------
  output$upload_status <- renderUI({
    if (!is.null(rv$upload_error)) {
      tags$span(class = "status-pill pill-error", icon("times-circle"), " Error")
    } else if (!is.null(rv$parsed_meta)) {
      tags$span(class = "status-pill pill-success", icon("check-circle"), " File loaded successfully")
    } else {
      tags$span(class = "status-pill pill-waiting", icon("hourglass-half"), " Awaiting upload")
    }
  })

  output$file_meta_display <- renderUI({
    if (!is.null(rv$upload_error)) {
      tags$div(class = "text-danger", tags$strong("Error: "), rv$upload_error)
    } else if (!is.null(rv$parsed_meta)) {
      m     <- rv$parsed_meta
      is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

      tags$div(
        tags$table(
          class = "table table-sm",
          style = "font-size:13px; margin-bottom:0;",
          tags$tbody(
            tags$tr(tags$td(tags$strong("Filename")),         tags$td(m$filename)),
            tags$tr(tags$td(tags$strong("Exp. Date")),        tags$td(m$experimental_date)),
            tags$tr(tags$td(tags$strong("Scientist ID")),     tags$td(m$scientist_id)),
            tags$tr(tags$td(tags$strong("Instrument S/N")),   tags$td(m$instrument_serial %||% "--")),
            tags$tr(tags$td(tags$strong("Analyte(s)")),        tags$td(m$plate_analyte %||% "--")),
            # BioPlex: single combined row count from the generated pre-helper.
            # INTELLIFLEX: no combined helper exists yet at this stage, so the
            # per-analyte breakdown below (captured from the raw file) is shown
            # instead of a single "Rows parsed" value here.
            if (!is_ix)
              tags$tr(tags$td(tags$strong("Rows parsed")),
                      tags$td(if (!is.null(m$helper)) nrow(m$helper) else "--"))
          )
        ),
        if (is_ix && !is.null(m$analyte_row_counts) && length(m$analyte_row_counts) > 0)
          tags$div(
            style = "margin-top:8px;",
            tags$strong(style = "font-size:12.5px;", "Rows parsed per Analyte Name"),
            tags$table(
              class = "table table-sm table-bordered",
              style = "font-size:12px; margin:4px 0 0;",
              tags$thead(tags$tr(tags$th("Analyte Name"), tags$th("Rows parsed"))),
              tags$tbody(
                lapply(seq_along(m$analyte_row_counts), function(i) {
                  tags$tr(
                    tags$td(names(m$analyte_row_counts)[i]),
                    tags$td(unname(m$analyte_row_counts[i]))
                  )
                })
              )
            )
          )
      )
    }
  })

  # -- Raw data note ---------------------------------------------------------
  output$raw_data_note <- renderUI({
    if (is.null(rv$raw_df)) {
      tags$div(class = "info-banner", icon("upload"), "  Upload a .xlsx file above to preview the raw plate data.")
    }
  })

  # -- Raw data table --------------------------------------------------------
  output$raw_data_table <- renderDT({
    req(rv$raw_df)
    df <- rv$raw_df

    # Only show first 15 cols for readability
    display_cols <- min(ncol(df), 12)
    df <- df[, seq_len(display_cols), drop = FALSE]

    datatable(
      df,
      options = list(
        pageLength = 15,
        scrollX    = TRUE,
        dom        = "lfrtip",
        columnDefs = list(list(className = "dt-center", targets = "_all"))
      ),
      rownames = FALSE,
      class    = "stripe hover cell-border compact"
    )
  })

  # -- Helper prerequisite banner --------------------------------------------
  output$helper_prereq_banner <- renderUI({
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    if (is_ix) {
      missing_parts <- character(0)
      if (is.null(rv$raw_df))     missing_parts <- c(missing_parts, "raw INTELLIFLEX file (Step 2)")
      if (is.null(rv$well96_df))  missing_parts <- c(missing_parts, "96-well plate layout (Step 3)")
      if (length(missing_parts) > 0) {
        tags$div(
          class = "info-banner",
          style = "font-size:14px; margin-bottom:16px; border-left-color:#6a1b9a; background:#f3e5f5;",
          icon("exclamation-circle"),
          HTML(paste0(" To auto-generate the INTELLIFLEX helper, please upload: <strong>",
                      paste(missing_parts, collapse = "</strong> and <strong>"),
                      "</strong>."))
        )
      } else {
        tags$div(
          class = "info-banner",
          style = "font-size:13px; margin-bottom:16px; border-left-color:#27ae60; background:#eafaf1;",
          icon("check-circle"),
          HTML(" All sources loaded. Select control wells in <strong>Step 3</strong>,
                fill in <strong>Scientist ID</strong> above, then click
                <strong>Generate / Refresh Helper</strong> to build the downloadable file.")
        )
      }
    } else {
      if (is.null(rv$parsed_meta)) {
        tags$div(
          class = "info-banner",
          style = "font-size:14px; margin-bottom:16px;",
          icon("arrow-left"), "  Please upload a raw .xlsx file on the ",
          tags$strong("Raw Data Upload"), " tab first to auto-populate the helper."
        )
      }
    }
  })

  # ===========================================================================
  # INTELLIFLEX -- auto-generate helper from raw file + 96-well layout
  # ===========================================================================

  # ---------------------------------------------------------------------------
  # Core INTELLIFLEX helper-build logic as a plain function so it can be called
  # from both the reactive (for the live preview table) and the download handler
  # (which runs outside a reactive context where req() would silently abort).
  # All inputs are passed explicitly -- no reactive reads, no req() calls.
  # ---------------------------------------------------------------------------
  .ix_build_helper_fn <- function(raw_df, w96, inst_sn, sci_id,
                                   ctrl_wells, sel_bt) {
    if (is.null(raw_df)) return(NULL)

    # -- 1. Analyte names from AnalyteName_* columns --------------------------
    an_cols <- grep("^AnalyteName_", colnames(raw_df), value = TRUE)
    if (length(an_cols) == 0) return(NULL)
    analyte_names <- unique(unlist(lapply(an_cols, function(col) {
      vals <- trimws(as.character(raw_df[[col]]))
      vals[!is.na(vals) & nchar(vals) > 0 & toupper(vals) != "NA" &
           toupper(vals) != "BLANK"]
    })))
    analyte_names <- analyte_names[nchar(analyte_names) > 0]
    if (length(analyte_names) == 0) return(NULL)

    # -- 2. Build well -> base sample_id map from 96-well layout ---------------
    # Strip dilution suffixes: _D1/_D2... (INTELLIFLEX) and _1/_2... (BioPlex)
    .strip_dil <- function(v) {
      v <- trimws(v)
      if (nchar(v) == 0 || toupper(v) %in% c("NO_MAB", "NA")) return(v)
      sub("_D?\\d+$", "", v)
    }

    well_sample_map <- character(0)
    if (!is.null(w96)) {
      row_col  <- colnames(w96)[1]
      num_cols <- colnames(w96)[-1]
      for (r in seq_len(nrow(w96))) {
        row_ltr <- as.character(w96[[row_col]][r])
        for (cc in num_cols) {
          val <- trimws(as.character(w96[[cc]][r]))
          if (!is.na(val) && nchar(val) > 0 && toupper(val) != "NA")
            well_sample_map[paste0(row_ltr, cc)] <- .strip_dil(val)
        }
      }
    }

    # -- 3. Identify unique samples and their well ranges ---------------------
    all_wells <- names(well_sample_map)
    if (length(all_wells) == 0) {
      all_wells <- unique(trimws(as.character(raw_df[["Well"]])))
      for (w in all_wells) well_sample_map[w] <- w
    }

    sample_wells <- list()
    no_mab_wells <- character(0)
    for (w in all_wells) {
      sid <- well_sample_map[w]
      if (is.na(sid) || nchar(sid) == 0) next
      if (tolower(trimws(sid)) == "no_mab") {
        no_mab_wells <- c(no_mab_wells, w)
      } else {
        if (is.null(sample_wells[[sid]])) sample_wells[[sid]] <- character(0)
        sample_wells[[sid]] <- c(sample_wells[[sid]], w)
      }
    }

    compact_range <- function(wells) {
      if (length(wells) == 0) return("")
      rows <- toupper(gsub("[0-9]", "", wells))
      cols <- suppressWarnings(as.integer(gsub("[^0-9]", "", wells)))
      df_w <- data.frame(r = rows, c = cols, stringsAsFactors = FALSE)
      df_w <- df_w[!is.na(df_w$c), ]
      if (nrow(df_w) == 0) return(paste(sort(wells), collapse = ", "))
      df_w <- df_w[order(df_w$c, match(df_w$r, LETTERS)), ]
      r1 <- df_w$r[1]; r2 <- df_w$r[nrow(df_w)]
      c1 <- df_w$c[1]; c2 <- df_w$c[nrow(df_w)]
      if (r1 == r2 && c1 == c2) return(paste0(r1, c1))
      paste0(r1, c1, ":", r2, c2)
    }

    no_mab_range_str <- compact_range(no_mab_wells)

    samp_ids <- names(sample_wells)
    if (length(samp_ids) == 0) return(NULL)

    bama_val <- if (length(sel_bt) > 0) paste(sel_bt, collapse = " + ") else ""
    inst_val <- if (!is.na(inst_sn) && nchar(inst_sn) > 0) inst_sn else ""

    rows_list <- list()
    exp_seq   <- 1L
    for (ag in analyte_names) {
      for (sid in samp_ids) {
        wells_this <- sample_wells[[sid]]
        stype <- if (any(wells_this %in% ctrl_wells)) "control" else NA_character_
        rows_list[[length(rows_list) + 1L]] <- data.frame(
          plate_analyte       = ag,
          experiment_id       = sprintf("EXP%03d", exp_seq),
          sample_id           = sid,
          plate_range         = compact_range(wells_this),
          no_mab_range        = no_mab_range_str,
          sample_type         = stype,
          start_concentration = NA_real_,
          dilution_factor     = NA_real_,
          bama_type           = bama_val,
          instrument          = inst_val,
          scientist_id        = sci_id,
          samples             = "",
          stringsAsFactors    = FALSE
        )
        exp_seq <- exp_seq + 1L
      }
    }
    do.call(rbind, rows_list)
  }

  # Thin reactive wrapper -- reads reactive values and passes them to the fn.
  ix_build_helper <- reactive({
    req(!is.null(rv$instrument) && rv$instrument == "INTELLIFLEX")
    req(!is.null(rv$raw_df))
    meta   <- rv$parsed_meta %||% list()
    sci_id <- trimws(input$scientist_id %||% "")
    if (nchar(sci_id) == 0) sci_id <- meta$scientist_id %||% ""
    sel_bt <- input$bama_type_select %||% input$bama_type_select_ov %||% character(0)
    .ix_build_helper_fn(
      raw_df     = rv$raw_df,
      w96        = rv$well96_df,
      inst_sn    = meta$instrument_serial %||% NA_character_,
      sci_id     = sci_id,
      ctrl_wells = input$ix_ctrl_wells %||% character(0),
      sel_bt     = sel_bt
    )
  })

  # Auto-rebuild whenever layout, controls, or bama_type selection change (INTELLIFLEX only)
  observe({
    req(!is.null(rv$instrument) && rv$instrument == "INTELLIFLEX")
    req(!is.null(rv$raw_df))
    # Track all inputs that affect the helper so observer re-fires on any change
    w96        <- rv$well96_df
    ctrl_wells <- input$ix_ctrl_wells %||% character(0)
    sel_bt     <- input$bama_type_select %||% input$bama_type_select_ov %||% character(0)
    sci_input  <- input$scientist_id %||% ""
    meta       <- rv$parsed_meta %||% list()
    sci_id     <- trimws(sci_input)
    if (nchar(sci_id) == 0) sci_id <- meta$scientist_id %||% ""
    df <- tryCatch(
      .ix_build_helper_fn(
        raw_df     = rv$raw_df,
        w96        = w96,
        inst_sn    = meta$instrument_serial %||% NA_character_,
        sci_id     = sci_id,
        ctrl_wells = ctrl_wells,
        sel_bt     = sel_bt
      ),
      error = function(e) NULL
    )
    if (!is.null(df) && nrow(df) > 0) rv$helper_edited <- df
  })

  # -- Generate / refresh helper ---------------------------------------------
  observeEvent(input$btn_generate, {
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    if (is_ix) {
      # -- INTELLIFLEX: rebuild from raw file + 96-well layout + UI inputs --
      meta   <- rv$parsed_meta %||% list()
      sci_id <- trimws(input$scientist_id %||% "")
      if (nchar(sci_id) == 0) sci_id <- meta$scientist_id %||% ""
      sel_bt <- input$bama_type_select %||% input$bama_type_select_ov %||% character(0)
      df <- tryCatch(
        .ix_build_helper_fn(
          raw_df     = rv$raw_df,
          w96        = rv$well96_df,
          inst_sn    = meta$instrument_serial %||% NA_character_,
          sci_id     = sci_id,
          ctrl_wells = input$ix_ctrl_wells %||% character(0),
          sel_bt     = sel_bt
        ),
        error = function(e) {
          showNotification(paste("Helper build failed:", conditionMessage(e)), type = "error")
          NULL
        }
      )
      if (is.null(df) || nrow(df) == 0) return()
      rv$helper_edited <- df
      showNotification(
        paste0("INTELLIFLEX helper generated: ", nrow(df), " rows (", 
               length(unique(df$plate_analyte)), " analyte(s) x ",
               length(unique(df$sample_id)), " sample(s))."),
        type = "message", duration = 5
      )
    } else {
      # -- BioPlex: original path --------------------------------------------
      req(rv$helper_df)
      df <- rv$helper_df
      sel_bt <- input$bama_type_select
      if (!is.null(sel_bt) && length(sel_bt) > 0 && "bama_type" %in% colnames(df))
        df$bama_type <- paste(sel_bt, collapse = " + ")
      if (nchar(trimws(input$scientist_id)) > 0) df$scientist_id             = input$scientist_id
      if (nchar(trimws(input$instrument_sn))> 0) df$instrument_serial_number = input$instrument_sn
      rv$helper_edited <- df
      showNotification("Helper refreshed with updated parameters.", type = "message")
    }
  })

  # -- Editable helper table -------------------------------------------------
  editable_cols <- c("plate_analyte", "sample_id", "plate_range", "no_mab_range",
                     "sample_type", "start_concentration", "dilution_factor", "samples")

  output$helper_table <- renderDT({
    req(rv$helper_edited)
    df <- rv$helper_edited

    col_idx <- which(colnames(df) %in% editable_cols) - 1  # 0-based for DT

    datatable(
      df,
      editable = list(target = "cell", disable = list(columns = setdiff(0:(ncol(df)-1), col_idx))),
      options  = list(
        pageLength = 20,
        scrollX    = TRUE,
        dom        = "lfrtip",
        columnDefs = list(list(className = "dt-center", targets = "_all"))
      ),
      rownames = FALSE,
      class    = "stripe hover cell-border compact"
    )
  }, server = TRUE)

  # Capture cell edits
  observeEvent(input$helper_table_cell_edit, {
    info <- input$helper_table_cell_edit
    df   <- rv$helper_edited
    df[info$row, info$col + 1] <- info$value
    rv$helper_edited <- df
  })

  # -- Row count display -----------------------------------------------------
  output$helper_row_count <- renderUI({
    req(rv$helper_edited)
    tags$span(
      style = "font-size:13px; color:#555; line-height:38px;",
      icon("table"), sprintf("  %d rows", nrow(rv$helper_edited))
    )
  })

  # -- Download: xlsx --------------------------------------------------------
  # Produces a file that faithfully mirrors the intelliflex_helper_trial.xlsx
  # structure: dark-blue frozen header, alternating row shading, exact column
  # widths, samples dropdown validation, and a colour-coded legend block
  # appended below the data rows (source-tag rows + Controls annotation rows).
  output$dl_helper_xlsx <- downloadHandler(
    filename = function() {
      base <- if (!is.null(rv$parsed_meta)) rv$parsed_meta$filename else "helper"
      paste0(base, "_helper.xlsx")
    },
    content = function(file) {
      is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

      # Resolve data -- three-level fallback:
      # 1. rv$helper_edited  (user-edited table, or auto-built by the observe())
      # 2. ix_build_helper() (build fresh on demand -- INTELLIFLEX only;
      #                       rv$helper_df is never set for INTELLIFLEX)
      # 3. rv$helper_df      (BioPlex freshly-parsed helper)
      # 4. empty placeholder (nothing uploaded yet)
      df <- if (!is.null(rv$helper_edited) && nrow(rv$helper_edited) > 0) {
              rv$helper_edited
            } else if (is_ix) {
              tryCatch({
                meta   <- rv$parsed_meta %||% list()
                sci_id <- trimws(input$scientist_id %||% "")
                if (nchar(sci_id) == 0) sci_id <- meta$scientist_id %||% ""
                sel_bt <- input$bama_type_select %||% input$bama_type_select_ov %||% character(0)
                .ix_build_helper_fn(
                  raw_df     = rv$raw_df,
                  w96        = rv$well96_df,
                  inst_sn    = meta$instrument_serial %||% NA_character_,
                  sci_id     = sci_id,
                  ctrl_wells = input$ix_ctrl_wells %||% character(0),
                  sel_bt     = sel_bt
                )
              }, error = function(e) NULL)
            } else if (!is.null(rv$helper_df) && nrow(rv$helper_df) > 0) {
              rv$helper_df
            } else {
              NULL
            }

      if (is.null(df) || nrow(df) == 0) {
        df <- data.frame(
          plate_analyte = character(0), experiment_id = character(0),
          sample_id = character(0), plate_range = character(0),
          no_mab_range = character(0), sample_type = character(0),
          start_concentration = numeric(0), dilution_factor = numeric(0),
          bama_type = character(0), instrument = character(0),
          scientist_id = character(0), samples = character(0)
        )
      }

      # Apply last-minute UI overrides for INTELLIFLEX so all auto-extractable
      # fields are stamped even if the user has not clicked Generate yet.
      if (is_ix && nrow(df) > 0) {
        sci <- trimws(input$scientist_id %||% "")
        if (nchar(sci) > 0 && "scientist_id" %in% colnames(df))
          df$scientist_id <- sci

        sel_bt <- input$bama_type_select %||% input$bama_type_select_ov %||% character(0)
        if (length(sel_bt) > 0 && "bama_type" %in% colnames(df)) {
          bv <- paste(sel_bt, collapse = " + ")
          df$bama_type[is.na(df$bama_type) | trimws(df$bama_type) == ""] <- bv
        }

        sn_ui <- trimws(input$instrument_sn %||% "")
        if (nchar(sn_ui) > 0 && "instrument" %in% colnames(df))
          df$instrument[is.na(df$instrument) | trimws(df$instrument) == ""] <- sn_ui
      }

      tryCatch({
        # -- Canonical column order & exact widths from the trial file ----------
        col_names   <- c("plate_analyte", "experiment_id", "sample_id", "plate_range",
                         "no_mab_range", "sample_type", "start_concentration",
                         "dilution_factor", "bama_type", "instrument",
                         "scientist_id", "samples")
        col_widths_def <- c(17.33, 16.16, 15.83, 16.66, 18.66, 17.5,
                            18.33, 16.16, 16.83, 17.33, 19.5, 16.5)

        for (cn in col_names) {
          if (!(cn %in% colnames(df))) df[[cn]] <- ""
        }
        df <- df[, col_names, drop = FALSE]
        n_data <- nrow(df)

        wb  <- createWorkbook()
        sht <- "Sheet1"
        addWorksheet(wb, sht)

        # -- Styles -------------------------------------------------------------
        hdr_style <- createStyle(
          fontName = "Arial", fontSize = 10,
          fontColour = "#FFFFFF", fgFill = "#1E4976",
          textDecoration = "bold", halign = "center",
          border = "TopBottomLeftRight", borderColour = "#CCCCCC"
        )
        style_odd <- createStyle(
          fontName = "Arial", fontSize = 10, fgFill = "#FFFFFF",
          border = "TopBottomLeftRight", borderColour = "#CCCCCC"
        )
        style_even <- createStyle(
          fontName = "Arial", fontSize = 10, fgFill = "#EEF2FF",
          border = "TopBottomLeftRight", borderColour = "#CCCCCC"
        )

        # Write header + data rows
        writeData(wb, sht, df, startRow = 1L, headerStyle = hdr_style)
        if (n_data > 0) {
          for (i in seq_len(n_data)) {
            sty <- if (i %% 2 == 1L) style_odd else style_even
            addStyle(wb, sht, style = sty, rows = i + 1L,
                     cols = seq_len(12L), gridExpand = TRUE)
          }
        }

        # Column widths and freeze pane
        setColWidths(wb, sht, cols = seq_len(12L), widths = col_widths_def)
        freezePane(wb, sht, firstRow = TRUE)

        # samples column dropdown (col 12)
        if (n_data > 0) {
          dataValidation(wb, sht,
            col = 12L, rows = seq(2L, n_data + 1L),
            type = "list", operator = "equal",
            value = '"serum,mab"'
          )
        }

        saveWorkbook(wb, file, overwrite = TRUE)
      }, error = function(e) {
        # Fallback: plain xlsx without styling
        openxlsx::write.xlsx(df, file, rowNames = FALSE)
      })
    }
  )

  # -- Download: csv ---------------------------------------------------------
  output$dl_helper_csv <- downloadHandler(
    filename = function() {
      base <- if (!is.null(rv$parsed_meta)) rv$parsed_meta$filename else "helper"
      paste0(base, "_helper.csv")
    },
    content = function(file) {
      is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
      df <- if (!is.null(rv$helper_edited) && nrow(rv$helper_edited) > 0)
              rv$helper_edited
            else if (is_ix)
              tryCatch({
                meta   <- rv$parsed_meta %||% list()
                sci_id <- trimws(input$scientist_id %||% "")
                if (nchar(sci_id) == 0) sci_id <- meta$scientist_id %||% ""
                sel_bt <- input$bama_type_select %||% input$bama_type_select_ov %||% character(0)
                .ix_build_helper_fn(
                  raw_df     = rv$raw_df,
                  w96        = rv$well96_df,
                  inst_sn    = meta$instrument_serial %||% NA_character_,
                  sci_id     = sci_id,
                  ctrl_wells = input$ix_ctrl_wells %||% character(0),
                  sel_bt     = sel_bt
                )
              }, error = function(e) NULL)
            else
              rv$helper_df
      req(!is.null(df) && nrow(df) > 0)
      write.csv(df, file, row.names = FALSE)
    }
  )

  # -- Download: empty template xlsx -----------------------------------------
  output$dl_empty_xlsx <- downloadHandler(
    filename = function() {
      base <- if (!is.null(rv$parsed_meta)) rv$parsed_meta$filename else "helper"
      paste0(base, "_template_empty.xlsx")
    },
    content = function(file) {
      # Build a zero-row data frame with the correct columns
      col_names <- c("plate_analyte", "experiment_id", "sample_id", "plate_range", "no_mab_range", "sample_type", "start_concentration", "dilution_factor", "bama_type", "instrument", "scientist_id", "samples")
      empty_df <- setNames(data.frame(matrix(ncol = length(col_names), nrow = 0)), col_names)

      wb <- createWorkbook()
      addWorksheet(wb, "Sheet1")

      header_style <- createStyle(
        fontName = "Arial", fontSize = 10, fontColour = "#FFFFFF",
        fgFill = "#1e4976", textDecoration = "bold", halign = "center",
        border = "TopBottomLeftRight", borderColour = "#CCCCCC"
      )
      writeData(wb, "Sheet1", empty_df, startRow = 1, headerStyle = header_style)
      setColWidths(wb, "Sheet1", cols = seq_len(ncol(empty_df)),
                   widths = c(28, 12, 24, 14, 20, 16, 20, 16, 18, 22, 14, 12))
      freezePane(wb, "Sheet1", firstRow = TRUE)
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )

  # -- Download: empty template csv ------------------------------------------
  output$dl_empty_csv <- downloadHandler(
    filename = function() {
      base <- if (!is.null(rv$parsed_meta)) rv$parsed_meta$filename else "helper"
      paste0(base, "_template_empty.csv")
    },
    content = function(file) {
      col_names <- c("plate_analyte", "experiment_id", "sample_id", "plate_range", "no_mab_range", "sample_type", "start_concentration", "dilution_factor", "bama_type", "instrument", "scientist_id", "samples")
      empty_df <- setNames(data.frame(matrix(ncol = length(col_names), nrow = 0)), col_names)
      write.csv(empty_df, file, row.names = FALSE)
    }
  )

  # -- BAMA type selector -- sync overview <-> helper panel, render pills --------
  # The overview tab has bama_type_select_ov; the helper panel has bama_type_select.
  # We keep them in sync via two observers so either one updates the other.
  # "Point-based" is a required assay type and is always kept selected.
  observeEvent(input$bama_type_select_ov, {
    isolate({
      sel_ov <- input$bama_type_select_ov
      if (!"Point-based" %in% sel_ov) {
        updateSelectizeInput(session, "bama_type_select_ov",
                              selected = c("Point-based", sel_ov))
        return()
      }
      if (!isTRUE(all(input$bama_type_select == input$bama_type_select_ov)))
        updateCheckboxGroupInput(session, "bama_type_select", selected = input$bama_type_select_ov)
    })
  }, ignoreNULL = FALSE)

  observeEvent(input$bama_type_select, {
    isolate({
      sel_hp <- input$bama_type_select
      if (!"Point-based" %in% sel_hp) {
        updateCheckboxGroupInput(session, "bama_type_select",
                                  selected = c("Point-based", sel_hp))
        return()
      }
      if (!isTRUE(all(input$bama_type_select == input$bama_type_select_ov)))
        updateSelectizeInput(session, "bama_type_select_ov", selected = input$bama_type_select)
    })
  }, ignoreNULL = FALSE)

  # Pill display for overview BAMA type box
  output$bama_type_pills_display <- renderUI({
    sel <- input$bama_type_select_ov
    if (is.null(sel) || length(sel) == 0) {
      tags$p(style = "color:#aaa; font-size:12px; margin-top:8px;",
             icon("info-circle"), " No type selected yet.")
    } else {
      pill_colours <- c(
        "Point-based"    = "#1e88e5",
        "Titration"      = "#43a047",
        "Quantification" = "#8e24aa"
      )
      tags$div(
        style = "margin-top:10px;",
        tags$strong(style = "font-size:12px; color:#555;", "Selected: "),
        lapply(sel, function(bt) {
          col <- unname(pill_colours[bt])
          if (is.na(col)) col <- "#555"
          tags$span(
            style = sprintf(
              "display:inline-block; background:%s; color:#fff; border-radius:12px;
               padding:3px 12px; font-size:11px; font-weight:700; margin:2px 3px;",
              col),
            icon("check"), " ", bt)
        })
      )
    }
  })

  # Dynamic sidebar: toggle analysis sub-items via shinyjs based on bama_type_select.
  # All three items are in the DOM; we show/hide their wrapper divs.
  observe({
    sel      <- input$bama_type_select
    show_all <- is.null(sel) || length(sel) == 0
    shinyjs::toggle("sidebar_item_point",     condition = show_all || "Point-based"    %in% sel)
    shinyjs::toggle("sidebar_item_titration", condition = show_all || "Titration"      %in% sel)
    shinyjs::toggle("sidebar_item_quant",     condition = show_all || "Quantification" %in% sel)
  })

  # -- Overview stat boxes ---------------------------------------------------
  output$ov_plates  <- renderUI({
    if (isTRUE(rv$analysis_run) && !is.null(rv$raw_df))
      tags$span(style="font-size:2.2rem; font-weight:700;", "1")
    else
      tags$span(style="font-size:2.2rem; font-weight:700;", "0")
  })
  output$ov_samples <- renderUI({
    if (isTRUE(rv$analysis_run) && !is.null(rv$raw_df) && "Type" %in% colnames(rv$raw_df)) {
      n <- sum(startsWith(as.character(rv$raw_df$Type), "S"), na.rm = TRUE)
      tags$span(style="font-size:2.2rem; font-weight:700;", n)
    } else {
      tags$span(style="font-size:2.2rem; font-weight:700;", "0")
    }
  })
  output$ov_viruses <- renderUI({
    if (isTRUE(rv$analysis_run) && !is.null(rv$raw_df)) {
      is_ag <- grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(rv$raw_df)) & !startsWith(colnames(rv$raw_df), "Beads_")
      ag_cols <- colnames(rv$raw_df)[is_ag]
      ag_cols <- ag_cols[!startsWith(toupper(ag_cols), "BLANK")]
      tags$span(style="font-size:2.2rem; font-weight:700;", length(ag_cols))
    } else {
      tags$span(style="font-size:2.2rem; font-weight:700;", "0")
    }
  })
  output$ov_run_summary <- renderUI({
    if (!isTRUE(rv$analysis_run)) {
      tags$p(style = "color:#555; font-size:14px; margin:0;", "No analysis run yet.")
    } else {
      tags$table(
        class = "table table-sm", style = "font-size:13px; margin:0;",
        tags$tbody(
          tags$tr(tags$td(tags$strong("Run timestamp")),  tags$td(rv$analysis_ts)),
          tags$tr(tags$td(tags$strong("Analyst")),        tags$td(rv$run_analyst)),
          tags$tr(tags$td(tags$strong("Analysis date")),  tags$td(rv$run_date_used)),
          tags$tr(tags$td(tags$strong("BAMA type")),
                  tags$td({
                    sel <- input$bama_type_select
                    if (!is.null(sel) && length(sel) == 0) sel <- NULL
                    if (is.null(sel) && !is.null(rv$parsed_meta) && !is.null(rv$parsed_meta$bama_type)) {
                      sel <- strsplit(as.character(rv$parsed_meta$bama_type), "\\s*\\+\\s*")[[1]]
                    }
                    if (!is.null(sel) && length(sel) > 0) {
                      # Generic palette cycled by position -- no assumptions about
                      # which specific assay type labels are present.
                      pill_palette <- c("#1e88e5", "#43a047", "#8e24aa", "#fb8c00", "#e53935", "#00897b")
                      pills <- lapply(seq_along(sel), function(i) {
                        col <- pill_palette[((i - 1) %% length(pill_palette)) + 1]
                        tags$span(
                          style = sprintf(
                            "display:inline-block; background:%s; color:#fff; border-radius:12px;
                             padding:2px 10px; font-size:11px; font-weight:700; margin:1px 3px 1px 0;",
                            col),
                          sel[i])
                      })
                      do.call(tags$div, c(list(style = "display:inline-block; line-height:1.8;"), pills))
                    } else {
                      "--"
                    }
                  })),
          tags$tr(tags$td(tags$strong("Raw rows")),
                  tags$td((if (!is.null(rv$raw_df)) nrow(rv$raw_df) else "--")))
        )
      )
    }
  })

  # -- 96-well plate upload (required INTELLIFLEX / optional BioPlex) --------
  observeEvent(input$well96_file, {
    req(input$well96_file)
    rv$well96_error <- NULL
    rv$well96_df    <- NULL
    skip_rows <- if (!is.null(rv$instrument) && rv$instrument == "INTELLIFLEX") 0 else 7
    tryCatch({
      df96 <- suppressMessages(
        read_excel(input$well96_file$datapath, col_names = TRUE, skip = skip_rows)
      )
      rv$well96_df <- as.data.frame(df96)
    }, error = function(e) {
      rv$well96_error <- conditionMessage(e)
    })
  })

  output$well96_status <- renderUI({
    if (!is.null(rv$well96_error)) {
      tags$span(class = "status-pill pill-error",
                icon("times-circle"), " Error: ", rv$well96_error)
    } else if (!is.null(rv$well96_df)) {
      tags$span(class = "status-pill pill-success",
                icon("check-circle"),
                sprintf(" Loaded: %d rows \u00d7 %d cols",
                        nrow(rv$well96_df), ncol(rv$well96_df)))
    } else {
      req_label <- if (!is.null(rv$instrument) && rv$instrument == "INTELLIFLEX")
        tags$span(class = "intelliflex-required-badge", "REQUIRED")
      else
        tags$span(style = "color:#888; font-size:12px;", "(optional)")
      tags$span(class = "status-pill pill-waiting",
                icon("hourglass-half"), " No file uploaded  ", req_label)
    }
  })

  # -- Dynamic upload tab titles ---------------------------------------------
  output$upload_box_title <- renderUI({
    if (!is.null(rv$instrument) && rv$instrument == "INTELLIFLEX")
      tags$span(icon("microchip"),
                " Step 2: Upload Raw INTELLIFLEX File (.xlsx) -- 'Results' sheet")
    else
      tags$span(icon("server"),
                " Step 2: Upload Raw BioPlex File (.xlsx) -- Plate Data Upload")
  })

  output$upload_filename_hint <- renderUI({
    if (!is.null(rv$instrument) && rv$instrument == "INTELLIFLEX")
      tags$small("Required sheet: ", tags$code("Results"),
                 " | Cols: WELL LOCATION + auto-detected ANALYTE NAME / MEDIAN per region")
    else
      tags$small("Expected filename format: ",
                 tags$code("YYYYMMDD_Study_ScientistID.xlsx"))
  })

  # -- 96-well box (required/optional depending on instrument) ---------------
  output$well96_box_ui <- renderUI({
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    if (is_ix) {
      box(
        title = tags$span(
          icon("table"), " Step 3: Upload 96-Well Plate Layout File (.xlsx)",
          tags$span(class = "intelliflex-required-badge", "REQUIRED for INTELLIFLEX")
        ),
        status = "primary", solidHeader = TRUE, width = 12,
        tags$div(
          class = "info-banner",
          style = "border-left-color:#6a1b9a; background:#f3e5f5;",
          icon("exclamation-circle"),
          HTML(" <strong>INTELLIFLEX requires this file.</strong>
                Format: first column = row letters (A-H),
                column headers = 1-12, cell values = sample IDs.
                Use <code>no_mab</code> for blank/no-antibody wells.")
        ),
        fluidRow(
          column(6,
            tags$div(class = "upload-zone", style = "border-color:#6a1b9a;",
              tags$div(class = "upload-icon", style = "color:#6a1b9a;", icon("table")),
              fileInput("well96_file", NULL,
                accept = c(".xlsx",".xls"), buttonLabel = "Browse Files...",
                placeholder = "96-well plate layout .xlsx"),
              tags$small("Rows A-H \u00d7 Cols 1-12, sample IDs in each cell")
            )
          ),
          column(6,
            tags$div(class = "helper-meta-card", style = "border-left-color:#6a1b9a;",
              tags$h4(icon("info-circle"), " Plate Layout Status"),
              uiOutput("well96_status")
            )
          )
        ),
        tags$hr(),
        tags$div(
          class = "info-banner", style = "border-left-color:#6a1b9a; background:#f3e5f5;",
          icon("info-circle"),
          HTML(" Select <strong>all control wells</strong> from the list below.
                Wells are populated from the uploaded plate layout file.
                Selected wells will be marked as <code>sample_type = control</code>
                in the generated helper.")
        ),
        tags$div(class = "helper-meta-card",
          style = "border-left-color:#6a1b9a; padding:12px 16px;",
          tags$div(
            style = "display:inline-block; background:#6a1b9a; color:#fff;
                     border-radius:12px; padding:2px 12px; font-size:11px;
                     font-weight:700; margin-bottom:8px;",
            icon("vial"), " Control Wells"
          ),
          tags$p(style = "font-size:12px; color:#555; margin-bottom:6px;",
            "Select all wells that are controls (positive or negative)."),
          uiOutput("ix_ctrl_wells_ui")
        )
      )
    } else {
      NULL
    }
  })

  # -- INTELLIFLEX: dynamic well-picker selectize inputs (from 96-well layout) --
  # Returns sorted well labels e.g. A1, A2 ... H12 from rv$well96_df
  # Single well-picker for all control wells (INTELLIFLEX)
  ix_well_choices <- reactive({
    w96 <- rv$well96_df
    if (is.null(w96)) return(character(0))
    row_col  <- colnames(w96)[1]
    num_cols <- colnames(w96)[-1]
    wells <- unlist(lapply(seq_len(nrow(w96)), function(r) {
      row_ltr <- as.character(w96[[row_col]][r])
      paste0(row_ltr, num_cols)
    }))
    labels <- unlist(lapply(seq_len(nrow(w96)), function(r) {
      row_ltr <- as.character(w96[[row_col]][r])
      sapply(num_cols, function(cc) {
        val <- as.character(w96[[cc]][r])
        if (!is.na(val) && nchar(trimws(val)) > 0)
          paste0(row_ltr, cc, "  --  ", trimws(val))
        else
          paste0(row_ltr, cc)
      })
    }))
    setNames(wells, labels)
  })

  output$ix_ctrl_wells_ui <- renderUI({
    choices <- ix_well_choices()
    selectizeInput(
      "ix_ctrl_wells", NULL,
      choices  = choices,
      selected = NULL,
      multiple = TRUE,
      options  = list(placeholder = "Select control well(s)...",
                      plugins     = list("remove_button")),
      width    = "100%"
    )
  })

  # -- Helper tab: lock banner for INTELLIFLEX (kept for backwards compat) ----
  output$helper_instrument_lock_banner <- renderUI({
    if (!is.null(rv$instrument) && rv$instrument == "INTELLIFLEX") {
      tags$div(
        class = "helper-locked-banner",
        tags$div(class = "lock-icon", icon("lock")),
        tags$h3("Helper Setup is not required for INTELLIFLEX"),
        tags$p(
          "The INTELLIFLEX pipeline reads sample identities directly from the ",
          tags$strong("96-well plate layout file"), " uploaded in Step 3 of the ",
          tags$strong("Plate Data Upload"), " tab. ",
          "Switch back to ", tags$strong("BioPlex"), " to use the Helper Setup."
        )
      )
    }
  })


  # Parse the raw uploaded file into per-well FI data for heatmaps
  plate_data_parsed <- reactive({
    req(rv$raw_df)
    df    <- rv$raw_df
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    # Identify antigen columns by "(NN)" suffix
    is_ag   <- grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(df)) & !startsWith(colnames(df), "Beads_")
    ag_cols <- colnames(df)[is_ag]
    ag_cols <- ag_cols[!startsWith(toupper(ag_cols), "BLANK")]

    if (length(ag_cols) == 0 || !"Well" %in% colnames(df))
      return(NULL)

    parse_fi <- function(x) {
      x <- gsub("\\s*\\(\\d+\\)\\s*$", "", as.character(x))
      suppressWarnings(as.numeric(gsub(",", ".", x)))
    }

    fi_mat   <- sapply(ag_cols, function(col) parse_fi(df[[col]]))
    if (is.vector(fi_mat)) fi_mat <- matrix(fi_mat, ncol = 1)
    rlu_vals <- rowMeans(fi_mat, na.rm = TRUE)

    well_ids  <- as.character(df[["Well"]])
    well_type <- if ("Type" %in% colnames(df)) as.character(df[["Type"]]) else rep("S", nrow(df))
    desc      <- if ("Description" %in% colnames(df)) as.character(df[["Description"]]) else well_type

    if (is_ix) {
      # INTELLIFLEX: sample names from 96-well layout file
      if (!is.null(rv$well96_df)) {
        w96     <- rv$well96_df
        row_col <- colnames(w96)[1]
        num_cols <- colnames(w96)[-1]
        ix_labels <- do.call(rbind, lapply(seq_len(nrow(w96)), function(r) {
          row_ltr <- as.character(w96[[row_col]][r])
          do.call(rbind, lapply(num_cols, function(cc) {
            data.frame(Well = paste0(row_ltr, cc),
                       SampleName = as.character(w96[[cc]][r]),
                       stringsAsFactors = FALSE)
          }))
        }))
        ix_map <- setNames(ix_labels$SampleName, ix_labels$Well)
        for (i in seq_along(well_ids)) {
          nm <- ix_map[well_ids[i]]
          if (!is.null(nm) && !is.na(nm)) desc[i] <- nm
        }
      }
      ctrl_all  <- trimws(isolate(input$ix_ctrl_wells) %||% character(0))
      ctrl_all  <- ctrl_all[nchar(ctrl_all) > 0]

      # Resolve pos/neg from saved ag_config or live pb_neg_ctrl input
      cur_ag_ix  <- isolate(input$pb_selected_antigen) %||% ""
      cur_cfg_ix <- rv$ag_config[[cur_ag_ix]]
      # Map ctrl wells -> sample names using the 96-well layout
      if (!is.null(rv$well96_df) && length(ctrl_all) > 0) {
        w96_t  <- rv$well96_df; rc_t <- colnames(w96_t)[1]; nc_t <- colnames(w96_t)[-1]
        lbl_t  <- do.call(rbind, lapply(seq_len(nrow(w96_t)), function(r)
                   do.call(rbind, lapply(nc_t, function(cc)
                     data.frame(Well = paste0(as.character(w96_t[[rc_t]][r]), cc),
                                Name = as.character(w96_t[[cc]][r]),
                                stringsAsFactors = FALSE)))))
        mp_t   <- setNames(lbl_t$Name, lbl_t$Well)
        ctrl_names_ix <- vapply(ctrl_all, function(w) {
          nm <- mp_t[[w]]; if (!is.null(nm) && !is.na(nm) && nchar(trimws(nm)) > 0) trimws(nm) else w
        }, character(1))
      } else {
        ctrl_names_ix <- setNames(ctrl_all, ctrl_all)
      }
      saved_neg_ix <- if (!is.null(cur_cfg_ix)) cur_cfg_ix$neg_ctrl %||% character(0) else character(0)
      live_neg_ix  <- isolate(input$pb_neg_ctrl)
      neg_names_ix <- if (length(live_neg_ix) > 0) live_neg_ix else saved_neg_ix
      neg_wells    <- ctrl_all[ctrl_names_ix %in% neg_names_ix]
      pos_wells    <- ctrl_all[!ctrl_all %in% neg_wells]

      well_type <- dplyr::case_when(
        length(pos_wells) > 0 & well_ids %in% pos_wells ~ "C_pos",
        length(neg_wells) > 0 & well_ids %in% neg_wells ~ "C_neg",
        desc == "no_mab"                                 ~ "B",
        TRUE                                             ~ well_type
      )
    } else {
      # BIOPLEX: use raw type + optional control-well overrides + helper/description detection
      ctrl_all  <- trimws(isolate(input$ix_ctrl_wells) %||% character(0))
      ctrl_all  <- ctrl_all[nchar(ctrl_all) > 0]
      pos_wells <- ctrl_all
      neg_wells <- character(0)

      # Build a well -> "C_pos" / "C_neg" / "C" map from helper sample_type column.
      # Uses: helper sample_type=="control" rows + ag_config pos_ctrl/neg_ctrl selections.
      helper_ctrl_map <- character(0)

      if (!is.null(rv$helper_edited)) {
        h       <- rv$helper_edited
        h_lc    <- setNames(h, tolower(trimws(colnames(h))))
        hc      <- colnames(h_lc)

        st_col_hb  <- grep("^sample_type$",   hc, value = TRUE)[1]
        sid_col_hb <- grep("^sample_id$",     hc, value = TRUE)[1]
        pr_col_hb  <- grep("^plate_range$",   hc, value = TRUE)[1]
        ag_col_hb  <- grep("^plate_analyte$", hc, value = TRUE)[1]

        # Get pos/neg selections for the current antigen from ag_config
        cur_ag    <- isolate(input$pb_selected_antigen) %||% ""
        cur_cfg   <- rv$ag_config[[cur_ag]]
        pos_sids  <- cur_cfg$pos_ctrl %||% character(0)
        neg_sids  <- cur_cfg$neg_ctrl %||% character(0)

        if (!is.na(st_col_hb) && !is.na(sid_col_hb) && !is.na(pr_col_hb)) {
          ctrl_mask_hb <- tolower(trimws(as.character(h_lc[[st_col_hb]]))) == "control"
          # Narrow to current analyte rows if possible
          if (!is.na(ag_col_hb) && nchar(cur_ag) > 0) {
            ag_match <- trimws(as.character(h_lc[[ag_col_hb]])) == cur_ag
            ctrl_mask_hb <- ctrl_mask_hb & ag_match
          }
          ctrl_h_rows <- which(ctrl_mask_hb)

          for (i in ctrl_h_rows) {
            sid_val <- trimws(as.character(h_lc[[sid_col_hb]][i]))
            pr_val  <- trimws(as.character(h_lc[[pr_col_hb]][i]))
            if (nchar(sid_val) == 0 || nchar(pr_val) == 0) next

            # Determine category for this control sample
            cat_val <- if (sid_val %in% pos_sids) "C_pos" else
                       if (sid_val %in% neg_sids) "C_neg" else "C"

            # Expand plate_range to individual wells
            segs <- trimws(unlist(strsplit(pr_val, ",")))
            for (seg in segs) {
              seg <- trimws(seg)
              if (grepl(":", seg)) {
                parts <- strsplit(seg, ":")[[1]]
                if (length(parts) != 2) next
                r1 <- toupper(gsub("[0-9]", "", trimws(parts[1])))
                c1 <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", trimws(parts[1]))))
                r2 <- toupper(gsub("[0-9]", "", trimws(parts[2])))
                c2 <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", trimws(parts[2]))))
                if (is.na(c1) || is.na(c2)) next
                ri1 <- match(r1, LETTERS); ri2 <- match(r2, LETTERS)
                if (is.na(ri1) || is.na(ri2)) next
                for (r in LETTERS[ri1:ri2])
                  for (cc in c1:c2)
                    helper_ctrl_map[paste0(r, cc)] <- cat_val
              } else if (nchar(seg) > 0) {
                helper_ctrl_map[seg] <- cat_val
              }
            }
          }
        }

        # Legacy fallback: helper has a 'control' column with "pos"/"neg" text
        if (length(helper_ctrl_map) == 0) {
          w_col_h <- which(hc == "well")[1]
          t_col_h <- which(hc == "type")[1]
          c_col_h <- which(hc == "control")[1]
          if (!is.na(w_col_h) && !is.na(t_col_h) && !is.na(c_col_h)) {
            for (i in seq_len(nrow(h_lc))) {
              wt  <- as.character(h_lc[[t_col_h]][i])
              wid <- as.character(h_lc[[w_col_h]][i])
              if (!startsWith(wt, "C")) next
              ctrl_val <- tolower(trimws(as.character(h_lc[[c_col_h]][i])))
              if      (grepl("pos", ctrl_val)) helper_ctrl_map[wid] <- "C_pos"
              else if (grepl("neg", ctrl_val)) helper_ctrl_map[wid] <- "C_neg"
            }
          }
        }
      }

      # Apply: explicit UI overrides > helper/description map > raw well_type
      mapped_types <- helper_ctrl_map[well_ids]
      well_type <- dplyr::case_when(
        length(pos_wells) > 0 & well_ids %in% pos_wells       ~ "C_pos",
        length(neg_wells) > 0 & well_ids %in% neg_wells       ~ "C_neg",
        !is.na(mapped_types) & nchar(mapped_types) > 0        ~ mapped_types,
        TRUE                                                   ~ well_type
      )
      desc <- ifelse(well_type == "B", "no_mab", desc)
    }

    data.frame(
      Well     = well_ids,
      RLU      = rlu_vals,
      WellType = well_type,
      Label    = substr(desc, 1, 12),
      SampleID = desc,
      stringsAsFactors = FALSE
    )
  })

  # Expand to full plate grid (size determined dynamically from raw data)
  full_grid <- reactive({
    dims      <- plate_dims()
    all_wells <- paste0(rep(dims$row_letters, each = dims$n_cols),
                        rep(seq_len(dims$n_cols), dims$n_rows))
    base <- data.frame(
      Well = all_wells,
      Row  = substr(all_wells, 1, 1),
      Col  = as.integer(sub("[A-Za-z]", "", all_wells)),
      stringsAsFactors = FALSE
    )
    pd <- plate_data_parsed()
    if (is.null(pd)) {
      base$RLU      <- NA_real_
      base$WellType <- "empty"
      base$Label    <- "no_sample"
    } else {
      base <- merge(base, pd, by = "Well", all.x = TRUE)
      base$RLU      <- ifelse(is.na(base$RLU), 0, base$RLU)
      base$WellType <- ifelse(is.na(base$WellType), "empty", base$WellType)
      base$Label    <- ifelse(is.na(base$Label), "no_sample", base$Label)
    }
    base$Row <- factor(base$Row, levels = rev(dims$row_letters))
    base$Col <- factor(base$Col, levels = seq_len(dims$n_cols))
    base
  })

  # -- No-run banners (shown on locked tabs before analysis is run) ----------
  .no_run_ui <- function(tab_label) {
    renderUI({
      if (!isTRUE(rv$analysis_run)) {
        tags$div(
          class = "no-run-banner",
          tags$div(class = "no-run-icon", icon("exclamation-triangle")),
          tags$h3("No data found for the current filter selection."),
          tags$p(
            paste0("The ", tab_label, " tab is available after a successful analysis run. ",
                   "Please upload your raw data and helper file, then click "),
            tags$strong("Run Analysis"), " in the sidebar."
          ),
          tags$span(class = "run-hint",
                    icon("play"), "  Run Analysis \u2192 sidebar")
        )
      }
    })
  }

  output$review_no_run_banner            <- .no_run_ui("Plate Review")
  output$dataframe_no_run_banner         <- .no_run_ui("Data Frame")
  output$export_no_run_banner            <- .no_run_ui("Export")

  # -- Export: active BAMA selection helper ----------------------------------
  .export_bama_sel <- function() {
    sel <- unique(c(input$bama_type_select_ov, input$bama_type_select))
    sel[!is.na(sel) & nchar(trimws(sel)) > 0]
  }

  # -- Export: dynamic sheet list shown in the UI card -----------------------
  output$export_sheet_list_ui <- renderUI({
    req(rv$analysis_run)
    sel  <- .export_bama_sel()
    base <- c("01_plate_setup", "02_plate_map", "03_raw_MFIs", "04_processed_MFI", "05_no_mAb_sub", "06_blank_sub", "07_background_sub")
    cond <- c(
      if ("Point-based"    %in% sel) "08_QC summary"     else NULL,
      if ("Point-based"    %in% sel) "09_bead_information" else NULL,
      if ("Point-based"    %in% sel) "10_control_plots"    else NULL,
      if ("Titration"      %in% sel) "11_titration_data"   else NULL,
      if ("Titration"      %in% sel) "12_titration_curves" else NULL,
      if ("Titration"      %in% sel) "13_auc_summary"      else NULL,
      if ("Titration"      %in% sel) "14_auc_plots"        else NULL,
      if ("Quantification" %in% sel) "15_standard_curve"   else NULL,
      if ("Quantification" %in% sel) "16_sample_concentration" else NULL
    )
    all_sheets <- c(base, cond)
    tags$p(
      style = "font-size:13px; color:#555; margin-bottom:8px;",
      tags$strong("Sheets: "),
      paste(all_sheets, collapse = ", "), "."
    )
  })

  # -- Export: run metadata display ------------------------------------------
  output$export_run_metadata_ui <- renderUI({
    req(rv$analysis_run)
    raw_rows    <- if (!is.null(rv$raw_df))       nrow(rv$raw_df)       else "--"
    helper_rows <- if (!is.null(rv$helper_edited)) nrow(rv$helper_edited) else "--"
    bama_sel    <- {
      sel <- .export_bama_sel()
      if (length(sel) > 0) paste(sel, collapse = " + ") else "--"
    }
    instr <- if (!is.null(rv$instrument) && nchar(trimws(rv$instrument)) > 0)
               rv$instrument else "BioPlex"

    tags$pre(
      style = paste0(
        "background:#f4f6f9; border:1px solid #d0d5dd; border-radius:6px;",
        "padding:14px 18px; font-size:13px; color:#1a3a5c; line-height:1.8;"
      ),
      paste0(
        "Run date    : ", rv$run_date_used,  "\n",
        "Analyst     : ", rv$run_analyst,     "\n",
        "Timestamp   : ", rv$analysis_ts,     "\n",
        "Instrument  : ", instr,              "\n",
        "BAMA type   : ", bama_sel,           "\n",
        "Raw rows    : ", raw_rows,           "\n",
        "Helper rows : ", helper_rows
      )
    )
  })

  # -- Export: legend title --------------------------------------------------
  output$export_legend_title_ui <- renderUI({
    sh <- input$export_legend_sheet %||% "11_titration_data"
    tags$span(icon("list-alt"), paste0("  Fields Legend \u2014 ", sh))
  })

  # -- Export: dynamic legend sheet selector --------------------------------
  output$export_legend_sheet_selector_ui <- renderUI({
    sel <- .export_bama_sel()
    base <- c("01_plate_setup", "02_plate_map", "03_raw_MFIs", "04_processed_MFI", "05_no_mAb_sub", "06_blank_sub", "07_background_sub")
    cond <- c(
      if ("Point-based"    %in% sel) "08_QC summary"     else NULL,
      if ("Point-based"    %in% sel) "09_bead_information" else NULL,
      if ("Point-based"    %in% sel) "10_control_plots"    else NULL,
      if ("Titration"      %in% sel) "11_titration_data"   else NULL,
      if ("Titration"      %in% sel) "12_titration_curves" else NULL,
      if ("Titration"      %in% sel) "13_auc_summary"      else NULL,
      if ("Titration"      %in% sel) "14_auc_plots"        else NULL,
      if ("Quantification" %in% sel) "15_standard_curve"   else NULL,
      if ("Quantification" %in% sel) "16_sample_concentration" else NULL
    )
    choices  <- c(base, cond)
    cur      <- isolate(input$export_legend_sheet)
    selected <- if (!is.null(cur) && cur %in% choices) cur else "11_titration_data"
    selectInput("export_legend_sheet", NULL,
                choices = choices, selected = selected, width = "100%")
  })

  # -- Export: field legends -------------------------------------------------
  .export_legends <- list(
    "01_plate_setup" = data.frame(
      Field       = c("plate_range", "sample_id", "sample_type", "samples",
                      "start_concentration", "dilution_factor", "control",
                      "analyst_id"),
      Description = c(
        "Well range covered by this helper row (e.g. A1:B2)",
        "Sample identifier as entered in the helper file",
        "Classification: sample / control / standard_curve",
        "Probe type used: mab or serum",
        "Starting concentration or dilution (numeric)",
        "Dilution factor applied at each step",
        "Control label, if applicable",
        "Analyst name as entered in the dashboard sidebar"
      ),
      stringsAsFactors = FALSE
    ),
    "02_plate_map" = data.frame(
      Field       = c("Plate Layout header", "Row", "<1..12>"),
      Description = c(
        "Single 8x12 plate-layout block -- sample/control/no_mab labels are the same for every antigen, so only one block is shown",
        "Plate row letter (A-H)",
        "Sample label / description for each well in that column"
      ),
      stringsAsFactors = FALSE
    ),
    "03_raw_MFIs" = data.frame(
      Field       = c("ANTIGEN: <name> header",
                      "Row", "<1..12>"),
      Description = c(
        "Antigen block label -- one 8x12 plate-map block per antigen",
        "Plate row letter (A-H)",
        "Raw MFI value in that well column (columns 1-12)"
      ),
      stringsAsFactors = FALSE
    ),
    "04_processed_MFI" = data.frame(
      Field       = c("ANTIGEN: <name> header",
                      "Row", "<1..12>"),
      Description = c(
        "Antigen block label -- one 8x12 plate-map block per antigen (same layout as 03_raw_MFIs)",
        "Plate row letter (A-H)",
        "Fully background-corrected MFI value in that well column (columns 1-12) -- blank-bead (R44/BLANK) + avg no_mab subtraction, negative values floored to 0; same values as 07_background_sub, shown in plate-map format"
      ),
      stringsAsFactors = FALSE
    ),
    "05_no_mAb_sub" = data.frame(
      Field       = c("Well", "Type", "Sample_ID", "<antigen_cols>"),
      Description = c(
        "Well identifier",
        "Well type code",
        "Sample identifier",
        "Raw MFI with only the average no_mab subtraction applied (blank-bead subtraction not included); negative values floored to 0"
      ),
      stringsAsFactors = FALSE
    ),
    "06_blank_sub" = data.frame(
      Field       = c("Well", "Type", "Sample_ID", "<antigen_cols>"),
      Description = c(
        "Well identifier",
        "Well type code",
        "Sample identifier",
        "Raw MFI with only the blank-bead (R44/BLANK) subtraction applied (no_mab subtraction not included); negative values floored to 0"
      ),
      stringsAsFactors = FALSE
    ),
    "07_background_sub" = data.frame(
      Field       = c("Well", "Type", "Sample_ID", "<antigen_cols>"),
      Description = c(
        "Well identifier",
        "Well type code",
        "Sample identifier",
        "Fully background-corrected MFI: blank-bead (R44/BLANK) + avg no_mab subtraction; negative values floored to 0"
      ),
      stringsAsFactors = FALSE
    ),
    "08_QC summary" = data.frame(
      Field       = c("Plot image", "One image per antigen"),
      Description = c(
        "QC Run summary -- PASS / CAUTION / N/A status for QC 1 (Bead Acquisition), QC 2 (Negative Controls), and QC 3 (Positive Controls) per antigen, matching the Point-based tab's QC Run scorecards (QC SUMMARY / overall status is not included here)",
        "Each antigen occupies its own embedded PNG in the sheet"
      ),
      stringsAsFactors = FALSE
    ),
    "09_bead_information" = data.frame(
      Field       = c("Well", "Type", "Sample_ID", "<antigen> Beads",
                      "% Agg Beads"),
      Description = c(
        "Well position",
        "Instrument well-type code",
        "Sample identifier",
        "Bead count per antigen -- green if >= the configurable Bead Count Threshold (default 50), red if below (matches the QC 1: Bead Acquisition table)",
        "% Agg Beads well-level acquisition-gating metric -- one shared column (same value applies to every antigen, for both BioPlex and INTELLIFLEX); amber-highlighted where the value exceeds the configurable % Agg Threshold set on the QC 1 tab"
      ),
      stringsAsFactors = FALSE
    ),
    "10_control_plots" = data.frame(
      Field       = c("Plot image", "One image per antigen"),
      Description = c(
        "Point-based QC Plot per Antigen -- MFI distribution for Positive Control, Negative Control, and Samples with LOD line",
        "Each antigen occupies its own embedded PNG in the sheet"
      ),
      stringsAsFactors = FALSE
    ),
    "11_titration_data" = data.frame(
      Field       = c("analyte", "Sample_ID", "Type", "sample_type", "sample_kind",
                      "dilution", "rep1_MFI", "rep2_MFI", "mean_MFI",
                      "per_std_dev"),
      Description = c(
        "Antigen / bead-region display name",
        "Sample identifier (both replicate IDs joined with '; ' when pivoted)",
        "Well type code (e.g. X1, X5) -- same code on two wells = replicate pair",
        "Classification: sample / control / standard_curve",
        "Probe type: mab or serum",
        "Concentration (ug/mL for mAb) or dilution step value (for serum) -- renamed from x_value",
        "MFI for the first replicate well",
        "MFI for the second replicate well",
        "Mean MFI across both replicates",
        "Coefficient of variation across replicates: sd(MFI) / mean(MFI) x 100 (NA if single replicate)"
      ),
      stringsAsFactors = FALSE
    ),
    "12_titration_curves" = data.frame(
      Field       = c("Plot image", "One image per analyte"),
      Description = c(
        "Titration curve -- average MFI per analyte vs concentration/dilution (log10 x-axis, loess smooth); dotted line shows the per-antigen LOD from Point-based QC Configuration (BRILLIANT BioPlex assays / INTELLIFLEX with a saved LOD only)",
        "Each analyte occupies its own embedded PNG in the sheet"
      ),
      stringsAsFactors = FALSE
    ),
    "13_auc_summary" = data.frame(
      Field       = c("analyte", "sample_id", "AUC", "n_points", "x_min", "x_max"),
      Description = c(
        "Antigen / bead-region display name",
        "Sample identifier",
        "Area under the titration curve (trapezoidal rule, log10-concentration axis)",
        "Number of dilution points used in AUC calculation",
        "Minimum concentration or dilution value included",
        "Maximum concentration or dilution value included"
      ),
      stringsAsFactors = FALSE
    ),
    "14_auc_plots" = data.frame(
      Field       = c("Plot image", "One image per analyte"),
      Description = c(
        "AUC bar chart -- area under the titration curve per sample",
        "Each analyte occupies its own embedded PNG in the sheet"
      ),
      stringsAsFactors = FALSE
    )
  )

  output$export_legend_table <- renderTable({
    sh  <- input$export_legend_sheet %||% "11_titration_data"
    leg <- .export_legends[[sh]]
    if (is.null(leg)) return(data.frame(Field = character(0), Description = character(0)))
    leg
  },
    striped  = TRUE, bordered = TRUE, hover = TRUE,
    spacing  = "s", width    = "100%", rownames = FALSE
  )

  # -- Export: download handler -----------------------------------------------
  output$dl_results_xlsx <- downloadHandler(
    filename = function() {
      inst_tag <- if (!is.null(rv$instrument) && rv$instrument == "INTELLIFLEX")
                    "inteliflex" else "bioplex"
      meta   <- rv$parsed_meta %||% list()
      sci_id <- trimws(input$scientist_id %||% "")
      if (nchar(sci_id) == 0) sci_id <- trimws(meta$scientist_id %||% "")
      if (nchar(sci_id) == 0) sci_id <- "ScientistID"
      sci_id   <- gsub("\\s+", "_", sci_id)
      run_date <- if (!is.null(rv$run_date_used) && nchar(trimws(rv$run_date_used)) > 0)
                    trimws(rv$run_date_used) else format(Sys.Date(), "%Y%m%d")
      ext <- if (inst_tag == "bioplex") "xls" else "xlsx"
      paste0("bama_results_", inst_tag, "_", sci_id, "_", run_date, ".", ext)
    },

    content = function(file) {
      wb      <- openxlsx::createWorkbook()
      is_ix   <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
      bama_sel <- .export_bama_sel()

      # -- Shared styles ------------------------------------------------------
      hdr_style <- openxlsx::createStyle(
        fontColour = "#FFFFFF", fgFill = "#1a3a5c",
        halign = "CENTER", textDecoration = "Bold",
        border = "Bottom", borderColour = "#AAAAAA"
      )
      ag_title_style <- openxlsx::createStyle(
        fgFill = "#1F4E79", fontColour = "white",
        textDecoration = "bold", halign = "left"
      )
      plate_hdr_style <- openxlsx::createStyle(
        fgFill = "#1F4E79", fontColour = "white",
        textDecoration = "bold", halign = "center", valign = "center"
      )
      border_style <- openxlsx::createStyle(
        border = "TopBottomLeftRight", borderColour = "black"
      )
      nomab_style   <- openxlsx::createStyle(fontColour = "#FF0000")
      nosample_style <- openxlsx::createStyle(fontColour = "#0070C0")

      # -- Helper: flat tabular sheet -----------------------------------------
      # Strip ANSI terminal escape codes (e.g. from drc warnings) so they
      # never reach openxlsx as invalid XML characters.
      .clean_msg <- function(msg) {
        gsub("\x1b\\[[0-9;]*m", "", msg, perl = TRUE)
      }

      .add_flat <- function(wb, name, df) {
        openxlsx::addWorksheet(wb, name)
        if (is.null(df) || nrow(df) == 0) {
          openxlsx::writeData(wb, name,
            data.frame(Note = "No data available for this sheet."))
          return(invisible(NULL))
        }
        openxlsx::writeData(wb, name, df, headerStyle = hdr_style)
        openxlsx::setColWidths(wb, name, cols = seq_len(ncol(df)), widths = "auto")
        openxlsx::freezePane(wb, name, firstRow = TRUE)
      }

      # Temp PNG files for plots -- cleaned up AFTER saveWorkbook, not before.
      # openxlsx stores the file PATH and reads it only at saveWorkbook() time,
      # so we must not delete these files with on.exit() inside .embed_plot.
      .plot_tmp_files <- character(0)

      # -- Helper: save a ggplot to a temp PNG and embed it ------------------
      .embed_plot <- function(wb, sheet_nm, p, width_in = 10, height_in = 5,
                              start_row = 1, start_col = 1) {
        # Use a persistent temp file -- NOT cleaned up here; cleaned after saveWorkbook
        tmp <- tempfile(fileext = ".png")
        .plot_tmp_files <<- c(.plot_tmp_files, tmp)

        # -- Render plot to PNG (try multiple devices) -------------------------
        save_ok <- tryCatch({
          # Method 1: ragg -- best for headless Linux, no X11 needed
          ggplot2::ggsave(tmp, plot = p,
                          device = ragg::agg_png,
                          width = width_in, height = height_in,
                          units = "in", res = 150)
          file.exists(tmp) && file.info(tmp)$size > 0
        }, error = function(e) FALSE)

        if (!isTRUE(save_ok)) {
          save_ok <- tryCatch({
            # Method 2: cairo PNG via ggsave
            ggplot2::ggsave(tmp, plot = p, device = "png",
                            type = "cairo",
                            width = width_in, height = height_in, dpi = 150)
            file.exists(tmp) && file.info(tmp)$size > 0
          }, error = function(e) FALSE)
        }

        if (!isTRUE(save_ok)) {
          save_ok <- tryCatch({
            # Method 3: base-R png() with cairo fallback
            grDevices::png(tmp,
                           width  = as.integer(width_in  * 150),
                           height = as.integer(height_in * 150),
                           res    = 150,
                           type   = if (capabilities("cairo")) "cairo" else "Xlib")
            print(p)
            grDevices::dev.off()
            file.exists(tmp) && file.info(tmp)$size > 0
          }, error = function(e) {
            tryCatch(grDevices::dev.off(), error = function(x) NULL)
            FALSE
          })
        }

        if (!isTRUE(save_ok)) {
          openxlsx::writeData(wb, sheet_nm,
            data.frame(Note = paste0(
              "Plot could not be rendered. ",
              "Install the 'ragg' package or ensure Cairo is available on this server.")),
            startRow = start_row)
          return(FALSE)
        }

        tryCatch({
          openxlsx::insertImage(wb, sheet_nm, tmp,
                                startRow = start_row, startCol = start_col,
                                width = width_in, height = height_in, units = "in")
          TRUE
        }, error = function(e) {
          openxlsx::writeData(wb, sheet_nm,
            data.frame(Note = paste("Image insert failed:", conditionMessage(e))),
            startRow = start_row)
          FALSE
        })
      }

      # -- Helper: antigen columns from raw_df (BioPlex & INTELLIFLEX) --------
      .get_ag_cols <- function(df, exclude_blank = TRUE) {
        if (is.null(df)) return(character(0))
        if (is_ix) {
          # INTELLIFLEX: antigen cols match the NN (region) suffix pattern
          cols <- colnames(df)
          cols <- cols[grepl("^.+\\s*\\(\\d+\\)\\s*$", cols) & !startsWith(cols, "Beads_")]
          cols <- cols[!grepl("^AnalyteName_", cols)]
          # R44 is the blank-bead region, not a real antigen/analyte for
          # analysis purposes -- excluded by default, but callers that need
          # every plate region represented (e.g. the raw plate-map / raw-MFI
          # export sheets) can pass exclude_blank = FALSE to keep it.
          if (exclude_blank)
            cols <- cols[!toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", cols))) %in% c("R44")]
          cols
        } else {
          # BioPlex: columns strictly between Description and the first
          # trailing-metadata column (Region, Gate, Total, Acquisition Time, etc.)
          # This mirrors the trailing_pat used in parse_plate_data().
          all_cols <- colnames(df)
          desc_idx <- which(tolower(all_cols) == "description")[1]
          if (is.na(desc_idx)) return(character(0))
          trailing_pat <- paste0("^(Region|Gate|Total|Location|X\\.Agg|%|",
                                 "Sampling|Plate\\.ID|Plate.ID|Bead\\.Count|",
                                 "Bead.Count|Acquisition|X\\.)")
          trailing_idx <- which(grepl(trailing_pat, all_cols, ignore.case = TRUE))[1]
          if (is.na(trailing_idx)) trailing_idx <- length(all_cols) + 1L
          cands <- all_cols[seq(desc_idx + 1L, trailing_idx - 1L)]
          cands <- cands[!grepl("BLANK", cands, ignore.case = TRUE)]
          cands <- cands[!startsWith(cands, "Beads_")]
          cands
        }
      }

      # -- Helper: resolve display (antigen/analyte) name for a raw MFI column
      # -- INTELLIFLEX: column is like "R26 (26)"; the real antigen name lives
      #    in the companion "AnalyteName_R26" column (same value every row).
      # -- BioPlex: the raw column name already *is* the antigen name.
      .get_ag_display_name <- function(col, df) {
        base_name <- gsub("\\s*\\(\\d+\\)\\s*$", "", col)
        if (is_ix) {
          region_tag  <- trimws(base_name)
          analyte_col <- paste0("AnalyteName_", region_tag)
          if (!is.null(df) && analyte_col %in% colnames(df)) {
            vals    <- as.character(df[[analyte_col]])
            non_na  <- vals[!is.na(vals) & nchar(trimws(vals)) > 0]
            if (length(non_na) > 0) return(trimws(non_na[1]))
          }
        }
        trimws(base_name)
      }

      # -- Helper: clean one MFI column to numeric ----------------------------
      .clean_mfi <- function(x) {
        x <- as.character(x)
        x <- gsub(",", ".", x)
        x <- gsub("\\s*\\(.*\\)", "", x)
        suppressWarnings(as.numeric(x))
      }

      # -- SHEET 1: plate_setup ---------------------------------------------
      {
        ps_df <- if (!is.null(rv$helper_edited)) rv$helper_edited else NULL
        if (!is.null(ps_df) && nrow(ps_df) > 0) {
          analyst_val <- if (!is.null(rv$run_analyst) && nchar(trimws(rv$run_analyst)) > 0)
                           trimws(rv$run_analyst) else NA_character_
          ps_df$analyst_id <- analyst_val
        }
        .add_flat(wb, "01_plate_setup", ps_df)
      }

      # -- SHEET 2: plate_data  (plate-map format -- single layout block) ----
      #   Row 1 : "Plate Layout"  (single cell -- labels are identical for
      #           every antigen, so only one block is emitted, unlike
      #           03_raw_MFIs which shows per-antigen raw values)
      #   Row 2 : "Row" | "1" | "2" | ... | "12"
      #   Rows 3-10: "A"-"H" | description per well
      openxlsx::addWorksheet(wb, "02_plate_map")
      tryCatch({
        df_raw_pd <- rv$raw_df
        if (is.null(df_raw_pd)) stop("No raw data available.")

        # Get Description (sample labels) per well
        wells_pd <- as.character(df_raw_pd[["Well"]])
        desc_pd  <- if ("Description" %in% colnames(df_raw_pd))
                      as.character(df_raw_pd[["Description"]])
                    else wells_pd

        # INTELLIFLEX: override labels from 96-well layout if available
        if (!is.null(rv$instrument) && rv$instrument == "INTELLIFLEX" &&
            !is.null(rv$well96_df)) {
          w96pd  <- rv$well96_df
          rc_pd  <- colnames(w96pd)[1]
          nc_pd  <- colnames(w96pd)[-1]
          lbl_pd <- do.call(rbind, lapply(seq_len(nrow(w96pd)), function(r) {
            rl_pd <- as.character(w96pd[[rc_pd]][r])
            do.call(rbind, lapply(nc_pd, function(cc_pd) {
              data.frame(Well = paste0(rl_pd, cc_pd),
                         SampleName = as.character(w96pd[[cc_pd]][r]),
                         stringsAsFactors = FALSE)
            }))
          }))
          mp_pd <- setNames(lbl_pd$SampleName, lbl_pd$Well)
          for (ii_pd in seq_along(wells_pd)) {
            nm_pd <- mp_pd[wells_pd[ii_pd]]
            if (!is.null(nm_pd) && !is.na(nm_pd)) desc_pd[ii_pd] <- nm_pd
          }
        }

        # Fallback labelling for wells with no Description text -- uses the
        # raw file's Type column (same B=no_mab convention as .classify_type):
        #   Type "B"   -> "no_mab"    (blank / no-antibody wells)
        #   anything else / blank -> "no_sample"
        # This covers BioPlex (whose Description column is often empty for
        # no_mab / no_sample wells) as well as any INTELLIFLEX wells not
        # covered by the 96-well layout override above.
        type_lookup_pd <- if ("Type" %in% colnames(df_raw_pd))
                             setNames(toupper(trimws(as.character(df_raw_pd[["Type"]]))), wells_pd)
                           else character(0)
        for (ii_pd in seq_along(desc_pd)) {
          if (is.na(desc_pd[ii_pd]) || nchar(trimws(desc_pd[ii_pd])) == 0) {
            t_code_pd <- type_lookup_pd[wells_pd[ii_pd]]
            desc_pd[ii_pd] <- if (!is.na(t_code_pd) && grepl("^B$", t_code_pd))
                                 "no_mab" else "no_sample"
          }
        }

        # Build a well -> label lookup
        label_lookup <- setNames(desc_pd, wells_pd)

        # Dynamic plate dimensions
        dims_pd  <- plate_dims()
        nrow_pd  <- dims_pd$n_rows
        ncol_pd  <- dims_pd$n_cols
        row_letters <- LETTERS[1:nrow_pd]
        col_nums    <- as.character(seq_len(ncol_pd))

        # Per-cell font styles for special well types
        nomab_font_style    <- openxlsx::createStyle(fontColour = "#FF0000")  # red
        nosample_font_style <- openxlsx::createStyle(fontColour = "#7030A0")  # purple
        control_font_style  <- openxlsx::createStyle(fontColour = "#7030A0")  # purple

        # Build set of control Description labels (Type ^C\d$ wells)
        ctrl_labels_pd <- character(0)
        tryCatch({
          if (!is.null(rv$raw_df) &&
              all(c("Type", "Description") %in% colnames(rv$raw_df))) {
            type_vals_pd <- toupper(trimws(as.character(rv$raw_df[["Type"]])))
            ctrl_mask_pd <- grepl("^C\\d*$", type_vals_pd)
            ctrl_labels_pd <- unique(trimws(as.character(
              rv$raw_df[["Description"]][ctrl_mask_pd])))
            ctrl_labels_pd <- ctrl_labels_pd[
              !is.na(ctrl_labels_pd) & nchar(ctrl_labels_pd) > 0]
          }
        }, error = function(e) NULL)

        pm_row <- 1L

        # -- Title row --------------------------------------------------------
        openxlsx::writeData(wb, "02_plate_map",
                            data.frame(V1 = "Plate Layout"),
                            startRow = pm_row, colNames = FALSE)
        openxlsx::addStyle(wb, "02_plate_map", ag_title_style,
                           rows = pm_row, cols = 1, stack = TRUE)
        pm_row <- pm_row + 1L

        # -- Header row: "Row" | "1" | "2" | ... | "12" -------------------
        hdr_row_df <- as.data.frame(
          matrix(c("Row", col_nums), nrow = 1),
          stringsAsFactors = FALSE)
        openxlsx::writeData(wb, "02_plate_map", hdr_row_df,
                            startRow = pm_row, colNames = FALSE)
        openxlsx::addStyle(wb, "02_plate_map", plate_hdr_style,
                           rows = pm_row, cols = 1:(ncol_pd + 1L),
                           gridExpand = FALSE, stack = TRUE)

        # -- Data rows: A-? x 1-? -----------------------------------------
        for (r_idx in seq_along(row_letters)) {
          row_ltr   <- row_letters[r_idx]
          cell_vals <- vapply(col_nums, function(col_n) {
            w <- paste0(row_ltr, col_n)
            v <- label_lookup[w]
            if (is.null(v) || is.na(v)) "" else as.character(v)
          }, character(1))
          data_row_df <- as.data.frame(
            matrix(c(row_ltr, cell_vals), nrow = 1),
            stringsAsFactors = FALSE)
          data_r <- pm_row + r_idx
          openxlsx::writeData(wb, "02_plate_map", data_row_df,
                              startRow = data_r, colNames = FALSE)
          # Row-letter header style in col 1
          openxlsx::addStyle(wb, "02_plate_map", plate_hdr_style,
                             rows = data_r, cols = 1L, stack = TRUE)

          # -- Per-cell color: no_mab=red, no_sample=purple, control=purple --
          for (c_idx in seq_along(cell_vals)) {
            val_raw <- trimws(cell_vals[c_idx])
            val_lc  <- tolower(val_raw)
            if (val_lc == "no_mab") {
              openxlsx::addStyle(wb, "02_plate_map", nomab_font_style,
                                 rows = data_r, cols = c_idx + 1L,
                                 stack = TRUE)
            } else if (val_lc == "no_sample") {
              openxlsx::addStyle(wb, "02_plate_map", nosample_font_style,
                                 rows = data_r, cols = c_idx + 1L,
                                 stack = TRUE)
            } else if (length(ctrl_labels_pd) > 0 &&
                       val_raw %in% ctrl_labels_pd) {
              openxlsx::addStyle(wb, "02_plate_map", control_font_style,
                                 rows = data_r, cols = c_idx + 1L,
                                 stack = TRUE)
            }
          }
        }

        # -- Border around the full block ----------------------------------
        openxlsx::addStyle(wb, "02_plate_map", border_style,
                           rows = pm_row:(pm_row + nrow_pd),
                           cols = 1:(ncol_pd + 1L),
                           gridExpand = TRUE, stack = TRUE)

        pm_row <- pm_row + nrow_pd + 2L  # 1 header + n data rows + 1 blank gap

        openxlsx::setColWidths(wb, "02_plate_map",
                               cols   = 1:(ncol_pd + 1L),
                               widths = c(4, rep(14, ncol_pd)))
      }, error = function(e) {
        openxlsx::writeData(wb, "02_plate_map",
          data.frame(Note = paste("Could not generate plate map:", conditionMessage(e))))
      })

      # -- SHEET 3: raw_MFIs  (plate-map format, one 8x12 block per antigen) -
      # Logic adapted from Bioplex-script.R, extended for INTELLIFLEX.
      tryCatch({
        df_raw   <- rv$raw_df
        req(!is.null(df_raw))

        ag_cols  <- .get_ag_cols(df_raw, exclude_blank = FALSE)
        if (length(ag_cols) == 0) stop("No antigen columns found in raw data.")

        # Ensure Row / Column exist
        df_raw <- df_raw %>%
          dplyr::mutate(
            Row    = gsub("[0-9]", "", Well),
            Column = suppressWarnings(as.numeric(gsub("[^0-9]", "", Well)))
          )

        full_plate <- expand.grid(Row    = plate_dims()$row_letters,
                                  Column = seq_len(plate_dims()$n_cols),
                                  stringsAsFactors = FALSE)

        # Build per-antigen plate-map blocks
        raw_mfi_list <- lapply(ag_cols, function(ag) {
          df_tmp <- df_raw %>%
            dplyr::mutate(value = .clean_mfi(.data[[ag]]))
          full_plate %>%
            dplyr::left_join(df_tmp[, c("Row", "Column", "value")],
                             by = c("Row", "Column")) %>%
            tidyr::pivot_wider(names_from = Column, values_from = value) %>%
            dplyr::arrange(Row)
        })
        names(raw_mfi_list) <- ag_cols

        # Write sheet
        openxlsx::addWorksheet(wb, "03_raw_MFIs")
        cur_row <- 1L

        for (ag in names(raw_mfi_list)) {
          plate <- raw_mfi_list[[ag]]

          # Antigen title row
          openxlsx::writeData(wb, "03_raw_MFIs",
                              data.frame(V1 = paste("ANTIGEN:",
                                          .get_ag_display_name(ag, df_raw))),
                              startRow = cur_row, colNames = FALSE)
          openxlsx::addStyle(wb, "03_raw_MFIs", ag_title_style,
                             rows = cur_row, cols = 1, stack = TRUE)
          cur_row <- cur_row + 1L

          # Plate map (header row + 8 data rows = 9 rows total)
          openxlsx::writeData(wb, "03_raw_MFIs", plate, startRow = cur_row)

          # Style: column headers (Row + 1..12)
          openxlsx::addStyle(wb, "03_raw_MFIs", plate_hdr_style,
                             rows = cur_row, cols = 1:(ncol(plate)),
                             gridExpand = FALSE, stack = TRUE)
          # Style: row labels (col 1, rows cur_row+1 .. cur_row+8)
          openxlsx::addStyle(wb, "03_raw_MFIs", plate_hdr_style,
                             rows = (cur_row + 1):(cur_row + plate_dims()$n_rows), cols = 1,
                             gridExpand = FALSE, stack = TRUE)
          openxlsx::addStyle(wb, "03_raw_MFIs", border_style,
                             rows = cur_row:(cur_row + plate_dims()$n_rows),
                             cols = 1:(ncol(plate)),
                             gridExpand = TRUE, stack = TRUE)

          cur_row <- cur_row + plate_dims()$n_rows + 2L
        }

        openxlsx::setColWidths(wb, "03_raw_MFIs",
                               cols   = 1:(plate_dims()$n_cols + 1L),
                               widths = c(4, rep(8, plate_dims()$n_cols)))

      }, error = function(e) {
        if (!"03_raw_MFIs" %in% names(wb$worksheets)) {
          openxlsx::addWorksheet(wb, "03_raw_MFIs")
        }
        openxlsx::writeData(wb, "03_raw_MFIs",
          data.frame(Note = paste("Could not generate raw MFI maps:", conditionMessage(e))))
      })

      # -- SHEET 4: processed_MFI (plate-map format, one 8x12 block per antigen) --
      # Same plate-map layout as 03_raw_MFIs, but showing the fully
      # background-corrected MFI values (blank-bead + avg no_mab subtraction,
      # from mfi_dataframe()$full -- i.e. the 07_background_sub values, shown
      # here in plate-map form instead of flat well-list form).
      tryCatch({
        mfi_pm <- mfi_dataframe()
        req(!is.null(mfi_pm) && !is.null(mfi_pm$full))
        df_proc <- mfi_pm$full
        df_proc <- df_proc[, !startsWith(colnames(df_proc), "AggPct_") &
                              !startsWith(colnames(df_proc), "Beads_"), drop = FALSE]

        ag_cols_proc <- setdiff(colnames(df_proc), c("Well", "Type", "Sample_ID"))
        if (length(ag_cols_proc) == 0) stop("No antigen columns found in processed data.")

        df_proc <- df_proc %>%
          dplyr::mutate(
            Row    = gsub("[0-9]", "", Well),
            Column = suppressWarnings(as.numeric(gsub("[^0-9]", "", Well)))
          )

        full_plate_proc <- expand.grid(Row    = plate_dims()$row_letters,
                                       Column = seq_len(plate_dims()$n_cols),
                                       stringsAsFactors = FALSE)

        # Build per-antigen plate-map blocks (antigen columns in df_proc are
        # already display names, so no lookup is needed for the title).
        proc_mfi_list <- lapply(ag_cols_proc, function(ag) {
          df_tmp <- df_proc[, c("Row", "Column", ag)]
          colnames(df_tmp)[3] <- "value"
          full_plate_proc %>%
            dplyr::left_join(df_tmp, by = c("Row", "Column")) %>%
            tidyr::pivot_wider(names_from = Column, values_from = value) %>%
            dplyr::arrange(Row)
        })
        names(proc_mfi_list) <- ag_cols_proc

        # Write sheet
        openxlsx::addWorksheet(wb, "04_processed_MFI")
        cur_row_proc <- 1L

        for (ag in names(proc_mfi_list)) {
          plate <- proc_mfi_list[[ag]]

          # Antigen title row
          openxlsx::writeData(wb, "04_processed_MFI",
                              data.frame(V1 = paste("ANTIGEN:", ag)),
                              startRow = cur_row_proc, colNames = FALSE)
          openxlsx::addStyle(wb, "04_processed_MFI", ag_title_style,
                             rows = cur_row_proc, cols = 1, stack = TRUE)
          cur_row_proc <- cur_row_proc + 1L

          # Plate map (header row + 8 data rows = 9 rows total)
          openxlsx::writeData(wb, "04_processed_MFI", plate, startRow = cur_row_proc)

          # Style: column headers (Row + 1..12)
          openxlsx::addStyle(wb, "04_processed_MFI", plate_hdr_style,
                             rows = cur_row_proc, cols = 1:(ncol(plate)),
                             gridExpand = FALSE, stack = TRUE)
          # Style: row labels (col 1, rows cur_row_proc+1 .. cur_row_proc+8)
          openxlsx::addStyle(wb, "04_processed_MFI", plate_hdr_style,
                             rows = (cur_row_proc + 1):(cur_row_proc + plate_dims()$n_rows), cols = 1,
                             gridExpand = FALSE, stack = TRUE)
          openxlsx::addStyle(wb, "04_processed_MFI", border_style,
                             rows = cur_row_proc:(cur_row_proc + plate_dims()$n_rows),
                             cols = 1:(ncol(plate)),
                             gridExpand = TRUE, stack = TRUE)

          cur_row_proc <- cur_row_proc + plate_dims()$n_rows + 2L
        }

        openxlsx::setColWidths(wb, "04_processed_MFI",
                               cols   = 1:(plate_dims()$n_cols + 1L),
                               widths = c(4, rep(8, plate_dims()$n_cols)))

      }, error = function(e) {
        if (!"04_processed_MFI" %in% names(wb$worksheets)) {
          openxlsx::addWorksheet(wb, "04_processed_MFI")
        }
        openxlsx::writeData(wb, "04_processed_MFI",
          data.frame(Note = paste("Could not generate processed MFI maps:", conditionMessage(e))))
      })

      # -- SHEET 5/6/7: no_mAb / blank-bead / background subtraction tables --
      # Pulled straight from the Processed tab's three parallel data frames
      # (mfi_dataframe()$mabs / $blank / $full). AggPct_* (QC-only) and
      # Beads_* columns are stripped -- bead counts are covered separately
      # in 09_bead_information.
      .strip_qc_cols <- function(d) {
        if (is.null(d)) return(NULL)
        d[, !startsWith(colnames(d), "AggPct_") & !startsWith(colnames(d), "Beads_"),
          drop = FALSE]
      }

      tryCatch({
        mfi <- mfi_dataframe()
        .add_flat(wb, "05_no_mAb_sub", .strip_qc_cols(if (!is.null(mfi)) mfi$mabs else NULL))
      }, error = function(e) {
        openxlsx::addWorksheet(wb, "05_no_mAb_sub")
        openxlsx::writeData(wb, "05_no_mAb_sub",
          data.frame(Note = paste("Could not generate no_mAb subtraction data:", conditionMessage(e))))
      })

      tryCatch({
        mfi <- mfi_dataframe()
        .add_flat(wb, "06_blank_sub", .strip_qc_cols(if (!is.null(mfi)) mfi$blank else NULL))
      }, error = function(e) {
        openxlsx::addWorksheet(wb, "06_blank_sub")
        openxlsx::writeData(wb, "06_blank_sub",
          data.frame(Note = paste("Could not generate blank-bead subtraction data:", conditionMessage(e))))
      })

      tryCatch({
        mfi <- mfi_dataframe()
        .add_flat(wb, "07_background_sub", .strip_qc_cols(if (!is.null(mfi)) mfi$full else NULL))
      }, error = function(e) {
        openxlsx::addWorksheet(wb, "07_background_sub")
        openxlsx::writeData(wb, "07_background_sub",
          data.frame(Note = paste("Could not generate background subtraction data:", conditionMessage(e))))
      })

      # -- SHEET 8: QC summary (Point-based only) ----------------------------
      # One image per antigen showing QC 1-3 pass/caution/N-A status, matching
      # the "QC Run" scorecard row in the Point-based tab (QC SUMMARY / overall
      # System Suitability is intentionally excluded here -- only 1 to 3).
      if ("Point-based" %in% bama_sel) {
        openxlsx::addWorksheet(wb, "08_QC summary")
        tryCatch({
          df_raw <- rv$raw_df
          if (is.null(df_raw)) {
            openxlsx::writeData(wb, "08_QC summary",
              data.frame(Note = "No raw data available."))
          } else {
            is_ag_qc   <- grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(df_raw)) &
                          !startsWith(colnames(df_raw), "Beads_")
            ag_cols_qc <- colnames(df_raw)[is_ag_qc]
            ag_cols_qc <- ag_cols_qc[!startsWith(toupper(ag_cols_qc), "BLANK")]
            ag_cols_qc <- ag_cols_qc[!toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "",
                                                          ag_cols_qc))) %in% c("R44")]
            if (length(ag_cols_qc) == 0) {
              openxlsx::writeData(wb, "08_QC summary",
                data.frame(Note = "No antigens detected."))
            } else {
              # Compact 3-card QC status graphic per antigen (mirrors the
              # "QC Run" scorecards shown 1-3 on the Point-based tab).
              .build_qc_summary_plot <- function(ag_name, p1, p2, p3) {
                card <- data.frame(
                  x        = 1:3,
                  num      = c("1", "2", "3"),
                  title    = c("Bead Acquisition", "Negative Controls", "Positive Controls"),
                  subtitle = c("\u2265 50 beads / well", "(Blank & Blank Well MFI)",
                               "(Low & High Pos Ctrl)"),
                  pass     = c(p1, p2, p3),
                  stringsAsFactors = FALSE
                )
                card$badge <- ifelse(is.na(card$pass), "N/A",
                                ifelse(card$pass, "PASS", "CAUTION"))
                card$color <- ifelse(is.na(card$pass), "#888888",
                                ifelse(card$pass, "#27ae60", "#e6a817"))
                card$bg    <- ifelse(is.na(card$pass), "#f8f9fa",
                                ifelse(card$pass, "#f0fff4", "#fffbea"))
                ggplot2::ggplot(card) +
                  ggplot2::geom_rect(
                    ggplot2::aes(xmin = x - 0.45, xmax = x + 0.45, ymin = 0, ymax = 1,
                                 fill = bg, colour = color),
                    linewidth = 1.1) +
                  ggplot2::geom_text(
                    ggplot2::aes(x = x, y = 0.80, label = paste0(num, ". ", title)),
                    size = 3.0, fontface = "bold", colour = "#555555") +
                  ggplot2::geom_text(
                    ggplot2::aes(x = x, y = 0.55, label = badge, colour = color),
                    size = 6.5, fontface = "bold") +
                  ggplot2::geom_text(
                    ggplot2::aes(x = x, y = 0.28, label = subtitle),
                    size = 2.6, colour = "#555555") +
                  ggplot2::scale_fill_identity() +
                  ggplot2::scale_colour_identity() +
                  ggplot2::coord_cartesian(xlim = c(0.4, 3.6), ylim = c(0, 1)) +
                  ggplot2::labs(title = paste("QC Run \u2014", ag_name)) +
                  ggplot2::theme_void(base_size = 12) +
                  ggplot2::theme(
                    plot.title  = ggplot2::element_text(face = "bold", size = 13,
                                    colour = "#1a3a5c", hjust = 0.5),
                    plot.margin = ggplot2::margin(10, 10, 10, 10)
                  )
              }

              img_row <- 1L; img_h <- 2.6; img_w <- 7.5
              for (ag_col in ag_cols_qc) {
                tryCatch({
                  cfg  <- .ag_cfg_get(ag_col)
                  d_ag <- .pb_antigen_data_core(ag_col, cfg$lod)
                  q1   <- .pb_qc1_core(d_ag, cfg$lod, ag_col,
                                       bead_thresh = suppressWarnings(as.numeric(input$pb_bead_count_threshold)))
                  q2   <- .pb_qc2_core(d_ag, cfg$lod, cfg$neg_ctrl)
                  q3   <- .pb_qc3_core(d_ag, cfg$lod, cfg$cv, cfg$pos_ctrl)
                  ag_display <- .ag_raw_to_display(ag_col)
                  p_qc <- .build_qc_summary_plot(ag_display, q1$pass, q2$pass, q3$pass)
                  .embed_plot(wb, "08_QC summary", p_qc,
                              width_in = img_w, height_in = img_h,
                              start_row = img_row, start_col = 1)
                  img_row <- img_row + ceiling(img_h * 6) + 1L
                }, error = function(e) {
                  # Don't let one antigen's failure blank out the whole sheet --
                  # log a note in its place and keep going with the rest.
                  openxlsx::writeData(wb, "08_QC summary",
                    data.frame(Note = paste0("Could not generate QC summary for '",
                                             ag_col, "': ", .clean_msg(conditionMessage(e)))),
                    startRow = img_row)
                  img_row <<- img_row + 2L
                })
              }
            }
          }
        }, error = function(e) {
          openxlsx::writeData(wb, "08_QC summary",
            data.frame(Note = paste("Could not generate QC summary:",
                                    .clean_msg(conditionMessage(e)))))
        })
      } # end Point-based (08_QC summary)

      # -- SHEET 9: bead_information (Point-based only) -----------------------
      # Static export of the "All Samples -- Bead Counts per Antigen" table
      # shown on the QC 1: Bead Acquisition tab, including the % Agg Beads
      # column(s) and the same green/red (>=50 beads) and amber (% Agg over
      # threshold) colour coding used on-screen.
      if ("Point-based" %in% bama_sel) {
        openxlsx::addWorksheet(wb, "09_bead_information")
        tryCatch({
          mfi_lst_bd <- mfi_dataframe()
          if (is.null(mfi_lst_bd) || is.null(mfi_lst_bd$full)) {
            openxlsx::writeData(wb, "09_bead_information",
              data.frame(Note = "No bead-count data available."))
          } else {
            df_full_bd <- mfi_lst_bd$full
            is_ix_bd   <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

            Beads_cols_bd <- grep("^Beads_", colnames(df_full_bd), value = TRUE)
            if (length(Beads_cols_bd) == 0) {
              openxlsx::writeData(wb, "09_bead_information",
                data.frame(Note = "No bead-count columns found."))
            } else {
              ag_names_bd    <- sub("^Beads_", "", Beads_cols_bd)
              AggPct_cols_bd <- paste0("AggPct_", ag_names_bd)
              fixed_cols_bd  <- c("Well", "Type", "Sample_ID")
              have_agg_bd    <- AggPct_cols_bd %in% colnames(df_full_bd)
              beads_disp_bd  <- paste0(ag_names_bd, " Beads")

              # % Agg Beads is a well-level metric -- the same value applies to
              # every antigen (both BioPlex and INTELLIFLEX), so it is shown
              # only once, aligned by well, rather than once per antigen.
              agg_disp_bd   <- "% Agg Beads"
              src_order_bd  <- c(fixed_cols_bd,
                                 if (any(have_agg_bd)) AggPct_cols_bd[which(have_agg_bd)[1]] else NULL,
                                 Beads_cols_bd)
              disp_names_bd <- c(fixed_cols_bd,
                                 if (any(have_agg_bd)) agg_disp_bd else NULL,
                                 beads_disp_bd)
              bead_df_bd <- df_full_bd[, src_order_bd, drop = FALSE]
              colnames(bead_df_bd) <- disp_names_bd

              # Round % Agg columns to 2 dp, matching the on-screen table
              agg_disp_present_bd <- agg_disp_bd[agg_disp_bd %in% colnames(bead_df_bd)]
              for (col in agg_disp_present_bd) {
                bead_df_bd[[col]] <- round(suppressWarnings(as.numeric(bead_df_bd[[col]])), 2)
              }

              openxlsx::writeData(wb, "09_bead_information", bead_df_bd, headerStyle = hdr_style)
              openxlsx::setColWidths(wb, "09_bead_information",
                                     cols = seq_len(ncol(bead_df_bd)), widths = "auto")
              openxlsx::freezePane(wb, "09_bead_information", firstRow = TRUE)

              n_rows_bd <- nrow(bead_df_bd)
              if (n_rows_bd > 0) {
                # Colour each bead-count column: green if >= threshold, red if < threshold
                # (same input$pb_bead_count_threshold used on-screen)
                bead_thresh_bd <- suppressWarnings(as.numeric(input$pb_bead_count_threshold))
                if (length(bead_thresh_bd) == 0 || is.na(bead_thresh_bd) || !is.finite(bead_thresh_bd))
                  bead_thresh_bd <- 50
                pass_style_bd <- openxlsx::createStyle(fontColour = "#155724",
                                   fgFill = "#d4edda", textDecoration = "bold")
                fail_style_bd <- openxlsx::createStyle(fontColour = "#721c24",
                                   fgFill = "#f8d7da", textDecoration = "bold")
                for (col in beads_disp_bd) {
                  if (col %in% colnames(bead_df_bd)) {
                    col_idx_bd   <- which(colnames(bead_df_bd) == col)
                    vals_bd      <- suppressWarnings(as.numeric(bead_df_bd[[col]]))
                    pass_rows_bd <- which(!is.na(vals_bd) & vals_bd >= bead_thresh_bd)
                    fail_rows_bd <- which(!is.na(vals_bd) & vals_bd < bead_thresh_bd)
                    if (length(pass_rows_bd) > 0)
                      openxlsx::addStyle(wb, "09_bead_information", pass_style_bd,
                                         rows = pass_rows_bd + 1L, cols = col_idx_bd, stack = TRUE)
                    if (length(fail_rows_bd) > 0)
                      openxlsx::addStyle(wb, "09_bead_information", fail_style_bd,
                                         rows = fail_rows_bd + 1L, cols = col_idx_bd, stack = TRUE)
                  }
                }

                # Flag % Agg columns: highlight values BELOW the configurable
                # threshold (same input$pb_agg_threshold used on-screen)
                thresh_bd <- suppressWarnings(as.numeric(input$pb_agg_threshold))
                if (length(thresh_bd) == 0 || is.na(thresh_bd) || !is.finite(thresh_bd))
                  thresh_bd <- 1e9
                flag_style_bd <- openxlsx::createStyle(fontColour = "#856404",
                                   fgFill = "#fff3cd", textDecoration = "bold")
                for (col in agg_disp_present_bd) {
                  col_idx_bd   <- which(colnames(bead_df_bd) == col)
                  vals_bd      <- suppressWarnings(as.numeric(bead_df_bd[[col]]))
                  flag_rows_bd <- which(!is.na(vals_bd) & vals_bd < thresh_bd)
                  if (length(flag_rows_bd) > 0)
                    openxlsx::addStyle(wb, "09_bead_information", flag_style_bd,
                                       rows = flag_rows_bd + 1L, cols = col_idx_bd, stack = TRUE)
                }
              }
            }
          }
        }, error = function(e) {
          openxlsx::writeData(wb, "09_bead_information",
            data.frame(Note = paste("Could not generate bead information table:",
                                    .clean_msg(conditionMessage(e)))))
        })
      } # end Point-based (09_bead_information)

      # -- SHEET 10: control_plots (Point-based only) -------------------------
      # One PNG per antigen, stacked vertically in a single sheet.
      if ("Point-based" %in% bama_sel) {
        openxlsx::addWorksheet(wb, "10_control_plots")
        tryCatch({
          d_ctrl <- tit_ctrl_data()
          if (is.null(d_ctrl) || nrow(d_ctrl) == 0) {
            openxlsx::writeData(wb, "10_control_plots",
              data.frame(Note = "No control/QC data available. Configure Point-based QC."))
          } else {
            analytes_ctrl <- sort(unique(as.character(d_ctrl$analyte)))
            # Reconstruct per-antigen version of the control plot used in the UI.
            # We reuse the same styling/LOD logic from tit_ctrl_plot_reactive but
            # subset to one analyte at a time so each antigen gets its own page.

            # Collect LOD per analyte (mirrors tit_ctrl_plot_reactive)
            ag_cfg_display_lod <- list()
            for (raw_key in names(rv$ag_config)) {
              dk <- trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", raw_key))
              ag_cfg_display_lod[[dk]] <- rv$ag_config[[raw_key]]
            }
            global_lod <- suppressWarnings(as.numeric(input$pb_lod)) %||% 2674
            global_lod <- if (is.na(global_lod)) 2674 else global_lod

            lod_for <- function(ag) {
              cfg     <- ag_cfg_display_lod[[ag]]
              lod_val <- suppressWarnings(as.numeric(cfg$lod))
              if (!is.null(cfg) && length(lod_val) == 1L && !is.na(lod_val) && lod_val > 0)
                list(lod = lod_val, locked = isTRUE(cfg$lod_locked))
              else
                list(lod = global_lod, locked = FALSE)
            }

            cat_colours <- c(
              "Positive\nControl"     = "#e74c3c",
              "Negative\nControl"     = "#6c6cdb",
              "Unclassified\nControl" = "#95a5a6",
              "Sample"               = "#555555"
            )
            cat_shapes <- c(
              "Positive\nControl"     = 16L,
              "Negative\nControl"     = 16L,
              "Unclassified\nControl" = 1L,
              "Sample"               = 16L
            )

            img_row    <- 1L
            img_height <- 4.5   # inches per antigen panel
            img_width  <- 8

            for (ag in analytes_ctrl) {
              tryCatch({
              d_ag <- d_ctrl[d_ctrl$analyte == ag, , drop = FALSE]
              if (nrow(d_ag) == 0) next

              d_ag$x_group <- dplyr::case_when(
                d_ag$ctrl_category == "positive"     ~ "Positive\nControl",
                d_ag$ctrl_category == "negative"     ~ "Negative\nControl",
                d_ag$ctrl_category == "unclassified" ~ "Unclassified\nControl",
                TRUE                                  ~ "Sample"
              )
              group_order    <- c("Positive\nControl", "Negative\nControl",
                                  "Unclassified\nControl", "Sample")
              present_groups <- intersect(group_order, unique(d_ag$x_group))
              d_ag$x_group   <- factor(d_ag$x_group, levels = present_groups)

              jitter_pos <- ggplot2::position_jitter(width = 0.18, height = 0, seed = 42)

              p_ag <- ggplot2::ggplot(d_ag,
                          ggplot2::aes(x = x_group, y = MFI,
                                       colour = x_group, shape = x_group,
                                       label = Sample_ID)) +
                ggplot2::geom_point(position = jitter_pos, size = 3.2, alpha = 0.88) +
                ggrepel::geom_text_repel(
                  position = jitter_pos, size = 2.6, max.overlaps = 20,
                  box.padding = 0.25, segment.size = 0.3,
                  segment.colour = "#aaaaaa", show.legend = FALSE) +
                ggplot2::scale_colour_manual(
                  values = cat_colours[present_groups], name = NULL, drop = TRUE) +
                ggplot2::scale_shape_manual(
                  values = cat_shapes[present_groups],  name = NULL, drop = TRUE) +
                ggplot2::scale_x_discrete(drop = TRUE) +
                ggplot2::scale_y_continuous(
                  labels = scales::label_comma(accuracy = 1),
                  expand = ggplot2::expansion(mult = c(0.05, 0.20))) +
                ggplot2::labs(x = NULL, y = "MFI",
                              title = paste("QC Plot \u2014", ag)) +
                ggplot2::theme_minimal(base_size = 12) +
                ggplot2::theme(
                  axis.text.x        = ggplot2::element_text(size = 11, face = "bold"),
                  axis.text.y        = ggplot2::element_text(size = 9),
                  panel.grid.major.x = ggplot2::element_blank(),
                  panel.grid.minor   = ggplot2::element_blank(),
                  axis.title.y       = ggplot2::element_text(size = 11, face = "bold"),
                  legend.position    = "right",
                  legend.text        = ggplot2::element_text(size = 9),
                  panel.border       = ggplot2::element_rect(colour = "#cccccc",
                                         fill = NA, linewidth = 0.4),
                  plot.title         = ggplot2::element_text(face = "bold", size = 13,
                                         colour = "#1a3a5c", hjust = 0.5)
                )

              # LOD line
              lod_info <- lod_for(ag)
              if (isTRUE(!is.na(lod_info$lod) && lod_info$lod > 0)) {
                lod_colour   <- if (isTRUE(lod_info$locked)) "#c0392b" else "#888888"
                lod_linetype <- "dashed"
                p_ag <- p_ag +
                  ggplot2::geom_hline(
                    yintercept = lod_info$lod,
                    linetype   = lod_linetype, colour = lod_colour,
                    linewidth  = 0.75, inherit.aes = FALSE) +
                  ggplot2::annotate(
                    "text", x = -Inf, y = lod_info$lod,
                    label   = paste0("LOD = ", round(lod_info$lod, 1),
                                     if (isTRUE(lod_info$locked)) " \u25cf" else ""),
                    hjust = -0.08, vjust = -0.45,
                    size  = 2.9, colour = lod_colour,
                    fontface = if (isTRUE(lod_info$locked)) "bold.italic" else "italic",
                    inherit.aes = FALSE)
              }

              .embed_plot(wb, "10_control_plots", p_ag,
                          width_in  = img_width,
                          height_in = img_height,
                          start_row = img_row,
                          start_col = 1)
              # Advance row pointer: openxlsx row unit ~ 20px; 1 in ~ 4 rows
              img_row <- img_row + ceiling(img_height * 6) + 1L
              }, error = function(e) {
                # Don't let one antigen's failure blank out the whole sheet --
                # log a note in its place and keep going with the rest.
                openxlsx::writeData(wb, "10_control_plots",
                  data.frame(Note = paste0("Could not generate control plot for '",
                                           ag, "': ", .clean_msg(conditionMessage(e)))),
                  startRow = img_row)
                img_row <<- img_row + 2L
              })
            }
          }
        }, error = function(e) {
          openxlsx::writeData(wb, "10_control_plots",
            data.frame(Note = paste("Could not generate control plots:", conditionMessage(e))))
        })
      } # end Point-based

      # ======================================================================
      # SHEETS 8-11: Titration only
      # ======================================================================
      if ("Titration" %in% bama_sel) {

        # -- SHEET 11: titration_data ------------------------------------------
        tryCatch({
          d <- tit_base_data()
          if (!is.null(d) && nrow(d) > 0) {
            d_out <- d[is.na(d$sample_type) |
                       tolower(trimws(d$sample_type)) != "standard_curve", , drop = FALSE]

            # -- perc_binding: (MFI - lower_4pl) / (upper_4pl - lower_4pl) * 100
            # Fit 4PL per (analyte, base_sample_id) on averaged MFI to get the
            # lower/upper asymptotes, then normalise EACH raw MFI row to 0-100%.
            # Values are clamped to [0, 100] -- the 4PL asymptotes define the scale.
            d_out$perc_binding <- NA_real_
            if ("base_sample_id" %in% colnames(d_out)) {
              d_samp_pb <- d[!is.na(d$x_value) & d$x_value > 0 &
                             !is.na(d$sample_type) &
                             tolower(trimws(d$sample_type)) == "sample" &
                             !is.na(d$base_sample_id), , drop = FALSE]
              if (nrow(d_samp_pb) > 0) {
                agg_pb <- as.data.frame(
                  d_samp_pb %>%
                    dplyr::group_by(analyte, base_sample_id, x_value) %>%
                    dplyr::summarise(avg_MFI = mean(MFI, na.rm = TRUE), .groups = "drop")
                )
                keys_pb <- unique(paste(agg_pb$analyte, agg_pb$base_sample_id, sep = "\x01"))
                pb_df <- do.call(rbind, lapply(keys_pb, function(key) {
                  parts <- strsplit(key, "\x01", fixed = TRUE)[[1]]
                  ag <- parts[1]; bid <- parts[2]
                  sub <- agg_pb[agg_pb$analyte == ag & agg_pb$base_sample_id == bid, ]
                  params <- tryCatch({
                    withCallingHandlers(
                      suppressWarnings({
                        fit <- drc::drm(avg_MFI ~ x_value, data = sub,
                                        fct = drc::LL.4(names = c("Slope","Lower","Upper","EC50")))
                        cf <- coef(fit)
                        list(
                          lower = unname(cf[grep("Lower", names(cf), ignore.case = TRUE)[1]]),
                          upper = unname(cf[grep("Upper", names(cf), ignore.case = TRUE)[1]])
                        )
                      }),
                      message = function(m) invokeRestart("muffleMessage")
                    )
                  }, error = function(e) NULL)
                  # Fallback to observed min/max if model fails
                  if (is.null(params) || any(!is.finite(c(params$lower, params$upper))) ||
                      abs(params$upper - params$lower) < 1e-6) {
                    params <- list(lower = min(sub$avg_MFI, na.rm = TRUE),
                                   upper = max(sub$avg_MFI, na.rm = TRUE))
                  }
                  data.frame(key = key, lower = params$lower, upper = params$upper,
                             stringsAsFactors = FALSE)
                }))
                out_key <- paste(d_out$analyte, d_out$base_sample_id, sep = "\x01")
                matched <- match(out_key, pb_df$key)
                lower_v <- pb_df$lower[matched]
                upper_v <- pb_df$upper[matched]
                pb_raw  <- (d_out$MFI - lower_v) / (upper_v - lower_v) * 100
                # Clamp to [0, 100] -- individual wells can sit outside asymptotes
                pb_raw  <- pmax(0, pmin(100, pb_raw))
                d_out$perc_binding <- ifelse(
                  !is.na(lower_v) & !is.na(upper_v) & abs(upper_v - lower_v) > 1e-6,
                  round(pb_raw, 2),
                  NA_real_
                )
              }
            }
            d_out$perc_binding[!is.na(d_out$perc_binding) &
                               !is.finite(d_out$perc_binding)] <- NA_real_


            has_replicates <- FALSE
            if ("base_sample_id" %in% colnames(d_out)) {
              type_counts <- tapply(
                seq_len(nrow(d_out)),
                paste(d_out$analyte, d_out$base_sample_id, d_out$Type, sep = "\x01"),
                length
              )
              has_replicates <- any(type_counts > 1, na.rm = TRUE)
            }

            if (has_replicates && "base_sample_id" %in% colnames(d_out)) {
              d_out <- d_out[order(d_out$analyte, d_out$base_sample_id,
                                   d_out$Type, d_out$Well), , drop = FALSE]
              grp_key <- paste(d_out$analyte, d_out$base_sample_id,
                               d_out$Type, sep = "\x01")
              d_out$.rep_num <- ave(seq_len(nrow(d_out)), grp_key, FUN = seq_along)
              pivot_key <- paste(d_out$analyte, d_out$base_sample_id,
                                 d_out$Type, sep = "\x01")
              uniq_keys <- unique(pivot_key)
              meta_take <- intersect(
                c("base_sample_id", "analyte", "Type", "x_value",
                  "sample_type", "sample_kind", "perc_binding"),
                colnames(d_out)
              )
              wide_rows <- lapply(uniq_keys, function(k) {
                rows <- d_out[pivot_key == k, , drop = FALSE]
                out  <- rows[1, meta_take, drop = FALSE]
                out$Sample_ID <- paste(sort(unique(rows$Sample_ID)), collapse = "; ")
                r1 <- rows$MFI[rows$.rep_num == 1]
                r2 <- rows$MFI[rows$.rep_num == 2]
                out$rep1_MFI <- if (length(r1) > 0) r1[1] else NA_real_
                out$rep2_MFI <- if (length(r2) > 0) r2[1] else NA_real_
                out$mean_MFI <- mean(rows$MFI, na.rm = TRUE)
                # per_std_dev: SD of the two perc_binding values (one per replicate)
                pb1 <- if (length(r1) > 0) rows$perc_binding[rows$.rep_num == 1][1] else NA_real_
                pb2 <- if (length(r2) > 0) rows$perc_binding[rows$.rep_num == 2][1] else NA_real_
                out$per_std_dev <- if (!is.na(pb1) && !is.na(pb2))
                                     paste0(round(sd(c(pb1, pb2), na.rm = TRUE), 2), "%")
                                   else NA_character_
                # Add % suffix to perc_binding (take from rows[1] after pivot)
                pb_val <- rows$perc_binding[1]
                out$perc_binding <- if (!is.na(pb_val)) paste0(pb_val, "%") else NA_character_
                out
                out
              })
              wide <- do.call(rbind, wide_rows)
              wide$base_sample_id <- NULL
              wide$.rep_num        <- NULL
              # Rename x_value -> dilution
              names(wide)[names(wide) == "x_value"] <- "dilution"
              # Final column order
              final_cols <- intersect(
                c("analyte", "Sample_ID", "Type", "sample_type", "sample_kind",
                  "dilution", "rep1_MFI", "rep2_MFI", "mean_MFI",
                  "per_std_dev"),
                colnames(wide)
              )
              wide <- wide[, final_cols, drop = FALSE]
              wide <- wide[order(wide$analyte, wide$Sample_ID), , drop = FALSE]
              rownames(wide) <- NULL
              .add_flat(wb, "11_titration_data", wide)
            } else {
              # Non-replicate path: keep only specified columns, rename x_value
              keep <- intersect(
                c("analyte", "Sample_ID", "Type", "sample_type", "sample_kind",
                  "x_value", "MFI"),
                colnames(d_out)
              )
              d_final <- d_out[, keep, drop = FALSE]
              names(d_final)[names(d_final) == "x_value"] <- "dilution"
              # per_std_dev not applicable (no replicates) -- add as NA column
              d_final$rep1_MFI    <- NA_real_
              d_final$rep2_MFI    <- NA_real_
              d_final$mean_MFI    <- d_final$MFI
              d_final$per_std_dev <- NA_character_
              d_final$MFI         <- NULL
              final_cols_nrep <- intersect(
                c("analyte", "Sample_ID", "Type", "sample_type", "sample_kind",
                  "dilution", "rep1_MFI", "rep2_MFI", "mean_MFI",
                  "per_std_dev"),
                colnames(d_final)
              )
              d_final <- d_final[, final_cols_nrep, drop = FALSE]
              d_final <- d_final[order(d_final$analyte, d_final$Sample_ID), , drop = FALSE]
              rownames(d_final) <- NULL
              .add_flat(wb, "11_titration_data", d_final)
            }
          } else {
            openxlsx::addWorksheet(wb, "11_titration_data")
            openxlsx::writeData(wb, "11_titration_data",
              data.frame(Note = "No titration data available. Run a Titration analysis first."))
          }
        }, error = function(e) {
          openxlsx::addWorksheet(wb, "11_titration_data")
          openxlsx::writeData(wb, "11_titration_data",
            data.frame(Note = paste("Could not generate titration data:",
                                    .clean_msg(conditionMessage(e)))))
        })
        # -- SHEET 12: titration_curves ------------------------------------------
        openxlsx::addWorksheet(wb, "12_titration_curves")
        tryCatch({
          d_tit <- tit_base_data()
          if (is.null(d_tit) || nrow(d_tit) == 0) {
            openxlsx::writeData(wb, "12_titration_curves",
              data.frame(Note = "No titration data available."))
          } else {
            d_tit <- d_tit[
              (is.na(d_tit$sample_type) |
               tolower(trimws(d_tit$sample_type)) != "standard_curve") &
              !is.na(d_tit$x_value) & d_tit$x_value > 0, , drop = FALSE]
            d_tit <- as.data.frame(
              d_tit %>%
                dplyr::group_by(base_sample_id, analyte, Type, x_value,
                                sample_type, sample_kind) %>%
                dplyr::summarise(avg_MFI = mean(MFI, na.rm = TRUE),
                                 n_reps  = dplyr::n(), .groups = "drop") %>%
                dplyr::arrange(base_sample_id, analyte, x_value)
            )
            if (nrow(d_tit) == 0) {
              openxlsx::writeData(wb, "12_titration_curves",
                data.frame(Note = "No concentration/dilution data to plot."))
            } else {
              analytes_tit <- sort(unique(as.character(d_tit$analyte)))
              kinds        <- tolower(trimws(unique(
                d_tit$sample_kind[!is.na(d_tit$sample_kind)])))
              x_label <- if (length(kinds) == 1 && kinds == "mab")
                           "Concentration (\u00b5g/mL)"
                         else if (length(kinds) == 1 && kinds == "serum")
                           "Dilution"
                         else "Concentration / Dilution"
              img_row <- 1L; img_h <- 5; img_w <- 9
              for (ag in analytes_tit) {
                d_ag <- d_tit[d_tit$analyte == ag, , drop = FALSE]
                if (nrow(d_ag) == 0) next
                colour_grps <- sort(unique(d_ag$base_sample_id))
                pal         <- setNames(.make_pal(length(colour_grps)), colour_grps)
                lod_df_ag   <- .tit_lod_lookup(ag)
                lod_val_ag  <- suppressWarnings(as.numeric(lod_df_ag$lod[1]))
                p_tit <- ggplot2::ggplot(d_ag,
                  ggplot2::aes(x = x_value, y = avg_MFI,
                               colour = base_sample_id, group = base_sample_id)) +
                  ggplot2::geom_line(linewidth = 0.55, alpha = 0.55) +
                  ggplot2::geom_smooth(method = "loess", se = FALSE, span = 0.75,
                                       linewidth = 1.1, alpha = 0.9) +
                  ggplot2::geom_point(size = 3, alpha = 0.9) +
                  ggplot2::scale_x_log10(labels = scales::label_comma(accuracy = 0.001)) +
                  ggplot2::scale_y_continuous(labels = scales::label_comma(accuracy = 1)) +
                  ggplot2::scale_colour_manual(values = pal, name = "Sample") +
                  ggplot2::labs(x = x_label, y = "Average MFI",
                                title = {
                                  ag_kinds <- tolower(trimws(unique(
                                    d_ag$sample_kind[!is.na(d_ag$sample_kind)])))
                                  curve_type <- if (length(ag_kinds) == 1 && ag_kinds == "mab")
                                                  "mAb titration curve"
                                                else if (length(ag_kinds) == 1 && ag_kinds == "serum")
                                                  "Serum titration curve"
                                                else "Titration curve"
                                  paste(curve_type, "\u2014", ag)
                                }) +
                  ggplot2::theme_minimal(base_size = 12) +
                  ggplot2::theme(
                    legend.position  = "right",
                    legend.text      = ggplot2::element_text(size = 9),
                    legend.title     = ggplot2::element_text(size = 10, face = "bold"),
                    panel.grid.minor = ggplot2::element_blank(),
                    axis.title       = ggplot2::element_text(size = 11, face = "bold"),
                    plot.title       = ggplot2::element_text(face = "bold", size = 13,
                                         colour = "#1a3a5c", hjust = 0.5),
                    plot.margin      = ggplot2::margin(10, 10, 10, 10)
                  )
                # -- Per-antigen LOD -- dotted reference line (BRILLIANT BioPlex /
                # INTELLIFLEX with a saved LOD only; forced to 0 otherwise) --------
                if (!is.na(lod_val_ag) && lod_val_ag > 0) {
                  p_tit <- p_tit +
                    ggplot2::geom_hline(yintercept = lod_val_ag,
                                        linetype = "dotted", colour = "gray30",
                                        linewidth = 0.7) +
                    ggplot2::annotate("text", x = -Inf, y = lod_val_ag,
                                      label = paste0("LOD = ", round(lod_val_ag, 1)),
                                      hjust = -0.05, vjust = -0.4, size = 3,
                                      colour = "gray30", fontface = "italic",
                                      inherit.aes = FALSE)
                }
                .embed_plot(wb, "12_titration_curves", p_tit,
                            width_in = img_w, height_in = img_h,
                            start_row = img_row, start_col = 1)
                img_row <- img_row + ceiling(img_h * 6) + 1L
              }
            }
          }
        }, error = function(e) {
          openxlsx::writeData(wb, "12_titration_curves",
            data.frame(Note = paste("Could not generate titration curves:",
                                    .clean_msg(conditionMessage(e)))))
        })

        # -- SHEET 13: auc_summary ---------------------------------------------
        tryCatch({
          d_auc <- auc_plot_data()
          if (!is.null(d_auc) && nrow(d_auc) > 0) {
            kinds  <- tolower(trimws(unique(d_auc$sample_kind[!is.na(d_auc$sample_kind)])))
            x_lbl  <- if (length(kinds) == 1 && kinds == "mab")   "conc_min_ug_mL" else
                      if (length(kinds) == 1 && kinds == "serum") "dil_min" else "x_min"
            x_lbl2 <- if (length(kinds) == 1 && kinds == "mab")   "conc_max_ug_mL" else
                      if (length(kinds) == 1 && kinds == "serum") "dil_max" else "x_max"
            auc_out <- data.frame(
              analyte   = d_auc$analyte,
              sample_id = d_auc$Sample_ID,
              AUC       = round(d_auc$AUC, 4),
              n_points  = d_auc$n_points,
              stringsAsFactors = FALSE
            )
            auc_out[[x_lbl]]  <- round(d_auc$x_min, 5)
            auc_out[[x_lbl2]] <- round(d_auc$x_max, 3)
            auc_out <- auc_out[order(auc_out$analyte, -auc_out$AUC), ]
            .add_flat(wb, "13_auc_summary", auc_out)
          } else {
            openxlsx::addWorksheet(wb, "13_auc_summary")
            openxlsx::writeData(wb, "13_auc_summary",
              data.frame(Note = "No AUC data available. Run a Titration analysis first."))
          }
        }, error = function(e) {
          openxlsx::addWorksheet(wb, "13_auc_summary")
          openxlsx::writeData(wb, "13_auc_summary",
            data.frame(Note = paste("Could not generate AUC data:",
                                    .clean_msg(conditionMessage(e)))))
        })

        # -- SHEET 14: auc_plots (all antigens stacked on one sheet) ----------
        openxlsx::addWorksheet(wb, "14_auc_plots")
        tryCatch({
          d_auc_plots <- auc_plot_data()
          if (is.null(d_auc_plots) || nrow(d_auc_plots) == 0) {
            openxlsx::writeData(wb, "14_auc_plots",
              data.frame(Note = "No AUC data available. Run a Titration analysis first."))
          } else {
            analytes_auc <- sort(unique(as.character(d_auc_plots$analyte)))
            img_row_auc  <- 1L; img_h_auc <- 5; img_w_auc <- 9

            for (ag_auc in analytes_auc) {
              d_ag_auc <- d_auc_plots[d_auc_plots$analyte == ag_auc, , drop = FALSE]
              if (nrow(d_ag_auc) == 0) next

              colour_grps_auc <- sort(unique(as.character(d_ag_auc$Sample_ID)))
              pal_auc         <- setNames(.make_pal(length(colour_grps_auc)),
                                         colour_grps_auc)

              p_auc <- ggplot2::ggplot(
                d_ag_auc,
                ggplot2::aes(x = reorder(Sample_ID, -AUC), y = AUC, fill = Sample_ID)) +
                ggplot2::geom_col(alpha = 0.85, width = 0.65) +
                ggplot2::scale_fill_manual(values = pal_auc) +
                ggplot2::scale_y_continuous(
                  labels = scales::label_comma(accuracy = 0.01),
                  expand = ggplot2::expansion(mult = c(0, 0.15))) +
                ggplot2::labs(
                  x     = "Sample",
                  y     = "AUC (trapezoidal)",
                  title = paste("AUC \u2014", ag_auc),
                  fill  = "Sample") +
                ggplot2::theme_minimal(base_size = 12) +
                ggplot2::theme(
                  axis.text.x        = ggplot2::element_text(angle = 35, hjust = 1, size = 9),
                  panel.grid.major.x = ggplot2::element_blank(),
                  panel.grid.minor   = ggplot2::element_blank(),
                  legend.position    = "none",
                  axis.title         = ggplot2::element_text(size = 11, face = "bold"),
                  plot.title         = ggplot2::element_text(face = "bold", size = 13,
                                         colour = "#1a3a5c", hjust = 0.5),
                  plot.margin        = ggplot2::margin(10, 10, 10, 10)
                )
              .embed_plot(wb, "14_auc_plots", p_auc,
                          width_in = img_w_auc, height_in = img_h_auc,
                          start_row = img_row_auc, start_col = 1)
              img_row_auc <- img_row_auc + ceiling(img_h_auc * 6) + 1L
            }
          }
        }, error = function(e) {
          openxlsx::writeData(wb, "14_auc_plots",
            data.frame(Note = paste("Could not generate AUC plots:",
                                    .clean_msg(conditionMessage(e)))))
        })

      } # end Titration

      # ======================================================================
      # SHEET 15: standard_curve (Quantification only)
      # ======================================================================
      if ("Quantification" %in% bama_sel) {
        openxlsx::addWorksheet(wb, "15_standard_curve")
        tryCatch({
          d_sc <- sc_base_data()
          if (is.null(d_sc) || nrow(d_sc) == 0) {
            openxlsx::writeData(wb, "15_standard_curve",
              data.frame(Note = paste(
                "No standard curve data found.",
                "Ensure the helper file has sample_type = 'standard_curve'",
                "and start_concentration filled for those wells."
              )))
          } else {
            # -- Flat data table ---------------------------------------------
            keep_sc <- intersect(
              c("analyte", "base_sample_id", "Well", "Type", "x_value",
                "dilution_label", "MFI", "sample_kind"),
              colnames(d_sc)
            )
            d_sc_out <- d_sc[, keep_sc, drop = FALSE]
            colnames(d_sc_out)[colnames(d_sc_out) == "base_sample_id"] <- "standard_curve_id"
            d_sc_out <- d_sc_out[order(d_sc_out$analyte,
                                       d_sc_out$standard_curve_id,
                                       d_sc_out$x_value), , drop = FALSE]
            rownames(d_sc_out) <- NULL
            openxlsx::writeData(wb, "15_standard_curve", d_sc_out)

            # -- Dashboard-matching plots: one per analyte -------------------
            analytes_sc <- sort(unique(as.character(d_sc$analyte)))

            # Respect the UI fit-range and log_x settings at export time
            log_x_exp   <- !is.null(input$sc_log_x) && input$sc_log_x == "log"
            fit_range   <- input$sc_fit_log_range
            use_dil_exp <- !is.null(input$sc_x_label_type) &&
                           input$sc_x_label_type == "dilution"

            kinds_sc    <- tolower(trimws(unique(
              d_sc$sample_kind[!is.na(d_sc$sample_kind)])))
            x_lbl_sc <- if (use_dil_exp)
                           "Dilution"
                         else if (log_x_exp)
                           "Known Concentration (\u00b5g/mL, log\u2081\u2080 scale)"
                         else if (length(kinds_sc) == 1 && kinds_sc == "mab")
                           "Concentration (\u00b5g/mL)"
                         else if (length(kinds_sc) == 1 && kinds_sc == "serum")
                           "Dilution"
                         else "Concentration / Dilution"

            # Get sample concentration estimates for ALL analytes for overlay
            est_tbl_exp <- tryCatch(sc_sample_conc_all(), error = function(e) NULL)

            img_row_sc <- nrow(d_sc_out) + 3L
            img_h_sc   <- 5.5; img_w_sc <- 9

            for (ag in analytes_sc) {
              d_ag_sc <- d_sc[d_sc$analyte == ag & !is.na(d_sc$x_value) &
                              d_sc$x_value > 0, , drop = FALSE]
              if (nrow(d_ag_sc) == 0) next

              d_avg_sc <- as.data.frame(
                d_ag_sc %>%
                  dplyr::group_by(base_sample_id, analyte, x_value) %>%
                  dplyr::summarise(avg_MFI = mean(MFI, na.rm = TRUE),
                                   n_reps  = dplyr::n(),
                                   .groups = "drop")
              )

              # Apply fit range to determine in/out subsets
              d_fit_exp <- if (!is.null(fit_range) && length(fit_range) == 2) {
                lo <- fit_range[1]; hi <- fit_range[2]
                d_avg_sc[log10(d_avg_sc$x_value) >= lo - 1e-9 &
                         log10(d_avg_sc$x_value) <= hi + 1e-9, , drop = FALSE]
              } else d_avg_sc
              fit_active_exp <- nrow(d_fit_exp) < nrow(d_avg_sc)

              d_avg_sc$in_fit_range <- if (fit_active_exp) {
                lo <- fit_range[1]; hi <- fit_range[2]
                log10(d_avg_sc$x_value) >= lo - 1e-9 &
                  log10(d_avg_sc$x_value) <= hi + 1e-9
              } else rep(TRUE, nrow(d_avg_sc))

              sc_grps <- sort(unique(d_avg_sc$base_sample_id))
              pal_sc  <- setNames(.make_pal(length(sc_grps)), sc_grps)

              # Build dil_map for secondary axis labels
              dil_map_exp <- local({
                idx <- !is.na(d_ag_sc$dilution_label) & !is.na(d_ag_sc$x_value)
                if (any(idx)) {
                  pairs <- unique(d_ag_sc[idx, c("x_value", "dilution_label"),
                                          drop = FALSE])
                  setNames(pairs$dilution_label, as.character(pairs$x_value))
                } else {
                  xv <- sort(unique(d_ag_sc$x_value), decreasing = TRUE)
                  setNames(as.character(xv), as.character(xv))
                }
              })

              p_sc <- ggplot2::ggplot(d_avg_sc,
                ggplot2::aes(x      = x_value,
                             y      = avg_MFI,
                             colour = base_sample_id,
                             fill   = base_sample_id,
                             group  = base_sample_id)) +
                ggplot2::geom_line(linewidth = 0.55, linetype = "dashed", alpha = 0.55) +
                ggplot2::geom_point(
                  data  = d_avg_sc[d_avg_sc$in_fit_range, , drop = FALSE],
                  ggplot2::aes(size = n_reps), alpha = 0.92, shape = 16) +
                {if (fit_active_exp && any(!d_avg_sc$in_fit_range))
                  ggplot2::geom_point(
                    data  = d_avg_sc[!d_avg_sc$in_fit_range, , drop = FALSE],
                    ggplot2::aes(size = n_reps), alpha = 0.35, shape = 1, stroke = 1.2)
                } +
                ggplot2::geom_smooth(
                  data   = d_fit_exp,
                  method = "lm", se = FALSE, linewidth = 1.1,
                  formula = if (log_x_exp) y ~ log10(x) else y ~ x
                ) +
                ggplot2::scale_colour_manual(values = pal_sc, guide = "none") +
                ggplot2::scale_fill_manual(  values = pal_sc, guide = "none") +
                ggplot2::scale_size_continuous(range = c(2.5, 5.5), guide = "none") +
                ggplot2::scale_y_continuous(
                  labels = scales::label_comma(accuracy = 1),
                  expand = ggplot2::expansion(mult = c(0.05, 0.22))) +
                ggplot2::labs(x = x_lbl_sc, y = "Average MFIs",
                              title = paste("Standard curve \u2014", ag)) +
                ggplot2::theme_minimal(base_size = 12) +
                ggplot2::theme(
                  legend.position   = "none",
                  panel.grid.minor  = ggplot2::element_blank(),
                  axis.title        = ggplot2::element_text(size = 11, face = "bold"),
                  plot.title        = ggplot2::element_text(face = "bold", size = 13,
                                       colour = "#1a3a5c", hjust = 0.5),
                  strip.text        = ggplot2::element_text(face = "bold", size = 11,
                                       colour = "#1a3a5c"),
                  strip.background  = ggplot2::element_rect(fill = "#eef3f9",
                                       colour = NA),
                  plot.margin       = ggplot2::margin(10, 16, 10, 10)
                )

              # x-axis scale with secondary dilution axis
              x_breaks_exp <- sort(unique(d_avg_sc$x_value))
              x_dil_exp    <- dil_map_exp[as.character(x_breaks_exp)]
              if (log_x_exp)
                p_sc <- p_sc + ggplot2::scale_x_log10(
                  labels   = scales::label_comma(accuracy = 0.001),
                  sec.axis = ggplot2::sec_axis(~ ., breaks = x_breaks_exp,
                                               labels = x_dil_exp,
                                               name   = "Dilution"))
              else
                p_sc <- p_sc + ggplot2::scale_x_continuous(
                  labels   = scales::label_comma(accuracy = 0.001),
                  sec.axis = ggplot2::sec_axis(~ ., breaks = x_breaks_exp,
                                               labels = x_dil_exp,
                                               name   = "Dilution"))

              # Sample estimate overlay (grey dots + labels)
              if (!is.null(est_tbl_exp) && nrow(est_tbl_exp) > 0 && !use_dil_exp) {
                est_ag <- est_tbl_exp[est_tbl_exp$analyte == ag &
                                      !is.na(est_tbl_exp$est_conc_ug_mL) &
                                      est_tbl_exp$est_conc_ug_mL > 0 &
                                      est_tbl_exp$status == "OK", , drop = FALSE]
                if (nrow(est_ag) > 0) {
                  p_sc <- p_sc +
                    ggplot2::geom_point(
                      data        = est_ag,
                      ggplot2::aes(x = est_conc_ug_mL, y = avg_MFI),
                      inherit.aes = FALSE,
                      shape = 21, size = 3.5,
                      colour = "grey25", fill = "grey70", alpha = 0.90) +
                    ggplot2::geom_label(
                      data        = est_ag,
                      ggplot2::aes(x = est_conc_ug_mL, y = avg_MFI,
                                   label = paste0(sample_id, "\n",
                                                  formatC(est_conc_ug_mL,
                                                          format = "fg", digits = 3),
                                                  " \u00b5g/mL")),
                      inherit.aes   = FALSE,
                      size          = 2.8, colour = "grey20", fill = "white",
                      alpha         = 0.85, label.padding = ggplot2::unit(0.18, "lines"),
                      label.size    = 0.25, hjust = -0.10, vjust = 0.5,
                      show.legend   = FALSE)
                }
              }

              # R² annotation per sample group
              if (nrow(d_fit_exp) >= 2) {
                r2_rows_exp <- lapply(sc_grps, function(sid) {
                  sub <- d_fit_exp[d_fit_exp$base_sample_id == sid, , drop = FALSE]
                  if (nrow(sub) < 2) return(NULL)
                  fit_lm <- tryCatch(
                    if (log_x_exp) lm(avg_MFI ~ log10(x_value), data = sub)
                    else           lm(avg_MFI ~ x_value,         data = sub),
                    error = function(e) NULL)
                  if (is.null(fit_lm)) return(NULL)
                  r2 <- summary(fit_lm)$r.squared
                  data.frame(base_sample_id = sid,
                             label = sprintf("R\u00b2 = %.4f", r2),
                             stringsAsFactors = FALSE)
                })
                r2_df_exp <- do.call(rbind, Filter(Negate(is.null), r2_rows_exp))
                if (!is.null(r2_df_exp) && nrow(r2_df_exp) > 0) {
                  p_sc <- p_sc +
                    ggplot2::geom_label(
                      data        = r2_df_exp,
                      ggplot2::aes(label = label),
                      x = -Inf, y = Inf,
                      hjust = -0.08, vjust = 1.3,
                      inherit.aes = FALSE,
                      size = 3.2, colour = "#1a3a5c", fill = "#eef3f9",
                      label.size = 0.25, alpha = 0.88)
                }
              }

              # Facet when multiple standard curve samples
              if (length(sc_grps) > 1)
                p_sc <- p_sc + ggplot2::facet_wrap(
                  ~ base_sample_id, scales = "free_y", ncol = 2)

              .embed_plot(wb, "15_standard_curve", p_sc,
                          width_in = img_w_sc, height_in = img_h_sc,
                          start_row = img_row_sc, start_col = 1)
              img_row_sc <- img_row_sc + ceiling(img_h_sc * 6) + 1L
            }
          }
        }, error = function(e) {
          openxlsx::writeData(wb, "15_standard_curve",
            data.frame(Note = paste("Could not generate standard curve:",
                                    .clean_msg(conditionMessage(e)))))
        })

        # -- SHEET 16: sample_concentration (all antigens combined) -----------
        tryCatch({
          conc_tbl <- sc_sample_conc_all()
          if (!is.null(conc_tbl) && nrow(conc_tbl) > 0) {
            conc_out <- conc_tbl
            colnames(conc_out) <- c("Analyte", "Sample ID", "x Value (Conc/Dil)",
                                    "Std Curve Sample", "Avg MFI", "N Reps",
                                    "Est. Conc. (ug/mL) [linear]",
                                    "Est. Conc. [log10]",
                                    "Status")
            .add_flat(wb, "16_sample_concentration", conc_out)
          } else {
            openxlsx::addWorksheet(wb, "16_sample_concentration")
            openxlsx::writeData(wb, "16_sample_concentration",
              data.frame(Note = "No sample concentration estimates available."))
          }
        }, error = function(e) {
          openxlsx::addWorksheet(wb, "16_sample_concentration")
          openxlsx::writeData(wb, "16_sample_concentration",
            data.frame(Note = paste("Could not generate concentration estimates:",
                                    .clean_msg(conditionMessage(e)))))
        })

      } # end Quantification

      openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
      # Clean up temp PNG files now that saveWorkbook has embedded them
      if (length(.plot_tmp_files) > 0) suppressWarnings(unlink(.plot_tmp_files))
    }
  )

  output$analysis_point_no_run_banner    <- .no_run_ui("Point-based Analysis")
  output$analysis_titration_no_run_banner <- .no_run_ui("Titration Analysis")
  output$analysis_quant_no_run_banner    <- .no_run_ui("Quantification")

  # .bama_gate_ui: shows a locked-style banner when this tab's BAMA type is
  # not selected on the Overview tab, or prompts to select when nothing chosen.
  # bama_key = "Point-based"|"Titration"|"Quantification"
  .bama_gate_ui <- function(bama_key, tab_label) {
    renderUI({
      sel_ov <- input$bama_type_select_ov
      sel_hp <- input$bama_type_select
      sel    <- unique(c(sel_ov, sel_hp))
      pill_colours <- c("Point-based"    = "#1e88e5",
                        "Titration"      = "#43a047",
                        "Quantification" = "#8e24aa")
      col <- unname(pill_colours[bama_key])
      if (is.na(col)) col <- "#555"

      # Nothing selected yet -- prompt user
      if (is.null(sel) || length(sel) == 0) {
        return(tags$div(
          class = "helper-locked-banner",
          style = paste0("border-left-color:", col, "; background:#f8f9fb;"),
          tags$div(class = "lock-icon", style = paste0("color:", col, ";"), icon("mouse-pointer")),
          tags$h3(style = paste0("color:", col, ";"),
                  paste0(tab_label, " \u2014 No Assay Type Selected")),
          tags$p(
            "Please go to the ", tags$strong("Overview"), " tab and select ",
            tags$strong(bama_key), " under ",
            tags$strong("BAMA Type \u2014 Assay Selection"), " to enable this analysis."
          )
        ))
      }

      # Type is selected -- show nothing (let content through)
      if (bama_key %in% sel) return(NULL)

      # Type NOT selected (others are) -- full locked banner
      tags$div(
        class = "helper-locked-banner",
        style = paste0("border-left-color:", col, "; background:#f8f9fb;"),
        tags$div(class = "lock-icon", style = paste0("color:", col, ";"), icon("lock")),
        tags$h3(style = paste0("color:", col, ";"),
                paste0(tab_label, " is not enabled")),
        tags$p(
          tags$strong(bama_key), " was not selected as an assay type for this run.",
          tags$br(),
          "To enable this tab, go to the ", tags$strong("Overview"), " tab and select ",
          tags$strong(bama_key), " under ",
          tags$strong("BAMA Type \u2014 Assay Selection"), "."
        )
      )
    })
  }
  output$analysis_point_bama_gate     <- .bama_gate_ui("Point-based",    "Point-based Analysis")
  output$analysis_titration_bama_gate <- .bama_gate_ui("Titration",       "Titration Analysis")
  output$analysis_quant_bama_gate     <- .bama_gate_ui("Quantification",  "Quantification")


  # -- Dynamic Plate Review titles -------------------------------------------
  output$plate_heatmap_title <- renderUI({
    req(rv$raw_df)
    df    <- rv$raw_df
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    sel   <- input$selected_antigen
    lbl <- if (!is.null(sel) && nchar(trimws(sel)) > 0) {
      base_name <- gsub("\\s*\\(\\d+\\)\\s*$", "", sel)
      display <- if (is_ix) {
        acol <- paste0("AnalyteName_", trimws(base_name))
        if (acol %in% colnames(df)) {
          v <- as.character(df[[acol]])
          v <- v[!is.na(v) & nchar(trimws(v)) > 0]
          if (length(v) > 0) trimws(v[1]) else base_name
        } else base_name
      } else base_name
      paste0("MFI -- ", display, if (is_ix) " (blank-bead & no-mab subtracted)" else "")
    } else {
      if (is_ix) "MFI (blank-bead & no-mab subtracted)" else "Plate MFI Values"
    }
    tags$span(style = "font-weight:700;", lbl)
  })

  output$plate_map_title <- renderUI({
    tags$span(style = "font-weight:700;", "Plate Layout")
  })

  # Plate selector -- one choice per antigen detected in the raw file
  output$plate_selector_ui <- renderUI({
    req(rv$raw_df)
    df    <- rv$raw_df
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    is_ag   <- grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(df)) & !startsWith(colnames(df), "Beads_")
    ag_cols <- colnames(df)[is_ag]
    ag_cols <- ag_cols[!startsWith(toupper(ag_cols), "BLANK")]
    ag_cols <- ag_cols[!toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_cols))) %in% c("R44")]

    if (length(ag_cols) == 0) return(NULL)

    ag_display <- sapply(ag_cols, function(col) {
      base_name <- gsub("\\s*\\(\\d+\\)\\s*$", "", col)
      if (is_ix) {
        analyte_col <- paste0("AnalyteName_", trimws(base_name))
        if (analyte_col %in% colnames(df)) {
          vals   <- as.character(df[[analyte_col]])
          non_na <- vals[!is.na(vals) & nchar(trimws(vals)) > 0]
          if (length(non_na) > 0) return(trimws(non_na[1]))
        }
      }
      base_name
    }, USE.NAMES = FALSE)

    choices <- setNames(ag_cols, ag_display)
    selectInput("selected_antigen", label = NULL,
                choices  = choices,
                selected = ag_cols[1],
                width    = "300px")
  })

  # MFI / RLU heatmap
  # Helper: build per-well subtracted MFI for a single antigen column.
  # avg_nomab_ref: if provided (not NULL), use it directly (shared reference from first antigen);
  # otherwise compute it from this antigen's own no_mab wells (fallback / first-antigen case).
  make_antigen_grid <- function(df, ag_col, blank_vals, is_nomab_vec, avg_nomab_ref = NULL) {
    parse_fi <- function(x) {
      x <- gsub("\\s*\\(\\d+\\)\\s*$", "", as.character(x))
      suppressWarnings(as.numeric(gsub(",", ".", x)))
    }
    raw_mfi     <- parse_fi(df[[ag_col]])
    neg_to_zero <- raw_mfi - blank_vals
    neg_to_zero <- ifelse(is.na(neg_to_zero) | neg_to_zero < 0, 0, neg_to_zero)
    avg_nomab   <- if (!is.null(avg_nomab_ref)) {
      avg_nomab_ref
    } else if (any(is_nomab_vec, na.rm = TRUE)) {
      v <- mean(neg_to_zero[is_nomab_vec], na.rm = TRUE)
      if (is.na(v)) 0 else v
    } else 0
    final <- neg_to_zero - avg_nomab
    ifelse(is.na(final) | final < 0, 0, final)
  }

  # Helper: compute shared avg_nomab from the first antigen column
  # (matches lab Excel protocol: single no_mab average applied to all antigens)
  compute_shared_avg_nomab <- function(df, ag_cols, blank_vals, is_nomab_vec) {
    if (length(ag_cols) == 0 || !any(is_nomab_vec, na.rm = TRUE)) return(0)
    parse_fi <- function(x) {
      x <- gsub("\\s*\\(\\d+\\)\\s*$", "", as.character(x))
      suppressWarnings(as.numeric(gsub(",", ".", x)))
    }
    raw_first <- parse_fi(df[[ag_cols[1]]])
    ntz_first <- raw_first - blank_vals
    ntz_first <- ifelse(is.na(ntz_first) | ntz_first < 0, 0, ntz_first)
    v         <- mean(ntz_first[is_nomab_vec], na.rm = TRUE)
    if (is.na(v)) 0 else v
  }

  make_plate_heatmap <- function(grid, antigen_name, plate_id, n_rows = 8L, n_cols = 12L) {
    rlu_max   <- max(grid$RLU, na.rm = TRUE)
    if (!is.finite(rlu_max) || rlu_max == 0) rlu_max <- 1
    scale_mid <- rlu_max / 2

    # Assign text colour per cell: white on dark (above midpoint), black on light
    grid$text_col <- ifelse(is.na(grid$RLU) | grid$RLU <= scale_mid, "black", "white")

    # Format label for every cell; blank only for truly missing wells
    grid$label <- ifelse(
      is.na(grid$RLU), "",
      ifelse(grid$RLU >= 1e6, paste0(round(grid$RLU / 1e6, 1), "M"),
      ifelse(grid$RLU >= 1e3, paste0(round(grid$RLU / 1e3,  1), "k"),
             as.character(round(grid$RLU, 0))))
    )

    title_label <- paste0("MFI \u2014 ", antigen_name, "  [", plate_id, "]")

    # Split into black-text and white-text subsets so colour is a fixed scalar
    # per geom_text call -- completely avoids any colour scale conflict
    g_black <- grid[grid$text_col == "black", , drop = FALSE]
    g_white <- grid[grid$text_col == "white", , drop = FALSE]

    p <- ggplot(grid, aes(x = Col, y = Row)) +
      geom_tile(aes(fill = RLU), color = "white", linewidth = 0.5) +
      scale_fill_gradient2(
        low = "white", mid = "#e57c00", high = "#b71c1c",
        midpoint = scale_mid,
        na.value = "#eeeeee", name = "MFI", labels = scales::comma
      ) +
      annotate("rect",
               xmin = 0.5, xmax = n_cols + 0.5,
               ymin = 0.5, ymax = n_rows + 0.5,
               fill = NA, color = "black", linewidth = 1.5) +
      labs(title = title_label, x = "Column", y = "Row") +
      theme_minimal(base_size = 11) +
      theme(panel.grid   = element_blank(),
            axis.text    = element_text(size = 9),
            plot.title   = element_text(face = "bold", size = 12, hjust = 0.5),
            legend.title = element_text(size = 9),
            legend.text  = element_text(size = 8),
            panel.border = element_blank())

    if (nrow(g_black) > 0)
      p <- p + geom_text(data = g_black, aes(x = Col, y = Row, label = label),
                         colour = "black", size = 2.5, inherit.aes = FALSE)
    if (nrow(g_white) > 0)
      p <- p + geom_text(data = g_white, aes(x = Col, y = Row, label = label),
                         colour = "white", size = 2.5, inherit.aes = FALSE)
    p
  }

  # Single-antigen MFI heatmap -- reacts to input$selected_antigen from the selector
  output$plate_rlu_heatmap_ui <- renderUI({
    req(rv$raw_df, input$selected_antigen)
    dims  <- plate_dims()
    px_ht <- if (dims$n_rows > 8L) "560px" else "380px"
    plotOutput("plate_rlu_selected_ag", height = px_ht)
  })

  output$plate_rlu_selected_ag <- renderPlot({
    req(rv$raw_df, input$selected_antigen)
    df2    <- rv$raw_df
    is_ix2 <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    ag_col_i <- input$selected_antigen   # the raw column name chosen in the dropdown

    # Resolve display name for the selected column
    base_name2  <- gsub("\\s*\\(\\d+\\)\\s*$", "", ag_col_i)
    ag_name_i <- if (is_ix2) {
      analyte_col2 <- paste0("AnalyteName_", trimws(base_name2))
      if (analyte_col2 %in% colnames(df2)) {
        vals2   <- as.character(df2[[analyte_col2]])
        non_na2 <- vals2[!is.na(vals2) & nchar(trimws(vals2)) > 0]
        if (length(non_na2) > 0) trimws(non_na2[1]) else base_name2
      } else base_name2
    } else base_name2

    parse_fi2 <- function(x) {
      x <- gsub("\\s*\\(\\d+\\)\\s*$", "", as.character(x))
      suppressWarnings(as.numeric(gsub(",", ".", x)))
    }

    wells2 <- as.character(df2[["Well"]])
    wt2    <- if ("Type" %in% colnames(df2)) as.character(df2[["Type"]]) else rep("S", nrow(df2))
    desc2  <- if ("Description" %in% colnames(df2)) as.character(df2[["Description"]]) else wells2

    if (is_ix2 && !is.null(rv$well96_df)) {
      w96b <- rv$well96_df
      rc2  <- colnames(w96b)[1]
      nc2  <- colnames(w96b)[-1]
      lbl2 <- do.call(rbind, lapply(seq_len(nrow(w96b)), function(r) {
        rl2 <- as.character(w96b[[rc2]][r])
        do.call(rbind, lapply(nc2, function(cc2) {
          data.frame(Well = paste0(rl2, cc2),
                     SampleName = as.character(w96b[[cc2]][r]),
                     stringsAsFactors = FALSE)
        }))
      }))
      mp2 <- setNames(lbl2$SampleName, lbl2$Well)
      for (ii in seq_along(wells2)) {
        nm2 <- mp2[wells2[ii]]
        if (!is.null(nm2) && !is.na(nm2)) desc2[ii] <- nm2
      }
      is_nomab2 <- (desc2 == "no_mab")
    } else {
      is_nomab2 <- (wt2 == "B")
    }

    blank_col2 <- if (is_ix2) {
      stripped_hm <- toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", colnames(df2))))
      cols_r44_hm <- colnames(df2)[stripped_hm == "R44"]
      if (length(cols_r44_hm) > 0) cols_r44_hm[1] else NA_character_
    } else {
      cols_bl_hm <- colnames(df2)[startsWith(toupper(colnames(df2)), "BLANK")]
      if (length(cols_bl_hm) > 0) cols_bl_hm[1] else NA_character_
    }
    blank_vals2 <- if (!is.na(blank_col2)) parse_fi2(df2[[blank_col2]]) else rep(0, nrow(df2))
    blank_vals2[is.na(blank_vals2)] <- 0

    # Identify all antigen columns so we can compute the shared avg_nomab from the first one
    is_ag2   <- grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(df2))
    ag_cols2 <- colnames(df2)[is_ag2]
    ag_cols2 <- ag_cols2[!startsWith(toupper(ag_cols2), "BLANK")]
    ag_cols2 <- ag_cols2[!toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_cols2))) %in% c("R44")]

    shared_nomab2 <- compute_shared_avg_nomab(df2, ag_cols2, blank_vals2, is_nomab2)
    mfi_vec <- make_antigen_grid(df2, ag_col_i, blank_vals2, is_nomab2, avg_nomab_ref = shared_nomab2)

    dims2     <- plate_dims()
    all_wells <- paste0(rep(dims2$row_letters, each = dims2$n_cols),
                        rep(seq_len(dims2$n_cols), dims2$n_rows))
    base_g <- data.frame(
      Well = all_wells,
      Row  = factor(substr(all_wells, 1, 1), levels = rev(dims2$row_letters)),
      Col  = factor(as.integer(sub("[A-Za-z]", "", all_wells)), levels = seq_len(dims2$n_cols)),
      stringsAsFactors = FALSE
    )
    mfi_df <- data.frame(Well = wells2, RLU = mfi_vec, stringsAsFactors = FALSE)
    grid2  <- merge(base_g, mfi_df, by = "Well", all.x = TRUE)

    make_plate_heatmap(grid2, ag_name_i, ag_name_i,
                       n_rows = dims2$n_rows, n_cols = dims2$n_cols)
  })

  # Plate Layout (Well Type colour + sample label)
  output$plate_map_heatmap <- renderPlot({
    grid        <- full_grid()
    sel_ag <- if (!is.null(input$selected_antigen) && nchar(trimws(input$selected_antigen)) > 0) {
      base_nm  <- gsub("\\s*\\(\\d+\\)\\s*$", "", input$selected_antigen)
      is_ix_pm <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
      raw_df_pm <- rv$raw_df
      if (is_ix_pm && !is.null(raw_df_pm)) {
        acol_pm <- paste0("AnalyteName_", trimws(base_nm))
        if (acol_pm %in% colnames(raw_df_pm)) {
          vv <- as.character(raw_df_pm[[acol_pm]])
          vv <- vv[!is.na(vv) & nchar(trimws(vv)) > 0]
          if (length(vv) > 0) trimws(vv[1]) else base_nm
        } else base_nm
      } else base_nm
    } else "Plate"
    title_label <- paste0("Plate Layout: ", sel_ag)

    # Colour palette -- controls: red, no_mab: dark red, samples/unknowns: orange
    # NOTE: "no_mab" wells (Type B) are real measured wells with no antibody
    # added, used as the background baseline. They are NOT the same as
    # "Blank" wells, which are physically empty/unused wells on the plate
    # (no data at all) and are shown separately in grey.
    grid$TypeColor <- dplyr::case_when(
      grid$WellType == "B"              ~ "no_mab",
      grid$WellType == "C_pos"          ~ "Ctrl_Pos",
      grid$WellType == "C_neg"          ~ "Ctrl_Neg",
      startsWith(grid$WellType, "C")    ~ "Control",
      startsWith(grid$WellType, "S")    ~ "Sample",
      startsWith(grid$WellType, "X")    ~ "Sample",
      grid$WellType == "empty"          ~ "Blank",
      TRUE                              ~ "Sample"
    )

    fill_lum <- c(
      "no_mab"   = 0.05,
      "Ctrl_Pos" = 0.28,
      "Ctrl_Neg" = 0.38,
      "Control"  = 0.30,
      "Sample"   = 0.60,
      "Blank"    = 0.85
    )
    grid$label_col <- ifelse(
      fill_lum[grid$TypeColor] >= 0.55, "black", "white"
    )
    grid$label_col[is.na(grid$label_col)] <- "white"
    dims_pm <- plate_dims()

    ggplot(grid, aes(x = Col, y = Row, fill = TypeColor)) +
      geom_tile(color = "white", linewidth = 0.5) +
      # Black outer border around entire plate
      annotate("rect",
               xmin = 0.5, xmax = dims_pm$n_cols + 0.5,
               ymin = 0.5, ymax = dims_pm$n_rows + 0.5,
               fill = NA, color = "black", linewidth = 1.5) +
      scale_fill_manual(
        name   = "Well Type",
        drop   = FALSE,
        values = c(
          "no_mab"   = "#8B0000",
          "Ctrl_Pos" = "#27ae60",
          "Ctrl_Neg" = "#e53935",
          "Control"  = "#e53935",
          "Sample"   = "#e67e22",
          "Blank"    = "#cccccc"
        ),
        labels = c(
          "no_mab"   = "No mAb",
          "Ctrl_Pos" = "Positive Control",
          "Ctrl_Neg" = "Negative Control",
          "Control"  = "Control",
          "Sample"   = "Sample / Unknown",
          "Blank"    = "Blank (empty well)"
        )
      ) +
      scale_color_identity(guide = "none") +
      labs(title = title_label, x = "Column", y = "Row") +
      theme_minimal(base_size = 11) +
      theme(
        panel.grid      = element_blank(),
        axis.text       = element_text(size = 9),
        plot.title      = element_text(face = "bold", size = 12, hjust = 0.5),
        legend.position = "right",
        legend.title    = element_text(size = 9),
        legend.text     = element_text(size = 8),
        panel.border    = element_blank()
      )
  })

  # Dynamic height wrappers so 384-well layout plots don't clip
  output$plate_map_heatmap_ui <- renderUI({
    dims  <- plate_dims()
    px_ht <- if (dims$n_rows > 8L) "560px" else "380px"
    tags$div(
      style = "position:relative;",
      plotOutput("plate_map_heatmap", height = px_ht,
                 hover = hoverOpts("plate_map_hover", delay = 80, delayType = "debounce")),
      uiOutput("plate_map_tooltip")
    )
  })

  output$plate_map_tooltip <- renderUI({
    hover <- input$plate_map_hover
    if (is.null(hover)) return(NULL)

    grid <- full_grid()
    if (is.null(grid) || nrow(grid) == 0) return(NULL)

    col_idx <- round(hover$x)
    row_idx <- round(hover$y)

    dims <- plate_dims()
    if (col_idx < 1 || col_idx > dims$n_cols || row_idx < 1 || row_idx > dims$n_rows) return(NULL)

    # Row factor is stored reversed (H=1, A=8 for a standard 96-well plate)
    row_letter <- rev(dims$row_letters)[row_idx]
    well_id    <- paste0(row_letter, col_idx)

    well_row <- grid[grid$Well == well_id, ]
    if (nrow(well_row) == 0) return(NULL)

    sample_name   <- well_row$SampleID[1]
    well_type_raw <- well_row$WellType[1]
    well_type_label <- dplyr::case_when(
      well_type_raw == "B"           ~ "No mAb",
      well_type_raw == "C_pos"       ~ "Positive Control",
      well_type_raw == "C_neg"       ~ "Negative Control",
      well_type_raw == "empty"       ~ "Blank (empty well)",
      startsWith(well_type_raw, "C") ~ "Control",
      TRUE                           ~ "Sample / Unknown"
    )

    left_px <- hover$coords_css$x + 14
    top_px  <- hover$coords_css$y - 44

    tags$div(
      style = paste0(
        "position:absolute; left:", left_px, "px; top:", top_px, "px; ",
        "background:rgba(20,20,20,0.90); color:#fff; padding:7px 12px; border-radius:5px; ",
        "font-size:12px; pointer-events:none; z-index:9999; white-space:nowrap; ",
        "box-shadow:0 2px 8px rgba(0,0,0,0.4); line-height:1.6;"
      ),
      tags$strong(style = "font-size:13px;", well_id),
      tags$br(),
      tags$span(paste0("Sample: ", sample_name)),
      tags$br(),
      tags$span(style = "color:#bbb;", paste0("Type: ", well_type_label))
    )
  })

  output$pb_plate_map_ui <- renderUI({
    dims  <- plate_dims()
    px_ht <- if (dims$n_rows > 8L) "500px" else "320px"
    plotOutput("pb_plate_map", height = px_ht)
  })

  # Control Averages -- per selected antigen, per sample (duplicates), blank+no_mab subtracted
  output$control_summary_table <- renderDT({
    req(rv$raw_df, input$selected_antigen)
    df     <- rv$raw_df
    ag_col <- input$selected_antigen
    is_ix  <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    if (!all(c("Well", "Type", "Description") %in% colnames(df))) return(NULL)
    if (!ag_col %in% colnames(df)) return(NULL)

    parse_fi <- function(x) {
      x <- gsub("\\s*\\(\\d+\\)\\s*$", "", as.character(x))
      suppressWarnings(as.numeric(gsub(",", ".", x)))
    }

    df$raw_mfi <- parse_fi(df[[ag_col]])
    df$type    <- as.character(df[["Type"]])
    df$desc    <- as.character(df[["Description"]])
    df$well    <- as.character(df[["Well"]])

    # --- Blank-bead subtraction (BLANK column) ---
    blank_col <- {
      cols_bl <- colnames(df)[startsWith(toupper(colnames(df)), "BLANK")]
      if (length(cols_bl) > 0) cols_bl[1] else NA_character_
    }
    blank_vals <- if (!is.na(blank_col)) parse_fi(df[[blank_col]]) else rep(0, nrow(df))
    blank_vals[is.na(blank_vals)] <- 0

    ntz <- pmax(df$raw_mfi - blank_vals, 0, na.rm = TRUE)
    ntz[is.na(ntz)] <- 0

    # --- no_mab subtraction ---
    # For INTELLIFLEX: locate no_mab wells from the helper's no_mab_range column.
    # For BioPlex: use the conventional Type == "B" rows.
    avg_nomab <- local({
      if (is_ix) {
        helper_nm <- rv$helper_edited
        if (!is.null(helper_nm) && nrow(helper_nm) > 0) {
          h_nm    <- setNames(helper_nm, tolower(trimws(colnames(helper_nm))))
          nm_col  <- grep("^no_mab_range$", colnames(h_nm), value = TRUE)[1]
          if (!is.na(nm_col)) {
            # Collect all unique no_mab wells across helper rows for this antigen
            base_ag_nm  <- trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_col))
            acol_nm     <- paste0("AnalyteName_", base_ag_nm)
            cur_ag_nm   <- if (acol_nm %in% colnames(df)) {
              v <- as.character(df[[acol_nm]]); v <- v[!is.na(v) & nchar(trimws(v)) > 0]
              if (length(v) > 0) trimws(v[1]) else base_ag_nm
            } else base_ag_nm
            ag_col_nm <- grep("^plate_analyte$", colnames(h_nm), value = TRUE)[1]
            rows_nm <- if (!is.na(ag_col_nm)) {
              which(trimws(as.character(h_nm[[ag_col_nm]])) == cur_ag_nm)
            } else seq_len(nrow(h_nm))

            expand_range_nm <- function(range_str) {
              range_str <- trimws(as.character(range_str))
              if (is.na(range_str) || nchar(range_str) == 0) return(character(0))
              segs <- trimws(unlist(strsplit(range_str, ",")))
              out  <- character(0)
              for (seg in segs) {
                seg <- trimws(seg)
                if (grepl(":", seg)) {
                  pts <- strsplit(seg, ":")[[1]]
                  r1  <- toupper(gsub("[0-9]", "", trimws(pts[1])))
                  c1  <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", trimws(pts[1]))))
                  r2  <- toupper(gsub("[0-9]", "", trimws(pts[2])))
                  c2  <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", trimws(pts[2]))))
                  ri1 <- match(r1, LETTERS); ri2 <- match(r2, LETTERS)
                  if (is.na(c1) || is.na(c2) || is.na(ri1) || is.na(ri2)) next
                  out <- c(out, as.vector(outer(LETTERS[ri1:ri2], c1:c2, paste0)))
                } else if (nchar(seg) > 0) out <- c(out, seg)
              }
              out
            }

            nm_wells <- unique(unlist(lapply(rows_nm, function(i)
              expand_range_nm(h_nm[[nm_col]][i]))))
            idx_nm <- which(df$well %in% nm_wells)
            if (length(idx_nm) > 0) {
              v <- mean(ntz[idx_nm], na.rm = TRUE); if (!is.na(v)) v else 0
            } else 0
          } else 0
        } else 0
      } else {
        is_nomab <- df$type == "B"
        if (any(is_nomab, na.rm = TRUE)) {
          v <- mean(ntz[is_nomab], na.rm = TRUE); if (!is.na(v)) v else 0
        } else 0
      }
    })

    df$mfi_sub <- pmax(ntz - avg_nomab, 0, na.rm = TRUE)
    df$mfi_sub[is.na(df$mfi_sub)] <- 0

    if (is_ix) {
      # ---------------------------------------------------------------------
      # INTELLIFLEX: raw Type codes are always "S" (no X/U classification is
      # exported by the instrument), so sample wells can't be located that
      # way. Instead, locate them from the Universal helper file: rows whose
      # sample_type == "sample" give the sample_id, and that row's
      # plate_range gives the wells holding that sample's measurements.
      # ---------------------------------------------------------------------
      helper <- rv$helper_edited
      if (is.null(helper) || nrow(helper) == 0)
        return(datatable(data.frame(Message = "No helper file loaded -- upload the Universal helper file to compute Sample Averages for INTELLIFLEX."),
                         rownames = FALSE, options = list(dom = "t")))

      h    <- helper
      h_lc <- setNames(h, tolower(trimws(colnames(h))))
      hc   <- colnames(h_lc)

      pr_col  <- grep("^plate_range$",   hc, value = TRUE)[1]
      st_col  <- grep("^sample_type$",   hc, value = TRUE)[1]
      sid_col <- grep("^sample_id$",     hc, value = TRUE)[1]
      ag_col_h <- grep("^plate_analyte$", hc, value = TRUE)[1]

      if (is.na(pr_col) || is.na(st_col))
        return(datatable(data.frame(Message = "Helper file is missing 'sample_type' and/or 'plate_range' columns."),
                         rownames = FALSE, options = list(dom = "t")))

      # Resolve the real antigen name for the selected column (e.g. "R34 (34)"
      # -> "HA") so helper rows can be narrowed to the current antigen via
      # its plate_analyte column, when present.
      base_ag   <- trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_col))
      acol      <- paste0("AnalyteName_", base_ag)
      cur_ag    <- if (acol %in% colnames(df)) {
        v <- as.character(df[[acol]]); v <- v[!is.na(v) & nchar(trimws(v)) > 0]
        if (length(v) > 0) trimws(v[1]) else base_ag
      } else base_ag

      mask_sample <- tolower(trimws(as.character(h_lc[[st_col]]))) == "sample"
      mask_sample[is.na(mask_sample)] <- FALSE
      if (!is.na(ag_col_h)) {
        ag_match <- trimws(as.character(h_lc[[ag_col_h]])) == cur_ag
        ag_match[is.na(ag_match)] <- FALSE
        if (any(mask_sample & ag_match, na.rm = TRUE)) mask_sample <- mask_sample & ag_match
      }
      sample_rows <- which(mask_sample)

      if (length(sample_rows) == 0)
        return(datatable(data.frame(Message = "No rows with sample_type == 'sample' found in the helper file for this antigen."),
                         rownames = FALSE, options = list(dom = "t")))

      # Expand "A1:H1", "A2", "A11:F11, A12:F12" style plate_range strings
      # into individual well labels.
      expand_range <- function(range_str) {
        range_str <- trimws(as.character(range_str))
        if (is.na(range_str) || nchar(range_str) == 0) return(character(0))
        segments  <- trimws(unlist(strsplit(range_str, ",")))
        wells_out <- character(0)
        for (seg in segments) {
          seg <- trimws(seg)
          if (grepl(":", seg)) {
            parts <- strsplit(seg, ":")[[1]]
            if (length(parts) != 2) next
            r1 <- toupper(gsub("[0-9]",  "", trimws(parts[1])))
            c1 <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", trimws(parts[1]))))
            r2 <- toupper(gsub("[0-9]",  "", trimws(parts[2])))
            c2 <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", trimws(parts[2]))))
            if (is.na(c1) || is.na(c2)) next
            ri1 <- match(r1, LETTERS); ri2 <- match(r2, LETTERS)
            if (is.na(ri1) || is.na(ri2)) next
            rows <- LETTERS[ri1:ri2]
            cols <- c1:c2
            wells_out <- c(wells_out, as.vector(outer(rows, cols, paste0)))
          } else {
            if (nchar(seg) > 0) wells_out <- c(wells_out, seg)
          }
        }
        wells_out
      }

      # Map well -> sample_id (last matching helper row wins on overlap)
      well_to_sid <- list()
      for (i in sample_rows) {
        sid_val <- if (!is.na(sid_col)) trimws(as.character(h_lc[[sid_col]][i])) else paste0("Sample_", i)
        for (w in expand_range(h_lc[[pr_col]][i]))
          well_to_sid[[w]] <- sid_val
      }

      if (length(well_to_sid) == 0)
        return(datatable(data.frame(Message = "No sample wells found -- check the helper's plate_range values."),
                         rownames = FALSE, options = list(dom = "t")))

      samp_df <- df[df$well %in% names(well_to_sid), , drop = FALSE]
      samp_df$desc <- vapply(samp_df$well, function(w) well_to_sid[[w]], character(1))
      samp_df <- samp_df[!is.na(samp_df$mfi_sub), ]
    } else {
      # --- Keep only sample wells (Type X or U) ---
      samp_df <- df[grepl("^[XxUu]", df$type), ]
      samp_df <- samp_df[!is.na(samp_df$mfi_sub), ]
    }

    if (nrow(samp_df) == 0)
      return(datatable(data.frame(Message = "No sample wells (Type X/U) found for this antigen."),
                       rownames = FALSE, options = list(dom = "t")))

    thresh <- as.numeric(input$cv_threshold %||% 20)

    # Build one row per unique sample_id; assume duplicates share the same Description
    sample_ids <- unique(samp_df$desc)
    summary_df <- do.call(rbind, lapply(sample_ids, function(sid) {
      rows    <- samp_df[samp_df$desc == sid, , drop = FALSE]
      vals    <- sort(rows$mfi_sub, na.last = TRUE)
      n       <- length(vals)
      rep1    <- if (n >= 1) round(vals[1], 2) else NA_real_
      rep2    <- if (n >= 2) round(vals[2], 2) else NA_real_
      avg     <- if (n >= 1) round(mean(vals, na.rm = TRUE), 2)  else NA_real_
      sdev    <- if (n >= 2) round(sd(vals,   na.rm = TRUE), 2)  else NA_real_
      cv      <- if (!is.na(avg) && avg != 0 && !is.na(sdev))
                   round(sdev / avg * 100, 2) else NA_real_
      data.frame(
        Sample_ID = sid,
        `MFI Rep1` = rep1,
        `MFI Rep2` = rep2,
        Avg        = avg,
        SD         = sdev,
        `CV%`      = cv,
        stringsAsFactors = FALSE, check.names = FALSE
      )
    }))
    row.names(summary_df) <- NULL

    datatable(
      summary_df,
      options = list(
        pageLength = 20, scrollX = TRUE, dom = "lfrtip",
        columnDefs = list(list(className = "dt-center", targets = "_all"))
      ),
      rownames = FALSE,
      class    = "stripe hover cell-border compact"
    ) %>%
      formatRound(columns = c("MFI Rep1", "MFI Rep2", "Avg", "SD"), digits = 2) %>%
      formatRound(columns = "CV%", digits = 2) %>%
      formatStyle(
        "CV%",
        backgroundColor = styleInterval(c(thresh, thresh * 1.5),
                                        c("transparent", "#fff3cd", "#f8d7da")),
        color           = styleInterval(c(thresh, thresh * 1.5),
                                        c("inherit", "#856404", "#721c24")),
        fontWeight      = styleInterval(c(thresh), c("normal", "bold"))
      )
  })

  # CV% threshold alert badge -- reacts to selected antigen
  output$cv_threshold_alert <- renderUI({
    req(rv$raw_df, input$selected_antigen)
    df     <- rv$raw_df
    ag_col <- input$selected_antigen
    if (!all(c("Type", "Description") %in% colnames(df))) return(NULL)
    if (!ag_col %in% colnames(df)) return(NULL)

    parse_fi <- function(x) {
      x <- gsub("\\s*\\(\\d+\\)\\s*$", "", as.character(x))
      suppressWarnings(as.numeric(gsub(",", ".", x)))
    }

    df$rlu  <- parse_fi(df[[ag_col]])
    df$type <- as.character(df[["Type"]])
    df$desc <- as.character(df[["Description"]])

    thresh <- as.numeric(input$cv_threshold %||% 20)

    ctrl_df <- df[startsWith(df$type, "C") | df$type == "B", ]
    ctrl_df <- ctrl_df[!is.na(ctrl_df$rlu), ]
    ctrl_df$lbl <- ifelse(ctrl_df$type == "B", "no_mab",
                          paste0(ctrl_df$type, " -- ", ctrl_df$desc))

    cvs <- sapply(split(ctrl_df$rlu, ctrl_df$lbl), function(x) {
      m <- mean(x, na.rm = TRUE)
      if (is.na(m) || m == 0 || length(x) < 2) return(NA_real_)
      sd(x, na.rm = TRUE) / m * 100
    })
    n_exceed <- sum(!is.na(cvs) & cvs > thresh, na.rm = TRUE)

    if (n_exceed == 0) {
      tags$div(class = "cv-threshold-ok",  icon("check-circle"),
               sprintf(" All controls within CV\u2264%.0f%%", thresh))
    } else {
      tags$div(class = "cv-threshold-alert", icon("exclamation-triangle"),
               sprintf(" %d control(s) exceed CV %.0f%%", n_exceed, thresh))
    }
  })

  # -- Data Frame tab -- live MFI table from raw data -------------------------

  # Helper: build the structured MFI data frame (Well, Type, Sample_ID, subtracted MFI cols)
  mfi_dataframe <- reactive({
    req(rv$raw_df)
    df    <- rv$raw_df
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    parse_fi <- function(x) {
      x <- gsub("\\s*\\(\\d+\\)\\s*$", "", as.character(x))
      suppressWarnings(as.numeric(gsub(",", ".", x)))
    }

    # Extract the bead count from the trailing "(NN)" in each cell value
    parse_bead_count <- function(x) {
      m <- regmatches(as.character(x), regexpr("\\((\\d+)\\)\\s*$", as.character(x)))
      v <- suppressWarnings(as.integer(gsub("[^0-9]", "", m)))
      ifelse(length(v) == 0 | is.na(v), NA_integer_, v)
    }
    parse_bead_counts <- function(col_vec) {
      sapply(as.character(col_vec), function(x) {
        m <- regmatches(x, regexpr("\\((\\d+)\\)\\s*$", x))
        if (length(m) == 0 || nchar(m) == 0) return(NA_integer_)
        suppressWarnings(as.integer(gsub("[^0-9]", "", m)))
      }, USE.NAMES = FALSE)
    }

    # Antigen columns (exclude BLANK and R44)
    is_ag   <- grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(df)) & !startsWith(colnames(df), "Beads_")
    ag_cols <- colnames(df)[is_ag]
    ag_cols <- ag_cols[!startsWith(toupper(ag_cols), "BLANK")]
    ag_cols <- ag_cols[!toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_cols))) %in% c("R44")]

    if (length(ag_cols) == 0 || !"Well" %in% colnames(df))
      return(NULL)

    # Clean antigen names (strip bead-count suffix)
    # For INTELLIFLEX: resolve actual antigen name from AnalyteName_<region> column
    # using the first non-NA value (all rows share the same analyte name per region)
    ag_display <- sapply(ag_cols, function(col) {
      base_name <- gsub("\\s*\\(\\d+\\)\\s*$", "", col)
      if (is_ix) {
        # Extract region tag from column name like "R34 (34)" -> "R34"
        region_tag <- trimws(base_name)
        analyte_col <- paste0("AnalyteName_", region_tag)
        if (analyte_col %in% colnames(df)) {
          vals <- as.character(df[[analyte_col]])
          non_na <- vals[!is.na(vals) & nchar(trimws(vals)) > 0]
          if (length(non_na) > 0) return(trimws(non_na[1]))
        }
      }
      base_name
    }, USE.NAMES = FALSE)

    wells     <- as.character(df[["Well"]])
    well_type <- if ("Type" %in% colnames(df)) as.character(df[["Type"]]) else rep("S", nrow(df))
    desc      <- if ("Description" %in% colnames(df)) as.character(df[["Description"]]) else wells

    # ---- Sample_ID resolution -----------------------------------------------
    if (is_ix) {
      # INTELLIFLEX: 96-well file for sample names
      if (!is.null(rv$well96_df)) {
        w96      <- rv$well96_df
        row_col  <- colnames(w96)[1]
        num_cols <- colnames(w96)[-1]
        ix_labels <- do.call(rbind, lapply(seq_len(nrow(w96)), function(r) {
          row_ltr <- as.character(w96[[row_col]][r])
          do.call(rbind, lapply(num_cols, function(cc) {
            data.frame(Well = paste0(row_ltr, cc),
                       SampleName = as.character(w96[[cc]][r]),
                       stringsAsFactors = FALSE)
          }))
        }))
        ix_map <- setNames(ix_labels$SampleName, ix_labels$Well)
        for (i in seq_along(wells)) {
          nm <- ix_map[wells[i]]
          if (!is.null(nm) && !is.na(nm)) desc[i] <- nm
        }
      }

      # Control-type mapping from Step 3 inputs
      ctrl_all  <- trimws(isolate(input$ix_ctrl_wells) %||% character(0))
      ctrl_all  <- ctrl_all[nchar(ctrl_all) > 0]
      pos_wells <- ctrl_all
      neg_wells <- character(0)

      # Identify which rows are controls and which are samples
      is_ctrl  <- (length(pos_wells) > 0 & wells %in% pos_wells) |
                  (length(neg_wells) > 0 & wells %in% neg_wells) |
                  startsWith(well_type, "C")
      is_blank <- desc == "no_mab"
      is_samp  <- !is_ctrl & !is_blank

      # Assign numbered C labels: same Sample_ID -> same C number
      ctrl_ids   <- ifelse(is_ctrl, desc, NA_character_)
      uniq_ctrl  <- unique(ctrl_ids[!is.na(ctrl_ids)])
      ctrl_num   <- setNames(seq_along(uniq_ctrl), uniq_ctrl)
      ctrl_label <- ifelse(is_ctrl,
                           paste0("C", ctrl_num[ctrl_ids]),
                           NA_character_)

      # Assign numbered X labels: same Sample_ID -> same X number
      samp_ids  <- ifelse(is_samp, desc, NA_character_)
      uniq_samp <- unique(samp_ids[!is.na(samp_ids)])
      samp_num  <- setNames(seq_along(uniq_samp), uniq_samp)
      samp_label <- ifelse(is_samp,
                           paste0("X", samp_num[samp_ids]),
                           NA_character_)

      type_display <- dplyr::case_when(
        is_blank             ~ "B",
        !is.na(ctrl_label)   ~ ctrl_label,
        !is.na(samp_label)   ~ samp_label,
        TRUE                 ~ well_type
      )
    } else {
      # BIOPLEX: use 96-well file if supplied (optional), else Description
      if (!is.null(rv$well96_df)) {
        w96      <- rv$well96_df
        # BioPlex 96-well: skip=7 -> Well | Type | Description columns
        w_col <- which(toupper(trimws(colnames(w96))) == "WELL")[1]
        d_col <- which(toupper(trimws(colnames(w96))) == "DESCRIPTION")[1]
        t_col <- which(toupper(trimws(colnames(w96))) == "TYPE")[1]
        if (!is.na(w_col) && !is.na(d_col)) {
          bp_map <- setNames(as.character(w96[[d_col]]), as.character(w96[[w_col]]))
          for (i in seq_along(wells)) {
            nm <- bp_map[wells[i]]
            if (!is.null(nm) && !is.na(nm) && nchar(nm) > 0) desc[i] <- nm
          }
        }
      }

      # Build type_display: plain C1/C2/... (no Positive/Negative label)
      type_display <- dplyr::case_when(
        well_type == "B"           ~ "B",
        startsWith(well_type, "C") ~ well_type,
        TRUE                       ~ well_type
      )
      desc <- ifelse(well_type == "B", "no_mab", desc)
    }

    # Build base data frame
    base_df <- data.frame(
      Well      = wells,
      Type      = type_display,
      Sample_ID = desc,
      stringsAsFactors = FALSE
    )

    # -- Subtraction logic (mirrors add_cal_columns + no_mab step in script) --
    # Step 1: locate the blank-bead column for per-row subtraction.
    #   BioPlex  : column whose name starts with "BLANK" (e.g. "BLANK (44)")
    #   INTELLIFLEX: column whose stripped name is exactly "R44" (e.g. "R44 (44)")
    blank_col <- if (is_ix) {
      stripped <- toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", colnames(df))))
      cols_r44 <- colnames(df)[stripped == "R44"]
      if (length(cols_r44) > 0) cols_r44[1] else NA_character_
    } else {
      cols_bl <- colnames(df)[startsWith(toupper(colnames(df)), "BLANK")]
      if (length(cols_bl) > 0) cols_bl[1] else NA_character_
    }
    blank_vals <- if (!is.na(blank_col)) parse_fi(df[[blank_col]]) else rep(0, nrow(df))
    blank_vals[is.na(blank_vals)] <- 0

    is_nomab_well <- (type_display == "B")   # no_mab rows (Type B)

    # Step 2: compute avg_nomab ONCE from the FIRST antigen column only.
    # This matches the lab Excel protocol: a single shared no_mab average (from the
    # first/reference antigen, e.g. GT1.1) is applied identically to ALL antigens
    # (equivalent to $AQ$45 absolute-reference formula in the results workbook).
    avg_nomab_shared <- if (length(ag_cols) > 0 && any(is_nomab_well, na.rm = TRUE)) {
      raw_first      <- parse_fi(df[[ag_cols[1]]])
      ntz_first      <- raw_first - blank_vals
      ntz_first      <- ifelse(is.na(ntz_first) | ntz_first < 0, 0, ntz_first)
      v              <- mean(ntz_first[is_nomab_well], na.rm = TRUE)
      if (is.na(v)) 0 else v
    } else 0

    # Three parallel data frames capturing each subtraction stage
    df_mabs  <- base_df   # Tab 1: avg_nomab only
    df_blank <- base_df   # Tab 2: blank-bead only
    df_full  <- base_df   # Tab 3: blank-bead + avg_nomab (background subtraction)

    # ---- % Agg Beads (well-level; same value applies to every antigen) -------
    # INTELLIFLEX: pre-computed at upload time from TOTAL EVENTS/TOTAL GATED EVENTS
    #              into raw_df$PctAggBeads.
    # BioPlex:     raw file has a direct "% Agg Beads" column (comma-decimal).
    pct_agg_values <- if (is_ix) {
      if ("PctAggBeads" %in% colnames(df)) suppressWarnings(as.numeric(df[["PctAggBeads"]]))
      else rep(NA_real_, nrow(df))
    } else {
      agg_col <- colnames(df)[grepl("Agg[^A-Za-z0-9]*Bead", colnames(df), ignore.case = TRUE)][1]
      if (!is.na(agg_col) && !is.null(agg_col)) {
        suppressWarnings(as.numeric(gsub(",", ".", as.character(df[[agg_col]]))))
      } else rep(NA_real_, nrow(df))
    }

    for (i in seq_along(ag_cols)) {
      raw_mfi <- parse_fi(df[[ag_cols[i]]])

      # Blank-bead subtraction per row, floor to 0
      neg_to_zero <- raw_mfi - blank_vals
      neg_to_zero <- ifelse(is.na(neg_to_zero) | neg_to_zero < 0, 0, neg_to_zero)

      # Tab 1 -- avg_nomab subtraction only (no blank-bead)
      mabs_val <- raw_mfi - avg_nomab_shared
      df_mabs[[ag_display[i]]] <- ifelse(is.na(mabs_val) | mabs_val < 0, 0, mabs_val)

      # Tab 2 -- blank-bead subtraction only
      df_blank[[ag_display[i]]] <- neg_to_zero

      # Tab 3 -- both subtractions (full background correction)
      final_mfi <- neg_to_zero - avg_nomab_shared
      df_full[[ag_display[i]]] <- ifelse(is.na(final_mfi) | final_mfi < 0, 0, final_mfi)

      # Extract bead counts ("Beads_<antigen>") for EVERY antigen column.
      # BioPlex: the raw cell looks like "23028,00 (63)" -- the (63) is the bead count.
      # INTELLIFLEX: real bead counts come from the ": RP1 COUNT" / ": RP2 COUNT"
      # column for this region, captured earlier as "Beads_<median_colname>".
      bead_col_name <- paste0("Beads_", ag_display[i])
      bead_counts   <- if (is_ix) {
        src_col <- paste0("Beads_", ag_cols[i])
        if (src_col %in% colnames(df)) suppressWarnings(as.numeric(df[[src_col]]))
        else parse_bead_counts(df[[ag_cols[i]]])
      } else {
        parse_bead_counts(df[[ag_cols[i]]])
      }
      df_mabs[[bead_col_name]]  <- bead_counts
      df_blank[[bead_col_name]] <- bead_counts
      df_full[[bead_col_name]]  <- bead_counts

      # % Agg Beads is a well-level metric (aggregation/doublet gating happens
      # once per well, not per analyte) -- the same value is shown alongside
      # each antigen's bead count for easy side-by-side QC review.
      agg_col_name <- paste0("AggPct_", ag_display[i])
      df_mabs[[agg_col_name]]  <- pct_agg_values
      df_blank[[agg_col_name]] <- pct_agg_values
      df_full[[agg_col_name]]  <- pct_agg_values
    }

    list(mabs = df_mabs, blank = df_blank, full = df_full)
  })


  # Helper: render a standard per-well MFI DT with type colour coding
  render_mfi_dt <- function(df, lod = NULL) {
    # Drop QC-only AggPct_* and Beads_* columns -- bead counts are shown
    # separately on the QC 1: Bead Acquisition table / 09_bead_information export.
    df <- df[, !startsWith(colnames(df), "AggPct_") & !startsWith(colnames(df), "Beads_"),
             drop = FALSE]
    fixed_cols <- c("Well", "Type", "Sample_ID")
    mfi_cols   <- setdiff(colnames(df), fixed_cols)

    dt <- datatable(
      df,
      options = list(
        pageLength = 20, scrollX = TRUE, dom = "lfrtip",
        columnDefs = list(list(className = "dt-center", targets = "_all"))
      ),
      rownames = FALSE, class = "stripe hover cell-border compact"
    )

    # Round only MFI columns
    if (length(mfi_cols) > 0) dt <- formatRound(dt, columns = mfi_cols, digits = 1)
    dt
  }

  # Helper: render duplicate-average DT
  render_avg_dt <- function(df) {
    # Drop QC-only AggPct_* columns (only shown in QC 1: Bead Acquisition table)
    df <- df[, !startsWith(colnames(df), "AggPct_"), drop = FALSE]
    # Exclude Beads_* from averaging -- bead counts aren't averaged
    Beads_cols <- grep("^Beads_", colnames(df), value = TRUE)
    mfi_cols <- setdiff(colnames(df), c("Well", "Type", "Sample_ID", Beads_cols))
    if (length(mfi_cols) == 0) return(NULL)
    dup_ids <- names(which(table(df$Sample_ID) > 1))
    if (length(dup_ids) == 0) {
      return(datatable(
        data.frame(Message = "No duplicate Sample_IDs detected -- all wells are unique."),
        rownames = FALSE, options = list(dom = "t")
      ))
    }
    df_dups <- df[df$Sample_ID %in% dup_ids, ]
    avg_df  <- do.call(rbind, lapply(dup_ids, function(sid) {
      rows <- df_dups[df_dups$Sample_ID == sid, , drop = FALSE]
      avgs <- sapply(mfi_cols, function(col) mean(rows[[col]], na.rm = TRUE))
      res  <- data.frame(Sample_ID = sid, Type = paste(unique(rows$Type), collapse = "/"),
                         Wells = paste(rows$Well, collapse = ", "), n_reps = nrow(rows),
                         stringsAsFactors = FALSE)
      for (col in mfi_cols) res[[col]] <- avgs[[col]]
      res
    }))
    avg_mfi_cols <- intersect(mfi_cols, colnames(avg_df))
    datatable(avg_df,
      options = list(pageLength = 20, scrollX = TRUE, dom = "lfrtip",
                     columnDefs = list(list(className = "dt-center", targets = "_all"))),
      rownames = FALSE, class = "stripe hover cell-border compact"
    ) %>% formatRound(columns = avg_mfi_cols, digits = 1)
  }

  # -- Tab 1: mAbs subtraction --------------------------------------------------
  output$df_mabs_table <- renderDT({
    res <- mfi_dataframe(); req(res)
    render_mfi_dt(res$mabs, lod = pb_cfg()$lod)
  })
  output$df_mabs_avg_table <- renderDT({
    res <- mfi_dataframe(); req(res); render_avg_dt(res$mabs)
  })

  # -- Tab 2: blank beads subtraction -------------------------------------------
  output$df_blank_table <- renderDT({
    res <- mfi_dataframe(); req(res)
    render_mfi_dt(res$blank, lod = pb_cfg()$lod)
  })
  output$df_blank_avg_table <- renderDT({
    res <- mfi_dataframe(); req(res); render_avg_dt(res$blank)
  })

  output$dataframe_mfi_table <- renderDT({
    res <- mfi_dataframe(); req(res)
    render_mfi_dt(res$full, lod = pb_cfg()$lod)
  })

  output$dataframe_avg_table <- renderDT({
    res <- mfi_dataframe(); req(res); render_avg_dt(res$full)
  })

  # -- Plate Review: MFU per antigen table (blank-bead + no_mab subtracted) ---
  output$review_mfu_table <- renderDT({
    req(rv$raw_df)
    df    <- rv$raw_df
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    parse_fi <- function(x) {
      x <- gsub("\\s*\\(\\d+\\)\\s*$", "", as.character(x))
      suppressWarnings(as.numeric(gsub(",", ".", x)))
    }

    # Identify all columns with bead-count suffix
    all_ag_raw <- colnames(df)[grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(df))]

    # Separate BLANK cols from antigen cols; also exclude R44
    blank_cols <- all_ag_raw[startsWith(toupper(all_ag_raw), "BLANK")]
    ag_cols    <- all_ag_raw[!startsWith(toupper(all_ag_raw), "BLANK")]
    ag_cols    <- ag_cols[!toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_cols))) %in% c("R44")]

    if (length(ag_cols) == 0) {
      return(datatable(
        data.frame(Message = "No antigen columns detected."),
        rownames = FALSE, options = list(dom = "t")
      ))
    }

    # Well type vector -- for INTELLIFLEX, identify no_mab from 96-well layout
    wells_rv  <- as.character(df[["Well"]])
    well_type <- if ("Type" %in% colnames(df)) as.character(df[["Type"]]) else rep("S", nrow(df))
    desc_rv   <- if ("Description" %in% colnames(df)) as.character(df[["Description"]]) else wells_rv

    if (is_ix && !is.null(rv$well96_df)) {
      w96_rv    <- rv$well96_df
      row_col96 <- colnames(w96_rv)[1]
      num_cols96 <- colnames(w96_rv)[-1]
      ix_lbl96 <- do.call(rbind, lapply(seq_len(nrow(w96_rv)), function(r) {
        rl <- as.character(w96_rv[[row_col96]][r])
        do.call(rbind, lapply(num_cols96, function(cc) {
          data.frame(Well = paste0(rl, cc),
                     SampleName = as.character(w96_rv[[cc]][r]),
                     stringsAsFactors = FALSE)
        }))
      }))
      ix_map96 <- setNames(ix_lbl96$SampleName, ix_lbl96$Well)
      for (i in seq_along(wells_rv)) {
        nm <- ix_map96[wells_rv[i]]
        if (!is.null(nm) && !is.na(nm)) desc_rv[i] <- nm
      }
      is_nomab_well <- (desc_rv == "no_mab")
    } else {
      is_nomab_well <- (well_type == "B")
    }

    # Locate blank-bead column: R44 (INTELLIFLEX) or BLANK prefix (BioPlex)
    blank_col <- if (is_ix) {
      stripped_rmt <- toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", colnames(df))))
      cols_r44_rmt <- colnames(df)[stripped_rmt == "R44"]
      if (length(cols_r44_rmt) > 0) cols_r44_rmt[1] else NA_character_
    } else {
      cols_bl_rmt <- colnames(df)[startsWith(toupper(colnames(df)), "BLANK")]
      if (length(cols_bl_rmt) > 0) cols_bl_rmt[1] else NA_character_
    }
    blank_vals <- if (!is.na(blank_col)) parse_fi(df[[blank_col]]) else rep(0, nrow(df))
    blank_vals[is.na(blank_vals)] <- 0

    # Compute MFU summary per antigen
    ag_display <- sapply(ag_cols, function(col) {
      base_name <- gsub("\\s*\\(\\d+\\)\\s*$", "", col)
      if (is_ix) {
        region_tag  <- trimws(base_name)
        analyte_col <- paste0("AnalyteName_", region_tag)
        if (analyte_col %in% colnames(df)) {
          vals    <- as.character(df[[analyte_col]])
          non_na  <- vals[!is.na(vals) & nchar(trimws(vals)) > 0]
          if (length(non_na) > 0) return(trimws(non_na[1]))
        }
      }
      base_name
    }, USE.NAMES = FALSE)

    # Compute avg_nomab ONCE from the first antigen (shared reference, matching the
    # lab Excel protocol where $AQ$45 -- GT1.1's no_mab average -- is applied to all antigens)
    avg_nomab_shared <- if (length(ag_cols) > 0 && any(is_nomab_well, na.rm = TRUE)) {
      raw_first <- parse_fi(df[[ag_cols[1]]])
      ntz_first <- raw_first - blank_vals
      ntz_first <- ifelse(is.na(ntz_first) | ntz_first < 0, 0, ntz_first)
      v         <- mean(ntz_first[is_nomab_well], na.rm = TRUE)
      if (is.na(v)) 0 else v
    } else 0

    mfu_rows <- lapply(seq_along(ag_cols), function(i) {
      raw_mfi <- parse_fi(df[[ag_cols[i]]])

      # Step 1: subtract BLANK (44) per row, floor to 0 (neg-to-zero)
      neg_to_zero <- raw_mfi - blank_vals
      neg_to_zero <- ifelse(is.na(neg_to_zero) | neg_to_zero < 0, 0, neg_to_zero)

      # Step 2: subtract the shared avg_nomab (keep negatives)
      mfu <- neg_to_zero - avg_nomab_shared
      mfu <- ifelse(is.na(mfu) | mfu < 0, 0, mfu)

      # Summary stats exclude no_mab wells
      mfu_samples <- mfu[!is_nomab_well]

      data.frame(
        Antigen         = ag_display[i],
        N_Wells         = sum(!is_nomab_well & !is.na(raw_mfi)),
        Blank_Bead_Mean = round(mean(blank_vals, na.rm = TRUE), 1),
        NoMab_Mean      = round(avg_nomab_shared, 1),
        MFU_Mean        = round(mean(mfu_samples, na.rm = TRUE), 1),
        MFU_SD          = round(sd(mfu_samples,   na.rm = TRUE), 1),
        MFU_Min         = round(min(mfu_samples,  na.rm = TRUE), 1),
        MFU_Max         = round(max(mfu_samples,  na.rm = TRUE), 1),
        stringsAsFactors = FALSE
      )
    })

    mfu_df <- do.call(rbind, mfu_rows)
    row.names(mfu_df) <- NULL

    colnames(mfu_df) <- c(
      "Antigen", "N Wells", "Blank-Bead Mean", "no_mab Mean",
      "MFIs Mean", "MFIs SD", "MFIs Min", "MFIs Max"
    )

    numeric_cols <- c("Blank-Bead Mean", "no_mab Mean", "MFIs Mean", "MFIs SD", "MFIs Min", "MFIs Max")

    datatable(
      mfu_df,
      options = list(
        pageLength = 20,
        scrollX    = TRUE,
        dom        = "lfrtip",
        columnDefs = list(list(className = "dt-center", targets = "_all"))
      ),
      rownames = FALSE,
      class    = "stripe hover cell-border compact"
    ) %>%
      formatRound(columns = numeric_cols, digits = 1) %>%
      formatStyle(
        "MFIs Mean",
        background = styleColorBar(range(mfu_df[["MFIs Mean"]], na.rm = TRUE), "#b3d9ff"),
        backgroundSize = "98% 88%",
        backgroundRepeat = "no-repeat",
        backgroundPosition = "center"
      )
  })

  # ===========================================================================
  # POINT-BASED QC  (BioPlex only)
  # ===========================================================================

  # -- Helper: resolve a raw antigen column key (e.g. "R34 (34)") to its
  # display name (e.g. "CFp10_SG" for INTELLIFLEX, or just the stripped base
  # name for BioPlex). Mirrors the AnalyteName_<region> resolution used when
  # building ag_display / labels elsewhere, so callers can reliably match
  # against out$analyte / df_full column names built the same way.
  .ag_raw_to_display <- function(raw_key) {
    raw_key <- trimws(raw_key %||% "")
    if (nchar(raw_key) == 0) return("")
    df    <- rv$raw_df
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    base  <- trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", raw_key))
    if (is_ix && !is.null(df)) {
      acol <- paste0("AnalyteName_", base)
      if (acol %in% colnames(df)) {
        v <- as.character(df[[acol]])
        v <- v[!is.na(v) & nchar(trimws(v)) > 0]
        if (length(v) > 0) return(trimws(v[1]))
      }
    }
    base
  }

  # -- Helper: per-antigen LOD lookup, shared by the Titration curve plot,
  # the QC Plot per Antigen, and their Excel-export equivalents.
  # For BioPlex, LOD is 0 for non-BRILLIANT runs (forced via pb_cfg /
  # .save_ag_cfg), so a 0/missing LOD naturally means "don't draw a line".
  # Returns data.frame(analyte, lod, lod_locked) for the given analytes.
  .tit_lod_lookup <- function(analytes) {
    ag_cfg_display_lod <- list()
    for (raw_key in names(rv$ag_config)) {
      dk <- .ag_raw_to_display(raw_key)
      ag_cfg_display_lod[[dk]] <- rv$ag_config[[raw_key]]
    }
    live_ag_raw     <- input$pb_selected_antigen %||% ""
    live_ag_display <- .ag_raw_to_display(live_ag_raw)
    live_lod        <- suppressWarnings(as.numeric(input$pb_lod)) %||% 2674
    global_lod      <- live_lod   # live input is the most up-to-date global fallback

    lod_rows <- lapply(analytes, function(ag) {
      # For the currently selected antigen use the live LOD input directly
      if (nchar(live_ag_display) > 0 && ag == live_ag_display) {
        return(data.frame(analyte = ag, lod = live_lod, lod_locked = FALSE,
                          stringsAsFactors = FALSE))
      }
      cfg     <- ag_cfg_display_lod[[ag]]
      lod_val <- suppressWarnings(as.numeric(cfg$lod))
      # Use per-antigen saved value if present; otherwise fall back to global input
      if (!is.null(cfg) && length(lod_val) == 1L && !is.na(lod_val) && lod_val > 0)
        data.frame(analyte = ag, lod = lod_val, lod_locked = isTRUE(cfg$lod_locked),
                   stringsAsFactors = FALSE)
      else
        data.frame(analyte = ag, lod = global_lod, lod_locked = FALSE,
                   stringsAsFactors = FALSE)
    })
    do.call(rbind, c(Filter(Negate(is.null), lod_rows),
                     list(data.frame(analyte = character(0), lod = numeric(0),
                                     lod_locked = logical(0)))))
  }

  # -- Antigen selector for Point-based tab -----------------------------------
  output$pb_antigen_selector_ui <- renderUI({
    req(rv$raw_df)
    df    <- rv$raw_df
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    is_ag <- grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(df)) & !startsWith(colnames(df), "Beads_")
    ag_cols <- colnames(df)[is_ag]
    ag_cols <- ag_cols[!startsWith(toupper(ag_cols), "BLANK")]
    ag_cols <- ag_cols[!toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_cols))) %in% c("R44")]
    if (length(ag_cols) == 0) return(tags$em("No antigens detected."))

    # For INTELLIFLEX resolve the actual analyte name from AnalyteName_<region>
    labels <- sapply(ag_cols, function(col) {
      base <- trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", col))
      if (is_ix) {
        acol <- paste0("AnalyteName_", base)
        if (acol %in% colnames(df)) {
          v <- as.character(df[[acol]])
          v <- v[!is.na(v) & nchar(trimws(v)) > 0]
          if (length(v) > 0) return(trimws(v[1]))
        }
      }
      base
    }, USE.NAMES = FALSE)

    choices <- setNames(ag_cols, labels)

    # Preserve whatever antigen the user had selected (e.g. after clicking
    # Run Analysis again) instead of always resetting to the first antigen --
    # only fall back to ag_cols[1] if there was no prior selection or it's
    # no longer among the available choices.
    prev_sel <- isolate(input$pb_selected_antigen)
    sel_val  <- if (!is.null(prev_sel) && prev_sel %in% ag_cols) prev_sel else ag_cols[1]

    selectInput("pb_selected_antigen", NULL,
                choices = choices, selected = sel_val, width = "100%")
  })


  # -- Reactive: control sample names from helper file -------------------------
  # Reads the helper's sample_id column for rows where sample_type == "control".
  # Returns the real names (e.g. "cap2", "human_IgG") for use in selectInput.
  # -- Reactive: control names from helper's sample_type column ---------------
  # Finds all rows where sample_type == "control" and returns their sample_id
  # values as a named character vector for use in pos/neg selectInputs.
  # Per-analyte filtering: if a plate_analyte column exists, returns controls
  # relevant to the currently selected analyte (or all if none selected).
  # -- Reactive: control names from raw MFI file Type + Description columns ---
  # Controls are rows where Type starts with "C" (C1, C2, C3...).
  # Returns unique Description values (e.g. "cap 2", "human IgG") as choices.
  pb_ctrl_choices <- reactive({
    # INTELLIFLEX: controls are the wells selected in Step 3 (input$ix_ctrl_wells),
    # labelled with their sample names from the 96-well plate layout.
    if (identical(rv$instrument, "INTELLIFLEX")) {
      ctrl_wells <- trimws(input$ix_ctrl_wells %||% character(0))
      ctrl_wells <- ctrl_wells[nchar(ctrl_wells) > 0]
      if (length(ctrl_wells) == 0) return(character(0))

      w96 <- rv$well96_df
      if (is.null(w96)) return(setNames(ctrl_wells, ctrl_wells))

      row_col  <- colnames(w96)[1]
      num_cols <- colnames(w96)[-1]
      ix_labels <- do.call(rbind, lapply(seq_len(nrow(w96)), function(r) {
        row_ltr <- as.character(w96[[row_col]][r])
        do.call(rbind, lapply(num_cols, function(cc) {
          data.frame(Well = paste0(row_ltr, cc),
                     SampleName = as.character(w96[[cc]][r]),
                     stringsAsFactors = FALSE)
        }))
      }))
      ix_map <- setNames(ix_labels$SampleName, ix_labels$Well)
      labels <- vapply(ctrl_wells, function(w) {
        nm <- ix_map[[w]]
        if (!is.null(nm) && !is.na(nm) && nchar(trimws(nm)) > 0) trimws(nm) else w
      }, character(1))
      sids <- sort(unique(labels))
      return(setNames(sids, sids))
    }

    df <- rv$raw_df
    if (is.null(df)) return(character(0))
    if (!all(c("Type", "Description") %in% colnames(df))) return(character(0))

    type_vals <- toupper(trimws(as.character(df[["Type"]])))
    ctrl_mask <- grepl("^C\\d*$", type_vals)
    desc_vals <- trimws(as.character(df[["Description"]][ctrl_mask]))
    desc_vals <- desc_vals[!is.na(desc_vals) & nchar(desc_vals) > 0]
    sids <- sort(unique(desc_vals))
    setNames(sids, sids)  # name = value = Description string
  })

  # -- Reactive: shared plate/antigen choices --------------------------------
  # Same column-detection logic used by the single Select Plate/Antigen
  # dropdown; factored out so the "apply controls to multiple plates/antigens"
  # picker (Step 2) offers the identical list.
  pb_antigen_choices <- reactive({
    req(rv$raw_df)
    df    <- rv$raw_df
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    is_ag <- grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(df)) & !startsWith(colnames(df), "Beads_")
    ag_cols <- colnames(df)[is_ag]
    ag_cols <- ag_cols[!startsWith(toupper(ag_cols), "BLANK")]
    ag_cols <- ag_cols[!toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_cols))) %in% c("R44")]
    if (length(ag_cols) == 0) return(character(0))

    labels <- sapply(ag_cols, function(col) {
      base <- trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", col))
      if (is_ix) {
        acol <- paste0("AnalyteName_", base)
        if (acol %in% colnames(df)) {
          v <- as.character(df[[acol]])
          v <- v[!is.na(v) & nchar(trimws(v)) > 0]
          if (length(v) > 0) return(trimws(v[1]))
        }
      }
      base
    }, USE.NAMES = FALSE)

    setNames(ag_cols, labels)
  })

  # -- Bulk "apply controls to multiple plates/antigens" pickers -------------
  output$pb_bulk_pos_ctrl_ui <- renderUI({
    choices <- pb_ctrl_choices()
    if (length(choices) == 0) return(tags$em("No controls."))
    selectInput("pb_bulk_pos_ctrl", NULL, choices = choices,
                multiple = TRUE, width = "100%")
  })

  output$pb_bulk_neg_ctrl_ui <- renderUI({
    choices <- pb_ctrl_choices()
    if (length(choices) == 0) return(tags$em("No controls."))
    selectInput("pb_bulk_neg_ctrl", NULL, choices = choices,
                multiple = TRUE, width = "100%")
  })

  output$pb_bulk_antigen_ui <- renderUI({
    choices <- pb_antigen_choices()
    if (length(choices) == 0) return(tags$em("No plates/antigens detected."))
    selectInput("pb_bulk_antigens", NULL, choices = choices,
                multiple = TRUE, width = "100%")
  })

  # -- Apply Step 1 controls to every plate/antigen chosen in Step 2 ---------
  # (feedback shown via showNotification() inside the observer below)
  observeEvent(input$pb_bulk_apply_ctrl, {
    ags <- input$pb_bulk_antigens
    pos <- input$pb_bulk_pos_ctrl
    neg <- input$pb_bulk_neg_ctrl

    if (is.null(ags) || length(ags) == 0) {
      showNotification("Select at least one plate/antigen to apply the controls to.",
                        type = "warning")
      return(invisible(NULL))
    }
    if ((is.null(pos) || length(pos) == 0) && (is.null(neg) || length(neg) == 0)) {
      showNotification("Select at least one Positive or Negative control first.",
                        type = "warning")
      return(invisible(NULL))
    }

    for (ag in ags) {
      cfg <- .ag_cfg_get(ag)
      if (!is.null(pos) && length(pos) > 0) cfg$pos_ctrl <- pos
      if (!is.null(neg) && length(neg) > 0) cfg$neg_ctrl <- neg
      rv$ag_config[[ag]] <- cfg
    }

    # If the plate/antigen currently shown in the fine-tune panel below was
    # among those just updated, refresh its live widgets so what's on screen
    # matches what was just saved (pb_config_row itself only re-renders on
    # antigen switch / lock clicks, per the note above it).
    cur <- input$pb_selected_antigen
    if (!is.null(cur) && cur %in% ags) {
      if (!is.null(pos) && length(pos) > 0) updateSelectInput(session, "pb_pos_ctrl", selected = pos)
      if (!is.null(neg) && length(neg) > 0) updateSelectInput(session, "pb_neg_ctrl", selected = neg)
    }

    showNotification(
      sprintf("Applied controls to %d plate(s)/antigen(s).", length(ags)),
      type = "message"
    )
  })

  # ---------------------------------------------------------------------------
  # Per-antigen config: lock/unlock LOD, LOQ, CV%, pos/neg controls
  # ---------------------------------------------------------------------------

  # Helper: get saved config for an antigen (or defaults)
  # INTELLIFLEX defaults LOD/LOQ/CV to 0; BioPlex uses established thresholds.
  .ag_defaults <- function() {
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    list(
      lod      = if (is_ix) 0 else 2674,
      loq      = if (is_ix) 0 else 4301,
      cv       = if (is_ix) 0 else 30,
      pos_ctrl = NULL, neg_ctrl = NULL,
      lod_locked = FALSE, loq_locked = FALSE,
      cv_locked  = FALSE, pos_locked = FALSE, neg_locked = FALSE
    )
  }

  .ag_cfg_get <- function(ag) {
    cfg <- rv$ag_config[[ag]]
    if (is.null(cfg)) .ag_defaults() else cfg
  }

  # Reactive returning effective values for current antigen
  pb_cfg <- reactive({
    ag    <- input$pb_selected_antigen %||% ""
    cfg   <- .ag_cfg_get(ag)
    is_ix <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    # Non-BRILLIANT BioPlex runs have no validated LOD/LOQ -- force both to 0
    # regardless of saved/locked config or the raw numeric input values.
    is_bioplex_now   <- !is_ix
    is_non_brilliant <- is_bioplex_now && !isTRUE(rv$brilliant_bioplex)
    lod_def <- if (is_ix) 0 else 2674
    loq_def <- if (is_ix) 0 else 430
    cv_def  <- if (is_ix) 0 else 20
    list(
      lod      = if (is_non_brilliant) 0 else
                 if (isTRUE(cfg$lod_locked)) cfg$lod      else (as.numeric(input$pb_lod)      %||% lod_def),
      loq      = if (is_non_brilliant) 0 else
                 if (isTRUE(cfg$loq_locked)) cfg$loq      else (as.numeric(input$pb_loq)      %||% loq_def),
      cv       = if (isTRUE(cfg$cv_locked))  cfg$cv       else (as.numeric(input$pb_cv_accept) %||% cv_def),
      pos_ctrl = if (isTRUE(cfg$pos_locked)) cfg$pos_ctrl else input$pb_pos_ctrl,
      neg_ctrl = if (isTRUE(cfg$neg_ctrl_locked %||% cfg$neg_locked)) cfg$neg_ctrl else input$pb_neg_ctrl,
      lod_locked = isTRUE(cfg$lod_locked), loq_locked = isTRUE(cfg$loq_locked),
      cv_locked  = isTRUE(cfg$cv_locked),  pos_locked = isTRUE(cfg$pos_locked),
      neg_locked = isTRUE(cfg$neg_locked)
    )
  })

  # Save current inputs into ag_config for current antigen
  .save_ag_cfg <- function(ag) {
    cfg <- .ag_cfg_get(ag)
    cfg$lod      <- as.numeric(input$pb_lod)      %||% 2674
    cfg$loq      <- as.numeric(input$pb_loq)      %||% 430
    cfg$cv       <- as.numeric(input$pb_cv_accept) %||% 20
    cfg$pos_ctrl <- input$pb_pos_ctrl
    cfg$neg_ctrl <- input$pb_neg_ctrl
    rv$ag_config[[ag]] <- cfg
  }

  # Track previous antigen so we can save its config before switching
  prev_antigen <- reactiveVal(NULL)

  # When antigen selection changes: save previous antigen config, then restore new one
  observeEvent(input$pb_selected_antigen, {
    ag  <- input$pb_selected_antigen; req(ag)
    # Save config for the antigen we are leaving (if any)
    prev <- prev_antigen()
    if (!is.null(prev) && nchar(prev) > 0 && prev != ag) {
      .save_ag_cfg(prev)
    }
    prev_antigen(ag)
    # Restore saved config for the new antigen
    cfg <- .ag_cfg_get(ag)
    is_bioplex_now  <- is.null(rv$instrument) || rv$instrument == "BIOPLEX"
    is_brilliant_now <- !is_bioplex_now || isTRUE(rv$brilliant_bioplex)
    updateNumericInput(session, "pb_lod",       value = if (is_brilliant_now) cfg$lod else 0)
    updateNumericInput(session, "pb_loq",       value = if (is_brilliant_now) cfg$loq else 0)
    updateNumericInput(session, "pb_cv_accept", value = cfg$cv)
    # pos/neg controls restored via pb_config_row re-render (selected= driven by cfg)
  }, ignoreNULL = TRUE)

  # -- BRILLIANT assay toggle (BioPlex only) ---------------------------------
  # Non-BRILLIANT BioPlex runs have no validated LOD/LOQ, so both are forced
  # to 0 and the numeric inputs are locked read-only until switched back.
  observeEvent(input$pb_brilliant_toggle, {
    rv$brilliant_bioplex <- identical(input$pb_brilliant_toggle, "yes")
    if (!rv$brilliant_bioplex) {
      updateNumericInput(session, "pb_lod", value = 0)
      updateNumericInput(session, "pb_loq", value = 0)
      ag <- input$pb_selected_antigen
      if (!is.null(ag) && nchar(trimws(ag)) > 0) .save_ag_cfg(ag)
    } else {
      ag <- input$pb_selected_antigen
      if (!is.null(ag) && nchar(trimws(ag)) > 0) {
        cfg <- .ag_cfg_get(ag)
        updateNumericInput(session, "pb_lod", value = cfg$lod)
        updateNumericInput(session, "pb_loq", value = cfg$loq)
      }
    }
  }, ignoreNULL = TRUE, ignoreInit = TRUE)

  # Toggle lock observers -- each toggle saves current state then flips the lock flag
  observeEvent(input$pb_lod_lock_toggle, {
    ag <- input$pb_selected_antigen; req(ag)
    .save_ag_cfg(ag)
    cfg <- rv$ag_config[[ag]]
    cfg$lod_locked <- !isTRUE(cfg$lod_locked)
    rv$ag_config[[ag]] <- cfg
  })
  observeEvent(input$pb_loq_lock_toggle, {
    ag <- input$pb_selected_antigen; req(ag)
    .save_ag_cfg(ag)
    cfg <- rv$ag_config[[ag]]
    cfg$loq_locked <- !isTRUE(cfg$loq_locked)
    rv$ag_config[[ag]] <- cfg
  })
  observeEvent(input$pb_cv_lock_toggle, {
    ag <- input$pb_selected_antigen; req(ag)
    .save_ag_cfg(ag)
    cfg <- rv$ag_config[[ag]]
    cfg$cv_locked <- !isTRUE(cfg$cv_locked)
    rv$ag_config[[ag]] <- cfg
  })
  observeEvent(input$pb_pos_lock_toggle, {
    ag <- input$pb_selected_antigen; req(ag)
    .save_ag_cfg(ag)
    cfg <- rv$ag_config[[ag]]
    cfg$pos_locked <- !isTRUE(cfg$pos_locked)
    rv$ag_config[[ag]] <- cfg
  })
  observeEvent(input$pb_neg_lock_toggle, {
    ag <- input$pb_selected_antigen; req(ag)
    .save_ag_cfg(ag)
    cfg <- rv$ag_config[[ag]]
    cfg$neg_locked <- !isTRUE(cfg$neg_locked)
    rv$ag_config[[ag]] <- cfg
  })

  # ---------------------------------------------------------------------------
  # Auto-save the current antigen config whenever the pos/neg ctrl selection
  # changes on the Point-based tab.  This keeps rv$ag_config current so the
  # Titration QC Plot sees the correct neg/pos labels even when the user
  # navigates to the Titration tab without first switching to another antigen
  # (which is the normal trigger for .save_ag_cfg).
  # ---------------------------------------------------------------------------
  observeEvent(input$pb_pos_ctrl, {
    ag <- input$pb_selected_antigen
    if (!is.null(ag) && nchar(trimws(ag)) > 0) .save_ag_cfg(ag)
  }, ignoreNULL = FALSE, ignoreInit = TRUE)

  observeEvent(input$pb_neg_ctrl, {
    ag <- input$pb_selected_antigen
    if (!is.null(ag) && nchar(trimws(ag)) > 0) .save_ag_cfg(ag)
  }, ignoreNULL = FALSE, ignoreInit = TRUE)

  # Render the full config row -- re-renders ONLY on antigen switch, BRILLIANT
  # toggle, a lock/unlock button click, or a change in available control
  # choices. These are listed explicitly below.
  #
  # IMPORTANT: cfg is read via isolate() so this block deliberately does NOT
  # reactively depend on rv$ag_config as a whole. rv$ag_config is a single
  # reactiveValues list keyed by antigen, so *any* write to it (e.g. the
  # auto-save observers below that fire on every input$pb_pos_ctrl /
  # input$pb_neg_ctrl change) would otherwise invalidate this entire block --
  # tearing down and recreating the Positive/Negative Controls pickers on
  # every single click. That is what was throwing the cursor off and making
  # it impossible to smoothly remove more than one control at a time.
  # Antigen switches and lock-button clicks still force a fresh render (via
  # the explicit dependencies below), which is when the saved cfg actually
  # needs to be re-applied to the widgets.
  output$pb_config_row <- renderUI({
    ag           <- input$pb_selected_antigen %||% ""
    input$pb_brilliant_toggle
    input$pb_lod_lock_toggle; input$pb_loq_lock_toggle; input$pb_cv_lock_toggle
    input$pb_pos_lock_toggle; input$pb_neg_lock_toggle
    choices      <- pb_ctrl_choices()
    cfg          <- isolate(.ag_cfg_get(ag))
    is_ix_now    <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    is_bioplex_now  <- !is_ix_now
    is_brilliant_now <- !is_bioplex_now || isTRUE(rv$brilliant_bioplex)

    .lock_btn <- function(id, locked) {
      actionButton(id, label = if (locked) tagList(icon("lock"),   " Locked")
                                      else tagList(icon("unlock"), " Unlocked"),
        class = paste("lock-btn btn-xs", if (locked) "locked" else "unlocked"),
        style = if (locked) "background:#dc3545;color:#fff;border:none;font-size:11px;padding:2px 8px;margin-top:4px;"
                       else "background:#6c757d;color:#fff;border:none;font-size:11px;padding:2px 8px;margin-top:4px;")
    }

    # LOD, LOQ and CV% are freely editable for both BioPlex and INTELLIFLEX
    # (no admin gate); they can still be locked/unlocked per antigen via the
    # buttons below. The only thing that forces a field read-only is a
    # non-BRILLIANT BioPlex run, where LOD/LOQ must stay fixed at 0.
    .free_numeric <- function(id, value, min = 0, max = NA, step = 1, force_readonly = FALSE) {
      ni <- numericInput(id, NULL, value = value, min = min, step = step, width = "100%")
      if (!is.na(max)) ni <- numericInput(id, NULL, value = value, min = min, max = max, step = step, width = "100%")
      if (force_readonly) {
        # Inject readonly attribute via tagAppendAttributes on the input tag
        ni$children[[2]] <- tagAppendAttributes(ni$children[[2]], readonly = "readonly")
      }
      ni
    }

    # LOD/LOQ display values -- fixed at 0 for non-BRILLIANT BioPlex runs
    lod_display <- if (is_brilliant_now) cfg$lod else 0
    loq_display <- if (is_brilliant_now) cfg$loq else 0

    tagList(
      if (!is_brilliant_now) {
        tags$div(
          style = "background:#fdeaea; border:1px solid #e0a0a0; border-radius:6px;
                   padding:5px 12px; margin-bottom:10px; font-size:12px; color:#7a2e2e;",
          icon("info-circle"), " No LOD/LOQ values available -- LOD and LOQ are fixed at 0."
        )
      },
      fluidRow(
        column(2,
          tags$label(style="font-size:13px; font-weight:600;",
            "LOD", tags$span(style="font-size:11px; font-weight:400; color:#555;", " (Limit of Detection)")),
          tags$div(class="lod-loq-wrap",
            .free_numeric("pb_lod", lod_display, min = 0, step = 1, force_readonly = !is_brilliant_now),
            if (is_brilliant_now) .lock_btn("pb_lod_lock_toggle", cfg$lod_locked)
          )
        ),
        column(2,
          tags$label(style="font-size:13px; font-weight:600;",
            "LOQ", tags$span(style="font-size:11px; font-weight:400; color:#555;", " (Limit of Quantification)")),
          tags$div(class="lod-loq-wrap",
            .free_numeric("pb_loq", loq_display, min = 0, step = 1, force_readonly = !is_brilliant_now),
            if (is_brilliant_now) .lock_btn("pb_loq_lock_toggle", cfg$loq_locked)
          )
        ),
        column(2,
          tags$label(style="font-size:13px; font-weight:600;", "CV%",
            tags$span(style="font-size:11px; font-weight:400; color:#555;", " (Acceptance)")),
          tags$div(class="lod-loq-wrap",
            .free_numeric("pb_cv_accept", cfg$cv, min = 1, max = 100, step = 1),
            .lock_btn("pb_cv_lock_toggle", cfg$cv_locked)
          )
        ),
        column(3,
          tags$label(style="font-size:13px; font-weight:600;",
            icon("plus-circle", style="color:#155724;"), " Positive Controls"),
          if (length(choices) == 0) tags$em("No controls.") else
            selectInput("pb_pos_ctrl", NULL, choices = choices,
                        selected = if (!is.null(cfg$pos_ctrl)) cfg$pos_ctrl else NULL,
                        multiple = TRUE, width = "100%"),
          .lock_btn("pb_pos_lock_toggle", cfg$pos_locked)
        ),
        column(3,
          tags$label(style="font-size:13px; font-weight:600;",
            icon("minus-circle", style="color:#721c24;"), " Negative Controls"),
          if (length(choices) == 0) tags$em("No controls.") else
            selectInput("pb_neg_ctrl", NULL, choices = choices,
                        selected = if (!is.null(cfg$neg_ctrl)) cfg$neg_ctrl else NULL,
                        multiple = TRUE, width = "100%"),
          .lock_btn("pb_neg_lock_toggle", cfg$neg_locked)
        )
      )
    )
  })

  # -- Core function: per-well MFI + classification for a given antigen.
  # Parameterized by (ag_col, lod) so it can be called both by the reactive
  # pb_antigen_data() (for the currently-selected antigen) and by the export
  # code, which loops over every antigen to build 08_QC summary.
  .pb_antigen_data_core <- function(ag_col, lod) {
    df     <- rv$raw_df
    is_ix  <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    if (is.null(df) || !ag_col %in% colnames(df) || !"Well" %in% colnames(df)) return(NULL)

    parse_fi <- function(x) {
      x <- gsub("\\s*\\(\\d+\\)\\s*$", "", as.character(x))
      suppressWarnings(as.numeric(gsub(",", ".", x)))
    }

    wells     <- as.character(df[["Well"]])
    well_type <- if ("Type"        %in% colnames(df)) as.character(df[["Type"]])        else rep("S", nrow(df))
    desc      <- if ("Description" %in% colnames(df)) as.character(df[["Description"]]) else wells

    # ---- For INTELLIFLEX: resolve sample names from 96-well layout -----------
    if (is_ix && !is.null(rv$well96_df)) {
      w96      <- rv$well96_df
      row_col  <- colnames(w96)[1]
      num_cols <- colnames(w96)[-1]
      ix_labels <- do.call(rbind, lapply(seq_len(nrow(w96)), function(r) {
        row_ltr <- as.character(w96[[row_col]][r])
        do.call(rbind, lapply(num_cols, function(cc)
          data.frame(Well = paste0(row_ltr, cc),
                     SampleName = as.character(w96[[cc]][r]),
                     stringsAsFactors = FALSE)))
      }))
      ix_map <- setNames(ix_labels$SampleName, ix_labels$Well)
      for (i in seq_along(wells)) {
        nm <- ix_map[[wells[i]]]
        if (!is.null(nm) && !is.na(nm) && nchar(trimws(nm)) > 0) desc[i] <- trimws(nm)
      }
      # Mark control / no_mab wells from Step 3 selections
      ctrl_wells <- trimws(isolate(input$ix_ctrl_wells) %||% character(0))
      ctrl_wells <- ctrl_wells[nchar(ctrl_wells) > 0]
      well_type  <- dplyr::case_when(
        wells %in% ctrl_wells  ~ "C",
        desc  == "no_mab"      ~ "B",
        TRUE                   ~ well_type
      )
    }

    # ---- Blank-bead column: for INTELLIFLEX this is the R44 (blank bead) ----
    blank_col <- if (is_ix) {
      # R44 is the blank bead region; its median column is "R44 (44)"
      stripped <- toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", colnames(df))))
      cols_r44 <- colnames(df)[stripped == "R44"]
      if (length(cols_r44) > 0) cols_r44[1] else NA_character_
    } else {
      cols_bl <- colnames(df)[startsWith(toupper(colnames(df)), "BLANK")]
      if (length(cols_bl) > 0) cols_bl[1] else NA_character_
    }
    blank_vals <- if (!is.na(blank_col)) parse_fi(df[[blank_col]]) else rep(0, nrow(df))
    blank_vals[is.na(blank_vals)] <- 0
    is_nomab <- (well_type == "B")

    # ---- Shared no_mab average (same logic as mfi_dataframe) ----------------
    is_ag_all   <- grepl("^.+\\s*\\(\\d+\\)\\s*$", colnames(df)) &
                   !startsWith(colnames(df), "Beads_")
    ag_cols_all <- colnames(df)[is_ag_all]
    ag_cols_all <- ag_cols_all[!startsWith(toupper(ag_cols_all), "BLANK")]
    ag_cols_all <- ag_cols_all[!toupper(trimws(gsub("\\s*\\(\\d+\\)\\s*$",
                                                     "", ag_cols_all))) %in% "R44"]
    avg_nomab_shared <- if (length(ag_cols_all) > 0 && any(is_nomab, na.rm = TRUE)) {
      raw_first <- parse_fi(df[[ag_cols_all[1]]])
      ntz       <- pmax(raw_first - blank_vals, 0, na.rm = TRUE); ntz[is.na(ntz)] <- 0
      v <- mean(ntz[is_nomab], na.rm = TRUE); if (is.na(v)) 0 else v
    } else 0

    # ---- MFI: blank-bead + no_mab subtracted (mirrors mfi_dataframe $full) --
    raw_mfi <- parse_fi(df[[ag_col]])
    ntz     <- pmax(raw_mfi - blank_vals, 0, na.rm = TRUE); ntz[is.na(ntz)] <- 0
    mfu     <- pmax(ntz - avg_nomab_shared, 0, na.rm = TRUE); mfu[is.na(mfu)] <- 0

    # ---- Bead count: for INTELLIFLEX use R34: RP1 COUNT directly ------------
    # For BioPlex, the bead count is embedded as a trailing "(NN)" in the raw
    # MFI cell itself (e.g. "23028,00 (63)" -> 63 beads) -- parse it out so
    # QC 1 can apply the real >=50-beads criterion instead of falling back
    # to an MFI/LOD proxy.
    parse_bead_ct <- function(col_vec) {
      x <- as.character(col_vec)
      sapply(x, function(v) {
        m <- regmatches(v, regexpr("\\((\\d+)\\)\\s*$", v))
        if (length(m) == 0 || nchar(m) == 0) return(NA_integer_)
        suppressWarnings(as.integer(gsub("[^0-9]", "", m)))
      }, USE.NAMES = FALSE)
    }
    bead_counts <- if (is_ix) {
      # ag_col is like "R34 (34)"; strip "(34)" -> "R34"; find "Beads_R34 (34)"
      Beads_col <- paste0("Beads_", ag_col)
      if (Beads_col %in% colnames(df))
        suppressWarnings(as.numeric(df[[Beads_col]]))
      else rep(NA_real_, nrow(df))
    } else {
      parse_bead_ct(df[[ag_col]])
    }

    readout    <- ifelse(mfu > lod, "Positive", "Negative")
    type_label <- dplyr::case_when(
      well_type == "B"             ~ "Blank (no_mab)",
      startsWith(well_type, "C")   ~ paste0("Control (", well_type, ")"),
      startsWith(well_type, "S")   ~ paste0("Standard (", well_type, ")"),
      startsWith(well_type, "X") |
        startsWith(well_type, "U") ~ "Sample",
      TRUE                         ~ well_type
    )

    data.frame(Well = wells, WellType = well_type, TypeLabel = type_label,
               Sample_ID = desc, MFIs = round(mfu, 1), Readout = readout,
               BeadCount = bead_counts, is_nomab = is_nomab,
               stringsAsFactors = FALSE)
  }

  # -- Core reactive: per-well MFI + classification for the selected antigen --
  pb_antigen_data <- reactive({
    req(rv$raw_df, input$pb_selected_antigen)
    .pb_antigen_data_core(input$pb_selected_antigen, pb_cfg()$lod)
  })

  # -- Per-well table ---------------------------------------------------------
  output$pb_well_table <- renderDT({
    d <- pb_antigen_data(); req(d)
    display <- d[, c("Well","TypeLabel","Sample_ID","MFIs","Readout")]
    colnames(display) <- c("Well","Type","Sample_ID","MFIs","Readout")
    datatable(display,
      options = list(pageLength = 15, scrollX = TRUE, dom = "lfrtip",
                     columnDefs = list(list(className = "dt-center", targets = "_all"))),
      rownames = FALSE, class = "stripe hover cell-border compact"
    ) %>%
      formatRound("MFIs", 1) %>%
      formatStyle("MFIs",
        backgroundColor = styleInterval(pb_cfg()$lod,
                                        c("#f8d7da", "#d4edda")),
        color           = styleInterval(pb_cfg()$lod,
                                        c("#721c24", "#155724")),
        fontWeight = "bold") %>%
      formatStyle("Readout",
        backgroundColor = styleEqual(c("Positive","Negative"), c("#d4edda","#f8d7da")),
        color           = styleEqual(c("Positive","Negative"), c("#155724","#721c24")),
        fontWeight = "bold")
  })

  # -- Plate map (Positive/Negative readout) ----------------------------------
  output$pb_plate_map <- renderPlot({
    d <- pb_antigen_data(); req(d)
    dims_pb   <- plate_dims()
    all_wells <- paste0(rep(dims_pb$row_letters, each = dims_pb$n_cols),
                        rep(seq_len(dims_pb$n_cols), dims_pb$n_rows))
    base <- data.frame(
      Well = all_wells,
      Row  = factor(substr(all_wells, 1, 1), levels = rev(dims_pb$row_letters)),
      Col  = factor(as.integer(sub("[A-Za-z]", "", all_wells)), levels = seq_len(dims_pb$n_cols)),
      stringsAsFactors = FALSE
    )
    grid <- merge(base, d[, c("Well","WellType","Readout","Sample_ID","MFIs","is_nomab")],
                  by = "Well", all.x = TRUE)
    grid$Readout  <- ifelse(is.na(grid$Readout), "Empty", grid$Readout)
    grid$is_nomab <- ifelse(is.na(grid$is_nomab), FALSE, grid$is_nomab)
    grid$FillColor <- dplyr::case_when(
      grid$is_nomab              ~ "Blank",
      grid$Readout == "Positive" ~ "Positive",
      grid$Readout == "Negative" ~ "Negative",
      TRUE                       ~ "Empty"
    )
    grid$Label   <- ifelse(is.na(grid$Sample_ID), "", substr(as.character(grid$Sample_ID), 1, 8))
    grid$txt_col <- ifelse(grid$FillColor %in% c("Positive","Negative","Blank"), "white", "black")
    is_ix  <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    ag_raw <- input$pb_selected_antigen %||% ""
    ag_lbl <- if (is_ix) {
      df   <- rv$raw_df
      base <- trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_raw))
      acol <- paste0("AnalyteName_", base)
      if (!is.null(df) && acol %in% colnames(df)) {
        v <- as.character(df[[acol]]); v <- v[!is.na(v) & nchar(trimws(v)) > 0]
        if (length(v) > 0) trimws(v[1]) else base
      } else base
    } else gsub("\\s*\\(\\d+\\)\\s*$", "", ag_raw)
    lod_v  <- as.integer(pb_cfg()$lod)
    ggplot(grid, aes(x = Col, y = Row, fill = FillColor)) +
      geom_tile(color = "white", linewidth = 0.6) +
      geom_text(aes(label = Label, color = txt_col), size = 2.1, fontface = "bold") +
      scale_color_identity() +
      annotate("rect",
               xmin = 0.5, xmax = dims_pb$n_cols + 0.5,
               ymin = 0.5, ymax = dims_pb$n_rows + 0.5,
               fill = NA, color = "black", linewidth = 1.5) +
      scale_fill_manual(
        name   = "Readout",
        values = c("Positive" = "#27ae60", "Negative" = "#e53935",
                   "Blank"    = "#8B0000",  "Empty"    = "#d0d0d0"),
        labels = c(
          "Positive" = paste0("Positive (MFIs > ", lod_v, ")"),
          "Negative" = paste0("Negative (MFIs \u2264 ", lod_v, ")"),
          "Blank"    = "Blank / no_mab", "Empty" = "Empty")
      ) +
      labs(title = paste0("Plate Readout -- ", ag_lbl), x = "Column", y = "Row") +
      theme_minimal(base_size = 11) +
      theme(panel.grid = element_blank(), axis.text = element_text(size = 9),
            plot.title = element_text(face = "bold", size = 12, hjust = 0.5),
            legend.title = element_text(size = 9), legend.text = element_text(size = 8),
            panel.border = element_blank())
  })

  # -- Helper: PASS/CAUTION badge --------------------------------------------
  .qc_badge <- function(pass, summary_text = NULL) {
    color <- if (isTRUE(pass)) "#27ae60" else if (is.na(pass)) "#888" else "#e6a817"
    txt_color <- if (isTRUE(pass) || is.na(pass)) "#fff" else "#000"
    label <- if (isTRUE(pass)) "PASS" else if (is.na(pass)) "N/A" else "CAUTION"
    tags$div(
      style = "display:flex; align-items:center; gap:16px; margin-bottom:10px;",
      tags$div(
        style = paste0("background:", color, "; color:", txt_color, "; border-radius:8px;
                        padding:10px 24px; font-size:18px; font-weight:700;
                        min-width:110px; text-align:center;"),
        label
      ),
      if (!is.null(summary_text))
        tags$div(style = "font-size:13px; color:#333;", summary_text)
    )
  }

  # -- Antigen label banner shown above QC detail sub-tabs --------------------
  output$pb_qc_antigen_label <- renderUI({
    ag_raw <- input$pb_selected_antigen
    if (is.null(ag_raw) || nchar(trimws(ag_raw)) == 0) return(NULL)
    ag_name <- gsub("\\s*\\(\\d+\\)\\s*$", "", ag_raw)
    lod_v   <- pb_cfg()$lod
    tags$div(
      style = "background:#e3f2fd; border-left:4px solid #1e88e5; border-radius:4px;
               padding:8px 14px; font-size:13px; color:#0d3c61; margin-bottom:12px;
               display:flex; align-items:center; gap:10px;",
      icon("tag"),
      tags$span(tags$strong("Selected Antigen: "), ag_name),
      tags$span(style="margin-left:14px;", tags$strong("LOD: "), lod_v),
      tags$span(
        style="margin-left:14px; font-size:12px; color:#555; font-style:italic;",
        "QC checks below apply to this antigen selection."
      )
    )
  })

  # Helper: render a standard QC DT
  .qc_dt <- function(df, lod = NULL) {
    dt <- datatable(df, rownames = FALSE, class = "stripe compact cell-border",
      options = list(dom = "t", pageLength = 30,
        columnDefs = list(list(className = "dt-center", targets = "_all")))
    )
    if ("MFIs" %in% colnames(df)) {
      dt <- formatRound(dt, "MFIs", 1)
      if (!is.null(lod)) {
        dt <- formatStyle(dt, "MFIs",
          backgroundColor = styleInterval(lod, c("#f8d7da", "#d4edda")),
          color           = styleInterval(lod, c("#721c24", "#155724")),
          fontWeight = "bold")
      }
    }
    dt %>% formatStyle("Met",
      backgroundColor = styleEqual(c(TRUE, FALSE), c("#d4edda", "#f8d7da")),
      color           = styleEqual(c(TRUE, FALSE), c("#155724", "#721c24")),
      fontWeight = "bold")
  }

  # -- QC 1: Bead Acquisition -- >= Bead Count Threshold beads per well -------
  .pb_qc1_core <- function(d, lod, ag_col, bead_thresh = 50) {
    if (is.null(d) || nrow(d) == 0) return(list(pass = NA, df = NULL, msg = "No data available."))
    if (length(bead_thresh) == 0 || is.na(bead_thresh) || !is.finite(bead_thresh)) bead_thresh <- 50
    # Resolve display name for the antigen
    is_ix   <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"
    ag_name <- if (is_ix) {
      df   <- rv$raw_df
      base <- trimws(gsub("\\s*\\(\\d+\\)\\s*$", "", ag_col))
      acol <- paste0("AnalyteName_", base)
      if (!is.null(df) && acol %in% colnames(df)) {
        v <- as.character(df[[acol]])
        v <- v[!is.na(v) & nchar(trimws(v)) > 0]
        if (length(v) > 0) trimws(v[1]) else base
      } else base
    } else gsub("\\s*\\(\\d+\\)\\s*$", "", ag_col)

    # INTELLIFLEX: pb_antigen_data now includes BeadCount from R<N>: RP1 COUNT
    # BioPlex:     BeadCount parsed from the trailing "(NN)" in the raw MFI cell
    if ("BeadCount" %in% colnames(d) && any(!is.na(d$BeadCount))) {
      bead_df <- d[, c("Well", "WellType", "Sample_ID", "BeadCount"), drop = FALSE]
      bead_df$Criterion <- paste0("\u2265 ", bead_thresh, " beads per well (", ag_name, ")")
      bead_df$Met       <- !is.na(bead_df$BeadCount) & bead_df$BeadCount >= bead_thresh
      list(
        pass = all(bead_df$Met, na.rm = TRUE),
        df   = bead_df,
        msg  = sprintf("%d/%d wells \u2265 %s beads for %s",
                       sum(bead_df$Met, na.rm = TRUE), nrow(bead_df), bead_thresh, ag_name)
      )
    } else {
      # BioPlex fallback: no bead-count column. Check that sample wells pass MFI > LOD.
      # QC 1 criterion: sample (non-blank, non-control) wells should have signal > LOD,
      # indicating the assay acquired enough beads / signal for reliable readout.
      rows <- d[!d$is_nomab &
                !startsWith(as.character(d$WellType), "C") &
                d$WellType != "B", , drop = FALSE]
      if (nrow(rows) == 0)
        return(list(pass = NA, df = NULL, msg = "No sample wells found."))
      rows$Criterion <- paste0("MFIs > LOD (", lod, ")")
      rows$Met       <- rows$MFIs > lod
      list(
        pass = all(rows$Met, na.rm = TRUE),
        df   = rows[, c("Well", "Sample_ID", "MFIs", "Criterion", "Met")],
        msg  = sprintf("%d/%d sample wells pass (Signal > LOD=%g)",
                       sum(rows$Met, na.rm = TRUE), nrow(rows), lod)
      )
    }
  }
  pb_qc1 <- reactive({
    d   <- pb_antigen_data(); req(d)
    lod <- pb_cfg()$lod
    ag_col <- input$pb_selected_antigen %||% ""
    bead_thresh <- suppressWarnings(as.numeric(input$pb_bead_count_threshold))
    .pb_qc1_core(d, lod, ag_col, bead_thresh = bead_thresh)
  })
  output$pb_qc1_result <- renderUI({ r <- pb_qc1(); .qc_badge(r$pass, r$msg) })

  # -- QC 1: All-samples bead count table (all analytes, all wells) -----------
  # % Agg Beads is a well-level acquisition-gating metric:
  #   BioPlex:     read directly from the raw file's single "% Agg Beads"
  #                column -- shown here as ONE column aligned by well, exactly
  #                as it appears on the raw file (not repeated per antigen).
  #   INTELLIFLEX: derived as (1 - TOTAL GATED EVENTS/TOTAL EVENTS) * 100,
  #                also well-level, but shown alongside each antigen's bead
  #                count column for easy side-by-side QC review.
  output$pb_qc1_all_table <- renderDT({
    mfi_lst <- mfi_dataframe(); req(!is.null(mfi_lst))
    df_full  <- mfi_lst$full;  req(!is.null(df_full))
    is_ix    <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    # Extract Beads_ and matching AggPct_ columns + well identifiers
    Beads_cols <- grep("^Beads_", colnames(df_full), value = TRUE)
    if (length(Beads_cols) == 0) return(NULL)
    ag_names    <- sub("^Beads_", "", Beads_cols)
    AggPct_cols <- paste0("AggPct_", ag_names)

    fixed_cols <- c("Well", "Type", "Sample_ID")
    have_agg   <- AggPct_cols %in% colnames(df_full)
    beads_disp <- paste0(ag_names, " Beads")

    # % Agg Beads is a well-level metric -- the same value applies to every
    # antigen (both BioPlex and INTELLIFLEX), so it is shown only once,
    # aligned by well, taken straight from the raw file / derived value --
    # placed right after the fixed identifiers.
    agg_disp <- "% Agg Beads"
    src_order  <- c(fixed_cols, if (any(have_agg)) AggPct_cols[which(have_agg)[1]] else NULL, Beads_cols)
    disp_names <- c(fixed_cols, if (any(have_agg)) agg_disp else NULL, beads_disp)
    bead_df <- df_full[, src_order, drop = FALSE]
    colnames(bead_df) <- disp_names

    # % Agg flag threshold (user-configurable input box above the table)
    thresh <- suppressWarnings(as.numeric(input$pb_agg_threshold))
    if (length(thresh) == 0 || is.na(thresh) || !is.finite(thresh)) thresh <- 1e9

    # Bead count pass/fail threshold (user-configurable input box above the table)
    bead_thresh <- suppressWarnings(as.numeric(input$pb_bead_count_threshold))
    if (length(bead_thresh) == 0 || is.na(bead_thresh) || !is.finite(bead_thresh)) bead_thresh <- 50

    dt <- datatable(
      bead_df,
      rownames = FALSE,
      class    = "stripe hover cell-border compact",
      options  = list(
        pageLength = 96, scrollX = TRUE, dom = "lfrtip",
        columnDefs = list(list(className = "dt-center", targets = "_all"))
      )
    )

    agg_disp_present <- agg_disp[agg_disp %in% colnames(bead_df)]
    if (length(agg_disp_present) > 0) dt <- dt %>% formatRound(columns = agg_disp_present, digits = 2)

    # Colour each bead count column: green if >= threshold, red if < threshold
    for (col in beads_disp) {
      if (col %in% colnames(bead_df)) {
        dt <- dt %>%
          formatStyle(
            col,
            backgroundColor = styleInterval(bead_thresh - 0.001, c("#f8d7da", "#d4edda")),
            color           = styleInterval(bead_thresh - 0.001, c("#721c24", "#155724")),
            fontWeight      = "bold"
          )
      }
    }

    # Flag % Agg Beads columns: highlight any value BELOW the threshold
    for (col in agg_disp_present) {
      dt <- dt %>%
        formatStyle(
          col,
          backgroundColor = styleInterval(thresh, c("#fff3cd", "#ffffff")),
          color           = styleInterval(thresh, c("#856404", "#333333")),
          fontWeight      = styleInterval(thresh, c("bold", "normal"))
        )
    }
    dt
  })

  # -- QC 2: Negative Controls -- selected C-type controls, MFI <= LOD ---------
  .pb_qc2_core <- function(d, lod, neg_sel) {
    if (is.null(d) || nrow(d) == 0) return(list(pass = NA, df = NULL, msg = "No data available."))
    if (is.null(neg_sel) || length(neg_sel) == 0)
      return(list(pass = NA, df = NULL, msg = "No negative controls selected."))
    rows <- d[d$Sample_ID %in% neg_sel, , drop = FALSE]
    if (nrow(rows) == 0)
      return(list(pass = NA, df = NULL, msg = paste0("No wells found for selected negative controls: ", paste(neg_sel, collapse=", "))))
    rows$Criterion <- paste0("MFI \u2264 LOD (", lod, ")")
    rows$Met       <- rows$MFIs <= lod
    list(pass = all(rows$Met, na.rm = TRUE),
         df   = rows[, c("Well","WellType","Sample_ID","MFIs","Readout","Criterion","Met")],
         msg  = sprintf("%d/%d negative control wells \u2264 LOD=%g",
                        sum(rows$Met, na.rm=TRUE), nrow(rows), lod))
  }
  pb_qc2 <- reactive({
    d <- pb_antigen_data(); req(d)
    .pb_qc2_core(d, pb_cfg()$lod, pb_cfg()$neg_ctrl)
  })
  output$pb_qc2_result <- renderUI({ r <- pb_qc2(); .qc_badge(r$pass, r$msg) })
  output$pb_qc2_table  <- renderDT({ r <- pb_qc2(); req(!is.null(r$df)); .qc_dt(r$df, lod = pb_cfg()$lod) })

  # -- QC 3: Positive Controls -- selected C-type controls, MFU > LOD, CV < threshold
  .pb_qc3_core <- function(d, lod, cv_accept, pos_sel) {
    if (is.null(d) || nrow(d) == 0) return(list(pass = NA, df = NULL, msg = "No data available."))
    if (is.null(pos_sel) || length(pos_sel) == 0)
      return(list(pass = NA, df = NULL, msg = "No positive controls selected."))
    ctrl <- d[d$Sample_ID %in% pos_sel, , drop = FALSE]
    if (nrow(ctrl) == 0)
      return(list(pass = NA, df = NULL, msg = paste0("No wells found for selected positive controls: ", paste(pos_sel, collapse=", "))))
    ctrl$Criterion <- paste0("MFI > LOD (", lod, ") & CV < ", cv_accept, "%")
    ctrl$Met       <- ctrl$MFIs > lod
    # Compute per-group CV% and map back to each well row
    grp_cvs <- sapply(split(ctrl$MFIs, ctrl$Sample_ID), function(x) {
      m <- mean(x, na.rm = TRUE)
      if (is.na(m) || m == 0 || length(x) < 2) return(NA_real_)
      round(sd(x, na.rm = TRUE) / m * 100, 2)
    })
    ctrl$CV_pct <- grp_cvs[ctrl$Sample_ID]
    ctrl_cv <- all(is.na(grp_cvs) | grp_cvs < cv_accept, na.rm = TRUE)
    list(pass = any(ctrl$Met, na.rm = TRUE) && isTRUE(ctrl_cv),
         df   = ctrl[, c("Well","WellType","Sample_ID","MFIs","CV_pct","Readout","Criterion","Met")],
         msg  = sprintf("%d/%d positive control wells > LOD=%g",
                        sum(ctrl$Met, na.rm=TRUE), nrow(ctrl), lod))
  }
  pb_qc3 <- reactive({
    d <- pb_antigen_data(); req(d)
    .pb_qc3_core(d, pb_cfg()$lod, pb_cfg()$cv, pb_cfg()$pos_ctrl)
  })
  output$pb_qc3_result <- renderUI({ r <- pb_qc3(); .qc_badge(r$pass, r$msg) })
  output$pb_qc3_table  <- renderDT({
    r <- pb_qc3(); req(!is.null(r$df))
    cv_accept <- pb_cfg()$cv
    lod       <- pb_cfg()$lod
    df <- r$df
    dt <- datatable(df, rownames = FALSE, class = "stripe compact cell-border",
      options = list(dom = "t", pageLength = 30,
        columnDefs = list(list(className = "dt-center", targets = "_all")))
    )
    if ("MFIs" %in% colnames(df)) {
      dt <- formatRound(dt, "MFIs", 1) %>%
        formatStyle("MFIs",
          backgroundColor = styleInterval(lod, c("#f8d7da", "#d4edda")),
          color           = styleInterval(lod, c("#721c24", "#155724")),
          fontWeight = "bold")
    }
    if ("CV_pct" %in% colnames(df)) {
      dt <- formatRound(dt, "CV_pct", 2) %>%
        formatStyle("CV_pct",
          color      = styleInterval(cv_accept, c("#155724", "#c0392b")),
          fontWeight = "bold")
    }
    dt %>% formatStyle("Met",
      backgroundColor = styleEqual(c(TRUE, FALSE), c("#d4edda", "#f8d7da")),
      color           = styleEqual(c(TRUE, FALSE), c("#155724", "#721c24")),
      fontWeight = "bold")
  })


  # (pb_qc5 and pb_qc6 retained as internal helpers for future use but not shown in UI)

  # -- QC 5: Replicate Precision (Intra-assay CV%) ----------------------------
  pb_qc5 <- reactive({
    d <- pb_antigen_data(); req(d)
    cv_accept <- pb_cfg()$cv
    samp    <- d[!d$is_nomab & !startsWith(d$WellType, "S"), , drop = FALSE]
    dup_ids <- names(which(table(samp$Sample_ID) > 1))
    if (length(dup_ids) == 0)
      return(list(pass = NA, df = NULL, msg = "No replicate wells found."))
    cv_df <- do.call(rbind, lapply(dup_ids, function(sid) {
      rows <- samp[samp$Sample_ID == sid, , drop = FALSE]
      m <- mean(rows$MFIs, na.rm = TRUE); s <- sd(rows$MFIs, na.rm = TRUE)
      cv <- if (!is.na(m) && m != 0) round(s / m * 100, 2) else NA_real_
      data.frame(Sample_ID = sid, N_Reps = nrow(rows),
                 Mean_MFIs = round(m, 1), SD_MFIs = round(s, 1),
                 CV_pct = cv,
                 Criterion = paste0("CV% <= ", cv_accept, "%"),
                 Met = !is.na(cv) && cv <= cv_accept,
                 stringsAsFactors = FALSE)
    }))
    list(pass = all(cv_df$Met, na.rm = TRUE), df = cv_df,
         msg = sprintf("%d/%d sample groups meet CV<=%g%%",
                       sum(cv_df$Met, na.rm=TRUE), nrow(cv_df), cv_accept))
  })
  output$pb_qc5_result <- renderUI({ r <- pb_qc5(); .qc_badge(r$pass, r$msg) })
  output$pb_qc5_table  <- renderDT({
    r <- pb_qc5(); req(!is.null(r$df))
    cv <- pb_cfg()$cv
    datatable(r$df, rownames = FALSE, class = "stripe compact cell-border",
      options = list(dom = "t", pageLength = 30,
        columnDefs = list(list(className = "dt-center", targets = "_all")))
    ) %>%
      formatRound(c("Mean_MFIs","SD_MFIs","CV_pct"), 2) %>%
      formatStyle("Met",
        backgroundColor = styleEqual(c(TRUE,FALSE), c("#d4edda","#f8d7da")),
        color           = styleEqual(c(TRUE,FALSE), c("#155724","#721c24")),
        fontWeight = "bold") %>%
      formatStyle("CV_pct",
        backgroundColor = styleInterval(c(cv, cv * 1.5),
          c("transparent","#fff3cd","#f8d7da")))
  })

  # -- QC 6: Inter-assay Precision (control CV% on current plate) -------------
  pb_qc6 <- reactive({
    d <- pb_antigen_data(); req(d)
    cv_accept <- pb_cfg()$cv
    ctrl <- d[startsWith(d$WellType, "C") & !d$is_nomab, , drop = FALSE]
    if (nrow(ctrl) == 0)
      return(list(pass = NA, df = NULL, msg = "No control wells found."))
    cv_df <- do.call(rbind, lapply(unique(ctrl$WellType), function(gid) {
      rows <- ctrl[ctrl$WellType == gid, , drop = FALSE]
      m <- mean(rows$MFIs, na.rm = TRUE); s <- sd(rows$MFIs, na.rm = TRUE)
      cv <- if (!is.na(m) && m != 0) round(s / m * 100, 2) else NA_real_
      data.frame(Control_Group = gid, N_Wells = nrow(rows),
                 Mean_MFIs = round(m, 1), SD_MFIs = round(s, 1),
                 CV_pct = cv,
                 Criterion = paste0("CV% <= ", cv_accept, "%"),
                 Met = !is.na(cv) && cv <= cv_accept,
                 stringsAsFactors = FALSE)
    }))
    list(pass = all(cv_df$Met, na.rm = TRUE), df = cv_df,
         msg = sprintf("%d/%d control groups meet CV<=%g%%",
                       sum(cv_df$Met, na.rm=TRUE), nrow(cv_df), cv_accept))
  })
  output$pb_qc6_result <- renderUI({ r <- pb_qc6(); .qc_badge(r$pass, r$msg) })
  output$pb_qc6_table  <- renderDT({
    r <- pb_qc6(); req(!is.null(r$df))
    datatable(r$df, rownames = FALSE, class = "stripe compact cell-border",
      options = list(dom = "t", pageLength = 30,
        columnDefs = list(list(className = "dt-center", targets = "_all")))
    ) %>%
      formatRound(c("Mean_MFIs","SD_MFIs","CV_pct"), 2) %>%
      formatStyle("Met",
        backgroundColor = styleEqual(c(TRUE,FALSE), c("#d4edda","#f8d7da")),
        color           = styleEqual(c(TRUE,FALSE), c("#155724","#721c24")),
        fontWeight = "bold")
  })

  # -- QC 4: QC SUMMARY (Overall) ------------------------------------
  pb_qc4 <- reactive({
    q1 <- pb_qc1(); q2 <- pb_qc2(); q3 <- pb_qc3()
    results <- list(
      "Bead Acquisition"   = q1$pass,
      "Negative Controls"  = q2$pass,
      "Positive Controls"  = q3$pass
    )
    evaluated <- results[!sapply(results, is.na)]
    pass <- if (length(evaluated) == 0) NA else all(unlist(evaluated))
    list(pass = pass, results = results)
  })

  # ===========================================================================
  # TITRATION ANALYSIS  (shared base data + Titration & AUC sub-tabs)
  # ===========================================================================

  # ---------------------------------------------------------------------------
  # Shared utility: build long-form data set (well x analyte) from processed
  # MFI joined with helper concentration / sample_type / control columns.
  # ---------------------------------------------------------------------------
  # .tit_base_data_core() takes the already-computed MFI data frame (one of
  # mfi_dataframe()$mabs / $blank / $full) and does all the helper-metadata
  # joining / x-value computation. Factored out (rather than baked into the
  # tit_base_data reactive below) so the Quantification tab's standard curve
  # can rebuild this same long-form data set against whichever subtraction
  # method the user has toggled to, without touching Titration/AUC, which
  # always use the fully background-corrected ("full") values.
  .tit_base_data_core <- function(df_full) {
    req(!is.null(df_full))
    helper   <- rv$helper_edited
    is_ix    <- !is.null(rv$instrument) && rv$instrument == "INTELLIFLEX"

    parse_fi <- function(x) suppressWarnings(as.numeric(gsub(",", ".", as.character(x))))

    fixed_cols  <- c("Well", "Type", "Sample_ID")
    Beads_cols  <- grep("^Beads_", colnames(df_full), value = TRUE)
    AggPct_cols <- grep("^AggPct_", colnames(df_full), value = TRUE)
    ag_cols     <- setdiff(colnames(df_full), c(fixed_cols, Beads_cols, AggPct_cols))
    if (length(ag_cols) == 0) return(NULL)

    df_work <- df_full[, c(fixed_cols, ag_cols), drop = FALSE]

    # Initialise metadata columns
    df_work$sample_type    <- NA_character_
    df_work$x_value        <- NA_real_
    df_work$control        <- NA_character_
    df_work$sample_kind    <- if (is_ix) "mab" else NA_character_
    df_work$base_sample_id <- NA_character_
    df_work$start_conc     <- NA_real_   # carried through for INTELLIFLEX step-based calc below
    df_work$dil_fac        <- NA_real_   # carried through for INTELLIFLEX step-based calc below

    # -------------------------------------------------------------------------
    # Attach helper metadata via plate_range expansion.
    #
    # Design:
    #   Each helper row covers ONE dilution step for ONE sample.
    #   The sample_id has the form "FH1_1", "FH1_2", ... where the _N suffix
    #   identifies the dilution step number for that base sample ("FH1").
    #   start_concentration and dilution_factor are filled in the helper for
    #   EACH row (dilution step), so the x-value is computed directly:
    #     mab   : x = start_concentration / dilution_factor   (one step down)
    #     serum : x = start_concentration * dilution_factor   (one step up)
    #   If dilution_factor is blank/NA we fall back to start_concentration alone.
    # -------------------------------------------------------------------------
    if (!is.null(helper)) {
      h    <- helper
      h_lc <- setNames(h, tolower(trimws(colnames(h))))
      hc   <- colnames(h_lc)

      pr_col        <- grep("^plate_range$",         hc, value = TRUE)[1]
      st_col_h      <- grep("^sample_type$",         hc, value = TRUE)[1]
      startconc_col <- grep("^start_concentration$", hc, value = TRUE)[1]
      dilfac_col    <- grep("^dilution_factor$",     hc, value = TRUE)[1]
      samples_col   <- grep("^samples$",             hc, value = TRUE)[1]
      ctrl_col_h    <- grep("^control$",             hc, value = TRUE)[1]
      sid_col_h     <- grep("^sample_id$",           hc, value = TRUE)[1]

      # ------------------------------------------------------------------
      # Expand plate_range strings (e.g. "A1:B2", "A11:F11, A12:F12")
      # into individual well labels.
      # ------------------------------------------------------------------
      expand_range <- function(range_str) {
        range_str <- trimws(as.character(range_str))
        if (is.na(range_str) || nchar(range_str) == 0) return(character(0))
        segments  <- trimws(unlist(strsplit(range_str, ",")))
        wells_out <- character(0)
        for (seg in segments) {
          seg <- trimws(seg)
          if (grepl(":", seg)) {
            parts <- strsplit(seg, ":")[[1]]
            if (length(parts) != 2) next
            r1 <- toupper(gsub("[0-9]",  "", trimws(parts[1])))
            c1 <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", trimws(parts[1]))))
            r2 <- toupper(gsub("[0-9]",  "", trimws(parts[2])))
            c2 <- suppressWarnings(as.integer(gsub("[A-Za-z]", "", trimws(parts[2]))))
            if (is.na(c1) || is.na(c2)) next
            ri1 <- match(r1, LETTERS); ri2 <- match(r2, LETTERS)
            if (is.na(ri1) || is.na(ri2)) next
            rows <- LETTERS[ri1:ri2]
            cols <- c1:c2
            wells_out <- c(wells_out, as.vector(outer(rows, cols, paste0)))
          } else {
            if (nchar(seg) > 0) wells_out <- c(wells_out, seg)
          }
        }
        wells_out
      }

      # Build well -> helper-row index map (last match wins if overlap)
      well_to_hrow <- list()
      if (!is.na(pr_col)) {
        for (i in seq_len(nrow(h_lc))) {
          for (w in expand_range(h_lc[[pr_col]][i]))
            well_to_hrow[[w]] <- i
        }
      }

      # Assign per-well metadata from helper
      for (i in seq_len(nrow(df_work))) {
        w     <- as.character(df_work$Well[i])
        h_row <- well_to_hrow[[w]]
        if (is.null(h_row)) next

        # sample_type
        if (!is.na(st_col_h))
          df_work$sample_type[i] <- as.character(h_lc[[st_col_h]][h_row])

        # sample kind: "mab" or "serum"
        sk_val <- if (!is.na(samples_col))
          tolower(trimws(as.character(h_lc[[samples_col]][h_row])))
        else ""
        df_work$sample_kind[i] <- sk_val

        # control label
        if (!is.na(ctrl_col_h))
          df_work$control[i] <- as.character(h_lc[[ctrl_col_h]][h_row])

        # base_sample_id: the helper sample_id is already the base sample name
        # (e.g. "VMI_02", "VMI_03", "Standard_S1_100").
        # Only strip a _S\d+ suffix (standard_curve naming convention).
        # Do NOT strip _\d+ from plain sample names like VMI_02 / VMI_03 --
        # that would collapse all samples into the same base ID ("VMI").
        # We detect whether the helper sample_id itself carries a dilution-step
        # suffix by checking whether the well-level Description has an extra
        # digit suffix beyond the helper sample_id; if so, the helper sample_id
        # is already the base and should be used as-is.
        raw_sid <- if (!is.na(sid_col_h))
          as.character(h_lc[[sid_col_h]][h_row])
        else as.character(df_work$Sample_ID[i])
        df_work$base_sample_id[i] <- {
          # Standard curve samples: strip _S\d+ suffix (e.g. "Standard_S1_100" -> "Standard")
          stripped_s <- sub("_[Ss]\\d+.*$", "", raw_sid)
          if (stripped_s != raw_sid) {
            stripped_s
          } else {
            # For non-standard-curve rows: keep raw_sid as-is.
            # The helper sample_id (e.g. "VMI_02") is already the base name.
            # Stripping a trailing _\d+ would wrongly merge VMI_02/VMI_03/VMI_14 -> "VMI".
            raw_sid
          }
        }

        # x_value
        row_stype <- if (!is.na(st_col_h))
          tolower(trimws(as.character(h_lc[[st_col_h]][h_row]))) else ""

        # start_concentration / dilution_factor for this row -- needed both by
        # the BioPlex post-loop calc below and by the INTELLIFLEX post-loop calc
        start_conc <- if (!is.na(startconc_col))
          parse_fi(h_lc[[startconc_col]][h_row]) else NA_real_
        dil_fac    <- if (!is.na(dilfac_col))
          parse_fi(h_lc[[dilfac_col]][h_row])    else NA_real_
        df_work$start_conc[i] <- start_conc
        df_work$dil_fac[i]    <- dil_fac

        if (row_stype == "standard_curve") {
          # Use start_concentration directly for standard curve points
          if (!is.na(start_conc)) df_work$x_value[i] <- start_conc
        } else {
          # Both BioPlex and INTELLIFLEX samples: defer x_value computation to
          # post-loop where we can determine each well's correct ordinal step
          # within its own sample group using well-sort order.
          # This avoids the bug where Type-code numbers (X1, X9, X17 ...) are
          # used as global step indices instead of per-sample step indices.
          df_work$x_value[i] <- NA_real_   # filled in post-loop
        }
      }

      # ---- Post-loop: compute x_value for ALL sample wells (BioPlex and INTELLIFLEX)
      # that still have NA x_value after the per-well loop above.
      # Strategy:
      #   1. Use the trailing digit in Description (e.g. "VMI_02_3" -> step 3)
      #      as the per-sample step number.  This correctly handles plates where
      #      the same dilution step appears in two wells (replicates): both
      #      "VMI_02_1" wells get step=1 and thus the same x_value, so they
      #      are averaged together by group_by(x_value) downstream.
      #   2. If Description carries no step digit, fall back to well-sort order
      #      within the sample group (same logic as INTELLIFLEX).
      {
        well_sort_key <- function(w) {
          col_n <- suppressWarnings(as.integer(gsub("[^0-9]", "", w)))
          row_n <- match(toupper(gsub("[0-9]", "", w)), LETTERS)
          if (is.na(col_n)) col_n <- 999L
          if (is.na(row_n)) row_n <- 999L
          col_n * 100L + row_n
        }
        # Only process sample-type wells with NA x_value
        samp_type_mask <- !is.na(df_work$sample_type) &
                          tolower(trimws(df_work$sample_type)) == "sample" &
                          is.na(df_work$x_value)
        for (bid in unique(na.omit(df_work$base_sample_id[samp_type_mask]))) {
          idx <- which(!is.na(df_work$sample_type) &
                       tolower(trimws(df_work$sample_type)) == "sample" &
                       is.na(df_work$x_value) &
                       !is.na(df_work$base_sample_id) &
                       df_work$base_sample_id == bid)
          if (length(idx) == 0) next

          sk_grp <- tolower(trimws(df_work$sample_kind[idx[1]]))
          sc_grp <- df_work$start_conc[idx[1]]
          df_grp <- df_work$dil_fac[idx[1]]

          # Try to extract step number from Description suffix (e.g. "VMI_02_3" -> 3).
          # Sample_ID holds the raw Description string (set from desc in mfi_dataframe).
          desc_vals <- trimws(as.character(df_work$Sample_ID[idx]))
          step_from_desc <- suppressWarnings(
            as.integer(sub(".*_(\\d+)$", "\\1", desc_vals))
          )
          # Accept desc-derived steps when ALL wells parsed successfully (no NAs).
          # This correctly handles pure replicates (e.g. two "VMI_02_1" wells) where
          # unique(step_from_desc) == 1: both share step 1, so they get the same
          # x_value and are averaged together -- NOT treated as separate dilution steps.
          has_desc_steps <- !any(is.na(step_from_desc)) && length(step_from_desc) > 0L

          if (has_desc_steps) {
            # Use Description-derived step numbers -- preserves replicate grouping
            steps_for_idx <- step_from_desc
          } else {
            # Fallback: assign ordinal steps by well-sort order
            ord        <- order(sapply(as.character(df_work$Well[idx]), well_sort_key))
            steps_for_idx <- integer(length(idx))
            steps_for_idx[ord] <- seq_along(idx)
          }

          if (!is.na(sc_grp) && !is.na(df_grp) && df_grp > 0) {
            df_work$x_value[idx] <- if (sk_grp == "mab") {
              sc_grp / (df_grp ^ (steps_for_idx - 1L))
            } else {
              sc_grp * (df_grp ^ (steps_for_idx - 1L))
            }
          } else {
            # No concentration info: use the step number itself as x_value
            # (results in ordinal axis 1, 2, 3 ... n)
            df_work$x_value[idx] <- steps_for_idx
          }
        }
      }
    } # end if (!is.null(helper))

    # For INTELLIFLEX without a helper, assign base_sample_id from Sample_ID
    # and x_value from well ordinal within each sample group
    if (is_ix && all(is.na(df_work$base_sample_id))) {
      df_work$base_sample_id <- as.character(df_work$Sample_ID)
      well_sort_key <- function(w) {
        col_n <- suppressWarnings(as.integer(gsub("[^0-9]", "", w)))
        row_n <- match(toupper(gsub("[0-9]", "", w)), LETTERS)
        if (is.na(col_n)) col_n <- 999L
        if (is.na(row_n)) row_n <- 999L
        col_n * 100L + row_n
      }
      for (bid in unique(na.omit(df_work$base_sample_id))) {
        idx <- which(!is.na(df_work$base_sample_id) & df_work$base_sample_id == bid &
                     !startsWith(as.character(df_work$Type), "C") &
                     df_work$Type != "B")
        if (length(idx) == 0) next
        wells_here <- as.character(df_work$Well[idx])
        ord        <- order(sapply(wells_here, well_sort_key))
        df_work$x_value[idx[ord]] <- seq_along(idx)
      }
    }

    # Extract "1:N" dilution label from Description (e.g. "VRC01_S1_1:50" -> "1:50").
    df_work$dilution_label <- {
      raw  <- as.character(df_work$Sample_ID)
      full <- rep(NA_character_, length(raw))
      for (k in seq_along(raw)) {
        m <- regexpr("1:[0-9,]+", raw[k])
        if (m > 0) full[k] <- regmatches(raw[k], m)
      }
      full
    }

    # Pivot to long form -- one row per (well x analyte) --------------------
    long_df <- do.call(rbind, lapply(ag_cols, function(ag) {
      data.frame(
        Well           = df_work$Well,
        Type           = df_work$Type,
        Sample_ID      = df_work$Sample_ID,
        base_sample_id = df_work$base_sample_id,
        sample_type    = df_work$sample_type,
        x_value        = df_work$x_value,
        dilution_label = df_work$dilution_label,
        control        = df_work$control,
        sample_kind    = df_work$sample_kind,
        start_conc     = df_work$start_conc,
        dil_fac        = df_work$dil_fac,
        analyte        = ag,
        MFI            = parse_fi(df_work[[ag]]),
        stringsAsFactors = FALSE
      )
    }))


    long_df <- long_df[!is.na(long_df$MFI) & !is.na(long_df$Type) & long_df$Type != "B", , drop = FALSE]

    # ---- Bead-count QC: attach each well's bead count for this analyte
    # (INTELLIFLEX RP1/RP2 COUNT, parsed into Beads_<antigen> columns
    # upstream) for reference/flagging purposes, but do NOT drop rows with
    # low bead counts -- all values (including < 50 beads) are retained so
    # the titration curve reflects the complete data. BioPlex data has no
    # Beads_* columns, so this is a no-op for that instrument. --------------
    if (length(Beads_cols) > 0) {
      Beads_long <- do.call(rbind, lapply(ag_cols, function(ag) {
        Beads_col <- paste0("Beads_", ag)
        bc <- if (Beads_col %in% colnames(df_full))
          suppressWarnings(as.numeric(df_full[[Beads_col]])) else NA_real_
        data.frame(Well = df_full$Well, analyte = ag, BeadCount = bc,
                   stringsAsFactors = FALSE)
      }))
      long_df <- merge(long_df, Beads_long, by = c("Well", "analyte"), all.x = TRUE)
      long_df$BeadCount <- NULL
    }

    # Ensure no NA analyte/Type rows remain (can appear after merge)
    long_df <- long_df[!is.na(long_df$analyte) & !is.na(long_df$Well), , drop = FALSE]

    long_df
  }

  # Titration / AUC tabs: unchanged behavior -- always built from the fully
  # background-corrected ("full") MFI values.
  tit_base_data <- reactive({
    req(rv$analysis_run)
    mfi_list <- mfi_dataframe(); req(!is.null(mfi_list))
    .tit_base_data_core(mfi_list$full)
  })

  # ---------------------------------------------------------------------------
  # TITRATION sub-tab helpers
  # ---------------------------------------------------------------------------

  # Shared filter helper -- applies Analyte / Sample Type / Control selectors
  .tit_apply_filters <- function(d, analyte_in, stype_in, ctrl_in) {
    if (!is.null(analyte_in) && length(analyte_in) == 1 && analyte_in != "All Analytes")
      d <- d[!is.na(d$analyte) & d$analyte == analyte_in, , drop = FALSE]
    if (!is.null(stype_in) && length(stype_in) == 1 && stype_in != "All Types") {
      type_match <- (!is.na(d$sample_type) & d$sample_type == stype_in) |
                    (!is.na(d$Type) & d$Type == stype_in)
      d <- d[type_match, , drop = FALSE]
    }
    if (!is.null(ctrl_in) && length(ctrl_in) == 1 && ctrl_in != "All Controls")
      d <- d[!is.na(d$control) & d$control == ctrl_in, , drop = FALSE]
    d
  }

  # Palette helper -------------------------------------------------------
  .make_pal <- function(n) {
    base_cols <- c("#e41a1c","#377eb8","#4daf4a","#984ea3",
                   "#ff7f00","#a65628","#f781bf","#555555")
    if (n <= length(base_cols)) base_cols[seq_len(n)]
    else colorRampPalette(base_cols)(n)
  }

  # UI selectors ---------------------------------------------------------
  output$tit_analyte_selector_ui <- renderUI({
    d <- tit_base_data(); req(!is.null(d))
    selectInput("tit_analyte", NULL,
                choices  = c("All Analytes", sort(unique(d$analyte))),
                selected = "All Analytes", width = "100%")
  })
  output$tit_control_selector_ui <- renderUI({
    d <- tit_base_data(); req(!is.null(d))
    ctrls <- sort(unique(d$control[!is.na(d$control) &
                                   nchar(trimws(d$control)) > 0 &
                                   d$control != "NA"]))
    selectInput("tit_control", NULL,
                choices  = c("All Controls", ctrls),
                selected = "All Controls", width = "100%")
  })

  # Filtered base -- Analyte and Control only (no sample_type filter)
  tit_filtered_data <- reactive({
    d <- tit_base_data(); req(!is.null(d) && nrow(d) > 0)
    if (!is.null(input$tit_analyte) && length(input$tit_analyte) == 1 &&
        input$tit_analyte != "All Analytes")
      d <- d[!is.na(d$analyte) & d$analyte == input$tit_analyte, , drop = FALSE]
    if (!is.null(input$tit_control) && length(input$tit_control) == 1 &&
        input$tit_control != "All Controls")
      d <- d[!is.na(d$control) & d$control == input$tit_control, , drop = FALSE]
    d
  })

  # Sample selector -- shows base_sample_id as tag-style pills.
  # Each pill is a base sample name (FH1, gl-VRC01, PGT145, VRC01, ...).
  # Selecting/deselecting a pill includes/excludes all its _N dilution rows.
  output$tit_sample_selector_ui <- renderUI({
    d <- tit_filtered_data(); req(!is.null(d))
    non_ctrl <- d[!startsWith(as.character(d$Type), "C"), , drop = FALSE]
    # Use base_sample_id if populated from helper, else strip _N from Sample_ID
    bids <- non_ctrl$base_sample_id
    bids[is.na(bids) | nchar(trimws(bids)) == 0] <- {
      raw <- as.character(non_ctrl$Sample_ID[is.na(bids) | nchar(trimws(bids)) == 0])
      stripped <- sub("_[Ss]\\d+.*$", "", raw)
      ifelse(stripped == raw, sub("_\\d+$", "", raw), stripped)
    }
    base_ids <- sort(unique(bids[nchar(trimws(bids)) > 0]))
    if (length(base_ids) == 0) base_ids <- sort(unique(d$Sample_ID))

    # Render as a flexbox tag-token UI (matches Neut appearance)
    selected_init <- base_ids
    tagList(
      tags$div(
        id    = "tit_sample_tag_container",
        style = paste0(
          "display:flex; flex-wrap:wrap; gap:4px; padding:4px 0;",
          "max-height:200px; overflow-y:auto;"
        ),
        lapply(base_ids, function(sid) {
          tags$div(
            class       = "tit-sample-tag selected",
            `data-sid`  = sid,
            style       = paste0(
              "display:inline-flex; align-items:center; gap:4px;",
              "background:#1a3a5c; color:#fff; border-radius:4px;",
              "padding:2px 8px; font-size:11px; cursor:pointer;",
              "user-select:none; white-space:nowrap;"
            ),
            tags$span(sid)
          )
        })
      ),
      # Hidden input updated by JS -- holds comma-separated selected base_ids
      tags$input(
        type  = "text",
        id    = "tit_selected_samples",
        style = "display:none;",
        value = paste(selected_init, collapse = ",")
      ),
      tags$script(HTML(sprintf("
        (function() {
          var allIds = %s;
          function getSelected() {
            return Array.from(
              document.querySelectorAll('#tit_sample_tag_container .tit-sample-tag.selected')
            ).map(function(el) { return el.getAttribute('data-sid'); });
          }
          function updateInput() {
            var sel = getSelected();
            var inp = document.getElementById('tit_selected_samples');
            if (inp) {
              inp.value = sel.join(',');
              inp.dispatchEvent(new Event('change'));
              Shiny.setInputValue('tit_selected_samples', sel.join(','), {priority: 'event'});
            }
          }
          document.querySelectorAll('#tit_sample_tag_container .tit-sample-tag').forEach(function(tag) {
            tag.addEventListener('click', function() {
              this.classList.toggle('selected');
              if (this.classList.contains('selected')) {
                this.style.background = '#1a3a5c';
                this.style.color      = '#fff';
              } else {
                this.style.background = '#e8edf3';
                this.style.color      = '#1a3a5c';
              }
              updateInput();
            });
          });
        })();
      ", jsonlite::toJSON(base_ids))))
    )
  })

  # Parse the comma-string from the tag UI into a character vector
  tit_selected_base_ids <- reactive({
    raw <- input$tit_selected_samples %||% ""
    ids <- trimws(unlist(strsplit(raw, ",")))
    ids[nchar(ids) > 0]
  })

  # Final aggregated plot data (sample wells only) -----------------------
  # Selection is by base_sample_id (FH1, gl-VRC01, ...).
  # All _N dilution rows for selected bases are included.
  # Rows with the same base_sample_id x analyte x x_value are averaged
  # (handles well-level replicates at the same dilution step).
  # One curve per base_sample_id in the final plot.
  tit_plot_data <- reactive({
    d <- tit_filtered_data(); req(!is.null(d))
    # Exclude standard_curve rows -- those belong on the Quantification tab only
    d <- d[is.na(d$sample_type) | tolower(trimws(d$sample_type)) != "standard_curve", , drop = FALSE]
    d_samp <- d[!startsWith(as.character(d$Type), "C"), , drop = FALSE]

    # Fill base_sample_id fallback
    miss <- is.na(d_samp$base_sample_id) | nchar(trimws(d_samp$base_sample_id)) == 0
    d_samp$base_sample_id[miss] <- {
      raw <- as.character(d_samp$Sample_ID[miss])
      stripped <- sub("_[Ss]\\d+.*$", "", raw)
      ifelse(stripped == raw, sub("_\\d+$", "", raw), stripped)
    }

    sel <- tit_selected_base_ids()
    if (length(sel) > 0)
      d_samp <- d_samp[d_samp$base_sample_id %in% sel, , drop = FALSE]
    if (nrow(d_samp) == 0) return(d_samp)

    as.data.frame(
      d_samp %>%
        group_by(base_sample_id, analyte, Type, x_value, sample_type, sample_kind) %>%
        summarise(avg_MFI = mean(MFI, na.rm = TRUE), n_reps = n(), .groups = "drop") %>%
        arrange(base_sample_id, analyte, x_value)
    )
  })

  # Control plot data (Type starting with C) -----------------------------
  # Averages duplicate wells per Sample_ID x analyte.
  # Tags each Sample_ID as "positive", "negative", or "sample" based on
  # ag_config pos_ctrl / neg_ctrl selections (stored as Type codes).
  # All control sample names: Type starting with C in raw MFI -> Description value
  helper_ctrl_names <- reactive({
    # Primary: raw MFI file Type column (C1, C2 ... = controls)
    df <- rv$raw_df
    if (!is.null(df) && all(c("Type", "Description") %in% colnames(df))) {
      type_vals <- toupper(trimws(as.character(df[["Type"]])))
      ctrl_mask <- grepl("^C\\d*$", type_vals)
      desc_vals <- trimws(as.character(df[["Description"]][ctrl_mask]))
      sids <- unique(desc_vals[!is.na(desc_vals) & nchar(desc_vals) > 0])
      if (length(sids) > 0) return(sids)
    }
    # Fallback: helper sample_type == "control" rows
    h <- rv$helper_edited
    if (is.null(h) || nrow(h) == 0) return(character(0))
    h_lc   <- setNames(h, tolower(trimws(colnames(h))))
    st_col  <- grep("^sample_type$", colnames(h_lc), value = TRUE)[1]
    sid_col <- grep("^sample_id$",   colnames(h_lc), value = TRUE)[1]
    if (is.na(st_col) || is.na(sid_col)) return(character(0))
    ctrl_mask <- tolower(trimws(as.character(h_lc[[st_col]]))) == "control"
    sids <- trimws(as.character(h_lc[[sid_col]][ctrl_mask]))
    unique(sids[!is.na(sids) & nchar(sids) > 0])
  })

  tit_ctrl_data <- reactive({
    d <- tit_filtered_data(); req(!is.null(d))

    # Identify control wells: Type starts with C (C1, C2 ...) OR Description in helper ctrl names
    # OR sample_type == "control" (populated from helper -- covers INTELLIFLEX where
    # control wells are identified via the 96-well layout / Step 3 selector rather
    # than a raw instrument Type code).
    ctrl_names <- helper_ctrl_names()
    is_ctrl <- grepl("^C\\d*$", toupper(trimws(as.character(d$Type)))) |
               (length(ctrl_names) > 0 & as.character(d$Sample_ID) %in% ctrl_names) |
               (!is.na(d$sample_type) & tolower(trimws(as.character(d$sample_type))) == "control")

    # Sample wells: Type starts with X or U (X1..Xn, U1..Un)
    # Also include rows where sample_type == "sample" from the helper (covers
    # INTELLIFLEX wells that may not carry an X/U Type code).
    is_samp <- (grepl("^[XxUu]\\d*$", trimws(as.character(d$Type))) |
                (!is.na(d$sample_type) & tolower(trimws(as.character(d$sample_type))) == "sample")) &
               !is_ctrl

    d_ctrl <- d[is_ctrl, , drop = FALSE]
    d_samp <- d[is_samp,  , drop = FALSE]

    if (nrow(d_ctrl) == 0 && nrow(d_samp) == 0) return(NULL)

    # ---- Controls: raw per-well rows, tagged as positive/negative/unclassified ----
    raw_ctrl <- if (nrow(d_ctrl) > 0) {
      out <- data.frame(
        Well      = d_ctrl$Well,
        Sample_ID = as.character(d_ctrl$Sample_ID),
        analyte   = d_ctrl$analyte,
        MFI       = d_ctrl$MFI,
        Type      = as.character(d_ctrl$Type),
        stringsAsFactors = FALSE
      )
      out$ctrl_category <- "unclassified"

      # Build a display-keyed config map from rv$ag_config (saved state for all
      # antigens).  ag_config keys are raw column names like "R34 (34)";
      # out$analyte holds RESOLVED display names (e.g. "CFp10_SG" for
      # INTELLIFLEX, via AnalyteName_<region>, or the stripped base name for
      # BioPlex). Use the same resolver used to build out$analyte so the keys
      # actually match -- naive suffix-stripping breaks for INTELLIFLEX since
      # the raw region code ("R34") differs from the resolved analyte name.
      ag_cfg_display <- list()
      for (raw_key in names(rv$ag_config)) {
        display_key <- .ag_raw_to_display(raw_key)
        ag_cfg_display[[display_key]] <- rv$ag_config[[raw_key]]
      }

      # The CURRENTLY-selected antigen's live inputs (input$pb_pos_ctrl /
      # input$pb_neg_ctrl) always reflect what the user has chosen in the
      # Point-based QC Configuration panel, but those values are only flushed
      # into rv$ag_config when the user navigates away or clicks a lock button.
      # To make the QC Plot update immediately, overlay the live inputs on top
      # of the saved config for whichever antigen is currently selected.
      live_ag_raw     <- input$pb_selected_antigen %||% ""
      live_ag_display <- .ag_raw_to_display(live_ag_raw)
      live_pos        <- input$pb_pos_ctrl
      live_neg        <- input$pb_neg_ctrl

      # For INTELLIFLEX: when neither pos nor neg has been set at all for an
      # antigen (user has not yet visited the Point-based tab), default all
      # control wells to positive -- mirroring the BioPlex selectInput which
      # pre-selects every choice as positive by default.
      is_ix_ctrl       <- identical(rv$instrument, "INTELLIFLEX")
      all_ctrl_choices <- names(pb_ctrl_choices())
      if (is_ix_ctrl && length(all_ctrl_choices) == 0)
        all_ctrl_choices <- helper_ctrl_names()

      for (ag in unique(out$analyte)) {
        cfg       <- ag_cfg_display[[ag]]
        pos_names <- if (!is.null(cfg)) cfg$pos_ctrl else NULL
        neg_names <- if (!is.null(cfg)) cfg$neg_ctrl else NULL

        # Override with live inputs when this is the currently-selected antigen
        # so the plot reflects Point-based panel selections in real time.
        # Only override when the live inputs are non-NULL (i.e. the selectInput
        # has been rendered and the user has interacted with it).  When NULL the
        # UI hasn't initialised yet, so keep the saved config to avoid wiping it.
        if (nchar(live_ag_display) > 0 && ag == live_ag_display) {
          if (!is.null(live_pos)) pos_names <- live_pos
          if (!is.null(live_neg)) neg_names <- live_neg
        }

        # When no config has been saved for this antigen, controls remain
        # "unclassified" (the default set above) so the plot shows them as open
        # grey circles.  The user must explicitly select pos/neg per antigen in
        # the Point-based QC Configuration panel -- intentional, since the same
        # physical control can be positive for one analyte and negative for another.

        rows <- out$analyte == ag
        if (!is.null(pos_names) && length(pos_names) > 0)
          out$ctrl_category[rows & out$Sample_ID %in% pos_names] <- "positive"
        if (!is.null(neg_names) && length(neg_names) > 0)
          out$ctrl_category[rows & out$Sample_ID %in% neg_names] <- "negative"
      }
      out
    } else data.frame(Well=character(0), Sample_ID=character(0), analyte=character(0),
                      MFI=numeric(0), Type=character(0), ctrl_category=character(0))

    # ---- Samples: raw per-well rows, use base_sample_id as label ----
    raw_samp <- if (nrow(d_samp) > 0) {
      # Resolve display label: base_sample_id if available, else strip _N from Sample_ID
      label <- d_samp$base_sample_id
      fallback <- is.na(label) | nchar(trimws(label)) == 0
      label[fallback] <- {
        raw <- as.character(d_samp$Sample_ID[fallback])
        stripped <- sub("_[Ss]\\d+.*$", "", raw)
        ifelse(stripped == raw, sub("_\\d+$", "", raw), stripped)
      }
      data.frame(
        Well          = d_samp$Well,
        Sample_ID     = label,
        analyte       = d_samp$analyte,
        MFI           = d_samp$MFI,
        Type          = as.character(d_samp$Type),
        ctrl_category = "sample",
        stringsAsFactors = FALSE
      )
    } else data.frame(Well=character(0), Sample_ID=character(0), analyte=character(0),
                      MFI=numeric(0), Type=character(0), ctrl_category=character(0))

    out <- rbind(raw_ctrl, raw_samp)
    out <- out[!is.na(out$MFI), , drop = FALSE]

    cat_order <- c("positive", "negative", "unclassified", "sample")
    out$ctrl_category <- factor(out$ctrl_category, levels = cat_order)
    out <- out[order(out$ctrl_category, out$Sample_ID), ]
    out
  })

  # Dynamic title ---------------------------------------------------------
  output$tit_plot_title <- renderUI({
    ag   <- input$tit_analyte %||% "All Analytes"
    ctrl <- input$tit_control %||% "All Controls"
    tags$span(
      style = "font-size:13px;",
      tags$strong(style = "color:#1a3a5c;", paste0("Analyte: ", ag)),
      tags$span(style = "color:#aaa;", "  |  "),
      tags$span(style = "color:#1a3a5c;", paste0("Control: ", ctrl))
    )
  })

  # ---------------------------------------------------------------------------
  # PLOT 1 -- Titration curves (sample wells)
  #
  # Design:
  #   - Always faceted by analyte (one panel per analyte)
  #   - Colour  = base_sample_id  (FH1_1 and FH1_2 share the same colour)
  #   - Linetype = replicate index from _N suffix (solid / dashed / dotted ...)
  #   - Each individual Sample_ID is its own curve
  #   - NO LOD line here (LOD is only on the Control Point Plot)
  # ---------------------------------------------------------------------------
  tit_plot_reactive <- reactive({
    d   <- tit_plot_data(); req(!is.null(d) && nrow(d) > 0)

    d_plot <- d[!is.na(d$x_value) & d$x_value > 0, , drop = FALSE]
    if (nrow(d_plot) == 0) {
      return(ggplot() +
        annotate("text", x = 0.5, y = 0.5,
                 label = paste0(
                   "No concentration/dilution values found.\n",
                   "Fill 'start_concentration', 'dilution_factor', and\n",
                   "'samples' (mab / serum) columns in the helper file."),
                 size = 5, hjust = 0.5, colour = "#888") +
        theme_void())
    }

    # X-axis label
    kinds   <- tolower(trimws(unique(d_plot$sample_kind[!is.na(d_plot$sample_kind)])))
    x_label <- if (length(kinds) == 1 && kinds == "mab")   "Concentration (\u00b5g/mL)" else
               if (length(kinds) == 1 && kinds == "serum") "Dilution"                   else
               "Concentration / Dilution"

    # One colour per base_sample_id -- each is a single titration curve
    colour_grps <- sort(unique(d_plot$base_sample_id))
    pal         <- setNames(.make_pal(length(colour_grps)), colour_grps)

    p <- ggplot(d_plot,
                aes(x      = x_value,
                    y      = avg_MFI,
                    colour = base_sample_id,
                    group  = interaction(base_sample_id, analyte))) +
      geom_line(aes(linetype = analyte), linewidth = 0.45, alpha = 0.55) +
      geom_smooth(aes(linetype = analyte), method = "loess", se = FALSE, span = 0.75,
                  linewidth = 1.1, alpha = 0.9) +
      geom_point(size = 3, alpha = 0.9) +
      scale_y_continuous(labels = scales::label_comma(accuracy = 1)) +
      scale_colour_manual(values = pal, name = "Sample") +
      labs(x = x_label, y = "Average MFIs", linetype = "Analyte") +
      theme_minimal(base_size = 12) +
      theme(
        legend.position  = "right",
        legend.text      = element_text(size = 9),
        legend.title     = element_text(size = 10, face = "bold"),
        panel.grid.minor = element_blank(),
        axis.title       = element_text(size = 11, face = "bold"),
        strip.text       = element_text(face = "bold", size = 11, colour = "#1a3a5c"),
        strip.background = element_rect(fill = "#eef3f9", colour = NA),
        plot.margin      = margin(10, 10, 10, 10)
      )

    p <- p + scale_x_log10(labels = scales::label_comma(accuracy = 0.001))

    # ---- Per-antigen LOD -- dotted reference line(s) ---------------------
    # Uses the same Point-based QC Configuration lookup as the QC Plot per
    # Antigen. For non-BRILLIANT BioPlex runs LOD is forced to 0 (and
    # INTELLIFLEX defaults to 0 until configured), so no line is drawn then.
    analytes_tit <- sort(unique(as.character(d_plot$analyte)))
    lod_df_tit   <- .tit_lod_lookup(analytes_tit)
    lod_df_tit   <- lod_df_tit[!is.na(lod_df_tit$lod) & lod_df_tit$lod > 0, , drop = FALSE]
    if (nrow(lod_df_tit) > 0) {
      p <- p +
        geom_hline(data        = lod_df_tit,
                   aes(yintercept = lod),
                   linetype    = "dotted", colour = "gray30",
                   linewidth   = 0.7, inherit.aes = FALSE) +
        geom_text(data        = lod_df_tit,
                  aes(x = -Inf, y = lod,
                      label = paste0(analyte, " LOD = ", round(lod, 1))),
                  hjust       = -0.05, vjust = -0.4,
                  size        = 2.6, colour = "gray30", fontface = "italic",
                  inherit.aes = FALSE, check_overlap = TRUE)
    }

    # Single-analyte data: drop the redundant linetype legend
    n_analytes <- length(unique(d_plot$analyte))
    if (n_analytes <= 1) p <- p + guides(linetype = "none")

    p
  })
  output$tit_curve_plot <- renderPlot({ tit_plot_reactive() }, res = 96)

  output$tit_save_png <- downloadHandler(
    filename = function() paste0("titration_curve_", format(Sys.Date(), "%Y%m%d"), ".png"),
    content  = function(file) {
      p <- tit_plot_reactive(); req(!is.null(p))
      ggplot2::ggsave(file, plot = p, device = "png", width = 10, height = 6, dpi = 150)
    }
  )

  # ---------------------------------------------------------------------------
  # PLOT 2 -- Controls point plot (LOD shown here, per analyte)
  # ---------------------------------------------------------------------------
  tit_ctrl_plot_reactive <- reactive({
    d <- tit_ctrl_data()

    # Respect the "Sample (sample_id)" pill selector from the Titration
    # sub-tab: only restrict the "sample" rows (ctrl_category == "sample"),
    # leaving Positive/Negative/Unclassified control points untouched, and
    # leaving tit_ctrl_data() itself unfiltered since it is also reused by
    # the Excel export's 10_control_plots sheet.
    if (!is.null(d) && nrow(d) > 0) {
      sel_samp <- tit_selected_base_ids()
      if (length(sel_samp) > 0) {
        drop_rows <- d$ctrl_category == "sample" & !(d$Sample_ID %in% sel_samp)
        d <- d[!drop_rows, , drop = FALSE]
      }
    }

    if (is.null(d) || nrow(d) == 0) {
      return(ggplot() +
        annotate("text", x = 0.5, y = 0.5,
                 label = paste0(
                   "No data to plot.\n",
                   "Controls need Type C1-Cn with Positive/Negative selected in\n",
                   "Point-based QC Configuration. Samples need Type X1-Xn / U1-Un."),
                 size = 4.5, hjust = 0.5, colour = "#888") +
        theme_void())
    }

    # ---- Per-antigen LOD from ag_config (Point-based QC) ----------------------
    # Draw for every analyte that has a LOD configured in the Point-based tab,
    # whether or not it is locked.  Fall back to the global pb_cfg LOD when no
    # per-analyte value has been saved yet.
    # For the currently-selected antigen the live input$pb_lod is used so the
    # LOD line updates immediately without requiring the user to navigate away.
    analytes <- sort(unique(as.character(d$analyte)))
    lod_df   <- .tit_lod_lookup(analytes)

    # ---- X-axis grouping: Positive Control | Negative Control | Sample --------
    # Controls classified in the Point-based QC tab show as Positive (red) or
    # Negative (blue). Unclassified controls remain visible as open grey circles
    # so the user is reminded to classify them in the Point-based QC tab.
    d$x_group <- dplyr::case_when(
      d$ctrl_category == "positive"     ~ "Positive\nControl",
      d$ctrl_category == "negative"     ~ "Negative\nControl",
      d$ctrl_category == "unclassified" ~ "Unclassified\nControl",
      TRUE                              ~ "Sample"
    )
    group_order    <- c("Positive\nControl", "Negative\nControl",
                        "Unclassified\nControl", "Sample")
    present_groups <- intersect(group_order, unique(d$x_group))
    d$x_group      <- factor(d$x_group, levels = present_groups)

    # ---- Colour + shape scheme ------------------------------------------------
    # Matches sketch: red = positive, blue = negative,
    #                 open grey = unclassified, dark grey = sample
    cat_colours <- c(
      "Positive\nControl"     = "#e74c3c",   # red
      "Negative\nControl"     = "#6c6cdb",   # blue/purple
      "Unclassified\nControl" = "#95a5a6",   # grey
      "Sample"                 = "#555555"    # dark grey
    )
    cat_shapes <- c(
      "Positive\nControl"     = 16L,   # filled circle
      "Negative\nControl"     = 16L,   # filled circle
      "Unclassified\nControl" = 1L,    # open circle
      "Sample"                 = 16L    # filled circle
    )

    jitter_pos <- position_jitter(width = 0.18, height = 0, seed = 42)

    p <- ggplot(d, aes(x      = x_group,
                       y      = MFI,
                       colour = x_group,
                       shape  = x_group,
                       label  = Sample_ID)) +
      geom_point(position = jitter_pos, size = 3.2, alpha = 0.88) +
      ggrepel::geom_text_repel(
        position        = jitter_pos,
        size            = 2.6,
        max.overlaps    = 20,
        box.padding     = 0.25,
        segment.size    = 0.3,
        segment.colour  = "#aaaaaa",
        show.legend     = FALSE
      ) +
      scale_colour_manual(values = cat_colours, name = NULL, drop = TRUE) +
      scale_shape_manual (values = cat_shapes,  name = NULL, drop = TRUE) +
      scale_x_discrete(drop = TRUE) +
      scale_y_continuous(
        labels = scales::label_comma(accuracy = 1),
        expand = expansion(mult = c(0.05, 0.20))
      ) +
      facet_wrap(~ analyte, scales = "free_y") +
      labs(x = NULL, y = "MFI",
           title = "QC Plot per Antigen") +
      theme_minimal(base_size = 12) +
      theme(
        axis.text.x        = element_text(size = 11, face = "bold"),
        axis.text.y        = element_text(size = 9),
        panel.grid.major.x = element_blank(),
        panel.grid.minor   = element_blank(),
        axis.title.y       = element_text(size = 11, face = "bold"),
        strip.text         = element_text(face = "bold", size = 11, colour = "#1a3a5c"),
        strip.background   = element_rect(fill = "#eef3f9", colour = NA),
        legend.position    = "right",
        legend.text        = element_text(size = 9),
        panel.border       = element_rect(colour = "#cccccc", fill = NA, linewidth = 0.4),
        plot.title         = element_text(face = "bold", size = 13, colour = "#1a3a5c",
                                          hjust = 0.5)
      )

    # ---- LOD horizontal dashed line -- always drawn from Point-based QC config --
    # Solid line when LOD is locked; dashed when using current (unlocked) value.
    if (nrow(lod_df) > 0) {
      lod_locked_rows   <- lod_df[lod_df$lod_locked == TRUE,  , drop = FALSE]
      lod_unlocked_rows <- lod_df[lod_df$lod_locked != TRUE, , drop = FALSE]

      if (nrow(lod_locked_rows) > 0) {
        p <- p +
          geom_hline(data      = lod_locked_rows,
                     aes(yintercept = lod),
                     linetype  = "dashed", colour = "#c0392b",
                     linewidth = 0.85, inherit.aes = FALSE) +
          geom_text(data       = lod_locked_rows,
                    aes(x = -Inf, y = lod,
                        label = paste0("LOD = ", round(lod, 1), " *")),
                    hjust = -0.08, vjust = -0.45,
                    size = 2.9, colour = "#c0392b", fontface = "bold.italic",
                    inherit.aes = FALSE)
      }
      if (nrow(lod_unlocked_rows) > 0) {
        p <- p +
          geom_hline(data      = lod_unlocked_rows,
                     aes(yintercept = lod),
                     linetype  = "dashed", colour = "#888888",
                     linewidth = 0.65, inherit.aes = FALSE) +
          geom_text(data       = lod_unlocked_rows,
                    aes(x = -Inf, y = lod,
                        label = paste0("LOD = ", round(lod, 1))),
                    hjust = -0.08, vjust = -0.45,
                    size = 2.9, colour = "#888888", fontface = "italic",
                    inherit.aes = FALSE)
      }
    }

    p
  })
  output$tit_ctrl_plot <- renderPlot({ tit_ctrl_plot_reactive() }, res = 96)

  output$tit_save_ctrl_png <- downloadHandler(
    filename = function() paste0("control_plot_", format(Sys.Date(), "%Y%m%d"), ".png"),
    content  = function(file) {
      p <- tit_ctrl_plot_reactive(); req(!is.null(p))
      ggplot2::ggsave(file, plot = p, device = "png", width = 10, height = 5, dpi = 150)
    }
  )

  # ---------------------------------------------------------------------------
  # DATA TABLE (Titration sub-tab)
  # ---------------------------------------------------------------------------
  output$tit_data_table <- renderDT({
    d <- tit_plot_data(); req(!is.null(d) && nrow(d) > 0)
    kinds <- tolower(trimws(unique(d$sample_kind[!is.na(d$sample_kind)])))
    x_col_lbl <- if (length(kinds) == 1 && kinds == "mab")   "concentration_ug_mL" else
                 if (length(kinds) == 1 && kinds == "serum") "dilution"             else
                 "concentration_or_dilution"
    tbl <- data.frame(
      analyte        = d$analyte,
      sample_id      = d$base_sample_id,
      sample_type    = d$sample_type,
      sample_kind    = d$sample_kind,
      stringsAsFactors = FALSE
    )
    tbl[[x_col_lbl]] <- round(d$x_value, 5)
    tbl$avg_MFI <- round(d$avg_MFI, 1)
    tbl$n_reps  <- d$n_reps
    tbl <- tbl[order(tbl$analyte, tbl$sample_id, tbl[[x_col_lbl]]), ]
    row.names(tbl) <- NULL
    datatable(tbl, filter = "top",
              options = list(pageLength = 16, scrollX = TRUE, dom = "lfrtip",
                             columnDefs = list(list(className = "dt-center", targets = "_all"))),
              rownames = FALSE, class = "stripe hover cell-border compact") %>%
      formatRound(columns = c(x_col_lbl, "avg_MFI"), digits = 3)
  })

  # ===========================================================================
  # AUC sub-tab
  # ===========================================================================
  # AUC method: trapezoidal rule on log10(concentration) vs average MFI.
  # For each sample x analyte pair, sort by concentration, compute:
  #   AUC = \u03a3 0.5 * (log10(c[i+1]) - log10(c[i])) * (MFI[i+1] + MFI[i])
  # This is equivalent to the area in log-concentration space (common in BAMA).

  .calc_auc <- function(conc, mfi) {
    # Remove rows with missing / non-positive concentration
    ok  <- !is.na(conc) & !is.na(mfi) & conc > 0
    if (sum(ok) < 2) return(NA_real_)
    ord  <- order(conc[ok])
    cx   <- log10(conc[ok][ord])
    yx   <- mfi[ok][ord]
    # Use DescTools::AUC (trapezoidal / linear interpolation)
    auc  <- tryCatch(
      DescTools::AUC(cx, yx, method = "trapezoid"),
      error = function(e) {
        # Fallback: manual trapezoidal rule
        n <- length(cx)
        sum(0.5 * diff(cx) * (yx[-n] + yx[-1]))
      }
    )
    round(auc, 4)
  }

  # ---------------------------------------------------------------------------
  # EC50 helper using drc::drm (4-parameter log-logistic, LL.4 / LL.5 model)
  # Returns a named list: ec50, slope, lower, upper, converged, model
  # Inputs: conc (numeric vector), mfi (numeric vector, 0-100% inhibition)
  # Mirrors the previous titers-sheet pattern: ic50_4pl / model_converged / slope_4pl etc.
  # ---------------------------------------------------------------------------
  .calc_ec50 <- function(conc, mfi) {
    blank <- list(
      ec50      = NA_real_,
      slope     = NA_real_,
      lower     = NA_real_,
      upper     = NA_real_,
      converged = FALSE,
      model     = NULL
    )
    ok <- !is.na(conc) & !is.na(mfi) & conc > 0
    if (sum(ok) < 4) return(blank)
    cx <- conc[ok]
    yx <- mfi[ok]

    fit <- tryCatch(
      suppressWarnings(
        drc::drm(yx ~ cx, fct = drc::LL.4(
          names = c("Slope", "Lower", "Upper", "EC50")
        ))
      ),
      error = function(e) NULL
    )
    if (is.null(fit)) {
      # Fallback: LL.5 (5-parameter)
      fit <- tryCatch(
        suppressWarnings(
          drc::drm(yx ~ cx, fct = drc::LL.5(
            names = c("Slope", "Lower", "Upper", "EC50", "f")
          ))
        ),
        error = function(e) NULL
      )
    }
    if (is.null(fit)) return(blank)

    cf <- tryCatch(coef(fit), error = function(e) NULL)
    if (is.null(cf)) return(blank)

    # drc coef names: "Slope:(Intercept)", "Lower:(Intercept)", etc.
    get_coef <- function(name_pat) {
      idx <- grep(name_pat, names(cf), ignore.case = TRUE)
      if (length(idx) > 0) unname(cf[idx[1]]) else NA_real_
    }

    list(
      ec50      = round(get_coef("EC50"),  4),
      slope     = round(get_coef("Slope"), 4),
      lower     = round(get_coef("Lower"), 4),
      upper     = round(get_coef("Upper"), 4),
      converged = TRUE,
      model     = fit
    )
  }

  # Build EC50 summary table (mirrors the previous titers-sheet columns)
  # Returns a data.frame with one row per (analyte, base_sample_id)
  ec50_summary <- reactive({
    d <- tit_base_data()
    if (is.null(d) || nrow(d) == 0) return(NULL)

    # Keep only sample rows that have concentration values
    d <- d[!is.na(d$x_value) & d$x_value > 0 &
           !is.na(d$sample_type) &
           tolower(trimws(d$sample_type)) == "sample", , drop = FALSE]
    if (nrow(d) == 0) return(NULL)

    # Average replicates at each x_value per (analyte, base_sample_id)
    agg <- as.data.frame(
      d %>%
        dplyr::group_by(analyte, base_sample_id, sample_kind, x_value) %>%
        dplyr::summarise(avg_MFI = mean(MFI, na.rm = TRUE),
                         n_reps  = dplyr::n(), .groups = "drop")
    )
    if (nrow(agg) == 0) return(NULL)

    meta <- rv$helper_edited
    sci_id <- if (!is.null(rv$run_analyst) && nchar(trimws(rv$run_analyst)) > 0)
                trimws(rv$run_analyst) else
              if (!is.null(meta) && "scientist_id" %in% colnames(meta))
                as.character(meta$scientist_id[1]) else ""
    run_date_val <- if (!is.null(rv$run_date_used) && nchar(trimws(rv$run_date_used)) > 0)
                      trimws(rv$run_date_used) else format(Sys.Date(), "%Y%m%d")

    rows_list <- list()
    for (ag in unique(agg$analyte)) {
      for (bid in unique(agg$base_sample_id[agg$analyte == ag])) {
        sub <- agg[agg$analyte == ag & agg$base_sample_id == bid, , drop = FALSE]
        n_pts <- nrow(sub)
        sk    <- if (nrow(sub) > 0) tolower(trimws(sub$sample_kind[1])) else ""

        ec50_res <- .calc_ec50(sub$x_value, sub$avg_MFI)

        # Classify outcome (mirrors the previous titers-sheet logic)
        outcome <- if (!ec50_res$converged) {
          "Not-Converged"
        } else if (!is.na(ec50_res$ec50) &&
                   ec50_res$ec50 > min(sub$x_value, na.rm = TRUE) &&
                   ec50_res$ec50 < max(sub$x_value, na.rm = TRUE)) {
          "Titered"
        } else {
          "FLAT"
        }

        rows_list[[length(rows_list) + 1L]] <- data.frame(
          scientist_id    = sci_id,
          analyte         = ag,
          experiment_date = run_date_val,
          sample_id       = bid,
          sample_kind     = sk,
          n_points        = n_pts,
          outcome         = outcome,
          ec50_4pl        = if (ec50_res$converged) ec50_res$ec50  else NA_real_,
          slope_4pl       = if (ec50_res$converged) ec50_res$slope else NA_real_,
          lower_4pl       = if (ec50_res$converged) ec50_res$lower else NA_real_,
          upper_4pl       = if (ec50_res$converged) ec50_res$upper else NA_real_,
          model_converged = if (ec50_res$converged) "Yes" else "No",
          stringsAsFactors = FALSE
        )
      }
    }
    if (length(rows_list) == 0) return(NULL)
    do.call(rbind, rows_list)
  })

  # AUC filter selectors (independent from Titration selectors)

  output$auc_analyte_selector_ui <- renderUI({
    d <- tit_base_data(); req(!is.null(d))
    selectInput("auc_analyte", NULL,
                choices  = c("All Analytes", sort(unique(d$analyte))),
                selected = "All Analytes", width = "100%")
  })
  output$auc_sample_type_selector_ui <- renderUI({
    d <- tit_base_data(); req(!is.null(d))
    types <- sort(unique(d$sample_type[!is.na(d$sample_type) & nchar(d$sample_type) > 0]))
    if (length(types) == 0) types <- sort(unique(d$Type))
    selectInput("auc_sample_type", NULL,
                choices  = c("All Types", types),
                selected = "All Types", width = "100%")
  })
  output$auc_control_selector_ui <- renderUI({
    d <- tit_base_data(); req(!is.null(d))
    ctrls <- sort(unique(d$control[!is.na(d$control) &
                                   nchar(trimws(d$control)) > 0 &
                                   d$control != "NA"]))
    selectInput("auc_control", NULL,
                choices  = c("All Controls", ctrls),
                selected = "All Controls", width = "100%")
  })

  # Filtered + aggregated data for AUC ------------------------------------
  auc_plot_data <- reactive({
    d <- tit_base_data(); req(!is.null(d))
    d <- .tit_apply_filters(d, input$auc_analyte, input$auc_sample_type, input$auc_control)
    # Exclude control-type wells from AUC (keep samples only)
    d <- d[!startsWith(as.character(d$Type), "C"), , drop = FALSE]
    if (nrow(d) == 0) return(NULL)
    # Average replicates per (base_sample_id, analyte, x_value, sample_kind)
    agg <- as.data.frame(
      d %>%
        filter(!is.na(x_value), x_value > 0) %>%
        group_by(base_sample_id, analyte, x_value, sample_kind) %>%
        summarise(avg_MFI = mean(MFI, na.rm = TRUE), .groups = "drop")
    )
    if (nrow(agg) == 0) return(NULL)
    # Compute AUC per (base_sample_id, analyte)
    auc_df <- as.data.frame(
      agg %>%
        group_by(base_sample_id, analyte, sample_kind) %>%
        summarise(
          AUC      = .calc_auc(x_value, avg_MFI),
          n_points = n(),
          x_min    = min(x_value, na.rm = TRUE),
          x_max    = max(x_value, na.rm = TRUE),
          .groups  = "drop"
        ) %>%
        filter(!is.na(AUC)) %>%
        arrange(analyte, desc(AUC))
    )
    # Rename base_sample_id -> Sample_ID for downstream compatibility
    auc_df$Sample_ID <- auc_df$base_sample_id
    auc_df
  })

  # AUC dynamic title -----------------------------------------------------
  output$auc_plot_title <- renderUI({
    ag   <- input$auc_analyte    %||% "All Analytes"
    st   <- input$auc_sample_type %||% "All Types"
    ctrl <- input$auc_control    %||% "All Controls"
    tags$span(
      style = "font-size:13px;",
      tags$strong(style = "color:#1a3a5c;", paste0("AUC -- Analyte: ", ag)),
      tags$span(style = "color:#aaa;", "  |  "),
      tags$span(style = "color:#1a3a5c;", paste0("Type: ", st))
    )
  })

  # AUC bar chart reactive ------------------------------------------------
  auc_plot_reactive <- reactive({
    df  <- auc_plot_data(); req(!is.null(df) && nrow(df) > 0)

    multi_ag <- length(unique(df$analyte)) > 1
    n_s      <- length(unique(df$Sample_ID))
    pal      <- .make_pal(n_s)

    # Determine x-axis descriptor for AUC label
    kinds <- tolower(trimws(unique(df$sample_kind[!is.na(df$sample_kind)])))
    auc_x_lbl <- if (length(kinds) == 1 && kinds == "mab") "log\u2081\u2080 concentration"
                 else if (length(kinds) == 1 && kinds == "serum") "log\u2081\u2080 dilution"
                 else "log\u2081\u2080 concentration/dilution"

    # Reorder Sample_ID by descending AUC within each analyte
    df$Sample_ID <- factor(df$Sample_ID,
                           levels = unique(df$Sample_ID[order(df$analyte, -df$AUC)]))

    p <- ggplot(df, aes(x = Sample_ID, y = AUC, fill = Sample_ID)) +
      geom_col(colour = "white", linewidth = 0.3, alpha = 0.92) +
      geom_text(aes(label = formatC(AUC, format = "f", digits = 2)),
                hjust = -0.1, size = 3, colour = "#333333") +
      scale_fill_manual(values = pal, name = "Sample") +
      scale_y_continuous(expand = expansion(mult = c(0, 0.18)),
                         labels = scales::label_comma(accuracy = 0.01)) +
      coord_flip() +
      labs(x = NULL, y = paste0("AUC  (trapezoidal, ", auc_x_lbl, ")")) +
      theme_minimal(base_size = 12) +
      theme(
        legend.position  = "none",
        panel.grid.major.y = element_blank(),
        panel.grid.minor   = element_blank(),
        axis.title.x     = element_text(size = 10, face = "bold"),
        axis.text.y      = element_text(size = 10),
        strip.text       = element_text(face = "bold", size = 11),
        plot.margin      = margin(10, 30, 10, 10)
      )

    if (multi_ag) p <- p + facet_wrap(~ analyte, scales = "free")

    p
  })
  output$auc_bar_plot <- renderPlot({ auc_plot_reactive() }, res = 96)

  output$auc_save_png <- downloadHandler(
    filename = function() paste0("auc_barplot_", format(Sys.Date(), "%Y%m%d"), ".png"),
    content  = function(file) {
      p <- auc_plot_reactive(); req(!is.null(p))
      ggplot2::ggsave(file, plot = p, device = "png", width = 10, height = 7, dpi = 150)
    }
  )

  # AUC summary table -------------------------------------------------------
  output$auc_summary_table <- renderDT({
    df <- auc_plot_data(); req(!is.null(df) && nrow(df) > 0)
    kinds    <- tolower(trimws(unique(df$sample_kind[!is.na(df$sample_kind)])))
    x_lbl    <- if (length(kinds) == 1 && kinds == "mab")   "conc_min_ug_mL" else
                if (length(kinds) == 1 && kinds == "serum") "dil_min"        else "x_min"
    x_lbl2   <- if (length(kinds) == 1 && kinds == "mab")   "conc_max_ug_mL" else
                if (length(kinds) == 1 && kinds == "serum") "dil_max"        else "x_max"
    tbl <- data.frame(
      analyte  = df$analyte,
      sample_id = df$Sample_ID,
      AUC      = round(df$AUC, 4),
      n_points = df$n_points,
      stringsAsFactors = FALSE
    )
    tbl[[x_lbl]]  <- round(df$x_min, 5)
    tbl[[x_lbl2]] <- round(df$x_max, 3)
    tbl <- tbl[order(tbl$analyte, -tbl$AUC), ]
    row.names(tbl) <- NULL
    datatable(tbl, filter = "top",
              options = list(pageLength = 16, scrollX = TRUE, dom = "lfrtip",
                             columnDefs = list(list(className = "dt-center", targets = "_all"))),
              rownames = FALSE, class = "stripe hover cell-border compact") %>%
      formatRound(columns = c("AUC", x_lbl, x_lbl2), digits = 4) %>%
      formatStyle("AUC",
        background         = styleColorBar(c(0, max(df$AUC, na.rm = TRUE)), "#b3d9ff"),
        backgroundSize     = "98% 88%",
        backgroundRepeat   = "no-repeat",
        backgroundPosition = "center")
  })

  # ===========================================================================
  # QUANTIFICATION TAB -- Standard Curve
  # ===========================================================================

  # Reactive: which Processed-data MFI variant to build the standard curve
  # from -- mirrors the 3 sub-tabs already on the Processed page (mAbs-only,
  # blank-beads-only, and the full background subtraction used everywhere
  # else). Defaults to "full" so behavior is unchanged unless the user
  # switches the toggle.
  sc_mfi_variant <- reactive({
    mfi_list <- mfi_dataframe(); req(!is.null(mfi_list))
    switch(input$sc_subtraction_mode %||% "full",
           "mabs"  = mfi_list$mabs,
           "blank" = mfi_list$blank,
           mfi_list$full)
  })

  # Quantification-tab base data (ALL sample types -- standard_curve AND
  # sample rows), rebuilt against the currently-toggled MFI variant. Standard
  # curve points and the sample wells back-calculated against them must use
  # the same subtraction method, so both sc_base_data (below) and
  # sc_sample_base_data (used for back-calculation, further down) are built
  # from this rather than from tit_base_data(), which always stays on "full".
  sc_full_data <- reactive({
    req(rv$analysis_run)
    .tit_base_data_core(sc_mfi_variant())
  })

  # Reactive: filter sc_full_data to standard_curve rows only -----------------
  sc_base_data <- reactive({
    req(rv$analysis_run)
    d <- sc_full_data(); req(!is.null(d))
    d_sc <- d[!is.na(d$sample_type) &
              tolower(trimws(d$sample_type)) == "standard_curve" &
              !is.na(d$x_value) & d$x_value > 0, , drop = FALSE]
    d_sc
  })

  # UI selectors ---------------------------------------------------------------
  # Analyte: single selection -- one at a time, defaults to first available
  output$sc_analyte_selector_ui <- renderUI({
    d <- sc_base_data()
    if (is.null(d) || nrow(d) == 0)
      return(tags$span(style = "color:#888; font-size:12px;",
                       "No standard_curve data found."))
    ags <- sort(unique(d$analyte))
    selectInput("sc_analyte", NULL,
                choices  = ags,
                selected = ags[1],
                width    = "100%")
  })

  # Sample: checkboxes, scoped to the selected analyte, all ticked by default
  output$sc_sample_selector_ui <- renderUI({
    d <- sc_base_data()
    if (is.null(d) || nrow(d) == 0) return(NULL)
    ag <- input$sc_analyte
    if (!is.null(ag) && ag %in% d$analyte)
      d <- d[d$analyte == ag, , drop = FALSE]
    bids <- sort(unique(d$base_sample_id[
      !is.na(d$base_sample_id) & nchar(trimws(d$base_sample_id)) > 0
    ]))
    if (length(bids) == 0) bids <- sort(unique(d$Sample_ID))
    checkboxGroupInput("sc_sample", NULL,
                       choices  = bids,
                       selected = bids,
                       width    = "100%")
  })

  # Fit-range slider: log10 min/max of x_value for selected analyte+samples ----
  output$sc_fit_range_ui <- renderUI({
    d <- sc_base_data()
    if (is.null(d) || nrow(d) == 0) return(NULL)

    ag <- input$sc_analyte
    if (!is.null(ag) && ag %in% d$analyte)
      d <- d[d$analyte == ag, , drop = FALSE]

    samps <- input$sc_sample
    if (!is.null(samps) && length(samps) > 0)
      d <- d[!is.na(d$base_sample_id) & d$base_sample_id %in% samps, , drop = FALSE]

    xv <- sort(unique(d$x_value[!is.na(d$x_value) & d$x_value > 0]))
    if (length(xv) < 2) return(tags$span(style = "color:#888; font-size:12px;",
                                          "Not enough points to set a range."))

    log_xv <- round(log10(xv), 4)
    sliderInput(
      "sc_fit_log_range",
      NULL,
      min   = min(log_xv),
      max   = max(log_xv),
      value = c(min(log_xv), max(log_xv)),
      step  = 0.0001,
      width = "100%",
      ticks = TRUE
    )
  })

  # Dynamic plot title ---------------------------------------------------------
  output$sc_plot_title <- renderUI({
    ag   <- input$sc_analyte %||% "\u2014"
    scl  <- if (!is.null(input$sc_log_x) && input$sc_log_x == "log")
               "log\u2081\u2080 scale" else "linear scale"
    xlbl <- if (!is.null(input$sc_x_label_type) && input$sc_x_label_type == "dilution")
               "dilution ratio" else "concentration (\u00b5g/mL)"
    sub_lbl <- switch(input$sc_subtraction_mode %||% "full",
                       "mabs"  = "no_mAbs subtraction only",
                       "blank" = "blank beads subtraction only",
                       "background subtraction (blank + no_mAbs)")
    tags$span(
      style = "font-size:13px;",
      tags$strong(style = "color:#1a3a5c;",
                  paste0("Standard Curve \u2014 Analyte: ", ag,
                         "  |  X: ", scl, ", ", xlbl)),
      tags$br(),
      tags$span(style = "font-size:11px; color:#666;",
                paste0("MFI basis: ", sub_lbl))
    )
  })

  # Filtered + averaged standard curve data ------------------------------------
  sc_plot_data <- reactive({
    d <- sc_base_data(); req(!is.null(d) && nrow(d) > 0)

    # Filter to the single selected analyte
    ag <- input$sc_analyte; req(!is.null(ag) && nchar(ag) > 0)
    d  <- d[d$analyte == ag, , drop = FALSE]

    # Filter to ticked samples
    samps <- input$sc_sample
    if (!is.null(samps) && length(samps) > 0)
      d <- d[!is.na(d$base_sample_id) & d$base_sample_id %in% samps, , drop = FALSE]

    if (nrow(d) == 0) return(NULL)

    # Average replicate wells at the same (sample x concentration)
    as.data.frame(
      d %>%
        group_by(base_sample_id, analyte, Type, x_value, dilution_label) %>%
        summarise(avg_MFI = mean(MFI, na.rm = TRUE),
                  n_reps  = n(),
                  .groups = "drop") %>%
        arrange(base_sample_id, x_value)
    )
  })

  # Standard curve plot --------------------------------------------------------
  sc_plot_reactive <- reactive({
    d              <- sc_plot_data()
    log_x          <- !is.null(input$sc_log_x) && input$sc_log_x == "log"
    use_dil_labels <- !is.null(input$sc_x_label_type) &&
                      input$sc_x_label_type == "dilution"
    show_r2           <- isTRUE(input$sc_show_r2)
    show_samp_labels  <- isTRUE(input$sc_show_sample_labels)

    if (is.null(d) || nrow(d) == 0) {
      return(ggplot() +
        annotate("text", x = 0.5, y = 0.5,
                 label = paste0(
                   "No standard curve data found.\n",
                   "Ensure the helper file has sample_type = 'standard_curve'\n",
                   "and start_concentration filled for those wells."),
                 size = 5, hjust = 0.5, colour = "#888") +
        theme_void())
    }

    samp_grps <- sort(unique(d$base_sample_id))
    pal        <- setNames(rep("blue", length(samp_grps)), samp_grps)

    # -- Fit-range filter (log10 slider) ---------------------------------------
    # d_fit = subset used for geom_smooth and R2; d = all points still plotted.
    fit_range <- input$sc_fit_log_range
    d_fit <- if (!is.null(fit_range) && length(fit_range) == 2 &&
                  any(!is.na(d$x_value) & d$x_value > 0)) {
      lo <- fit_range[1]; hi <- fit_range[2]
      d[!is.na(d$x_value) & d$x_value > 0 &
        log10(d$x_value) >= lo - 1e-9 & log10(d$x_value) <= hi + 1e-9, , drop = FALSE]
    } else {
      d
    }
    fit_subset_active <- !is.null(fit_range) && length(fit_range) == 2 &&
                         nrow(d_fit) < nrow(d)


    # Build x_value -> "1:N" label from the dilution_label column stored in data.
    dil_map <- local({
      idx <- !is.na(d$dilution_label) & !is.na(d$x_value)
      if (any(idx)) {
        pairs <- unique(d[idx, c("x_value", "dilution_label"), drop = FALSE])
        setNames(pairs$dilution_label, as.character(pairs$x_value))
      } else {
        xv <- sort(unique(d$x_value), decreasing = TRUE)
        setNames(as.character(xv), as.character(xv))
      }
    })

    x_label        <- if (use_dil_labels) "Dilution"
                      else if (log_x) "Known Concentration (\u00b5g/mL, log\u2081\u2080 scale)"
                      else "Known Concentration (\u00b5g/mL)"

    # -- Dilution mode: ordered categorical x-axis ----------------------------
    if (use_dil_labels) {
      dil_order   <- dil_map[as.character(sort(unique(d$x_value), decreasing = TRUE))]
      d$x_factor  <- factor(dil_map[as.character(d$x_value)], levels = dil_order)

      p <- ggplot(d, aes(x      = x_factor,
                         y      = avg_MFI,
                         colour = base_sample_id,
                         group  = base_sample_id)) +
        geom_line(linewidth = 0.9, alpha = 0.75) +
        geom_point(aes(size = n_reps), alpha = 0.92, shape = 16) +
        scale_x_discrete(drop = FALSE) +
        scale_colour_manual(values = pal, guide = "none") +
        scale_fill_manual(  values = pal, guide = "none") +
        scale_size_continuous(range = c(2.5, 5.5), guide = "none") +
        scale_y_continuous(labels = scales::label_comma(accuracy = 1),
                           expand = expansion(mult = c(0.05, 0.18))) +
        labs(x = x_label, y = "Average MFIs") +
        theme_minimal(base_size = 12) +
        theme(
          legend.position   = "none",
          panel.grid.minor  = element_blank(),
          axis.title        = element_text(size = 11, face = "bold"),
          axis.text.x       = element_text(angle = 45, hjust = 1, size = 9),
          strip.text        = element_text(face = "bold", size = 12, colour = "#1a3a5c"),
          strip.background  = element_rect(fill = "#eef3f9", colour = NA),
          plot.margin       = margin(10, 10, 10, 10)
        )

    # -- Concentration mode: numeric x-axis with dilution secondary axis ------
    } else {
      # Mark which points are inside vs outside the fit range
      d$in_fit_range <- if (fit_subset_active) {
        lo <- fit_range[1]; hi <- fit_range[2]
        !is.na(d$x_value) & d$x_value > 0 &
          log10(d$x_value) >= lo - 1e-9 & log10(d$x_value) <= hi + 1e-9
      } else {
        rep(TRUE, nrow(d))
      }

      # -- 5PL logistic fit per sample group -------------------------------
      # Fits FI = Lower + (Upper - Lower) / (1 + (Conc/EC50)^Slope)^Asym
      # (drc::LL.5, the same 5-parameter logistic used elsewhere in this app)
      # and draws it as a smooth solid curve in place of a raw point-to-point
      # dashed line -- matching standard Bio-Plex-style standard curve plots.
      fit_curve_list <- lapply(samp_grps, function(sid) {
        sub <- d_fit[d_fit$base_sample_id == sid & !is.na(d_fit$x_value) &
                     d_fit$x_value > 0 & !is.na(d_fit$avg_MFI), , drop = FALSE]
        if (nrow(sub) < 4) return(NULL)

        fit <- tryCatch(
          suppressWarnings(
            drc::drm(avg_MFI ~ x_value, data = sub,
                     fct = drc::LL.5(names = c("Slope", "Lower", "Upper", "EC50", "Asym")))
          ),
          error = function(e) NULL
        )
        if (is.null(fit)) {
          fit <- tryCatch(
            suppressWarnings(
              drc::drm(avg_MFI ~ x_value, data = sub,
                       fct = drc::LL.4(names = c("Slope", "Lower", "Upper", "EC50")))
            ),
            error = function(e) NULL
          )
        }
        if (is.null(fit)) return(NULL)

        x_rng <- range(sub$x_value, na.rm = TRUE)
        x_seq <- if (log_x)
                   10^seq(log10(x_rng[1]), log10(x_rng[2]), length.out = 200)
                 else
                   seq(x_rng[1], x_rng[2], length.out = 200)

        y_seq <- tryCatch(
          suppressWarnings(as.numeric(predict(fit, newdata = data.frame(x_value = x_seq)))),
          error = function(e) NULL
        )
        if (is.null(y_seq)) return(NULL)

        data.frame(base_sample_id = sid, x_value = x_seq, avg_MFI = y_seq,
                   stringsAsFactors = FALSE)
      })
      fit_curve_df <- do.call(rbind, Filter(Negate(is.null), fit_curve_list))

      p <- ggplot(d, aes(x      = x_value,
                         y      = avg_MFI,
                         colour = base_sample_id,
                         fill   = base_sample_id,
                         group  = base_sample_id)) +
        # All points plotted; excluded ones shown as hollow/dimmed
        geom_point(data = d[d$in_fit_range, , drop = FALSE],
                   aes(size = n_reps), alpha = 0.92, shape = 16) +
        {if (fit_subset_active && any(!d$in_fit_range))
          geom_point(data = d[!d$in_fit_range, , drop = FALSE],
                     aes(size = n_reps), alpha = 0.35, shape = 1, stroke = 1.2)
        } +
        {if (!is.null(fit_curve_df) && nrow(fit_curve_df) > 0)
          geom_line(data = fit_curve_df,
                    aes(x = x_value, y = avg_MFI, group = base_sample_id),
                    inherit.aes = FALSE, colour = "blue",
                    linewidth = 0.9, alpha = 0.9)
        } +
        scale_colour_manual(values = pal, guide = "none") +
        scale_fill_manual(  values = pal, guide = "none") +
        scale_size_continuous(range = c(2.5, 5.5), guide = "none") +
        scale_y_continuous(labels = scales::label_comma(accuracy = 1),
                           expand = expansion(mult = c(0.05, 0.22))) +
        labs(x = x_label, y = "Average MFIs") +
        theme_minimal(base_size = 12) +
        theme(
          legend.position   = "none",
          panel.grid.minor  = element_blank(),
          axis.title        = element_text(size = 11, face = "bold"),
          strip.text        = element_text(face = "bold", size = 12, colour = "#1a3a5c"),
          strip.background  = element_rect(fill = "#eef3f9", colour = NA),
          plot.margin       = margin(10, 16, 10, 10)
        )

      x_breaks <- sort(unique(d$x_value))
      x_dil_labels <- dil_map[as.character(x_breaks)]

      if (log_x)
        p <- p + scale_x_log10(
              labels   = scales::label_comma(accuracy = 0.001),
              sec.axis = sec_axis(~ ., breaks = x_breaks,
                                  labels = x_dil_labels,
                                  name   = "Dilution"))
      else
        p <- p + scale_x_continuous(
              labels   = scales::label_comma(accuracy = 0.001),
              sec.axis = sec_axis(~ ., breaks = x_breaks,
                                  labels = x_dil_labels,
                                  name   = "Dilution"))
    }

    # -- Sample estimate overlay from Quantification table --------------------
    # For each sample with a valid est_conc_ug_mL, draw a grey filled point
    # and label showing where the sample's MFI intersects the fitted curve.
    est_tbl <- tryCatch(sc_sample_conc(), error = function(e) NULL)
    if (!is.null(est_tbl) && nrow(est_tbl) > 0 && !use_dil_labels) {
      est_ok <- est_tbl[!is.na(est_tbl$est_conc_ug_mL) &
                        est_tbl$est_conc_ug_mL > 0 &
                        est_tbl$status == "OK", , drop = FALSE]
      if (nrow(est_ok) > 0) {
        p <- p +
          geom_point(
            data        = est_ok,
            aes(x = est_conc_ug_mL, y = avg_MFI),
            inherit.aes = FALSE,
            shape       = 21,
            size        = 3.5,
            colour      = "grey25",
            fill        = "grey70",
            alpha       = 0.90
          )

        if (show_samp_labels) {
          p <- p +
            geom_label(
              data        = est_ok,
              aes(x = est_conc_ug_mL, y = avg_MFI,
                  label = paste0(sample_id, "\n",
                                 formatC(est_conc_ug_mL, format = "fg", digits = 3),
                                 " \u00b5g/mL")),
              inherit.aes   = FALSE,
              size          = 2.8,
              colour        = "grey20",
              fill          = "white",
              alpha         = 0.85,
              label.padding = unit(0.18, "lines"),
              label.size    = 0.25,
              hjust         = -0.10,
              vjust         = 0.5,
              show.legend   = FALSE
            )
        }
      }
    }

    # -- R2 annotation (per sample panel, top-left corner) --------------------
    if (show_r2 && !use_dil_labels && nrow(d_fit) >= 2) {
      r2_rows <- lapply(samp_grps, function(sid) {
        sub <- d_fit[d_fit$base_sample_id == sid, , drop = FALSE]
        if (nrow(sub) < 2) return(NULL)
        fit <- tryCatch(
          if (log_x) lm(avg_MFI ~ log10(x_value), data = sub)
          else       lm(avg_MFI ~ x_value,         data = sub),
          error = function(e) NULL
        )
        if (is.null(fit)) return(NULL)
        r2 <- summary(fit)$r.squared
        data.frame(base_sample_id = sid,
                   label = sprintf("R\u00b2 = %.4f", r2),
                   stringsAsFactors = FALSE)
      })
      r2_df <- do.call(rbind, Filter(Negate(is.null), r2_rows))
      if (!is.null(r2_df) && nrow(r2_df) > 0) {
        p <- p +
          geom_label(
            data        = r2_df,
            aes(label = label),
            x           = -Inf, y = Inf,
            hjust       = -0.08, vjust = 1.3,
            inherit.aes = FALSE,
            size        = 3.2,
            colour      = "#1a3a5c",
            fill        = "#eef3f9",
            label.size  = 0.25,
            alpha       = 0.88
          )
      }
    }

    # -- Facet: one panel per sample -------------------------------------------
    if (length(samp_grps) > 1)
      p <- p + facet_wrap(~ base_sample_id, scales = "free_y", ncol = 2)
    else
      p <- p + labs(caption = paste0("Sample: ", samp_grps[1]))

    p
  })
  output$sc_curve_plot <- renderPlot({ sc_plot_reactive() }, res = 96)

  output$sc_save_png <- downloadHandler(
    filename = function() {
      ag <- gsub("[^A-Za-z0-9_]", "_", input$sc_analyte %||% "analyte")
      paste0("standard_curve_", ag, "_", format(Sys.Date(), "%Y%m%d"), ".png")
    },
    content  = function(file) {
      p <- sc_plot_reactive(); req(!is.null(p))
      d <- sc_plot_data()
      n_samps <- if (!is.null(d)) length(unique(d$base_sample_id)) else 1
      plot_h  <- max(5, ceiling(n_samps / 2) * 3.5)
      ggplot2::ggsave(file, plot = p, device = "png",
                      width = 11, height = plot_h, dpi = 150)
    }
  )

  # Standard curve data table --------------------------------------------------
  output$sc_data_table <- renderDT({
    d <- sc_plot_data(); req(!is.null(d) && nrow(d) > 0)

    tbl <- data.frame(
      analyte        = d$analyte,
      sample_id      = d$base_sample_id,
      dilution       = ifelse(!is.na(d$dilution_label), d$dilution_label, as.character(d$x_value)),
      known_conc     = round(d$x_value, 5),
      avg_MFI        = round(d$avg_MFI, 1),
      n_reps         = d$n_reps,
      stringsAsFactors = FALSE
    )
    tbl <- tbl[order(tbl$analyte, tbl$sample_id, tbl$known_conc), ]
    row.names(tbl) <- NULL
    datatable(tbl, filter = "top",
              options = list(pageLength = 20, scrollX = TRUE, dom = "lfrtip",
                             columnDefs = list(list(className = "dt-center",
                                                    targets = "_all"))),
              rownames = FALSE, class = "stripe hover cell-border compact") %>%
      formatRound(columns = c("known_conc", "avg_MFI"), digits = 3) %>%
      formatStyle("avg_MFI",
        background         = styleColorBar(c(0, max(tbl$avg_MFI, na.rm = TRUE)), "#c8e6c9"),
        backgroundSize     = "98% 88%",
        backgroundRepeat   = "no-repeat",
        backgroundPosition = "center")
  })

  # ===========================================================================
  # QUANTIFICATION -- Back-calculated sample concentrations
  # ===========================================================================

  # Reactive: sample rows for the currently selected analyte -------------------
  sc_sample_base_data <- reactive({
    req(rv$analysis_run)
    d <- sc_full_data(); req(!is.null(d))
    ag <- input$sc_analyte; req(!is.null(ag) && nchar(ag) > 0)
    d_samp <- d[!is.na(d$sample_type) &
                tolower(trimws(d$sample_type)) == "sample" &
                !is.na(d$analyte) & d$analyte == ag, , drop = FALSE]
    d_samp
  })

  # Reactive: back-calculate concentration for each sample well ----------------
  sc_sample_conc <- reactive({
    d_samp <- sc_sample_base_data()
    if (is.null(d_samp) || nrow(d_samp) == 0) return(NULL)

    # Get standard curve fit data (fit subset, respecting slider)
    d_sc   <- sc_base_data(); req(!is.null(d_sc) && nrow(d_sc) > 0)
    ag     <- input$sc_analyte; req(!is.null(ag))
    d_sc   <- d_sc[d_sc$analyte == ag, , drop = FALSE]
    log_x  <- !is.null(input$sc_log_x) && input$sc_log_x == "log"

    # Apply fit range filter
    fit_range <- input$sc_fit_log_range
    if (!is.null(fit_range) && length(fit_range) == 2) {
      lo <- fit_range[1]; hi <- fit_range[2]
      d_sc <- d_sc[!is.na(d_sc$x_value) & d_sc$x_value > 0 &
                   log10(d_sc$x_value) >= lo - 1e-9 &
                   log10(d_sc$x_value) <= hi + 1e-9, , drop = FALSE]
    }
    if (nrow(d_sc) < 2) return(NULL)

    # Average replicates within the standard curve data per (sample x x_value)
    sc_avg <- as.data.frame(
      d_sc %>%
        group_by(base_sample_id, analyte, x_value) %>%
        summarise(avg_MFI = mean(MFI, na.rm = TRUE), .groups = "drop")
    )

    # Fit one LM per standard curve sample (base_sample_id)
    # Then for each unknown sample well, average MFI is projected per curve
    sc_ids   <- sort(unique(sc_avg$base_sample_id))
    mfi_upper_per_curve <- sapply(sc_ids, function(sid) {
      sub <- sc_avg[sc_avg$base_sample_id == sid, , drop = FALSE]
      if (nrow(sub) < 2) return(NA_real_)
      max(sub$avg_MFI, na.rm = TRUE)
    })
    mfi_lower_per_curve <- sapply(sc_ids, function(sid) {
      sub <- sc_avg[sc_avg$base_sample_id == sid, , drop = FALSE]
      if (nrow(sub) < 2) return(NA_real_)
      min(sub$avg_MFI, na.rm = TRUE)
    })
    fit_list <- lapply(sc_ids, function(sid) {
      sub <- sc_avg[sc_avg$base_sample_id == sid, , drop = FALSE]
      if (nrow(sub) < 2) return(NULL)
      if (log_x) lm(avg_MFI ~ log10(x_value), data = sub)
      else       lm(avg_MFI ~ x_value,         data = sub)
    })
    names(fit_list)          <- sc_ids
    names(mfi_upper_per_curve) <- sc_ids
    names(mfi_lower_per_curve) <- sc_ids

    # Average the sample MFI replicates per well group (base_sample_id x analyte x x_value).
    # Samples sharing the same base_sample_id AND x_value are true duplicates (same X-number,
    # e.g. both X1) -- they must be averaged into one row and yield a single concentration.
    # Type is NOT included in the grouping so that duplicate wells at the same dilution step
    # are always merged regardless of how Type codes happen to differ across wells.
    samp_avg <- as.data.frame(
      d_samp %>%
        group_by(base_sample_id, analyte, x_value) %>%
        summarise(avg_MFI = mean(MFI, na.rm = TRUE),
                  n_reps  = n(),
                  .groups = "drop")
    )
    # Build full sample label: base_sample_id + "_" + zero-padded step rank.
    # Step rank is the ordinal position of x_value within each base_sample_id
    # (ascending), so the first dilution point is _01, second is _02, etc.
    samp_avg <- samp_avg %>%
      group_by(base_sample_id) %>%
      mutate(step_rank     = rank(x_value, ties.method = "first"),
             full_sample_id = paste0(base_sample_id, "_",
                                     sprintf("%02d", as.integer(step_rank)))) %>%
      ungroup() %>%
      as.data.frame()

    # Back-calculate: for each sample x each standard curve, invert the LM
    # LM:  MFI = b0 + b1 * log10(x)  =>  log10(x) = (MFI - b0) / b1
    # => x = 10^((MFI - b0) / b1)   [log mode]
    # LM:  MFI = b0 + b1 * x         =>  x = (MFI - b0) / b1  [linear mode]
    rows_out <- list()
    for (i in seq_len(nrow(samp_avg))) {
      srow    <- samp_avg[i, , drop = FALSE]
      mfi_val <- srow$avg_MFI

      for (sc_id in sc_ids) {
        fit <- fit_list[[sc_id]]
        if (is.null(fit)) next
        b0  <- coef(fit)[1]; b1 <- coef(fit)[2]
        upper <- mfi_upper_per_curve[[sc_id]]
        lower <- mfi_lower_per_curve[[sc_id]]

        status <- if (is.na(mfi_val)) "No MFI data"
                  else if (mfi_val > upper) "> upper limit"
                  else if (mfi_val < lower) "< lower limit"
                  else "OK"

        est_conc <- if (status == "OK") {
          if (log_x) 10^((mfi_val - b0) / b1)
          else       (mfi_val - b0) / b1
        } else NA_real_

        log10_conc <- if (status == "OK") {
          if (log_x) (mfi_val - b0) / b1           # direct log10 result
          else       log10(max((mfi_val - b0) / b1, 1e-15))  # log of linear result
        } else NA_real_

        rows_out[[length(rows_out) + 1L]] <- data.frame(
          analyte          = srow$analyte,
          sample_id        = srow$full_sample_id,
          x_value          = srow$x_value,
          std_curve_sample = sc_id,
          avg_MFI          = round(mfi_val, 1),
          n_reps           = srow$n_reps,
          est_conc_ug_mL   = if (is.na(est_conc))    NA_real_ else round(est_conc,    5),
          log10_conc       = if (is.na(log10_conc))  NA_real_ else round(log10_conc,  5),
          status           = status,
          stringsAsFactors = FALSE
        )
      }
    }
    if (length(rows_out) == 0) return(NULL)
    result <- do.call(rbind, rows_out)
    result[order(result$analyte, result$sample_id, result$x_value, result$std_curve_sample), ]
  })

  # -- All-antigen concentration table (for 16_sample_concentration export) ----
  # Mirrors sc_sample_conc logic but iterates over every analyte in sc_base_data,
  # using the current log_x and fit_range UI settings.
  sc_sample_conc_all <- reactive({
    req(rv$analysis_run)
    d_all  <- sc_full_data(); req(!is.null(d_all))
    d_sc_all <- sc_base_data(); req(!is.null(d_sc_all) && nrow(d_sc_all) > 0)

    log_x     <- !is.null(input$sc_log_x) && input$sc_log_x == "log"
    fit_range <- input$sc_fit_log_range

    all_analytes <- sort(unique(as.character(d_sc_all$analyte)))

    results_list <- lapply(all_analytes, function(ag) {
      # Standard curve data for this analyte
      d_sc <- d_sc_all[d_sc_all$analyte == ag, , drop = FALSE]

      # Apply fit range
      if (!is.null(fit_range) && length(fit_range) == 2) {
        lo <- fit_range[1]; hi <- fit_range[2]
        d_sc <- d_sc[!is.na(d_sc$x_value) & d_sc$x_value > 0 &
                     log10(d_sc$x_value) >= lo - 1e-9 &
                     log10(d_sc$x_value) <= hi + 1e-9, , drop = FALSE]
      }
      if (nrow(d_sc) < 2) return(NULL)

      sc_avg <- as.data.frame(
        d_sc %>%
          dplyr::group_by(base_sample_id, analyte, x_value) %>%
          dplyr::summarise(avg_MFI = mean(MFI, na.rm = TRUE), .groups = "drop")
      )
      sc_ids <- sort(unique(sc_avg$base_sample_id))

      mfi_upper <- sapply(sc_ids, function(sid) {
        sub <- sc_avg[sc_avg$base_sample_id == sid, ]
        if (nrow(sub) < 2) NA_real_ else max(sub$avg_MFI, na.rm = TRUE)
      })
      mfi_lower <- sapply(sc_ids, function(sid) {
        sub <- sc_avg[sc_avg$base_sample_id == sid, ]
        if (nrow(sub) < 2) NA_real_ else min(sub$avg_MFI, na.rm = TRUE)
      })
      fit_list <- lapply(sc_ids, function(sid) {
        sub <- sc_avg[sc_avg$base_sample_id == sid, ]
        if (nrow(sub) < 2) return(NULL)
        tryCatch(
          if (log_x) lm(avg_MFI ~ log10(x_value), data = sub)
          else       lm(avg_MFI ~ x_value,         data = sub),
          error = function(e) NULL)
      })
      names(fit_list)  <- sc_ids
      names(mfi_upper) <- sc_ids
      names(mfi_lower) <- sc_ids

      # Sample wells for this analyte
      d_samp <- d_all[!is.na(d_all$sample_type) &
                      tolower(trimws(d_all$sample_type)) == "sample" &
                      !is.na(d_all$analyte) & d_all$analyte == ag, , drop = FALSE]
      if (nrow(d_samp) == 0) return(NULL)

      samp_avg <- as.data.frame(
        d_samp %>%
          dplyr::group_by(base_sample_id, analyte, x_value) %>%
          dplyr::summarise(avg_MFI = mean(MFI, na.rm = TRUE),
                           n_reps  = dplyr::n(), .groups = "drop")
      )
      samp_avg <- samp_avg %>%
        dplyr::group_by(base_sample_id) %>%
        dplyr::mutate(step_rank      = rank(x_value, ties.method = "first"),
                      full_sample_id = paste0(base_sample_id, "_",
                                              sprintf("%02d", as.integer(step_rank)))) %>%
        dplyr::ungroup() %>%
        as.data.frame()

      rows_out <- list()
      for (i in seq_len(nrow(samp_avg))) {
        srow    <- samp_avg[i, , drop = FALSE]
        mfi_val <- srow$avg_MFI
        for (sc_id in sc_ids) {
          fit <- fit_list[[sc_id]]
          if (is.null(fit)) next
          b0 <- coef(fit)[1]; b1 <- coef(fit)[2]
          upper <- mfi_upper[[sc_id]]; lower <- mfi_lower[[sc_id]]
          status <- if (is.na(mfi_val))       "No MFI data"
                    else if (mfi_val > upper)  "> upper limit"
                    else if (mfi_val < lower)  "< lower limit"
                    else "OK"
          est_conc <- if (status == "OK") {
            if (log_x) 10^((mfi_val - b0) / b1) else (mfi_val - b0) / b1
          } else NA_real_
          log10_conc <- if (status == "OK") {
            if (log_x) (mfi_val - b0) / b1
            else       log10(max((mfi_val - b0) / b1, 1e-15))
          } else NA_real_
          rows_out[[length(rows_out) + 1L]] <- data.frame(
            analyte          = srow$analyte,
            sample_id        = srow$full_sample_id,
            x_value          = srow$x_value,
            std_curve_sample = sc_id,
            avg_MFI          = round(mfi_val, 1),
            n_reps           = srow$n_reps,
            est_conc_ug_mL   = if (is.na(est_conc))   NA_real_ else round(est_conc,   5),
            log10_conc       = if (is.na(log10_conc)) NA_real_ else round(log10_conc, 5),
            status           = status,
            stringsAsFactors = FALSE
          )
        }
      }
      if (length(rows_out) == 0) return(NULL)
      do.call(rbind, rows_out)
    })

    combined <- do.call(rbind, Filter(Negate(is.null), results_list))
    if (is.null(combined) || nrow(combined) == 0) return(NULL)
    combined[order(combined$analyte, combined$sample_id,
                   combined$x_value, combined$std_curve_sample), ]
  })

  output$sc_sample_conc_table <- renderDT({
    tbl <- sc_sample_conc(); req(!is.null(tbl) && nrow(tbl) > 0)
    row.names(tbl) <- NULL

    dt <- datatable(
      tbl,
      filter  = "top",
      options = list(pageLength = 20, scrollX = TRUE, dom = "lfrtip",
                     columnDefs = list(list(className = "dt-center", targets = "_all"))),
      rownames = FALSE,
      class    = "stripe hover cell-border compact",
      colnames = c("Analyte", "Sample ID", "x Value (Conc/Dil)", "Std Curve Sample",
                   "Avg MFI", "N Reps",
                   "Est. Conc. (\u00b5g/mL) [linear]",
                   "Est. Conc. [log\u2081\u2080]",
                   "Status")
    ) %>%
      formatRound(columns = c("est_conc_ug_mL", "log10_conc"), digits = 4) %>%
      formatStyle(
        "status",
        color      = styleEqual(
          c("> upper limit", "< lower limit", "OK", "No MFI data"),
          c("#c0392b",       "#888888",       "#27ae60", "#888888")),
        fontWeight = "bold"
      ) %>%
      formatStyle(
        "avg_MFI",
        background         = styleColorBar(
          c(0, max(tbl$avg_MFI[!is.na(tbl$avg_MFI)], 1)), "#b3d9ff"),
        backgroundSize     = "98% 88%",
        backgroundRepeat   = "no-repeat",
        backgroundPosition = "center"
      )
    dt
  })

  output$sc_sample_conc_dl <- downloadHandler(
    filename = function() {
      ag <- gsub("[^A-Za-z0-9_]", "_", input$sc_analyte %||% "analyte")
      paste0("sample_conc_estimates_", ag, "_", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      tbl <- sc_sample_conc(); req(!is.null(tbl))
      colnames(tbl) <- c("Analyte", "Sample ID", "x Value (Conc/Dil)", "Std Curve Sample",
                         "Avg MFI", "N Reps",
                         "Est. Conc. (ug/mL) [linear]",
                         "Est. Conc. [log10]",
                         "Status")
      openxlsx::write.xlsx(tbl, file, rowNames = FALSE)
    }
  )

  output$pb_qc4_result <- renderUI({
    r <- pb_qc4()
    overall_badge <- .qc_badge(
      r$pass,
      if (isTRUE(r$pass))  "All evaluated QC checks PASS -- System is suitable."
      else if (is.na(r$pass)) "No QC data available yet."
      else "One or more QC checks require attention -- Review individual QC sections above."
    )
    summary_rows <- lapply(names(r$results), function(nm) {
      val   <- r$results[[nm]]
      color <- if (is.na(val)) "#888" else if (isTRUE(val)) "#27ae60" else "#e6a817"
      lbl   <- if (is.na(val)) "N/A"  else if (isTRUE(val)) "PASS"    else "CAUTION"
      tags$div(
        style = "display:inline-flex; align-items:center; gap:8px;
                 background:#f8f9fa; border-radius:6px; padding:6px 14px;
                 margin:4px; font-size:13px;",
        tags$span(style = paste0("color:", color, "; font-weight:700;"), lbl),
        tags$span(style = "color:#333;", nm)
      )
    })
    tags$div(overall_badge, tags$div(style = "margin-top:12px;", summary_rows))
  })

  # -- Top scorecard row (4 cards) --------------------------------------------
  output$pb_qc_scorecards <- renderUI({
    ag_lbl <- if (!is.null(input$pb_selected_antigen))
      gsub("\\s*\\(\\d+\\)\\s*$", "", input$pb_selected_antigen) else "--"
    lod <- pb_cfg()$lod
    loq <- pb_cfg()$loq
    cv  <- pb_cfg()$cv

    .card <- function(num, title, subtitle, pass_val, criteria_txt) {
      color <- if (is.na(pass_val)) "#888" else if (isTRUE(pass_val)) "#27ae60" else "#e6a817"
      badge <- if (is.na(pass_val)) "N/A"  else if (isTRUE(pass_val)) "PASS"    else "CAUTION"
      bg    <- if (is.na(pass_val)) "#f8f9fa" else if (isTRUE(pass_val)) "#f0fff4" else "#fffbea"
      tags$div(
        style = paste0("background:", bg, "; border:2px solid ", color, ";
                        border-radius:10px; padding:12px 10px; text-align:center;
                        min-width:130px; flex:1;"),
        tags$div(style = "font-size:10px; color:#888; font-weight:700;
                          text-transform:uppercase; letter-spacing:0.5px; margin-bottom:2px;",
          paste0(num, ". ", title)),
        tags$div(style = paste0("font-size:18px; font-weight:800; color:", color, ";"), badge),
        tags$div(style = "font-size:11px; color:#555; margin-top:4px;", subtitle),
        tags$div(style = "font-size:10px; color:#777; margin-top:3px; font-style:italic;",
          criteria_txt)
      )
    }

    q1 <- pb_qc1(); q2 <- pb_qc2(); q3 <- pb_qc3()

    tags$div(
      tags$div(style = "font-size:12px; color:#555; margin-bottom:10px;",
        icon("tag"), tags$strong(" Antigen: "), ag_lbl,
        tags$span(style = "margin-left:16px;",
          icon("ruler"), tags$strong(" LOD: "), lod),
        tags$span(style = "margin-left:12px;",
          icon("ruler-horizontal"), tags$strong(" LOQ: "), loq),
        tags$span(style = "margin-left:12px;",
          icon("percent"), tags$strong(" CV Acceptance: "), paste0(cv, "%"))
      ),
      tags$div(
        style = "display:flex; gap:8px; flex-wrap:wrap;",
        .card("1", "Bead Acquisition",   "\u2265 50 beads / well",   q1$pass, "Beads \u2265 50"),
        .card("2", "Negative Controls",  "(Blank & Blank Well MFI)", q2$pass, "Blank MFI \u2264 LOD"),
        .card("3", "Positive Controls",  "(Low & High Pos Ctrl)",    q3$pass, "MFI within range")
      )
    )
  })

  # ---------------------------------------------------------------------------
  # Per-tab refresh buttons -- reload the session when clicked
  # ---------------------------------------------------------------------------
  refresh_ids <- c("refresh_upload", "refresh_helper", "refresh_overview",
                   "refresh_review", "refresh_export", "refresh_dataframe",
                   "refresh_point", "refresh_titration", "refresh_quant")
  for (.rid in refresh_ids) {
    local({
      rid <- .rid
      observeEvent(input[[rid]], { session$reload() }, ignoreNULL = TRUE, ignoreInit = TRUE)
    })
  }

}

shinyApp(ui, server)
