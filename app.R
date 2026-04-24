packages <- c(
  "shiny",
  "shinyWidgets",
  "openxlsx",
  "dplyr",
  "tidyr",
  "ggplot2",
  "plotly",
  "DT",
  "stringr",
  "scales",
  "sf",
  "rlang"
)

missing_packages <- setdiff(packages, rownames(installed.packages()))
if (length(missing_packages) > 0) {
  stop(
    "Install the missing packages before running the dashboard: ",
    paste(missing_packages, collapse = ", ")
  )
}

invisible(lapply(setdiff(packages, "plotly"), library, character.only = TRUE))
options(scipen = 999)

`%||%` <- function(x, y) {
  if (is.null(x) || length(x) == 0 || all(is.na(x))) y else x
}

resolve_existing_path <- function(candidates) {
  hit <- candidates[file.exists(candidates)]
  if (length(hit) == 0) {
    return(NA_character_)
  }
  hit[[1]]
}

data_path <- resolve_existing_path(c("output_updated.xlsx", "output.xlsx", "data/output.xlsx"))
if (is.na(data_path)) {
  stop("Could not find output_updated.xlsx or output.xlsx in the project root or in a data/ folder.")
}

map_path <- resolve_existing_path(c("map.gpkg", "data/map.gpkg"))

dataset_keys <- c("msz_pre", "msz_act", "zorgkos", "huisart", "medicij", "msz")

dataset_labels <- c(
  msz_pre = "MSZ prestaties",
  msz_act = "MSZ activiteiten",
  zorgkos = "Zorgkosten",
  huisart = "Huisartszorg",
  medicij = "Medicijnen",
  msz = "MSZ"
)

condition_labels <- c(
  diabetes = "Diabetes",
  migraine = "Migraine",
  schildkl = "Schildklier",
  schildkl_a = "Schildklier A",
  schildkl_b = "Schildklier B"
)

type_labels <- c(
  gemiddelde_per_persoon = "Gemiddelde per persoon",
  gemiddelde_per_gebruiker = "Gemiddelde per gebruiker",
  mediaan_per_persoon = "Mediaan per persoon",
  mediaan_per_gebruiker = "Mediaan per gebruiker",
  n_totaal_gebruikers = "Aantal gebruikers",
  sum_totaal_groep = "Totale som"
)

family_labels <- c(
  heeft = "Gebruik / prevalentie",
  kosten = "Kosten",
  n = "Aantal contacten"
)

sex_palette <- c(
  Mannen = "#4A96CF",
  Vrouwen = "#A878C9"
)

profile_palette_base <- c(
  with_condition = "#C47C4E",
  without_condition = "#5D8F73"
)

pretty_default <- function(x) {
  x |>
    stringr::str_replace_all("_", " ") |>
    stringr::str_squish() |>
    stringr::str_to_title()
}

pretty_dataset <- function(x) {
  unname(dataset_labels[x] %||% pretty_default(x))
}

pretty_condition <- function(x) {
  if (is.na(x) || !nzchar(x)) {
    return("Alle personen")
  }
  unname(condition_labels[x] %||% pretty_default(x))
}

pretty_type <- function(x) {
  unname(type_labels[x] %||% pretty_default(x))
}

pretty_binary <- function(x) {
  x_chr <- as.character(x)
  dplyr::case_when(
    x_chr %in% c("1", "Ja", "TRUE", "true") ~ "Ja",
    x_chr %in% c("0", "Nee", "FALSE", "false") ~ "Nee",
    TRUE ~ x_chr
  )
}

pretty_metric_name <- function(x) {
  out <- x
  out <- stringr::str_replace(out, "^gebruikt_", "gebruik ")
  out <- stringr::str_replace(out, "^heeft_", "")
  out <- stringr::str_replace(out, "^kosten_", "kosten ")
  out <- stringr::str_replace(out, "^n_", "aantal ")
  out <- stringr::str_replace_all(out, "_", " ")
  out <- stringr::str_squish(out)
  out <- stringr::str_to_title(out)
  out <- stringr::str_replace_all(out, "Poh Ggz", "POH GGZ")
  out <- stringr::str_replace_all(out, "Ggz", "GGZ")
  out <- stringr::str_replace_all(out, "Msz", "MSZ")
  out
}

pretty_profile_condition <- function(x, condition_label = NA_character_) {
  dplyr::case_when(
    as.character(x) == "1" ~ "Met aandoening",
    as.character(x) == "0" ~ "Zonder aandoening",
    TRUE ~ as.character(x)
  )
}

family_key_from_metric <- function(x) {
  dplyr::case_when(
    stringr::str_starts(x, "heeft_") ~ "heeft",
    stringr::str_starts(x, "kosten_") ~ "kosten",
    stringr::str_starts(x, "n_") ~ "n",
    TRUE ~ "other"
  )
}

pretty_family <- function(x) {
  unname(family_labels[x] %||% pretty_default(x))
}

ordered_area_levels <- function(x) {
  values <- unique(as.character(x))
  preferred <- c("Amsterdam", "NL")
  c(intersect(preferred, values), setdiff(sort(values), preferred))
}

ordered_sex_levels <- function(x) {
  values <- unique(as.character(x))
  preferred <- c("Mannen", "Vrouwen")
  c(intersect(preferred, values), setdiff(sort(values), preferred))
}

condition_status_label <- function(x, condition_label = NA_character_) {
  x_bin <- pretty_binary(x)

  if (is.na(condition_label) || !nzchar(condition_label) || identical(condition_label, "Alle personen")) {
    return(x_bin)
  }

  dplyr::case_when(
    x_bin == "Ja" ~ paste("Met", condition_label),
    x_bin == "Nee" ~ paste("Zonder", condition_label),
    TRUE ~ x_bin
  )
}

condition_linetype_values <- function(labels) {
  stats::setNames(
    ifelse(startsWith(labels, "Zonder") | labels %in% c("Nee", "No"), "dashed", "solid"),
    labels
  )
}

palette_for_values <- function(values) {
  values <- unique(as.character(values))
  if (length(values) == 0) {
    return(character(0))
  }
  if (all(values %in% names(sex_palette))) {
    return(sex_palette[values])
  }
  stats::setNames(build_palette(length(values)), values)
}

profile_palette_for_labels <- function(labels) {
  labels <- unique(as.character(labels))
  fallback <- build_palette(length(labels))
  stats::setNames(
    vapply(seq_along(labels), function(i) {
      label <- labels[[i]]
      if (startsWith(label, "Met ")) {
        profile_palette_base[["with_condition"]]
      } else if (startsWith(label, "Zonder ")) {
        profile_palette_base[["without_condition"]]
      } else {
        fallback[[i]]
      }
    }, character(1)),
    labels
  )
}

sanitize_filename <- function(x) {
  gsub("[^A-Za-z0-9_-]+", "_", x)
}

build_export_name <- function(...) {
  parts <- c(...)
  parts <- parts[!is.na(parts) & nzchar(parts)]
  sanitize_filename(paste(parts, collapse = "_"))
}

compose_title_parts <- function(...) {
  parts <- c(...)
  parts <- as.character(parts)
  parts <- parts[!is.na(parts) & nzchar(parts)]
  paste(parts, collapse = " | ")
}

choose_preferred_choice <- function(choices, current = NULL, fallback = NULL) {
  choices <- unique(as.character(choices))
  if (length(choices) == 0) {
    return(character(0))
  }
  if (!is.null(current) && nzchar(current) && current %in% choices) {
    return(current)
  }
  if (!is.null(fallback) && nzchar(fallback) && fallback %in% choices) {
    return(fallback)
  }
  choices[[1]]
}

first_or_empty <- function(x) {
  if (length(x) == 0) {
    return(character(0))
  }
  x[[1]]
}

normalize_hex_color <- function(value, fallback) {
  if (is.null(value) || !nzchar(value)) {
    return(fallback)
  }

  value <- as.character(value)[[1]]
  if (!grepl("^#[0-9A-Fa-f]{6}$", value)) {
    return(fallback)
  }

  value
}

color_picker_input <- function(id, label, value) {
  shinyWidgets::colorPickr(
    inputId = id,
    label = label,
    selected = value,
    swatches = NULL,
    preview = FALSE,
    hue = TRUE,
    opacity = FALSE,
    interaction = list(
      cancel = FALSE,
      clear = FALSE,
      save = FALSE,
      hex = FALSE,
      rgba = FALSE,
      input = FALSE
    ),
    theme = "classic",
    update = "changestop",
    position = "bottom-middle",
    hideOnSave = TRUE,
    useAsButton = FALSE,
    pickr_width = "240px",
    width = "100%"
  )
}

apply_plotly_transparent_layout <- function(plot_obj, ...) {
  plotly::layout(
    plot_obj,
    paper_bgcolor = "rgba(0,0,0,0)",
    plot_bgcolor = "rgba(0,0,0,0)",
    ...
  ) |>
    plotly::config(displayModeBar = FALSE, displaylogo = FALSE)
}

apply_plotly_y_range <- function(plot_obj, lower, upper) {
  if (is.null(upper) || !is.finite(upper)) {
    return(plot_obj)
  }

  layout_names <- names(plot_obj$x$layout %||% list())
  yaxis_names <- layout_names[stringr::str_detect(layout_names, "^yaxis[0-9]*$")]

  if (length(yaxis_names) == 0) {
    yaxis_names <- "yaxis"
  }

  for (axis_name in yaxis_names) {
    axis <- plot_obj$x$layout[[axis_name]] %||% list()
    axis$range <- c(lower, upper)
    axis$autorange <- FALSE
    plot_obj$x$layout[[axis_name]] <- axis
  }

  plot_obj
}

format_script_goal <- function(path) {
  lines <- readLines(path, warn = FALSE, n = 12)
  goal_line <- lines[stringr::str_detect(lines, "^# Goal:")]
  if (length(goal_line) == 0) {
    return("")
  }
  stringr::str_trim(sub("^# Goal:\\s*", "", goal_line[[1]]))
}

parse_dataset_context <- function(base_name) {
  matched_key <- dataset_keys[
    vapply(
      dataset_keys,
      function(key) identical(base_name, key) || startsWith(base_name, paste0(key, "_")),
      logical(1)
    )
  ][1]

  if (is.na(matched_key)) {
    matched_key <- base_name
  }

  condition_key <- if (identical(base_name, matched_key)) {
    NA_character_
  } else {
    sub(paste0("^", matched_key, "_"), "", base_name)
  }

  data.frame(
    dataset_key = matched_key,
    dataset_label = pretty_dataset(matched_key),
    condition_key = condition_key,
    condition_label = pretty_condition(condition_key),
    stringsAsFactors = FALSE
  )
}

build_sheet_index <- function(sheets, parser) {
  rows <- lapply(sheets, function(sheet) {
    sheet_info <- parser(sheet)
    context <- parse_dataset_context(sheet_info$base_name)
    view_key <- sheet_info$view_key %||% NA_character_
    view_label <- sheet_info$view_label %||% NA_character_
    base_label <- if (is.na(context$condition_key)) {
      paste(context$dataset_label, "| Alle personen |", sheet_info$period_label)
    } else {
      paste(context$dataset_label, "|", context$condition_label, "|", sheet_info$period_label)
    }

    transform(
      context,
      sheet = sheet,
      base_name = sheet_info$base_name,
      period_key = sheet_info$period_key,
      period_label = sheet_info$period_label,
      view_key = view_key,
      view_label = view_label,
      sheet_label = if (!is.na(view_label) && nzchar(view_label)) {
        paste(base_label, "|", view_label)
      } else {
        base_label
      }
    )
  })

  dplyr::bind_rows(rows) |>
    dplyr::arrange(dataset_label, condition_label, period_label, view_label, sheet)
}

is_yearly_sheet <- function(sheet) {
  stringr::str_detect(sheet, "_yr$")
}

parse_yearly_sheet <- function(sheet) {
  if (stringr::str_detect(sheet, "_yr$")) {
    return(list(base_name = sub("_yr$", "", sheet), period_key = "all_years", period_label = "Alle jaren"))
  }

  stop("Unsupported yearly sheet: ", sheet)
}

is_es_mean_sheet <- function(sheet) {
  stringr::str_detect(sheet, "_es$")
}

parse_es_mean_sheet <- function(sheet) {
  if (stringr::str_detect(sheet, "_es$")) {
    return(list(
      base_name = sub("_es$", "", sheet),
      period_key = "all_years",
      period_label = "Alle jaren"
    ))
  }

  stop("Unsupported event-study means sheet: ", sheet)
}

is_snapshot_level_sheet <- function(sheet) {
  stringr::str_detect(sheet, "_23$") &
    !stringr::str_detect(sheet, "^profile_") &
    !stringr::str_detect(sheet, "_es_diff_23$") &
    !stringr::str_detect(sheet, "_es_23$")
}

parse_snapshot_level_sheet <- function(sheet) {
  if (stringr::str_detect(sheet, "_23$")) {
    return(list(
      base_name = sub("_23$", "", sheet),
      period_key = "2023",
      period_label = "2023"
    ))
  }

  stop("Unsupported 2023 levels sheet: ", sheet)
}

parse_es_ci_sheet <- function(sheet) {
  if (stringr::str_detect(sheet, "_es_ci$")) {
    return(list(base_name = sub("_es_ci$", "", sheet), period_key = "all_years", period_label = "Alle jaren"))
  }

  stop("Unsupported estimated-effects sheet: ", sheet)
}

sheet_names <- openxlsx::getSheetNames(data_path)

yearly_index <- build_sheet_index(sheet_names[is_yearly_sheet(sheet_names)], parse_yearly_sheet)
es_mean_index <- build_sheet_index(sheet_names[is_es_mean_sheet(sheet_names)], parse_es_mean_sheet)
es_ci_index <- build_sheet_index(sheet_names[stringr::str_detect(sheet_names, "_es_ci$")], parse_es_ci_sheet)
snapshot_level_index <- build_sheet_index(
  sheet_names[is_snapshot_level_sheet(sheet_names)],
  parse_snapshot_level_sheet
)

yearly_index <- yearly_index |>
  dplyr::mutate(
    sheet_label = ifelse(
      is.na(condition_key),
      paste(dataset_label, "| Alle personen"),
      paste(dataset_label, "|", condition_label)
    )
  )

es_mean_index <- es_mean_index |>
  dplyr::mutate(
    sheet_label = ifelse(
      is.na(condition_key),
      paste(dataset_label, "| Alle personen"),
      paste(dataset_label, "|", condition_label)
    )
  )

es_ci_index <- es_ci_index |>
  dplyr::mutate(
    sheet_label = ifelse(
      is.na(condition_key),
      paste(dataset_label, "| Alle personen"),
      paste(dataset_label, "|", condition_label)
    )
  )

snapshot_level_index <- snapshot_level_index |>
  dplyr::mutate(
    sheet_label = ifelse(
      is.na(condition_key),
      paste(dataset_label, "| Alle personen"),
      paste(dataset_label, "|", condition_label)
    )
  )

profile_index <- {
  profile_sheets <- sheet_names[stringr::str_detect(sheet_names, "^profile_")]
  rows <- lapply(profile_sheets, function(sheet) {
    match <- stringr::str_match(sheet, "^profile_(.*)_(before|after)_(match|23)$")
    condition_key <- match[, 2]
    stage_key <- match[, 3]
    period_key <- match[, 4]
    data.frame(
      sheet = sheet,
      base_id = paste(condition_key, stage_key, sep = "__"),
      condition_key = condition_key,
      condition_label = pretty_condition(condition_key),
      stage_key = stage_key,
      period_key = period_key,
      period_label = dplyr::case_when(
        period_key == "23" ~ "2023",
        TRUE ~ "Alle jaren"
      ),
      stage_label = dplyr::case_when(
        stage_key == "before" ~ "Voor matching",
        stage_key == "after" ~ "Na matching",
        TRUE ~ pretty_default(stage_key)
      ),
      sheet_label = paste(pretty_condition(condition_key), "|", dplyr::case_when(
        stage_key == "before" ~ "Voor matching",
        stage_key == "after" ~ "Na matching",
        TRUE ~ pretty_default(stage_key)
      ), "|", dplyr::case_when(
        period_key == "23" ~ "2023",
        TRUE ~ "Alle jaren"
      )),
      stringsAsFactors = FALSE
    )
  })
  dplyr::bind_rows(rows) |>
    dplyr::arrange(condition_label, period_label, stage_key)
}

profile_base_index <- profile_index |>
  dplyr::distinct(base_id, condition_label, stage_label) |>
  dplyr::mutate(base_label = paste(condition_label, "|", stage_label)) |>
  dplyr::arrange(condition_label, stage_label)

snapshot_base_index <- snapshot_level_index |>
  dplyr::distinct(base_name, dataset_label, condition_key, condition_label) |>
  dplyr::mutate(
    base_label = ifelse(
      is.na(condition_key),
      paste(dataset_label, "| Alle personen"),
      paste(dataset_label, "|", condition_label)
    )
  ) |>
  dplyr::arrange(dataset_label, condition_label)

script_files <- list.files("code", pattern = "\\.R$", full.names = TRUE)
script_index <- dplyr::bind_rows(lapply(script_files, function(path) {
  info <- file.info(path)
  data.frame(
    script = basename(path),
    goal = format_script_goal(path),
    modified = format(info$mtime, "%Y-%m-%d %H:%M"),
    stringsAsFactors = FALSE
  )
}))

cache_env <- new.env(parent = emptyenv())

read_sheet_cached <- function(sheet) {
  cache_key <- paste0("sheet__", sheet)
  if (!exists(cache_key, envir = cache_env, inherits = FALSE)) {
    df <- openxlsx::read.xlsx(data_path, sheet = sheet)
    assign(cache_key, df, envir = cache_env)
  }
  get(cache_key, envir = cache_env, inherits = FALSE)
}

condition_flag_col <- function(df) {
  known_cols <- c(
    "year", "years_since_diagnosis", "geslacht", "n_totaal", "variable",
    "value", "name", "type", "amsterdam", "se", "outcome", "sample",
    "t", "coef", "lo", "hi", "n", "n_met_conditie", "n_zonder_conditie"
  )
  candidates <- setdiff(names(df), known_cols)
  if (length(candidates) == 0) {
    return(NA_character_)
  }
  candidates[[1]]
}

metric_labeler <- function(metric_name, stat_type = NA_character_) {
  is_share_metric <- stringr::str_starts(metric_name, "heeft_") || stringr::str_starts(metric_name, "gebruikt_")
  is_count_type <- !is.na(stat_type) && stat_type %in% c("n_totaal_gebruikers", "sum_totaal_groep")

  if (is_share_metric && !is_count_type) {
    return(scales::label_percent(accuracy = 1, decimal.mark = ","))
  }
  scales::label_number(big.mark = ".", decimal.mark = ",")
}

axis_metric_labeler <- function(metric_name, stat_type = NA_character_) {
  base_labeler <- metric_labeler(metric_name, stat_type)

  if (!difference_metric_is_share(metric_name, stat_type)) {
    return(base_labeler)
  }

  function(x) {
    labels <- base_labeler(x)
    labels[is.finite(x) & x > 1] <- ""
    labels
  }
}

difference_metric_is_share <- function(metric_name, stat_type = NA_character_) {
  is_share_metric <- stringr::str_starts(metric_name, "heeft_") || stringr::str_starts(metric_name, "gebruikt_")
  is_count_type <- !is.na(stat_type) && stat_type %in% c("n_totaal_gebruikers", "sum_totaal_groep")
  isTRUE(is_share_metric && !is_count_type)
}

metric_axis_upper_limit <- function(metric_name, stat_type = NA_character_) {
  NULL
}

difference_metric_labeler <- function(metric_name, stat_type = NA_character_) {
  if (difference_metric_is_share(metric_name, stat_type)) {
    return(function(x) scales::number(x * 100, accuracy = 0.1, big.mark = ".", decimal.mark = ","))
  }

  metric_labeler(metric_name, stat_type)
}

compute_axis_lower_limit <- function(values, include_zero = FALSE, upper_limit = NULL) {
  values <- values[is.finite(values)]

  if (length(values) == 0) {
    lower <- 0
  } else {
    lower <- min(values, na.rm = TRUE)
  }

  if (isTRUE(include_zero)) {
    lower <- min(lower, 0)
  }

  if (!is.null(upper_limit) && (!is.finite(lower) || lower >= upper_limit)) {
    lower <- min(0, upper_limit - max(abs(upper_limit) * 0.05, 0.01))
  }

  lower
}

compute_axis_upper_bound <- function(values, lower_limit = NULL, upper_limit = NULL) {
  values <- values[is.finite(values)]

  if (length(values) == 0) {
    return(upper_limit %||% 1)
  }

  if (is.null(lower_limit) || !is.finite(lower_limit)) {
    lower_limit <- min(values, na.rm = TRUE)
  }

  upper_value <- max(values, na.rm = TRUE)
  span <- max(
    upper_value - lower_limit,
    abs(upper_value) * 0.12,
    if (!is.null(upper_limit) && is.finite(upper_limit)) upper_limit * 0.02 else 1e-06
  )

  upper_bound <- upper_value + 0.08 * span

  if (!is.null(upper_limit) && is.finite(upper_limit)) {
    upper_bound <- min(upper_bound, upper_limit)
  }

  if (!is.finite(upper_bound) || upper_bound <= lower_limit) {
    if (!is.null(upper_limit) && is.finite(upper_limit) && upper_limit > lower_limit) {
      return(upper_limit)
    }

    return(lower_limit + max(abs(lower_limit) * 0.05, 0.01))
  }

  upper_bound
}

compute_condition_difference <- function(df, condition_col) {
  if (is.na(condition_col) || !(condition_col %in% names(df))) {
    return(df[0, , drop = FALSE])
  }

  df <- df |>
    dplyr::mutate(.condition_status = pretty_binary(.data[[condition_col]]))

  if (!all(c("Ja", "Nee") %in% unique(df$.condition_status))) {
    return(df[0, setdiff(names(df), c(condition_col, ".condition_status")), drop = FALSE])
  }

  group_cols <- setdiff(names(df), c(condition_col, ".condition_status", "value", "n_totaal"))

  df |>
    dplyr::group_by(dplyr::across(dplyr::all_of(group_cols))) |>
    dplyr::summarise(
      n_met_conditie = dplyr::first(n_totaal[.condition_status == "Ja"]),
      n_zonder_conditie = dplyr::first(n_totaal[.condition_status == "Nee"]),
      value = dplyr::first(value[.condition_status == "Ja"]) - dplyr::first(value[.condition_status == "Nee"]),
      .groups = "drop"
    )
}

display_table <- function(df, export_name = "gegevens", show_excel = TRUE) {
  table_options <- list(
    dom = if (isTRUE(show_excel)) "Bfrt" else "frt",
    buttons = if (isTRUE(show_excel)) list(list(
      extend = "excel",
      text = "Gegevens downloaden",
      filename = export_name,
      title = ""
    )) else list(),
    paging = FALSE,
    scrollX = TRUE
  )

  table_args <- list(
    data = df,
    rownames = FALSE,
    filter = "top",
    options = table_options
  )

  if (isTRUE(show_excel)) {
    table_args$extensions <- "Buttons"
  }

  do.call(DT::datatable, table_args)
}

build_palette <- function(n) {
  colorRampPalette(
    c("#b0145b", "#1f6e8c", "#d67c25", "#4b7f52", "#745c97", "#8c564b")
  )(max(1, n))
}

theme_dashboard <- function() {
  ggplot2::theme_minimal(base_size = 14) +
    ggplot2::theme(
      plot.title = element_text(face = "bold", color = "#1f3c44", size = 17),
      plot.subtitle = element_text(color = "#576574", margin = margin(b = 12)),
      panel.grid.minor = element_blank(),
      panel.grid.major.x = element_blank(),
      panel.spacing.x = grid::unit(0.3, "lines"),
      panel.spacing.y = grid::unit(0.4, "lines"),
      legend.position = "bottom",
      legend.title = element_blank(),
      panel.background = element_rect(fill = "transparent", colour = NA),
      plot.background = element_rect(fill = "transparent", colour = NA),
      legend.background = element_rect(fill = "transparent", colour = NA),
      legend.key = element_rect(fill = "transparent", colour = NA),
      strip.background = element_rect(fill = "transparent", colour = NA),
      rect = element_rect(fill = "transparent", colour = NA),
      axis.title = element_blank(),
      axis.text.x = element_text(color = "#36454f"),
      axis.text.y = element_text(color = "#36454f")
    )
}

build_series_label <- function(df, condition_col = NA_character_, condition_label = NA_character_, include_area = TRUE) {
  parts <- list()

  if ("geslacht" %in% names(df) && dplyr::n_distinct(df$geslacht) > 1) {
    parts[[length(parts) + 1]] <- paste("Geslacht:", df$geslacht)
  }

  if (isTRUE(include_area) && "amsterdam" %in% names(df) && dplyr::n_distinct(df$amsterdam) > 1) {
    parts[[length(parts) + 1]] <- paste("Gebied:", df$amsterdam)
  }

  if (!is.na(condition_col) && condition_col %in% names(df) && dplyr::n_distinct(df[[condition_col]]) > 1) {
    prefix <- if (is.na(condition_label) || !nzchar(condition_label)) {
      "Status"
    } else {
      condition_label
    }
    parts[[length(parts) + 1]] <- paste(prefix, pretty_binary(df[[condition_col]]), sep = ": ")
  }

  if (length(parts) == 0) {
    return(rep("Waarde", nrow(df)))
  }

  series_matrix <- do.call(cbind, parts)
  apply(series_matrix, 1, paste, collapse = " | ")
}

prepare_time_series_aesthetics <- function(df, condition_col = NA_character_, condition_label = NA_character_) {
  has_sex <- "geslacht" %in% names(df) && all(unique(as.character(stats::na.omit(df$geslacht))) %in% names(sex_palette))
  has_condition <- !is.na(condition_col) &&
    condition_col %in% names(df) &&
    dplyr::n_distinct(df[[condition_col]]) > 1

  if (has_sex) {
    df$color_value <- factor(as.character(df$geslacht), levels = ordered_sex_levels(df$geslacht))
  } else if (has_condition) {
    status_labels <- condition_status_label(df[[condition_col]], condition_label)
    df$color_value <- factor(status_labels, levels = unique(condition_status_label(c("Nee", "Ja"), condition_label)))
  } else {
    df$color_value <- factor(rep("Waarde", nrow(df)))
  }

  if (has_condition) {
    status_labels <- condition_status_label(df[[condition_col]], condition_label)
    df$linetype_value <- factor(
      status_labels,
      levels = unique(condition_status_label(c("Nee", "Ja"), condition_label))
    )
  }

  group_parts <- list(as.character(df$color_value))
  if ("linetype_value" %in% names(df)) {
    group_parts[[length(group_parts) + 1]] <- as.character(df$linetype_value)
  }

  df$plot_group <- interaction(group_parts, drop = TRUE, lex.order = TRUE)
  df
}

build_time_series_plot <- function(
  df,
  x_var,
  title,
  subtitle,
  y_labeler,
  axis_labeler = y_labeler,
  vline = NULL,
  facet_var = NULL,
  compact_facets = FALSE,
  show_confidence = FALSE,
  include_zero = FALSE,
  y_max = NULL
) {
  color_col <- if ("color_value" %in% names(df)) "color_value" else "series"
  group_col <- if ("plot_group" %in% names(df)) "plot_group" else color_col
  has_linetype <- "linetype_value" %in% names(df) && dplyr::n_distinct(df$linetype_value) > 1
  has_confidence <- isTRUE(show_confidence) && "se" %in% names(df) && any(is.finite(df$se))
  color_levels <- if (is.factor(df[[color_col]])) levels(df[[color_col]]) else unique(as.character(df[[color_col]]))
  color_levels <- color_levels[color_levels %in% unique(as.character(df[[color_col]]))]
  palette_values <- palette_for_values(color_levels)

  if (has_confidence) {
    df <- df |>
      dplyr::mutate(
        ci_low = value - 1.96 * se,
        ci_high = value + 1.96 * se
      )
  }

  df$tooltip <- paste0(
    "Reeks: ", df$series, "<br>",
    if (identical(x_var, "year")) "Jaar" else "Jaren sinds diagnose", ": ", df[[x_var]], "<br>",
    "Waarde: ", y_labeler(df$value)
  )

  if (has_confidence) {
    df$tooltip <- paste0(
      df$tooltip,
      "<br>95%-BI: ",
      y_labeler(df$ci_low),
      " tot ",
      y_labeler(df$ci_high)
    )
  }

  p <- ggplot(df, aes(
    x = .data[[x_var]],
    y = value,
    color = .data[[color_col]],
    group = .data[[group_col]],
    text = tooltip
  ))

  if (has_confidence) {
    p <- p +
      geom_ribbon(
        data = df,
        aes(
          x = .data[[x_var]],
          ymin = ci_low,
          ymax = ci_high,
          fill = .data[[color_col]],
          group = .data[[group_col]]
        ),
        inherit.aes = FALSE,
        alpha = 0.12,
        color = NA,
        show.legend = FALSE
      )
  }

  if (has_linetype) {
    linetype_levels <- if (is.factor(df$linetype_value)) levels(df$linetype_value) else unique(as.character(df$linetype_value))
    linetype_levels <- linetype_levels[linetype_levels %in% unique(as.character(df$linetype_value))]
    p <- p +
      geom_line(aes(linetype = linetype_value), linewidth = 1.1) +
      geom_point(size = 2.5, show.legend = FALSE) +
      scale_linetype_manual(values = condition_linetype_values(linetype_levels))
  } else {
    p <- p +
      geom_line(linewidth = 1.1) +
      geom_point(size = 2.5, show.legend = FALSE)
  }

  p <- p +
    scale_color_manual(values = palette_values) +
    {if (has_confidence) scale_fill_manual(values = palette_values, guide = "none")} +
    scale_x_continuous(breaks = sort(unique(df[[x_var]]))) +
    scale_y_continuous(labels = axis_labeler) +
    labs(title = title, subtitle = subtitle, color = NULL, linetype = NULL) +
    theme_dashboard()

  if (isTRUE(include_zero)) {
    p <- p + expand_limits(y = 0)
  }

  if (!is.null(y_max)) {
    y_values <- if (has_confidence) c(df$ci_low, df$ci_high) else df$value
    y_lower <- compute_axis_lower_limit(y_values, include_zero = include_zero, upper_limit = y_max)
    y_upper <- compute_axis_upper_bound(y_values, lower_limit = y_lower, upper_limit = y_max)
    p <- p + coord_cartesian(
      ylim = c(
        y_lower,
        y_upper
      )
    )
  }

  if (!is.null(vline)) {
    p <- p + geom_vline(xintercept = vline, color = "#b0145b", linetype = "dashed", linewidth = 0.7)
  }

  if (!is.null(facet_var) && facet_var %in% names(df) && dplyr::n_distinct(df[[facet_var]]) > 1) {
    p <- p +
      facet_wrap(stats::as.formula(paste("~", facet_var)), nrow = 1) +
      theme(
        panel.spacing.x = if (isTRUE(compact_facets)) grid::unit(0.05, "lines") else grid::unit(0.3, "lines"),
        panel.spacing.y = if (isTRUE(compact_facets)) grid::unit(0.15, "lines") else grid::unit(0.4, "lines")
      )
  }

  p
}

build_snapshot_plot <- function(
  level_df,
  title,
  subtitle,
  y_labeler,
  axis_labeler = y_labeler,
  diff_labeler = y_labeler,
  diff_df = NULL,
  y_max = NULL
) {
  condition_col <- condition_flag_col(level_df)
  has_condition <- !is.na(condition_col) &&
    condition_col %in% names(level_df) &&
    dplyr::n_distinct(level_df[[condition_col]]) > 1
  sex_values <- unique(as.character(stats::na.omit(level_df$geslacht %||% character(0))))
  has_sex <- "geslacht" %in% names(level_df) &&
    length(sex_values) > 0 &&
    all(sex_values %in% names(sex_palette))
  time_text <- if ("year" %in% names(level_df) && dplyr::n_distinct(level_df$year) == 1) {
    paste0("<br>Jaar: ", level_df$year)
  } else if ("years_since_diagnosis" %in% names(level_df)) {
    paste0("<br>Jaren sinds diagnose: ", level_df$years_since_diagnosis)
  } else {
    ""
  }

  if (has_condition) {
    condition_levels <- unique(condition_status_label(c("Nee", "Ja")))

    level_plot_df <- level_df |>
      dplyr::mutate(
        condition_display = condition_status_label(.data[[condition_col]]),
        condition_legend = dplyr::case_when(
          condition_display == "Ja" ~ "Met aandoening",
          condition_display == "Nee" ~ "Zonder aandoening",
          TRUE ~ condition_display
        ),
        x_value = factor(condition_display, levels = condition_levels),
        fill_value = factor(condition_display, levels = condition_levels),
        tooltip = paste0(
          "Gebied: ", amsterdam,
          if (has_sex) paste0("<br>Geslacht: ", geslacht) else "",
          "<br>Diagnosestatus: ", condition_legend,
          time_text,
          if ("n_totaal" %in% names(level_df) && any(!is.na(level_df$n_totaal))) {
            paste0(
              "<br>n: ",
              scales::number(n_totaal, accuracy = 1, big.mark = ".", decimal.mark = ",")
            )
          } else {
            ""
          },
          "<br>Waarde: ", y_labeler(value)
        )
      )

    if (has_sex) {
      fill_values <- sex_palette
      level_plot_df$fill_value <- factor(as.character(level_plot_df$geslacht), levels = ordered_sex_levels(level_plot_df$geslacht))
      level_plot_df$alpha_value <- factor(
        level_plot_df$condition_legend,
        levels = c("Met aandoening", "Zonder aandoening")
      )
    } else {
      fill_values <- profile_palette_for_labels(condition_levels)
    }

    p <- ggplot(
      level_plot_df,
      aes(
        x = x_value,
        y = value,
        fill = fill_value,
        text = tooltip
      )
    )

    if (has_sex) {
      p <- p +
        geom_col(aes(alpha = alpha_value), width = 0.62, color = NA) +
        scale_alpha_manual(
          values = c("Met aandoening" = 1, "Zonder aandoening" = 0.45),
          name = NULL
        )
    } else {
      p <- p + geom_col(width = 0.62, color = NA)
    }

    p <- p +
      scale_fill_manual(values = fill_values) +
      scale_y_continuous(
        labels = axis_labeler,
        expand = ggplot2::expansion(mult = c(0.05, if (!is.null(diff_df) && nrow(diff_df) > 0) 0.22 else 0.08))
      ) +
      labs(
        title = title,
        subtitle = subtitle,
        fill = NULL
      ) +
      theme_dashboard() +
      expand_limits(y = 0)

    y_lower <- compute_axis_lower_limit(level_plot_df$value, include_zero = TRUE, upper_limit = y_max)
    y_upper <- compute_axis_upper_bound(level_plot_df$value, lower_limit = y_lower, upper_limit = y_max)

    if (!is.null(y_max) || (!is.null(diff_df) && nrow(diff_df) > 0)) {
      p <- p + coord_cartesian(
        ylim = c(
          y_lower,
          y_upper
        )
      )
    }

    if (!is.null(diff_df) && nrow(diff_df) > 0) {
      panel_keys <- c()
      if (has_sex) {
        panel_keys <- c(panel_keys, "geslacht")
      }
      if ("amsterdam" %in% names(level_plot_df) && dplyr::n_distinct(level_plot_df$amsterdam) > 1) {
        panel_keys <- c(panel_keys, "amsterdam")
      }

      if (length(panel_keys) > 0) {
        panel_max <- level_plot_df |>
          dplyr::group_by(dplyr::across(dplyr::all_of(panel_keys))) |>
          dplyr::summarise(
            panel_max = max(value, na.rm = TRUE),
            panel_min = min(value, na.rm = TRUE),
            .groups = "drop"
          )

        diff_labels <- diff_df |>
          dplyr::left_join(panel_max, by = panel_keys)
      } else {
        diff_labels <- diff_df |>
          dplyr::mutate(
            panel_max = max(level_plot_df$value, na.rm = TRUE),
            panel_min = min(level_plot_df$value, na.rm = TRUE)
          )
      }

      diff_labels <- diff_labels |>
        dplyr::mutate(
          span = pmax(panel_max - pmin(panel_min, 0), abs(panel_max) * 0.12, 1e-06),
          available_upper = pmax(y_upper - 0.02 * span, panel_max),
          available_gap = pmax(available_upper - panel_max, 0),
          bracket_y = dplyr::if_else(
            available_gap < 0.05 * span,
            panel_max - 0.04 * span,
            pmin(panel_max + 0.05 * span, panel_max + 0.55 * available_gap)
          ),
          label_y = dplyr::if_else(
            available_gap < 0.05 * span,
            bracket_y + 0.16 * span,
            pmin(bracket_y + 0.1 * span, panel_max + 0.92 * available_gap)
          ),
          diff_label = diff_labeler(value),
          x_left = 1,
          x_right = 2,
          x_mid = 1.5
        )

      p <- p +
        geom_segment(
          data = diff_labels,
          aes(x = x_left, xend = x_right, y = bracket_y, yend = bracket_y),
          inherit.aes = FALSE,
          color = "#4b5563",
          linewidth = 0.5
        ) +
        geom_segment(
          data = diff_labels,
          aes(x = x_left, xend = x_left, y = bracket_y, yend = bracket_y - 0.03 * span),
          inherit.aes = FALSE,
          color = "#4b5563",
          linewidth = 0.5
        ) +
        geom_segment(
          data = diff_labels,
          aes(x = x_right, xend = x_right, y = bracket_y, yend = bracket_y - 0.03 * span),
          inherit.aes = FALSE,
          color = "#4b5563",
          linewidth = 0.5
        ) +
        geom_text(
          data = diff_labels,
          aes(x = x_mid, y = label_y, label = diff_label),
          inherit.aes = FALSE,
          color = "#4b5563",
          size = 4,
          vjust = 0
        )
    }

    if (has_sex && "amsterdam" %in% names(level_plot_df) && dplyr::n_distinct(level_plot_df$amsterdam) > 1) {
      p <- p + facet_grid(geslacht ~ amsterdam)
    } else if (has_sex) {
      p <- p + facet_wrap(~geslacht, nrow = 1)
    } else if ("amsterdam" %in% names(level_plot_df) && dplyr::n_distinct(level_plot_df$amsterdam) > 1) {
      p <- p + facet_wrap(~amsterdam, nrow = 1)
    }

    return(p)
  }

  x_levels <- if (has_sex) ordered_sex_levels(level_df$geslacht) else "Waarde"
  fill_values <- if (has_sex) palette_for_values(x_levels) else c(Waarde = "#4A96CF")

  plot_df <- level_df |>
    dplyr::mutate(
      x_value = factor(if (has_sex) as.character(geslacht) else "Waarde", levels = x_levels),
      fill_value = factor(if (has_sex) as.character(geslacht) else "Waarde", levels = x_levels),
      tooltip = paste0(
        "Gebied: ", amsterdam,
        if (has_sex) paste0("<br>Geslacht: ", geslacht) else "",
        time_text,
        if ("n_totaal" %in% names(level_df) && any(!is.na(level_df$n_totaal))) {
          paste0(
            "<br>n: ",
            scales::number(n_totaal, accuracy = 1, big.mark = ".", decimal.mark = ",")
          )
        } else {
          ""
        },
        "<br>Waarde: ", y_labeler(value)
      )
    )

  p <- ggplot(
    plot_df,
    aes(
      x = x_value,
      y = value,
      fill = fill_value,
      text = tooltip
    )
  ) +
    geom_col(width = 0.62, color = NA) +
    scale_fill_manual(values = fill_values) +
    scale_y_continuous(labels = axis_labeler) +
    labs(
      title = title,
      subtitle = subtitle,
      fill = NULL
    ) +
    theme_dashboard() +
    expand_limits(y = 0)

  if (!is.null(y_max)) {
    y_lower <- compute_axis_lower_limit(plot_df$value, include_zero = TRUE, upper_limit = y_max)
    y_upper <- compute_axis_upper_bound(plot_df$value, lower_limit = y_lower, upper_limit = y_max)
    p <- p + coord_cartesian(
      ylim = c(
        y_lower,
        y_upper
      )
    )
  }

  if ("amsterdam" %in% names(plot_df) && dplyr::n_distinct(plot_df$amsterdam) > 1) {
    p <- p + facet_wrap(~amsterdam, nrow = 1)
  }

  if (!has_sex) {
    p <- p + theme(legend.position = "none")
  }

  p
}

make_line_only_legend <- function(plot_obj, order_mode = c("default", "sex_then_linetype")) {
  order_mode <- match.arg(order_mode)
  traces <- plot_obj$x$data %||% list()
  if (length(traces) == 0) {
    return(plot_obj)
  }

  normalize_legend_label <- function(label) {
    if (is.null(label) || !nzchar(label)) {
      return(label)
    }

    label <- as.character(label)
    if (grepl("^\\(.*\\)$", label)) {
      label <- sub("^\\((.*)\\)$", "\\1", label)
    }

    if (grepl("^[^,]+,\\d+$", label)) {
      label <- sub(",\\d+$", "", label)
    }

    label
  }

  normalize_trace <- function(trace) {
    trace$name <- normalize_legend_label(trace$name %||% "")
    if (!is.null(trace$legendgroup) && nzchar(trace$legendgroup)) {
      trace$legendgroup <- normalize_legend_label(trace$legendgroup)
    }
    if ((is.null(trace$legendgroup) || !nzchar(trace$legendgroup)) &&
        !is.null(trace$name) && nzchar(trace$name)) {
      trace$legendgroup <- trace$name
    }
    trace
  }

  plot_obj$x$data <- lapply(traces, normalize_trace)

  legend_traces <- Filter(function(trace) {
    mode <- trace$mode %||% ""
    !is.null(trace$name) &&
      nzchar(trace$name) &&
      (grepl("lines", mode) || !is.null(trace$line))
  }, plot_obj$x$data)

  if (length(legend_traces) == 0) {
    return(plot_obj)
  }

  legend_keys <- vapply(legend_traces, function(trace) {
    paste(
      trace$name %||% "",
      trace$legendgroup %||% "",
      trace$line$color %||% trace$marker$color %||% "",
      trace$line$dash %||% "solid",
      sep = "||"
    )
  }, character(1))
  legend_traces <- legend_traces[!duplicated(legend_keys)]

  if (identical(order_mode, "sex_then_linetype")) {
    sex_rank <- function(trace_name) {
      sex_name <- names(sex_palette)[vapply(names(sex_palette), startsWith, logical(1), x = trace_name %||% "")]
      if (length(sex_name) == 0) {
        return(length(sex_palette) + 1)
      }
      match(sex_name[[1]], names(sex_palette))
    }

    dash_rank <- function(trace) {
      dash_value <- trace$line$dash %||% "solid"
      match(dash_value, c("solid", "dash", "dot", "dashdot")) %||% 99
    }

    trace_order <- order(
      vapply(legend_traces, function(trace) sex_rank(trace$name), numeric(1)),
      vapply(legend_traces, dash_rank, numeric(1)),
      vapply(legend_traces, function(trace) trace$name %||% "", character(1))
    )
    legend_traces <- legend_traces[trace_order]
  }

  plot_obj$x$data <- lapply(plot_obj$x$data, function(trace) {
    trace$showlegend <- FALSE
    trace
  })

  for (trace in legend_traces) {
    plot_obj <- plotly::add_trace(
      plot_obj,
      x = 0,
      y = 0,
      type = "scatter",
      mode = "lines",
      visible = "legendonly",
      hoverinfo = "skip",
      showlegend = TRUE,
      name = trace$name,
      legendgroup = trace$legendgroup %||% trace$name,
      line = list(
        color = trace$line$color %||% trace$marker$color %||% "#4b5563",
        dash = trace$line$dash %||% "solid",
        width = trace$line$width %||% 2
      ),
      inherit = FALSE
    )
  }

  plot_obj
}

build_event_study_annotations <- function(df) {
  if (!("linetype_value" %in% names(df)) || dplyr::n_distinct(df$linetype_value) <= 1) {
    return(list())
  }

  sexes <- ordered_sex_levels(df$geslacht)
  sexes <- sexes[sexes %in% unique(as.character(df$geslacht))]
  if (length(sexes) == 0) {
    return(list())
  }

  linetype_labels <- if (is.factor(df$linetype_value)) levels(df$linetype_value) else unique(as.character(df$linetype_value))
  linetype_labels <- linetype_labels[linetype_labels %in% unique(as.character(df$linetype_value))]
  display_labels <- generic_condition_legend_label(linetype_labels)
  dash_map <- condition_linetype_values(linetype_labels)
  dash_order <- match(unname(dash_map), c("solid", "dashed", "dot", "dashdot"))
  dash_order[is.na(dash_order)] <- 99
  order_idx <- order(dash_order, linetype_labels)
  linetype_labels <- linetype_labels[order_idx]
  display_labels <- display_labels[order_idx]

  build_sample <- function(label, dash_value, color) {
    line_html <- if (identical(dash_value, "solid")) {
      "&#9473;&#9473;&#9473;"
    } else {
      "&#9473; &#9473; &#9473;"
    }
    paste0(
      "<span style='color:", color, ";'>", line_html, "</span>",
      "&nbsp;", label
    )
  }

  y_positions <- seq(-0.18, by = -0.09, length.out = length(sexes))

  lapply(seq_along(sexes), function(i) {
    sex <- sexes[[i]]
    color <- sex_palette[[sex]] %||% "#4b5563"
    row_text <- paste(
      vapply(seq_along(linetype_labels), function(j) {
        build_sample(
          display_labels[[j]],
          dash_map[[linetype_labels[[j]]]] %||% "solid",
          color
        )
      }, character(1)),
      collapse = "&nbsp;&nbsp;&nbsp;&nbsp;"
    )

    list(
      x = 0,
      y = y_positions[[i]],
      xref = "paper",
      yref = "paper",
      xanchor = "left",
      yanchor = "top",
      align = "left",
      showarrow = FALSE,
      text = paste0(
        "<span style='color:", color, "; font-weight:600;'>", sex, ":</span>",
        "&nbsp;&nbsp;",
        row_text
      ),
      font = list(size = 12, color = "#4b5563")
    )
  })
}

build_sex_annotations <- function(df) {
  sexes <- ordered_sex_levels(df$geslacht)
  sexes <- sexes[sexes %in% unique(as.character(df$geslacht))]
  if (length(sexes) == 0) {
    return(list())
  }

  row_text <- paste(
    vapply(sexes, function(sex) {
      color <- sex_palette[[sex]] %||% "#4b5563"
      paste0(
        "<span style='color:", color, ";'>&#9473;&#9473;&#9473;</span>",
        "&nbsp;", sex
      )
    }, character(1)),
    collapse = "&nbsp;&nbsp;&nbsp;&nbsp;"
  )

  list(
    list(
      x = 0,
      y = -0.11,
      xref = "paper",
      yref = "paper",
      xanchor = "left",
      yanchor = "top",
      align = "left",
      showarrow = FALSE,
      text = row_text,
      font = list(size = 12, color = "#4b5563")
    )
  )
}

css_rgba <- function(color, alpha = 1) {
  rgb_values <- grDevices::col2rgb(color)
  sprintf(
    "rgba(%d, %d, %d, %.3f)",
    rgb_values[1, 1],
    rgb_values[2, 1],
    rgb_values[3, 1],
    alpha
  )
}

generic_condition_legend_label <- function(x) {
  x_chr <- as.character(x)
  dplyr::case_when(
    stringr::str_detect(x_chr, "^Met\\b") | x_chr %in% c("Ja", "1", "TRUE", "true") ~ "Met aandoening",
    stringr::str_detect(x_chr, "^Zonder\\b") | x_chr %in% c("Nee", "0", "FALSE", "false") ~ "Zonder aandoening",
    TRUE ~ x_chr
  )
}

build_snapshot_condition_annotations <- function(df) {
  condition_col <- condition_flag_col(df)
  if (
    is.na(condition_col) ||
    !("geslacht" %in% names(df)) ||
    dplyr::n_distinct(df[[condition_col]]) <= 1
  ) {
    return(list())
  }

  sexes <- ordered_sex_levels(df$geslacht)
  sexes <- sexes[sexes %in% unique(as.character(df$geslacht))]
  if (length(sexes) == 0) {
    return(list())
  }

  status_labels <- c("Met aandoening", "Zonder aandoening")
  alpha_values <- c("Met aandoening" = 1, "Zonder aandoening" = 0.45)

  build_sample <- function(label, color) {
    paste0(
      "<span style='color:", css_rgba(color, alpha_values[[label]]), ";'>&#9608;&#9608;</span>",
      "&nbsp;", label
    )
  }

  y_positions <- seq(-0.12, by = -0.07, length.out = length(sexes))

  lapply(seq_along(sexes), function(i) {
    sex <- sexes[[i]]
    color <- sex_palette[[sex]] %||% "#4b5563"
    row_text <- paste(
      vapply(status_labels, build_sample, character(1), color = color),
      collapse = "&nbsp;&nbsp;&nbsp;&nbsp;"
    )

    list(
      x = 0,
      y = y_positions[[i]],
      xref = "paper",
      yref = "paper",
      xanchor = "left",
      yanchor = "top",
      align = "left",
      showarrow = FALSE,
      text = paste0(
        "<span style='color:", color, "; font-weight:600;'>", sex, ":</span>",
        "&nbsp;&nbsp;",
        row_text
      ),
      font = list(size = 12, color = "#4b5563")
    )
  })
}

load_map_sf <- function(path) {
  map_df <- sf::st_read(path, quiet = TRUE)

  if ("geom" %in% names(map_df) && !"geometry" %in% names(map_df)) {
    names(map_df)[names(map_df) == "geom"] <- "geometry"
    sf::st_geometry(map_df) <- "geometry"
  }

  map_df |>
    dplyr::mutate(stadsdelen = as.character(stadsdelen))
}

build_map_long <- function(df, group_key) {
  id_cols <- intersect(c("stadsdelen", "geslacht", "n_totaal"), names(df))
  special_regex <- "(_n_totaal_gebruikers$|_ratio$|_stadsdelen_verwacht(_leeftijd)?$|_stadsdelen_verwacht(_leeftijd)?_n_totaal_gebruikers$)"
  base_cols <- setdiff(names(df), id_cols)
  outcome_cols <- base_cols[!stringr::str_detect(base_cols, special_regex)]

  dplyr::bind_rows(lapply(outcome_cols, function(outcome_key) {
    expected_col <- c(
      paste0(outcome_key, "_stadsdelen_verwacht"),
      paste0(outcome_key, "_stadsdelen_verwacht_leeftijd")
    )
    expected_col <- expected_col[expected_col %in% names(df)][1]

    ratio_col <- if (paste0(outcome_key, "_ratio") %in% names(df)) {
      paste0(outcome_key, "_ratio")
    } else {
      NA_character_
    }

    count_col <- if (paste0(outcome_key, "_n_totaal_gebruikers") %in% names(df)) {
      paste0(outcome_key, "_n_totaal_gebruikers")
    } else {
      NA_character_
    }

    common <- data.frame(
      group_key = group_key,
      group_label = if (group_key == "total") "Totaal" else "Geslacht",
      stadsdelen = as.character(df$stadsdelen),
      geslacht = if ("geslacht" %in% names(df)) as.character(df$geslacht) else NA_character_,
      n_totaal = if ("n_totaal" %in% names(df)) as.numeric(df$n_totaal) else NA_real_,
      family_key = family_key_from_metric(outcome_key),
      family_label = pretty_family(family_key_from_metric(outcome_key)),
      outcome_key = outcome_key,
      outcome_label = pretty_metric_name(outcome_key),
      count = if (!is.na(count_col)) as.numeric(df[[count_col]]) else NA_real_,
      stringsAsFactors = FALSE
    )

    rows <- list()

    observed_row <- common
    observed_row$value_kind <- "observed"
    observed_row$value_label <- "Geobserveerd"
    observed_row$value <- as.numeric(df[[outcome_key]])
    rows[[length(rows) + 1]] <- observed_row

    if (!is.na(expected_col)) {
      expected_row <- common
      expected_row$value_kind <- "expected"
      expected_row$value_label <- "Verwacht op basis van leeftijd"
      expected_row$value <- as.numeric(df[[expected_col]])
      rows[[length(rows) + 1]] <- expected_row
    }

    if (!is.na(ratio_col) || !is.na(expected_col)) {
      ratio_row <- common
      ratio_row$value_kind <- "ratio"
      ratio_row$value_label <- "Index"
      if (!is.na(ratio_col)) {
        ratio_row$value <- as.numeric(df[[ratio_col]])
      } else {
        observed_value <- as.numeric(df[[outcome_key]])
        expected_value <- as.numeric(df[[expected_col]])
        ratio_row$value <- dplyr::if_else(
          is.na(expected_value) | expected_value == 0,
          NA_real_,
          observed_value / expected_value
        )
      }
      rows[[length(rows) + 1]] <- ratio_row
    }

    dplyr::bind_rows(rows)
  })) |>
    dplyr::mutate(
      display_value = dplyr::case_when(
        value_kind != "ratio" & family_key == "heeft" ~ value * 100,
        TRUE ~ value
      ),
      short_label = sub("^[A-Za-z]\\s+", "", stadsdelen)
    )
}

expand_equal_range <- function(values) {
  values <- values[is.finite(values)]
  if (length(values) == 0) {
    return(c(0, 1))
  }
  rng <- range(values, na.rm = TRUE)
  if (diff(rng) == 0) {
    bump <- if (rng[[1]] == 0) 0.1 else abs(rng[[1]]) * 0.05
    rng <- c(rng[[1]] - bump, rng[[2]] + bump)
  }
  rng
}

map_legend_title <- function(family_key, value_kind) {
  if (value_kind == "ratio") {
    return("Index")
  }
  if (family_key == "heeft") {
    return("Gebruik")
  }
  if (family_key == "kosten") {
    return("Kosten")
  }
  "Aantal"
}

format_map_tooltip_value <- function(value, family_key, value_kind) {
  if (!is.finite(value)) {
    return("Geen waarde beschikbaar")
  }

  if (identical(value_kind, "ratio")) {
    return(scales::number(value, accuracy = 0.01, big.mark = ".", decimal.mark = ","))
  }

  if (identical(family_key, "heeft")) {
    return(paste0(scales::number(value, accuracy = 0.1, big.mark = ".", decimal.mark = ","), "%"))
  }

  scales::number(value, accuracy = 1, big.mark = ".", decimal.mark = ",")
}

map_numeric_vector <- function(x) {
  if (is.list(x)) {
    return(vapply(x, function(item) {
      item <- unlist(item, recursive = TRUE, use.names = FALSE)
      if (length(item) == 0) {
        return(NA_real_)
      }
      suppressWarnings(as.numeric(item[[1]]))
    }, numeric(1)))
  }

  suppressWarnings(as.numeric(x))
}

map_character_vector <- function(x) {
  if (is.list(x)) {
    return(vapply(x, function(item) {
      item <- unlist(item, recursive = TRUE, use.names = FALSE)
      if (length(item) == 0) {
        return(NA_character_)
      }
      as.character(item[[1]])
    }, character(1)))
  }

  as.character(x)
}

parse_map_min_count <- function(value) {
  if (is.null(value) || !nzchar(value)) {
    return(NA_real_)
  }
  suppressWarnings(as.numeric(value))
}

apply_map_masks <- function(df, exclude_westpoort = FALSE, min_count = NA_real_) {
  count_values <- map_numeric_vector(df$count)
  display_values <- map_numeric_vector(df$display_value)
  stadsdelen_values <- map_character_vector(df$stadsdelen)

  df |>
    dplyr::mutate(
      masked_value = dplyr::case_when(
        isTRUE(exclude_westpoort) & stadsdelen_values == "B Westpoort" ~ NA_real_,
        is.finite(min_count) & !is.na(count_values) & count_values < min_count ~ NA_real_,
        TRUE ~ display_values
      )
    )
}

build_map_tooltip <- function(df_map, exclude_westpoort = FALSE, min_count = NA_real_) {
  stadsdelen_values <- map_character_vector(df_map$stadsdelen)
  sex_values <- map_character_vector(df_map$geslacht)
  count_values <- map_numeric_vector(df_map$count)
  masked_values <- map_numeric_vector(df_map$masked_value)
  family_values <- map_character_vector(df_map$family_key)
  kind_values <- map_character_vector(df_map$value_kind)

  value_text <- vapply(seq_len(nrow(df_map)), function(i) {
    if (isTRUE(exclude_westpoort) && identical(stadsdelen_values[[i]], "B Westpoort")) {
      return("Uitgesloten")
    }

    if (is.finite(min_count) && !is.na(count_values[[i]]) && count_values[[i]] < min_count) {
      return(paste0("Verborgen (n < ", min_count, ")"))
    }

    format_map_tooltip_value(
      value = masked_values[[i]],
      family_key = family_values[[i]],
      value_kind = kind_values[[i]]
    )
  }, character(1))

  sex_text <- vapply(seq_along(sex_values), function(i) {
    if (is.na(sex_values[[i]]) || !nzchar(sex_values[[i]])) {
      return("")
    }
    paste0("<br>Geslacht: ", sex_values[[i]])
  }, character(1))

  count_text <- vapply(seq_along(count_values), function(i) {
    if (is.na(count_values[[i]])) {
      return("")
    }
    paste0("<br>n: ", scales::number(count_values[[i]], accuracy = 1, big.mark = ".", decimal.mark = ","))
  }, character(1))

  paste0(
    "Stadsdeel: ", stadsdelen_values,
    sex_text,
    count_text,
    "<br>Waarde: ", value_text
  )
}

map_sf <- if (!is.na(map_path)) load_map_sf(map_path) else NULL

map_long <- {
  total_sheet <- if ("huisart_stad_total" %in% sheet_names) read_sheet_cached("huisart_stad_total") else NULL
  sex_sheet <- if ("huisart_stad_sex" %in% sheet_names) read_sheet_cached("huisart_stad_sex") else NULL

  dplyr::bind_rows(
    if (!is.null(total_sheet)) build_map_long(total_sheet, "total"),
    if (!is.null(sex_sheet)) build_map_long(sex_sheet, "sex")
  )
}

ui <- navbarPage(
  title = "Ongelijkheid tussen mannen en vrouwen",
  id = "main_nav",
  header = tagList(
    tags$head(
      tags$style(HTML("
        html,
        body,
        .navbar,
        .navbar-default,
        .navbar-default .navbar-collapse,
        .navbar-default .navbar-form,
        .container-fluid,
        .tab-content,
        .navbar-header,
        .navbar-nav,
        .main-container,
        .well,
        .row,
        .col-sm-4,
        .col-sm-8,
        .col-sm-12,
        .tab-pane,
        .sidebarPanel,
        .mainPanel,
        .dataTables_wrapper,
        table.dataTable,
        table.dataTable thead th,
        table.dataTable tbody td,
        table.dataTable tbody tr,
        .panel,
        .js-plotly-plot .plot-container,
        .js-plotly-plot .svg-container,
        .shiny-plot-output,
        .plot-container {
          background: #ffffff !important;
          background-color: #ffffff !important;
        }

        .well {
          border: none !important;
          box-shadow: none !important;
        }
      "))
    ),
    tags$div(
      style = "padding: 12px 18px 4px 18px; color: #4b5563; font-size: 14px;",
      "Interactieve tool voor het verkennen van verschillen tussen mannen en vrouwen in zorggebruik, zorgkosten en aandoeningsprevalentie in Amsterdam en Nederland."
    )
  ),
  tabPanel(
    "Jaarlijkse trends",
    sidebarLayout(
      sidebarPanel(
        selectInput(
          "year_sheet",
          "Dataset",
          choices = stats::setNames(yearly_index$sheet, yearly_index$sheet_label),
          selected = yearly_index$sheet[[1]]
        ),
        selectInput("year_outcome", "Uitkomst", choices = NULL),
        selectInput("year_type", "Statistiek", choices = NULL),
        checkboxGroupInput("year_sex", "Geslacht", choices = NULL),
        checkboxGroupInput("year_region", "Gebied", choices = NULL),
        checkboxGroupInput(
          "year_include_zero",
          "Nul op y-as opnemen",
          choices = c("Ja" = "yes"),
          selected = "yes"
        ),
        uiOutput("year_condition_ui"),
        downloadButton("dl_year_plot", "Grafiek downloaden")
      ),
      mainPanel(
        plotly::plotlyOutput("plot_year", height = "520px"),
        DTOutput("tbl_year")
      )
    )
  ),
  tabPanel(
    "Event studies",
    fluidPage(
      tabsetPanel(
        tabPanel(
          "Geobserveerde gemiddelden",
          sidebarLayout(
            sidebarPanel(
              selectInput(
                "es_mean_sheet",
                "Dataset",
                choices = stats::setNames(es_mean_index$sheet, es_mean_index$sheet_label),
                selected = es_mean_index$sheet[[1]]
              ),
              selectInput("es_mean_outcome", "Uitkomst", choices = NULL),
              selectInput("es_mean_type", "Statistiek", choices = NULL),
              checkboxGroupInput("es_mean_sex", "Geslacht", choices = NULL),
              checkboxGroupInput("es_mean_region", "Gebied", choices = NULL),
              uiOutput("es_mean_diff_ui"),
              checkboxGroupInput(
                "es_mean_include_zero",
                "Nul op y-as opnemen",
                choices = c("Ja" = "yes"),
                selected = "yes"
              ),
              uiOutput("es_mean_condition_ui"),
              downloadButton("dl_es_mean_plot", "Grafiek downloaden")
            ),
            mainPanel(
              plotly::plotlyOutput("plot_es_mean", height = "520px"),
              DTOutput("tbl_es_mean")
            )
          )
        ),
        tabPanel(
          "Geschatte effecten",
          sidebarLayout(
            sidebarPanel(
              selectInput(
                "es_ci_sheet",
                "Dataset",
                choices = stats::setNames(es_ci_index$sheet, es_ci_index$sheet_label),
                selected = es_ci_index$sheet[[1]]
              ),
              selectInput("es_ci_outcome", "Uitkomst", choices = NULL),
              checkboxGroupInput("es_ci_sex", "Geslacht", choices = NULL),
              checkboxGroupInput("es_ci_sample", "Gebied", choices = NULL),
              downloadButton("dl_es_ci_plot", "Grafiek downloaden")
            ),
            mainPanel(
              plotly::plotlyOutput("plot_es_ci", height = "520px"),
              DTOutput("tbl_es_ci")
            )
          )
        )
      )
    )
  ),
  tabPanel(
    "2023",
    sidebarLayout(
      sidebarPanel(
        selectInput(
          "snapshot_base",
          "Dataset",
          choices = stats::setNames(snapshot_base_index$base_name, snapshot_base_index$base_label),
          selected = first_or_empty(snapshot_base_index$base_name)
        ),
        selectInput("snapshot_level_outcome", "Uitkomst", choices = NULL),
        selectInput("snapshot_level_type", "Statistiek", choices = NULL),
        checkboxGroupInput("snapshot_level_sex", "Geslacht", choices = NULL),
        checkboxGroupInput("snapshot_level_region", "Gebied", choices = NULL),
        downloadButton("dl_snapshot_level_plot", "Grafiek downloaden")
      ),
      mainPanel(
        plotly::plotlyOutput("plot_snapshot_level", height = "560px"),
        DTOutput("tbl_snapshot_level")
      )
    )
  ),
  tabPanel(
    "Profielen",
    sidebarLayout(
      sidebarPanel(
        selectInput(
          "profile_base",
          "Dataset",
          choices = stats::setNames(profile_base_index$base_id, profile_base_index$base_label),
          selected = first_or_empty(profile_base_index$base_id)
        ),
        uiOutput("profile_period_ui"),
        selectInput("profile_category", "Categorie", choices = NULL),
        checkboxGroupInput("profile_sex", "Geslacht", choices = NULL),
        checkboxGroupInput("profile_region", "Gebied", choices = NULL),
        downloadButton("dl_profile_plot", "Grafiek downloaden")
      ),
      mainPanel(
        plotly::plotlyOutput("plot_profile", height = "560px"),
        DTOutput("tbl_profile")
      )
    )
  ),
  tabPanel(
    "Kaarten",
    sidebarLayout(
      sidebarPanel(
        selectInput(
          "map_group",
          "Kaartniveau",
          choices = stats::setNames(c("total", "sex"), c("Totaal", "Geslacht")),
          selected = if ("sex" %in% unique(map_long$group_key)) "sex" else "total"
        ),
        uiOutput("map_sex_ui"),
        selectInput("map_value_kind", "Kaarttype", choices = NULL),
        selectInput("map_family", "Uitkomstgroep", choices = NULL),
        selectInput("map_outcome", "Uitkomst", choices = NULL),
        conditionalPanel(
          condition = "input.map_value_kind == 'observed'",
          color_picker_input(
            "map_observed_high_color",
            "Selecteer kleur (hoogst)",
            "#8B0A50"
          )
        ),
        conditionalPanel(
          condition = "input.map_value_kind == 'ratio'",
          color_picker_input(
            "map_ratio_low_color",
            "Selecteer kleur (laagst)",
            "#483D8B"
          ),
          color_picker_input(
            "map_ratio_high_color",
            "Selecteer kleur (hoogst)",
            "#FF0000"
          )
        ),
        selectInput(
          "map_min_count",
          "Stadsdelen verbergen bij n kleiner dan",
          choices = c("10" = "10", "30" = "30", "50" = "50"),
          selected = "10"
        ),
        checkboxInput("map_exclude_westpoort", "Westpoort uitsluiten", value = FALSE),
        checkboxInput("map_fixed_legend", "Legenda voor gekozen uitkomst vastzetten", value = TRUE),
        checkboxInput("map_show_labels", "Labels van stadsdelen tonen", value = TRUE),
        conditionalPanel(
          condition = "input.map_value_kind == 'ratio' || input.map_value_kind == 'observed'",
          checkboxInput("map_manual_scale", "Kaartschaal handmatig instellen", value = FALSE),
          conditionalPanel(
            condition = "input.map_manual_scale",
            conditionalPanel(
              condition = "input.map_value_kind == 'ratio'",
              helpText("Midden blijft altijd 1 (wit).")
            ),
            conditionalPanel(
              condition = "input.map_value_kind == 'observed'",
              helpText("Min is wit en max is de gekozen kleur.")
            ),
            fluidRow(
              column(6, numericInput("map_scale_min", "Min", value = 0.9, step = 0.01)),
              column(6, numericInput("map_scale_max", "Max", value = 1.1, step = 0.01))
            )
          )
        ),
        downloadButton("dl_map", "Kaart downloaden")
      ),
      mainPanel(
        tags$div(
          style = "position: relative;",
          plotOutput(
            "plot_map",
            height = "700px",
            hover = hoverOpts("plot_map_hover", delay = 100, delayType = "debounce")
          ),
          uiOutput("plot_map_hover_ui")
        ),
        DTOutput("tbl_map")
      )
    )
  )
)

server <- function(input, output, session) {
  shared_outcome_selection <- reactiveVal(NULL)

  current_snapshot_sheet <- reactive({
    req(input$snapshot_base)
    candidates <- snapshot_level_index |>
      dplyr::filter(base_name == input$snapshot_base)
    req(nrow(candidates) > 0)

    selected_sheet <- candidates$sheet[[1]]
    req(length(selected_sheet) > 0, !is.na(selected_sheet))
    selected_sheet
  })

  current_profile_sheet <- reactive({
    req(input$profile_base)
    candidates <- profile_index |>
      dplyr::filter(base_id == input$profile_base)
    req(nrow(candidates) > 0)

    selected_period <- choose_preferred_choice(
      choices = candidates$period_key,
      current = input$profile_period,
      fallback = "match"
    )
    selected_sheet <- candidates$sheet[candidates$period_key == selected_period][1]
    req(length(selected_sheet) > 0, !is.na(selected_sheet))
    selected_sheet
  })

  available_outcomes_for_input <- function(kind) {
    if (identical(kind, "year")) {
      req(input$year_sheet)
      return(sort(unique(as.character(read_sheet_cached(input$year_sheet)$name))))
    }
    if (identical(kind, "es_mean")) {
      req(input$es_mean_sheet)
      return(sort(unique(as.character(read_sheet_cached(input$es_mean_sheet)$name))))
    }
    if (identical(kind, "es_ci")) {
      req(input$es_ci_sheet)
      return(sort(unique(as.character(read_sheet_cached(input$es_ci_sheet)$outcome))))
    }
    if (identical(kind, "snapshot_level")) {
      return(sort(unique(as.character(read_sheet_cached(current_snapshot_sheet())$name))))
    }
    character(0)
  }

  sync_outcome_across_tabs <- function(source_kind, outcome_value) {
    if (is.null(outcome_value) || !nzchar(outcome_value)) {
      return()
    }

    shared_outcome_selection(outcome_value)

    target_specs <- list(
      year = "year_outcome",
      es_mean = "es_mean_outcome",
      es_ci = "es_ci_outcome",
      snapshot_level = "snapshot_level_outcome"
    )

    for (kind in setdiff(names(target_specs), source_kind)) {
      choices <- tryCatch(available_outcomes_for_input(kind), error = function(e) character(0))
      if (!(outcome_value %in% choices)) {
        next
      }

      input_id <- target_specs[[kind]]
      if (identical(input[[input_id]] %||% "", outcome_value)) {
        next
      }

      updateSelectInput(session, input_id, selected = outcome_value)
    }
  }

  save_png_export <- function(file, plot_obj) {
    ggplot2::ggsave(
      file,
      plot = plot_obj,
      scale = 0.7,
      width = 14,
      height = 10,
      dpi = 300,
      bg = "transparent"
    )
  }

  save_plot_png <- function(file, plot_obj) {
    save_png_export(file, plot_obj)
  }

  map_color_values <- shiny::debounce(reactive({
    list(
      observed_high = normalize_hex_color(input$map_observed_high_color, "#8B0A50"),
      ratio_low = normalize_hex_color(input$map_ratio_low_color, "#483D8B"),
      ratio_high = normalize_hex_color(input$map_ratio_high_color, "#FF0000")
    )
  }), millis = 100)

  output$tbl_scripts <- renderDT({
    display_table(script_index, export_name = "code_overzicht")
  })

  output$tbl_sheets <- renderDT({
    display_table(
      data.frame(sheet = sheet_names, stringsAsFactors = FALSE),
      export_name = "werkbladen_output_xlsx"
    )
  })

  output$profile_period_ui <- renderUI({
    req(input$profile_base)
    candidates <- profile_index |>
      dplyr::filter(base_id == input$profile_base) |>
      dplyr::distinct(period_key, period_label)
    if (nrow(candidates) == 0) {
      return(NULL)
    }
    candidates <- candidates |>
      dplyr::mutate(period_order = dplyr::case_when(
        period_key == "match" ~ 1,
        period_key == "23" ~ 2,
        TRUE ~ 99
      )) |>
      dplyr::arrange(period_order, period_label)

    radioButtons(
      "profile_period",
      "Periode",
      choices = stats::setNames(candidates$period_key, candidates$period_label),
      selected = choose_preferred_choice(
        candidates$period_key,
        current = input$profile_period,
        fallback = "match"
      )
    )
  })

  observeEvent(input$year_sheet, {
    df <- read_sheet_cached(input$year_sheet)
    outcomes <- sort(unique(as.character(df$name)))
    if (length(outcomes) == 0) {
      updateSelectInput(session, "year_outcome", choices = character(0), selected = character(0))
      updateSelectInput(session, "year_type", choices = character(0), selected = character(0))
      return()
    }
    updateSelectInput(
      session, "year_outcome",
      choices = stats::setNames(outcomes, pretty_metric_name(outcomes)),
      selected = choose_preferred_choice(
        outcomes,
        current = input$year_outcome,
        fallback = shared_outcome_selection()
      )
    )
    updateCheckboxGroupInput(
      session, "year_sex",
      choices = ordered_sex_levels(df$geslacht),
      selected = ordered_sex_levels(df$geslacht)
    )
    updateCheckboxGroupInput(
      session, "year_region",
      choices = ordered_area_levels(df$amsterdam),
      selected = ordered_area_levels(df$amsterdam)
    )
  }, ignoreNULL = FALSE)

  observeEvent(input$year_outcome, {
    sync_outcome_across_tabs("year", input$year_outcome)
  }, ignoreInit = TRUE)

  observeEvent(list(input$year_sheet, input$year_outcome), {
    df <- read_sheet_cached(input$year_sheet)
    outcomes <- sort(unique(as.character(df$name)))
    if (length(outcomes) == 0) {
      updateSelectInput(session, "year_type", choices = character(0), selected = character(0))
      return()
    }
    active_outcome <- input$year_outcome
    if (is.null(active_outcome) || !nzchar(active_outcome) || !(active_outcome %in% outcomes)) {
      active_outcome <- outcomes[[1]]
    }
    types <- df |>
      dplyr::filter(name == active_outcome) |>
      dplyr::pull(type) |>
      as.character() |>
      unique() |>
      sort()
    if (length(types) == 0) {
      updateSelectInput(session, "year_type", choices = character(0), selected = character(0))
      return()
    }
    updateSelectInput(
      session, "year_type",
      choices = stats::setNames(types, pretty_type(types)),
      selected = choose_preferred_choice(
        types,
        current = input$year_type
      )
    )
  }, ignoreNULL = FALSE)

  output$year_condition_ui <- renderUI({
    df <- read_sheet_cached(input$year_sheet)
    flag_col <- condition_flag_col(df)
    if (is.na(flag_col)) {
      return(NULL)
    }
    values <- sort(unique(as.character(df[[flag_col]])))
    checkboxGroupInput(
      "year_condition_filter",
      "Diagnosestatus",
      choices = stats::setNames(values, pretty_binary(values)),
      selected = values
    )
  })

  year_filtered <- reactive({
    req(input$year_sheet, input$year_outcome, input$year_type)
    df <- read_sheet_cached(input$year_sheet)
    flag_col <- condition_flag_col(df)
    selected_sex <- input$year_sex %||% unique(as.character(df$geslacht))
    selected_region <- input$year_region %||% unique(as.character(df$amsterdam))

    df <- df |>
      dplyr::filter(
        name == input$year_outcome,
        type == input$year_type,
        geslacht %in% selected_sex,
        amsterdam %in% selected_region
      )

    if (!is.na(flag_col)) {
      selected_status <- input$year_condition_filter %||% unique(as.character(df[[flag_col]]))
      df <- df |>
        dplyr::filter(.data[[flag_col]] %in% selected_status)
    }

    df <- df |>
      dplyr::mutate(amsterdam = factor(as.character(amsterdam), levels = ordered_area_levels(selected_region)))

    sheet_meta <- yearly_index[yearly_index$sheet == input$year_sheet, , drop = FALSE]
    df$series <- build_series_label(df, flag_col, sheet_meta$condition_label[[1]] %||% NA_character_, include_area = FALSE)
    df <- prepare_time_series_aesthetics(df, flag_col, sheet_meta$condition_label[[1]] %||% NA_character_)
    df
  })

  year_plot_obj <- reactive({
    df <- year_filtered()
    validate(need(nrow(df) > 0, "Geen gegevens beschikbaar voor de huidige selectie."))
    sheet_meta <- yearly_index[yearly_index$sheet == input$year_sheet, , drop = FALSE]
    build_time_series_plot(
      df = df,
      x_var = "year",
      title = compose_title_parts(sheet_meta$dataset_label[[1]], sheet_meta$condition_label[[1]]),
      subtitle = compose_title_parts(pretty_metric_name(input$year_outcome), pretty_type(input$year_type)),
      y_labeler = metric_labeler(input$year_outcome, input$year_type),
      axis_labeler = axis_metric_labeler(input$year_outcome, input$year_type),
      facet_var = "amsterdam",
      compact_facets = TRUE,
      show_confidence = TRUE,
      include_zero = "yes" %in% (input$year_include_zero %||% character(0)),
      y_max = metric_axis_upper_limit(input$year_outcome, input$year_type)
    )
  })

  output$plot_year <- plotly::renderPlotly({
    df <- year_filtered()
    custom_annotations <- build_event_study_annotations(df)
    if (length(custom_annotations) == 0) {
      custom_annotations <- build_sex_annotations(df)
    }
    bottom_margin <- if (length(custom_annotations) > 1) 115 else 90
    y_max <- metric_axis_upper_limit(input$year_outcome, input$year_type)
    y_values <- if ("se" %in% names(df) && any(is.finite(df$se))) {
      c(df$value - 1.96 * df$se, df$value + 1.96 * df$se)
    } else {
      df$value
    }

    plot_obj <- plotly::ggplotly(year_plot_obj(), tooltip = "text")

    if (!is.null(y_max)) {
      y_lower <- compute_axis_lower_limit(
        y_values,
        include_zero = "yes" %in% (input$year_include_zero %||% character(0)),
        upper_limit = y_max
      )
      y_upper <- compute_axis_upper_bound(y_values, lower_limit = y_lower, upper_limit = y_max)
      plot_obj <- apply_plotly_y_range(
        plot_obj,
        lower = y_lower,
        upper = y_upper
      )
    }

    apply_plotly_transparent_layout(
      plot_obj,
      showlegend = FALSE,
      annotations = custom_annotations,
      margin = list(b = bottom_margin)
    )
  })

  output$dl_year_plot <- downloadHandler(
    filename = function() {
      paste0(
        build_export_name(
          "grafiek_jaarlijkse_trends",
          input$year_sheet %||% "dataset",
          input$year_outcome %||% "uitkomst",
          input$year_type %||% "statistiek"
        ),
        ".png"
      )
    },
    content = function(file) {
      save_plot_png(file, year_plot_obj())
    }
  )

  output$tbl_year <- renderDT({
    df <- year_filtered()
    display_table(
      df,
      export_name = build_export_name(
        "jaarlijkse_trends",
        input$year_sheet %||% "dataset",
        input$year_outcome %||% "uitkomst",
        input$year_type %||% "statistiek"
      )
    )
  })

  observeEvent(input$es_mean_sheet, {
    df <- read_sheet_cached(input$es_mean_sheet)
    outcomes <- sort(unique(as.character(df$name)))
    if (length(outcomes) == 0) {
      updateSelectInput(session, "es_mean_outcome", choices = character(0), selected = character(0))
      updateSelectInput(session, "es_mean_type", choices = character(0), selected = character(0))
      return()
    }
    updateSelectInput(
      session, "es_mean_outcome",
      choices = stats::setNames(outcomes, pretty_metric_name(outcomes)),
      selected = choose_preferred_choice(
        outcomes,
        current = input$es_mean_outcome,
        fallback = shared_outcome_selection()
      )
    )
    updateCheckboxGroupInput(
      session, "es_mean_sex",
      choices = ordered_sex_levels(df$geslacht),
      selected = ordered_sex_levels(df$geslacht)
    )
    updateCheckboxGroupInput(
      session, "es_mean_region",
      choices = ordered_area_levels(df$amsterdam),
      selected = ordered_area_levels(df$amsterdam)
    )
  }, ignoreNULL = FALSE)

  observeEvent(input$es_mean_outcome, {
    sync_outcome_across_tabs("es_mean", input$es_mean_outcome)
  }, ignoreInit = TRUE)

  observeEvent(list(input$es_mean_sheet, input$es_mean_outcome), {
    df <- read_sheet_cached(input$es_mean_sheet)
    outcomes <- sort(unique(as.character(df$name)))
    if (length(outcomes) == 0) {
      updateSelectInput(session, "es_mean_type", choices = character(0), selected = character(0))
      return()
    }
    active_outcome <- input$es_mean_outcome
    if (is.null(active_outcome) || !nzchar(active_outcome) || !(active_outcome %in% outcomes)) {
      active_outcome <- outcomes[[1]]
    }
    types <- df |>
      dplyr::filter(name == active_outcome) |>
      dplyr::pull(type) |>
      as.character() |>
      unique() |>
      sort()
    if (length(types) == 0) {
      updateSelectInput(session, "es_mean_type", choices = character(0), selected = character(0))
      return()
    }
    updateSelectInput(
      session, "es_mean_type",
      choices = stats::setNames(types, pretty_type(types)),
      selected = choose_preferred_choice(
        types,
        current = input$es_mean_type
      )
    )
  }, ignoreNULL = FALSE)

  output$es_mean_diff_ui <- renderUI({
    df <- read_sheet_cached(input$es_mean_sheet)
    flag_col <- condition_flag_col(df)
    if (is.na(flag_col) || dplyr::n_distinct(df[[flag_col]]) <= 1) {
      return(NULL)
    }
    checkboxGroupInput(
      "es_mean_show_diff",
      "Verschil tonen",
      choices = c("Ja" = "yes"),
      selected = if ("yes" %in% (input$es_mean_show_diff %||% character(0))) "yes" else character(0)
    )
  })

  es_mean_view_is_diff <- reactive({
    df <- read_sheet_cached(input$es_mean_sheet)
    flag_col <- condition_flag_col(df)
    !is.na(flag_col) &&
      dplyr::n_distinct(df[[flag_col]]) > 1 &&
      "yes" %in% (input$es_mean_show_diff %||% character(0))
  })

  output$es_mean_condition_ui <- renderUI({
    if (isTRUE(es_mean_view_is_diff())) {
      return(NULL)
    }
    df <- read_sheet_cached(input$es_mean_sheet)
    flag_col <- condition_flag_col(df)
    if (is.na(flag_col)) {
      return(NULL)
    }
    values <- sort(unique(as.character(df[[flag_col]])))
    checkboxGroupInput(
      "es_mean_condition_filter",
      "Diagnosestatus",
      choices = stats::setNames(values, pretty_binary(values)),
      selected = values
    )
  })

  es_mean_filtered <- reactive({
    req(input$es_mean_sheet, input$es_mean_outcome, input$es_mean_type)
    df <- read_sheet_cached(input$es_mean_sheet)
    flag_col <- condition_flag_col(df)
    selected_sex <- input$es_mean_sex %||% unique(as.character(df$geslacht))
    selected_region <- input$es_mean_region %||% unique(as.character(df$amsterdam))

    df <- df |>
      dplyr::filter(
        name == input$es_mean_outcome,
        type == input$es_mean_type,
        geslacht %in% selected_sex,
        amsterdam %in% selected_region
      )

    if (isTRUE(es_mean_view_is_diff())) {
      df <- compute_condition_difference(df, flag_col)
      flag_col <- NA_character_
    } else if (!is.na(flag_col)) {
      selected_status <- input$es_mean_condition_filter %||% unique(as.character(df[[flag_col]]))
      df <- df |>
        dplyr::filter(.data[[flag_col]] %in% selected_status)
    }

    df <- df |>
      dplyr::mutate(amsterdam = factor(as.character(amsterdam), levels = ordered_area_levels(selected_region)))

    sheet_meta <- es_mean_index[es_mean_index$sheet == input$es_mean_sheet, , drop = FALSE]
    df$series <- build_series_label(df, flag_col, sheet_meta$condition_label[[1]] %||% NA_character_, include_area = FALSE)
    df <- prepare_time_series_aesthetics(df, flag_col, sheet_meta$condition_label[[1]] %||% NA_character_)
    df
  })

  es_mean_plot_obj <- reactive({
    df <- es_mean_filtered()
    validate(need(nrow(df) > 0, "Geen gegevens beschikbaar voor de huidige selectie."))
    sheet_meta <- es_mean_index[es_mean_index$sheet == input$es_mean_sheet, , drop = FALSE]
    build_time_series_plot(
      df = df,
      x_var = "years_since_diagnosis",
      title = compose_title_parts(
        sheet_meta$dataset_label[[1]],
        sheet_meta$condition_label[[1]]
      ),
      subtitle = compose_title_parts(
        pretty_metric_name(input$es_mean_outcome),
        pretty_type(input$es_mean_type),
        if (isTRUE(es_mean_view_is_diff())) "Verschil (met - zonder)"
      ),
      y_labeler = metric_labeler(input$es_mean_outcome, input$es_mean_type),
      axis_labeler = axis_metric_labeler(input$es_mean_outcome, input$es_mean_type),
      vline = 0,
      facet_var = "amsterdam",
      compact_facets = TRUE,
      include_zero = "yes" %in% (input$es_mean_include_zero %||% character(0)),
      y_max = metric_axis_upper_limit(input$es_mean_outcome, input$es_mean_type)
    )
  })

  output$plot_es_mean <- plotly::renderPlotly({
    df <- es_mean_filtered()
    custom_annotations <- build_event_study_annotations(df)
    if (length(custom_annotations) == 0) {
      custom_annotations <- build_sex_annotations(df)
    }
    bottom_margin <- if (length(custom_annotations) > 1) 115 else 90
    y_max <- metric_axis_upper_limit(input$es_mean_outcome, input$es_mean_type)

    plot_obj <- plotly::ggplotly(es_mean_plot_obj(), tooltip = "text")

    if (!is.null(y_max)) {
      y_lower <- compute_axis_lower_limit(
        df$value,
        include_zero = "yes" %in% (input$es_mean_include_zero %||% character(0)),
        upper_limit = y_max
      )
      y_upper <- compute_axis_upper_bound(df$value, lower_limit = y_lower, upper_limit = y_max)
      plot_obj <- apply_plotly_y_range(
        plot_obj,
        lower = y_lower,
        upper = y_upper
      )
    }

    apply_plotly_transparent_layout(
      plot_obj,
      showlegend = FALSE,
      annotations = custom_annotations,
      margin = list(b = bottom_margin),
      legend = list(
        orientation = "h",
        x = 0,
        xanchor = "left",
        y = -0.2
      )
    )
  })

  output$dl_es_mean_plot <- downloadHandler(
    filename = function() {
      paste0(
        build_export_name(
          "grafiek_event_study_geobserveerde_gemiddelden",
          input$es_mean_sheet %||% "dataset",
          input$es_mean_outcome %||% "uitkomst",
          input$es_mean_type %||% "statistiek",
          if (isTRUE(es_mean_view_is_diff())) "verschil" else "niveaus"
        ),
        ".png"
      )
    },
    content = function(file) {
      save_plot_png(file, es_mean_plot_obj())
    }
  )

  output$tbl_es_mean <- renderDT({
    display_table(
      es_mean_filtered(),
      export_name = build_export_name(
        "event_study_geobserveerde_gemiddelden",
        input$es_mean_sheet %||% "dataset",
        input$es_mean_outcome %||% "uitkomst",
        input$es_mean_type %||% "statistiek",
        if (isTRUE(es_mean_view_is_diff())) "verschil" else "niveaus"
      )
    )
  })

  observeEvent(input$snapshot_base, {
    df <- read_sheet_cached(current_snapshot_sheet())
    outcomes <- sort(unique(as.character(df$name)))
    if (length(outcomes) == 0) {
      updateSelectInput(session, "snapshot_level_outcome", choices = character(0), selected = character(0))
      updateSelectInput(session, "snapshot_level_type", choices = character(0), selected = character(0))
      return()
    }
    updateSelectInput(
      session, "snapshot_level_outcome",
      choices = stats::setNames(outcomes, pretty_metric_name(outcomes)),
      selected = choose_preferred_choice(
        outcomes,
        current = input$snapshot_level_outcome,
        fallback = shared_outcome_selection()
      )
    )
    updateCheckboxGroupInput(
      session, "snapshot_level_sex",
      choices = ordered_sex_levels(df$geslacht),
      selected = ordered_sex_levels(df$geslacht)
    )
    updateCheckboxGroupInput(
      session, "snapshot_level_region",
      choices = ordered_area_levels(df$amsterdam),
      selected = ordered_area_levels(df$amsterdam)
    )
  }, ignoreNULL = FALSE)

  observeEvent(input$snapshot_level_outcome, {
    sync_outcome_across_tabs("snapshot_level", input$snapshot_level_outcome)
  }, ignoreInit = TRUE)

  observeEvent(list(input$snapshot_base, input$snapshot_level_outcome), {
    df <- read_sheet_cached(current_snapshot_sheet())
    outcomes <- sort(unique(as.character(df$name)))
    if (length(outcomes) == 0) {
      updateSelectInput(session, "snapshot_level_type", choices = character(0), selected = character(0))
      return()
    }
    active_outcome <- input$snapshot_level_outcome
    if (is.null(active_outcome) || !nzchar(active_outcome) || !(active_outcome %in% outcomes)) {
      active_outcome <- outcomes[[1]]
    }
    types <- df |>
      dplyr::filter(name == active_outcome) |>
      dplyr::pull(type) |>
      as.character() |>
      unique() |>
      sort()
    if (length(types) == 0) {
      updateSelectInput(session, "snapshot_level_type", choices = character(0), selected = character(0))
      return()
    }
    updateSelectInput(
      session, "snapshot_level_type",
      choices = stats::setNames(types, pretty_type(types)),
      selected = choose_preferred_choice(
        types,
        current = input$snapshot_level_type
      )
    )
  }, ignoreNULL = FALSE)

  snapshot_level_filtered <- reactive({
    req(current_snapshot_sheet(), input$snapshot_level_outcome, input$snapshot_level_type)
    df <- read_sheet_cached(current_snapshot_sheet())
    selected_sex <- input$snapshot_level_sex %||% unique(as.character(df$geslacht))
    selected_region <- input$snapshot_level_region %||% unique(as.character(df$amsterdam))

    df <- df |>
      dplyr::filter(
        name == input$snapshot_level_outcome,
        type == input$snapshot_level_type,
        geslacht %in% selected_sex,
        amsterdam %in% selected_region
      ) |>
      dplyr::mutate(amsterdam = factor(as.character(amsterdam), levels = ordered_area_levels(selected_region)))

    df
  })

  snapshot_level_diff <- reactive({
    df <- snapshot_level_filtered()
    flag_col <- condition_flag_col(df)
    compute_condition_difference(df, flag_col) |>
      dplyr::mutate(amsterdam = factor(as.character(amsterdam), levels = levels(df$amsterdam)))
  })

  snapshot_level_table_data <- reactive({
    level_df <- snapshot_level_filtered() |>
      dplyr::mutate(weergave = "Niveaus")
    diff_df <- snapshot_level_diff()

    if (nrow(diff_df) == 0) {
      return(level_df)
    }

    dplyr::bind_rows(
      level_df,
      diff_df |>
        dplyr::mutate(weergave = "Verschil (met - zonder)")
    )
  })

  snapshot_level_plot_obj <- reactive({
    df <- snapshot_level_filtered()
    validate(need(nrow(df) > 0, "Geen gegevens beschikbaar voor de huidige selectie."))
    sheet_meta <- snapshot_level_index[snapshot_level_index$sheet == current_snapshot_sheet(), , drop = FALSE]
    build_snapshot_plot(
      level_df = df,
      title = compose_title_parts(
        sheet_meta$dataset_label[[1]],
        sheet_meta$condition_label[[1]],
        "2023"
      ),
      subtitle = compose_title_parts(
        pretty_metric_name(input$snapshot_level_outcome),
        pretty_type(input$snapshot_level_type),
        if (nrow(snapshot_level_diff()) > 0) "inclusief verschil (met - zonder)"
      ),
      y_labeler = metric_labeler(input$snapshot_level_outcome, input$snapshot_level_type),
      axis_labeler = axis_metric_labeler(input$snapshot_level_outcome, input$snapshot_level_type),
      diff_labeler = difference_metric_labeler(input$snapshot_level_outcome, input$snapshot_level_type),
      diff_df = snapshot_level_diff(),
      y_max = metric_axis_upper_limit(input$snapshot_level_outcome, input$snapshot_level_type)
    )
  })

  output$plot_snapshot_level <- plotly::renderPlotly({
    df <- snapshot_level_filtered()
    y_max <- metric_axis_upper_limit(input$snapshot_level_outcome, input$snapshot_level_type)
    plot_obj <- plotly::ggplotly(snapshot_level_plot_obj(), tooltip = "text")
    custom_annotations <- build_snapshot_condition_annotations(df)
    if (length(custom_annotations) == 0) {
      custom_annotations <- build_sex_annotations(df)
    }
    bottom_margin <- if (length(custom_annotations) > 1) 130 else if (length(custom_annotations) == 1) 100 else 80

    if (!is.null(y_max)) {
      y_lower <- compute_axis_lower_limit(
        df$value,
        include_zero = TRUE,
        upper_limit = y_max
      )
      y_upper <- compute_axis_upper_bound(
        df$value,
        lower_limit = y_lower,
        upper_limit = y_max
      )
      plot_obj <- apply_plotly_y_range(
        plot_obj,
        lower = y_lower,
        upper = y_upper
      )
    }

    apply_plotly_transparent_layout(
      plot_obj,
      showlegend = FALSE,
      annotations = custom_annotations,
      margin = list(b = bottom_margin),
      legend = list(
        orientation = "h",
        x = 0,
        xanchor = "left",
        y = -0.2
      )
    )
  })

  output$dl_snapshot_level_plot <- downloadHandler(
    filename = function() {
      paste0(
        build_export_name(
          "grafiek_2023_overzicht",
          input$snapshot_base %||% "dataset",
          input$snapshot_level_outcome %||% "uitkomst",
          input$snapshot_level_type %||% "statistiek"
        ),
        ".png"
      )
    },
    content = function(file) {
      save_plot_png(file, snapshot_level_plot_obj())
    }
  )

  output$tbl_snapshot_level <- renderDT({
    display_table(
      snapshot_level_table_data(),
      export_name = build_export_name(
        "gegevens_2023_overzicht",
        input$snapshot_base %||% "dataset",
        input$snapshot_level_outcome %||% "uitkomst",
        input$snapshot_level_type %||% "statistiek"
      )
    )
  })

  observeEvent(input$es_ci_sheet, {
    df <- read_sheet_cached(input$es_ci_sheet)
    outcomes <- sort(unique(as.character(df$outcome)))
    if (length(outcomes) == 0) {
      updateSelectInput(session, "es_ci_outcome", choices = character(0), selected = character(0))
      return()
    }
    updateSelectInput(
      session, "es_ci_outcome",
      choices = stats::setNames(outcomes, pretty_metric_name(outcomes)),
      selected = choose_preferred_choice(
        outcomes,
        current = input$es_ci_outcome,
        fallback = shared_outcome_selection()
      )
    )
    updateCheckboxGroupInput(
      session, "es_ci_sex",
      choices = ordered_sex_levels(df$geslacht),
      selected = ordered_sex_levels(df$geslacht)
    )
    updateCheckboxGroupInput(
      session, "es_ci_sample",
      choices = ordered_area_levels(df$sample),
      selected = ordered_area_levels(df$sample)
    )
  }, ignoreNULL = FALSE)

  observeEvent(input$es_ci_outcome, {
    sync_outcome_across_tabs("es_ci", input$es_ci_outcome)
  }, ignoreInit = TRUE)

  es_ci_filtered <- reactive({
    req(input$es_ci_sheet, input$es_ci_outcome)
    df <- read_sheet_cached(input$es_ci_sheet)
    selected_sex <- input$es_ci_sex %||% unique(as.character(df$geslacht))
    selected_sample <- input$es_ci_sample %||% unique(as.character(df$sample))

    df |>
      dplyr::filter(
        outcome == input$es_ci_outcome,
        geslacht %in% selected_sex,
        sample %in% selected_sample
      )
  })

  es_ci_plot_obj <- reactive({
    df <- es_ci_filtered()
    validate(need(nrow(df) > 0, "Geen gegevens beschikbaar voor de huidige selectie."))
    sheet_meta <- es_ci_index[es_ci_index$sheet == input$es_ci_sheet, , drop = FALSE]
    palette_values <- palette_for_values(ordered_sex_levels(df$geslacht))
    df$tooltip <- paste0(
      "Geslacht: ", df$geslacht, "<br>",
      "Gebied: ", df$sample, "<br>",
      "Jaren sinds diagnose: ", df$t, "<br>",
      "Schatting: ", scales::number(df$coef, big.mark = ".", decimal.mark = ",", accuracy = 0.001), "<br>",
      "95%-BI: [",
      scales::number(df$lo, big.mark = ".", decimal.mark = ",", accuracy = 0.001),
      ", ",
      scales::number(df$hi, big.mark = ".", decimal.mark = ",", accuracy = 0.001),
      "]"
    )

    p <- ggplot(df, aes(x = t, y = coef, color = geslacht, group = geslacht, text = tooltip)) +
      geom_hline(yintercept = 0, color = "#4b5563", linetype = "dashed", linewidth = 0.7) +
      geom_vline(xintercept = 0, color = "#b0145b", linetype = "dashed", linewidth = 0.7) +
      geom_errorbar(aes(ymin = lo, ymax = hi), width = 0.15, linewidth = 0.8) +
      geom_line(linewidth = 1) +
      geom_point(size = 2.5) +
      scale_color_manual(values = palette_values) +
      scale_x_continuous(breaks = sort(unique(df$t))) +
      scale_y_continuous(labels = scales::label_number(big.mark = ".", decimal.mark = ",")) +
      labs(
        title = compose_title_parts(sheet_meta$dataset_label[[1]], sheet_meta$condition_label[[1]]),
        subtitle = pretty_metric_name(input$es_ci_outcome),
        color = NULL
      ) +
      theme_dashboard() +
      theme(
        panel.spacing.x = grid::unit(0.05, "lines"),
        panel.spacing.y = grid::unit(0.15, "lines")
      )

    if (dplyr::n_distinct(df$sample) > 1) {
      p <- p + facet_wrap(~sample)
    }

    p
  })

  output$plot_es_ci <- plotly::renderPlotly({
    plotly::ggplotly(es_ci_plot_obj(), tooltip = "text") |>
      apply_plotly_transparent_layout(
        legend = list(
          orientation = "h",
          x = 0,
          xanchor = "left",
          y = -0.2
        )
      )
  })

  output$dl_es_ci_plot <- downloadHandler(
    filename = function() {
      paste0(
        build_export_name(
          "grafiek_event_study_geschatte_effecten",
          input$es_ci_sheet %||% "dataset",
          input$es_ci_outcome %||% "uitkomst"
        ),
        ".png"
      )
    },
    content = function(file) {
      save_plot_png(file, es_ci_plot_obj())
    }
  )

  output$tbl_es_ci <- renderDT({
    display_table(
      es_ci_filtered(),
      export_name = build_export_name(
        "event_study_geschatte_effecten",
        input$es_ci_sheet %||% "dataset",
        input$es_ci_outcome %||% "uitkomst"
      )
    )
  })

  observeEvent(list(input$profile_base, input$profile_period), {
    df <- read_sheet_cached(current_profile_sheet())
    categories <- sort(unique(as.character(df$category)))
    updateSelectInput(
      session, "profile_category",
      choices = stats::setNames(categories, pretty_default(categories)),
      selected = categories[[1]]
    )
    updateCheckboxGroupInput(
      session, "profile_sex",
      choices = ordered_sex_levels(df$geslacht),
      selected = ordered_sex_levels(df$geslacht)
    )
    updateCheckboxGroupInput(
      session, "profile_region",
      choices = ordered_area_levels(df$amsterdam),
      selected = ordered_area_levels(df$amsterdam)
    )
  }, ignoreNULL = FALSE)

  profile_filtered <- reactive({
    req(current_profile_sheet(), input$profile_category)
    df <- read_sheet_cached(current_profile_sheet())
    selected_sex <- input$profile_sex %||% unique(as.character(df$geslacht))
    selected_region <- input$profile_region %||% unique(as.character(df$amsterdam))
    profile_meta <- profile_index[profile_index$sheet == current_profile_sheet(), , drop = FALSE]

    df |>
      dplyr::filter(
        category == input$profile_category,
        geslacht %in% selected_sex,
        amsterdam %in% selected_region
      ) |>
      dplyr::mutate(
        condition_label = pretty_profile_condition(condition, profile_meta$condition_label[[1]]),
        category_label = pretty_default(category)
      )
  })

  profile_plot_obj <- reactive({
    df <- profile_filtered()
    validate(need(nrow(df) > 0, "Geen gegevens beschikbaar voor de huidige selectie."))
    palette_values <- profile_palette_for_labels(unique(df$condition_label))
    title_text <- compose_title_parts(
      profile_index$condition_label[profile_index$sheet == current_profile_sheet()],
      profile_index$stage_label[profile_index$sheet == current_profile_sheet()],
      profile_index$period_label[profile_index$sheet == current_profile_sheet()]
    )

    if (identical(input$profile_category, "leeftijd")) {
      df$tooltip <- paste0(
        "Conditie: ", df$condition_label, "<br>",
        "Geslacht: ", df$geslacht, "<br>",
        "Gebied: ", df$amsterdam, "<br>",
        "Waarde: ", scales::number(df$share, big.mark = ".", decimal.mark = ",", accuracy = 0.1)
      )

      ggplot(df, aes(x = condition_label, y = share, fill = condition_label, text = tooltip)) +
        geom_col(width = 0.65) +
        facet_grid(geslacht ~ amsterdam) +
        scale_fill_manual(values = palette_values) +
        scale_y_continuous(labels = scales::label_number(big.mark = ".", decimal.mark = ",")) +
        labs(
          title = title_text,
          subtitle = "Gemiddelde leeftijd",
          fill = NULL
        ) +
        theme_dashboard()
    } else {
      df$tooltip <- paste0(
        "Categorie: ", pretty_default(input$profile_category), "<br>",
        "Niveau: ", df$level, "<br>",
        "Conditie: ", df$condition_label, "<br>",
        "Geslacht: ", df$geslacht, "<br>",
        "Gebied: ", df$amsterdam, "<br>",
        "Waarde: ", scales::percent(df$share, accuracy = 0.1, decimal.mark = ",")
      )

      ggplot(df, aes(x = level, y = share, fill = condition_label, text = tooltip)) +
        geom_col(position = position_dodge(width = 0.8), width = 0.72) +
        coord_flip() +
        facet_grid(geslacht ~ amsterdam) +
        scale_fill_manual(values = palette_values) +
        scale_y_continuous(labels = scales::label_percent(accuracy = 1, decimal.mark = ",")) +
        labs(
          title = title_text,
          subtitle = pretty_default(input$profile_category),
          fill = NULL
        ) +
        theme_dashboard()
    }
  })

  output$plot_profile <- plotly::renderPlotly({
    plotly::ggplotly(profile_plot_obj(), tooltip = "text") |>
      apply_plotly_transparent_layout(
        legend = list(
          orientation = "h",
          x = 0,
          xanchor = "left",
          y = -0.2
        )
      )
  })

  output$dl_profile_plot <- downloadHandler(
    filename = function() {
      paste0(
        build_export_name(
          "grafiek_profielen",
          input$profile_base %||% "dataset",
          input$profile_period %||% "periode",
          input$profile_category %||% "categorie"
        ),
        ".png"
      )
    },
    content = function(file) {
      save_plot_png(file, profile_plot_obj())
    }
  )

  output$tbl_profile <- renderDT({
    display_table(
      profile_filtered(),
      export_name = build_export_name(
        "profielen",
        input$profile_base %||% "dataset",
        input$profile_period %||% "periode",
        input$profile_category %||% "categorie"
      )
    )
  })

  output$map_sex_ui <- renderUI({
    if (!identical(input$map_group, "sex")) {
      return(NULL)
    }
    sex_choices <- map_long |>
      dplyr::filter(group_key == "sex") |>
      dplyr::pull(geslacht) |>
      as.character() |>
      unique() |>
      sort()
    if (length(sex_choices) == 0) {
      return(NULL)
    }
    radioButtons(
      "map_sex",
      "Geslacht",
      choices = stats::setNames(sex_choices, sex_choices),
      selected = sex_choices[[1]]
    )
  })

  observeEvent(input$map_group, {
    df <- map_long |>
      dplyr::filter(group_key == input$map_group)
    kinds <- setdiff(unique(df$value_kind), "expected")
    kind_labels <- c(
      "Geobserveerd" = "observed",
      "Index" = "ratio"
    )
    if (length(kinds) == 0) {
      updateSelectInput(session, "map_value_kind", choices = character(0), selected = character(0))
      return()
    }
    available_kind_choices <- kind_labels[match(kinds, unname(kind_labels))]
    selected_kind <- input$map_value_kind
    if (is.null(selected_kind) || !nzchar(selected_kind) || !(selected_kind %in% kinds)) {
      selected_kind <- kinds[[1]]
    }
    updateSelectInput(
      session, "map_value_kind",
      choices = available_kind_choices,
      selected = selected_kind
    )
  }, ignoreNULL = FALSE)

  observeEvent(list(input$map_group, input$map_value_kind, input$map_sex), {
    req(input$map_group, input$map_value_kind)
    df <- map_long |>
      dplyr::filter(group_key == input$map_group, value_kind == input$map_value_kind)

    if (identical(input$map_group, "sex")) {
      df <- df |>
        dplyr::filter(geslacht == (input$map_sex %||% unique(df$geslacht)[1]))
    }

    families <- df |>
      dplyr::distinct(family_key, family_label) |>
      dplyr::arrange(family_label)
    if (nrow(families) == 0) {
      updateSelectInput(session, "map_family", choices = character(0), selected = character(0))
      return()
    }

    selected_family <- input$map_family
    if (is.null(selected_family) || !nzchar(selected_family) || !(selected_family %in% families$family_key)) {
      selected_family <- families$family_key[[1]]
    }

    updateSelectInput(
      session, "map_family",
      choices = stats::setNames(families$family_key, families$family_label),
      selected = selected_family
    )
  }, ignoreNULL = FALSE)

  observeEvent(list(input$map_group, input$map_value_kind, input$map_family, input$map_sex), {
    req(input$map_group, input$map_value_kind, input$map_family)
    df <- map_long |>
      dplyr::filter(
        group_key == input$map_group,
        value_kind == input$map_value_kind,
        family_key == input$map_family
      )

    if (identical(input$map_group, "sex")) {
      df <- df |>
        dplyr::filter(geslacht == (input$map_sex %||% unique(df$geslacht)[1]))
    }

    outcomes <- df |>
      dplyr::distinct(outcome_key, outcome_label) |>
      dplyr::arrange(outcome_label)
    if (nrow(outcomes) == 0) {
      updateSelectInput(session, "map_outcome", choices = character(0), selected = character(0))
      return()
    }

    selected_outcome <- input$map_outcome
    if (is.null(selected_outcome) || !nzchar(selected_outcome) || !(selected_outcome %in% outcomes$outcome_key)) {
      selected_outcome <- outcomes$outcome_key[[1]]
    }

    updateSelectInput(
      session, "map_outcome",
      choices = stats::setNames(outcomes$outcome_key, outcomes$outcome_label),
      selected = selected_outcome
    )
  }, ignoreNULL = FALSE)

  map_selected_values <- reactive({
    req(input$map_group, input$map_value_kind, input$map_family, input$map_outcome)
    min_count <- parse_map_min_count(input$map_min_count)
    df <- map_long |>
      dplyr::filter(
        group_key == input$map_group,
        value_kind == input$map_value_kind,
        family_key == input$map_family,
        outcome_key == input$map_outcome
      )

    if (identical(input$map_group, "sex")) {
      df <- df |>
        dplyr::filter(geslacht == (input$map_sex %||% unique(df$geslacht)[1]))
    }

    apply_map_masks(
      df,
      exclude_westpoort = isTRUE(input$map_exclude_westpoort),
      min_count = min_count
    )
  })

  map_scale_source <- reactive({
    req(input$map_group, input$map_value_kind, input$map_family, input$map_outcome)
    min_count <- parse_map_min_count(input$map_min_count)
    df <- map_long |>
      dplyr::filter(
        group_key == input$map_group,
        value_kind == input$map_value_kind,
        family_key == input$map_family,
        outcome_key == input$map_outcome
      )

    apply_map_masks(
      df,
      exclude_westpoort = isTRUE(input$map_exclude_westpoort),
      min_count = min_count
    )
  })

  map_prefill_scale_limits <- reactive({
    req(input$map_value_kind)
    values <- if (isTRUE(input$map_fixed_legend)) map_scale_source()$masked_value else map_selected_values()$masked_value
    limits <- expand_equal_range(values)

    if (identical(input$map_value_kind, "ratio")) {
      limits[[1]] <- min(limits[[1]], 0.99)
      limits[[2]] <- max(limits[[2]], 1.01)
      if (limits[[1]] >= 1) {
        limits[[1]] <- 0.99
      }
      if (limits[[2]] <= 1) {
        limits[[2]] <- 1.01
      }
    }

    limits
  })

  observeEvent(
    list(
      input$map_group,
      input$map_value_kind,
      input$map_family,
      input$map_outcome,
      input$map_sex,
      input$map_fixed_legend,
      input$map_min_count,
      input$map_exclude_westpoort
    ),
    {
      req(input$map_value_kind %in% c("observed", "ratio"))
      limits <- map_prefill_scale_limits()
      step_value <- if (identical(input$map_value_kind, "ratio")) 0.01 else 0.1

      updateNumericInput(session, "map_scale_min", value = signif(limits[[1]], 4), step = step_value)
      updateNumericInput(session, "map_scale_max", value = signif(limits[[2]], 4), step = step_value)
    },
    ignoreNULL = FALSE
  )

  map_joined <- reactive({
    validate(need(!is.null(map_sf), "map.gpkg is niet gevonden, dus de kaart kan niet worden getekend."))
    map_sf |>
      dplyr::left_join(map_selected_values(), by = "stadsdelen")
  })

  map_plot_obj <- reactive({
    df_map <- map_joined()
    selected_values <- map_selected_values()
    scale_values <- if (isTRUE(input$map_fixed_legend)) map_scale_source()$masked_value else selected_values$masked_value
    scale_limits <- expand_equal_range(scale_values)
    selected_limits <- expand_equal_range(selected_values$masked_value)
    legend_title <- map_legend_title(input$map_family, input$map_value_kind)
    map_colors <- map_color_values()
    observed_high_color <- map_colors$observed_high
    ratio_low_color <- map_colors$ratio_low
    ratio_high_color <- map_colors$ratio_high
    manual_scale_limits <- NULL

    if (identical(input$map_value_kind, "ratio") && isTRUE(input$map_manual_scale)) {
      min_value <- suppressWarnings(as.numeric(input$map_scale_min))
      max_value <- suppressWarnings(as.numeric(input$map_scale_max))
      validate(need(is.finite(min_value) && is.finite(max_value), "Vul geldige waarden in voor min en max."))
      validate(need(min_value < max_value, "Min moet kleiner zijn dan max."))
      validate(need(min_value < 1 && max_value > 1, "Kies min < 1 en max > 1 zodat wit gelijk blijft aan 1."))
      manual_scale_limits <- c(min_value, max_value)
    }

    if (identical(input$map_value_kind, "observed") && isTRUE(input$map_manual_scale)) {
      min_value <- suppressWarnings(as.numeric(input$map_scale_min))
      max_value <- suppressWarnings(as.numeric(input$map_scale_max))
      validate(need(is.finite(min_value) && is.finite(max_value), "Vul geldige waarden in voor min en max."))
      validate(need(min_value < max_value, "Min moet kleiner zijn dan max."))
      manual_scale_limits <- c(min_value, max_value)
    }

    validate(need(nrow(df_map) > 0, "Geen kaartgegevens beschikbaar voor de huidige selectie."))

    p <- ggplot(df_map) +
      geom_sf(aes(fill = masked_value), color = alpha("#415a77", 0.45), linewidth = 0.25)

    if (isTRUE(input$map_show_labels)) {
      label_points <- suppressWarnings(sf::st_point_on_surface(df_map))
      p <- p + geom_sf_text(data = label_points, aes(label = short_label), size = 4.1, color = "black")
    }

    if (identical(input$map_value_kind, "ratio")) {
      p <- p +
        scale_fill_gradient2(
          low = ratio_low_color,
          mid = "white",
          high = ratio_high_color,
          midpoint = 1,
          limits = manual_scale_limits %||% if (isTRUE(input$map_fixed_legend)) scale_limits else NULL,
          na.value = "lightgrey",
          name = legend_title
        )
    } else {
      p <- p +
        scale_fill_gradient2(
          low = "white",
          mid = "white",
          high = observed_high_color,
          midpoint = if (!is.null(manual_scale_limits)) manual_scale_limits[[1]] else if (isTRUE(input$map_fixed_legend)) scale_limits[[1]] else selected_limits[[1]],
          limits = manual_scale_limits %||% if (isTRUE(input$map_fixed_legend)) scale_limits else NULL,
          na.value = "lightgrey",
          name = legend_title
        )
    }

    subtitle_parts <- c(
      map_selected_values()$value_label[[1]],
      pretty_family(input$map_family),
      if (identical(input$map_group, "sex")) input$map_sex else "Totaal"
    )

    p +
      guides(fill = guide_colorbar(barheight = grid::unit(7, "cm"), barwidth = grid::unit(1, "cm"))) +
      labs(
        title = map_selected_values()$outcome_label[[1]],
        subtitle = paste(subtitle_parts, collapse = " | ")
      ) +
      coord_sf(datum = NA) +
      theme_minimal(base_size = 14) +
      theme(
        plot.title = element_text(face = "bold", color = "#1f3c44", size = 17),
        plot.subtitle = element_text(color = "#5b6470"),
        panel.grid = element_blank(),
        panel.background = element_rect(fill = "transparent", colour = NA),
        plot.background = element_rect(fill = "transparent", colour = NA),
        axis.text = element_blank(),
        axis.title = element_blank(),
        axis.ticks = element_blank(),
        legend.position = "right",
        legend.background = element_rect(fill = "transparent", colour = NA),
        legend.key = element_rect(fill = "transparent", colour = NA),
        legend.title = element_text(face = "bold"),
        rect = element_rect(fill = "transparent", colour = NA)
      )
  })

  output$plot_map <- renderPlot({
    map_plot_obj()
  }, res = 110, bg = "transparent")

  map_hovered_row <- reactive({
    hover <- input$plot_map_hover
    req(hover$x, hover$y)
    df_map <- map_joined()
    point_sf <- sf::st_sfc(
      sf::st_point(c(hover$x, hover$y)),
      crs = sf::st_crs(df_map)
    )
    hit_index <- which(as.vector(sf::st_intersects(df_map, point_sf, sparse = FALSE)))
    if (length(hit_index) == 0) {
      return(NULL)
    }
    df_map[hit_index[[1]], , drop = FALSE]
  })

  output$plot_map_hover_ui <- renderUI({
    hover <- input$plot_map_hover
    row <- map_hovered_row()
    if (is.null(hover) || is.null(row)) {
      return(NULL)
    }

    min_count <- parse_map_min_count(input$map_min_count)
    tooltip_html <- build_map_tooltip(
      sf::st_drop_geometry(row),
      exclude_westpoort = isTRUE(input$map_exclude_westpoort),
      min_count = min_count
    )[[1]]

    left_pos <- hover$coords_css$x + 15
    top_pos <- hover$coords_css$y + 15

    tags$div(
      style = paste(
        "position:absolute;",
        sprintf("left:%spx;", left_pos),
        sprintf("top:%spx;", top_pos),
        "max-width:240px;",
        "padding:10px 12px;",
        "background:rgba(255,255,255,0.96);",
        "border:1px solid #cfd8df;",
        "border-radius:6px;",
        "box-shadow:0 2px 10px rgba(0,0,0,0.12);",
        "pointer-events:none;",
        "font-size:13px;",
        "line-height:1.4;",
        "z-index:20;"
      ),
      HTML(tooltip_html)
    )
  })

  output$dl_map <- downloadHandler(
    filename = function() {
      group_part <- if (identical(input$map_group, "sex")) input$map_sex %||% "sex" else "total"
      paste0(
        "map_",
        sanitize_filename(input$map_outcome %||% "outcome"),
        "_",
        sanitize_filename(input$map_value_kind %||% "type"),
        "_",
        sanitize_filename(group_part),
        ".png"
      )
    },
    content = function(file) {
      save_png_export(file, map_plot_obj())
    }
  )

  output$tbl_map <- renderDT({
    df <- map_selected_values() |>
      dplyr::select(
        stadsdelen,
        geslacht,
        n_totaal,
        count,
        value_label,
        outcome_label,
        display_value,
        masked_value
      )
    display_table(
      df,
      export_name = build_export_name(
        "kaartgegevens",
        input$map_outcome %||% "uitkomst",
        input$map_value_kind %||% "type",
        if (identical(input$map_group, "sex")) input$map_sex %||% "sex" else "total"
      )
    )
  })
}

shinyApp(ui, server)
