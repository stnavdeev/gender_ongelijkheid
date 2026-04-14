packages <- c(
  "shiny",
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

data_path <- resolve_existing_path(c("output.xlsx", "data/output.xlsx"))
if (is.na(data_path)) {
  stop("Could not find output.xlsx in the project root or in a data/ folder.")
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

pretty_profile_condition <- function(x, condition_label) {
  dplyr::case_when(
    as.character(x) == "1" ~ paste("Met", condition_label),
    as.character(x) == "0" ~ paste("Zonder", condition_label),
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

build_sheet_index <- function(sheets, suffix_regex) {
  rows <- lapply(sheets, function(sheet) {
    base_name <- sub(suffix_regex, "", sheet)
    context <- parse_dataset_context(base_name)
    transform(
      context,
      sheet = sheet,
      sheet_label = if (is.na(context$condition_key)) {
        paste(context$dataset_label, "| Alle personen")
      } else {
        paste(context$dataset_label, "|", context$condition_label)
      }
    )
  })

  dplyr::bind_rows(rows) |>
    dplyr::arrange(dataset_label, condition_label, sheet)
}

sheet_names <- openxlsx::getSheetNames(data_path)

yearly_index <- build_sheet_index(sheet_names[stringr::str_detect(sheet_names, "_yr$")], "_yr$")
es_mean_index <- build_sheet_index(sheet_names[stringr::str_detect(sheet_names, "_es$")], "_es$")
es_ci_index <- build_sheet_index(sheet_names[stringr::str_detect(sheet_names, "_es_ci$")], "_es_ci$")

profile_index <- {
  profile_sheets <- sheet_names[stringr::str_detect(sheet_names, "^profile_")]
  rows <- lapply(profile_sheets, function(sheet) {
    match <- stringr::str_match(sheet, "^profile_(.*)_(before|after)_match$")
    condition_key <- match[, 2]
    stage_key <- match[, 3]
    data.frame(
      sheet = sheet,
      condition_key = condition_key,
      condition_label = pretty_condition(condition_key),
      stage_key = stage_key,
      stage_label = dplyr::case_when(
        stage_key == "before" ~ "Voor matching",
        stage_key == "after" ~ "Na matching",
        TRUE ~ pretty_default(stage_key)
      ),
      sheet_label = paste(pretty_condition(condition_key), "|", dplyr::case_when(
        stage_key == "before" ~ "Voor matching",
        stage_key == "after" ~ "Na matching",
        TRUE ~ pretty_default(stage_key)
      )),
      stringsAsFactors = FALSE
    )
  })
  dplyr::bind_rows(rows) |>
    dplyr::arrange(condition_label, stage_key)
}

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
    "t", "coef", "lo", "hi", "n"
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

display_table <- function(df, export_name = "gegevens", show_excel = TRUE) {
  table_options <- list(
    dom = if (isTRUE(show_excel)) "Bfrtip" else "frtip",
    buttons = if (isTRUE(show_excel)) list(list(
      extend = "excel",
      text = "Gegevens downloaden",
      filename = export_name
    )) else list(),
    pageLength = 10,
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
  vline = NULL,
  facet_var = NULL,
  compact_facets = FALSE,
  show_confidence = FALSE,
  include_zero = FALSE
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
    scale_y_continuous(labels = y_labeler) +
    labs(title = title, subtitle = subtitle, color = NULL, linetype = NULL) +
    theme_dashboard()

  if (isTRUE(include_zero)) {
    p <- p + expand_limits(y = 0)
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
  dash_map <- condition_linetype_values(linetype_labels)
  dash_order <- match(unname(dash_map), c("solid", "dashed", "dot", "dashdot"))
  dash_order[is.na(dash_order)] <- 99
  linetype_labels <- linetype_labels[order(dash_order, linetype_labels)]

  build_sample <- function(label, color) {
    dash_value <- dash_map[[label]] %||% "solid"
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
      vapply(linetype_labels, build_sample, character(1), color = color),
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
      ratio_row$value_label <- "Afwijking t.o.v. verwachting"
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
    return("Afwijking t.o.v. verwachting")
  }
  if (family_key == "heeft") {
    return("Prevalentie (%)")
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
  header = tags$div(
    style = "padding: 12px 18px 4px 18px; color: #4b5563; font-size: 14px;",
    "Interactieve tool voor het verkennen van verschillen tussen mannen en vrouwen in zorggebruik, zorgkosten en aandoeningsprevalentie in Amsterdam en Nederland."
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
        uiOutput("year_condition_ui")
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
              checkboxGroupInput(
                "es_mean_include_zero",
                "Nul op y-as opnemen",
                choices = c("Ja" = "yes"),
                selected = "yes"
              ),
              uiOutput("es_mean_condition_ui")
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
              checkboxGroupInput("es_ci_sample", "Gebied", choices = NULL)
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
    "Profielen",
    sidebarLayout(
      sidebarPanel(
        selectInput(
          "profile_sheet",
          "Dataset",
          choices = stats::setNames(profile_index$sheet, profile_index$sheet_label),
          selected = profile_index$sheet[[1]]
        ),
        selectInput("profile_category", "Categorie", choices = NULL),
        checkboxGroupInput("profile_sex", "Geslacht", choices = NULL),
        checkboxGroupInput("profile_region", "Gebied", choices = NULL)
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
              helpText("Min is wit en max is donkerroze.")
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
  output$tbl_scripts <- renderDT({
    display_table(script_index, export_name = "code_overzicht")
  })

  output$tbl_sheets <- renderDT({
    display_table(
      data.frame(sheet = sheet_names, stringsAsFactors = FALSE),
      export_name = "werkbladen_output_xlsx"
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
      selected = outcomes[[1]]
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
      selected = types[[1]]
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

  output$plot_year <- plotly::renderPlotly({
    df <- year_filtered()
    validate(need(nrow(df) > 0, "Geen gegevens beschikbaar voor de huidige selectie."))
    sheet_meta <- yearly_index[yearly_index$sheet == input$year_sheet, , drop = FALSE]
    custom_annotations <- build_event_study_annotations(df)
    if (length(custom_annotations) == 0) {
      custom_annotations <- build_sex_annotations(df)
    }
    bottom_margin <- if (length(custom_annotations) > 1) 115 else 90

    p <- build_time_series_plot(
      df = df,
      x_var = "year",
      title = paste(sheet_meta$dataset_label[[1]], "-", sheet_meta$condition_label[[1]]),
      subtitle = paste(pretty_metric_name(input$year_outcome), "|", pretty_type(input$year_type)),
      y_labeler = metric_labeler(input$year_outcome, input$year_type),
      facet_var = "amsterdam",
      compact_facets = TRUE,
      show_confidence = TRUE,
      include_zero = "yes" %in% (input$year_include_zero %||% character(0))
    )

    plotly::ggplotly(p, tooltip = "text") |>
      plotly::layout(
        showlegend = FALSE,
        annotations = custom_annotations,
        margin = list(b = bottom_margin)
      )
  })

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
      selected = outcomes[[1]]
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
      selected = types[[1]]
    )
  }, ignoreNULL = FALSE)

  output$es_mean_condition_ui <- renderUI({
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

    if (!is.na(flag_col)) {
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

  output$plot_es_mean <- plotly::renderPlotly({
    df <- es_mean_filtered()
    validate(need(nrow(df) > 0, "Geen gegevens beschikbaar voor de huidige selectie."))
    sheet_meta <- es_mean_index[es_mean_index$sheet == input$es_mean_sheet, , drop = FALSE]
    custom_annotations <- build_event_study_annotations(df)
    bottom_margin <- if (length(custom_annotations) > 0) 115 else 70

    p <- build_time_series_plot(
      df = df,
      x_var = "years_since_diagnosis",
      title = paste(sheet_meta$dataset_label[[1]], "-", sheet_meta$condition_label[[1]]),
      subtitle = paste(pretty_metric_name(input$es_mean_outcome), "|", pretty_type(input$es_mean_type)),
      y_labeler = metric_labeler(input$es_mean_outcome, input$es_mean_type),
      vline = 0,
      facet_var = "amsterdam",
      compact_facets = TRUE,
      include_zero = "yes" %in% (input$es_mean_include_zero %||% character(0))
    )

    plotly::ggplotly(p, tooltip = "text") |>
      plotly::layout(
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

  output$tbl_es_mean <- renderDT({
    display_table(
      es_mean_filtered(),
      export_name = build_export_name(
        "event_study_geobserveerde_gemiddelden",
        input$es_mean_sheet %||% "dataset",
        input$es_mean_outcome %||% "uitkomst",
        input$es_mean_type %||% "statistiek"
      )
    )
  })

  observeEvent(input$es_ci_sheet, {
    df <- read_sheet_cached(input$es_ci_sheet)
    outcomes <- sort(unique(as.character(df$outcome)))
    updateSelectInput(
      session, "es_ci_outcome",
      choices = stats::setNames(outcomes, pretty_metric_name(outcomes)),
      selected = outcomes[[1]]
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

  output$plot_es_ci <- plotly::renderPlotly({
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
        title = paste(sheet_meta$dataset_label[[1]], "-", sheet_meta$condition_label[[1]]),
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

    plotly::ggplotly(p, tooltip = "text") |>
      plotly::layout(
        legend = list(
          orientation = "h",
          x = 0,
          xanchor = "left",
          y = -0.2
        )
      )
  })

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

  observeEvent(input$profile_sheet, {
    df <- read_sheet_cached(input$profile_sheet)
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
    req(input$profile_sheet, input$profile_category)
    df <- read_sheet_cached(input$profile_sheet)
    selected_sex <- input$profile_sex %||% unique(as.character(df$geslacht))
    selected_region <- input$profile_region %||% unique(as.character(df$amsterdam))
    profile_meta <- profile_index[profile_index$sheet == input$profile_sheet, , drop = FALSE]

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

  output$plot_profile <- plotly::renderPlotly({
    df <- profile_filtered()
    validate(need(nrow(df) > 0, "Geen gegevens beschikbaar voor de huidige selectie."))
    palette_values <- profile_palette_for_labels(unique(df$condition_label))
    title_text <- paste(
      profile_index$condition_label[profile_index$sheet == input$profile_sheet],
      "-",
      profile_index$stage_label[profile_index$sheet == input$profile_sheet]
    )

    p <- if (identical(input$profile_category, "leeftijd")) {
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

    plotly::ggplotly(p, tooltip = "text") |>
      plotly::layout(
        legend = list(
          orientation = "h",
          x = 0,
          xanchor = "left",
          y = -0.2
        )
      )
  })

  output$tbl_profile <- renderDT({
    display_table(
      profile_filtered(),
      export_name = build_export_name(
        "profielen",
        input$profile_sheet %||% "dataset",
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
      "Afwijking t.o.v. verwachting" = "ratio"
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
          low = "darkslateblue",
          mid = "white",
          high = "red",
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
          high = "deeppink4",
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
        axis.text = element_blank(),
        axis.title = element_blank(),
        axis.ticks = element_blank(),
        legend.position = "right",
        legend.title = element_text(face = "bold"),
        rect = element_blank()
      )
  })

  output$plot_map <- renderPlot({
    map_plot_obj()
  }, res = 110)

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
      ggplot2::ggsave(file, plot = map_plot_obj(), scale = 0.7, width = 14, height = 10, dpi = 300, bg = "white")
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
