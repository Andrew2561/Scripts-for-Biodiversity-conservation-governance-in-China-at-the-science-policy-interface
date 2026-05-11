#!/usr/bin/env Rscript

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(tidyr)
  library(stringr)
  library(ggplot2)
  library(ggrepel)
  library(scales)
  library(tibble)
  library(openxlsx)
})

options(stringsAsFactors = FALSE, scipen = 999)
Sys.setenv(LANG = "en_US.UTF-8", LC_ALL = "en_US.UTF-8")
try(Sys.setlocale("LC_ALL", "en_US.UTF-8"), silent = TRUE)
try(Sys.setlocale("LC_CTYPE", "en_US.UTF-8"), silent = TRUE)

get_script_path <- function() {
  cmd_args <- commandArgs(trailingOnly = FALSE)
  file_arg <- "--file="
  file_match <- grep(file_arg, cmd_args)

  if (length(file_match) > 0) {
    script_path <- sub(file_arg, "", cmd_args[file_match[1]])
    return(normalizePath(script_path, winslash = "/", mustWork = TRUE))
  }

  if (!is.null(sys.frames()[[1]]$ofile)) {
    return(normalizePath(sys.frames()[[1]]$ofile, winslash = "/", mustWork = TRUE))
  }

  if (requireNamespace("rstudioapi", quietly = TRUE)) {
    script_path <- tryCatch(
      rstudioapi::getActiveDocumentContext()$path,
      error = function(e) ""
    )
    if (nzchar(script_path)) {
      return(normalizePath(script_path, winslash = "/", mustWork = TRUE))
    }
  }

  stop(
    "Cannot determine script path.\n",
    "Please run this file as a saved script, not as copied console lines."
  )
}

resolve_project_root <- function() {
  script_path <- get_script_path()
  script_dir <- dirname(script_path)
  parent <- normalizePath(file.path(script_dir, ".."), winslash = "/", mustWork = TRUE)

  if (
    dir.exists(file.path(parent, "policy text")) &&
    dir.exists(file.path(parent, "bibliometric analysis"))
  ) {
    return(parent)
  }

  if (
    dir.exists(file.path(script_dir, "policy text")) &&
    dir.exists(file.path(script_dir, "bibliometric analysis"))
  ) {
    return(script_dir)
  }

  stop(
    "Cannot resolve project_root from script location.\n",
    "Script path: ", script_path, "\n",
    "Script directory: ", script_dir, "\n",
    "Checked:\n",
    "  1) ", parent, "\n",
    "  2) ", script_dir, "\n",
    "But neither contains both 'policy text' and 'bibliometric analysis'."
  )
}

resolve_plot_font_family <- function() {
  "sans"
}

configure_plot_text_defaults <- function(font_family) {
  for (geom_name in c("text", "label")) {
    try(update_geom_defaults(geom_name, list(family = font_family)), silent = TRUE)
  }
}

theme_keyword_map <- function(base_size = 11) {
  theme_minimal(base_size = base_size, base_family = plot_font_family) +
    theme(
      plot.title = element_text(face = "bold", size = base_size + 4, colour = "#16203A"),
      plot.subtitle = element_text(size = base_size, colour = "#4F5D75"),
      plot.caption = element_text(size = base_size - 2, colour = "#667085"),
      axis.title = element_text(face = "bold", colour = "#16203A"),
      axis.text = element_text(colour = "#243447"),
      panel.grid.minor = element_blank(),
      panel.grid.major.y = element_blank(),
      panel.grid.major.x = element_line(colour = alpha("#94A3B8", 0.18), linewidth = 0.35),
      legend.position = "bottom",
      legend.title = element_text(face = "bold"),
      plot.background = element_rect(fill = "#FFFFFF", colour = NA),
      panel.background = element_rect(fill = "#FFFFFF", colour = NA),
      axis.ticks = element_blank()
    )
}

save_plot <- function(plot_obj, filename, width = 18, height = 22, dpi = 320) {
  ggsave_args <- list(
    filename = filename,
    plot = plot_obj,
    width = width,
    height = height,
    dpi = dpi,
    bg = "#FFFFFF"
  )

  if (tolower(tools::file_ext(filename)) == "png" && requireNamespace("ragg", quietly = TRUE)) {
    ggsave_args$device <- ragg::agg_png
  }

  do.call(ggsave, ggsave_args)
}

save_plot_with_pdf <- function(plot_obj, filename, width = 18, height = 22, dpi = 320) {
  save_plot(plot_obj, filename, width = width, height = height, dpi = dpi)
  ggsave(
    filename = paste0(tools::file_path_sans_ext(filename), ".pdf"),
    plot = plot_obj,
    width = width,
    height = height,
    bg = "#FFFFFF",
    device = grDevices::pdf,
    useDingbats = FALSE
  )
}

draw_key_line_only <- function(data, params, size) {
  line_colour <- if (!is.null(data$colour) && !is.na(data$colour)) data$colour else "black"
  line_alpha <- if (!is.null(data$alpha) && !is.na(data$alpha)) data$alpha else 1
  line_width <- if (!is.null(data$linewidth) && !is.na(data$linewidth)) data$linewidth else 0.5
  line_type <- if (!is.null(data$linetype) && !is.na(data$linetype)) data$linetype else 1

  grid::segmentsGrob(
    x0 = 0.08,
    y0 = 0.5,
    x1 = 0.92,
    y1 = 0.5,
    gp = grid::gpar(
      col = alpha(line_colour, line_alpha),
      lwd = line_width * 2.845276,
      lty = line_type,
      lineend = "butt"
    )
  )
}

make_display_labels <- function(theme_en) {
  str_wrap(theme_en, width = 14)
}

standardize_figure_keyword_name <- function(theme_en) {
  dplyr::recode(
    theme_en,
    "Returning Farmland to Forest" = "Sloping Land Conversion Program",
    "Sanjiangyuan" = "Sanjiang Plain",
    .default = theme_en
  )
}

keyword_match_key <- function(theme_en) {
  theme_en |>
    as.character() |>
    standardize_figure_keyword_name() |>
    str_replace_all("[._]", " ") |>
    str_squish() |>
    str_to_lower()
}

order_policy_domains <- function(domains) {
  clean_domains <- str_replace_na(as.character(domains), "Other Policy Domains")
  observed_domains <- unique(clean_domains[clean_domains != ""])
  c(
    sort(setdiff(observed_domains, "Other Policy Domains")),
    intersect("Other Policy Domains", observed_domains)
  )
}

assign_band_positions <- function(df, y_min, y_max, order_cols) {
  if (nrow(df) == 0) {
    return(df |> mutate(y = numeric(0)))
  }

  inner_min <- y_min + 0.20
  inner_max <- y_max - 0.20

  ordered_df <- df |>
    arrange(year, across(all_of(order_cols))) |>
    group_by(year) |>
    arrange(across(all_of(order_cols)), .by_group = TRUE) |>
    mutate(
      n_in_year = n(),
      rank_in_year = row_number(),
      y = if_else(
        n_in_year == 1L,
        (inner_min + inner_max) / 2,
        inner_max - ((rank_in_year - 1) * (inner_max - inner_min) / (n_in_year - 1))
      )
    ) |>
    ungroup()

  ordered_df
}

extract_tag_value <- function(record_lines, tag) {
  tag_prefix <- paste0(tag, " ")
  tag_start <- which(startsWith(record_lines, tag_prefix))[1]

  if (is.na(tag_start)) {
    return(NA_character_)
  }

  collected <- sub(paste0("^", tag, "\\s+"), "", record_lines[tag_start])
  next_line <- tag_start + 1L

  while (
    next_line <= length(record_lines) &&
      !str_detect(record_lines[next_line], "^[A-Z0-9]{2} ")
  ) {
    collected <- c(collected, record_lines[next_line])
    next_line <- next_line + 1L
  }

  str_squish(paste(collected, collapse = " "))
}

parse_citespace_download <- function(path) {
  record_lines <- readLines(path, warn = FALSE, encoding = "UTF-8")
  record_starts <- which(startsWith(record_lines, "PT "))

  if (length(record_starts) == 0) {
    return(tibble(record_id = character(0), year = integer(0), keyword_zh = character(0)))
  }

  parsed_records <- lapply(seq_along(record_starts), function(i) {
    record_start <- record_starts[i]
    record_end <- if (i < length(record_starts)) {
      record_starts[i + 1L] - 1L
    } else {
      length(record_lines)
    }

    current_record <- record_lines[record_start:record_end]
    record_year <- suppressWarnings(as.integer(extract_tag_value(current_record, "PY")))
    record_keywords <- extract_tag_value(current_record, "DE")
    record_ut <- extract_tag_value(current_record, "UT")

    if (is.na(record_year) || is.na(record_keywords) || !nzchar(record_keywords)) {
      return(tibble(record_id = character(0), year = integer(0), keyword_zh = character(0)))
    }

    tibble(
      record_id = if_else(
        is.na(record_ut) || !nzchar(record_ut),
        paste(path, i, sep = "::"),
        record_ut
      ),
      year = record_year,
      keyword_zh = str_split(record_keywords, ";", simplify = FALSE)[[1]]
    ) |>
      mutate(keyword_zh = str_squish(keyword_zh)) |>
      filter(keyword_zh != "")
  })

  bind_rows(parsed_records)
}

load_keyword_translation_dictionary <- function(path) {
  if (!file.exists(path)) {
    stop("Required keyword translation dictionary does not exist: ", path)
  }

  read.xlsx(path, colNames = FALSE) |>
    transmute(
      theme_en = standardize_figure_keyword_name(str_squish(as.character(X1))),
      theme_key = keyword_match_key(theme_en),
      keyword_zh = str_squish(as.character(X2))
    ) |>
    filter(
      !is.na(theme_en),
      !is.na(keyword_zh),
      theme_en != "",
      keyword_zh != "",
      !str_detect(theme_en, "^#")
    ) |>
    distinct(theme_key, keyword_zh, .keep_all = TRUE)
}

load_policy_keyword_yearly_appearance <- function(path) {
  if (!file.exists(path)) {
    stop("Required policy keyword match workbook does not exist: ", path)
  }

  read.xlsx(path, sheet = "KeywordMatchesLong") |>
    transmute(
      theme_en = standardize_figure_keyword_name(str_squish(as.character(keyword_en))),
      theme_key = keyword_match_key(theme_en),
      target_domain = "Policy",
      year = suppressWarnings(as.integer(year)),
      policy_id = as.character(policy_id)
    ) |>
    filter(!is.na(theme_key), theme_key != "", !is.na(year)) |>
    group_by(theme_key, target_domain, year) |>
    summarise(appearance_count = n_distinct(policy_id), .groups = "drop") |>
    mutate(present = appearance_count > 0)
}

load_research_keyword_yearly_matrix <- function(path) {
  if (!file.exists(path)) {
    return(tibble(theme_key = character(0), target_domain = character(0), year = integer(0), appearance_count = numeric(0), present = logical(0)))
  }

  yearly_matrix <- read.xlsx(path)
  if (ncol(yearly_matrix) < 2) {
    return(tibble(theme_key = character(0), target_domain = character(0), year = integer(0), appearance_count = numeric(0), present = logical(0)))
  }

  names(yearly_matrix)[1] <- "theme_en"
  year_columns <- intersect(names(yearly_matrix), as.character(year_min:year_max))

  if (length(year_columns) == 0) {
    return(tibble(theme_key = character(0), target_domain = character(0), year = integer(0), appearance_count = numeric(0), present = logical(0)))
  }

  yearly_matrix |>
    transmute(
      theme_en = standardize_figure_keyword_name(str_squish(as.character(theme_en))),
      theme_key = keyword_match_key(theme_en),
      across(all_of(year_columns), ~suppressWarnings(as.numeric(.x)))
    ) |>
    pivot_longer(
      cols = all_of(year_columns),
      names_to = "year",
      values_to = "appearance_count"
    ) |>
    mutate(
      target_domain = "Research",
      year = suppressWarnings(as.integer(year)),
      appearance_count = replace_na(appearance_count, 0),
      present = appearance_count > 0
    ) |>
    filter(!is.na(theme_key), theme_key != "", !is.na(year))
}

load_research_keyword_yearly_appearance <- function(cluster_dir, dictionary_path, yearly_matrix_path) {
  keyword_dictionary <- load_keyword_translation_dictionary(dictionary_path)
  download_paths <- list.files(
    cluster_dir,
    pattern = "^download_cluster_.*\\.txt$",
    recursive = TRUE,
    full.names = TRUE
  )

  record_based_appearance <- if (length(download_paths) > 0) {
    bind_rows(lapply(download_paths, parse_citespace_download)) |>
      inner_join(keyword_dictionary |> select(theme_key, keyword_zh), by = "keyword_zh") |>
      group_by(theme_key, target_domain = "Research", year) |>
      summarise(appearance_count = n_distinct(record_id), .groups = "drop") |>
      mutate(present = appearance_count > 0)
  } else {
    tibble(theme_key = character(0), target_domain = character(0), year = integer(0), appearance_count = numeric(0), present = logical(0))
  }

  matrix_based_appearance <- load_research_keyword_yearly_matrix(yearly_matrix_path)

  bind_rows(record_based_appearance, matrix_based_appearance) |>
    group_by(theme_key, target_domain, year) |>
    summarise(
      appearance_count = max(appearance_count, na.rm = TRUE),
      present = any(present, na.rm = TRUE),
      .groups = "drop"
    )
}

project_root <- resolve_project_root()
output_root <- file.path(project_root, "analysis_outputs", "biodiversity_policy_english")
interaction_dir <- file.path(output_root, "literature", "04_policy_research_interaction")
input_path <- file.path(interaction_dir, "policy_research_timing_evidence.csv")
policy_keyword_matches_path <- file.path(
  output_root,
  "policy",
  "03_keyword_network",
  "policy_document_keyword_matches_231.xlsx"
)
research_keyword_yearly_matrix_path <- file.path(
  project_root,
  "bibliometric analysis",
  "关键词发文量",
  "20250916政策关键词年发文量.xlsx"
)
research_cluster_dir <- file.path(project_root, "bibliometric analysis", "clusters")
research_keyword_dictionary_path <- file.path(project_root, "bibliometric analysis", "关键词中译英.xlsx")

if (!file.exists(input_path)) {
  stop(
    "Required file does not exist: ", input_path, "\n",
    "Please run analysis_scripts/run_biodiversity_policy_pipeline_final.R first."
  )
}

dir.create(interaction_dir, recursive = TRUE, showWarnings = FALSE)

plot_font_family <- resolve_plot_font_family()
configure_plot_text_defaults(plot_font_family)
figure_axis_title_size <- 14.5
figure_axis_text_size <- 12
figure_legend_text_size <- 11.2
figure_legend_title_size <- 12
figure_legend_text_geom_size <- figure_legend_text_size / 2.845276
figure_legend_title_geom_size <- figure_legend_title_size / 2.845276

message("Using plot font family: ", plot_font_family)
message("Reading timing evidence from: ", input_path)

evidence_raw <- read_csv(input_path, show_col_types = FALSE)

required_cols <- c(
  "theme_en",
  "theme_zh",
  "policy_first_year",
  "research_first_year"
)
missing_cols <- setdiff(required_cols, names(evidence_raw))
if (length(missing_cols) > 0) {
  stop(
    "Timing evidence file is missing required columns: ",
    paste(missing_cols, collapse = ", ")
  )
}

evidence <- evidence_raw |>
  transmute(
    theme_en = standardize_figure_keyword_name(as.character(theme_en)),
    theme_key = keyword_match_key(theme_en),
    theme_zh = as.character(theme_zh),
    theme_group = if ("theme_group" %in% names(evidence_raw)) {
      str_replace_na(as.character(theme_group), "Other Policy Domains")
    } else {
      "Other Policy Domains"
    },
    display_label = make_display_labels(theme_en),
    policy_first_year = suppressWarnings(as.integer(policy_first_year)),
    research_first_year = suppressWarnings(as.integer(research_first_year))
  ) |>
  filter(!is.na(theme_en), theme_en != "") |>
  distinct(theme_key, .keep_all = TRUE)

year_min <- min(c(evidence$policy_first_year, evidence$research_first_year), na.rm = TRUE)
year_max <- max(c(evidence$policy_first_year, evidence$research_first_year), na.rm = TRUE)

zone_spec <- tribble(
  ~zone, ~zone_label, ~y_min, ~y_max,
  "Shared-policy", "Shared keywords\n(policy first appearance)", 10.4, 18.0,
  "Shared-research", "Shared keywords\n(research first appearance)", 1.0, 8.6
)

shared_themes <- evidence |>
  filter(!is.na(policy_first_year), !is.na(research_first_year)) |>
  mutate(
    relation_type = case_when(
      policy_first_year < research_first_year ~ "Policy-first",
      research_first_year < policy_first_year ~ "Research-first",
      TRUE ~ "Same-year"
    ),
    policy_to_research_years = case_when(
      relation_type == "Policy-first" ~ research_first_year - policy_first_year,
      TRUE ~ NA_integer_
    ),
    research_to_policy_years = case_when(
      relation_type == "Research-first" ~ policy_first_year - research_first_year,
      TRUE ~ NA_integer_
    )
  )

shared_policy_nodes <- shared_themes |>
  transmute(
    theme_key,
    theme_en,
    display_label,
    zone = "Shared-policy",
    year = policy_first_year,
    partner_year = research_first_year,
    policy_first_year,
    research_first_year,
    policy_to_research_years,
    research_to_policy_years,
    order_partner_year = research_first_year,
    order_label = display_label,
    relation_type
  )

shared_research_nodes <- shared_themes |>
  transmute(
    theme_key,
    theme_en,
    display_label,
    zone = "Shared-research",
    year = research_first_year,
    partner_year = policy_first_year,
    policy_first_year,
    research_first_year,
    policy_to_research_years,
    research_to_policy_years,
    order_partner_year = policy_first_year,
    order_label = display_label,
    relation_type
  )

nodes <- bind_rows(
  assign_band_positions(shared_policy_nodes, 10.4, 18.0, c("order_partner_year", "order_label")),
  assign_band_positions(shared_research_nodes, 1.0, 8.6, c("order_partner_year", "order_label"))
) |>
  mutate(
    zone_label = if_else(
      zone == "Shared-policy",
      "Shared keywords\n(policy first appearance)",
      "Shared keywords\n(research first appearance)"
    )
  )

shared_node_pairs <- shared_themes |>
  select(
    theme_key,
    theme_en,
    theme_group,
    display_label,
    policy_first_year,
    research_first_year,
    relation_type,
    policy_to_research_years,
    research_to_policy_years
  ) |>
  left_join(
    nodes |>
      filter(zone == "Shared-policy") |>
      select(theme_en, policy_x = year, policy_y = y, policy_zone = zone),
    by = "theme_en"
  ) |>
  left_join(
    nodes |>
      filter(zone == "Shared-research") |>
      select(theme_en, research_x = year, research_y = y, research_zone = zone),
    by = "theme_en"
  ) |>
  mutate(
    x = case_when(
      relation_type == "Policy-first" ~ policy_x,
      relation_type == "Research-first" ~ research_x,
      TRUE ~ NA_real_
    ),
    y = case_when(
      relation_type == "Policy-first" ~ policy_y,
      relation_type == "Research-first" ~ research_y,
      TRUE ~ NA_real_
    ),
    xend = case_when(
      relation_type == "Policy-first" ~ research_x,
      relation_type == "Research-first" ~ policy_x,
      TRUE ~ NA_real_
    ),
    yend = case_when(
      relation_type == "Policy-first" ~ research_y,
      relation_type == "Research-first" ~ policy_y,
      TRUE ~ NA_real_
    )
  )

directed_edges <- shared_node_pairs |>
  filter(relation_type != "Same-year")

synchronous_edges <- shared_node_pairs |>
  filter(relation_type == "Same-year")

conversion_summary_spec <- tribble(
  ~summary_id, ~cohort_label, ~relation_type, ~first_year_col, ~cohort_start_year, ~cohort_end_year, ~conversion_year_col,
  "policy_first_1990_2014_to_research", "Keywords first appearing in policy during 1990-2014 and later appearing in research", "Policy-first", "policy_first_year", 1990L, 2014L, "policy_to_research_years",
  "research_first_1990_2014_to_policy", "Keywords first appearing in research during 1990-2014 and later appearing in policy", "Research-first", "research_first_year", 1990L, 2014L, "research_to_policy_years",
  "policy_first_2015_2024_to_research", "Keywords first appearing in policy during 2015-2024 and later appearing in research", "Policy-first", "policy_first_year", 2015L, 2024L, "policy_to_research_years"
)

build_conversion_summary_row <- function(spec_row, shared_theme_data) {
  cohort_data <- shared_theme_data |>
    filter(
      relation_type == spec_row$relation_type,
      .data[[spec_row$first_year_col]] >= spec_row$cohort_start_year,
      .data[[spec_row$first_year_col]] <= spec_row$cohort_end_year,
      !is.na(.data[[spec_row$conversion_year_col]])
    )

  tibble(
    summary_id = spec_row$summary_id,
    cohort_label = spec_row$cohort_label,
    relation_type = spec_row$relation_type,
    first_appearance_domain = if_else(spec_row$relation_type == "Policy-first", "Policy", "Research"),
    conversion_target_domain = if_else(spec_row$relation_type == "Policy-first", "Research", "Policy"),
    first_year_start = spec_row$cohort_start_year,
    first_year_end = spec_row$cohort_end_year,
    keyword_count = nrow(cohort_data),
    mean_conversion_years = if (nrow(cohort_data) > 0) {
      round(mean(cohort_data[[spec_row$conversion_year_col]], na.rm = TRUE), 3)
    } else {
      NA_real_
    }
  )
}

conversion_summary <- bind_rows(
  lapply(
    seq_len(nrow(conversion_summary_spec)),
    function(i) build_conversion_summary_row(conversion_summary_spec[i, ], shared_themes)
  )
)

conversion_year_range <- 1990:2024

yearly_conversion_summary <- bind_rows(
  shared_themes |>
    filter(
      relation_type == "Policy-first",
      policy_first_year %in% conversion_year_range,
      !is.na(policy_to_research_years)
    ) |>
    group_by(first_year = policy_first_year) |>
    summarise(
      keyword_count = n(),
      mean_conversion_years = round(mean(policy_to_research_years), 3),
      .groups = "drop"
    ) |>
    mutate(
      relation_type = "Policy-first",
      first_appearance_domain = "Policy",
      conversion_target_domain = "Research"
    ),
  shared_themes |>
    filter(
      relation_type == "Research-first",
      research_first_year %in% conversion_year_range,
      !is.na(research_to_policy_years)
    ) |>
    group_by(first_year = research_first_year) |>
    summarise(
      keyword_count = n(),
      mean_conversion_years = round(mean(research_to_policy_years), 3),
      .groups = "drop"
    ) |>
    mutate(
      relation_type = "Research-first",
      first_appearance_domain = "Research",
      conversion_target_domain = "Policy"
    )
)

yearly_conversion_plot_data <- bind_rows(
  tibble(
    relation_type = "Policy-first",
    first_year = conversion_year_range,
    first_appearance_domain = "Policy",
    conversion_target_domain = "Research"
  ),
  tibble(
    relation_type = "Research-first",
    first_year = conversion_year_range,
    first_appearance_domain = "Research",
    conversion_target_domain = "Policy"
  )
) |>
  left_join(
    yearly_conversion_summary,
    by = c("relation_type", "first_year", "first_appearance_domain", "conversion_target_domain")
  ) |>
  mutate(
    direction_label = recode(
      relation_type,
      "Policy-first" = "Policy to research",
      "Research-first" = "Research to policy"
    ),
    keyword_count = if_else(is.na(keyword_count), 0L, keyword_count)
  ) |>
  group_by(relation_type) |>
  arrange(first_year, .by_group = TRUE) |>
  mutate(segment_id = cumsum(is.na(mean_conversion_years))) |>
  ungroup()

zone_midpoints <- zone_spec |>
  mutate(y_mid = (y_min + y_max) / 2)

relation_fill_palette <- c(
  "Policy-first" = "#E1B22E",
  "Research-first" = "#7FA5DE",
  "Same-year" = "#7BB56A"
)

relation_arrow_palette <- relation_fill_palette
yearly_observed_palette <- c(
  "Policy-first" = "#E7C867",
  "Research-first" = "#9DBAE8"
)
yearly_trend_palette <- c(
  "Policy-first" = "#9A7300",
  "Research-first" = "#4E74B8"
)
yearly_point_alpha <- 0.98
yearly_trend_ci_alpha <- 0.15
yearly_ci_legend_alpha <- 0.28
yearly_point_size_range <- c(2.0, 5.4)

policy_label_band <- c(10.6, 17.8)
research_label_band <- c(1.2, 8.4)

band_separator_y <- mean(c(zone_spec$y_min[1], zone_spec$y_max[2]))

yearly_panel_order <- c("Policy to research", "Research to policy")

yearly_conversion_observed_data <- yearly_conversion_plot_data |>
  filter(keyword_count > 0) |>
  mutate(direction_label = factor(direction_label, levels = yearly_panel_order))

yearly_point_size_domain <- sqrt(range(yearly_conversion_observed_data$keyword_count, na.rm = TRUE))

rescale_yearly_point_size <- function(keyword_count) {
  if (abs(diff(yearly_point_size_domain)) < .Machine$double.eps) {
    return(rep(mean(yearly_point_size_range), length(keyword_count)))
  }

  scales::rescale(
    sqrt(keyword_count),
    to = yearly_point_size_range,
    from = yearly_point_size_domain
  )
}

yearly_conversion_observed_data <- yearly_conversion_observed_data |>
  mutate(point_size = rescale_yearly_point_size(keyword_count))

yearly_size_legend_breaks <- c(1L, 5L, 10L)
yearly_size_legend_breaks <- yearly_size_legend_breaks[
  yearly_size_legend_breaks <= max(yearly_conversion_observed_data$keyword_count, na.rm = TRUE)
]
yearly_size_legend_breaks <- unique(yearly_size_legend_breaks)
yearly_size_legend_sizes <- rescale_yearly_point_size(yearly_size_legend_breaks)

build_yearly_trend_data <- function(df) {
  if (nrow(df) < 2) {
    return(tibble(
      first_year = numeric(0),
      fit = numeric(0),
      conf_low = numeric(0),
      conf_high = numeric(0)
    ))
  }

  fit_model <- lm(mean_conversion_years ~ first_year, data = df)
  prediction_years <- seq(min(df$first_year), max(df$first_year))

  if (nrow(df) < 3) {
    fitted_values <- as.numeric(predict(fit_model, newdata = tibble(first_year = prediction_years)))
    return(tibble(
      first_year = prediction_years,
      fit = fitted_values,
      conf_low = fitted_values,
      conf_high = fitted_values
    ))
  }

  prediction <- predict(
    fit_model,
    newdata = tibble(first_year = prediction_years),
    interval = "confidence"
  )

  tibble(
    first_year = prediction_years,
    fit = prediction[, "fit"],
    conf_low = prediction[, "lwr"],
    conf_high = prediction[, "upr"]
  )
}

yearly_trend_data <- yearly_conversion_observed_data |>
  group_by(relation_type, direction_label) |>
  group_modify(~build_yearly_trend_data(.x)) |>
  ungroup()

yearly_y_max <- max(
  c(yearly_conversion_observed_data$mean_conversion_years, yearly_trend_data$conf_high),
  na.rm = TRUE
)
yearly_y_max <- ceiling(yearly_y_max / 5) * 5
yearly_y_breaks <- seq(0, yearly_y_max, by = 5)
yearly_trend_data <- yearly_trend_data |>
  mutate(
    conf_low = pmax(conf_low, 0),
    conf_high = pmin(conf_high, yearly_y_max)
  )

build_yearly_direction_plot <- function(
    relation_code,
    direction_label,
    observed_colour,
    trend_colour,
    panel_tag,
    show_x_title = TRUE
) {
  observed_df <- yearly_conversion_observed_data |>
    filter(relation_type == relation_code)

  trend_df <- yearly_trend_data |>
    filter(relation_type == relation_code)

  ggplot(observed_df, aes(x = first_year, y = mean_conversion_years)) +
    geom_ribbon(
      data = trend_df,
      aes(x = first_year, ymin = conf_low, ymax = conf_high, fill = "95% CI"),
      inherit.aes = FALSE,
      alpha = yearly_trend_ci_alpha,
      colour = NA,
      show.legend = TRUE
    ) +
    geom_line(
      data = trend_df,
      aes(x = first_year, y = fit),
      inherit.aes = FALSE,
      colour = trend_colour,
      linewidth = 1.15,
      show.legend = FALSE
    ) +
    geom_point(
      aes(colour = direction_label, size = point_size),
      alpha = yearly_point_alpha
    ) +
    scale_colour_manual(
      values = stats::setNames(observed_colour, direction_label),
      breaks = direction_label,
      labels = direction_label,
      name = NULL
    ) +
    scale_fill_manual(
      values = c("95% CI" = trend_colour),
      name = NULL
    ) +
    scale_size_identity() +
    scale_x_continuous(
      breaks = conversion_year_range,
      limits = c(min(conversion_year_range) - 0.5, max(conversion_year_range) + 0.5),
      expand = c(0, 0)
    ) +
    scale_y_continuous(
      breaks = yearly_y_breaks,
      limits = c(0, yearly_y_max),
      labels = label_number(accuracy = 1),
      expand = expansion(mult = c(0, 0.04))
    ) +
    labs(
      title = NULL,
      subtitle = NULL,
      x = if (show_x_title) "First Appearance Year" else NULL,
      y = "Average Years",
      tag = panel_tag
    ) +
    theme_keyword_map(11) +
    theme(
      panel.grid.major.y = element_line(colour = alpha("#94A3B8", 0.24), linewidth = 0.35),
      axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8),
      legend.position = "none",
      plot.tag = element_text(face = "bold", size = 13, colour = "#0F172A"),
      plot.tag.position = c(0, 1),
      plot.margin = margin(8, 16, 6, 28)
    ) +
    coord_cartesian(clip = "off")
}

build_yearly_direction_legend <- function(direction_label, observed_colour, trend_colour) {
  
  ggplot() +
    
    # direction label
    annotate(
      "point",
      x = 1.2,
      y = 1.05,
      colour = observed_colour,
      size = 2.6,
      alpha = yearly_point_alpha
    ) +
    
    annotate(
      "text",
      x = 1.45,
      y = 1.05,
      label = direction_label,
      hjust = 0,
      vjust = 0.5,
      size = 3.2,
      family = plot_font_family,
      colour = "#0F172A"
    ) +
    
    # linear fit
    annotate(
      "segment",
      x = 3.1,
      xend = 3.6,
      y = 1.05,
      yend = 1.05,
      colour = trend_colour,
      linewidth = 1.1,
      lineend = "butt"
    ) +
    
    annotate(
      "text",
      x = 3.75,
      y = 1.05,
      label = "Linear fit",
      hjust = 0,
      vjust = 0.5,
      size = 3.2,
      family = plot_font_family,
      colour = "#0F172A"
    ) +
    
    # CI
    annotate(
      "rect",
      xmin = 4.9,
      xmax = 5.35,
      ymin = 0.96,
      ymax = 1.14,
      fill = alpha(trend_colour, yearly_ci_legend_alpha),
      colour = NA
    ) +
    
    annotate(
      "text",
      x = 5.5,
      y = 1.05,
      label = "95% CI",
      hjust = 0,
      vjust = 0.5,
      size = 3.2,
      family = plot_font_family,
      colour = "#0F172A"
    ) +
    
    # keyword count title (closer to points)
    annotate(
      "text",
      x = 2.35,
      y = 0.82,
      label = "Keyword Count",
      hjust = 0,
      vjust = 0.5,
      size = 3.2,
      family = plot_font_family,
      colour = "#0F172A"
    ) +
    
    # keyword count points (more compact)
    annotate(
      "point",
      x = c(3.75, 4.2, 4.7)[seq_along(yearly_size_legend_breaks)],
      y = 0.82,
      colour = observed_colour,
      size = yearly_size_legend_sizes,
      alpha = yearly_point_alpha
    ) +
    
    annotate(
      "text",
      x = c(3.93, 4.38, 4.88)[seq_along(yearly_size_legend_breaks)],
      y = 0.82,
      label = yearly_size_legend_breaks,
      hjust = 0,
      vjust = 0.5,
      size = 3.05,
      family = plot_font_family,
      colour = "#0F172A"
    ) +
    
    # centered canvas
    coord_cartesian(
      xlim = c(0.8, 6.6),
      ylim = c(0.72, 1.18),
      clip = "off"
    ) +
    
    theme_void() +
    
    theme(
      plot.margin = margin(0, 0, 0, 0)
    )
}

# Original non-overlaid two-panel version retained for reference.
# policy_yearly_plot <- build_yearly_direction_plot(
#   relation_code = "Policy-first",
#   direction_label = "Policy to research",
#   observed_colour = yearly_observed_palette[["Policy-first"]],
#   trend_colour = yearly_trend_palette[["Policy-first"]],
#   panel_tag = "a",
#   show_x_title = FALSE
# )
# policy_yearly_legend <- build_yearly_direction_legend(
#   direction_label = "Policy to research",
#   observed_colour = yearly_observed_palette[["Policy-first"]],
#   trend_colour = yearly_trend_palette[["Policy-first"]]
# )
#
# research_yearly_plot <- build_yearly_direction_plot(
#   relation_code = "Research-first",
#   direction_label = "Research to policy",
#   observed_colour = yearly_observed_palette[["Research-first"]],
#   trend_colour = yearly_trend_palette[["Research-first"]],
#   panel_tag = "b",
#   show_x_title = TRUE
# )
# research_yearly_legend <- build_yearly_direction_legend(
#   direction_label = "Research to policy",
#   observed_colour = yearly_observed_palette[["Research-first"]],
#   trend_colour = yearly_trend_palette[["Research-first"]]
# )
#
# policy_yearly_panel <- patchwork::wrap_plots(
#   policy_yearly_plot,
#   policy_yearly_legend,
#   ncol = 1,
#   heights = c(1, 0.18)
# )
#
# research_yearly_panel <- patchwork::wrap_plots(
#   research_yearly_plot,
#   research_yearly_legend,
#   ncol = 1,
#   heights = c(1, 0.18)
# )
#
# yearly_conversion_plot <- patchwork::wrap_plots(
#   policy_yearly_panel,
#   research_yearly_panel,
#   ncol = 1,
#   heights = c(1, 1)
# )

yearly_direction_observed_palette <- c(
  "Policy to research" = yearly_observed_palette[["Policy-first"]],
  "Research to policy" = yearly_observed_palette[["Research-first"]]
)

yearly_direction_trend_palette <- c(
  "Policy to research" = yearly_trend_palette[["Policy-first"]],
  "Research to policy" = yearly_trend_palette[["Research-first"]]
)

yearly_conversion_overlay_plot <- ggplot() +
  geom_ribbon(
    data = yearly_trend_data,
    aes(
      x = first_year,
      ymin = conf_low,
      ymax = conf_high,
      fill = direction_label,
      group = direction_label
    ),
    alpha = yearly_trend_ci_alpha,
    colour = NA,
    show.legend = FALSE
  ) +
  geom_line(
    data = yearly_trend_data,
    aes(x = first_year, y = fit, colour = direction_label, group = direction_label),
    linewidth = 1.15,
    alpha = 0.94,
    show.legend = FALSE
  ) +
  geom_point(
    data = yearly_conversion_observed_data,
    aes(
      x = first_year,
      y = mean_conversion_years,
      fill = direction_label,
      colour = direction_label,
      size = point_size
    ),
    shape = 21,
    stroke = 0.45,
    alpha = yearly_point_alpha,
    show.legend = FALSE
  ) +
  scale_colour_manual(values = yearly_direction_trend_palette) +
  scale_fill_manual(values = yearly_direction_observed_palette) +
  scale_size_identity() +
  scale_x_continuous(
    breaks = conversion_year_range,
    limits = c(min(conversion_year_range) - 0.5, max(conversion_year_range) + 0.5),
    expand = c(0, 0)
  ) +
  scale_y_continuous(
    breaks = yearly_y_breaks,
    limits = c(0, yearly_y_max),
    labels = label_number(accuracy = 1),
    expand = expansion(mult = c(0, 0.04))
  ) +
  labs(
    x = "First Appearance Year",
    y = "Average Years",
    tag = "d"
  ) +
  theme_keyword_map(11) +
  theme(
    panel.grid.major.y = element_line(colour = alpha("#94A3B8", 0.24), linewidth = 0.35),
    axis.text = element_text(size = figure_axis_text_size, colour = "#243447"),
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = figure_axis_text_size),
    axis.title = element_text(size = figure_axis_title_size, face = "bold", colour = "#16203A"),
    legend.position = "none",
    plot.tag = element_text(face = "bold", size = 13, colour = "#0F172A"),
    plot.tag.position = c(0, 1),
    plot.margin = margin(8, 16, 6, 28)
  ) +
  coord_cartesian(clip = "off")

build_yearly_overlay_legend <- function() {
  key_x <- c(1.25, 3.25)
  key_labels <- c("Policy to research", "Research to policy")
  count_x <- c(6.5, 7.1, 7.7)[seq_along(yearly_size_legend_breaks)]

  ggplot() +
    annotate(
      "rect",
      xmin = key_x,
      xmax = key_x + 0.48,
      ymin = 0.96,
      ymax = 1.14,
      fill = alpha(yearly_direction_observed_palette[key_labels], yearly_ci_legend_alpha),
      colour = NA
    ) +
    annotate(
      "segment",
      x = key_x,
      xend = key_x + 0.48,
      y = 1.05,
      yend = 1.05,
      colour = yearly_direction_trend_palette[key_labels],
      linewidth = 1.15,
      lineend = "butt"
    ) +
    annotate(
      "point",
      x = key_x + 0.24,
      y = 1.05,
      fill = yearly_direction_observed_palette[key_labels],
      colour = yearly_direction_trend_palette[key_labels],
      shape = 21,
      stroke = 0.45,
      size = 3.0,
      alpha = yearly_point_alpha
    ) +
    annotate(
      "text",
      x = key_x + 0.62,
      y = 1.05,
      label = key_labels,
      hjust = 0,
      vjust = 0.5,
      size = figure_legend_text_geom_size,
      family = plot_font_family,
      colour = "#0F172A"
    ) +
    annotate(
      "text",
      x = 5.25,
      y = 1.05,
      label = "Keyword count",
      hjust = 0,
      vjust = 0.5,
      size = figure_legend_title_geom_size,
      family = plot_font_family,
      fontface = "bold",
      colour = "#0F172A"
    ) +
    annotate(
      "point",
      x = count_x,
      y = 1.05,
      fill = "#FFFFFF",
      colour = "#1E293B",
      shape = 21,
      stroke = 0.45,
      size = yearly_size_legend_sizes,
      alpha = yearly_point_alpha
    ) +
    annotate(
      "text",
      x = count_x + 0.22,
      y = 1.05,
      label = yearly_size_legend_breaks,
      hjust = 0,
      vjust = 0.5,
      size = figure_legend_text_geom_size,
      family = plot_font_family,
      colour = "#0F172A"
    ) +
    coord_cartesian(
      xlim = c(0.55, 8.92),
      ylim = c(0.82, 1.24),
      clip = "off"
    ) +
    theme_void() +
    theme(plot.margin = margin(0, 0, 0, 0))
}

yearly_conversion_legend <- build_yearly_overlay_legend()

yearly_conversion_plot <- patchwork::wrap_plots(
  yearly_conversion_overlay_plot,
  yearly_conversion_legend,
  ncol = 1,
  heights = c(1, 0.18)
)

yearly_conversion_plot_for_panel <- patchwork::wrap_plots(
  yearly_conversion_overlay_plot +
    theme(plot.margin = margin(4, 14, 2, 24)),
  yearly_conversion_legend,
  ncol = 1,
  heights = c(1, 0.12)
)

cohort_horizon_years <- 5L
max_complete_cohort_year <- year_max - cohort_horizon_years
cohort_year_range <- year_min:max_complete_cohort_year

cohort_conversion_source_data <- bind_rows(
  evidence |>
    filter(!is.na(policy_first_year)) |>
    transmute(
      theme_key,
      theme_en,
      direction_id = "P2R_5yr",
      relation_type = "Policy-first",
      direction_label = "Policy to research",
      first_appearance_domain = "Policy",
      conversion_target_domain = "Research",
      source_year = policy_first_year,
      target_year = research_first_year
    ),
  evidence |>
    filter(!is.na(research_first_year)) |>
    transmute(
      theme_key,
      theme_en,
      direction_id = "R2P_5yr",
      relation_type = "Research-first",
      direction_label = "Research to policy",
      first_appearance_domain = "Research",
      conversion_target_domain = "Policy",
      source_year = research_first_year,
      target_year = policy_first_year
    )
) |>
  filter(
    source_year %in% cohort_year_range,
    is.na(target_year) | target_year > source_year
  ) |>
  mutate(
    conversion_lag_years = target_year - source_year,
    converted_within_5yr = !is.na(target_year) & conversion_lag_years <= cohort_horizon_years
  )

cohort_conversion_annual_summary <- cohort_conversion_source_data |>
  group_by(direction_id, relation_type, direction_label, first_appearance_domain, conversion_target_domain, source_year) |>
  summarise(
    cohort_keyword_count = n(),
    converted_5yr_count = sum(converted_within_5yr, na.rm = TRUE),
    conversion_probability_5yr = converted_5yr_count / cohort_keyword_count,
    .groups = "drop"
  ) |>
  arrange(direction_id, source_year)

cohort_conversion_palette <- c(
  "Policy to research" = yearly_observed_palette[["Policy-first"]],
  "Research to policy" = yearly_observed_palette[["Research-first"]]
)

cohort_conversion_y_max <- max(cohort_conversion_annual_summary$conversion_probability_5yr, na.rm = TRUE)
cohort_conversion_y_max <- if_else(is.finite(cohort_conversion_y_max), cohort_conversion_y_max, 1)
cohort_conversion_y_max <- min(1, max(0.4, ceiling((cohort_conversion_y_max + 0.05) * 10) / 10))

cohort_conversion_plot <- ggplot(
  cohort_conversion_annual_summary,
  aes(
    x = source_year,
    y = conversion_probability_5yr,
    colour = direction_label,
    group = direction_label
  )
) +
  geom_line(linewidth = 1.05, alpha = 0.86) +
  geom_point(
    aes(size = cohort_keyword_count),
    alpha = 0.96
  ) +
  scale_colour_manual(values = cohort_conversion_palette, name = NULL) +
  scale_size_area(
    max_size = 7,
    breaks = pretty_breaks(n = 3),
    name = "Cohort keywords"
  ) +
  scale_x_continuous(
    breaks = seq(min(cohort_year_range), max(cohort_year_range), by = 2),
    expand = expansion(mult = c(0.02, 0.06))
  ) +
  scale_y_continuous(
    labels = label_percent(accuracy = 1),
    limits = c(0, cohort_conversion_y_max),
    breaks = pretty_breaks(n = 5),
    expand = expansion(mult = c(0, 0.03))
  ) +
  labs(
    x = "Source cohort year",
    y = "5-year conversion probability"
  ) +
  guides(
    colour = guide_legend(order = 1, nrow = 1, byrow = TRUE),
    size = guide_legend(order = 2, nrow = 1)
  ) +
  theme_keyword_map(11) +
  theme(
    panel.grid.major.y = element_line(colour = alpha("#94A3B8", 0.24), linewidth = 0.35),
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8.5),
    legend.box = "horizontal",
    plot.margin = margin(8, 18, 8, 20)
  )

p2r_annual_summary <- cohort_conversion_annual_summary |>
  filter(direction_id == "P2R_5yr") |>
  transmute(
    source_year,
    p2r_cohort_count = cohort_keyword_count,
    p2r_converted_5yr_count = converted_5yr_count,
    p2r_5yr = conversion_probability_5yr
  )

r2p_annual_summary <- cohort_conversion_annual_summary |>
  filter(direction_id == "R2P_5yr") |>
  transmute(
    source_year,
    r2p_cohort_count = cohort_keyword_count,
    r2p_converted_5yr_count = converted_5yr_count,
    r2p_5yr = conversion_probability_5yr
  )

directionality_index_summary <- full_join(
  p2r_annual_summary,
  r2p_annual_summary,
  by = "source_year"
) |>
  mutate(
    total_cohort_count = coalesce(p2r_cohort_count, 0L) + coalesce(r2p_cohort_count, 0L),
    bidirectional_uptake_5yr = rowMeans(cbind(p2r_5yr, r2p_5yr), na.rm = FALSE),
    policy_leading_directionality_5yr = p2r_5yr - r2p_5yr,
    conditional_directionality_5yr = if_else(
      !is.na(p2r_5yr) & !is.na(r2p_5yr) & (p2r_5yr + r2p_5yr) > 0,
      (p2r_5yr - r2p_5yr) / (p2r_5yr + r2p_5yr),
      NA_real_
    )
  ) |>
  arrange(source_year)

directionality_plot_data <- directionality_index_summary |>
  filter(!is.na(p2r_5yr), !is.na(r2p_5yr))

directionality_label_data <- directionality_plot_data |>
  arrange(source_year)

directionality_axis_max <- max(c(directionality_plot_data$p2r_5yr, directionality_plot_data$r2p_5yr), na.rm = TRUE)
directionality_axis_max <- if_else(is.finite(directionality_axis_max), directionality_axis_max, 1)
directionality_axis_max <- min(1, max(0.25, ceiling((directionality_axis_max + 0.04) * 20) / 20))

directionality_state_space_plot <- ggplot(
  directionality_plot_data,
  aes(x = r2p_5yr, y = p2r_5yr)
) +
  annotate(
    "rect",
    xmin = 0,
    xmax = directionality_axis_max,
    ymin = 0,
    ymax = directionality_axis_max,
    fill = "#F8FAFC",
    colour = NA
  ) +
  geom_abline(
    slope = 1,
    intercept = 0,
    colour = alpha("#475569", 0.55),
    linewidth = 0.65,
    linetype = "22"
  ) +
  geom_point(
    aes(size = total_cohort_count, fill = source_year),
    shape = 21,
    colour = "#0F172A",
    stroke = 0.6,
    alpha = 0.96
  ) +
  geom_text_repel(
    data = directionality_label_data,
    aes(label = source_year),
    seed = 42,
    size = 3.25,
    family = plot_font_family,
    colour = "#0F172A",
    box.padding = 0.18,
    point.padding = 0.11,
    min.segment.length = 0,
    segment.size = 0.14,
    segment.alpha = 0.34,
    max.overlaps = Inf
  ) +
  annotate(
    "text",
    x = directionality_axis_max * 0.23,
    y = directionality_axis_max * 0.88,
    label = "Higher policy-to-research\n5-year uptake",
    hjust = 0,
    size = 4.05,
    family = plot_font_family,
    colour = "#8A6410"
  ) +
  annotate(
    "text",
    x = directionality_axis_max * 0.58,
    y = directionality_axis_max * 0.12,
    label = "Higher research-to-policy\n5-year uptake",
    hjust = 0,
    size = 4.05,
    family = plot_font_family,
    colour = "#3D65A8"
  ) +
  scale_fill_gradientn(
    colours = c("#E7C867", "#A8C3E8", "#4E74B8"),
    breaks = unique(c(seq(year_min, max_complete_cohort_year, by = 5), max_complete_cohort_year)),
    name = "Cohort year",
    guide = guide_colourbar(
      order = 1,
      title.position = "top",
      barwidth = grid::unit(5.2, "cm"),
      barheight = grid::unit(0.3, "cm")
    )
  ) +
  scale_size_area(
    max_size = 10.5,
    breaks = pretty_breaks(n = 3),
    name = "Cohort keywords"
  ) +
  scale_x_continuous(
    labels = label_percent(accuracy = 1),
    limits = c(0, directionality_axis_max),
    breaks = pretty_breaks(n = 5),
    expand = c(0, 0)
  ) +
  scale_y_continuous(
    labels = label_percent(accuracy = 1),
    limits = c(0, directionality_axis_max),
    breaks = pretty_breaks(n = 5),
    expand = c(0, 0)
  ) +
  labs(
    x = "Research-to-policy 5-year conversion probability",
    y = "Policy-to-research 5-year conversion probability",
    tag = "a"
  ) +
  guides(
    size = guide_legend(order = 2, nrow = 1)
  ) +
  theme_keyword_map(13) +
  theme(
    panel.grid.major.y = element_line(colour = alpha("#94A3B8", 0.22), linewidth = 0.35),
    axis.text = element_text(size = figure_axis_text_size, colour = "#243447"),
    axis.title = element_text(size = figure_axis_title_size, face = "bold", colour = "#16203A"),
    legend.box = "horizontal",
    legend.text = element_text(size = figure_legend_text_size),
    legend.title = element_text(size = figure_legend_title_size, face = "bold"),
    plot.tag = element_text(face = "bold", size = 16, colour = "#0F172A"),
    plot.tag.position = c(0, 1),
    plot.margin = margin(8, 18, 8, 20)
  ) +
  coord_equal(clip = "off")

directionality_metric_plot_data <- directionality_plot_data |>
  select(
    source_year,
    total_cohort_count,
    policy_leading_directionality_5yr,
    conditional_directionality_5yr
  ) |>
  pivot_longer(
    cols = c(policy_leading_directionality_5yr, conditional_directionality_5yr),
    names_to = "directionality_metric",
    values_to = "directionality_value"
  ) |>
  mutate(
    directionality_metric = recode(
      directionality_metric,
      "policy_leading_directionality_5yr" = "P2R-R2P 5-year uptake gap",
      "conditional_directionality_5yr" = "Conditional 5-year uptake gap"
    ),
    directionality_metric = factor(
      directionality_metric,
      levels = c("P2R-R2P 5-year uptake gap", "Conditional 5-year uptake gap")
    ),
    direction_class = case_when(
      directionality_value > 0 ~ "Policy-leading",
      directionality_value < 0 ~ "Research-leading",
      TRUE ~ "Balanced"
    ),
    direction_class = factor(
      direction_class,
      levels = c("Policy-leading", "Research-leading", "Balanced")
    )
  ) |>
  filter(!is.na(directionality_value))

directionality_leading_palette <- c(
  "Policy-leading" = cohort_conversion_palette[["Policy to research"]],
  "Research-leading" = cohort_conversion_palette[["Research to policy"]],
  "Balanced" = "#94A3B8"
)

directionality_indices_plot <- ggplot(
  directionality_metric_plot_data,
  aes(x = source_year)
) +
  geom_hline(yintercept = 0, colour = alpha("#334155", 0.58), linewidth = 0.55) +
  geom_segment(
    aes(
      xend = source_year,
      y = 0,
      yend = directionality_value,
      colour = direction_class,
      linetype = directionality_metric
    ),
    linewidth = 0.72,
    alpha = 0.78,
    lineend = "round"
  ) +
  geom_point(
    aes(
      y = directionality_value,
      colour = direction_class,
      shape = directionality_metric
    ),
    size = 6.4,
    stroke = 1.45,
    alpha = 0.96
  ) +
  scale_colour_manual(
    values = directionality_leading_palette,
    labels = c(
      "Policy-leading" = "Net P2R",
      "Research-leading" = "Net R2P",
      "Balanced" = "Balanced"
    ),
    name = "Direction",
    drop = FALSE
  ) +
  scale_linetype_manual(
    values = c(
      "P2R-R2P 5-year uptake gap" = "solid",
      "Conditional 5-year uptake gap" = "22"
    ),
    labels = c(
      "P2R-R2P 5-year uptake gap" = "P2R-R2P gap",
      "Conditional 5-year uptake gap" = "Conditional gap"
    ),
    name = "Index"
  ) +
  scale_shape_manual(
    values = c(
      "P2R-R2P 5-year uptake gap" = 16,
      "Conditional 5-year uptake gap" = 1
    ),
    labels = c(
      "P2R-R2P 5-year uptake gap" = "P2R-R2P gap",
      "Conditional 5-year uptake gap" = "Conditional gap"
    ),
    name = "Index"
  ) +
  scale_x_continuous(
    breaks = conversion_year_range,
    labels = conversion_year_range,
    limits = c(min(conversion_year_range) - 0.5, max(conversion_year_range) + 0.5),
    expand = c(0, 0)
  ) +
  scale_y_continuous(
    labels = label_percent(accuracy = 1),
    limits = c(-1, 1),
    breaks = seq(-1, 1, by = 0.5),
    position = "right",
    expand = expansion(mult = c(0.03, 0.03))
  ) +
  labs(
    x = "Source cohort year",
    y = "Directionality index value"
  ) +
  guides(
    colour = guide_legend(
      order = 1,
      nrow = 1,
      byrow = TRUE,
      title.position = "left",
      override.aes = list(size = 6.4)
    ),
    linetype = guide_legend(
      order = 2,
      nrow = 1,
      byrow = TRUE,
      title.position = "left"
    ),
    shape = guide_legend(
      order = 2,
      nrow = 1,
      byrow = TRUE,
      title.position = "left",
      override.aes = list(size = 6.4)
    )
  ) +
  theme_keyword_map(12) +
  theme(
    panel.grid.major.y = element_line(colour = alpha("#94A3B8", 0.18), linewidth = 0.3),
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 0.5, size = figure_axis_text_size),
    axis.text.y.right = element_text(angle = 90, vjust = 0.5, hjust = 0.5, size = figure_axis_text_size),
    axis.title.x = element_text(angle = 180, size = figure_axis_title_size, face = "bold", colour = "#16203A"),
    axis.title.y.right = element_text(angle = 90, size = figure_axis_title_size, face = "bold", colour = "#16203A"),
    legend.position = "inside",
    legend.position.inside = c(0.5, 0.035),
    legend.justification = c(0.5, 0),
    legend.box = "vertical",
    legend.background = element_rect(fill = alpha("#FFFFFF", 0.86), colour = NA),
    legend.key = element_rect(fill = alpha("#FFFFFF", 0), colour = NA),
    legend.key.width = grid::unit(0.52, "cm"),
    legend.key.height = grid::unit(0.46, "cm"),
    legend.spacing.y = grid::unit(0.05, "cm"),
    legend.margin = margin(3, 5, 3, 5),
    legend.box.margin = margin(0, 0, 0, 0),
    legend.text = element_text(size = figure_legend_text_size),
    legend.title = element_text(size = figure_legend_title_size, face = "bold"),
    plot.margin = margin(8, 14, 8, 14)
  )

policy_keyword_yearly_appearance <- load_policy_keyword_yearly_appearance(policy_keyword_matches_path)
research_keyword_yearly_appearance <- load_research_keyword_yearly_appearance(
  research_cluster_dir,
  research_keyword_dictionary_path,
  research_keyword_yearly_matrix_path
)

annual_keyword_appearance <- bind_rows(
  policy_keyword_yearly_appearance,
  research_keyword_yearly_appearance
) |>
  group_by(theme_key, target_domain, year) |>
  summarise(
    appearance_count = max(appearance_count, na.rm = TRUE),
    present = any(present, na.rm = TRUE),
    .groups = "drop"
  )

annual_keyword_coverage <- annual_keyword_appearance |>
  distinct(theme_key, target_domain) |>
  mutate(annual_data_available = TRUE)

post_conversion_follow_up_years <- 1:10

converted_keyword_data <- shared_themes |>
  filter(relation_type %in% c("Policy-first", "Research-first")) |>
  transmute(
    theme_key,
    theme_en,
    relation_type,
    direction_label = recode(
      relation_type,
      "Policy-first" = "Policy to research",
      "Research-first" = "Research to policy"
    ),
    conversion_target_domain = recode(
      relation_type,
      "Policy-first" = "Research",
      "Research-first" = "Policy"
    ),
    conversion_year = if_else(relation_type == "Policy-first", research_first_year, policy_first_year),
    conversion_lag_years = if_else(relation_type == "Policy-first", policy_to_research_years, research_to_policy_years)
  ) |>
  left_join(
    annual_keyword_coverage,
    by = c("theme_key", "conversion_target_domain" = "target_domain")
  ) |>
  mutate(annual_data_available = replace_na(annual_data_available, FALSE))

post_conversion_persistence_coverage <- converted_keyword_data |>
  group_by(relation_type, direction_label, conversion_target_domain) |>
  summarise(
    converted_keyword_count = n(),
    annual_coverage_count = sum(annual_data_available),
    annual_coverage_share = annual_coverage_count / converted_keyword_count,
    .groups = "drop"
  )

post_conversion_persistence_detail <- converted_keyword_data |>
  filter(annual_data_available) |>
  tidyr::expand_grid(years_since_conversion = post_conversion_follow_up_years) |>
  mutate(
    observation_year = conversion_year + years_since_conversion,
    observable = observation_year <= year_max
  ) |>
  filter(observable) |>
  left_join(
    annual_keyword_appearance,
    by = c(
      "theme_key",
      "conversion_target_domain" = "target_domain",
      "observation_year" = "year"
    )
  ) |>
  mutate(
    appearance_count = replace_na(appearance_count, 0),
    retained = replace_na(present, FALSE)
  ) |>
  select(
    theme_key,
    theme_en,
    relation_type,
    direction_label,
    conversion_target_domain,
    conversion_year,
    conversion_lag_years,
    years_since_conversion,
    observation_year,
    appearance_count,
    retained
  )

post_conversion_persistence_summary <- post_conversion_persistence_detail |>
  group_by(relation_type, direction_label, conversion_target_domain, years_since_conversion) |>
  summarise(
    observable_keyword_count = n_distinct(theme_key),
    retained_keyword_count = sum(retained, na.rm = TRUE),
    retention_rate = retained_keyword_count / observable_keyword_count,
    .groups = "drop"
  ) |>
  arrange(relation_type, years_since_conversion)

post_conversion_persistence_plot <- ggplot(
  post_conversion_persistence_summary,
  aes(
    x = years_since_conversion,
    y = retention_rate,
    colour = direction_label,
    group = direction_label
  )
) +
  geom_path(linewidth = 1.05, alpha = 0.88) +
  geom_point(
    aes(size = observable_keyword_count),
    alpha = 0.96
  ) +
  scale_colour_manual(values = cohort_conversion_palette, name = "Direction") +
  scale_size_area(
    max_size = 6.5,
    breaks = pretty_breaks(n = 3),
    name = "Observable keywords"
  ) +
  scale_x_continuous(
    breaks = post_conversion_follow_up_years,
    labels = post_conversion_follow_up_years,
    limits = c(min(post_conversion_follow_up_years) - 0.25, max(post_conversion_follow_up_years) + 0.25),
    expand = c(0, 0)
  ) +
  scale_y_continuous(
    labels = label_percent(accuracy = 1),
    limits = c(0, 1),
    breaks = seq(0, 1, by = 0.2),
    position = "right",
    expand = expansion(mult = c(0, 0.03))
  ) +
  labs(
    x = "Years after first uptake by the receiving side",
    y = "Receiving-side retention rate"
  ) +
  guides(
    colour = guide_legend(
      order = 1,
      nrow = 1,
      byrow = TRUE,
      title.position = "left"
    ),
    size = guide_legend(
      order = 2,
      nrow = 1,
      byrow = TRUE,
      title.position = "left"
    )
  ) +
  theme_keyword_map(12) +
  theme(
    panel.grid.major.y = element_line(colour = alpha("#94A3B8", 0.24), linewidth = 0.35),
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 0.5, size = figure_axis_text_size),
    axis.text.y.right = element_text(angle = 90, vjust = 0.5, hjust = 0.5, size = figure_axis_text_size),
    axis.title.x = element_text(angle = 180, size = figure_axis_title_size, face = "bold", colour = "#16203A"),
    axis.title.y.right = element_text(angle = 90, size = figure_axis_title_size, face = "bold", colour = "#16203A"),
    legend.position = "inside",
    legend.position.inside = c(0.985, 0.035),
    legend.justification = c(1, 0),
    legend.box = "vertical",
    legend.background = element_rect(fill = alpha("#FFFFFF", 0.86), colour = NA),
    legend.key = element_rect(fill = alpha("#FFFFFF", 0), colour = NA),
    legend.key.width = grid::unit(0.52, "cm"),
    legend.key.height = grid::unit(0.46, "cm"),
    legend.spacing.y = grid::unit(0.05, "cm"),
    legend.margin = margin(3, 5, 3, 5),
    legend.box.margin = margin(0, 0, 0, 0),
    legend.text = element_text(size = figure_legend_text_size),
    legend.title = element_text(size = figure_legend_title_size, face = "bold"),
    plot.margin = margin(8, 14, 8, 14)
  )

directionality_state_space_plot_for_panel <- directionality_state_space_plot +
  guides(
    fill = guide_colourbar(
      order = 1,
      title.position = "top",
      barwidth = grid::unit(5.1, "cm"),
      barheight = grid::unit(0.26, "cm"),
      theme = theme(legend.text = element_text(size = figure_legend_text_size - 1.8))
    ),
    size = guide_legend(order = 2, nrow = 1)
  ) +
  theme(
    legend.margin = margin(0, 0, 0, 0),
    legend.box.margin = margin(-4, 0, 0, 0),
    legend.spacing.x = grid::unit(0.18, "cm"),
    legend.spacing.y = grid::unit(0, "cm"),
    plot.margin = margin(0, 2, 0, 2)
  )

directionality_metric_panel_data <- directionality_metric_plot_data |>
  mutate(
    direction_class_panel = recode(
      direction_class,
      "Policy-leading" = "Net P2R",
      "Research-leading" = "Net R2P",
      "Balanced" = "Balanced"
    ),
    directionality_metric_panel = recode(
      directionality_metric,
      "P2R-R2P 5-year uptake gap" = "P2R-R2P gap",
      "Conditional 5-year uptake gap" = "Conditional gap"
    )
  )

directionality_panel_palette <- c(
  "Net P2R" = directionality_leading_palette[["Policy-leading"]],
  "Net R2P" = directionality_leading_palette[["Research-leading"]],
  "Balanced" = directionality_leading_palette[["Balanced"]]
)

directionality_indices_plot_for_panel <- ggplot(
  directionality_metric_panel_data,
  aes(y = source_year)
) +
  geom_vline(xintercept = 0, colour = alpha("#334155", 0.58), linewidth = 0.55) +
  geom_segment(
    aes(
      x = 0,
      xend = directionality_value,
      yend = source_year,
      colour = direction_class_panel,
      linetype = directionality_metric_panel
    ),
    linewidth = 0.72,
    alpha = 0.78,
    lineend = "round"
  ) +
  geom_point(
    aes(
      x = directionality_value,
      colour = direction_class_panel,
      shape = directionality_metric_panel
    ),
    size = 6.4,
    stroke = 1.45,
    alpha = 0.96
  ) +
  scale_colour_manual(values = directionality_panel_palette, name = "Direction", drop = FALSE) +
  scale_linetype_manual(
    values = c(
      "P2R-R2P gap" = "solid",
      "Conditional gap" = "22"
    ),
    name = "Index"
  ) +
  scale_shape_manual(
    values = c(
      "P2R-R2P gap" = 16,
      "Conditional gap" = 1
    ),
    name = "Index"
  ) +
  scale_x_continuous(
    labels = label_percent(accuracy = 1),
    limits = c(-1, 1),
    breaks = seq(-1, 1, by = 0.5),
    expand = expansion(mult = c(0.03, 0.03))
  ) +
  scale_y_reverse(
    breaks = conversion_year_range,
    labels = conversion_year_range,
    limits = c(max(conversion_year_range) + 0.5, min(conversion_year_range) - 0.5),
    expand = c(0, 0)
  ) +
  labs(
    tag = "b",
    x = "Directionality index value",
    y = "Source cohort year"
  ) +
  guides(
    colour = guide_legend(
      order = 1,
      nrow = 1,
      byrow = TRUE,
      title.position = "left",
      override.aes = list(size = 6.4)
    ),
    linetype = guide_legend(
      order = 2,
      nrow = 1,
      byrow = TRUE,
      title.position = "left"
    ),
    shape = guide_legend(
      order = 2,
      nrow = 1,
      byrow = TRUE,
      title.position = "left",
      override.aes = list(size = 6.4)
    )
  ) +
  theme_keyword_map(12) +
  theme(
    panel.grid.major.x = element_line(colour = alpha("#94A3B8", 0.18), linewidth = 0.3),
    panel.grid.major.y = element_line(colour = alpha("#94A3B8", 0.18), linewidth = 0.3),
    axis.text = element_text(size = figure_axis_text_size, colour = "#243447"),
    axis.title = element_text(size = figure_axis_title_size, face = "bold", colour = "#16203A"),
    legend.position = "inside",
    legend.position.inside = c(0.5, 0.035),
    legend.justification = c(0.5, 0),
    legend.box.just = "center",
    legend.box = "vertical",
    legend.background = element_rect(fill = alpha("#FFFFFF", 0.86), colour = NA),
    legend.key = element_rect(fill = alpha("#FFFFFF", 0), colour = NA),
    legend.key.width = grid::unit(0.48, "cm"),
    legend.key.height = grid::unit(0.42, "cm"),
    legend.spacing.y = grid::unit(0.04, "cm"),
    legend.margin = margin(3, 5, 3, 5),
    legend.box.margin = margin(0, 0, 0, 0),
    legend.text = element_text(size = figure_legend_text_size),
    legend.title = element_text(size = figure_legend_title_size, face = "bold"),
    plot.tag = element_text(size = 16, face = "bold", colour = "#0F172A"),
    plot.tag.position = c(0.01, 0.99),
    plot.margin = margin(4, 8, 4, 8)
  )

post_conversion_persistence_plot_for_panel <- ggplot(
  post_conversion_persistence_summary,
  aes(
    x = retention_rate,
    y = years_since_conversion,
    colour = direction_label,
    group = direction_label
  )
) +
  geom_path(linewidth = 1.05, alpha = 0.88) +
  geom_point(
    aes(size = observable_keyword_count),
    alpha = 0.96
  ) +
  scale_colour_manual(
    values = cohort_conversion_palette,
    labels = c(
      "Policy to research" = "P2R",
      "Research to policy" = "R2P"
    ),
    name = "Direction"
  ) +
  scale_size_area(
    max_size = 6.5,
    breaks = pretty_breaks(n = 3),
    name = "Keywords"
  ) +
  scale_x_continuous(
    labels = label_percent(accuracy = 1),
    limits = c(0, 1),
    breaks = seq(0, 1, by = 0.2),
    expand = expansion(mult = c(0, 0.03))
  ) +
  scale_y_reverse(
    breaks = post_conversion_follow_up_years,
    labels = post_conversion_follow_up_years,
    limits = c(max(post_conversion_follow_up_years) + 1.25, min(post_conversion_follow_up_years) - 0.25),
    expand = c(0, 0)
  ) +
  labs(
    tag = "c",
    x = "Receiving-side retention rate",
    y = "Years after first uptake by the receiving side"
  ) +
  guides(
    colour = guide_legend(
      order = 1,
      nrow = 1,
      byrow = TRUE,
      title.position = "left"
    ),
    size = guide_legend(
      order = 2,
      nrow = 1,
      byrow = TRUE,
      title.position = "left"
    )
  ) +
  theme_keyword_map(12) +
  theme(
    panel.grid.major.x = element_line(colour = alpha("#94A3B8", 0.24), linewidth = 0.35),
    panel.grid.major.y = element_line(colour = alpha("#94A3B8", 0.24), linewidth = 0.35),
    axis.text = element_text(size = figure_axis_text_size, colour = "#243447"),
    axis.title = element_text(size = figure_axis_title_size, face = "bold", colour = "#16203A"),
    legend.position = "inside",
    legend.position.inside = c(0.5, 0.035),
    legend.justification = c(0.5, 0),
    legend.box.just = "center",
    legend.box = "vertical",
    legend.background = element_rect(fill = alpha("#FFFFFF", 0.86), colour = NA),
    legend.key = element_rect(fill = alpha("#FFFFFF", 0), colour = NA),
    legend.key.width = grid::unit(0.48, "cm"),
    legend.key.height = grid::unit(0.42, "cm"),
    legend.spacing.y = grid::unit(0.04, "cm"),
    legend.margin = margin(3, 5, 3, 5),
    legend.box.margin = margin(0, 0, 0, 0),
    legend.text = element_text(size = figure_legend_text_size),
    legend.title = element_text(size = figure_legend_title_size, face = "bold"),
    plot.tag = element_text(size = 16, face = "bold", colour = "#0F172A"),
    plot.tag.position = c(0.01, 0.99),
    plot.margin = margin(4, 8, 4, 8)
  )

directionality_interaction_four_panel_plot <- patchwork::wrap_plots(
  patchwork::wrap_plots(
    directionality_state_space_plot_for_panel,
    yearly_conversion_plot_for_panel,
    ncol = 1,
    heights = c(1.30, 1.06)
  ),
  directionality_indices_plot_for_panel,
  post_conversion_persistence_plot_for_panel,
  ncol = 3,
  widths = c(1.72, 0.62, 0.62)
)

inheritance_plot <- ggplot() +
  geom_hline(
    yintercept = band_separator_y,
    colour = alpha("#64748B", 0.45),
    linewidth = 0.55
  ) +
  geom_curve(
    data = directed_edges |> filter(relation_type == "Policy-first"),
    aes(x = x, y = y, xend = xend, yend = yend, colour = relation_type),
    curvature = 0.18,
    linewidth = 0.40,
    alpha = 0.58,
    arrow = grid::arrow(length = grid::unit(0.11, "inches"), type = "closed", angle = 12),
    lineend = "round"
  ) +
  geom_curve(
    data = directed_edges |> filter(relation_type == "Research-first"),
    aes(x = x, y = y, xend = xend, yend = yend, colour = relation_type),
    curvature = -0.18,
    linewidth = 0.40,
    alpha = 0.58,
    arrow = grid::arrow(length = grid::unit(0.11, "inches"), type = "closed", angle = 12),
    lineend = "round"
  ) +
  geom_segment(
    data = synchronous_edges,
    aes(x = policy_x, y = policy_y, xend = research_x, yend = research_y, colour = relation_type),
    linewidth = 0.42,
    alpha = 0.62,
    lineend = "round",
    arrow = grid::arrow(length = grid::unit(0.11, "inches"), type = "closed", ends = "both", angle = 12)
  ) +
  geom_point(
    data = nodes,
    aes(x = year, y = y, fill = relation_type),
    shape = 21,
    size = 2.15,
    stroke = 0,
    colour = "transparent",
    alpha = 0.98
  ) +
  geom_text_repel(
    data = nodes |> filter(zone == "Shared-policy"),
    aes(x = year, y = y, label = display_label),
    seed = 42,
    size = 2.0,
    lineheight = 0.84,
    colour = "#0F172A",
    direction = "both",
    force = 1.6,
    force_pull = 0.06,
    box.padding = 0.12,
    point.padding = 0.04,
    min.segment.length = 0,
    segment.size = 0.14,
    segment.alpha = 0.35,
    segment.color = "#94A3B8",
    max.overlaps = Inf,
    max.time = 3,
    max.iter = 20000,
    ylim = policy_label_band
  ) +
  geom_text_repel(
    data = nodes |> filter(zone == "Shared-research"),
    aes(x = year, y = y, label = display_label),
    seed = 84,
    size = 2.0,
    lineheight = 0.84,
    colour = "#0F172A",
    direction = "both",
    force = 1.9,
    force_pull = 0.05,
    box.padding = 0.13,
    point.padding = 0.04,
    min.segment.length = 0,
    segment.size = 0.14,
    segment.alpha = 0.35,
    segment.color = "#94A3B8",
    max.overlaps = Inf,
    max.time = 4,
    max.iter = 25000,
    ylim = research_label_band
  ) +
  scale_fill_manual(
    values = relation_fill_palette,
    breaks = c("Policy-first", "Research-first", "Same-year"),
    labels = c(
      "Policy-first" = "Policy-led keywords",
      "Research-first" = "Research-led keywords",
      "Same-year" = "Synchronous keywords"
    ),
    name = NULL
  ) +
  scale_colour_manual(
    values = relation_arrow_palette,
    breaks = c("Policy-first", "Research-first", "Same-year"),
    labels = c(
      "Policy-first" = "Policy drives research",
      "Research-first" = "Research drives policy",
      "Same-year" = "Synchronous emergence"
    ),
    name = NULL
  ) +
  scale_x_continuous(
    breaks = year_min:year_max,
    limits = c(year_min - 0.5, year_max + 0.5),
    expand = c(0, 0),
    sec.axis = dup_axis(
      name = NULL,
      breaks = year_min:year_max,
      labels = year_min:year_max
    )
  ) +
  scale_y_continuous(
    breaks = zone_midpoints$y_mid,
    labels = zone_midpoints$zone_label,
    limits = c(min(zone_spec$y_min) - 0.05, max(zone_spec$y_max) + 0.32),
    expand = c(0, 0)
  ) +
  labs(
    title = NULL,
    subtitle = NULL,
    x = "Year",
    y = NULL
  ) +
  guides(
    fill = guide_legend(order = 1, nrow = 1, byrow = TRUE, title.position = "left", title.hjust = 0),
    colour = guide_legend(order = 2, nrow = 1, byrow = TRUE, title.position = "left", title.hjust = 0)
  ) +
  theme_keyword_map(11) +
  theme(
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8),
    axis.text.x.top = element_text(angle = 90, vjust = 0.5, hjust = 0, size = 8, colour = "#243447"),
    axis.text.y = element_text(face = "bold", size = 11, colour = "#16203A"),
    axis.ticks.x.top = element_blank(),
    legend.key.width = grid::unit(1.8, "lines"),
    legend.box = "vertical",
    legend.box.just = "center",
    legend.justification = "center",
    legend.title = element_blank(),
    plot.margin = margin(6, 16, 6, 22)
  ) +
  coord_cartesian(clip = "off")

plot_png_path <- file.path(interaction_dir, "policy_research_keyword_inheritance_timeline.png")
plot_pdf_path <- file.path(interaction_dir, "policy_research_keyword_inheritance_timeline.pdf")
nodes_csv_path <- file.path(interaction_dir, "policy_research_keyword_inheritance_nodes.csv")
edges_csv_path <- file.path(interaction_dir, "policy_research_keyword_inheritance_arrows.csv")
conversion_summary_csv_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_conversion_summary.csv"
)
yearly_conversion_plot_png_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_yearly_conversion_time.png"
)
yearly_conversion_plot_pdf_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_yearly_conversion_time.pdf"
)
yearly_conversion_summary_csv_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_yearly_conversion_summary.csv"
)
cohort_conversion_plot_png_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_cohort_conversion_probability.png"
)
cohort_conversion_plot_pdf_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_cohort_conversion_probability.pdf"
)
cohort_conversion_annual_csv_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_cohort_conversion_probability_annual.csv"
)
directionality_plot_png_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_directionality_state_space.png"
)
directionality_plot_pdf_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_directionality_state_space.pdf"
)
directionality_summary_csv_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_directionality_index_summary.csv"
)
directionality_indices_plot_png_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_directionality_indices.png"
)
directionality_indices_plot_pdf_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_directionality_indices.pdf"
)
post_conversion_persistence_plot_png_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_post_conversion_persistence.png"
)
post_conversion_persistence_plot_pdf_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_post_conversion_persistence.pdf"
)
directionality_interaction_four_panel_plot_png_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_directionality_interaction_four_panel.png"
)
directionality_interaction_four_panel_plot_pdf_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_directionality_interaction_four_panel.pdf"
)
post_conversion_persistence_summary_csv_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_post_conversion_persistence_summary.csv"
)
post_conversion_persistence_detail_csv_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_post_conversion_persistence_detail.csv"
)
post_conversion_persistence_coverage_csv_path <- file.path(
  interaction_dir,
  "policy_research_keyword_inheritance_post_conversion_persistence_coverage.csv"
)

save_plot_with_pdf(inheritance_plot, plot_png_path, width = 18, height = 21.5, dpi = 320)
save_plot_with_pdf(yearly_conversion_plot, yearly_conversion_plot_png_path, width = 13.5, height = 7.1, dpi = 320)
save_plot_with_pdf(cohort_conversion_plot, cohort_conversion_plot_png_path, width = 10.8, height = 6.4, dpi = 320)
save_plot_with_pdf(directionality_state_space_plot, directionality_plot_png_path, width = 9.2, height = 8.2, dpi = 320)
save_plot_with_pdf(directionality_indices_plot, directionality_indices_plot_png_path, width = 14.4, height = 4.5, dpi = 320)
save_plot_with_pdf(post_conversion_persistence_plot, post_conversion_persistence_plot_png_path, width = 14.4, height = 4.5, dpi = 320)
save_plot_with_pdf(directionality_interaction_four_panel_plot, directionality_interaction_four_panel_plot_png_path, width = 23.5, height = 15.8, dpi = 320)
write_csv(nodes, nodes_csv_path)
write_csv(shared_node_pairs, edges_csv_path)
write_csv(conversion_summary, conversion_summary_csv_path)
write_csv(yearly_conversion_plot_data, yearly_conversion_summary_csv_path)
write_csv(cohort_conversion_annual_summary, cohort_conversion_annual_csv_path)
write_csv(directionality_index_summary, directionality_summary_csv_path)
write_csv(post_conversion_persistence_summary, post_conversion_persistence_summary_csv_path)
write_csv(post_conversion_persistence_detail, post_conversion_persistence_detail_csv_path)
write_csv(post_conversion_persistence_coverage, post_conversion_persistence_coverage_csv_path)
unlink(file.path(interaction_dir, "policy_research_keyword_inheritance_cohort_conversion_probability_period.csv"))

message("Saved plot to: ", plot_png_path)
message("Saved plot to: ", plot_pdf_path)
message("Saved yearly conversion plot to: ", yearly_conversion_plot_png_path)
message("Saved yearly conversion plot to: ", yearly_conversion_plot_pdf_path)
message("Saved cohort conversion plot to: ", cohort_conversion_plot_png_path)
message("Saved cohort conversion plot to: ", cohort_conversion_plot_pdf_path)
message("Saved directionality state-space plot to: ", directionality_plot_png_path)
message("Saved directionality state-space plot to: ", directionality_plot_pdf_path)
message("Saved directionality indices plot to: ", directionality_indices_plot_png_path)
message("Saved directionality indices plot to: ", directionality_indices_plot_pdf_path)
message("Saved post-conversion persistence plot to: ", post_conversion_persistence_plot_png_path)
message("Saved post-conversion persistence plot to: ", post_conversion_persistence_plot_pdf_path)
message("Saved four-panel interaction plot to: ", directionality_interaction_four_panel_plot_png_path)
message("Saved four-panel interaction plot to: ", directionality_interaction_four_panel_plot_pdf_path)
message("Saved node table to: ", nodes_csv_path)
message("Saved arrow table to: ", edges_csv_path)
message("Saved conversion summary to: ", conversion_summary_csv_path)
message("Saved yearly conversion summary to: ", yearly_conversion_summary_csv_path)
message("Saved cohort conversion annual summary to: ", cohort_conversion_annual_csv_path)
message("Saved directionality index summary to: ", directionality_summary_csv_path)
message("Saved post-conversion persistence summary to: ", post_conversion_persistence_summary_csv_path)
message("Saved post-conversion persistence detail to: ", post_conversion_persistence_detail_csv_path)
message("Saved post-conversion persistence coverage to: ", post_conversion_persistence_coverage_csv_path)
message(
  sprintf(
    "Theme counts: %d shared themes, including %d policy-first, %d research-first, and %d same-year.",
    length(unique(shared_node_pairs$theme_en)),
    sum(shared_node_pairs$relation_type == "Policy-first"),
    sum(shared_node_pairs$relation_type == "Research-first"),
    sum(shared_node_pairs$relation_type == "Same-year")
  )
)
invisible(
  lapply(
    seq_len(nrow(conversion_summary)),
    function(i) {
      summary_row <- conversion_summary[i, ]
      mean_years <- summary_row$mean_conversion_years[[1]]
      keyword_count <- summary_row$keyword_count[[1]]
      summary_id <- summary_row$summary_id[[1]]

      message(
        sprintf(
          "Average conversion time for %s: %s years across %s keywords.",
          summary_id,
          ifelse(is.na(mean_years), "NA", mean_years),
          keyword_count
        )
      )
    }
  )
)
