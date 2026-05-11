#!/usr/bin/env Rscript

suppressPackageStartupMessages({
  library(readxl)
  library(readr)
  library(openxlsx)
  library(dplyr)
  library(tidyr)
  library(purrr)
  library(stringr)
  library(forcats)
  library(ggplot2)
  library(ggrepel)
  library(igraph)
  library(ggraph)
  library(scales)
  library(tibble)
  library(lubridate)
  library(patchwork)
  library(xml2)
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

project_root <- resolve_project_root()
policy_dir <- file.path(project_root, "policy text")
literature_dir <- file.path(project_root, "bibliometric analysis")
output_root <- file.path(project_root, "analysis_outputs", "biodiversity_policy_english")
policy_year_min <- 1990L
policy_year_max <- 2024L
policy_object_reference_checked_path <- file.path(output_root, "policy", "03_keyword_network", "gephi_nodes_checked.csv")
policy_keyword_nodes_generated_path <- file.path(output_root, "policy", "03_keyword_network", "gephi_nodes_generated.csv")
keyword_stoplist_path <- file.path(output_root, "policy", "03_keyword_network", "keyword_stoplist.csv")
research_keyword_kept_path <- file.path(output_root, "literature", "04_policy_research_interaction", "research_keywords_kept.csv")
policy_domain_output_dir <- file.path(output_root, "policy", "01_type_subject_domain")
legacy_policy_domain_output_dir <- file.path(output_root, "policy", "01_subject_object")

# Document-level policy domain classification: two-step workflow.
# Step 1: the pipeline writes an auto-classified list on every run.
# Step 2: the user manually corrects misclassifications and saves as _checked.csv.
# On subsequent runs the checked file overrides the auto-classification.
doc_class_generated_path <- file.path(policy_domain_output_dir, "document_domain_classification.csv")
doc_class_checked_path   <- file.path(policy_domain_output_dir, "document_domain_classification_checked.csv")
doc_class_checked_legacy_path <- file.path(legacy_policy_domain_output_dir, "document_object_classification_checked.csv")

message("Resolved project_root: ", project_root)
message("policy_dir: ", policy_dir)
message("literature_dir: ", literature_dir)
message("output_root: ", output_root)

if (!dir.exists(policy_dir)) {
  stop("Directory does not exist: ", policy_dir)
}
if (!dir.exists(literature_dir)) {
  stop("Directory does not exist: ", literature_dir)
}

if (!file.exists(policy_object_reference_checked_path)) {
  stop(
    "Required checked keyword node file does not exist: ", policy_object_reference_checked_path, "\n",
    "Policy keyword plots and downstream operations now require gephi_nodes_checked.csv; ",
    "gephi_nodes.csv is no longer used as an input."
  )
}

message("Loading authoritative policy keyword reference from: ", policy_object_reference_checked_path)
policy_object_reference_raw <- read_csv(policy_object_reference_checked_path, show_col_types = FALSE)

dir.create(output_root, recursive = TRUE, showWarnings = FALSE)

paths <- list(
  policy_subject = policy_domain_output_dir,
  policy_network = file.path(output_root, "policy", "02_intergovernmental_network"),
  keyword_network = file.path(output_root, "policy", "03_keyword_network"),
  policy_tools = file.path(output_root, "policy", "04_policy_tools"),
  policy_pmc = file.path(output_root, "policy", "05_pmc_evaluation"),
  literature_stage = file.path(output_root, "literature", "01_research_stages"),
  literature_burst = file.path(output_root, "literature", "02_burstness"),
  literature_timeline = file.path(output_root, "literature", "03_timezone"),
  interaction = file.path(output_root, "literature", "04_policy_research_interaction")
)
walk(paths, ~ dir.create(.x, recursive = TRUE, showWarnings = FALSE))

message("Loading data and building the unified corpus.")

# ----- Shared plotting utilities -------------------------------------------------
# These helpers keep the plot layout code reusable and easier to inspect.

resolve_plot_font_family <- function() {
  candidate_families <- c(
    "Helvetica",
    "Liberation Sans",
    "DejaVu Sans",
    "Noto Sans"
  )
  
  if (!requireNamespace("systemfonts", quietly = TRUE)) {
    return("sans")
  }
  
  available_families <- unique(systemfonts::system_fonts()$family)
  matched_family <- candidate_families[candidate_families %in% available_families]
  
  if (length(matched_family) == 0) {
    return("sans")
  }
  
  matched_family[1]
}

configure_plot_text_defaults <- function(font_family) {
  for (geom_name in c("text", "label", "text_repel", "label_repel")) {
    try(update_geom_defaults(geom_name, list(family = font_family)), silent = TRUE)
  }
}

plot_font_family <- resolve_plot_font_family()
configure_plot_text_defaults(plot_font_family)
message("Using plot font family: ", plot_font_family)

theme_policy <- function(base_size = 12) {
  theme_minimal(base_size = base_size, base_family = plot_font_family) +
    theme(
      plot.title = element_text(face = "bold", size = base_size + 3, colour = "#16203A"),
      plot.subtitle = element_text(size = base_size, colour = "#4F5D75"),
      plot.caption = element_text(size = base_size - 2, colour = "#6B7280"),
      panel.grid.minor = element_blank(),
      panel.grid.major.x = element_blank(),
      panel.grid.major.y = element_blank(),
      axis.title = element_text(face = "bold", colour = "#16203A"),
      axis.text = element_text(colour = "#283044"),
      legend.position = "bottom",
      legend.box = "vertical",
      legend.title = element_text(face = "bold"),
      plot.background = element_rect(fill = "#FFFFFF", colour = NA),
      panel.background = element_rect(fill = "#FFFFFF", colour = NA)
    )
}

scale_fill_heat <- function(...) {
  scale_fill_gradientn(
    colours = c("#F6E944", "#72C96B", "#2A788E", "#404788"),
    ...
  )
}

save_plot <- function(plot_obj, filename, width = 12, height = 8, dpi = 320) {
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

save_combined_plot <- function(plot_obj, filename, width = 12, height = 8, dpi = 320) {
  save_plot(plot_obj, filename, width = width, height = height, dpi = dpi)
  
  pdf_args <- list(
    filename = paste0(tools::file_path_sans_ext(filename), ".pdf"),
    plot = plot_obj,
    width = width,
    height = height,
    bg = "#FFFFFF",
    device = grDevices::pdf,
    useDingbats = FALSE
  )
  
  do.call(ggsave, pdf_args)
}

# Stack two existing PNG outputs vertically and add panel tags to each image.
save_vertical_image_panel <- function(top_path, bottom_path, output_path, tags = c("a", "b"), output_width = 16, export_pdf = FALSE) {
  if (!requireNamespace("png", quietly = TRUE)) {
    stop("Package 'png' is required to build the comparison image panels.")
  }
  
  build_panel_image <- function(image_path, tag_label) {
    image_array <- png::readPNG(image_path)
    aspect_ratio <- dim(image_array)[1] / dim(image_array)[2]
    image_grob <- grid::grobTree(
      grid::rectGrob(gp = grid::gpar(fill = "#FFFFFF", col = NA)),
      grid::rasterGrob(
        image_array,
        x = 0.5,
        y = 0.5,
        width = 1,
        height = 1,
        default.units = "npc",
        interpolate = TRUE
      ),
      grid::textGrob(
        tag_label,
        x = grid::unit(0.02, "npc"),
        y = grid::unit(0.98, "npc"),
        just = c("left", "top"),
        gp = grid::gpar(col = "#101820", fontsize = 30, fontface = "bold", fontfamily = plot_font_family)
      )
    )
    list(grob = image_grob, ratio = aspect_ratio)
  }
  
  top_panel <- build_panel_image(top_path, tags[1])
  bottom_panel <- build_panel_image(bottom_path, tags[2])
  
  combined_panel <- wrap_elements(full = top_panel$grob) /
    wrap_elements(full = bottom_panel$grob) +
    plot_layout(heights = c(top_panel$ratio, bottom_panel$ratio))
  
  output_height <- output_width * (top_panel$ratio + bottom_panel$ratio)
  
  if (isTRUE(export_pdf)) {
    save_combined_plot(
      combined_panel,
      output_path,
      width = output_width,
      height = output_height
    )
  } else {
    save_plot(
      combined_panel,
      output_path,
      width = output_width,
      height = output_height
    )
  }
}

save_horizontal_image_panel <- function(left_path, right_path, output_path, tags = c("a", "b"), output_height = 10) {
  if (!requireNamespace("png", quietly = TRUE)) {
    stop("Package 'png' is required to build the comparison image panels.")
  }
  
  build_panel_image <- function(image_path, tag_label) {
    image_array <- png::readPNG(image_path)
    aspect_ratio <- dim(image_array)[1] / dim(image_array)[2]
    image_grob <- grid::grobTree(
      grid::rectGrob(gp = grid::gpar(fill = "#FFFFFF", col = NA)),
      grid::rasterGrob(
        image_array,
        x = 0.5,
        y = 0.5,
        width = 1,
        height = 1,
        default.units = "npc",
        interpolate = TRUE
      ),
      grid::textGrob(
        tag_label,
        x = grid::unit(0.02, "npc"),
        y = grid::unit(0.98, "npc"),
        just = c("left", "top"),
        gp = grid::gpar(col = "#101820", fontsize = 30, fontface = "bold", fontfamily = plot_font_family)
      )
    )
    list(grob = image_grob, ratio = aspect_ratio)
  }
  
  left_panel <- build_panel_image(left_path, tags[1])
  right_panel <- build_panel_image(right_path, tags[2])
  
  combined_panel <- wrap_elements(full = left_panel$grob) +
    wrap_elements(full = right_panel$grob) +
    plot_layout(widths = c(1 / left_panel$ratio, 1 / right_panel$ratio))
  
  save_plot(
    combined_panel,
    output_path,
    width = output_height * ((1 / left_panel$ratio) + (1 / right_panel$ratio)),
    height = output_height
  )
}

order_with_other_last <- function(x) {
  x <- unique(as.character(x))
  x <- x[!is.na(x) & x != ""]
  other_mask <- str_detect(x, regex("other", ignore_case = TRUE))
  c(sort(x[!other_mask]), sort(x[other_mask]))
}

panel_tag_theme <- theme(
  plot.tag = element_text(face = "bold", size = 16, colour = "#101820"),
  plot.tag.position = c(0.01, 0.99)
)

# Stack a compact top summary bar and right summary bar around a heatmap.
compose_heatmap_with_margins <- function(heat_plot, top_plot, right_plot, top_height = 0.24, right_width = 0.30) {
  wrap_plots(
    top_plot,
    plot_spacer(),
    heat_plot,
    right_plot,
    ncol = 2,
    byrow = TRUE,
    widths = c(1, right_width),
    heights = c(top_height, 1)
  )
}

# Stack a compact top summary bar above a heatmap.
compose_heatmap_with_top <- function(heat_plot, top_plot, top_height = 0.28) {
  wrap_plots(top_plot, heat_plot, ncol = 1, heights = c(top_height, 1))
}

# Build the top marginal bar chart that summarises totals by x-axis category.
make_top_margin_bar <- function(data, x_col, y_col, fill_values, y_title = NULL) {
  ggplot(data, aes(x = .data[[x_col]], y = .data[[y_col]], fill = .data[[x_col]])) +
    geom_col(width = 0.82, show.legend = FALSE) +
    scale_fill_manual(values = fill_values, drop = FALSE) +
    labs(x = NULL, y = y_title) +
    theme_policy(8) +
    theme(
      axis.text.y = element_text(size = 6.5),
      axis.text.x = element_blank(),
      axis.ticks.x = element_blank(),
      axis.title.x = element_blank(),
      panel.grid.major.x = element_blank(),
      plot.margin = margin(5.5, 5.5, 0, 5.5)
    )
}

# Build the right marginal bar chart that summarises totals by y-axis category.
make_right_margin_bar <- function(data, x_col, y_col, fill_values, x_title = NULL) {
  ggplot(data, aes(x = .data[[x_col]], y = .data[[y_col]], fill = .data[[y_col]])) +
    geom_col(width = 0.82, show.legend = FALSE) +
    scale_fill_manual(values = fill_values, drop = FALSE) +
    labs(x = x_title, y = NULL) +
    theme_policy(8) +
    theme(
      axis.text.x = element_text(size = 6.5),
      axis.text.y = element_blank(),
      axis.ticks.y = element_blank(),
      axis.title.y = element_blank(),
      panel.grid.major.y = element_blank(),
      plot.margin = margin(0, 5.5, 5.5, 0)
    )
}

# Use a shared neutral grey for marginal summary bars.
make_uniform_fill <- function(levels, colour = "#AEB7C2") {
  setNames(rep(colour, length(levels)), levels)
}

set.seed(1234)

rescale_to_unit <- function(x) {
  rng <- range(x, na.rm = TRUE)
  if (diff(rng) == 0) {
    return(rep(0, length(x)))
  }
  (x - mean(rng)) / (diff(rng) / 2)
}

rescale_to_range <- function(x, to = c(0, 1)) {
  rng <- range(x, na.rm = TRUE)
  if (!all(is.finite(rng)) || diff(rng) == 0) {
    return(rep(mean(to), length(x)))
  }
  scales::rescale(x, to = to, from = rng)
}

normalize_angle <- function(angle) {
  atan2(sin(angle), cos(angle))
}

wrap_angle_2pi <- function(angle) {
  angle <- angle %% (2 * pi)
  ifelse(angle < 0, angle + 2 * pi, angle)
}

largest_gap_start <- function(angle) {
  if (length(angle) <= 1) {
    return(-pi / 2)
  }
  
  angle_sorted <- sort(wrap_angle_2pi(angle))
  gaps <- diff(c(angle_sorted, angle_sorted[1] + 2 * pi))
  gap_idx <- which.max(gaps)[1]
  
  angle_sorted[gap_idx] + gaps[gap_idx] / 2
}

evenly_spaced_angles <- function(angle, start = NULL) {
  n <- length(angle)
  
  if (n == 0) {
    return(numeric())
  }
  if (n == 1) {
    return(normalize_angle(if (is.null(start)) -pi / 2 else start))
  }
  
  start <- if (is.null(start)) largest_gap_start(angle) else start
  rotated <- wrap_angle_2pi(angle - start)
  order_idx <- order(rotated, method = "radix")
  
  target <- seq(0, 2 * pi, length.out = n + 1)[seq_len(n)] + start
  target <- normalize_angle(target)
  
  out <- numeric(n)
  out[order_idx] <- target
  out
}

blend_angles <- function(source_angle, target_angle, weight = 0.9) {
  weight <- max(0, min(1, weight))
  
  normalize_angle(
    atan2(
      (1 - weight) * sin(source_angle) + weight * sin(target_angle),
      (1 - weight) * cos(source_angle) + weight * cos(target_angle)
    )
  )
}

default_inner_shell_size <- function(n) {
  if (n <= 20) {
    return(5L)
  }
  if (n <= 60) {
    return(6L)
  }
  if (n <= 140) {
    return(8L)
  }
  10L
}

allocate_shell_ids <- function(n, inner_shell_size = NULL) {
  if (n <= 0) {
    return(integer())
  }
  
  inner_shell_size <- if (is.null(inner_shell_size)) default_inner_shell_size(n) else inner_shell_size
  inner_shell_size <- max(3L, as.integer(inner_shell_size))
  
  shell_sizes <- integer()
  remaining <- n
  shell_idx <- 1L
  
  while (remaining > 0) {
    shell_capacity <- inner_shell_size * shell_idx
    shell_take <- min(remaining, shell_capacity)
    shell_sizes <- c(shell_sizes, shell_take)
    remaining <- remaining - shell_take
    shell_idx <- shell_idx + 1L
  }
  
  rep(seq_along(shell_sizes), times = shell_sizes)
}

circularize_coords <- function(coords,
                               radial_power = 0.78,
                               inner_radius = 0.16,
                               max_radius = 1,
                               angle_strength = 0.92,
                               radius_strength = 0.84,
                               angle_start = NULL,
                               radial_order = NULL,
                               inner_shell_size = NULL) {
  if (nrow(coords) == 0) {
    return(cbind(x = numeric(), y = numeric()))
  }
  
  x <- coords[, 1]
  y <- coords[, 2]
  
  x <- x - mean(x, na.rm = TRUE)
  y <- y - mean(y, na.rm = TRUE)
  
  x <- rescale_to_unit(x)
  y <- rescale_to_unit(y)
  
  angle <- atan2(y, x)
  dist0 <- sqrt(x^2 + y^2)
  n <- length(dist0)
  
  if (!is.null(radial_order) && length(radial_order) == n && any(is.finite(radial_order))) {
    radial_order <- as.numeric(radial_order)
    radial_floor <- min(radial_order[is.finite(radial_order)], na.rm = TRUE)
    radial_order[!is.finite(radial_order)] <- radial_floor
    
    radial_sort <- order(-radial_order, wrap_angle_2pi(angle), method = "radix")
    shell_ids_sorted <- allocate_shell_ids(n, inner_shell_size = inner_shell_size)
    shell_id <- integer(n)
    shell_id[radial_sort] <- shell_ids_sorted
    
    base_start <- if (is.null(angle_start)) largest_gap_start(angle) else angle_start
    angle_uniform <- numeric(n)
    
    for (shell in seq_len(max(shell_id))) {
      idx <- which(shell_id == shell)
      m <- length(idx)
      if (m == 0) {
        next
      }
      
      shell_order <- order(wrap_angle_2pi(angle[idx] - base_start), method = "radix")
      target <- seq(0, 2 * pi, length.out = m + 1)[seq_len(m)] + base_start
      if (m > 1) {
        target <- target + ((shell - 1) %% 2) * (pi / m)
      }
      target <- normalize_angle(target)
      
      shell_angles <- numeric(m)
      shell_angles[shell_order] <- target
      angle_uniform[idx] <- shell_angles
    }
    
    n_shells <- max(shell_id)
    shell_progress <- if (n_shells == 1) 0 else seq(0, 1, length.out = n_shells)
    shell_radii <- inner_radius + (max_radius - inner_radius) * (shell_progress^radial_power)
    
    radius_uniform <- shell_radii[shell_id]
  } else {
    angle_uniform <- evenly_spaced_angles(angle, start = angle_start)
    ord <- rank(dist0, ties.method = "average")
    radius_rank <- pmax(ord - 0.5, 0.5) / max(n, 1)
    radius_uniform <- inner_radius + (max_radius - inner_radius) * ((sqrt(radius_rank))^radial_power)
  }
  
  angle_new <- blend_angles(angle, angle_uniform, weight = angle_strength)
  radius_base <- rescale_to_range(dist0, to = c(inner_radius, max_radius))
  r_new <- (1 - radius_strength) * radius_base + radius_strength * radius_uniform
  
  cbind(
    x = r_new * cos(angle_new),
    y = r_new * sin(angle_new)
  )
}

fr_circle_layout <- function(graph,
                             niter = 12000,
                             start_temp = 0.08,
                             grid = "nogrid",
                             radial_power = 0.78,
                             inner_radius = 0.16,
                             max_radius = 1.05,
                             angle_strength = 0.92,
                             radius_strength = 0.84,
                             angle_start = NULL,
                             radial_order = NULL,
                             inner_shell_size = NULL) {
  coords <- igraph::layout_with_fr(
    graph,
    weights = E(graph)$weight,
    niter = niter,
    start.temp = start_temp,
    grid = grid
  )
  
  coords2 <- circularize_coords(
    coords,
    radial_power = radial_power,
    inner_radius = inner_radius,
    max_radius = max_radius,
    angle_strength = angle_strength,
    radius_strength = radius_strength,
    angle_start = angle_start,
    radial_order = radial_order,
    inner_shell_size = inner_shell_size
  )
  
  create_layout(
    graph,
    layout = "manual",
    x = coords2[, 1],
    y = coords2[, 2]
  )
}

build_radial_label_data <- function(layout_data,
                                    label_names,
                                    size_col,
                                    size_range = c(3, 5),
                                    label_radius = 1.24,
                                    segment_gap = 0.03) {
  layout_data |>
    filter(name %in% label_names) |>
    mutate(
      angle = atan2(y, x),
      side = if_else(cos(angle) >= 0, 1, -1),
      label_x = label_radius * cos(angle),
      label_y = label_radius * sin(angle),
      segment_x = label_x - side * segment_gap,
      segment_y = label_y,
      hjust = if_else(side > 0, 0, 1),
      label_size = rescale_to_range(.data[[size_col]], to = size_range)
    )
}

build_node_label_data <- function(layout_data,
                                  label_names,
                                  size_col,
                                  size_range = c(3, 5)) {
  layout_data |>
    filter(name %in% label_names) |>
    mutate(
      label_size = rescale_to_range(.data[[size_col]], to = size_range)
    )
}

load_excel_ascii <- function(path, ...) {
  if (!file.exists(path)) {
    stop("Source Excel file does not exist: ", path)
  }
  
  tmp <- tempfile(fileext = ".xlsx")
  ok <- file.copy(path, tmp, overwrite = TRUE)
  
  if (!ok || !file.exists(tmp)) {
    stop(
      "Failed to copy Excel file to temp location.\n",
      "Source: ", path, "\n",
      "Temp: ", tmp
    )
  }
  
  read_excel(tmp, ...)
}

split_agency_list <- function(x) {
  x |>
    str_split("\\s*(?:,|;|，|；|、)\\s*|\\r\\n|\\n", simplify = FALSE) |>
    unlist() |>
    str_squish() |>
    purrr::discard(~ .x == "")
}

split_bilingual <- function(x) {
  x <- ifelse(is.na(x), "", x)
  out <- str_split_fixed(x, "\\r\\n|\\n", 2)
  tibble(en = str_squish(out[, 1]), zh = str_squish(out[, 2]))
}

extract_year_from_date <- function(x) {
  x <- trimws(as.character(x))
  x[x %in% c("", "NA", "NaN")] <- NA_character_
  numeric_mask <- suppressWarnings(!is.na(as.numeric(x)))
  years <- rep(NA_integer_, length(x))
  if (any(numeric_mask, na.rm = TRUE)) {
    numeric_values <- as.numeric(x[numeric_mask])
    excel_date <- as.Date(numeric_values, origin = "1899-12-30")
    years[numeric_mask] <- year(excel_date)
  }
  string_mask <- !numeric_mask & !is.na(x)
  if (any(string_mask)) {
    years[string_mask] <- suppressWarnings(as.integer(str_extract(x[string_mask], "(19|20)\\d{2}")))
  }
  years
}

read_txt_utf8 <- function(path) {
  txt <- readLines(path, warn = FALSE, encoding = "UTF-8")
  paste(txt, collapse = "\n")
}

standardise_agency_name <- function(x) {
  x <- str_squish(x)
  x <- recode(
    x,
    "中共中央" = "Central Committee of the Communist Party of China",
    "State Environmental Protection Administration" = "Ministry of Ecology and Environment",
    "Ministry of Environmental Protection" = "Ministry of Ecology and Environment",
    "National Environmental Protection Agency" = "Ministry of Ecology and Environment",
    "Ministry of Agriculture" = "Ministry of Agriculture and Rural Affairs",
    "State Forestry Administration" = "National Forestry and Grassland Administration",
    "Ministry of Forestry" = "National Forestry and Grassland Administration",
    "State Oceanic Administration" = "National Oceanic Administration",
    "Ministry of Land and Resources" = "Ministry of Natural Resources",
    "Ministry of Geology and Mineral Resources" = "Ministry of Natural Resources",
    "Ministry of Communications" = "Ministry of Transport",
    "Others" = NA_character_
  )
  x
}

agency_level <- function(x) {
  case_when(
    str_detect(x, "Communist Party|CPC|Central Committee") ~ "Party Central",
    str_detect(x, "National People's Congress|Standing Committee") ~ "National Legislature",
    str_detect(x, "^State Council$|General Office of State Council") ~ "State Council",
    str_detect(x, "Supreme People's Court|Supreme People's Procuratorate") ~ "Judicial",
    str_detect(x, "Commission|Ministry|Administration|Bureau|Customs|Bank|Office|Agency|Headquarters|Committee") ~ "Ministries and Central Agencies",
    str_detect(x, "People's Government|Province|Provincial|Municipal|Autonomous Region|Shanghai|Sichuan|Ningxia|Yellow River Conservancy Commission|Huaihe River") ~ "Local and Regional Authorities",
    TRUE ~ "Other Actors"
  )
}

subject_type_from_agencies <- function(agencies) {
  lvls <- unique(na.omit(agency_level(agencies)))
  if (length(lvls) == 0) {
    return("Unknown")
  }
  if ("Party Central" %in% lvls && length(lvls) > 1) {
    return("Party-State Joint")
  }
  if (length(lvls) > 1) {
    return("Interagency Joint Issuance")
  }
  lvls[1]
}

topic_family <- function(topic_en) {
  policy_object_group(topic_en)
}

actor_object_dictionary <- list(
  "Local Governments" = c("地方各级人民政府", "县级以上人民政府", "地方政府", "省级人民政府", "市县人民政府"),
  "Enterprises and Market Actors" = c("企业", "经营主体", "市场主体", "社会资本", "金融机构", "用能单位", "生产经营单位"),
  "Research and Public Institutions" = c("科研院所", "高等学校", "科研机构", "医疗机构", "事业单位", "实验室"),
  "Rural Producers and Resource Users" = c("农户", "渔民", "牧民", "林农", "养殖", "捕捞", "林场职工"),
  "Public and Social Organisations" = c("公众", "社会组织", "志愿者", "群众", "公民", "基层群众性自治组织"),
  "Protected Area Managers" = c("国家公园", "自然保护区", "湿地公园", "森林公园", "林场", "保护区管理机构")
)

dominant_actor_object <- function(text) {
  counts <- map_int(actor_object_dictionary, ~ sum(str_detect(text, fixed(.x))))
  if (all(counts == 0)) {
    return("General Governance Targets")
  }
  names(which.max(counts))
}

goal_dictionary <- list(
  "Conservation and Restoration" = c("保护", "修复", "恢复", "红线", "湿地", "森林", "草原", "野生", "栖息地", "生物多样性"),
  "Governance and Regulation" = c("制度", "机制", "法治", "监管", "标准", "考核", "规划", "清单", "名录"),
  "Science and Capacity" = c("科技", "研发", "研究", "创新", "监测", "调查", "平台", "人才"),
  "Green Development" = c("绿色", "低碳", "循环", "产业", "开发利用", "高质量发展", "生态产品"),
  "Participation and Education" = c("宣传", "教育", "公众", "社会参与", "培训", "志愿"),
  "Risk Prevention and Biosecurity" = c("风险", "安全", "防控", "预警", "应急", "生物安全"),
  "Climate and Carbon" = c("碳达峰", "碳中和", "碳汇", "温室气体", "气候")
)

tool_dictionary <- tibble(
  major_tool = c(
    "Supply-side", "Supply-side", "Supply-side", "Supply-side",
    "Environmental-side", "Environmental-side", "Environmental-side", "Environmental-side", "Environmental-side",
    "Demand-side", "Demand-side", "Demand-side", "Demand-side", "Demand-side"
  ),
  subtool = c(
    "Funding and Subsidy",
    "Infrastructure and Equipment",
    "Science and Technology",
    "Training and Education",
    "Planning and Target Setting",
    "Regulation and Standards",
    "Supervision and Assessment",
    "Institutional Reform and Coordination",
    "Fiscal and Financial Incentives",
    "Government Procurement and Service Purchase",
    "Market Creation and Trading",
    "Pilots and Demonstration",
    "Social Participation and Partnerships",
    "International Cooperation"
  ),
  patterns = list(
    c("资金", "财政", "补助", "补贴", "专项资金", "经费", "投资"),
    c("基础设施", "工程建设", "平台建设", "设施", "装备", "道路", "站点", "电网"),
    c("科技", "研发", "研究", "技术攻关", "实验室", "监测", "创新"),
    c("培训", "教育", "人才", "队伍建设", "宣传教育"),
    c("规划", "方案", "计划", "路线图", "目标", "行动方案", "清单", "目录"),
    c("条例", "办法", "规则", "规范", "标准", "准入", "许可"),
    c("监督", "检查", "考核", "评价", "问责", "审计", "执法"),
    c("改革", "体制", "机制", "职责分工", "联席", "协调机制", "领导小组"),
    c("税收", "信贷", "债券", "基金", "保险", "电价", "财政政策", "绿色金融"),
    c("政府采购", "购买服务", "委托", "外包"),
    c("市场", "交易", "碳汇", "生态补偿", "排放权", "配额", "用能权"),
    c("试点", "示范", "示范区", "先行", "试验区"),
    c("社会资本", "社会组织", "公众参与", "志愿者", "合作机制"),
    c("国际合作", "一带一路", "南南合作", "对外交流")
  )
)

source_type_from_title <- function(x) {
  case_when(
    str_detect(x, "法$|法（|法律") ~ "Law",
    str_detect(x, "条例") ~ "Regulation",
    str_detect(x, "规划|纲要|方案|计划") ~ "Plan or Programme",
    str_detect(x, "意见") ~ "Opinion",
    str_detect(x, "通知|通告") ~ "Notice",
    str_detect(x, "办法|细则") ~ "Measures or Detailed Rules",
    TRUE ~ "Other Source"
  )
}

extract_sentences <- function(text) {
  text |>
    str_replace_all("[\r\n]+", "\n") |>
    str_split("(?<=[。！？；])|\\n", simplify = FALSE) |>
    unlist() |>
    str_squish() |>
    purrr::discard(~ .x == "" || str_length(.x) < 8)
}

categorise_goal <- function(sentence) {
  hits <- map_int(goal_dictionary, ~ sum(str_detect(sentence, fixed(.x))))
  if (all(hits == 0)) {
    return("Other Goals")
  }
  names(which.max(hits))
}

detect_tool_hits <- function(sentence) {
  tool_dictionary |>
    rowwise() |>
    mutate(hit = any(str_detect(sentence, fixed(unlist(patterns))))) |>
    ungroup() |>
    filter(hit) |>
    select(major_tool, subtool)
}

irony_patterns <- tibble(
  signal_type = c(
    "Conservation-development tension",
    "Market-control tension",
    "Participation-command tension"
  ),
  left_terms = list(
    c("保护", "严格", "禁止", "红线", "管控", "减排"),
    c("市场", "社会资本", "优化", "便利", "简化"),
    c("公众参与", "社会组织", "志愿者", "共治")
  ),
  right_terms = list(
    c("开发", "利用", "建设", "扩张", "生产", "采伐", "开采", "旅游"),
    c("监管", "执法", "审查", "考核", "问责", "严控"),
    c("统一领导", "集中统一", "必须", "一律", "行政首长负责制")
  )
)

contrast_markers <- c("但", "同时", "兼顾", "统筹", "既要", "又要", "一方面", "另一方面")

detect_irony_sentence <- function(sentence) {
  if (!any(str_detect(sentence, fixed(contrast_markers)))) {
    return(NULL)
  }
  out <- irony_patterns |>
    rowwise() |>
    mutate(
      left_hit = any(str_detect(sentence, fixed(left_terms))),
      right_hit = any(str_detect(sentence, fixed(right_terms)))
    ) |>
    ungroup() |>
    filter(left_hit & right_hit) |>
    mutate(sentence = sentence) |>
    select(signal_type, sentence)
  if (nrow(out) == 0) {
    return(NULL)
  }
  out
}

# ----- Negative / context filtering for policy tool coding -----------------------
# These patterns disqualify a sentence from being coded as a tool hit when the
# matched keyword appears only in a negated, historical-reference, or
# definitional context rather than as an actual policy instrument.

tool_negation_patterns <- c(
  "尚未.*建立",  # "have not yet established"
  "缺乏",        # "lack of"
  "不足",        # "insufficient"
  "薄弱",        # "weak"
  "缺少",        # "lacking"
  "尚未",        # "not yet"
  "没有.*落实",  # "have not implemented"
  "未能"         # "failed to"
)

tool_historical_patterns <- c(
  "曾经",        # "once / in the past"
  "历史上",      # "historically"
  "过去",        # "in the past"
  "此前"         # "previously"
)

tool_definitional_patterns <- c(
  "是指",        # "refers to"
  "所称.*是",    # "what is meant by ... is"
  "本法所称",    # "as used in this law"
  "本条例所称",  # "as used in these regulations"
  "定义"         # "definition"
)

#' Check whether a sentence should be excluded from tool coding.
#' Returns TRUE when the sentence is negated, purely definitional, or
#' only references the tool concept in a historical/retrospective context.
sentence_is_disqualified <- function(sentence) {
  any(str_detect(sentence, tool_negation_patterns)) ||
    any(str_detect(sentence, tool_historical_patterns)) ||
    any(str_detect(sentence, tool_definitional_patterns))
}

#' Context-aware wrapper around the original detect_tool_hits().
#' Sentences that match a negation / definitional / historical pattern are
#' excluded so that only genuinely prescriptive tool usage is counted.
detect_tool_hits_filtered <- function(sentence) {
  if (sentence_is_disqualified(sentence)) {
    return(tibble(major_tool = character(), subtool = character()))
  }
  detect_tool_hits(sentence)
}

manual_theme_dictionary <- tribble(
  ~en, ~zh,
  "Biodiversity", "生物多样性",
  "National Park", "国家公园",
  "Nature Conservation", "自然保护地",
  "Ecological Protection Redline", "生态保护红线",
  "Habitat", "栖息地",
  "Invasive Species", "外来物种入侵",
  "Biosecurity", "生物安全",
  "Mangrove", "红树林",
  "Seagrass Bed", "海草床",
  "Wildlife", "野生动物",
  "Wild Plants", "野生植物",
  "River Basin", "流域",
  "Yangtze River Basin", "长江流域",
  "Yellow River Basin", "黄河流域",
  "Ecological Compensation", "生态补偿",
  "Ecological Restoration", "生态修复",
  "Ecological Security", "生态安全",
  "Natural Resources", "自然资源",
  "Territorial Spatial Planning", "国土空间规划",
  "Carbon Sink", "碳汇"
)

generic_theme_stoplist <- c(
  "Management", "Development", "Use", "Economy", "Development", "Mechanism",
  "Influence", "Function", "Policy", "Recover", "Classification", "Ecosystem",
  "Collaboration", "Strategy", "Plant", "Value", "Analysis", "Method", "Renew",
  "Cycle", "Question", "Problem", "Model", "Evaluation", "Overview",
  "Research Progress", "System", "Index", "Dynamic", "Changing Trends", "City",
  "Development And Utilization", "Reasonable Use", "Adaption", "Function",
  "Influencing Factors", "Countermeasures", "China", "Impact Factor", "Resource",
  "Environment", "Ecology", "Distribution", "Risk", "Policy", "Influence",
  "Classification", "New Era", "Function", "Analysis", "Method"
)

normalize_policy_domain <- function(x) {
  recode(
    as.character(x),
    "Other Governance Topics" = "Other Policy Domains",
    .default = as.character(x)
  )
}

object_order <- order_with_other_last(c(
  "Nature Conservation",
  "Habitats and Ecological Space",
  "Forests and Grasslands",
  "Waters and Aquatic Systems",
  "Ecological Restoration",
  "Climate and Carbon",
  "Risk, Biosecurity and Disasters",
  "Land, Agriculture and Natural Resources",
  "Pollution Control",
  "Tourism and Recreation",
  "Species Conservation and Use",
  "Other Policy Domains"
))

agency_level_palette <- c(
  "Judicial" = "#E6C84F",
  "Ministries and Central Agencies" = "#4DAF4A",
  "National Legislature" = "#A6761D",
  "Local and Regional Authorities" = "#66A61E",
  "Party Central" = "#4C78A8",
  "State Council" = "#E8529B",
  "Other Actors" = "#E67E22"
)

object_group_palette <- c(
  "Nature Conservation" = "#E78AC3",
  "Habitats and Ecological Space" = "#FC8D62",
  "Forests and Grasslands" = "#BEBADA",
  "Waters and Aquatic Systems" = "#80B1D3",
  "Ecological Restoration" = "#1B9E77",
  "Climate and Carbon" = "#8DD3C7",
  "Risk, Biosecurity and Disasters" = "#E6C84F",
  "Land, Agriculture and Natural Resources" = "#4C9ED9",
  "Pollution Control" = "#A6D854",
  "Tourism and Recreation" = "#CCEBC5",
  "Species Conservation and Use" = "#BC80BD",
  "Other Policy Domains" = "#AEB7C2"
)

subject_order <- order_with_other_last(c(
  "Party Central",
  "Party-State Joint",
  "National Legislature",
  "State Council",
  "Ministries and Central Agencies",
  "Judicial",
  "Local and Regional Authorities",
  "Interagency Joint Issuance",
  "Other Actors",
  "Unknown"
))

subject_type_palette <- setNames(
  hcl.colors(length(subject_order), "Dark 3"),
  subject_order
)

standardise_document_type <- function(x) {
  case_when(
    str_detect(x, regex("^law$", ignore_case = TRUE)) | str_detect(x, regex("law", ignore_case = TRUE)) ~ "Law",
    str_detect(x, regex("administrative regulation", ignore_case = TRUE)) ~ "Administrative Regulation",
    str_detect(x, regex("departmental regulation", ignore_case = TRUE)) ~ "Departmental Regulation",
    str_detect(x, regex("state council", ignore_case = TRUE)) ~ "State Council Normative Document",
    str_detect(x, regex("department(al)? normative document", ignore_case = TRUE)) ~ "Departmental Normative Document",
    str_detect(x, regex("department(al)? working document", ignore_case = TRUE)) ~ "Departmental Working Document",
    str_detect(x, regex("party regulation", ignore_case = TRUE)) ~ "Party Regulation",
    TRUE ~ x
  )
}

nounify_theme <- function(x) {
  x <- recode(
    x,
    "Protected Areas" = "Nature Conservation",
    "Use" = "Utilization",
    "Pollute" = "Pollution",
    "Recover" = "Restoration",
    "Renew" = "Renewal",
    "Adaption" = "Adaptation",
    "Influence" = "Impact",
    "Changing Trends" = "Trend Change",
    "Dynamic" = "Dynamics",
    "Reasonable Use" = "Rational Utilization",
    "Development And Utilization" = "Development and Utilization",
    "Research Progress" = "Research Progress",
    "Ai" = "AI",
    "Usa" = "USA",
    "Hongkong" = "Hong Kong",
    "Configuration Optimizion" = "Configuration Optimization",
    "Optimize Configuration" = "Configuration Optimization",
    "Soil And Water Conservation" = "Soil and Water Conservation",
    "Northern Slope Of Tianshan Mountains" = "Northern Slope of Tianshan Mountains",
    .default = x
  )
  
  x |>
    str_replace_all("\\bAnd\\b", "and") |>
    str_replace_all("\\bOf\\b", "of") |>
    str_replace_all("\\bTo\\b", "to")
}

theme_stoplist_key <- function(x) {
  x |>
    as.character() |>
    str_replace_all("_", " ") |>
    str_squish() |>
    str_to_title() |>
    nounify_theme()
}

policy_keyword_label_aliases <- c(
  "Sanjiangyuan" = "Sanjiang Plain"
)

canonicalize_policy_keyword_label <- function(x) {
  x_norm <- theme_stoplist_key(x)
  recode(x_norm, !!!policy_keyword_label_aliases, .default = x_norm)
}

load_keyword_stoplist <- function(stoplist_path) {
  if (!file.exists(stoplist_path)) {
    message("No keyword stoplist file found at: ", stoplist_path)
    return(character())
  }
  
  stoplist_raw <- read_csv(stoplist_path, show_col_types = FALSE)
  if (!"keyword" %in% names(stoplist_raw)) {
    stop("keyword_stoplist.csv must contain a 'keyword' column: ", stoplist_path)
  }
  
  stoplist_terms <- stoplist_raw |>
    transmute(keyword = theme_stoplist_key(keyword)) |>
    filter(!is.na(keyword), nzchar(keyword)) |>
    distinct(keyword) |>
    pull(keyword)
  
  message("Loaded ", length(stoplist_terms), " keyword stoplist terms from: ", stoplist_path)
  stoplist_terms
}

generic_theme_stoplist <- unique(c(
  theme_stoplist_key(generic_theme_stoplist),
  load_keyword_stoplist(keyword_stoplist_path)
))

is_theme_stopword <- function(x) {
  theme_stoplist_key(x) %in% generic_theme_stoplist
}

classify_theme_group <- function(en) {
  case_when(
    en %in% c(
      "Ecological Environment", "Ecological Civilization", "Ecological Functions",
      "Ecological Benefits", "Ecological Services", "Ecological Construction",
      "Ecological Value", "Ecological Processes", "Ecological Impact",
      "Compensation Mechanism", "Compensation Standards", "Natural Renewal",
      "Adaptation", "Sustainable Development", "Coordinated Development",
      "Value Realization", "Indicator System", "Evaluation Indicators",
      "Formation Mechanism", "Dynamic Changes"
    ) ~ "Ecological Restoration",
    en %in% c(
      "Environmental Impact", "Environmental Protection Industry"
    ) ~ "Pollution Control",
    en %in% c(
      "Green Development", "Circular Economy"
    ) ~ "Climate and Carbon",
    en %in% c(
      "Terrestrial Space", "Ecological Space", "Public Participation", "Remote Sensing"
    ) ~ "Habitats and Ecological Space",
    en %in% c(
      "Community Structure", "Spartina Alterniflora"
    ) ~ "Species Conservation and Use",
    en %in% c(
      "Xishuangbanna"
    ) ~ "Forests and Grasslands",
    str_detect(en, regex("Water|River|Lake|Wetland|Run-?Off|Runoff|Groundwater|Mangrove|Mangroves|Seagrass|Sea|Delta|Baiyangdian|Poyang|Dongting|Taihu|Bohai Sea|Yangtze|Yellow River|Huaihe|Tarim River|Lancang River|Reservoir|Watershed|Chongming Dongtan|Yancheng|Estuary|Coastal Wetlands", ignore_case = TRUE)) ~ "Waters and Aquatic Systems",
    str_detect(en, regex("Pollution|Environmental Quality|Air Pollution|Soil Pollution|Non-Point Source Pollution|Heavy Metal|Eutrophication|Total Quantity Control|Clean Production|Sulfur Dioxide|Environmental Management|Environmental Governance|Environmental Issues", ignore_case = TRUE)) ~ "Pollution Control",
    str_detect(en, regex("Climate|Carbon|Greenhouse|Methane|Warming|Carbon Sink|Carbon Peaking|Carbon Neutrality|Carbon Storage|Carbon Cycle|Green Finance|Reducing Pollution and Carbon Emissions", ignore_case = TRUE)) ~ "Climate and Carbon",
    str_detect(en, regex("National Park|Protected Area|Protected Areas|Nature Conservation|Ecological Protection$|Environmental Protection$|Sanjiang Plain|Ex Situ Conservation", ignore_case = TRUE)) ~ "Nature Conservation",
    str_detect(en, regex("Territorial Spatial Planning|Ecological Protection Redline|Ecological Space|Delineation|Territorial Space|Habitat|Corridor|Landscape|Functional Zoning|Spatial Pattern|Spatial Distribution|Distribution Pattern|Urban Agglomeration|Configuration Optimization|Optimize Configuration|Land Use|Coastal Zone|Ecological Land|Ecological Network|Multi-Planning Integration", ignore_case = TRUE)) ~ "Habitats and Ecological Space",
    str_detect(en, regex("Forest|Grassland|Vegetation|Shrub|Alpine Meadow|Plantation|Korean Pine|Rainforest|Daxinganling|Lesser Khingan|Changbai Mountain|Helan Mountains|Liupan Mountain|Qinling Mountains|Wuyi Mountain|Hexi Corridor|Desert Grassland|Tropical Rainforest|Arbor|Grassland Resources|Forest Resources|Vegetation Restoration|Vegetation Cover|Vegetation Type", ignore_case = TRUE)) ~ "Forests and Grasslands",
    str_detect(en, regex("Biodiversity|Diversity|Wildlife|Wild Plants|Species|Birds|Specimen|Aquatic Plants|Plant Communities|Red-Crowned Crane|^Fish$", ignore_case = TRUE)) ~ "Species Conservation and Use",
    str_detect(en, regex("Biosecurity|Invasive|Disaster|Risk|Drought|Earthquake|Typhoon|Sensitivity|Vulnerability|Human Interference|Interference|Natural Enemies|Geological Disasters|Biological Invasions|Forest Fire|^Fire$|Flood|Seawater Intrusion", ignore_case = TRUE)) ~ "Risk, Biosecurity and Disasters",
    str_detect(en, regex("Agriculture|Land|Natural Resources|Arable|Farmland|Soil|Crop|Farmers|Food Security|Grazing|Mineral Resources|Industrialization|Productive Forces|Development and Utilization|Rational Utilization|Property|Rural Revitalization|Land Consolidation|Land Management|Land Resources|Oasis|Rice|Wheat|Corn|Paddy|Soil Fertility|Arable Land Resources|Land Acquisition|County|Urbanization", ignore_case = TRUE)) ~ "Land, Agriculture and Natural Resources",
    str_detect(en, regex("Ecological Environment|Ecological Restoration|^Restoration$|Ecological Security|Ecological Compensation|Ecological Products|Desertification|Soil Erosion|Returning Farmland to Forest", ignore_case = TRUE)) ~ "Ecological Restoration",
    str_detect(en, regex("Ecotourism|Tourism|Recreation|Urban Green Space", ignore_case = TRUE)) ~ "Tourism and Recreation",
    str_detect(en, regex("Xinjiang|Inner Mongolia|Shanghai|Yunnan|Guangxi|Ningxia|Beijing|Qinghai-Tibet Plateau|Loess Plateau|Tibet|Northeast China|Western Region|Shandong Province|Jiangsu Province|Henan Province|Heihe|Hainan Island|Anhui Province|Jilin Province|Hong Kong|USA", ignore_case = TRUE)) ~ "Land, Agriculture and Natural Resources",
    TRUE ~ "Other Policy Domains"
  )
}

col_or_na <- function(df, col_name) {
  if (col_name %in% names(df)) {
    df[[col_name]]
  } else {
    rep(NA, nrow(df))
  }
}

policy_object_reference <- tibble(
  label = col_or_na(policy_object_reference_raw, "label"),
  group = col_or_na(policy_object_reference_raw, "group"),
  checked_size = suppressWarnings(as.integer(col_or_na(policy_object_reference_raw, "size"))),
  checked_first_policy_year = suppressWarnings(as.integer(col_or_na(policy_object_reference_raw, "first_policy_year"))),
  checked_community = suppressWarnings(as.integer(col_or_na(policy_object_reference_raw, "community")))
) |>
  transmute(
    label = canonicalize_policy_keyword_label(label),
    group = normalize_policy_domain(nounify_theme(group)),
    checked_size,
    checked_first_policy_year,
    checked_community
  ) |>
  filter(
    !is.na(label), label != "",
    !is.na(group), group %in% object_order
  ) |>
  distinct(label, .keep_all = TRUE)

policy_object_reference_map <- setNames(policy_object_reference$group, policy_object_reference$label)

policy_object_group <- function(x) {
  x_norm <- theme_stoplist_key(x)
  lookup <- unname(policy_object_reference_map[x_norm])
  fallback <- classify_theme_group(x_norm)
  
  lookup[is.na(x_norm) | x_norm == ""] <- NA_character_
  dplyr::coalesce(lookup, fallback)
}

align_policy_keywords_to_checked <- function(keyword_summary, policy_theme_years, checked_reference) {
  keyword_stats <- keyword_summary |>
    left_join(
      policy_theme_years |>
        select(theme_en, first_policy_year),
      by = "theme_en"
    ) |>
    transmute(
      raw_theme_en = theme_en,
      theme_key = canonicalize_policy_keyword_label(theme_en),
      theme_zh,
      theme_group = as.character(theme_group),
      document_frequency,
      first_policy_year = suppressWarnings(as.integer(first_policy_year))
    )
  
  exact_reference <- checked_reference |>
    transmute(
      theme_key = label,
      exact_label = label,
      exact_group = as.character(group),
      exact_size = checked_size,
      exact_first_policy_year = checked_first_policy_year,
      exact_community = checked_community
    )
  
  signature_reference <- checked_reference |>
    transmute(
      signature = paste(as.character(group), checked_size, checked_first_policy_year, sep = "||"),
      signature_label = label,
      signature_group = as.character(group),
      signature_size = checked_size,
      signature_first_policy_year = checked_first_policy_year,
      signature_community = checked_community
    ) |>
    filter(!is.na(signature_size), !is.na(signature_first_policy_year)) |>
    group_by(signature) |>
    filter(n() == 1) |>
    ungroup()
  
  aligned <- keyword_stats |>
    left_join(exact_reference, by = "theme_key") |>
    mutate(signature = paste(theme_group, document_frequency, first_policy_year, sep = "||")) |>
    left_join(signature_reference, by = "signature") |>
    mutate(
      canonical_label = coalesce(exact_label, signature_label),
      canonical_group = coalesce(exact_group, signature_group),
      checked_size = coalesce(exact_size, signature_size),
      checked_first_policy_year = coalesce(exact_first_policy_year, signature_first_policy_year),
      checked_community = coalesce(exact_community, signature_community),
      match_basis = case_when(
        !is.na(exact_label) ~ "checked_label",
        !is.na(signature_label) ~ "checked_signature",
        TRUE ~ NA_character_
      )
    ) |>
    select(
      raw_theme_en,
      theme_zh,
      theme_group,
      document_frequency,
      first_policy_year,
      canonical_label,
      canonical_group,
      checked_size,
      checked_first_policy_year,
      checked_community,
      match_basis
    )
  
  dropped_terms <- aligned |>
    filter(is.na(canonical_label)) |>
    pull(raw_theme_en)
  
  if (length(dropped_terms) > 0) {
    message(
      "Dropping ", length(dropped_terms),
      " policy keywords not present in gephi_nodes_checked.csv: ",
      paste(head(dropped_terms, 10), collapse = ", "),
      if (length(dropped_terms) > 10) " ..." else ""
    )
  }
  
  aligned |>
    filter(!is.na(canonical_label))
}

build_policy_corpus <- function() {
  metadata_candidates <- c(
    file.path(policy_dir, "政策文件清单.xlsx"),
    file.path(project_root, "..", "政策文件1205最终版", "全部政策文件.xlsx")
  )
  required_metadata_cols <- c(
    "No.",
    "Document Name",
    "Enacting Department",
    "First Enactment Date",
    "Latest Revision Date",
    "Document Nature",
    "Document Topic"
  )
  
  metadata_raw <- NULL
  metadata_path <- NA_character_
  for (candidate_path in metadata_candidates) {
    if (!file.exists(candidate_path)) {
      next
    }
    candidate_raw <- load_excel_ascii(candidate_path)
    if (all(required_metadata_cols %in% names(candidate_raw))) {
      metadata_raw <- candidate_raw
      metadata_path <- candidate_path
      break
    }
  }
  
  if (is.null(metadata_raw)) {
    stop(
      "Could not find a policy metadata workbook with required columns. Checked:\n",
      paste(metadata_candidates, collapse = "\n")
    )
  }
  
  message("Using policy metadata source: ", metadata_path)
  metadata <- metadata_raw |>
    mutate(
      source_policy_id = as.integer(No.),
      id = source_policy_id,
      title_en = split_bilingual(`Document Name`)$en,
      title_zh = split_bilingual(`Document Name`)$zh,
      agencies_en_raw = split_bilingual(`Enacting Department`)$en,
      agencies_zh_raw = split_bilingual(`Enacting Department`)$zh,
      year = extract_year_from_date(`First Enactment Date`),
      revision_year = extract_year_from_date(`Latest Revision Date`),
      nature_en = split_bilingual(`Document Nature`)$en,
      nature_zh = split_bilingual(`Document Nature`)$zh,
      document_type = standardise_document_type(split_bilingual(`Document Nature`)$en),
      topic_en = split_bilingual(`Document Topic`)$en,
      topic_zh = split_bilingual(`Document Topic`)$zh,
      topic_family = normalize_policy_domain(topic_family(topic_en))
    ) |>
    filter(!is.na(year), year >= policy_year_min, year <= policy_year_max) |>
    arrange(year, source_policy_id) |>
    mutate(file_id = row_number())
  
  message(
    sprintf(
      "Restricting policy sample to %d-%d: kept %d of %d metadata rows.",
      policy_year_min, policy_year_max, nrow(metadata), nrow(metadata_raw)
    )
  )
  
  txt_files <- list.files(policy_dir, pattern = "[.]txt$", full.names = TRUE)
  if (length(txt_files) == 0) {
    stop("No .txt policy files found in: ", policy_dir)
  }
  
  txt_tbl <- tibble(
    file = txt_files,
    file_name = basename(txt_files),
    file_id = suppressWarnings(as.integer(str_extract(file_name, "^[0-9]+(?=、)"))),
    file_id_start = suppressWarnings(as.integer(str_extract(file_name, "^[0-9]+(?=-)"))),
    file_id_end = suppressWarnings(as.integer(str_extract(file_name, "(?<=-)[0-9]+(?=、)")))
  )
  
  combined_file <- txt_tbl |> filter(!is.na(file_id_start), !is.na(file_id_end))
  if (nrow(combined_file) > 1) {
    stop("Expected at most one combined policy text file with an 'x-y、' prefix, found: ", nrow(combined_file))
  }
  
  regular_txt <- txt_tbl |>
    filter((is.na(file_id_start) | is.na(file_id_end)) & !is.na(file_id))
  
  texts <- regular_txt |>
    group_by(file_id) |>
    summarise(
      source_files = paste(sort(file_name), collapse = " | "),
      text = paste(map_chr(file, read_txt_utf8), collapse = "\n\n"),
      .groups = "drop"
    )
  
  if (nrow(combined_file) == 1) {
    combined_text <- read_txt_utf8(combined_file$file[[1]])
    marker <- str_locate(combined_text, "国有林区改革指导意见")[1]
    if (is.na(marker)) {
      stop("Could not split combined file 98-99 because marker '国有林区改革指导意见' was not found.")
    }
    texts <- bind_rows(
      texts,
      tibble(
        file_id = c(combined_file$file_id_start[[1]], combined_file$file_id_end[[1]]),
        source_files = combined_file$file_name[[1]],
        text = c(str_sub(combined_text, 1, marker - 1), str_sub(combined_text, marker, -1))
      )
    )
  }
  
  policy <- metadata |>
    left_join(texts, by = "file_id") |>
    mutate(
      agencies_en = map(
        agencies_en_raw,
        ~ .x |>
          split_agency_list() |>
          map_chr(standardise_agency_name) |>
          purrr::discard(is.na) |>
          unique()
      ),
      subject_type = map_chr(agencies_en, subject_type_from_agencies),
      actor_object = map_chr(text, dominant_actor_object),
      year = ifelse(is.na(year), as.integer(str_extract(text, "(19|20)\\d{2}")), year)
    ) |>
    mutate(
      subject_type = factor(subject_type, levels = subject_order),
      topic_family = factor(topic_family, levels = object_order),
      document_type = factor(document_type)
    ) |>
    arrange(year, id)
  
  if (any(is.na(policy$text) | policy$text == "")) {
    missing_ids <- policy$id[is.na(policy$text) | policy$text == ""]
    stop("Missing text for policy IDs: ", paste(missing_ids, collapse = ", "))
  }
  
  policy
}

build_keyword_dictionary <- function() {
  keyword_dict_path <- file.path(literature_dir, "关键词中译英.xlsx")
  if (!file.exists(keyword_dict_path)) {
    stop("Required file does not exist: ", keyword_dict_path)
  }
  
  literature_dict <- load_excel_ascii(keyword_dict_path, col_names = FALSE) |>
    transmute(en = `...1`, zh = `...2`) |>
    filter(!is.na(zh), !str_detect(en, "^#"))
  
  bind_rows(literature_dict, manual_theme_dictionary) |>
    distinct(en, .keep_all = TRUE) |>
    mutate(en = nounify_theme(en)) |>
    mutate(group = policy_object_group(en))
}

# ----- Inter-coder reliability assessment -----------------------------------------
# This module generates a stratified coding sample for a second coder and
# computes Cohen's kappa once the second coder's results are available.

generate_icr_sample <- function(policy_df, sample_fraction = 0.20, seed = 42) {
  set.seed(seed)
  strata <- policy_df |>
    group_by(topic_family) |>
    slice_sample(prop = sample_fraction) |>
    ungroup() |>
    transmute(
      policy_id = id,
      policy_title = title_en,
      year,
      document_type = as.character(document_type),
      coder1_object = as.character(topic_family),
      coder2_object = NA_character_
    ) |>
    arrange(policy_id)
  strata
}

compute_cohens_kappa <- function(icr_df) {
  if (!requireNamespace("irr", quietly = TRUE)) {
    message("Package 'irr' not installed; skipping kappa computation.")
    message("Install with: install.packages('irr')")
    return(list(kappa = NA_real_, z = NA_real_, p = NA_real_, n = nrow(icr_df),
                agreement_pct = NA_real_))
  }
  complete <- icr_df |> filter(!is.na(coder1_object), !is.na(coder2_object))
  if (nrow(complete) < 5) {
    message("Fewer than 5 dual-coded documents; cannot compute kappa.")
    return(list(kappa = NA_real_, z = NA_real_, p = NA_real_, n = nrow(complete),
                agreement_pct = NA_real_))
  }
  rating_matrix <- cbind(complete$coder1_object, complete$coder2_object)
  k <- irr::kappa2(rating_matrix, weight = "unweighted")
  agreement_pct <- mean(complete$coder1_object == complete$coder2_object) * 100
  list(
    kappa = k$value,
    z = k$statistic,
    p = k$p.value,
    n = nrow(complete),
    agreement_pct = agreement_pct
  )
}

policy <- build_policy_corpus()
keyword_dict <- build_keyword_dictionary()

write_csv(
  policy |>
    transmute(
      policy_id = id,
      text_file_id = file_id,
      year,
      title_en,
      title_zh
    ),
  file.path(output_root, "policy_text_file_id_map.csv")
)

# ----- Document-level classification: auto-generate then allow manual override -----
message("Writing auto-classified document object list and checking for manual overrides.")

doc_class_auto <- policy |>
  transmute(
    policy_id = id,
    title_en,
    title_zh,
    topic_en,
    topic_zh,
    auto_topic_family = as.character(topic_family),
    # Leave a blank column for the second step (manual correction).
    corrected_topic_family = NA_character_
  )

write_excel_csv(doc_class_auto, doc_class_generated_path)
message("  Auto-classification written to: ", doc_class_generated_path)

doc_class_checked_input_path <- case_when(
  file.exists(doc_class_checked_path) ~ doc_class_checked_path,
  file.exists(doc_class_checked_legacy_path) ~ doc_class_checked_legacy_path,
  TRUE ~ NA_character_
)

if (!is.na(doc_class_checked_input_path)) {
  message("  Found manually checked document classification; applying overrides from: ", doc_class_checked_input_path)
  doc_class_checked <- read_csv(doc_class_checked_input_path, show_col_types = FALSE)
  
  # Use corrected_topic_family when available, otherwise keep auto.
  override_map <- doc_class_checked |>
    mutate(corrected_topic_family = normalize_policy_domain(corrected_topic_family)) |>
    filter(!is.na(corrected_topic_family), corrected_topic_family != "") |>
    select(policy_id, corrected_topic_family)
  
  if (nrow(override_map) > 0) {
    matched_overrides <- sum(policy$id %in% override_map$policy_id)
    policy <- policy |>
      left_join(override_map, by = c("id" = "policy_id")) |>
      mutate(
        topic_family = if_else(
          !is.na(corrected_topic_family),
          factor(corrected_topic_family, levels = levels(topic_family)),
          topic_family
        )
      ) |>
      select(-corrected_topic_family)
    message(sprintf("  Overrode %d document classifications from checked file.", matched_overrides))
  } else {
    message("  Checked file exists but contains no overrides.")
  }
} else {
  message("  No checked file found at: ", doc_class_checked_path)
  message("  To correct misclassifications: fill 'corrected_topic_family' in the auto-generated CSV,")
  message("  save as 'document_domain_classification_checked.csv' in the same folder, then re-run.")
}

# ----- ICR execution --------------------------------------------------------------
message("Generating inter-coder reliability sample and checking for second-coder file.")

icr_sample_path <- file.path(output_root, "icr_coding_sample.csv")
icr_completed_path <- file.path(output_root, "icr_coding_completed.csv")

icr_sample <- generate_icr_sample(policy, sample_fraction = 0.20, seed = 42)
write_csv(icr_sample, icr_sample_path)
message("  ICR sample (", nrow(icr_sample), " documents, ~20%) written to: ", icr_sample_path)

if (file.exists(icr_completed_path)) {
  message("  Found completed ICR file; computing Cohen's kappa.")
  icr_completed <- read_csv(icr_completed_path, show_col_types = FALSE)
  icr_result <- compute_cohens_kappa(icr_completed)
  message(sprintf(
    "  Cohen's kappa = %.3f (z = %.2f, p = %.4f, n = %d, raw agreement = %.1f%%)",
    icr_result$kappa, icr_result$z, icr_result$p, icr_result$n, icr_result$agreement_pct
  ))
  icr_report <- tibble(
    metric = c("Cohen's kappa", "z-statistic", "p-value", "n (dual-coded)", "Raw agreement (%)"),
    value = c(icr_result$kappa, icr_result$z, icr_result$p, icr_result$n, icr_result$agreement_pct)
  )
  write_csv(icr_report, file.path(output_root, "icr_kappa_report.csv"))
} else {
  message("  No completed ICR file found at: ", icr_completed_path)
  message("  To compute kappa: have a second coder fill in 'coder2_object' in the sample CSV,")
  message("  save as 'icr_coding_completed.csv' in the same directory, then re-run.")
}

# ----- Policy keyword matching ---------------------------------------------------
# Match the bilingual keyword dictionary against every policy full text.
message("Matching keywords across policies.")

keyword_presence <- map(keyword_dict$zh, function(term) {
  str_detect(policy$text, fixed(term))
}) |>
  setNames(keyword_dict$en) |>
  as_tibble()

keyword_summary_raw <- tibble(
  theme_en = keyword_dict$en,
  theme_zh = keyword_dict$zh,
  theme_group = keyword_dict$group,
  document_frequency = map_int(seq_len(nrow(keyword_dict)), ~ sum(keyword_presence[[keyword_dict$en[.x]]]))
) |>
  filter(document_frequency > 0)

policy_theme_years_raw <- map_dfr(seq_len(nrow(keyword_summary_raw)), function(i) {
  theme <- keyword_summary_raw$theme_en[i]
  docs <- policy |>
    mutate(match = keyword_presence[[theme]]) |>
    filter(match) |>
    summarise(
      theme_en = theme,
      theme_zh = keyword_summary_raw$theme_zh[i],
      theme_group = keyword_summary_raw$theme_group[i],
      first_policy_year = min(year, na.rm = TRUE),
      document_frequency = n()
    )
  docs
})

keyword_alignment <- align_policy_keywords_to_checked(
  keyword_summary = keyword_summary_raw,
  policy_theme_years = policy_theme_years_raw,
  checked_reference = policy_object_reference
)

keyword_alias_map <- keyword_alignment |>
  distinct(canonical_label, raw_theme_en)

canonical_label_order <- policy_object_reference$label[
  policy_object_reference$label %in% keyword_alias_map$canonical_label
]

keyword_incidence <- map(canonical_label_order, function(label) {
  raw_terms <- keyword_alias_map$raw_theme_en[keyword_alias_map$canonical_label == label]
  as.integer(rowSums(keyword_presence[, raw_terms, drop = FALSE]) > 0)
}) |>
  setNames(canonical_label_order) |>
  as_tibble()

policy_theme_years <- tibble(
  theme_en = canonical_label_order,
  theme_zh = map_chr(canonical_label_order, function(label) {
    zh_terms <- keyword_alignment$theme_zh[keyword_alignment$canonical_label == label]
    zh_terms <- unique(zh_terms[!is.na(zh_terms) & zh_terms != ""])
    paste(zh_terms, collapse = " / ")
  }),
  theme_group = policy_object_reference$group[match(canonical_label_order, policy_object_reference$label)],
  first_policy_year = map_int(canonical_label_order, function(label) {
    matched_years <- policy$year[as.logical(keyword_incidence[[label]])]
    if (length(matched_years) == 0) {
      return(NA_integer_)
    }
    min(matched_years, na.rm = TRUE)
  }),
  document_frequency = map_int(canonical_label_order, function(label) sum(keyword_incidence[[label]]))
) |>
  filter(document_frequency > 0)

keyword_keep <- policy_theme_years |>
  filter(
    !is_theme_stopword(theme_en),
    document_frequency >= 8,
    document_frequency <= round(nrow(policy) * 0.75)
  ) |>
  arrange(desc(document_frequency), theme_en)

keyword_matrix <- keyword_incidence |>
  select(all_of(keyword_keep$theme_en)) |>
  mutate(policy_id = policy$id, policy_year = policy$year)

# ----- Dictionary coverage diagnostics --------------------------------------------
# Report how many policies are covered by at least one keyword, flag potential gaps,
# and identify high-frequency Chinese terms not yet in the dictionary.
message("Running dictionary coverage diagnostics.")

keyword_any_match <- rowSums(keyword_matrix[, keyword_keep$theme_en, drop = FALSE]) > 0
coverage_pct <- mean(keyword_any_match) * 100
uncovered_ids <- policy$id[!keyword_any_match]

message(sprintf("  Dictionary covers %.1f%% of policies (%d / %d).",
                coverage_pct, sum(keyword_any_match), nrow(policy)))
if (length(uncovered_ids) > 0) {
  message("  Uncovered policy IDs: ", paste(head(uncovered_ids, 20), collapse = ", "),
          if (length(uncovered_ids) > 20) " ..." else "")
}

# Candidate expansion: extract frequent bigrams from uncovered policy texts.
if (length(uncovered_ids) > 0) {
  uncovered_texts <- policy$text[policy$id %in% uncovered_ids]
  # Simple 2-4 character Chinese token extraction (crude but informative).
  candidate_tokens <- uncovered_texts |>
    str_extract_all("[\\x{4e00}-\\x{9fff}]{2,4}") |>
    unlist() |>
    table() |>
    sort(decreasing = TRUE)
  # Filter out tokens already in the dictionary.
  existing_zh <- keyword_dict$zh
  candidate_tokens <- candidate_tokens[!names(candidate_tokens) %in% existing_zh]
  top_candidates <- head(candidate_tokens, 30)
  if (length(top_candidates) > 0) {
    candidate_df <- tibble(
      chinese_term = names(top_candidates),
      frequency_in_uncovered = as.integer(top_candidates)
    )
    write_csv(candidate_df, file.path(output_root, "dictionary_expansion_candidates.csv"))
    message("  Top dictionary expansion candidates written to: dictionary_expansion_candidates.csv")
  }
}

coverage_report <- tibble(
  metric = c("Total policies", "Covered by dictionary", "Coverage %", "Uncovered count",
             "Keywords in dictionary", "Keywords after filtering"),
  value = c(nrow(policy), sum(keyword_any_match), round(coverage_pct, 1),
            length(uncovered_ids), nrow(keyword_dict), nrow(keyword_keep))
)
write_csv(coverage_report, file.path(output_root, "dictionary_coverage_report.csv"))

# ----- Policy subject-domain analysis --------------------------------------------
# This module compares policy types, policy subjects, and policy domains.
message("Running policy type-subject-domain analysis.")

subject_object_table <- policy |>
  transmute(
    policy_id = id,
    year,
    policy_title = title_en,
    document_type = as.character(document_type),
    subject_type = as.character(subject_type),
    topic_family = as.character(topic_family),
    issuing_agencies = map_chr(agencies_en, ~ paste(.x, collapse = " | ")),
    actor_object
  )

document_type_levels <- order_with_other_last(subject_object_table$document_type)
object_alpha_levels <- order_with_other_last(subject_object_table$topic_family)

subject_object_table <- subject_object_table |>
  mutate(
    document_type = factor(document_type, levels = document_type_levels),
    subject_type = factor(subject_type, levels = subject_order),
    topic_family = factor(topic_family, levels = object_order)
  )

subject_object_summary <- subject_object_table |>
  count(subject_type, topic_family, document_type, sort = TRUE)

document_type_palette <- setNames(hcl.colors(length(document_type_levels), "Set 2"), document_type_levels)

subject_doc_counts <- subject_object_table |>
  mutate(subject_type = factor(subject_type, levels = rev(subject_order))) |>
  count(subject_type, document_type)

object_doc_counts <- subject_object_table |>
  mutate(topic_family = factor(topic_family, levels = rev(object_order))) |>
  count(topic_family, document_type)

subject_object_counts <- subject_object_table |>
  mutate(
    subject_type = factor(subject_type, levels = rev(subject_order)),
    topic_family = factor(topic_family, levels = object_order)
  ) |>
  count(subject_type, topic_family)

subject_doc_heatmap <- subject_doc_counts |>
  ggplot(aes(document_type, subject_type, fill = n)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.4) +
  scale_fill_heat(name = "Policies") +
  labs(x = "Policy Type", y = "Policy Subject") +
  theme_policy(10) +
  theme(axis.text.x = element_text(angle = 35, hjust = 1))

object_doc_heatmap <- object_doc_counts |>
  mutate(
    document_type = factor(document_type, levels = rev(document_type_levels)),
    topic_family = factor(topic_family, levels = object_order)
  ) |>
  ggplot(aes(topic_family, document_type, fill = n)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.4) +
  scale_fill_heat(name = "Policies") +
  labs(x = "Policy Domain", y = "Policy Type") +
  theme_policy(10) +
  theme(axis.text.x = element_text(angle = 35, hjust = 1))

subject_object_heatmap <- subject_object_counts |>
  mutate(
    subject_type = factor(subject_type, levels = subject_order),
    topic_family = factor(topic_family, levels = rev(object_order))
  ) |>
  ggplot(aes(subject_type, topic_family, fill = n)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.4) +
  scale_fill_heat(name = "Policies") +
  labs(x = "Policy Subject", y = "Policy Domain") +
  theme_policy(10) +
  theme(axis.text.x = element_text(angle = 35, hjust = 1))

document_type_year_plot <- subject_object_table |>
  count(year, document_type) |>
  ggplot(aes(year, n, fill = document_type)) +
  geom_col(width = 0.85) +
  labs(x = "Year", y = "Number of Policies", fill = NULL) +
  guides(fill = guide_legend(nrow = 3, ncol = 3, byrow = TRUE)) +
  theme_policy(10) +
  theme(
    legend.text = element_text(size = 6.5),
    legend.key.size = grid::unit(0.28, "cm")
  )

subject_year_plot <- subject_object_table |>
  count(year, subject_type) |>
  ggplot(aes(year, n, fill = subject_type)) +
  geom_col(width = 0.85) +
  labs(x = "Year", y = "Policies", fill = NULL) +
  guides(fill = guide_legend(nrow = 2, byrow = TRUE)) +
  theme_policy(10) +
  theme(
    legend.text = element_text(size = 6.5),
    legend.key.size = grid::unit(0.28, "cm")
  )

object_year_plot <- subject_object_table |>
  mutate(topic_family = factor(topic_family, levels = object_order)) |>
  count(year, topic_family) |>
  ggplot(aes(year, n, fill = topic_family)) +
  geom_col(width = 0.85) +
  scale_fill_manual(values = object_group_palette, drop = FALSE) +
  labs(x = "Year", y = "Policies", fill = NULL) +
  guides(fill = guide_legend(nrow = 4, ncol = 3, byrow = TRUE)) +
  theme_policy(10) +
  theme(
    legend.text = element_text(size = 6.5),
    legend.key.size = grid::unit(0.28, "cm")
  )

subject_doc_top_bar <- subject_doc_counts |>
  count(document_type, wt = n, name = "total") |>
  mutate(document_type = factor(document_type, levels = document_type_levels)) |>
  make_top_margin_bar("document_type", "total", make_uniform_fill(document_type_levels), y_title = "Number of\nPolicies")

object_doc_top_bar <- object_doc_counts |>
  count(topic_family, wt = n, name = "total") |>
  mutate(topic_family = factor(topic_family, levels = object_order)) |>
  make_top_margin_bar("topic_family", "total", make_uniform_fill(object_order), y_title = NULL)

subject_object_top_bar <- subject_object_counts |>
  count(subject_type, wt = n, name = "total") |>
  mutate(subject_type = factor(subject_type, levels = subject_order)) |>
  make_top_margin_bar("subject_type", "total", make_uniform_fill(subject_order), y_title = NULL)

subject_doc_panel <- compose_heatmap_with_top(
  subject_doc_heatmap,
  subject_doc_top_bar,
  top_height = 0.18
)

object_doc_panel <- compose_heatmap_with_top(
  object_doc_heatmap,
  object_doc_top_bar,
  top_height = 0.18
)

subject_object_panel_plot <- compose_heatmap_with_top(
  subject_object_heatmap,
  subject_object_top_bar,
  top_height = 0.18
)

document_type_year_panel_plot <- document_type_year_plot
subject_year_panel_plot <- subject_year_plot + labs(y = NULL)
object_year_panel_plot <- object_year_plot + labs(y = NULL)
subject_doc_panel_wrapped <- wrap_elements(full = subject_doc_panel)
object_doc_panel_wrapped <- wrap_elements(full = object_doc_panel)
subject_object_panel_wrapped <- wrap_elements(full = subject_object_panel_plot)

subject_object_panel <- ((document_type_year_panel_plot + subject_year_panel_plot + object_year_panel_plot) + plot_layout(ncol = 3)) /
  ((subject_doc_panel_wrapped + subject_object_panel_wrapped + object_doc_panel_wrapped) + plot_layout(ncol = 3)) +
  plot_layout(ncol = 1, heights = c(0.67, 1.33)) +
  plot_annotation(tag_levels = list(c("a", "b", "c", "d", "e", "f"))) &
  panel_tag_theme

save_plot(subject_doc_heatmap, file.path(paths$policy_subject, "subject_policy_type_heatmap.png"), 8, 6)
save_plot(object_doc_heatmap, file.path(paths$policy_subject, "domain_policy_type_heatmap.png"), 8, 6)
save_plot(subject_object_heatmap, file.path(paths$policy_subject, "subject_domain_heatmap.png"), 8, 6)
save_plot(document_type_year_plot, file.path(paths$policy_subject, "policy_type_by_year.png"), 9, 6)
save_plot(subject_year_plot, file.path(paths$policy_subject, "subject_type_share_by_year.png"), 9, 6)
save_plot(object_year_plot, file.path(paths$policy_subject, "policy_domain_by_year.png"), 9, 6)
save_combined_plot(subject_object_panel, file.path(paths$policy_subject, "type_subject_domain.png"), 20.5, 13.4)

write_csv(subject_object_table, file.path(paths$policy_subject, "policy_type_subject_domain_table.csv"))
write_csv(subject_object_summary, file.path(paths$policy_subject, "type_subject_domain_summary.csv"))

policy_link_base <- policy |>
  arrange(year, id) |>
  transmute(
    policy_id = id,
    policy_seq = row_number(),
    year,
    document_type = factor(as.character(document_type), levels = document_type_levels),
    topic_family = factor(as.character(topic_family), levels = object_order),
    agencies_en
  )

policy_band <- policy_link_base |>
  mutate(policy_y = rev(seq_len(n())))

agency_link_long <- policy_link_base |>
  select(policy_id, policy_seq, agencies_en) |>
  unnest_longer(agencies_en, values_to = "agency_name")

agency_band <- agency_link_long |>
  group_by(agency_name) |>
  summarise(link_mid = median(policy_seq), link_n = n(), .groups = "drop") |>
  arrange(link_mid, agency_name) |>
  mutate(
    agency_y = seq(max(policy_band$policy_y), min(policy_band$policy_y), length.out = n()),
    agency_color = hcl.colors(n(), "Dynamic")
  )

object_present <- object_order[object_order %in% unique(as.character(policy_link_base$topic_family))]
object_band <- tibble(topic_family = object_present) |>
  mutate(
    object_y = seq(
      mean(range(policy_band$policy_y)) + diff(range(policy_band$policy_y)) * 0.30,
      mean(range(policy_band$policy_y)) - diff(range(policy_band$policy_y)) * 0.30,
      length.out = n()
    ),
    object_color = unname(object_group_palette[topic_family])
  )

# ----- Agency-document-object linkage --------------------------------------------
# This Sankey-like layout traces each policy from issuing subject to file to object.
band_title_data <- tibble(
  x = c(1, 5, 9),
  y = max(policy_band$policy_y) + 7,
  label = c("Policy Subject", "Policy File", "Policy Domain")
)

left_links <- agency_link_long |>
  left_join(policy_band |> select(policy_id, policy_y), by = "policy_id") |>
  left_join(agency_band |> select(agency_name, agency_y, agency_color), by = "agency_name")

right_links <- policy_link_base |>
  select(policy_id, policy_seq, topic_family) |>
  left_join(policy_band |> select(policy_id, policy_y), by = "policy_id") |>
  left_join(object_band |> select(topic_family, object_y, object_color), by = "topic_family")

policy_unit_document_object_linkage <- left_links |>
  left_join(
    right_links |>
      select(policy_id, topic_family, object_y, object_color),
    by = "policy_id"
  ) |>
  left_join(policy_link_base |> select(policy_id, document_type), by = "policy_id")

linkage_plot <- ggplot() +
  geom_segment(
    data = left_links,
    aes(x = 1.15, xend = 4.85, y = agency_y, yend = policy_y),
    colour = left_links$agency_color,
    linewidth = 0.22,
    alpha = 0.22
  ) +
  geom_segment(
    data = right_links,
    aes(x = 5.15, xend = 8.85, y = policy_y, yend = object_y),
    colour = right_links$object_color,
    linewidth = 0.22,
    alpha = 0.34
  ) +
  geom_tile(
    data = agency_band,
    aes(x = 1, y = agency_y),
    width = 0.22,
    height = if (nrow(agency_band) > 1) abs(diff(agency_band$agency_y)[1]) * 0.82 else 1.2,
    fill = agency_band$agency_color
  ) +
  geom_tile(
    data = policy_band,
    aes(x = 5, y = policy_y, fill = document_type),
    width = 0.26,
    height = 0.88
  ) +
  geom_tile(
    data = object_band,
    aes(x = 9, y = object_y),
    width = 0.22,
    height = if (nrow(object_band) > 1) abs(diff(object_band$object_y)[1]) * 0.82 else 1.2,
    fill = object_band$object_color
  ) +
  geom_text(
    data = agency_band,
    aes(x = 0.84, y = agency_y, label = agency_name),
    hjust = 1,
    size = 3,
    fontface = "bold",
    colour = "#16203A"
  ) +
  geom_text(
    data = policy_band,
    aes(x = 5, y = policy_y, label = policy_seq),
    size = 1.7,
    fontface = "bold",
    colour = "#16203A"
  ) +
  geom_text(
    data = object_band,
    aes(x = 9.16, y = object_y, label = topic_family),
    hjust = 0,
    size = 4.3,
    fontface = "bold",
    colour = "#16203A"
  ) +
  geom_text(
    data = band_title_data,
    aes(x = x, y = y, label = label),
    size = 5.9,
    fontface = "bold",
    colour = "#16203A"
  ) +
  scale_fill_manual(values = document_type_palette, name = "Policy Type") +
  guides(fill = guide_legend(nrow = 2, ncol = 4, byrow = TRUE)) +
  coord_cartesian(xlim = c(0, 10), ylim = c(min(policy_band$policy_y) - 3, max(policy_band$policy_y) + 10), clip = "off") +
  theme_void(base_family = plot_font_family) +
  theme(
    plot.margin = margin(12, 150, 12, 220),
    plot.background = element_rect(fill = "#FFFFFF", colour = NA),
    legend.position = "bottom",
    legend.title = element_text(face = "bold"),
    legend.text = element_text(size = 8)
  )

save_combined_plot(linkage_plot, file.path(paths$policy_subject, "agency_document_domain_linkage.png"), 20, 30)
write_csv(policy_unit_document_object_linkage, file.path(paths$policy_subject, "agency_document_domain_linkage.csv"))

# ----- Intergovernmental power network -------------------------------------------
# This network shows co-issuance ties among policy subjects.
message("Building the intergovernmental power network.")

agency_edges <- map_dfr(seq_len(nrow(policy)), function(i) {
  agencies <- policy$agencies_en[[i]]
  if (length(agencies) < 2) {
    return(tibble())
  }
  combos <- combn(sort(unique(agencies)), 2)
  tibble(
    source = combos[1, ],
    target = combos[2, ],
    policy_id = policy$id[i],
    year = policy$year[i]
  )
})

intergov_edges <- agency_edges |>
  count(source, target, name = "weight", sort = TRUE)

# Compute Jaccard-normalised edge weights to control for agency prolificacy.
agency_doc_counts <- agency_link_long |>
  count(agency_name, name = "n_docs")

intergov_edges <- intergov_edges |>
  left_join(agency_doc_counts, by = c("source" = "agency_name")) |>
  rename(n_source = n_docs) |>
  left_join(agency_doc_counts, by = c("target" = "agency_name")) |>
  rename(n_target = n_docs) |>
  mutate(
    jaccard = weight / (n_source + n_target - weight),
    cosine = weight / sqrt(n_source * n_target)
  ) |>
  select(source, target, weight, jaccard, cosine)

intergov_nodes <- tibble(name = sort(unique(c(intergov_edges$source, intergov_edges$target)))) |>
  mutate(`level.x` = agency_level(name))

intergov_level_order <- order_with_other_last(intergov_nodes$`level.x`)
intergov_nodes <- intergov_nodes |>
  mutate(`level.x` = factor(`level.x`, levels = intergov_level_order))

# Build two graph objects: one with raw weights (for visual sizing) and one
# with Jaccard weights (for centrality computation).
g_intergov <- graph_from_data_frame(intergov_edges |> select(source, target, weight),
                                    directed = FALSE,
                                    vertices = intergov_nodes |> rename(level_x = `level.x`))

g_intergov_jaccard <- graph_from_data_frame(
  intergov_edges |> transmute(source, target, weight = jaccard),
  directed = FALSE,
  vertices = intergov_nodes |> rename(level_x = `level.x`)
)

intergov_nodes <- intergov_nodes |>
  mutate(
    # Raw-weight metrics (for node sizing in plots).
    weighted_degree = strength(g_intergov, vids = V(g_intergov), mode = "all",
                               weights = E(g_intergov)$weight),
    degree = degree(g_intergov),
    # Jaccard-normalised centrality metrics (reported in the paper).
    betweenness_jaccard = betweenness(g_intergov_jaccard, normalized = TRUE,
                                      weights = 1 / E(g_intergov_jaccard)$weight),
    eigenvector_jaccard = eigen_centrality(g_intergov_jaccard,
                                           weights = E(g_intergov_jaccard)$weight)$vector,
    closeness_jaccard = suppressWarnings(
      closeness(g_intergov_jaccard, normalized = TRUE,
                weights = 1 / E(g_intergov_jaccard)$weight)
    ),
    # Raw-weight centralities for comparison / sensitivity.
    betweenness_raw = betweenness(g_intergov, normalized = TRUE,
                                  weights = 1 / E(g_intergov)$weight),
    eigenvector_raw = eigen_centrality(g_intergov, weights = E(g_intergov)$weight)$vector,
    closeness_raw = suppressWarnings(
      closeness(g_intergov, normalized = TRUE,
                weights = 1 / E(g_intergov)$weight)
    ),
    # Additional structural metrics.
    clustering_coeff = transitivity(g_intergov, type = "local", isolates = "zero"),
    pagerank_raw = page_rank(g_intergov, weights = E(g_intergov)$weight)$vector,
    pagerank_jaccard = page_rank(g_intergov_jaccard, weights = E(g_intergov_jaccard)$weight)$vector,
    constraint_raw = constraint(g_intergov, weights = E(g_intergov)$weight),
    constraint_jaccard = constraint(g_intergov_jaccard, weights = E(g_intergov_jaccard)$weight),
    kcore = coreness(g_intergov)
  )

# Hub and authority scores (HITS algorithm).
intergov_hub_auth_raw <- hub_score(g_intergov, weights = E(g_intergov)$weight)
intergov_hub_auth_jac <- hub_score(g_intergov_jaccard, weights = E(g_intergov_jaccard)$weight)
intergov_auth_raw <- authority_score(g_intergov, weights = E(g_intergov)$weight)
intergov_auth_jac <- authority_score(g_intergov_jaccard, weights = E(g_intergov_jaccard)$weight)

intergov_nodes <- intergov_nodes |>
  mutate(
    hub_raw = intergov_hub_auth_raw$vector[match(name, V(g_intergov)$name)],
    hub_jaccard = intergov_hub_auth_jac$vector[match(name, V(g_intergov_jaccard)$name)],
    authority_raw = intergov_auth_raw$vector[match(name, V(g_intergov)$name)],
    authority_jaccard = intergov_auth_jac$vector[match(name, V(g_intergov_jaccard)$name)]
  )

# Report: rank agencies by theoretically meaningful metrics.
# Betweenness → brokerage power (条块分割 gatekeeping).
# Eigenvector → hierarchical influence (connected to other influential actors).
# Constraint → structural holes (low constraint = spanning structural holes).
# PageRank → network importance accounting for neighbour quality.
centrality_report <- intergov_nodes |>
  select(name, `level.x`, degree, weighted_degree,
         betweenness_jaccard, eigenvector_jaccard, closeness_jaccard,
         pagerank_jaccard, constraint_jaccard,
         hub_jaccard, authority_jaccard,
         clustering_coeff, kcore,
         betweenness_raw, eigenvector_raw, closeness_raw,
         pagerank_raw, constraint_raw,
         hub_raw, authority_raw) |>
  arrange(desc(betweenness_jaccard))

write_csv(centrality_report, file.path(paths$policy_network, "centrality_report.csv"))
message("  Centrality report written (betweenness, eigenvector, PageRank, constraint, hub/authority, clustering, k-core).")

V(g_intergov)$weighted_degree <- intergov_nodes$weighted_degree[match(V(g_intergov)$name, intergov_nodes$name)]
V(g_intergov)$level_x <- intergov_nodes$`level.x`[match(V(g_intergov)$name, intergov_nodes$name)]

label_intergov <- intergov_nodes

intergov_layout <- fr_circle_layout(
  g_intergov,
  niter = 15000,
  start_temp = 0.10,
  radial_power = 0.86,
  inner_radius = 0.20,
  max_radius = 1.10,
  angle_strength = 0.98,
  radius_strength = 0.98,
  radial_order = V(g_intergov)$weighted_degree,
  inner_shell_size = 6
)

intergov_label_data <- build_node_label_data(
  as_tibble(intergov_layout),
  label_intergov$name,
  size_col = "weighted_degree",
  size_range = c(1.8, 3.6)
)

intergov_network_plot <- ggraph(intergov_layout) +
  geom_edge_link(
    aes(width = weight),
    alpha = 0.12,
    colour = "#7C8EA8",
    lineend = "round"
  ) +
  geom_node_point(
    aes(size = weighted_degree, colour = level_x),
    alpha = 0.92
  ) +
  ggrepel::geom_text_repel(
    data = intergov_label_data,
    aes(x = x, y = y, label = name),
    inherit.aes = FALSE,
    size = intergov_label_data$label_size,
    family = plot_font_family,
    fontface = "plain",
    colour = "#16203A",
    box.padding = 0.10,
    point.padding = 0.06,
    force = 0.75,
    force_pull = 2.1,
    max.overlaps = Inf,
    min.segment.length = Inf,
    segment.colour = NA,
    seed = 1234
  ) +
  scale_edge_width(range = c(0.14, 2.15), name = "Edge weight") +
  scale_size_continuous(range = c(2.6, 9.6), name = "Node size") +
  scale_colour_manual(values = agency_level_palette, name = "Policy Subject", drop = FALSE) +
  guides(
    edge_width = guide_legend(order = 1, nrow = 1),
    size = guide_legend(order = 2, nrow = 1),
    colour = guide_legend(order = 3, nrow = 1)
  ) +
  coord_equal(xlim = c(-1.42, 1.42), ylim = c(-0.8, 1.2), clip = "off") +
  theme_void(base_family = plot_font_family) +
  theme(
    legend.position = "bottom",
    legend.box = "vertical",
    plot.margin = margin(4, 32, 4, 32),
    plot.background = element_rect(fill = "#FFFFFF", colour = NA)
  )

save_plot(intergov_network_plot, file.path(paths$policy_network, "intergovernmental_power_network.png"), 15, 10.5)
save_vertical_image_panel(
  file.path(paths$policy_network, "intergovernmental_power_network.png"),
  file.path(literature_dir, "20250922机构图.png"),
  file.path(paths$policy_network, "intergovernmental_power_network_comparison_panel.png"),
  output_width = 15,
  export_pdf = TRUE
)
write_csv(intergov_nodes, file.path(paths$policy_network, "gephi_nodes.csv"))
write_csv(intergov_edges, file.path(paths$policy_network, "gephi_edges.csv"))

# ----- Policy keyword co-occurrence network --------------------------------------
# This network highlights how policy domains co-appear inside policy texts.
message("Building the policy keyword network.")

keyword_incidence <- as.matrix(
  keyword_matrix |>
    select(all_of(keyword_keep$theme_en)) |>
    mutate(across(everything(), as.integer))
)

keyword_cooccurrence <- t(keyword_incidence) %*% keyword_incidence
diag(keyword_cooccurrence) <- 0

keyword_edges <- as.data.frame(as.table(keyword_cooccurrence)) |>
  transmute(
    source = as.character(Var1),
    target = as.character(Var2),
    weight = as.integer(Freq)
  ) |>
  filter(source < target, weight >= 6) |>
  arrange(desc(weight))

keyword_keep <- keyword_keep |>
  mutate(theme_group = factor(theme_group, levels = object_order))

keyword_nodes <- keyword_keep |>
  transmute(
    id = theme_en,
    label = theme_en,
    label_zh = theme_zh,
    group = as.character(theme_group),
    size = document_frequency,
    first_policy_year
  ) |>
  left_join(
    policy_object_reference |>
      select(label, checked_size, checked_first_policy_year, checked_community),
    by = c("id" = "label")
  ) |>
  mutate(
    size = coalesce(checked_size, size),
    first_policy_year = coalesce(checked_first_policy_year, first_policy_year)
  )

g_keyword <- graph_from_data_frame(keyword_edges, directed = FALSE, vertices = keyword_nodes |> rename(name = id))
keyword_membership_computed <- cluster_louvain(g_keyword, weights = E(g_keyword)$weight)$membership
keyword_membership <- if_else(
  is.na(keyword_nodes$checked_community),
  as.integer(keyword_membership_computed),
  keyword_nodes$checked_community
)

# Compute comprehensive keyword network centrality metrics.
kw_betweenness <- betweenness(g_keyword, normalized = TRUE,
                              weights = 1 / E(g_keyword)$weight)
kw_eigenvector <- eigen_centrality(g_keyword, weights = E(g_keyword)$weight)$vector
kw_closeness   <- suppressWarnings(
  closeness(g_keyword, normalized = TRUE, weights = 1 / E(g_keyword)$weight)
)
kw_pagerank    <- page_rank(g_keyword, weights = E(g_keyword)$weight)$vector
kw_constraint  <- constraint(g_keyword, weights = E(g_keyword)$weight)
kw_clustering  <- transitivity(g_keyword, type = "local", isolates = "zero")
kw_kcore       <- coreness(g_keyword)
kw_hub         <- hub_score(g_keyword, weights = E(g_keyword)$weight)$vector
kw_authority   <- authority_score(g_keyword, weights = E(g_keyword)$weight)$vector

keyword_nodes_out <- tibble(
  label = V(g_keyword)$name,
  group = factor(V(g_keyword)$group, levels = object_order),
  size = V(g_keyword)$size,
  first_policy_year = V(g_keyword)$first_policy_year,
  community = as.integer(keyword_membership),
  betweenness = kw_betweenness,
  eigenvector = kw_eigenvector,
  closeness = kw_closeness,
  pagerank = kw_pagerank,
  constraint = kw_constraint,
  clustering_coeff = kw_clustering,
  kcore = kw_kcore,
  hub = kw_hub,
  authority = kw_authority
)

V(g_keyword)$size <- keyword_nodes_out$size[match(V(g_keyword)$name, keyword_nodes_out$label)]
V(g_keyword)$group <- keyword_nodes_out$group[match(V(g_keyword)$name, keyword_nodes_out$label)]

label_keywords <- keyword_nodes_out

keyword_layout <- fr_circle_layout(
  g_keyword,
  niter = 18000,
  start_temp = 0.12,
  radial_power = 0.84,
  inner_radius = 0.18,
  max_radius = 1.12,
  angle_strength = 0.99,
  radius_strength = 0.99,
  radial_order = V(g_keyword)$size,
  inner_shell_size = 10
)

keyword_layout$community <- keyword_nodes_out$community[match(keyword_layout$name, keyword_nodes_out$label)]
keyword_layout$size <- keyword_nodes_out$size[match(keyword_layout$name, keyword_nodes_out$label)]
keyword_layout$group <- factor(
  keyword_nodes_out$group[match(keyword_layout$name, keyword_nodes_out$label)],
  levels = object_order
)

keyword_label_data <- build_node_label_data(
  as_tibble(keyword_layout),
  label_keywords$label,
  size_col = "size",
  size_range = c(1.1, 2.9)
)

keyword_network_plot <- ggraph(keyword_layout) +
  geom_edge_link(
    aes(width = weight),
    alpha = 0.08,
    colour = "#7C8EA8",
    lineend = "round"
  ) +
  geom_node_point(
    aes(size = size, colour = group),
    alpha = 0.90
  ) +
  ggrepel::geom_text_repel(
    data = keyword_label_data,
    aes(x = x, y = y, label = name),
    inherit.aes = FALSE,
    size = keyword_label_data$label_size,
    family = plot_font_family,
    fontface = "plain",
    colour = "#16203A",
    box.padding = 0.04,
    point.padding = 0.02,
    force = 0.40,
    force_pull = 1.6,
    max.overlaps = Inf,
    min.segment.length = Inf,
    segment.colour = NA,
    seed = 1234
  ) +
  scale_edge_width(range = c(0.12, 2.05), name = "Edge weight") +
  scale_size_continuous(range = c(1.8, 7.8), name = "Node size") +
  scale_colour_manual(values = object_group_palette, name = "Policy Domain", drop = FALSE) +
  guides(
    edge_width = guide_legend(order = 1, nrow = 1),
    size = guide_legend(order = 2, nrow = 1),
    colour = guide_legend(order = 3, nrow = 2, byrow = TRUE)
  ) +
  coord_equal(xlim = c(-1.46, 1.46), ylim = c(-1.1, 1.2), clip = "off") +
  theme_void(base_family = plot_font_family) +
  theme(
    legend.position = "bottom",
    legend.box = "vertical",
    plot.margin = margin(4, 34, 4, 34),
    plot.background = element_rect(fill = "#FFFFFF", colour = NA)
  )

save_plot(keyword_network_plot, file.path(paths$keyword_network, "policy_keyword_network.png"), 16, 11)
save_vertical_image_panel(
  file.path(paths$keyword_network, "policy_keyword_network.png"),
  file.path(literature_dir, "20250909V8关键词共现.png"),
  file.path(paths$keyword_network, "policy_keyword_network_comparison_panel.png"),
  output_width = 16,
  export_pdf = TRUE
)
write_csv(keyword_nodes_out, policy_keyword_nodes_generated_path)
write_csv(keyword_edges, file.path(paths$keyword_network, "gephi_edges.csv"))
write_csv(keyword_keep, file.path(paths$keyword_network, "policy_keyword_frequency_table.csv"))

# ----- Policy tool analysis ------------------------------------------------------
# Identify tool-coded sentences and compare tool mixes across groups.
message("Identifying policy tools.")

tool_hits <- map_dfr(seq_len(nrow(policy)), function(i) {
  sentences <- extract_sentences(policy$text[i])
  out <- map_dfr(sentences, function(sent) {
    # Use context-filtered detection: negated / definitional / historical
    # sentences are excluded from tool coding.
    hits <- detect_tool_hits_filtered(sent)
    if (nrow(hits) == 0) return(tibble())
    hits |> mutate(sentence = sent)
  })
  if (nrow(out) == 0) {
    return(tibble())
  }
  out |>
    mutate(
      policy_id = policy$id[i],
      year = policy$year[i],
      policy_title = policy$title_en[i],
      document_type = as.character(policy$document_type[i]),
      subject_type = as.character(policy$subject_type[i]),
      topic_family = as.character(policy$topic_family[i])
    )
})

tool_major_levels <- c("Demand-side", "Environmental-side", "Supply-side")

tool_axis_lookup <- tool_dictionary |>
  mutate(
    side_order = factor(major_tool, levels = tool_major_levels),
    subtool_label = str_c(major_tool, ": ", subtool)
  ) |>
  arrange(side_order, subtool)

tool_subtool_levels <- rev(tool_axis_lookup$subtool_label)

tool_hits <- tool_hits |>
  left_join(tool_axis_lookup |> select(major_tool, subtool, subtool_label), by = c("major_tool", "subtool")) |>
  mutate(
    subtool_label = factor(subtool_label, levels = tool_subtool_levels),
    document_type = factor(document_type, levels = document_type_levels),
    subject_type = factor(subject_type, levels = subject_order),
    topic_family = factor(topic_family, levels = object_alpha_levels)
  )

tool_summary_major <- tool_hits |>
  count(year, major_tool)

tool_policy_profile <- expand_grid(
  policy |>
    arrange(year, id) |>
    transmute(
      policy_id = id,
      policy_seq = row_number(),
      document_type = factor(document_type, levels = document_type_levels)
    ),
  subtool_label = tool_subtool_levels
) |>
  left_join(
    tool_hits |>
      count(policy_id, subtool_label, name = "tool_sentences"),
    by = c("policy_id", "subtool_label")
  ) |>
  mutate(
    tool_sentences = replace_na(tool_sentences, 0L),
    subtool_label = factor(subtool_label, levels = tool_subtool_levels)
  ) |>
  group_by(policy_id) |>
  mutate(
    total_tool_sentences = sum(tool_sentences),
    tool_share = if_else(total_tool_sentences > 0, tool_sentences / total_tool_sentences, 0)
  ) |>
  ungroup()

tool_doc_counts <- tool_hits |>
  count(subtool_label, document_type)

tool_subject_counts <- tool_hits |>
  count(subtool_label, subject_type)

tool_object_counts <- tool_hits |>
  count(subtool_label, topic_family)

tool_doc_heatmap <- tool_doc_counts |>
  ggplot(aes(document_type, subtool_label, fill = n)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.4) +
  scale_fill_heat(name = "Sentences") +
  labs(x = "Policy Type", y = "Policy Tool") +
  theme_policy(10) +
  coord_cartesian(clip = "off") +
  theme(
    axis.text.x = element_text(angle = 35, hjust = 1, size = 8.3),
    axis.title.y = element_text(face = "bold"),
    plot.margin = margin(5.5, 5.5, 5.5, 32)
  )

tool_subject_heatmap <- tool_subject_counts |>
  ggplot(aes(subject_type, subtool_label, fill = n)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.4) +
  scale_fill_heat(name = "Sentences") +
  labs(x = "Policy Subject", y = NULL) +
  theme_policy(10) +
  theme(
    axis.text.x = element_text(angle = 35, hjust = 1, size = 9.0),
    axis.text.y = element_blank(),
    plot.tag.position = c(-0.30, 1.02),
    plot.margin = margin(5.5, 5.5, 5.5, 14)
  )

tool_object_heatmap <- tool_object_counts |>
  ggplot(aes(topic_family, subtool_label, fill = n)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.4) +
  scale_fill_heat(name = "Sentences") +
  labs(x = "Policy Domain", y = NULL) +
  theme_policy(10) +
  theme(
    axis.text.x = element_text(angle = 35, hjust = 1),
    axis.text.y = element_blank(),
    plot.tag.position = c(-0.30, 1.02),
    plot.margin = margin(5.5, 5.5, 5.5, 14)
  )

tool_time_plot <- tool_summary_major |>
  ggplot(aes(year, n, fill = major_tool)) +
  geom_col(width = 0.85) +
  labs(x = "Year", y = "Number of Tool-coded Sentences", fill = NULL) +
  theme_policy(10)

tool_policy_profile_total_plot <- tool_policy_profile |>
  distinct(policy_id, policy_seq, document_type, total_tool_sentences) |>
  ggplot(aes(policy_seq, total_tool_sentences, fill = document_type)) +
  geom_col(width = 0.85) +
  scale_fill_manual(values = document_type_palette, name = "Policy Type", drop = FALSE) +
  labs(x = NULL, y = "Number of\nTool-coded\nSentences") +
  theme_policy(9) +
  theme(
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    axis.title.x = element_blank(),
    panel.grid.major.x = element_blank(),
    axis.title.y = element_text(size = 7.2, lineheight = 0.92),
    legend.position = "right",
    legend.justification = c(1, 1),
    legend.direction = "vertical",
    legend.box.background = element_rect(fill = alpha("#FFFFFF", 0.92), colour = NA),
    legend.key.size = grid::unit(0.22, "cm"),
    legend.text = element_text(size = 6.5),
    legend.title = element_text(size = 8.0, face = "bold"),
    legend.margin = margin(3, 4, 3, 4)
  )

tool_policy_profile_heatmap <- tool_policy_profile |>
  ggplot(aes(policy_seq, subtool_label, fill = tool_sentences)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.2) +
  scale_fill_heat(name = "Tool-coded\nSentences") +
  labs(x = "Policy Sequence", y = "Policy Tool") +
  theme_policy(9) +
  theme(
    axis.text.y = element_text(size = 6.8),
    axis.title.y = element_text(face = "bold"),
    plot.margin = margin(5.5, 0, 5.5, 5.5)
  )

tool_policy_profile_right_bar <- tool_policy_profile |>
  group_by(subtool_label) |>
  summarise(
    overall_share = sum(tool_sentences) / sum(total_tool_sentences),
    .groups = "drop"
  ) |>
  ggplot(aes(overall_share, subtool_label)) +
  geom_col(width = 0.82, show.legend = FALSE, fill = "#AEB7C2") +
  scale_x_continuous(limits = c(0, 0.25), labels = percent_format(accuracy = 1)) +
  labs(x = "Overall Share", y = NULL) +
  theme_policy(8) +
  theme(
    axis.text.y = element_blank(),
    axis.ticks.y = element_blank(),
    axis.title.y = element_blank(),
    panel.grid.major.y = element_blank(),
    plot.margin = margin(5.5, 5.5, 5.5, 0)
  )

tool_doc_top_bar <- tool_doc_counts |>
  count(document_type, wt = n, name = "total") |>
  mutate(document_type = factor(document_type, levels = document_type_levels)) |>
  make_top_margin_bar("document_type", "total", make_uniform_fill(document_type_levels), y_title = "Number of\nTool-coded Sentences") +
  theme(axis.title.y = element_text(size = 7.2, lineheight = 0.92))

tool_subject_top_bar <- tool_subject_counts |>
  count(subject_type, wt = n, name = "total") |>
  mutate(subject_type = factor(subject_type, levels = subject_order)) |>
  make_top_margin_bar("subject_type", "total", make_uniform_fill(subject_order), y_title = NULL)

tool_object_top_bar <- tool_object_counts |>
  count(topic_family, wt = n, name = "total") |>
  mutate(topic_family = factor(topic_family, levels = object_alpha_levels)) |>
  make_top_margin_bar("topic_family", "total", make_uniform_fill(object_alpha_levels), y_title = NULL)

tool_doc_panel <- wrap_elements(full = compose_heatmap_with_top(tool_doc_heatmap, tool_doc_top_bar, top_height = 0.20))
tool_subject_panel <- wrap_elements(full = compose_heatmap_with_top(tool_subject_heatmap, tool_subject_top_bar, top_height = 0.20))
tool_object_panel <- wrap_elements(full = compose_heatmap_with_top(tool_object_heatmap, tool_object_top_bar, top_height = 0.20))
tool_policy_profile_panel_top <- (tool_policy_profile_total_plot + guide_area() + plot_layout(widths = c(1.08, 0.12), guides = "collect")) &
  theme(
    legend.position = "right",
    legend.justification = c(1, 1),
    legend.direction = "vertical",
    legend.box.background = element_rect(fill = alpha("#FFFFFF", 0.92), colour = NA),
    legend.key.size = grid::unit(0.22, "cm"),
    legend.text = element_text(size = 6.5),
    legend.title = element_text(size = 8.0, face = "bold"),
    legend.margin = margin(3, 4, 3, 4)
  )

tool_policy_profile_panel <- wrap_elements(
  full = tool_policy_profile_panel_top /
    (tool_policy_profile_heatmap + tool_policy_profile_right_bar + plot_layout(widths = c(1.08, 0.12))) +
    plot_layout(heights = c(0.92, 2.15))
)

tool_panel <- tool_time_plot /
  tool_policy_profile_panel /
  ((tool_doc_panel + tool_subject_panel + tool_object_panel) + plot_layout(ncol = 3)) +
  plot_layout(ncol = 1, heights = c(0.52, 1.02, 1.32)) +
  plot_annotation(tag_levels = list(c("a", "b", "c", "d", "e"))) &
  panel_tag_theme

save_plot(tool_object_heatmap, file.path(paths$policy_tools, "policy_tool_mix_heatmap.png"), 9, 6)
save_combined_plot(tool_panel, file.path(paths$policy_tools, "policy_tool_panel.png"), 18.8, 19.6)
save_plot(tool_time_plot, file.path(paths$policy_tools, "policy_tool_shares_over_time.png"), 12, 6)
write_excel_csv(tool_hits, file.path(paths$policy_tools, "policy_tool_sentence_table.csv"))
write.xlsx(tool_hits, file.path(paths$policy_tools, "policy_tool_sentence_table.xlsx"), overwrite = TRUE)
write_csv(tool_summary_major, file.path(paths$policy_tools, "policy_tool_major_summary.csv"))

# ----- PMC policy evaluation -----------------------------------------------------
# This module scores each policy on ten PMC dimensions and compares group means.
message("Computing PMC policy evaluation scores.")

policy_tool_counts <- tool_hits |>
  count(policy_id, major_tool, name = "tool_sentences") |>
  pivot_wider(names_from = major_tool, values_from = tool_sentences, values_fill = 0)

pmc_features <- policy |>
  transmute(
    policy_id = id,
    year,
    policy_title = title_en,
    document_type = as.character(document_type),
    subject_type = as.character(subject_type),
    topic_family = as.character(topic_family),
    authority_strength = case_when(
      document_type == "Law" ~ 1,
      document_type == "Administrative Regulation" ~ 0.9,
      document_type == "Party Regulation" ~ 0.88,
      document_type == "State Council Normative Document" ~ 0.8,
      document_type == "Departmental Regulation" ~ 0.72,
      document_type == "Departmental Normative Document" ~ 0.65,
      TRUE ~ 0.55
    ),
    timeliness_planning = pmin(1, (str_count(text, "到20\\d{2}年|阶段|中长期|近期|远期|规划期") + 1) / 4),
    objective_clarity = pmin(1, (str_count(text, "目标|任务|总体要求|主要目标|量化|约束性指标") + 1) / 5),
    policy_scope = case_when(
      topic_family %in% c("Nature Conservation", "Habitats and Ecological Space", "Ecological Restoration") ~ 0.9,
      topic_family %in% c("Forests and Grasslands", "Waters and Aquatic Systems", "Climate and Carbon") ~ 0.82,
      TRUE ~ 0.74
    ),
    implementation_mechanisms = pmin(1, (str_count(text, "职责|责任|分工|协调|联席|落实|实施") + 1) / 6),
    support_measures = pmin(1, (str_count(text, "资金|财政|补贴|科技|人才|培训|平台|基础设施") + 1) / 6),
    participation_partnerships = pmin(1, (str_count(text, "公众|社会组织|志愿者|社会资本|合作|国际合作") + 1) / 5),
    monitoring_evaluation = pmin(1, (str_count(text, "监测|监督|评估|评价|考核|审计|动态调整") + 1) / 6),
    internal_consistency = pmax(0.35, 1 - pmin(str_count(text, "既要|又要|兼顾|统筹|同时"), 8) / 12)
  ) |>
  left_join(policy_tool_counts, by = "policy_id") |>
  mutate(
    `Supply-side` = replace_na(`Supply-side`, 0L),
    `Environmental-side` = replace_na(`Environmental-side`, 0L),
    `Demand-side` = replace_na(`Demand-side`, 0L),
    tool_diversity = ((`Supply-side` > 0) + (`Environmental-side` > 0) + (`Demand-side` > 0)) / 3
  )

pmc_dimension_names <- c(
  "Authority Strength",
  "Timeliness and Planning Horizon",
  "Objective Clarity",
  "Policy Scope",
  "Tool Diversity",
  "Implementation Mechanisms",
  "Support Measures",
  "Participation and Partnerships",
  "Monitoring and Evaluation",
  "Internal Consistency"
)

policy_pmc <- pmc_features |>
  transmute(
    policy_id,
    policy_seq = row_number(),
    year,
    policy_title,
    document_type = factor(document_type, levels = document_type_levels),
    subject_type = factor(subject_type, levels = subject_order),
    topic_family = factor(topic_family, levels = object_order),
    `Authority Strength` = authority_strength,
    `Timeliness and Planning Horizon` = timeliness_planning,
    `Objective Clarity` = objective_clarity,
    `Policy Scope` = policy_scope,
    `Tool Diversity` = tool_diversity,
    `Implementation Mechanisms` = implementation_mechanisms,
    `Support Measures` = support_measures,
    `Participation and Partnerships` = participation_partnerships,
    `Monitoring and Evaluation` = monitoring_evaluation,
    `Internal Consistency` = internal_consistency
  ) |>
  mutate(
    pmc_index = rowSums(across(all_of(pmc_dimension_names))),
    pmc_grade = case_when(
      pmc_index >= 8 ~ "Excellent",
      pmc_index >= 6.5 ~ "Good",
      pmc_index >= 5 ~ "Moderate",
      TRUE ~ "Weak"
    )
  )

pmc_long <- policy_pmc |>
  pivot_longer(
    cols = all_of(pmc_dimension_names),
    names_to = "dimension",
    values_to = "score"
  ) |>
  mutate(dimension = factor(dimension, levels = rev(pmc_dimension_names)))

pmc_total_plot <- policy_pmc |>
  ggplot(aes(policy_seq, pmc_index, fill = document_type)) +
  geom_col(width = 0.85) +
  scale_fill_manual(values = document_type_palette, name = "Policy Type", drop = FALSE) +
  labs(x = NULL, y = "PMC Index") +
  theme_policy(9) +
  theme(
    axis.text.x = element_blank(),
    axis.title.x = element_blank(),
    legend.position = "right",
    legend.justification = c(1, 1),
    legend.direction = "vertical",
    legend.box.background = element_rect(fill = alpha("#FFFFFF", 0.92), colour = NA),
    legend.key.size = grid::unit(0.22, "cm"),
    legend.text = element_text(size = 6.5),
    legend.title = element_text(size = 8.0, face = "bold"),
    legend.margin = margin(3, 4, 3, 4)
  )

pmc_all_heatmap <- pmc_long |>
  ggplot(aes(policy_seq, dimension, fill = score)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.2) +
  scale_fill_heat(limits = c(0, 1), name = "Score") +
  labs(x = "Policy Sequence", y = "PMC Criteria") +
  theme_policy(9)

pmc_dimension_bar <- pmc_long |>
  group_by(dimension) |>
  summarise(score = mean(score), .groups = "drop") |>
  ggplot(aes(score, dimension, fill = score)) +
  geom_col(width = 0.88) +
  scale_fill_gradientn(colours = c("#F6E944", "#72C96B", "#2A788E", "#404788"), guide = "none") +
  scale_x_continuous(limits = c(0, 1.02)) +
  labs(x = "Average PMC Score", y = NULL) +
  theme_policy(9) +
  theme(
    axis.text.y = element_blank(),
    axis.ticks.y = element_blank(),
    plot.margin = margin(5.5, 5.5, 5.5, 0)
  )

pmc_panel_a_top <- (pmc_total_plot + guide_area() + plot_layout(widths = c(1.08, 0.12), guides = "collect")) &
  theme(
    legend.position = "right",
    legend.justification = c(1, 1),
    legend.direction = "vertical",
    legend.box.background = element_rect(fill = alpha("#FFFFFF", 0.92), colour = NA),
    legend.key.size = grid::unit(0.22, "cm"),
    legend.text = element_text(size = 6.5),
    legend.title = element_text(size = 8.0, face = "bold"),
    legend.margin = margin(3, 4, 3, 4)
  )

pmc_panel_a <- pmc_panel_a_top /
  (pmc_all_heatmap + pmc_dimension_bar + plot_layout(widths = c(1.08, 0.12))) +
  plot_layout(heights = c(0.92, 2.15))

pmc_doc_heatmap <- pmc_long |>
  group_by(document_type, dimension) |>
  summarise(score = mean(score), .groups = "drop") |>
  ggplot(aes(document_type, dimension, fill = score)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.3) +
  scale_fill_heat(limits = c(0, 1), name = "Score") +
  labs(x = "Policy Type", y = "PMC Criteria") +
  theme_policy(9) +
  theme(
    axis.text.x = element_text(angle = 35, hjust = 1, size = 8.4),
    axis.title.y = element_text(face = "bold")
  )

pmc_subject_heatmap <- pmc_long |>
  group_by(subject_type, dimension) |>
  summarise(score = mean(score), .groups = "drop") |>
  ggplot(aes(subject_type, dimension, fill = score)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.3) +
  scale_fill_heat(limits = c(0, 1), name = "Score") +
  labs(x = "Policy Subject", y = NULL) +
  theme_policy(9) +
  theme(
    axis.text.x = element_text(angle = 35, hjust = 1, size = 8.4),
    axis.text.y = element_blank(),
    plot.tag.position = c(-0.30, 1.02),
    plot.margin = margin(5.5, 5.5, 5.5, 14)
  )

pmc_object_heatmap <- pmc_long |>
  mutate(topic_family = factor(topic_family, levels = object_order)) |>
  group_by(topic_family, dimension) |>
  summarise(score = mean(score), .groups = "drop") |>
  ggplot(aes(topic_family, dimension, fill = score)) +
  geom_tile(colour = "#FFFFFF", linewidth = 0.3) +
  scale_fill_heat(limits = c(0, 1), name = "Score") +
  labs(x = "Policy Domain", y = NULL) +
  theme_policy(9) +
  theme(
    axis.text.x = element_text(angle = 35, hjust = 1),
    axis.text.y = element_blank(),
    plot.tag.position = c(-0.30, 1.02),
    plot.margin = margin(5.5, 5.5, 5.5, 14)
  )

pmc_doc_index_bar <- policy_pmc |>
  group_by(document_type) |>
  summarise(pmc_index = mean(pmc_index), .groups = "drop") |>
  make_top_margin_bar("document_type", "pmc_index", make_uniform_fill(document_type_levels), y_title = "PMC Index")

pmc_subject_index_bar <- policy_pmc |>
  group_by(subject_type) |>
  summarise(pmc_index = mean(pmc_index), .groups = "drop") |>
  make_top_margin_bar("subject_type", "pmc_index", make_uniform_fill(subject_order), y_title = NULL)

pmc_object_index_bar <- policy_pmc |>
  group_by(topic_family) |>
  summarise(pmc_index = mean(pmc_index), .groups = "drop") |>
  make_top_margin_bar("topic_family", "pmc_index", make_uniform_fill(object_order), y_title = NULL)

pmc_doc_panel <- compose_heatmap_with_top(pmc_doc_heatmap, pmc_doc_index_bar, top_height = 0.22)
pmc_subject_panel <- compose_heatmap_with_top(pmc_subject_heatmap, pmc_subject_index_bar, top_height = 0.22)
pmc_object_panel <- compose_heatmap_with_top(pmc_object_heatmap, pmc_object_index_bar, top_height = 0.22)
pmc_doc_panel_wrapped <- wrap_elements(full = pmc_doc_panel)
pmc_subject_panel_wrapped <- wrap_elements(full = pmc_subject_panel)
pmc_object_panel_wrapped <- wrap_elements(full = pmc_object_panel)

pmc_panel <- wrap_elements(full = pmc_panel_a) /
  ((pmc_doc_panel_wrapped + pmc_subject_panel_wrapped + pmc_object_panel_wrapped) + plot_layout(ncol = 3)) +
  plot_layout(ncol = 1, heights = c(0.68, 1.32)) +
  plot_annotation(tag_levels = list(c("a", "b", "c", "d"))) &
  panel_tag_theme

save_combined_plot(pmc_panel, file.path(paths$policy_pmc, "pmc_panel.png"), 20.8, 13.6)
write_csv(policy_pmc, file.path(paths$policy_pmc, "pmc_policy_scores.csv"))

# ----- Literature-side evidence loading ------------------------------------------
# This module prepares the research-side yearly series and keyword metadata.
message("Loading literature-side evidence.")

literature_network_path <- file.path(literature_dir, "20250903V2network_summary_table.csv")
burst_xlsx_path <- file.path(literature_dir, "20250903V2突现词分析.xlsx")
docx_path_main <- file.path(literature_dir, "20260313除重.docx")

if (!file.exists(literature_network_path)) {
  stop("Required file does not exist: ", literature_network_path)
}
if (!file.exists(burst_xlsx_path)) {
  stop("Required file does not exist: ", burst_xlsx_path)
}
if (!file.exists(docx_path_main)) {
  stop("Required file does not exist: ", docx_path_main)
}
if (!file.exists(research_keyword_kept_path)) {
  stop("Required file does not exist: ", research_keyword_kept_path)
}

literature_network <- read_csv(literature_network_path, show_col_types = FALSE) |>
  filter(!is.na(Label), !is.na(Year)) |>
  mutate(
    Year = as.integer(Year),
    BurstBegin = suppressWarnings(as.integer(BurstBegin)),
    BurstEnd = suppressWarnings(as.integer(BurstEnd))
  )

research_keyword_kept <- read_csv(research_keyword_kept_path, show_col_types = FALSE)
required_research_keyword_cols <- c("keyword", "frequency", "first_year")
missing_research_keyword_cols <- setdiff(required_research_keyword_cols, names(research_keyword_kept))
if (length(missing_research_keyword_cols) > 0) {
  stop(
    "research_keywords_kept.csv is missing required columns: ",
    paste(missing_research_keyword_cols, collapse = ", ")
  )
}

research_keyword_kept <- tibble(
  keyword = col_or_na(research_keyword_kept, "keyword"),
  frequency = suppressWarnings(as.numeric(col_or_na(research_keyword_kept, "frequency"))),
  burst = suppressWarnings(as.numeric(col_or_na(research_keyword_kept, "burst"))),
  first_year = suppressWarnings(as.integer(col_or_na(research_keyword_kept, "first_year")))
) |>
  mutate(
    keyword = as.character(keyword)
  ) |>
  filter(
    !is.na(keyword), keyword != "",
    !is.na(first_year),
    first_year >= policy_year_min,
    first_year <= policy_year_max
  )

burst_raw <- load_excel_ascii(burst_xlsx_path, col_names = FALSE) |>
  transmute(
    c1 = `...1`,
    c2 = `...2`,
    c3 = `...3`,
    c4 = `...4`,
    c5 = `...5`,
    c6 = `...6`
  )

burst_table <- burst_raw |>
  slice(-1) |>
  rename(
    keyword = c1,
    first_year = c2,
    burst_strength = c3,
    burst_begin = c4,
    burst_end = c5,
    burst_bar = c6
  ) |>
  mutate(
    first_year = as.integer(first_year),
    burst_strength = as.numeric(burst_strength),
    burst_begin = as.integer(burst_begin),
    burst_end = as.integer(burst_end)
  ) |>
  filter(!is.na(keyword), !is.na(burst_begin))

extract_literature_year_counts <- function() {
  unzip_dir <- tempfile()
  dir.create(unzip_dir, recursive = TRUE)
  unzip(docx_path_main, exdir = unzip_dir)
  
  xml_file <- file.path(unzip_dir, "word", "document.xml")
  if (!file.exists(xml_file)) {
    stop("Could not find document.xml after unzipping: ", docx_path_main)
  }
  
  xml_doc <- xml2::read_xml(xml_file)
  ns <- xml2::xml_ns(xml_doc)
  paragraph_nodes <- xml2::xml_find_all(xml_doc, ".//w:p", ns)
  paragraphs <- vapply(paragraph_nodes, function(node) {
    text_nodes <- xml2::xml_find_all(node, ".//w:t", ns)
    str_squish(paste(xml2::xml_text(text_nodes), collapse = ""))
  }, character(1))
  paragraphs <- paragraphs[paragraphs != ""]
  
  years <- c()
  counts <- c()
  for (i in seq_len(length(paragraphs) - 1)) {
    if (str_detect(paragraphs[i], "^(19|20)\\d{2}$") && str_detect(paragraphs[i + 1], "^\\d+$")) {
      years <- c(years, as.integer(paragraphs[i]))
      counts <- c(counts, as.integer(paragraphs[i + 1]))
    }
  }
  
  out <- data.frame(
    year = years,
    publications = counts,
    stringsAsFactors = FALSE
  )
  
  if (nrow(out) == 0) {
    stop("Failed to recover yearly publication counts from 20260313除重.docx")
  }
  
  out <- out[!duplicated(out$year), , drop = FALSE]
  out <- out[order(out$year), , drop = FALSE]
  as_tibble(out)
}

literature_years <- extract_literature_year_counts()

segmented_breaks <- function(df, value_col) {
  series <- df |>
    transmute(
      year = as.integer(year),
      value = as.numeric(.data[[value_col]])
    ) |>
    filter(!is.na(year), !is.na(value))
  
  if (nrow(series) < 10) {
    idx <- unique(pmax(2, pmin(nrow(series) - 1, floor(c(nrow(series) / 3, 2 * nrow(series) / 3)))))
    return(series$year[idx][seq_len(min(2, length(idx)))])
  }
  
  yrs <- series$year
  y <- log(series$value + 1)
  best <- list(rss = Inf, i = NA_integer_, j = NA_integer_)
  
  for (i in 4:(nrow(series) - 6)) {
    for (j in (i + 4):(nrow(series) - 2)) {
      fit1 <- lm(y[1:i] ~ yrs[1:i])
      fit2 <- lm(y[(i + 1):j] ~ yrs[(i + 1):j])
      fit3 <- lm(y[(j + 1):nrow(series)] ~ yrs[(j + 1):nrow(series)])
      rss <- sum(resid(fit1)^2) + sum(resid(fit2)^2) + sum(resid(fit3)^2)
      if (rss < best$rss) {
        best <- list(rss = rss, i = i, j = j)
      }
    }
  }
  
  if (is.na(best$i) || is.na(best$j)) {
    idx <- unique(pmax(2, pmin(nrow(series) - 1, floor(c(nrow(series) / 3, 2 * nrow(series) / 3)))))
    return(series$year[idx][seq_len(min(2, length(idx)))])
  }
  
  c(series$year[best$i], series$year[best$j])
}

# ----- Literature stage comparison -----------------------------------------------
# Fixed stage windows make the policy and research timelines comparable.
stage_year_range <- policy_year_min:policy_year_max
policy_break_years <- c(2000L, 2015L)
research_break_years <- c(2005L, 2015L)

policy_year_counts <- tibble(year = stage_year_range) |>
  left_join(subject_object_table |>
              count(year, name = "policies"), by = "year") |>
  mutate(policies = replace_na(policies, 0))

policy_stage_lookup <- c(
  "Stage I" = "Stage I (1990-2000): Framework building\nunder single-statute governance",
  "Stage II" = "Stage II (2001-2015): Framework refinement\nunder multi-actor participation",
  "Stage III" = "Stage III (2016-2024):\nSystemic governance under\necological civilization"
)

research_stage_lookup <- c(
  "Stage I" = "Stage I (1990-2005): Conceptual Introduction and\nEcological Baseline Assessment",
  "Stage II" = "Stage II (2006-2015):\nMethodological Advancement\nand Ecological Process Research",
  "Stage III" = "Stage III (2016-2024):\nGovernance-Oriented Research\nand Multi-Scale System Integration"
)

literature_years <- literature_years |>
  mutate(
    stage_id = case_when(
      year <= research_break_years[1] ~ "Stage I",
      year <= research_break_years[2] ~ "Stage II",
      TRUE ~ "Stage III"
    ),
    stage = factor(research_stage_lookup[stage_id], levels = unname(research_stage_lookup))
  )

stage_themes <- literature_network |>
  filter(!is_theme_stopword(Label)) |>
  mutate(
    stage_id = case_when(
      Year <= research_break_years[1] ~ "Stage I",
      Year <= research_break_years[2] ~ "Stage II",
      TRUE ~ "Stage III"
    ),
    stage = factor(research_stage_lookup[stage_id], levels = unname(research_stage_lookup))
  ) |>
  group_by(stage) |>
  slice_max(order_by = Freq + ifelse(is.na(Burst), 0, Burst) * 8, n = 12, with_ties = FALSE) |>
  ungroup() |>
  select(stage, Label, Year, Freq, Burst, BurstBegin, BurstEnd)

policy_year_counts <- policy_year_counts |>
  mutate(
    stage_id = case_when(
      year <= policy_break_years[1] ~ "Stage I",
      year <= policy_break_years[2] ~ "Stage II",
      TRUE ~ "Stage III"
    ),
    stage = factor(policy_stage_lookup[stage_id], levels = unname(policy_stage_lookup))
  )

literature_years_full <- tibble(year = stage_year_range) |>
  left_join(literature_years |> select(year, publications), by = "year") |>
  mutate(
    publications = replace_na(publications, 0L),
    stage_id = case_when(
      year <= research_break_years[1] ~ "Stage I",
      year <= research_break_years[2] ~ "Stage II",
      TRUE ~ "Stage III"
    ),
    stage = factor(research_stage_lookup[stage_id], levels = unname(research_stage_lookup))
  )

policy_stage_palette <- c(
  "Stage I (1990-2000): Framework building\nunder single-statute governance" = "#D8E8C4",
  "Stage II (2001-2015): Framework refinement\nunder multi-actor participation" = "#F5D89B",
  "Stage III (2016-2024):\nSystemic governance under\necological civilization" = "#D6E2FF"
)

research_stage_palette <- c(
  "Stage I (1990-2005): Conceptual Introduction and\nEcological Baseline Assessment" = "#D8E8C4",
  "Stage II (2006-2015):\nMethodological Advancement\nand Ecological Process Research" = "#F5D89B",
  "Stage III (2016-2024):\nGovernance-Oriented Research\nand Multi-Scale System Integration" = "#D6E2FF"
)

make_stage_arrow_polygons <- function(stage_df, y_center, arrow_height) {
  map_dfr(seq_len(nrow(stage_df)), function(i) {
    head_width <- min(0.8, (stage_df$x_end[i] - stage_df$x_start[i]) / 5)
    tibble(
      stage = stage_df$stage[i],
      x = c(
        stage_df$x_start[i],
        stage_df$x_start[i] + head_width,
        stage_df$x_end[i] - head_width,
        stage_df$x_end[i],
        stage_df$x_end[i] - head_width,
        stage_df$x_start[i] + head_width
      ),
      y = c(
        y_center,
        y_center + arrow_height / 2,
        y_center + arrow_height / 2,
        y_center,
        y_center - arrow_height / 2,
        y_center - arrow_height / 2
      )
    )
  })
}

policy_stage_annotation <- tibble(
  stage = unname(policy_stage_lookup),
  x_start = c(policy_year_min, 2001, 2016),
  x_end = c(2000, 2015, 2024)
) |>
  mutate(
    x_mid = (x_start + x_end) / 2,
    stage_label_size = c(4.0, 4.0, 4.0)
  )

research_stage_annotation <- tibble(
  stage = unname(research_stage_lookup),
  x_start = c(policy_year_min, 2006, 2016),
  x_end = c(2005, 2015, 2024)
) |>
  mutate(
    x_mid = (x_start + x_end) / 2,
    stage_label_size = c(4.0, 4.0, 4.0)
  )

policy_max <- max(policy_year_counts$policies)
research_max <- max(literature_years_full$publications)
stage_panel_row_heights <- c(1.44, 0.56)
policy_plot_top <- policy_max * 2.30
research_plot_top <- research_max * 1.18
policy_stage_y <- policy_plot_top * 0.95
research_stage_y <- research_plot_top * 0.95
policy_arrow_height <- policy_plot_top * 0.085
research_arrow_height <- research_plot_top * ((policy_arrow_height / policy_plot_top) * stage_panel_row_heights[1] / stage_panel_row_heights[2])
policy_arrow_poly <- make_stage_arrow_polygons(policy_stage_annotation, policy_stage_y, policy_arrow_height)
research_arrow_poly <- make_stage_arrow_polygons(research_stage_annotation, research_stage_y, research_arrow_height)

# Manual label controls for Figure a event annotations.
# Edit `label_x` / `label_y` to move the label, `segment_end_y` to change connector length,
# and either `wrap_width` or `label_text` (with manual "\n" line breaks) to control wrapping.
prepare_policy_event_labels <- function(event_df, event_color) {
  event_df |>
    left_join(policy_year_counts |> select(year, policies), by = "year") |>
    mutate(
      policies = replace_na(policies, 0),
      label_x = coalesce(label_x, as.numeric(year)),
      label_y = coalesce(label_y, policies + policy_max * 0.20),
      segment_end_y = coalesce(segment_end_y, label_y - policy_max * 0.045),
      wrap_width = coalesce(wrap_width, 20L),
      label_hjust = coalesce(label_hjust, 0.5),
      event_color = event_color
    ) |>
    mutate(
      display_label = map2_chr(
        label_text,
        wrap_width,
        ~ if (str_detect(.x, fixed("\n"))) .x else str_wrap(.x, width = .y)
      )
    )
}

policy_international_events <- tribble(
  ~year, ~label_text, ~label_x, ~label_y, ~segment_end_y, ~wrap_width, ~label_hjust,
  1987L, "Brundtland Report published", 1987, policy_max * 1.8, NA_real_, 20L, 0.5,
  1992L, "China signs the CBD", 1992, policy_max * 1.8, NA_real_, 20L, 0.5,
  2002L, "CBD 2010 Biodiversity Target adopted", 2002, policy_max * 1.8, NA_real_, 18L, 0.5,
  2010L, "COP10:\nNagoya Protocol &\nAichi Targets", 2010, policy_max * 1.8, NA_real_, 18L, 0.5,
  2015L, "UN 2030 Agenda & Paris Agreement adopted", 2015, policy_max * 1.8, NA_real_, 18L, 0.5,
  2019L, "IPBES Global Biodiversity Assessment published", 2019, policy_max * 1.8, NA_real_, 18L, 0.5,
  2021L, "COP15 Kunming: Kunming Declaration", 2021, policy_max * 2.0, NA_real_, 18L, 0.5,
  2022L, "Kunming-Montreal GBF adopted", 2022, policy_max * 1.8, NA_real_, 18L, 0.5
) |>
  filter(year >= policy_year_min) |>
  prepare_policy_event_labels(event_color = "#3E6AA1")

policy_domestic_events <- tribble(
  ~year, ~label_text, ~label_x, ~label_y, ~segment_end_y, ~wrap_width, ~label_hjust,
  1980L, "China joins IUCN", 1980, policy_max * 1.0, NA_real_, 16L, 0.5,
  1983L, "Environmental protection as fundamental national policy", 1983, policy_max * 1.0, NA_real_, 16L, 0.5,
  1994L, "First National Biodiversity Action Plan", 1994, policy_max * 1.0, NA_real_, 16L, 0.5,
  1998L, "Catastrophic Yangtze floods", 1998, policy_max * 1.4, NA_real_, 16L, 0.5,
  1998L, "The NFPP launched\n(Phase I)", 1998, policy_max * 1.0, NA_real_, 16L, 0.5,
  1999L, "The SLCP was piloted", 1999, policy_max * 1.2, NA_real_, 16L, 0.5,
  2002L, "The SLCP was scaled nationally", 2002, policy_max * 1.2, NA_real_, 16L, 0.5,
  2003L, "SARS outbreak", 2003, policy_max * 1.0, NA_real_, 16L, 0.5,
  2005L, "\"Two Mountains\" concept introduced", 2005, policy_max * 1.2, NA_real_, 16L, 0.5,
  2007L, "17th CPC Congress introduces ecological civilization", 2007, policy_max * 1.0, NA_real_, 18L, 0.5,
  2010L, "National Biodiversity Strategy (2011-2030) released", 2010, policy_max * 1.0, NA_real_, 18L, 0.5,
  2011L, "National Biodiversity\nConservation Committee\nestablished", 2011, policy_max * 1.2, NA_real_, 18L, 0.5,
  2011L, "The NFPP (Phase II) implemented", 2011, policy_max * 0.8, NA_real_, 18L, 0.5,
  2012L, "18th CPC Congress:\necological civilization\nin Five-Sphere Plan", 2012, policy_max * 1.4, NA_real_, 20L, 0.5,
  2013L, "3rd Plenum: ecological civilization institutional framework", 2013, policy_max * 1.0, NA_real_, 18L, 0.5,
  2015L, "Overall Plan for Ecological Civilization Reform issued", 2015, policy_max * 1.2, NA_real_, 16L, 0.5,
  2016L, "Nationwide ban on commercial logging in natural forests", 2016, policy_max * 1.4, NA_real_, 16L, 0.5,
  2017L, "19th CPC\nCongress:\nNPCPAS", 2017, policy_max * 0.7, NA_real_, 16L, 0.5,
  2018L, "Institutional reform:\nMEE & MNR\nestablished", 2018, policy_max * 1.2, NA_real_, 16L, 0.5,
  2018L, "Ecological civilization enshrined in Constitution", 2018, policy_max * 1.0, NA_real_, 16L, 0.5,
  2019L, "4th Plenum of 19th CPC: improving national park conservation system", 2019, policy_max * 0.7, NA_real_, 16L, 0.5,
  2020L, "COVID-19 & accelerated Biosecurity Law enactment", 2020, policy_max * 1.6, NA_real_, 16L, 0.5,
  2020L, "Ten-year Yangtze fishing ban launched", 2020, policy_max * 1.4, NA_real_, 16L, 0.5,
  2021L, "First five national parks established", 2021, policy_max * 1.2, NA_real_, 16L, 0.5,
  2022L, "20th CPC\nCongress:\nadvancing\nNPCPAS", 2022, policy_max * 1.5, NA_real_, 16L, 0.5,
  2024L, "Government Work\nReport: advancing\nNPCPAS", 2024, policy_max * 1.5, NA_real_, 16L, 0.5
) |>
  filter(year >= policy_year_min) |>
  prepare_policy_event_labels(event_color = "#C76D3A")

policy_stage_bar <- ggplot(policy_year_counts, aes(year, policies, fill = stage)) +
  geom_col(width = 0.85) +
  geom_vline(xintercept = policy_break_years + 0.5, linetype = "dashed", colour = "#96A1AF", linewidth = 1.0) +
  geom_polygon(
    data = policy_arrow_poly,
    aes(x = x, y = y, group = stage, fill = stage),
    inherit.aes = FALSE,
    alpha = 0.88,
    colour = NA
  ) +
  geom_text(
    data = policy_stage_annotation,
    aes(x = x_mid, y = policy_stage_y, label = stage, size = stage_label_size),
    inherit.aes = FALSE,
    lineheight = 0.94,
    fontface = "bold",
    colour = "#16203A"
  ) +
  scale_size_identity() +
  geom_segment(
    data = policy_international_events,
    aes(x = year, xend = label_x, y = policies, yend = segment_end_y),
    inherit.aes = FALSE,
    colour = policy_international_events$event_color,
    linewidth = 0.42,
    linetype = "dashed",
    alpha = 0.8
  ) +
  geom_label(
    data = policy_international_events,
    aes(x = label_x, y = label_y, label = display_label, hjust = label_hjust),
    inherit.aes = FALSE,
    label.padding = grid::unit(0.11, "lines"),
    label.r = grid::unit(0.12, "lines"),
    linewidth = 0.36,
    size = 3.0,
    lineheight = 0.92,
    fill = "white",
    colour = policy_international_events$event_color
  ) +
  geom_segment(
    data = policy_domestic_events,
    aes(x = year, xend = label_x, y = policies, yend = segment_end_y),
    inherit.aes = FALSE,
    colour = policy_domestic_events$event_color,
    linewidth = 0.42,
    linetype = "dashed",
    alpha = 0.8
  ) +
  geom_label(
    data = policy_domestic_events,
    aes(x = label_x, y = label_y, label = display_label, hjust = label_hjust),
    inherit.aes = FALSE,
    label.padding = grid::unit(0.11, "lines"),
    label.r = grid::unit(0.12, "lines"),
    linewidth = 0.36,
    size = 3.0,
    lineheight = 0.92,
    fill = "white",
    colour = policy_domestic_events$event_color
  ) +
  scale_fill_manual(values = policy_stage_palette, guide = "none") +
  scale_x_continuous(
    limits = c(min(stage_year_range) - 0.5, max(stage_year_range) + 0.5),
    breaks = stage_year_range
  ) +
  labs(x = NULL, y = "Policies", fill = NULL) +
  theme_policy(10) +
  theme(
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8.0),
    legend.position = "none"
  ) +
  coord_cartesian(ylim = c(0, policy_plot_top), clip = "off")

research_stage_bar <- ggplot(literature_years_full, aes(year, publications, fill = stage)) +
  geom_col(width = 0.85) +
  geom_vline(xintercept = research_break_years + 0.5, linetype = "dashed", colour = "#96A1AF", linewidth = 1.0) +
  geom_polygon(
    data = research_arrow_poly,
    aes(x = x, y = y, group = stage, fill = stage),
    inherit.aes = FALSE,
    alpha = 0.88,
    colour = NA
  ) +
  geom_text(
    data = research_stage_annotation,
    aes(x = x_mid, y = research_stage_y, label = stage, size = stage_label_size),
    inherit.aes = FALSE,
    lineheight = 0.94,
    fontface = "bold",
    colour = "#16203A"
  ) +
  scale_size_identity() +
  scale_fill_manual(values = research_stage_palette, guide = "none") +
  scale_x_continuous(
    limits = c(min(stage_year_range) - 0.5, max(stage_year_range) + 0.5),
    breaks = stage_year_range
  ) +
  labs(x = "Year", y = "Publications", fill = NULL) +
  theme_policy(10) +
  theme(
    legend.position = "none",
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8.0)
  ) +
  coord_cartesian(ylim = c(0, research_plot_top), clip = "off")

# ----- Custom legend strip (solid rounded box with dashed line below) -----------
# A hand-drawn legend that mirrors the annotation style used in Figure a.
# Each legend key: solid-border rounded label on top, dashed connector hanging below.
legend_items <- tibble(
  x_box = c(1.8, 7.0),
  x_text = c(2.3, 7.5),
  box_y = c(0.52, 0.52),
  line_top = c(0.45, 0.45),
  line_bot = c(0.30, 0.30),
  text_y = c(0.52, 0.52),
  color = c("#3E6AA1", "#C76D3A"),
  label = c("International biodiversity conservation milestone",
            "Domestic biodiversity conservation milestone")
)

milestone_legend <- ggplot() +
  geom_segment(
    data = legend_items,
    aes(x = x_box, xend = x_box, y = line_top, yend = line_bot),
    colour = legend_items$color,
    linetype = "dashed",
    linewidth = 0.55
  ) +
  geom_label(
    data = legend_items,
    aes(x = x_box, y = box_y, label = "     "),
    colour = legend_items$color,
    fill = "white",
    label.r = unit(0.15, "lines"),
    label.padding = unit(0.18, "lines"),
    linewidth = 0.45,
    size = 3.5
  ) +
  geom_text(
    data = legend_items,
    aes(x = x_text, y = text_y, label = label),
    hjust = 0,
    vjust = 0.5,
    size = 3.6,
    lineheight = 0.95,
    colour = "#222222"
  ) +
  scale_x_continuous(limits = c(0, 12)) +
  scale_y_continuous(limits = c(0, 1)) +
  theme_void() +
  theme(plot.margin = margin(0, 0, 2, 0, "pt"))

# Combine two main panels and the custom legend strip at the bottom.
stage_panel <- policy_stage_bar / research_stage_bar / milestone_legend +
  plot_layout(heights = c(stage_panel_row_heights, 0.11)) +
  plot_annotation(tag_levels = list(c("a", "b", ""))) &
  panel_tag_theme &
  theme(
    plot.background = element_rect(fill = "white", colour = NA),
    panel.background = element_rect(fill = "white", colour = NA)
  )

save_plot(research_stage_bar, file.path(paths$literature_stage, "research_stage_segmentation.png"), 14, 5)
save_combined_plot(stage_panel, file.path(paths$literature_stage, "policy_research_stage_comparison.png"), 16.5, 12.6)

write_csv(literature_years_full, file.path(paths$literature_stage, "yearly_publications.csv"))
write_csv(policy_year_counts, file.path(paths$literature_stage, "yearly_policy_counts.csv"))
write_csv(stage_themes, file.path(paths$literature_stage, "stage_theme_summary.csv"))

# ----- Literature burstness and timezone -----------------------------------------
# These plots summarise burst keywords and the first appearance time of themes.
message("Rendering literature burstness and timezone visuals.")

burst_plot_data <- burst_table |>
  slice_head(n = 25) |>
  mutate(table_order = row_number()) |>
  arrange(table_order) |>
  mutate(
    keyword = factor(keyword, levels = rev(keyword))
  )

burst_plot <- ggplot(burst_plot_data, aes(y = keyword)) +
  geom_vline(
    xintercept = seq(policy_year_min, policy_year_max, by = 1),
    colour = "grey85",
    linewidth = 0.35
  ) +
  geom_segment(
    aes(x = first_year, xend = policy_year_max, yend = keyword),
    linewidth = 2.6,
    colour = "#D7EAF3"
  ) +
  geom_segment(
    aes(x = burst_begin, xend = burst_end, yend = keyword, colour = burst_strength),
    linewidth = 3.4
  ) +
  geom_text(
    aes(
      x = pmin(burst_end + 0.35, policy_year_max + 0.9),
      label = sprintf("%.2f", burst_strength)
    ),
    hjust = 0,
    size = 3
  ) +
  scale_colour_gradient(low = "#86BBD8", high = "#C44536") +
  scale_x_continuous(
    breaks = seq(policy_year_min, policy_year_max, by = 1),
    limits = c(policy_year_min, policy_year_max + 1.5),
    expand = c(0, 0)
  ) +
  labs(
    x = "Year",
    y = NULL,
    colour = "Burst\nstrength"
  ) +
  theme_policy(10) +
  theme(
    panel.grid.major.x = element_blank(),
    panel.grid.minor.x = element_blank(),
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)
  )

save_combined_plot(burst_plot, file.path(paths$literature_burst, "literature_burst_keywords.png"), 13, 9)

timeline_plot <- literature_network |>
  filter(!is_theme_stopword(Label)) |>
  slice_max(order_by = Freq, n = 50) |>
  mutate(
    label = str_replace_all(Label, "_", " "),
    stage = case_when(
      Year <= research_break_years[1] ~ "Stage I",
      Year <= research_break_years[2] ~ "Stage II",
      TRUE ~ "Stage III"
    )
  ) |>
  ggplot(aes(Year, reorder(label, Year), size = Freq, colour = stage)) +
  geom_point(alpha = 0.9) +
  scale_size_continuous(range = c(2.2, 10)) +
  scale_colour_manual(values = c("Stage I" = "#0B4F6C", "Stage II" = "#F26419", "Stage III" = "#7A306C")) +
  labs(
    x = "First appearance year",
    y = NULL,
    size = "Frequency",
    colour = NULL
  ) +
  theme_policy(10)

save_plot(timeline_plot, file.path(paths$literature_timeline, "literature_keyword_timezone.png"), 13, 11)

write_csv(burst_table, file.path(paths$literature_burst, "burst_keyword_table.csv"))
write_csv(literature_network, file.path(paths$literature_timeline, "literature_network_summary_table.csv"))

# ----- Policy-research timing interaction ----------------------------------------
# This module compares when policy themes and research themes first appear.
message("Comparing policy-first and research-first timing using both shared and non-overlapping themes.")

theme_group_lookup <- keyword_dict |>
  transmute(
    theme_en = canonicalize_policy_keyword_label(en),
    theme_group = group
  ) |>
  distinct(theme_en, .keep_all = TRUE)

policy_theme_evidence <- policy_theme_years |>
  filter(!is_theme_stopword(theme_en)) |>
  transmute(
    theme_en,
    theme_zh,
    theme_group,
    policy_first_year = first_policy_year,
    policy_document_frequency = document_frequency
  )

literature_theme_evidence <- research_keyword_kept |>
  transmute(
    theme_en = canonicalize_policy_keyword_label(keyword),
    research_first_year = first_year,
    research_frequency = frequency,
    research_burst = burst
  ) |>
  filter(!is.na(research_first_year)) |>
  filter(!is_theme_stopword(theme_en)) |>
  group_by(theme_en) |>
  summarise(
    research_first_year = min(research_first_year, na.rm = TRUE),
    research_frequency = max(research_frequency, na.rm = TRUE),
    research_burst = if_else(all(is.na(research_burst)), NA_real_, max(research_burst, na.rm = TRUE)),
    .groups = "drop"
  ) |>
  left_join(theme_group_lookup, by = "theme_en")

interaction_timing <- full_join(policy_theme_evidence, literature_theme_evidence, by = "theme_en", suffix = c(".policy", ".research")) |>
  mutate(
    theme_group = coalesce(theme_group.policy, theme_group.research),
    # ---- Tolerance window: ±1 year counts as "Synchronous" ----
    year_gap = as.integer(policy_first_year - research_first_year),
    tolerance = 1L,
    evidence_type = case_when(
      is.na(research_first_year) & !is.na(policy_first_year) ~ "Policy-led",
      is.na(policy_first_year) & !is.na(research_first_year) ~ "Research-led",
      !is.na(year_gap) & abs(year_gap) <= tolerance ~ "Synchronous",
      !is.na(year_gap) & year_gap < -tolerance ~ "Policy-led",
      !is.na(year_gap) & year_gap > tolerance ~ "Research-led",
      TRUE ~ NA_character_
    ),
    evidence_basis = case_when(
      is.na(research_first_year) ~ "Policy-only theme",
      is.na(policy_first_year) ~ "Research-only theme",
      abs(year_gap) <= tolerance ~ paste0("Shared theme: synchronous (gap=", year_gap, ")"),
      year_gap < -tolerance ~ paste0("Shared theme: policy appears earlier (gap=", year_gap, ")"),
      year_gap > tolerance ~ paste0("Shared theme: research appears earlier (gap=+", year_gap, ")"),
      TRUE ~ "Shared theme: same first year"
    ),
    evidence_year = case_when(
      evidence_type == "Policy-led" ~ policy_first_year,
      evidence_type == "Research-led" ~ research_first_year,
      evidence_type == "Synchronous" ~ pmin(policy_first_year, research_first_year, na.rm = TRUE),
      TRUE ~ NA_real_
    ),
    # ---- Document-frequency weighting ----
    # Weight each theme by how many documents mention it (sum of policy + research
    # frequency, capped at 1 after log-scaling) so that themes appearing in many
    # documents count more than single-mention curiosities.
    freq_policy = replace_na(policy_document_frequency, 0L),
    freq_research = replace_na(research_frequency, 0L),
    freq_weight = log1p(freq_policy + freq_research)
  ) |>
  mutate(evidence_year = as.integer(evidence_year)) |>
  arrange(evidence_year, theme_en)

# ---- Unweighted yearly shares (original method, for baseline comparison) ----
interaction_yearly_share <- interaction_timing |>
  filter(!is.na(evidence_year), !is.na(evidence_type)) |>
  count(evidence_year, evidence_type) |>
  group_by(evidence_year) |>
  mutate(share = n / sum(n)) |>
  ungroup()

# ---- Frequency-weighted yearly shares ----
interaction_yearly_share_weighted <- interaction_timing |>
  filter(!is.na(evidence_year), !is.na(evidence_type)) |>
  group_by(evidence_year, evidence_type) |>
  summarise(weighted_n = sum(freq_weight), .groups = "drop") |>
  group_by(evidence_year) |>
  mutate(share_weighted = weighted_n / sum(weighted_n)) |>
  ungroup()

interaction_yearly_share_full <- expand_grid(
  evidence_year = policy_year_min:policy_year_max,
  evidence_type = c("Policy-led", "Synchronous", "Research-led")
) |>
  left_join(interaction_yearly_share, by = c("evidence_year", "evidence_type")) |>
  left_join(interaction_yearly_share_weighted |> select(evidence_year, evidence_type, share_weighted),
            by = c("evidence_year", "evidence_type")) |>
  mutate(
    n = replace_na(n, 0L),
    share = replace_na(share, 0),
    share_weighted = replace_na(share_weighted, 0)
  )

dominance_periods <- interaction_yearly_share_full |>
  filter(evidence_type %in% c("Policy-led", "Research-led")) |>
  select(evidence_year, evidence_type, share) |>
  pivot_wider(names_from = evidence_type, values_from = share, values_fill = 0) |>
  mutate(has_observation = (`Policy-led` + `Research-led`) > 0) |>
  filter(has_observation) |>
  mutate(
    dominance = if_else(
      `Research-led` > `Policy-led`,
      "Research Dominates & Policy Dominated",
      "Policy Dominates & Research Dominated"
    ),
    block = cumsum(
      dominance != lag(dominance, default = first(dominance)) |
        evidence_year != lag(evidence_year, default = first(evidence_year)) + 1
    )
  ) |>
  group_by(block, dominance) |>
  summarise(
    xmin = pmax(min(evidence_year) - 0.5, policy_year_min),
    xmax = pmin(max(evidence_year) + 0.5, policy_year_max),
    border_color = if_else(
      first(dominance) == "Research Dominates & Policy Dominated",
      "#4169E1",
      "#FA9107"
    ),
    .groups = "drop"
  )

# Weighted dominance periods (same logic but using share_weighted).
dominance_periods_weighted <- interaction_yearly_share_full |>
  filter(evidence_type %in% c("Policy-led", "Research-led")) |>
  select(evidence_year, evidence_type, share_weighted) |>
  pivot_wider(names_from = evidence_type, values_from = share_weighted, values_fill = 0) |>
  mutate(has_observation = (`Policy-led` + `Research-led`) > 0) |>
  filter(has_observation) |>
  mutate(
    dominance = if_else(
      `Research-led` > `Policy-led`,
      "Research Dominates & Policy Dominated",
      "Policy Dominates & Research Dominated"
    ),
    block = cumsum(
      dominance != lag(dominance, default = first(dominance)) |
        evidence_year != lag(evidence_year, default = first(evidence_year)) + 1
    )
  ) |>
  group_by(block, dominance) |>
  summarise(
    xmin = pmax(min(evidence_year) - 0.5, policy_year_min),
    xmax = pmin(max(evidence_year) + 0.5, policy_year_max),
    border_color = if_else(
      first(dominance) == "Research Dominates & Policy Dominated",
      "#4169E1",
      "#F5784B"
    ),
    .groups = "drop"
  )

# Unweighted share plot (original).
interaction_share_plot <- ggplot(interaction_yearly_share_full, aes(evidence_year, share, fill = evidence_type)) +
  geom_area(position = "fill", alpha = 0.95, linewidth = 0.15, colour = "white") +
  geom_rect(
    data = dominance_periods,
    aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = 1, colour = dominance),
    inherit.aes = FALSE,
    fill = NA,
    linewidth = 1.15
  ) +
  scale_y_continuous(labels = percent_format()) +
  scale_fill_manual(values = c(
    "Policy-led" = "#FCFFA4",
    "Synchronous" = "#CFEDBE",
    "Research-led" = "#D7EAF3"
  )) +
  scale_colour_manual(values = c(
    "Research Dominates & Policy Dominated" = "#4169E1",
    "Policy Dominates & Research Dominated" = "#F5784B"
  ), name = NULL) +
  scale_x_continuous(
    limits = c(policy_year_min, policy_year_max),
    breaks = policy_year_min:policy_year_max
  ) +
  labs(x = "Year", y = "Share (unweighted, ±1yr tolerance)", fill = NULL) +
  guides(
    fill = guide_legend(order = 1, nrow = 1, byrow = TRUE),
    colour = guide_legend(order = 2, nrow = 1, byrow = TRUE, title.position = "left", title.hjust = 0)
  ) +
  theme_policy(10) +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8))

# Frequency-weighted share plot.
interaction_share_plot_weighted <- ggplot(interaction_yearly_share_full, aes(evidence_year, share_weighted, fill = evidence_type)) +
  geom_area(position = "fill", alpha = 0.95, linewidth = 0.15, colour = "white") +
  geom_rect(
    data = dominance_periods_weighted,
    aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = 1, colour = dominance),
    inherit.aes = FALSE,
    fill = NA,
    linewidth = 1.15
  ) +
  scale_y_continuous(labels = percent_format()) +
  scale_fill_manual(values = c(
    "Policy-led" = "#FCFFA4",
    "Synchronous" = "#CFEDBE",
    "Research-led" = "#D7EAF3"
  )) +
  scale_colour_manual(values = c(
    "Research Dominates & Policy Dominated" = "#4169E1",
    "Policy Dominates & Research Dominated" = "#F5784B"
  ), name = NULL) +
  scale_x_continuous(
    limits = c(policy_year_min, policy_year_max),
    breaks = policy_year_min:policy_year_max
  ) +
  labs(x = "Year", y = "Share (frequency-weighted, ±1yr tolerance)", fill = NULL) +
  guides(
    fill = guide_legend(order = 1, nrow = 1, byrow = TRUE),
    colour = guide_legend(order = 2, nrow = 1, byrow = TRUE, title.position = "left", title.hjust = 0)
  ) +
  theme_policy(10) +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8))

interaction_share_plot_tagged <- interaction_share_plot +
  labs(tag = "a") +
  panel_tag_theme

interaction_share_plot_weighted_tagged <- interaction_share_plot_weighted +
  labs(tag = "b") +
  panel_tag_theme

interaction_share_panel <- interaction_share_plot_tagged / interaction_share_plot_weighted_tagged +
  plot_layout(ncol = 1, guides = "collect") &
  theme(
    legend.position = "bottom",
    legend.box = "vertical",
    legend.box.just = "center",
    legend.justification = "center",
    plot.tag = element_text(face = "bold", size = 25, colour = "#101820")
  )

save_plot(interaction_share_plot_tagged, file.path(paths$interaction, "policy_research_timing_shares.png"), 13, 8)
save_plot(interaction_share_plot_weighted_tagged, file.path(paths$interaction, "policy_research_timing_shares_weighted.png"), 13, 8)
save_combined_plot(interaction_share_panel, file.path(paths$interaction, "policy_research_timing_shares_panel.png"), 13, 16)
write_csv(interaction_timing, file.path(paths$interaction, "policy_research_timing_evidence.csv"))
write_csv(interaction_yearly_share_full, file.path(paths$interaction, "policy_research_timing_shares.csv"))

message("Writing consolidated Excel workbooks.")

wb_policy <- createWorkbook()
addWorksheet(wb_policy, "TypeSubjectDomain")
writeDataTable(wb_policy, "TypeSubjectDomain", subject_object_table)
addWorksheet(wb_policy, "TypeSubjectDomainSummary")
writeDataTable(wb_policy, "TypeSubjectDomainSummary", subject_object_summary)
addWorksheet(wb_policy, "AgencyDocDomainLinks")
writeDataTable(wb_policy, "AgencyDocDomainLinks", policy_unit_document_object_linkage)
addWorksheet(wb_policy, "IntergovNodes")
writeDataTable(wb_policy, "IntergovNodes", intergov_nodes)
addWorksheet(wb_policy, "IntergovEdges")
writeDataTable(wb_policy, "IntergovEdges", intergov_edges)
addWorksheet(wb_policy, "KeywordFreq")
writeDataTable(wb_policy, "KeywordFreq", keyword_keep)
addWorksheet(wb_policy, "KeywordNodes")
writeDataTable(wb_policy, "KeywordNodes", keyword_nodes_out)
addWorksheet(wb_policy, "KeywordEdges")
writeDataTable(wb_policy, "KeywordEdges", keyword_edges)
addWorksheet(wb_policy, "ToolHits")
writeDataTable(wb_policy, "ToolHits", tool_hits)
addWorksheet(wb_policy, "PMC")
writeDataTable(wb_policy, "PMC", policy_pmc)
saveWorkbook(wb_policy, file.path(output_root, "policy_analysis_tables.xlsx"), overwrite = TRUE)

wb_literature <- createWorkbook()
addWorksheet(wb_literature, "YearlyPublications")
writeDataTable(wb_literature, "YearlyPublications", literature_years_full)
addWorksheet(wb_literature, "YearlyPolicies")
writeDataTable(wb_literature, "YearlyPolicies", policy_year_counts)
addWorksheet(wb_literature, "StageThemes")
writeDataTable(wb_literature, "StageThemes", stage_themes)
addWorksheet(wb_literature, "BurstKeywords")
writeDataTable(wb_literature, "BurstKeywords", burst_table)
addWorksheet(wb_literature, "TimezoneSummary")
writeDataTable(wb_literature, "TimezoneSummary", literature_network)
addWorksheet(wb_literature, "TimingEvidence")
writeDataTable(wb_literature, "TimingEvidence", interaction_timing)
addWorksheet(wb_literature, "TimingShares")
writeDataTable(wb_literature, "TimingShares", interaction_yearly_share_full)
saveWorkbook(wb_literature, file.path(output_root, "literature_interaction_tables.xlsx"), overwrite = TRUE)

report_lines <- c(
  "# Biodiversity Policy and Research Interaction Outputs",
  "",
  sprintf(
    "This folder contains fully English tables and figures generated from %d biodiversity policy documents and the local CiteSpace-related literature files.",
    nrow(policy)
  ),
  "",
  "## Policy modules",
  "- `policy/01_type_subject_domain`: policy type-subject-domain tables and visuals.",
  "  - `document_domain_classification.csv`: auto-generated document-to-domain mapping (output on every run).",
  "  - `document_domain_classification_checked.csv`: place manually corrected version here; fill `corrected_topic_family` column and re-run.",
  "- `policy/02_intergovernmental_network`: co-issuance network data for Gephi and network charts.",
  "  - Edge weights include raw co-occurrence, Jaccard, and cosine similarity.",
  "  - `centrality_report.csv`: agencies ranked by Jaccard-normalised betweenness (brokerage), eigenvector (hierarchical influence), PageRank, constraint (structural holes), hub/authority (HITS), clustering coefficient, and k-core.",
  "- `policy/03_keyword_network`: keyword co-occurrence network data for Gephi and keyword timeline charts.",
  "  - Keyword nodes include betweenness, eigenvector, closeness, PageRank, constraint, clustering coefficient, k-core, hub, and authority centrality.",
  "- `policy/04_policy_tools`: sentence-level policy tool coding and tool-mix charts.",
  "  - Tool detection uses context filtering: negated, definitional, and historical sentences are excluded.",
  "- `policy/05_pmc_evaluation`: automatic PMC evaluation tables and figures.",
  "",
  "## Literature modules",
  "- `literature/01_research_stages`: policy-versus-research annual stage comparison and stage themes.",
  "- `literature/02_burstness`: translated burstness chart and clean data table.",
  "- `literature/03_timezone`: reconstructed timezone figure and cleaned summary data.",
  "- `literature/04_policy_research_interaction`: timing evidence including shared and non-overlapping themes.",
  "  - Uses ±1-year tolerance window for 'Synchronous' classification.",
  "  - Includes both unweighted and frequency-weighted share plots.",
  "  - `analysis_scripts/plot_policy_research_keyword_inheritance_timeline.R` also writes annual complete 5-year cohort conversion, directionality, and post-conversion persistence figures.",
  "",
  "## Methodological robustness",
  "- `icr_coding_sample.csv`: stratified 20% sample for inter-coder reliability assessment.",
  "  - Fill in `coder2_object`, save as `icr_coding_completed.csv`, and re-run to get Cohen's kappa.",
  "- `dictionary_coverage_report.csv`: keyword dictionary coverage diagnostics.",
  "- `dictionary_expansion_candidates.csv`: high-frequency terms in uncovered policies (if any).",
  "- `policy_text_file_id_map.csv`: mapping from stable policy IDs to the time-ordered txt file numbering used in `policy text`.",
  "",
  "## Important assumptions",
  "- Policy texts were merged by the renumbered txt-file sequence after restricting the corpus to 1990-2024; `policy_text_file_id_map.csv` records the mapping back to stable policy IDs.",
  "- Policy IDs 98 and 99 were split from the combined text file using the heading `国有林区改革指导意见`.",
  "- Document-level policy domain classification uses a two-step workflow: auto-classify → manually correct → re-run. If `document_domain_classification_checked.csv` exists, its `corrected_topic_family` column overrides the auto-classification.",
  "- Literature yearly publication counts are extracted from `20260313除重.docx` covering 1990-2024.",
  "- Policy analyses are restricted to documents first enacted in 1990-2024 so the policy and research windows are directly comparable.",
  "- The stage comparison uses fixed windows of 1990-2000, 2001-2015, and 2016-2024 for policy, alongside 1990-2005, 2006-2015, and 2016-2024 for research.",
  "- Policy-research timing evidence includes themes unique to policy and themes unique to research, not only overlapping themes.",
  "- Policy-research timing shares use `research_keywords_kept.csv`, which excludes literature-side stopwords before computing yearly shares.",
  "- Interaction timing uses a ±1-year tolerance window: themes whose first policy and research appearances differ by 1 year or fewer are classified as 'Synchronous' rather than led by either side.",
  "- Cohort conversion probabilities use annual complete 5-year source cohorts; post-conversion persistence is measured from receiving-side yearly reappearances 1-10 years after uptake.",
  "- Frequency-weighted shares weight each theme by log(1 + policy_doc_freq + research_freq) so that themes mentioned in many documents carry more influence than single-mention curiosities.",
  "- Intergovernmental network centrality is reported with both raw and Jaccard-normalised edge weights; Jaccard normalisation controls for agency prolificacy.",
  "- Tool coding excludes sentences matching negation, definitional, or historical-reference patterns to improve precision.",
  "- Heatmap colours are standardised to a yellow-green-blue gradient and plot titles are intentionally removed from the figures; descriptive titles remain in the R script as comments."
)
writeLines(report_lines, file.path(output_root, "README.md"))

message("Pipeline completed. Outputs written to: ", output_root)




type_levels_small_to_large <- c(
  "Departmental Working Document",
  "Departmental Normative Document",
  "State Council Normative Document",
  "Departmental Regulation",
  "Administrative Regulation",
  "Law"
)

type_labels_wrapped <- c(
  "Departmental\nWorking\nDocument",
  "Departmental\nNormative\nDocument",
  "State Council\nNormative\nDocument",
  "Departmental\nRegulation",
  "Administrative\nRegulation",
  "Law"
)

plot_df <- policy_pmc |>
  filter(document_type %in% type_levels_small_to_large) |>
  mutate(
    document_type = factor(document_type, levels = type_levels_small_to_large),
    admin_level = case_when(
      document_type == "Departmental Working Document" ~ 1,
      document_type == "Departmental Normative Document" ~ 2,
      document_type == "State Council Normative Document" ~ 3,
      document_type == "Departmental Regulation" ~ 4,
      document_type == "Administrative Regulation" ~ 5,
      document_type == "Law" ~ 6
    )
  )


library(dplyr)
library(ggplot2)
library(patchwork)
library(cowplot)

type_levels_small_to_large <- c(
  "Departmental Working Document",
  "Departmental Normative Document",
  "State Council Normative Document",
  "Departmental Regulation",
  "Administrative Regulation",
  "Law"
)

type_labels_wrapped <- c(
  "Departmental\nWorking\nDocument",
  "Departmental\nNormative\nDocument",
  "State Council\nNormative\nDocument",
  "Departmental\nRegulation",
  "Administrative\nRegulation",
  "Law"
)

make_admin_level <- function(document_type) {
  case_when(
    document_type == "Departmental Working Document" ~ 1,
    document_type == "Departmental Normative Document" ~ 2,
    document_type == "State Council Normative Document" ~ 3,
    document_type == "Departmental Regulation" ~ 4,
    document_type == "Administrative Regulation" ~ 5,
    document_type == "Law" ~ 6
  )
}

type_palette_scatter <- document_type_palette[type_levels_small_to_large]

keyword_count_by_policy <- keyword_matrix |>
  mutate(keyword_count = rowSums(across(all_of(keyword_keep$theme_en)))) |>
  select(policy_id, keyword_count)

tool_count_by_policy <- tool_hits |>
  count(policy_id, name = "tool_coded_sentences")

plot_df <- policy_pmc |>
  filter(document_type %in% type_levels_small_to_large) |>
  left_join(keyword_count_by_policy, by = "policy_id") |>
  left_join(tool_count_by_policy, by = "policy_id") |>
  mutate(
    keyword_count = replace_na(keyword_count, 0L),
    tool_coded_sentences = replace_na(tool_coded_sentences, 0L),
    document_type = factor(document_type, levels = type_levels_small_to_large),
    admin_level = make_admin_level(document_type)
  )

common_scatter_theme <- theme_minimal(base_size = 12) +
  theme(
    panel.grid.minor = element_blank(),
    axis.text.x = element_text(lineheight = 0.9),
    plot.margin = margin(5.5, 10, 5.5, 24)
  )

panel_tag_theme_local <- theme(
  plot.tag = element_text(face = "bold", size = 18, colour = "#101820"),
  plot.tag.position = c(0.01, 0.99)
)

make_type_scatter_plot <- function(data, y_var, y_label) {
  ggplot(data, aes(x = admin_level, y = {{ y_var }}, colour = document_type)) +
    geom_point(
      position = position_jitter(width = 0.10, height = 0),
      size = 2.0,
      alpha = 0.75
    ) +
    geom_smooth(
      data = data,
      mapping = aes(
        x = admin_level,
        y = {{ y_var }}
      ),
      inherit.aes = FALSE,
      method = "lm",
      formula = y ~ poly(x, 2, raw = TRUE),
      level = 0.95,
      se = TRUE,
      colour = "#C44536",
      fill = "#F4B6A6",
      linewidth = 1.0,
      alpha = 0.20
    ) +
    scale_colour_manual(
      values = type_palette_scatter,
      drop = FALSE,
      name = "Policy Type"
    ) +
    scale_x_continuous(
      breaks = 1:6,
      labels = type_labels_wrapped,
      limits = c(0.7, 6.3)
    ) +
    labs(
      x = "Policy Type (Ordered by Administrative Hierarchy)",
      y = y_label
    ) +
    common_scatter_theme +
    theme(legend.position = "none")
}

plot_a <- make_type_scatter_plot(
  plot_df,
  keyword_count,
  "Number of Identified Keywords"
) + labs(tag = "a")

plot_b <- make_type_scatter_plot(
  plot_df,
  tool_coded_sentences,
  "Number of Tool-coded Sentences"
) + labs(tag = "b")

plot_c <- make_type_scatter_plot(
  plot_df,
  pmc_index,
  "PMC Index"
) + labs(tag = "c")

main_panel <- plot_a + plot_b + plot_c + plot_layout(ncol = 3)

model_legend_plot <- ggplot() +
  geom_ribbon(
    data = data.frame(x = c(1, 2), ymin = c(0, 0), ymax = c(1, 1)),
    aes(x = x, ymin = ymin, ymax = ymax, fill = "95% CI"),
    alpha = 0.20
  ) +
  geom_line(
    data = data.frame(x = c(1, 2), y = c(0.5, 0.5)),
    aes(x = x, y = y, linetype = "Quadratic fit"),
    colour = "#C44536",
    linewidth = 1.0
  ) +
  scale_fill_manual(
    values = c("95% CI" = "#F4B6A6"),
    name = NULL
  ) +
  scale_linetype_manual(
    values = c("Quadratic fit" = "solid"),
    name = NULL
  ) +
  guides(
    fill = guide_legend(
      order = 1,
      nrow = 1,
      byrow = TRUE,
      override.aes = list(alpha = 0.20)
    ),
    linetype = guide_legend(
      order = 2,
      nrow = 1,
      byrow = TRUE,
      override.aes = list(colour = "#C44536")
    )
  ) +
  theme_void() +
  theme(
    legend.position = "bottom",
    legend.box = "horizontal",
    legend.direction = "horizontal",
    legend.justification = "center"
  )

type_legend_plot <- ggplot(
  plot_df,
  aes(x = admin_level, y = pmc_index, colour = document_type)
) +
  geom_point(size = 2.0) +
  scale_colour_manual(
    values = type_palette_scatter,
    drop = FALSE,
    name = "Policy Type"
  ) +
  guides(
    colour = guide_legend(nrow = 2, byrow = TRUE)
  ) +
  theme_void() +
  theme(
    legend.position = "bottom",
    legend.justification = "center"
  )

model_legend <- cowplot::get_legend(model_legend_plot)
type_legend <- cowplot::get_legend(type_legend_plot)

legend_panel <- wrap_elements(full = model_legend) /
  wrap_elements(full = type_legend) +
  plot_layout(heights = c(0.55, 0.85))

policy_type_abc_panel <- main_panel /
  legend_panel +
  plot_layout(heights = c(1, 0.22)) &
  panel_tag_theme_local

save_combined_plot(
  policy_type_abc_panel,
  file.path(paths$policy_pmc, "policy_type_keyword_tool_pmc_panel.png"),
  18.5,
  7.8
)
