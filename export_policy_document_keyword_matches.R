#!/usr/bin/env Rscript

suppressPackageStartupMessages({
  library(readxl)
  library(readr)
  library(openxlsx)
  library(dplyr)
  library(tidyr)
  library(purrr)
  library(stringr)
  library(tibble)
  library(lubridate)
})

options(stringsAsFactors = FALSE, scipen = 999)
Sys.setenv(LANG = "en_US.UTF-8", LC_ALL = "en_US.UTF-8")
try(Sys.setlocale("LC_ALL", "en_US.UTF-8"), silent = TRUE)
try(Sys.setlocale("LC_CTYPE", "en_US.UTF-8"), silent = TRUE)

# This script exports document-level policy keyword matches for the authoritative
# 231 policy keywords stored in gephi_nodes_checked.csv.
# It writes tabular outputs only and does not currently emit module-level figures,
# so this script does not create companion PDF charts.
#
# Important note on the threshold:
# The "minimum threshold of 5 co-occurrences" (i.e., edges kept at >= 6) is not
# recalculated here. Instead, this script uses gephi_nodes_checked.csv as the
# authoritative final keyword list that already underlies the checked policy
# keyword network. The output therefore reports document-level matches only for
# those 231 vetted keywords.

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
    "Please run this file as a saved script."
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
    "Checked:\n",
    "  1) ", parent, "\n",
    "  2) ", script_dir
  )
}

project_root <- resolve_project_root()
policy_dir <- file.path(project_root, "policy text")
literature_dir <- file.path(project_root, "bibliometric analysis")
output_root <- file.path(project_root, "analysis_outputs", "biodiversity_policy_english")
keyword_output_dir <- file.path(output_root, "policy", "03_keyword_network")

policy_year_min <- 1990L
policy_year_max <- 2024L

policy_keyword_reference_path <- file.path(keyword_output_dir, "gephi_nodes_checked.csv")
policy_text_map_path <- file.path(output_root, "policy_text_file_id_map.csv")
output_xlsx_path <- file.path(keyword_output_dir, "policy_document_keyword_matches_231.xlsx")

dir.create(keyword_output_dir, recursive = TRUE, showWarnings = FALSE)

message("Resolved project_root: ", project_root)
message("Using policy_dir: ", policy_dir)
message("Using literature_dir: ", literature_dir)
message("Using reference keyword file: ", policy_keyword_reference_path)

if (!dir.exists(policy_dir)) {
  stop("Directory does not exist: ", policy_dir)
}
if (!dir.exists(literature_dir)) {
  stop("Directory does not exist: ", literature_dir)
}
if (!file.exists(policy_keyword_reference_path)) {
  stop("Required file does not exist: ", policy_keyword_reference_path)
}

load_excel_ascii <- function(path, ...) {
  if (!file.exists(path)) {
    stop("Source Excel file does not exist: ", path)
  }

  tmp <- tempfile(fileext = ".xlsx")
  ok <- file.copy(path, tmp, overwrite = TRUE)
  if (!ok || !file.exists(tmp)) {
    stop("Failed to copy Excel file to temp location: ", path)
  }

  read_excel(tmp, ...)
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

policy_stage <- function(year) {
  case_when(
    year <= 2000 ~ "Stage I (1990-2000)",
    year <= 2015 ~ "Stage II (2001-2015)",
    TRUE ~ "Stage III (2016-2024)"
  )
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

normalize_policy_domain <- function(x) {
  recode(
    as.character(x),
    "Other Governance Topics" = "Other Policy Domains",
    .default = as.character(x)
  )
}

object_order <- c(
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
)

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
    en %in% c("Environmental Impact", "Environmental Protection Industry") ~ "Pollution Control",
    en %in% c("Green Development", "Circular Economy") ~ "Climate and Carbon",
    en %in% c("Terrestrial Space", "Ecological Space", "Public Participation", "Remote Sensing") ~ "Habitats and Ecological Space",
    en %in% c("Community Structure", "Spartina Alterniflora") ~ "Species Conservation and Use",
    en %in% c("Xishuangbanna") ~ "Forests and Grasslands",
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

policy_object_reference_raw <- read_csv(policy_keyword_reference_path, show_col_types = FALSE)

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
      "Dropped ", length(dropped_terms),
      " dictionary entries not aligned to checked keywords: ",
      paste(head(dropped_terms, 10), collapse = ", "),
      if (length(dropped_terms) > 10) " ..." else ""
    )
  }

  aligned |>
    filter(!is.na(canonical_label))
}

build_text_lookup <- function() {
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
    stop("Expected at most one combined file with an 'x-y、' prefix, found: ", nrow(combined_file))
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
      stop("Could not split combined file because marker '国有林区改革指导意见' was not found.")
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

  texts |>
    arrange(file_id)
}

build_policy_map_from_existing_output <- function() {
  if (!file.exists(policy_text_map_path)) {
    return(NULL)
  }

  message("Using existing policy_text_file_id_map.csv: ", policy_text_map_path)

  read_csv(policy_text_map_path, show_col_types = FALSE) |>
    transmute(
      policy_id = as.integer(policy_id),
      text_file_id = as.integer(text_file_id),
      year = as.integer(year),
      title_en = title_en,
      title_zh = title_zh
    ) |>
    filter(!is.na(text_file_id), !is.na(year))
}

build_policy_map_from_metadata <- function() {
  metadata_candidates <- c(
    file.path(policy_dir, "政策文件清单.xlsx"),
    file.path(policy_dir, "policy documents_1990-2024.xlsx"),
    file.path(project_root, "..", "政策文件1205最终版", "全部政策文件.xlsx")
  )

  required_metadata_cols <- c("No.", "Document Name", "First Enactment Date")

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

  message("Using metadata workbook: ", metadata_path)

  metadata_raw |>
    mutate(
      source_policy_id = as.integer(No.),
      year = extract_year_from_date(`First Enactment Date`),
      title_en = split_bilingual(`Document Name`)$en,
      title_zh = split_bilingual(`Document Name`)$zh
    ) |>
    filter(!is.na(source_policy_id), !is.na(year), year >= policy_year_min, year <= policy_year_max) |>
    arrange(year, source_policy_id) |>
    mutate(text_file_id = row_number()) |>
    transmute(
      policy_id = source_policy_id,
      text_file_id,
      year,
      title_en,
      title_zh
    )
}

build_policy_corpus <- function() {
  texts <- build_text_lookup()
  policy_map <- build_policy_map_from_existing_output()
  if (is.null(policy_map)) {
    policy_map <- build_policy_map_from_metadata()
  }

  policy <- policy_map |>
    left_join(texts, by = c("text_file_id" = "file_id")) |>
    mutate(stage = policy_stage(year)) |>
    arrange(year, policy_id)

  if (any(is.na(policy$text) | policy$text == "")) {
    missing_ids <- policy$policy_id[is.na(policy$text) | policy$text == ""]
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
    mutate(
      en = nounify_theme(en),
      group = policy_object_group(en)
    )
}

format_keywords <- function(x) {
  x <- unique(x[!is.na(x) & x != ""])
  if (length(x) == 0) {
    return(NA_character_)
  }
  paste(x, collapse = "; ")
}

autosize_and_filter <- function(wb, sheet, data) {
  if (nrow(data) == 0 || ncol(data) == 0) {
    return(invisible(NULL))
  }
  freezePane(wb, sheet = sheet, firstRow = TRUE)
  addFilter(wb, sheet = sheet, rows = 1, cols = 1:ncol(data))
  setColWidths(wb, sheet = sheet, cols = 1:ncol(data), widths = "auto")
  invisible(NULL)
}

message("Building policy corpus.")
policy <- build_policy_corpus()

message("Loading keyword dictionary and aligning to authoritative 231 checked keywords.")
keyword_dict <- build_keyword_dictionary()

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
  policy |>
    mutate(match = keyword_presence[[theme]]) |>
    filter(match) |>
    summarise(
      theme_en = theme,
      theme_zh = keyword_summary_raw$theme_zh[i],
      theme_group = keyword_summary_raw$theme_group[i],
      first_policy_year = min(year, na.rm = TRUE),
      document_frequency = n()
    )
})

keyword_alignment <- align_policy_keywords_to_checked(
  keyword_summary = keyword_summary_raw,
  policy_theme_years = policy_theme_years_raw,
  checked_reference = policy_object_reference
)

keyword_alias_map <- keyword_alignment |>
  distinct(canonical_label, raw_theme_en)

canonical_label_order <- policy_object_reference$label

keyword_incidence <- map(canonical_label_order, function(label) {
  raw_terms <- keyword_alias_map$raw_theme_en[keyword_alias_map$canonical_label == label]
  if (length(raw_terms) == 0) {
    return(rep(0L, nrow(policy)))
  }
  as.integer(rowSums(keyword_presence[, raw_terms, drop = FALSE]) > 0)
}) |>
  setNames(canonical_label_order) |>
  as_tibble()

keyword_metadata <- policy_object_reference |>
  transmute(
    keyword_en = label,
    keyword_group = group,
    keyword_document_frequency = checked_size,
    first_policy_year = checked_first_policy_year,
    community = checked_community
  ) |>
  mutate(
    keyword_zh = map_chr(keyword_en, function(label) {
      zh_terms <- keyword_alignment$theme_zh[keyword_alignment$canonical_label == label]
      format_keywords(zh_terms)
    }),
    matched_in_dictionary = keyword_en %in% keyword_alias_map$canonical_label
  )

keyword_matrix <- bind_cols(
  policy |>
    transmute(
      policy_id,
      text_file_id,
      year,
      stage,
      title_en,
      title_zh,
      source_files
    ),
  keyword_incidence
)

keyword_long <- imap_dfr(canonical_label_order, function(label, idx) {
  match_rows <- which(keyword_incidence[[label]] == 1L)
  if (length(match_rows) == 0) {
    return(tibble())
  }

  tibble(
    policy_id = policy$policy_id[match_rows],
    text_file_id = policy$text_file_id[match_rows],
    year = policy$year[match_rows],
    stage = policy$stage[match_rows],
    title_en = policy$title_en[match_rows],
    title_zh = policy$title_zh[match_rows],
    keyword_en = label,
    keyword_zh = keyword_metadata$keyword_zh[match(label, keyword_metadata$keyword_en)],
    keyword_group = keyword_metadata$keyword_group[match(label, keyword_metadata$keyword_en)],
    source_files = policy$source_files[match_rows]
  )
}) |>
  arrange(year, policy_id, keyword_en)

document_summary <- keyword_long |>
  group_by(policy_id, text_file_id, year, stage, title_en, title_zh, source_files) |>
  summarise(
    keyword_count = n(),
    matched_keywords_en = format_keywords(keyword_en),
    matched_keywords_zh = format_keywords(keyword_zh),
    matched_keyword_groups = format_keywords(keyword_group),
    .groups = "drop"
  )

document_summary <- policy |>
  transmute(
    policy_id,
    text_file_id,
    year,
    stage,
    title_en,
    title_zh,
    source_files
  ) |>
  left_join(document_summary, by = c(
    "policy_id", "text_file_id", "year", "stage", "title_en", "title_zh", "source_files"
  )) |>
  mutate(
    keyword_count = coalesce(keyword_count, 0L),
    matched_keywords_en = if_else(keyword_count == 0L, NA_character_, matched_keywords_en),
    matched_keywords_zh = if_else(keyword_count == 0L, NA_character_, matched_keywords_zh),
    matched_keyword_groups = if_else(keyword_count == 0L, NA_character_, matched_keyword_groups)
  ) |>
  arrange(year, policy_id)

stage_summary <- document_summary |>
  group_by(stage) |>
  summarise(
    documents = n(),
    average_keyword_count = round(mean(keyword_count), 3),
    median_keyword_count = median(keyword_count),
    min_keyword_count = min(keyword_count),
    max_keyword_count = max(keyword_count),
    zero_keyword_documents = sum(keyword_count == 0L),
    .groups = "drop"
  )

message("Writing workbook: ", output_xlsx_path)

wb <- createWorkbook()

addWorksheet(wb, "DocumentSummary")
writeData(wb, "DocumentSummary", document_summary)
autosize_and_filter(wb, "DocumentSummary", document_summary)

addWorksheet(wb, "KeywordMatchesLong")
writeData(wb, "KeywordMatchesLong", keyword_long)
autosize_and_filter(wb, "KeywordMatchesLong", keyword_long)

addWorksheet(wb, "KeywordMatrix231")
writeData(wb, "KeywordMatrix231", keyword_matrix)
autosize_and_filter(wb, "KeywordMatrix231", keyword_matrix)

addWorksheet(wb, "KeywordReference231")
writeData(wb, "KeywordReference231", keyword_metadata)
autosize_and_filter(wb, "KeywordReference231", keyword_metadata)

addWorksheet(wb, "StageSummary")
writeData(wb, "StageSummary", stage_summary)
autosize_and_filter(wb, "StageSummary", stage_summary)

saveWorkbook(wb, output_xlsx_path, overwrite = TRUE)

message("Done.")
message("Documents exported: ", nrow(document_summary))
message("Authoritative keywords exported: ", nrow(keyword_metadata))
message("Matched document-keyword pairs exported: ", nrow(keyword_long))
message("No figure outputs are generated by this script.")
message("Workbook written to: ", output_xlsx_path)
